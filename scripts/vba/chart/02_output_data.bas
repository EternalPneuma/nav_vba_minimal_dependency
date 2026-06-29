Option Explicit
Private g_fillCounts As Object

Private Const SHEET_PRODUCT_INFO As String = "产品信息"
Private Const SHEET_CHART_NAV_DATA As String = "绘图净值数据"
Private Const SHEET_DATA_SUMMARY As String = "数据摘要"
Private Const COL_TRUST_CODE As String = "信托计划代码"
Private Const COL_PRODUCT_CODE As String = "产品编号"
Private Const COL_PRODUCT_CODE_ALT As String = "产品代码"
Private Const COL_CODE As String = "代码"
Private Const COL_PRODUCT_NAME As String = "产品名称"
Private Const COL_PRODUCT_FULL_NAME As String = "产品全称"
Private Const COL_TRUST_NAME As String = "信托计划名称"
Private Const COL_PRODUCT_SHORT As String = "产品简称"
Private Const COL_CHART_EXPORT As String = "图表导出"

'==============================================================
' 模块: 导出产品数据
' 功能: 按“绘图净值数据”中有记录且已配置的产品,逐个生成sheet到一个新xlsx文件
'       每个sheet以产品简称命名,包含该产品全部历史数据,按净值日期升序
'       同时合并原02_check_data,在导出工作簿中新增“数据摘要”
'==============================================================

Public Sub Chart02_ExportProductSummary()

    Dim t0 As Double: t0 = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim wsDim As Worksheet: Set wsDim = wbDB.Worksheets(SHEET_PRODUCT_INFO)
    Dim wsData As Worksheet: Set wsData = wbDB.Worksheets(SHEET_CHART_NAV_DATA)
    
    ' 初始化填补统计收集器
    Set g_fillCounts = CreateObject("Scripting.Dictionary")
    
    '--- 1. 读取产品信息: 产品编号/信托计划代码 -> 产品简称/产品名称 ---
    Dim dimDict As Object
    Set dimDict = CreateObject("Scripting.Dictionary")
    dimDict.CompareMode = vbTextCompare
    
    Dim dimHeaderMap As Object
    Set dimHeaderMap = BuildHeaderMap(wsDim, 1)
    
    Dim dimCodeCol As Long
    dimCodeCol = FindFirstExistingHeader(dimHeaderMap, Array(COL_TRUST_CODE, COL_PRODUCT_CODE, COL_PRODUCT_CODE_ALT, COL_CODE))
    If dimCodeCol = 0 Then
        MsgBox "产品净值汇总导出无法继续" & vbCrLf & vbCrLf & _
               "错误信息：产品信息中缺少可用于匹配的代码字段：" & COL_TRUST_CODE & " / " & COL_PRODUCT_CODE & " / " & COL_PRODUCT_CODE_ALT, _
               vbExclamation, "产品净值汇总导出"
        GoTo CleanUp
    End If
    
    Dim dimShortCol As Long
    dimShortCol = FindFirstExistingHeader(dimHeaderMap, Array(COL_PRODUCT_SHORT, COL_PRODUCT_NAME, COL_PRODUCT_FULL_NAME, COL_TRUST_NAME))
    If dimShortCol = 0 Then
        MsgBox "产品净值汇总导出无法继续" & vbCrLf & vbCrLf & _
               "错误信息：产品信息中缺少可用于命名sheet的名称字段：" & COL_PRODUCT_SHORT & " / " & COL_PRODUCT_NAME & " / " & COL_PRODUCT_FULL_NAME, _
               vbExclamation, "产品净值汇总导出"
        GoTo CleanUp
    End If
    
    ' “图表导出”为可选字段：存在时仅“是/Y/YES/1/TRUE/启用”可导出；不存在时保持原有全量导出行为。
    Dim dimChartExportCol As Long
    dimChartExportCol = FindFirstExistingHeader(dimHeaderMap, Array(COL_CHART_EXPORT))
    
    Dim dimExportDict As Object
    Set dimExportDict = CreateObject("Scripting.Dictionary")
    dimExportDict.CompareMode = vbTextCompare
    
    Dim dimLastRow As Long
    dimLastRow = wsDim.Cells(wsDim.Rows.Count, dimCodeCol).End(xlUp).Row
    If dimLastRow < 2 Then
        MsgBox "产品净值汇总导出无法继续" & vbCrLf & vbCrLf & _
               "错误信息：产品信息中没有数据，请先完善产品信息。", vbExclamation, "产品净值汇总导出"
        GoTo CleanUp
    End If
    
    Dim i As Long, prodCode As String, prodShort As String
    Dim chartExportEnabled As Boolean
    For i = 2 To dimLastRow
        prodCode = NormalizeText(wsDim.Cells(i, dimCodeCol).value)
        prodShort = NormalizeText(wsDim.Cells(i, dimShortCol).value)
        If Len(prodCode) > 0 And Len(prodShort) > 0 Then
            dimDict(prodCode) = prodShort
            If dimChartExportCol = 0 Then
                chartExportEnabled = True
            Else
                chartExportEnabled = IsChartExportEnabled(wsDim.Cells(i, dimChartExportCol).value)
            End If
            dimExportDict(prodCode) = chartExportEnabled
        End If
    Next i
        
    If dimDict.Count = 0 Then
        MsgBox "产品净值汇总导出无法继续" & vbCrLf & vbCrLf & _
               "错误信息：产品信息中未读取到有效的产品编号/信托计划代码和产品简称/产品名称。", vbExclamation, "产品净值汇总导出"
        GoTo CleanUp
    End If
        
    '--- 2. 读取“绘图净值数据”全部数据,按产品编号分组,同时去重 ---
    Dim dataLastRow As Long
    dataLastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    If dataLastRow < 2 Then
        MsgBox "产品净值汇总导出无法继续" & vbCrLf & vbCrLf & _
               "错误信息：绘图净值数据中没有数据可导出。", vbExclamation, "产品净值汇总导出"
        GoTo CleanUp
    End If
    
    ' 读取表头(第1行A-J)
    Dim header As Variant
    header = wsData.Range("A1:J1").value
    
    ' 读取数据
    Dim dataArr As Variant
    dataArr = wsData.Range("A2:J" & dataLastRow).value
    
    ' 按产品编号分组: groupDict(prodCode) = 子Dictionary( pk -> 行数组 )
    ' 同时用pk(日期+编号)做去重,后出现的覆盖先出现的
    Dim groupDict As Object
    Set groupDict = CreateObject("Scripting.Dictionary")
    
    Dim maxDate As Date: maxDate = 0
    
    Dim j As Long, c As Long
    Dim curCode As String, pk As String
    Dim curDate As Variant
        
    For j = 1 To UBound(dataArr, 1)
        curDate = dataArr(j, 1)
        curCode = Trim(CStr(dataArr(j, 2)))
        
        If Len(curCode) = 0 Then GoTo NextRow
        If IsEmpty(curDate) Or Not IsDate(curDate) Then GoTo NextRow
        
        ' 跟踪最大日期(用于文件名)
        If CDate(curDate) > maxDate Then maxDate = CDate(curDate)
        
        ' 主键
        pk = Format(CDate(curDate), "yyyy-mm-dd") & "|" & curCode
        
        ' 取出/创建该产品的子Dictionary
        Dim subDict As Object
        If groupDict.Exists(curCode) Then
            Set subDict = groupDict(curCode)
        Else
            Set subDict = CreateObject("Scripting.Dictionary")
            groupDict.Add curCode, subDict
        End If
    
        ' 缓存该行(数组形式)
        Dim rowArr(1 To 10) As Variant
        For c = 1 To 10
            rowArr(c) = dataArr(j, c)
        Next c
        subDict(pk) = rowArr  ' 重复主键自动覆盖
    
NextRow:
    Next j
    
    '--- 3. 创建新工作簿,逐个产品生成sheet ---
    Dim wbOut As Workbook
    Set wbOut = Workbooks.Add
    
    ' 删除默认创建的多余sheet,只留一个占位
    Do While wbOut.Sheets.Count > 1
        wbOut.Sheets(wbOut.Sheets.Count).Delete
    Loop
    
    Dim usedSheetNames As Object
    Set usedSheetNames = CreateObject("Scripting.Dictionary")
    usedSheetNames(SHEET_DATA_SUMMARY) = 1
    
    Dim exportedCount As Long: exportedCount = 0
    Dim exportedProducts As Object: Set exportedProducts = CreateObject("Scripting.Dictionary")
    Dim emptyProducts As Object: Set emptyProducts = CreateObject("Scripting.Dictionary")
    Dim missingProducts As Object: Set missingProducts = CreateObject("Scripting.Dictionary")
    Dim disabledProducts As Object: Set disabledProducts = CreateObject("Scripting.Dictionary")
    
    ' 产品信息已扩充为主数据。无绘图数据的产品仅记录到“数据摘要”，不再逐个写入提示。
    Dim codeKey As Variant
    For Each codeKey In dimDict.keys
        prodCode = CStr(codeKey)
        If Not groupDict.Exists(prodCode) Then emptyProducts(prodCode) = dimDict(prodCode)
    Next codeKey
    
    ' 绘图数据是导出候选的来源；缺少产品信息或未启用图表导出的产品不生成sheet。
    Dim dataCodeKey As Variant
    For Each dataCodeKey In groupDict.keys
        prodCode = CStr(dataCodeKey)
        If Not dimDict.Exists(prodCode) Then
            missingProducts(prodCode) = vbNullString
        ElseIf Not CBool(dimExportDict(prodCode)) Then
            disabledProducts(prodCode) = dimDict(prodCode)
        End If
    Next dataCodeKey
    
    Dim isFirstSheet As Boolean: isFirstSheet = True
    Dim wsOut As Worksheet
        
    ' 按产品信息表顺序遍历；只有绘图净值数据中实际存在的产品才生成sheet。
    For Each codeKey In dimDict.keys
        prodCode = CStr(codeKey)
        prodShort = dimDict(prodCode)
        
        ' 检查该产品是否有数据
        If Not groupDict.Exists(prodCode) Then GoTo NextProduct
        If Not CBool(dimExportDict(prodCode)) Then GoTo NextProduct
        
        Set subDict = groupDict(prodCode)
        If subDict.Count = 0 Then GoTo NextProduct
        
        ' 清洗sheet名
        Dim cleanName As String
        cleanName = CleanSheetName(prodShort, usedSheetNames)
        usedSheetNames(cleanName) = 1
    
        ' 创建sheet
        If isFirstSheet Then
            Set wsOut = wbOut.Sheets(1)
            wsOut.name = cleanName
            isFirstSheet = False
        Else
            Set wsOut = wbOut.Sheets.Add(After:=wbOut.Sheets(wbOut.Sheets.Count))
            wsOut.name = cleanName
        End If

        ' 写入表头
        wsOut.Range("A1:J1").value = header
    
        ' 把subDict里的数据按日期升序排列后写入
        Dim writeArr() As Variant
        Dim nRows As Long: nRows = subDict.Count
        ReDim writeArr(1 To nRows, 1 To 10)
    
        ' 先把所有行收集到临时数组(包含排序键)
        Dim tmpArr() As Variant
        ReDim tmpArr(1 To nRows, 1 To 11)  ' 第11列存排序键
    
        Dim idx As Long: idx = 0
        Dim pkKey As Variant
        For Each pkKey In subDict.keys
            idx = idx + 1
            Dim arr As Variant: arr = subDict(pkKey)
            For c = 1 To 10
                tmpArr(idx, c) = arr(c)
            Next c
            tmpArr(idx, 11) = CStr(pkKey)  ' 主键作为排序键(yyyy-mm-dd|code)
        Next pkKey
    
        ' 按第11列升序排序(冒泡,产品内数据量一般不大)
        SortByCol tmpArr, 11
    
       ' 对净值列(第4列)做分红日平滑,覆盖原值
        SmoothDividendDays tmpArr, 4
    
        ' 基于平滑后的净值,重算30日年化收益率,覆盖第7列(G列),并保留计算审计信息
        Dim yieldAuditArr() As Variant
        Calc30DayAnnualYield tmpArr, 1, 4, 7, prodCode, yieldAuditArr
    
        ' 拷贝到writeArr
        For idx = 1 To nRows
            For c = 1 To 10
                writeArr(idx, c) = tmpArr(idx, c)
            Next c
        Next idx
    
        ' 一次性写入原始数据区
        wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(nRows + 1, 10)).value = writeArr
    
        ' K:N 为30日年化收益率审计字段,不影响图表使用的 A:G 列
        wsOut.Range("K1").value = "30日年化收益率处理方式"
        wsOut.Range("L1").value = "30日年化收益率计算日期"
        wsOut.Range("M1").value = "30日年化收益率基准日期"
        wsOut.Range("N1").value = "30日年化收益率实际间隔(天)"
        wsOut.Range(wsOut.Cells(2, 11), wsOut.Cells(nRows + 1, 14)).value = yieldAuditArr
        
        ' 简单格式化: 表头加粗,日期列格式
        wsOut.Range("A1:N1").Font.Bold = True
        wsOut.Range("A:A").NumberFormat = "yyyy-mm-dd"
        wsOut.Range("L:M").NumberFormat = "yyyy-mm-dd"
        wsOut.Columns("A:N").AutoFit
        
        exportedCount = exportedCount + 1
        exportedProducts(prodCode) = prodShort
        
NextProduct:
    Next codeKey
                
    WriteDataSummarySheet wbOut, dimDict, groupDict, g_fillCounts, _
                          exportedProducts, emptyProducts, missingProducts, disabledProducts
                
    '--- 4. 保存文件 ---
    If exportedCount = 0 Then
        wbOut.Close SaveChanges:=False
        MsgBox "产品净值汇总导出无法继续" & vbCrLf & vbCrLf & _
               "错误信息：没有可导出的产品数据。", vbExclamation, "产品净值汇总导出"
        GoTo CleanUp
    End If
    
    Dim outFileName As String
    If maxDate > 0 Then
        outFileName = "产品净值汇总_" & Format(maxDate, "yyyymmdd") & ".xlsx"
    Else
        outFileName = "产品净值汇总_" & Format(Now, "yyyymmdd") & ".xlsx"
    End If
    
    Dim outPath As String
    outPath = wbDB.Path & "\" & outFileName
    
    ' 如果同名文件已存在,先删除(避免SaveAs弹窗)
    On Error Resume Next
    If Dir(outPath) <> "" Then Kill outPath
    On Error GoTo 0
    
    wbOut.SaveAs fileName:=outPath, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    
    '--- 5. 汇总提示 ---
    Dim msg As String
    msg = "产品净值汇总导出完成" & vbCrLf & vbCrLf & _
          "耗时：" & Format(Timer - t0, "0.00") & " 秒" & vbCrLf & vbCrLf & _
          "处理结果：" & vbCrLf & _
          "导出产品数：" & exportedCount & " 个" & vbCrLf & vbCrLf & _
          "输出文件：" & vbCrLf & outPath
    
    msg = msg & vbCrLf & vbCrLf & _
          "注意事项：" & vbCrLf & _
          "产品信息有、绘图净值数据无（仅记录）：" & emptyProducts.Count & " 个" & vbCrLf & _
          "绘图净值数据有、产品信息未配置（未导出）：" & missingProducts.Count & " 个"
    If dimChartExportCol > 0 Then
        msg = msg & vbCrLf & "绘图净值数据有、图表导出未启用（未导出）：" & disabledProducts.Count & " 个"
    End If
    msg = msg & vbCrLf & "详情见导出文件“数据摘要”sheet。"
        
    Dim totalFilled As Long: totalFilled = 0
    Dim kFill As Variant
    For Each kFill In g_fillCounts.keys
        totalFilled = totalFilled + g_fillCounts(kFill)
    Next kFill
        
    If totalFilled > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "30日年化收益率计算填补：" & totalFilled & " 条" & vbCrLf & _
              "详情见导出文件“数据摘要”sheet，原始有效值已保留。"
    End If
        
    MsgBox msg, vbInformation, "产品净值汇总导出"
        
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub
        
        
'==============================================================
' 辅助函数: 清洗sheet名
'  - 移除非法字符 \ / ? * [ ] :
'  - 长度截断到31
'  - 重名自动加后缀
'==============================================================
Private Function CleanSheetName(ByVal rawName As String, ByVal usedDict As Object) As String
    Dim s As String: s = rawName
    
    ' 替换非法字符
    Dim badChars As Variant
    badChars = Array("\", "/", "?", "*", "[", "]", ":")
    Dim k As Long
    For k = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(k), "-")
    Next k

    ' 去除首尾单引号(Excel不允许sheet名以单引号开头或结尾)
    Do While Left(s, 1) = "'"
        s = Mid(s, 2)
    Loop
    Do While Right(s, 1) = "'"
        s = Left(s, Len(s) - 1)
    Loop
    
    s = Trim(s)
    If Len(s) = 0 Then s = "未命名"

    ' 截断到31字符
    If Len(s) > 31 Then s = Left(s, 31)
    
    ' 处理重名
    Dim baseName As String: baseName = s
    Dim suffix As Long: suffix = 2
    Do While usedDict.Exists(s)
        Dim suffixStr As String: suffixStr = "_" & suffix
        If Len(baseName) + Len(suffixStr) > 31 Then
            s = Left(baseName, 31 - Len(suffixStr)) & suffixStr
        Else
            s = baseName & suffixStr
        End If
        suffix = suffix + 1
    Loop
    
    CleanSheetName = s
End Function

    
'==============================================================
' 辅助函数: 按指定列对二维数组升序排序(冒泡)
'==============================================================
Private Sub SortByCol(ByRef arr As Variant, ByVal sortCol As Long)
    Dim n As Long: n = UBound(arr, 1)
    Dim nCols As Long: nCols = UBound(arr, 2)
    
    Dim i As Long, j As Long, c As Long
    Dim tmp As Variant

    For i = 1 To n - 1
        For j = 1 To n - i
            If CStr(arr(j, sortCol)) > CStr(arr(j + 1, sortCol)) Then
                For c = 1 To nCols
                    tmp = arr(j, c)
                    arr(j, c) = arr(j + 1, c)
                    arr(j + 1, c) = tmp
                Next c
            End If
        Next j
    Next i
End Sub

'==============================================================
' 辅助过程: 对二维数组的指定列做"分红日平滑"
'   - 算法: 基于"反弹特征"识别孤立异常点
'   - 条件: |V(D) - (V(D-1)+V(D+1))/2| > MAX( |V(D+1)-V(D-1)|*K, V(D-1)*threshold )
'   - 替换: 异常点V(D) = (V(D-1)+V(D+1))/2
'   - 边界: 首尾点不处理
'   - 调用前提: 数组已按日期升序排序
'==============================================================
Private Sub SmoothDividendDays(ByRef arr As Variant, ByVal navCol As Long)
    Const K_RATIO As Double = 3#      ' 偏离倍数
    Const THRESHOLD As Double = 0.0005 ' 绝对阈值: 0.05%

    Dim n As Long: n = UBound(arr, 1)
    If n < 3 Then Exit Sub  ' 少于3个点无法平滑

    ' 收集有效数值的索引(跳过非数值/空值)
    ' 注意: 平滑只在"连续有效"的位置之间进行
    ' 即如果 arr(i, navCol) 是数值, 才参与判定

    Dim i As Long
    Dim prevV As Double, curV As Double, nextV As Double
    Dim expected As Double, jump As Double, baseline As Double
    Dim limit As Double

    ' 用一个新数组保存平滑后的净值,避免边平滑边判定(否则后一个点会基于已修改的前一个点判定)
    Dim newNav() As Double
    ReDim newNav(1 To n)
    Dim isValid() As Boolean
    ReDim isValid(1 To n)

    For i = 1 To n
        If IsNumeric(arr(i, navCol)) And Not IsEmpty(arr(i, navCol)) Then
            newNav(i) = CDbl(arr(i, navCol))
            isValid(i) = True
        Else
            isValid(i) = False
        End If
    Next i

    ' 对内部点(2 到 n-1)做判定,使用ORIGINAL值判断而非已修改值
    For i = 2 To n - 1
        If Not isValid(i) Then GoTo NextPoint
        If Not isValid(i - 1) Then GoTo NextPoint
        If Not isValid(i + 1) Then GoTo NextPoint

        prevV = CDbl(arr(i - 1, navCol))
        curV = CDbl(arr(i, navCol))
        nextV = CDbl(arr(i + 1, navCol))

        expected = (prevV + nextV) / 2
        jump = Abs(curV - expected)
        baseline = Abs(nextV - prevV)

        ' 判定阈值: max(baseline*K, prev*threshold)
        limit = baseline * K_RATIO
        If prevV * THRESHOLD > limit Then limit = prevV * THRESHOLD

        If jump > limit Then
            ' 异常点,替换
            newNav(i) = expected
        End If
NextPoint:
    Next i

    ' 把newNav写回arr
    For i = 1 To n
        If isValid(i) Then
            arr(i, navCol) = newNav(i)
        End If
    Next i
End Sub

'==============================================================
' 辅助过程: 基于净值序列重算30日年化收益率(单利公式,365天年)
'   - 公式: ((V_t / V_t-30 - 1) × (365 / 实际间隔)) × 100
'   - 匹配规则: 找日期-30,失败则向前匹配-31/-32/-33,最多4次
'   - 合并逻辑:
'       * 原始值有效(非空/非""/非0/非#N/A)且与计算值偏离<=阈值 → 保留原始值
'       * 原始值有效但偏离>阈值                              → 用计算值覆盖
'       * 原始值无效                                          → 用计算值填补
'       * 计算失败时,如果原始值有效则保留,否则写#N/A
'
' 参数:
'   arr        - 二维数组(已按日期升序排序)
'   dateCol    - 日期列
'   navCol     - 净值列
'   yieldCol   - 年化收益列
'   prodCode   - 当前产品编号(用于记录填补统计)
'   auditArr   - 审计数组: 处理方式/计算日期/基准日期/实际间隔(天)
'==============================================================
Private Sub Calc30DayAnnualYield(ByRef arr As Variant, _
                                  ByVal dateCol As Long, _
                                  ByVal navCol As Long, _
                                  ByVal yieldCol As Long, _
                                  ByVal prodCode As String, _
                                  ByRef auditArr() As Variant)
    Const TARGET_DAYS As Long = 30
    Const MAX_LOOKBACK As Long = 3
    Const DAYS_PER_YEAR As Double = 365#
    Const DEVIATION_THRESHOLD As Double = 0.5   ' 偏离阈值(百分点): 原始与计算差>0.5则用计算值
    Const FLAT_NAV_THRESHOLD As Double = 0.0001  ' 净值波动<1基点视为"不变"

    Dim n As Long: n = UBound(arr, 1)
    ReDim auditArr(1 To n, 1 To 4)
    If n < 2 Then
        auditArr(1, 1) = "无法计算（数据不足）"
        If IsDate(arr(1, dateCol)) Then
            auditArr(1, 2) = CDate(CLng(Fix(CDbl(CDate(arr(1, dateCol))))))
        End If
        Exit Sub
    End If

    '--- 0. 检测是否为"净值不变型"产品(净值=固定值,收益靠分红) ---
    Dim navMin As Double, navMax As Double
    Dim hasFirstNav As Boolean: hasFirstNav = False
    Dim ii As Long
    For ii = 1 To n
        If IsNumeric(arr(ii, navCol)) And Not IsEmpty(arr(ii, navCol)) Then
            Dim navVal As Double: navVal = CDbl(arr(ii, navCol))
            If navVal > 0 Then
                If Not hasFirstNav Then
                    navMin = navVal
                    navMax = navVal
                    hasFirstNav = True
                Else
                    If navVal < navMin Then navMin = navVal
                    If navVal > navMax Then navMax = navVal
                End If
            End If
        End If
    Next ii

    Dim isFlatNav As Boolean: isFlatNav = False
    If hasFirstNav Then
        If (navMax - navMin) < FLAT_NAV_THRESHOLD Then
            isFlatNav = True
        End If
    End If

    ' 净值不变型: 完全跳过计算,保留原始G列值不变
    ' (该产品的填补数记0)
    If isFlatNav Then
        For ii = 1 To n
            auditArr(ii, 1) = "净值不变型，不计算（保留原始值）"
            If IsDate(arr(ii, dateCol)) Then
                auditArr(ii, 2) = CDate(CLng(Fix(CDbl(CDate(arr(ii, dateCol))))))
            End If
        Next ii
        If Not g_fillCounts Is Nothing And Len(prodCode) > 0 Then
            g_fillCounts(prodCode) = 0
        End If
        Exit Sub
    End If

    '--- 1. 建立 日期(Long) -> 净值 的字典索引 ---
    Dim navDict As Object
    Set navDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim d As Date

    For i = 1 To n
        If IsDate(arr(i, dateCol)) And IsNumeric(arr(i, navCol)) Then
            d = CDate(arr(i, dateCol))
            If CDbl(arr(i, navCol)) > 0 Then
                navDict(CLng(Fix(CDbl(d)))) = CDbl(arr(i, navCol))
            End If
        End If
    Next i

    '--- 2. 逐行计算并合并 ---
    Dim curDate As Date, curNav As Double
    Dim curDateNum As Long, baseDateNum As Long, baseNav As Double
    Dim offset As Long, actualGap As Long
    Dim foundBase As Boolean
    Dim ratio As Double, calcYield As Double
    Dim hasCalc As Boolean
    Dim rawVal As Variant, rawValid As Boolean, rawNum As Double
    Dim filledCount As Long: filledCount = 0

    For i = 1 To n
        If IsDate(arr(i, dateCol)) Then
            auditArr(i, 2) = CDate(CLng(Fix(CDbl(CDate(arr(i, dateCol))))))
        End If

        '--- 2.1 判断原始值是否"有效" ---
        ' 有效定义: 非空 + 非"" + 非0 + 非#N/A错误
        rawValid = False
        rawNum = 0
        rawVal = arr(i, yieldCol)

        If Not IsError(rawVal) Then
            If Not IsEmpty(rawVal) Then
                If IsNumeric(rawVal) Then
                    rawNum = CDbl(rawVal)
                    If rawNum <> 0 Then
                        rawValid = True
                    End If
                End If
            End If
        End If
        ' IsError、IsEmpty、非数值、""、0 都视为无效

        '--- 2.2 尝试计算理论值 ---
        hasCalc = False
        calcYield = 0

        ' 当前行净值有效才能算
        If IsDate(arr(i, dateCol)) And IsNumeric(arr(i, navCol)) Then
            If CDbl(arr(i, navCol)) > 0 Then
                curDate = CDate(arr(i, dateCol))
                curDateNum = CLng(Fix(CDbl(curDate)))
                curNav = CDbl(arr(i, navCol))

                ' 查找 D-30, D-31, D-32, D-33
                foundBase = False
                For offset = 0 To MAX_LOOKBACK
                    baseDateNum = curDateNum - TARGET_DAYS - offset
                    If navDict.Exists(baseDateNum) Then
                        baseNav = navDict(baseDateNum)
                        actualGap = curDateNum - baseDateNum
                        auditArr(i, 3) = CDate(baseDateNum)
                        auditArr(i, 4) = actualGap
                        foundBase = True
                        Exit For
                    End If
                Next offset

                If foundBase Then
                    ratio = curNav / baseNav
                    If ratio > 0 Then
                        calcYield = ((ratio - 1) * (DAYS_PER_YEAR / actualGap)) * 100
                        hasCalc = True
                    End If
                End If
            End If
        End If

        '--- 2.3 合并逻辑 ---
        If rawValid And hasCalc Then
            ' 都有: 看偏离
            If Abs(rawNum - calcYield) > DEVIATION_THRESHOLD Then
                ' 偏离过大,用计算值覆盖
                arr(i, yieldCol) = calcYield
                filledCount = filledCount + 1
                auditArr(i, 1) = "计算覆盖原始值"
            Else
                ' 偏离合理,保留原始
                arr(i, yieldCol) = rawNum
                auditArr(i, 1) = "原始值保留（计算校验通过）"
            End If
        ElseIf rawValid And Not hasCalc Then
            ' 只有原始: 保留
            arr(i, yieldCol) = rawNum
            auditArr(i, 1) = "原始值保留（无法计算）"
        ElseIf Not rawValid And hasCalc Then
            ' 只有计算: 填补
            arr(i, yieldCol) = calcYield
            filledCount = filledCount + 1
            auditArr(i, 1) = "计算填补"
        Else
            ' 都没有: #N/A
            arr(i, yieldCol) = CVErr(xlErrNA)
            auditArr(i, 1) = "无法计算（无有效基准或净值）"
        End If
    Next i

    '--- 3. 记录该产品的填补条数 ---
    If Not g_fillCounts Is Nothing And Len(prodCode) > 0 Then
        g_fillCounts(prodCode) = filledCount
    End If
End Sub
'==============================================================
' 辅助过程: 在导出工作簿中新增“数据摘要”sheet
'   - 合并原02_check_data的数据完整度报告
'   - 同步写入30日年化收益率计算填补条数
'   - 不再创建、清空或写入数据库工作簿的Sheet3
'==============================================================
Private Sub WriteDataSummarySheet(ByVal wbOut As Workbook, ByVal dimDict As Object, ByVal groupDict As Object, ByVal fillDict As Object, _
                                  ByVal exportedDict As Object, ByVal emptyDict As Object, _
                                  ByVal missingDict As Object, ByVal disabledDict As Object)
    Dim wsRpt As Worksheet
    Set wsRpt = wbOut.Worksheets.Add(Before:=wbOut.Worksheets(1))
    wsRpt.name = SHEET_DATA_SUMMARY

    Dim headers As Variant
    headers = Array("产品编号", "产品简称", _
                    "净值日期最小值", "净值日期最大值", _
                    "净值最小值", "净值最大值", _
                    "数据条数", "总跨度(天)", _
                    "最大连续缺失天数", _
                    "首日净值", "末日净值", "累计涨跌幅(%)", _
                    "计算填补条数", "导出状态")

    Dim nCols As Long: nCols = UBound(headers) + 1
    wsRpt.Range(wsRpt.Cells(1, 1), wsRpt.Cells(1, nCols)).value = headers
    wsRpt.Range(wsRpt.Cells(1, 1), wsRpt.Cells(1, nCols)).Font.Bold = True

    If groupDict Is Nothing Then
        wsRpt.Columns("A:N").AutoFit
        Exit Sub
    End If

    If groupDict.Count = 0 Then
        wsRpt.Columns("A:N").AutoFit
        Exit Sub
    End If

    Dim codes() As String
    ReDim codes(0 To groupDict.Count - 1)

    Dim idx As Long: idx = 0
    Dim k As Variant
    For Each k In groupDict.keys
        codes(idx) = CStr(k)
        idx = idx + 1
    Next k
    SortStringArray codes

    Dim outArr() As Variant
    ReDim outArr(1 To groupDict.Count, 1 To nCols)

    Dim outRow As Long: outRow = 0
    Dim p As Long
    For p = LBound(codes) To UBound(codes)
        Dim code As String: code = codes(p)
        Dim subDict As Object
        Set subDict = groupDict(code)

        Dim validCount As Long: validCount = 0
        Dim pkKey As Variant
        For Each pkKey In subDict.keys
            Dim rowArr As Variant
            rowArr = subDict(pkKey)
            If IsDate(rowArr(1)) And IsNumeric(rowArr(4)) Then
                If CDbl(rowArr(4)) > 0 Then validCount = validCount + 1
            End If
        Next pkKey

        If validCount = 0 Then GoTo NextCode

        Dim dArr() As Long
        Dim vArr() As Double
        ReDim dArr(1 To validCount)
        ReDim vArr(1 To validCount)

        idx = 0
        For Each pkKey In subDict.keys
            rowArr = subDict(pkKey)
            If IsDate(rowArr(1)) And IsNumeric(rowArr(4)) Then
                If CDbl(rowArr(4)) > 0 Then
                    idx = idx + 1
                    dArr(idx) = CLng(Fix(CDbl(CDate(rowArr(1)))))
                    vArr(idx) = CDbl(rowArr(4))
                End If
            End If
        Next pkKey

        SortByDate dArr, vArr

        Dim minDate As Long, maxDate As Long
        minDate = dArr(1)
        maxDate = dArr(validCount)

        Dim minNav As Double, maxNav As Double
        minNav = vArr(1)
        maxNav = vArr(1)

        Dim i As Long
        For i = 2 To validCount
            If vArr(i) < minNav Then minNav = vArr(i)
            If vArr(i) > maxNav Then maxNav = vArr(i)
        Next i

        Dim maxGap As Long: maxGap = 0
        For i = 2 To validCount
            Dim gap As Long
            gap = dArr(i) - dArr(i - 1) - 1
            If gap > maxGap Then maxGap = gap
        Next i

        Dim firstNav As Double: firstNav = vArr(1)
        Dim lastNav As Double: lastNav = vArr(validCount)
        Dim retPct As Double
        If firstNav > 0 Then retPct = (lastNav / firstNav - 1) * 100

        Dim prodShort As String: prodShort = vbNullString
        If Not dimDict Is Nothing Then
            If dimDict.Exists(code) Then prodShort = CStr(dimDict(code))
        End If

        Dim filledCount As Long: filledCount = 0
        If Not fillDict Is Nothing Then
            If fillDict.Exists(code) Then filledCount = CLng(fillDict(code))
        End If

        outRow = outRow + 1
        outArr(outRow, 1) = code
        outArr(outRow, 2) = prodShort
        outArr(outRow, 3) = CDate(minDate)
        outArr(outRow, 4) = CDate(maxDate)
        outArr(outRow, 5) = minNav
        outArr(outRow, 6) = maxNav
        outArr(outRow, 7) = validCount
        outArr(outRow, 8) = maxDate - minDate
        outArr(outRow, 9) = maxGap
        outArr(outRow, 10) = firstNav
        outArr(outRow, 11) = lastNav
        outArr(outRow, 12) = retPct
        outArr(outRow, 13) = filledCount
        If Not dimDict.Exists(code) Then
            outArr(outRow, 14) = "产品信息未配置（未导出）"
        ElseIf Not exportedDict Is Nothing And exportedDict.Exists(code) Then
            outArr(outRow, 14) = "已导出"
        ElseIf Not disabledDict Is Nothing And disabledDict.Exists(code) Then
            outArr(outRow, 14) = "图表导出未启用（未导出）"
        Else
            outArr(outRow, 14) = "未导出"
        End If
NextCode:
    Next p

    If outRow > 0 Then
        Dim finalArr() As Variant
        ReDim finalArr(1 To outRow, 1 To nCols)

        Dim r As Long, c As Long
        For r = 1 To outRow
            For c = 1 To nCols
                finalArr(r, c) = outArr(r, c)
            Next c
        Next r

        wsRpt.Range(wsRpt.Cells(2, 1), wsRpt.Cells(1 + outRow, nCols)).value = finalArr
        wsRpt.Range(wsRpt.Cells(2, 3), wsRpt.Cells(1 + outRow, 4)).NumberFormat = "yyyy-mm-dd"
        wsRpt.Range(wsRpt.Cells(2, 5), wsRpt.Cells(1 + outRow, 6)).NumberFormat = "0.0000"
        wsRpt.Range(wsRpt.Cells(2, 10), wsRpt.Cells(1 + outRow, 11)).NumberFormat = "0.0000"
        wsRpt.Range(wsRpt.Cells(2, 12), wsRpt.Cells(1 + outRow, 12)).NumberFormat = "0.00"
    End If

    Dim nextRow As Long
    nextRow = outRow + 4
    If nextRow < 4 Then nextRow = 4
    WriteExceptionSection wsRpt, nextRow, "产品信息有、绘图净值数据无（仅记录，不影响导出）", emptyDict
    WriteExceptionSection wsRpt, nextRow, "绘图净值数据有、产品信息未配置（未导出）", missingDict
    WriteExceptionSection wsRpt, nextRow, "绘图净值数据有、图表导出未启用（未导出）", disabledDict

    wsRpt.Columns("A:N").AutoFit
End Sub

'==============================================================
' 辅助过程: 在“数据摘要”中写入导出异常清单
'==============================================================
Private Sub WriteExceptionSection(ByVal wsRpt As Worksheet, ByRef nextRow As Long, _
                                  ByVal titleText As String, ByVal itemDict As Object)
    Dim itemCount As Long
    If Not itemDict Is Nothing Then itemCount = itemDict.Count

    wsRpt.Cells(nextRow, 1).value = titleText & "：" & itemCount & " 个"
    wsRpt.Cells(nextRow, 1).Font.Bold = True
    nextRow = nextRow + 1

    wsRpt.Cells(nextRow, 1).value = "产品编号"
    wsRpt.Cells(nextRow, 2).value = "产品简称"
    wsRpt.Cells(nextRow, 1).Resize(1, 2).Font.Bold = True
    nextRow = nextRow + 1

    If itemCount > 0 Then
        Dim codes() As String
        ReDim codes(0 To itemCount - 1)

        Dim idx As Long: idx = 0
        Dim codeKey As Variant
        For Each codeKey In itemDict.keys
            codes(idx) = CStr(codeKey)
            idx = idx + 1
        Next codeKey
        SortStringArray codes

        Dim outArr() As Variant
        ReDim outArr(1 To itemCount, 1 To 2)
        Dim i As Long
        For i = LBound(codes) To UBound(codes)
            outArr(i + 1, 1) = codes(i)
            outArr(i + 1, 2) = itemDict(codes(i))
        Next i
        wsRpt.Cells(nextRow, 1).Resize(itemCount, 2).value = outArr
        nextRow = nextRow + itemCount
    End If

    nextRow = nextRow + 1
End Sub


'==============================================================
' 辅助过程: 字符串数组升序排序
'==============================================================
Private Sub SortStringArray(ByRef arr() As String)
    Dim n As Long: n = UBound(arr) - LBound(arr) + 1
    If n < 2 Then Exit Sub

    Dim i As Long, j As Long, tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = LBound(arr) To UBound(arr) - 1 - (i - LBound(arr))
            If arr(j) > arr(j + 1) Then
                tmp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = tmp
            End If
        Next j
    Next i
End Sub

'==============================================================
' 辅助过程: 按日期数组联合排序
'==============================================================
Private Sub SortByDate(ByRef dArr() As Long, ByRef vArr() As Double)
    Dim n As Long: n = UBound(dArr)
    If n < 2 Then Exit Sub

    Dim i As Long, j As Long
    Dim tmpD As Long, tmpV As Double
    For i = 1 To n - 1
        For j = 1 To n - i
            If dArr(j) > dArr(j + 1) Then
                tmpD = dArr(j): dArr(j) = dArr(j + 1): dArr(j + 1) = tmpD
                tmpV = vArr(j): vArr(j) = vArr(j + 1): vArr(j + 1) = tmpV
            End If
        Next j
    Next i
End Sub

'==============================================================
' 辅助函数: 构建表头 -> 列号映射
'==============================================================
Private Function BuildHeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim c As Long
    For c = 1 To lastCol
        Dim headerText As String
        headerText = NormalizeText(ws.Cells(headerRow, c).value)
        If Len(headerText) > 0 Then
            If Not dict.Exists(headerText) Then dict.Add headerText, c
        End If
    Next c

    Set BuildHeaderMap = dict
End Function

'==============================================================
' 辅助函数: 从候选字段名中返回第一个存在的列号
'==============================================================
Private Function FindFirstExistingHeader(ByVal headerMap As Object, ByVal candidates As Variant) As Long
    If headerMap Is Nothing Then Exit Function

    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If headerMap.Exists(CStr(candidates(i))) Then
            FindFirstExistingHeader = CLng(headerMap(CStr(candidates(i))))
            Exit Function
        End If
    Next i
End Function

'==============================================================
' 辅助函数: 文本规范化
'==============================================================
Private Function NormalizeText(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Then
        NormalizeText = vbNullString
        Exit Function
    End If

    Dim textValue As String
    textValue = CStr(value)
    textValue = Replace(textValue, ChrW$(12288), " ")
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    NormalizeText = WorksheetFunction.Trim(textValue)
End Function

'==============================================================
' 辅助函数: 判断“图表导出”配置是否启用
'==============================================================
Private Function IsChartExportEnabled(ByVal value As Variant) As Boolean
    Dim configValue As String
    configValue = UCase$(NormalizeText(value))

    Select Case configValue
        Case "是", "Y", "YES", "1", "TRUE", "启用"
            IsChartExportEnabled = True
    End Select
End Function

'==============================================================
' 辅助函数: 最后一列
'==============================================================
Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = foundCell.Column
    End If
End Function

