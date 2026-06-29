Option Explicit

'==============================================================
' 模块: 导入净值数据
' 功能: 扫描当前文件夹下所有[净值数据浏览表 yyyy-mm-dd至yyyy-mm-dd.xlsx]
'       将数据按主键(日期+产品编号)增量合并到“绘图净值数据”
'==============================================================

Private Const TARGET_SHEET_NAME As String = "绘图净值数据"
Private Const INTERNAL_SOURCE_SHEET_NAME As String = "上层产品净值数据(181)"
Private Const SOURCE_FILE_DESC As String = "净值数据浏览表 yyyy-mm-dd至yyyy-mm-dd.xlsx"

Private Const INTERNAL_COL_ACCOUNT_NAME As String = "账套名称"
Private Const INTERNAL_COL_NAV_DATE As String = "日期"
Private Const INTERNAL_COL_UNIT_NAV As String = "单位净值"
Private Const INTERNAL_COL_PRODUCT_CODE As String = "信托计划代码"

Public Sub Chart01_ImportNavData()
    
    Dim t0 As Double: t0 = Timer

    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox("运行前请确认：" & vbCrLf & vbCrLf & _
                           "1. 已经下载或导出“净值表”的数据。" & vbCrLf & _
                           "2. 净值表文件已保存到当前数据库工作簿同级目录。" & vbCrLf & _
                           "3. 文件名保持系统默认格式：" & SOURCE_FILE_DESC & vbCrLf & vbCrLf & _
                           "确认后开始导入。", _
                           vbQuestion + vbYesNo + vbDefaultButton2, "绘图净值数据导入")
    If confirmResult <> vbYes Then Exit Sub
    
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "绘图净值数据导入无法继续" & vbCrLf & vbCrLf & _
               "错误信息：当前数据库工作簿尚未保存，请先保存后再运行导入。", vbExclamation, "绘图净值数据导入"
        Exit Sub
    End If
    
    '--- 1. 准备环境 ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim wsDB As Worksheet: Set wsDB = wbDB.Worksheets(TARGET_SHEET_NAME)
    Dim folderPath As String: folderPath = wbDB.Path & "\"
    
    '--- 2. 扫描文件夹,正则匹配文件名,收集文件信息 ---
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "净值数据浏览表\s+(\d{4}-\d{2}-\d{2})至(\d{4}-\d{2}-\d{2})\.xlsx$"
    regex.IgnoreCase = True
    
    '用Dictionary存储: key=开始日期(用于排序), value=文件全路径
    Dim fileDict As Object
    Set fileDict = CreateObject("Scripting.Dictionary")
    
    Dim fileName As String
    fileName = Dir(folderPath & "净值数据浏览表 *.xlsx")
    
    Dim matches As Object, m As Object
    Dim dStart As Date, dEnd As Date
    Dim sortKey As String
    
    Do While Len(fileName) > 0
        If regex.Test(fileName) Then
            Set matches = regex.Execute(fileName)
            Set m = matches(0)
            dStart = CDate(m.SubMatches(0))
            dEnd = CDate(m.SubMatches(1))
            
            If dStart <= dEnd Then
                ' 排序键: 开始日期+文件名,确保唯一且按日期升序
                sortKey = Format(dStart, "yyyy-mm-dd") & "|" & fileName
                fileDict.Add sortKey, folderPath & fileName
            End If
        End If
        fileName = Dir()
    Loop
    
    If fileDict.Count = 0 Then
        ImportFromInternalNav181 wsDB, t0, folderPath
        GoTo CleanUp
    End If
    
    ' 按key升序排序文件列表
    Dim sortedKeys() As String
    sortedKeys = SortKeys(fileDict.keys)
    
    '--- 3. 读取数据库“绘图净值数据”现有数据到内存,建立主键索引 ---
    Dim dbDict As Object
    Set dbDict = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long, lastCol As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    lastCol = 10  ' A-J共10列
    
    Dim dbData As Variant
    Dim hasExistingData As Boolean
    
    If lastRow >= 2 Then
        hasExistingData = True
        dbData = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRow, lastCol)).value
        
        Dim i As Long, pk As String
        For i = 1 To UBound(dbData, 1)
            pk = BuildKey(dbData(i, 1), dbData(i, 2))
            If Len(pk) > 0 Then
                ' value存储行号(在dbData数组中的索引)
                dbDict(pk) = i
            End If
        Next i
    Else
        hasExistingData = False
    End If
    
    '--- 4. 依次打开每个净值数据浏览表,合并数据 ---
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim srcLastRow As Long
    Dim srcStartRow As Long
    Dim srcData As Variant
    Dim newRows As Object  ' 存储新增行(主键不存在的)
    Set newRows = CreateObject("Scripting.Dictionary")
    
    Dim updatedCount As Long, insertedCount As Long
    updatedCount = 0
    insertedCount = 0
    
    Dim k As Variant, filePath As String
    For Each k In sortedKeys
        filePath = fileDict(k)
        
        On Error Resume Next
        Set wbSrc = Workbooks.Open(fileName:=filePath, ReadOnly:=True, UpdateLinks:=0)
        If wbSrc Is Nothing Then
            MsgBox "绘图净值数据导入警告" & vbCrLf & vbCrLf & _
                   "错误信息：无法打开文件，可能被占用，已跳过。" & vbCrLf & _
                   "文件：" & filePath, vbExclamation, "绘图净值数据导入"
            On Error GoTo 0
            GoTo NextFile
        End If
        On Error GoTo 0
        
        Set wsSrc = wbSrc.Worksheets(1)
        srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
        srcStartRow = FindSourceDataStartRow(wsSrc)
        
        If srcStartRow > 0 And srcLastRow >= srcStartRow Then
            srcData = wsSrc.Range(wsSrc.Cells(srcStartRow, 1), wsSrc.Cells(srcLastRow, 10)).value
            
            Dim j As Long, srcKey As String
            For j = 1 To UBound(srcData, 1)
                srcKey = BuildKey(srcData(j, 1), srcData(j, 2))
                If Len(srcKey) > 0 Then
                    If dbDict.Exists(srcKey) Then
                        ' 覆盖更新内存中的数据
                        Dim rowIdx As Long
                        rowIdx = dbDict(srcKey)
                        Dim c As Long
                        For c = 1 To 10
                            dbData(rowIdx, c) = srcData(j, c)
                        Next c
                        updatedCount = updatedCount + 1
                    Else
                        ' 新增行,先缓存
                        Dim newRow(1 To 10) As Variant
                        For c = 1 To 10
                            newRow(c) = srcData(j, c)
                        Next c
                        newRows(srcKey) = newRow
                        ' 同时加入dbDict,防止同批次重复
                        dbDict(srcKey) = -1  ' -1表示在newRows里
                        insertedCount = insertedCount + 1
                    End If
                End If
            Next j
        End If
        
        wbSrc.Close SaveChanges:=False
        Set wbSrc = Nothing
NextFile:
    Next k
    
    '--- 5. 一次性写回“绘图净值数据” ---
    ' 5.1 先写回更新过的现有数据
    If hasExistingData And UBound(dbData, 1) > 0 Then
        wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(1 + UBound(dbData, 1), 10)).value = dbData
    End If
    
    ' 5.2 再追加新增数据
    If newRows.Count > 0 Then
        Dim writeArr() As Variant
        ReDim writeArr(1 To newRows.Count, 1 To 10)
        Dim idx As Long: idx = 0
        Dim key As Variant
        For Each key In newRows.keys
            idx = idx + 1
            Dim arr As Variant: arr = newRows(key)
            For c = 1 To 10
                writeArr(idx, c) = arr(c)
            Next c
        Next key
        
        Dim writeStartRow As Long
        writeStartRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row + 1
        If writeStartRow < 2 Then writeStartRow = 2
        
        wsDB.Range(wsDB.Cells(writeStartRow, 1), _
                   wsDB.Cells(writeStartRow + newRows.Count - 1, 10)).value = writeArr
    End If
    
    ' 5.3 按A列日期升序排序(可选)
    Dim finalLastRow As Long
    finalLastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    If finalLastRow >= 3 Then
        wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(finalLastRow, 10)).Sort _
            Key1:=wsDB.Cells(2, 1), Order1:=xlAscending, _
            Key2:=wsDB.Cells(2, 2), Order2:=xlAscending, _
            header:=xlNo
    End If
    
    MsgBox "绘图净值数据导入完成" & vbCrLf & vbCrLf & _
           "处理范围：" & folderPath & vbCrLf & _
           "耗时：" & Format(Timer - t0, "0.00") & " 秒" & vbCrLf & vbCrLf & _
           "处理结果：" & vbCrLf & _
           "符合要求净值表数：" & fileDict.Count & vbCrLf & _
           "覆盖更新行数：" & updatedCount & vbCrLf & _
           "新增行数：" & insertedCount, _
           vbInformation, "绘图净值数据导入"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


'==============================================================
' 辅助过程: 未找到外部净值表时,从“上层产品净值数据(181)”补充导入
'==============================================================
Private Sub ImportFromInternalNav181(ByVal wsDB As Worksheet, ByVal t0 As Double, ByVal folderPath As String)
    Dim wsSource As Worksheet
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(INTERNAL_SOURCE_SHEET_NAME)
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "绘图净值数据导入无法继续" & vbCrLf & vbCrLf & _
               "错误信息：未找到任何[" & SOURCE_FILE_DESC & "]文件，也未找到内部sheet：" & INTERNAL_SOURCE_SHEET_NAME & vbCrLf & _
               "处理范围：" & folderPath, vbExclamation, "绘图净值数据导入"
        Exit Sub
    End If
    On Error GoTo InternalFail

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(wsSource, 1)

    Dim sourceNameCol As Long
    Dim sourceDateCol As Long
    Dim sourceNavCol As Long
    Dim sourceCodeCol As Long

    sourceNameCol = RequireHeader(headerMap, INTERNAL_COL_ACCOUNT_NAME, INTERNAL_SOURCE_SHEET_NAME)
    sourceDateCol = RequireHeader(headerMap, INTERNAL_COL_NAV_DATE, INTERNAL_SOURCE_SHEET_NAME)
    sourceNavCol = RequireHeader(headerMap, INTERNAL_COL_UNIT_NAV, INTERNAL_SOURCE_SHEET_NAME)
    sourceCodeCol = RequireHeader(headerMap, INTERNAL_COL_PRODUCT_CODE, INTERNAL_SOURCE_SHEET_NAME)

    Dim dbDict As Object
    Set dbDict = CreateObject("Scripting.Dictionary")
    dbDict.CompareMode = vbTextCompare

    Dim targetProductSet As Object
    Set targetProductSet = CreateObject("Scripting.Dictionary")
    targetProductSet.CompareMode = vbTextCompare

    Dim targetProductName As Object
    Set targetProductName = CreateObject("Scripting.Dictionary")
    targetProductName.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row

    Dim r As Long
    Dim pk As String
    If lastRow >= 2 Then
        Dim dbData As Variant
        dbData = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRow, 3)).Value

        For r = 1 To UBound(dbData, 1)
            Dim targetCode As String
            targetCode = NormalizeText(dbData(r, 2))
            If Len(targetCode) > 0 Then
                targetProductSet(targetCode) = True
                If Len(NormalizeText(dbData(r, 3))) > 0 Then targetProductName(targetCode) = NormalizeText(dbData(r, 3))
            End If

            pk = BuildNormalizedKey(dbData(r, 1), dbData(r, 2))
            If Len(pk) > 0 Then dbDict(pk) = True
        Next r
    End If

    If targetProductSet.Count = 0 Then
        MsgBox "绘图净值数据导入无法继续" & vbCrLf & vbCrLf & _
               "错误信息：未找到外部净值表，且“" & TARGET_SHEET_NAME & "”中没有已有产品编号，无法从内部sheet补充缺失日期。", _
               vbExclamation, "绘图净值数据导入"
        Exit Sub
    End If

    Dim newRows As Object
    Set newRows = CreateObject("Scripting.Dictionary")
    newRows.CompareMode = vbTextCompare

    Dim sourceLastRow As Long
    sourceLastRow = LastUsedRow(wsSource)

    Dim productCode As String
    Dim productName As String
    Dim navDate As Date
    Dim navValue As Variant
    Dim newRow(1 To 10) As Variant
    Dim insertedCount As Long
    Dim skippedExistingCount As Long
    Dim invalidDateCount As Long
    Dim invalidNavCount As Long
    Dim zeroNavCount As Long
    Dim duplicateSourceCount As Long
    Dim skippedNoTargetProductCount As Long

    For r = 2 To sourceLastRow
        productCode = NormalizeText(wsSource.Cells(r, sourceCodeCol).Value)
        If Len(productCode) = 0 Then GoTo ContinueSourceRow

        If Not targetProductSet.Exists(productCode) Then
            skippedNoTargetProductCount = skippedNoTargetProductCount + 1
            GoTo ContinueSourceRow
        End If

        If Not TryParseDateValue(wsSource.Cells(r, sourceDateCol).Value, navDate) Then
            invalidDateCount = invalidDateCount + 1
            GoTo ContinueSourceRow
        End If

        pk = Format$(navDate, "yyyy-mm-dd") & "|" & productCode
        If dbDict.Exists(pk) Then
            skippedExistingCount = skippedExistingCount + 1
            GoTo ContinueSourceRow
        End If

        navValue = wsSource.Cells(r, sourceNavCol).Value
        If Not IsNumeric(navValue) Then
            invalidNavCount = invalidNavCount + 1
            GoTo ContinueSourceRow
        End If
        If CDbl(navValue) = 0 Then
            zeroNavCount = zeroNavCount + 1
            GoTo ContinueSourceRow
        End If

        productName = vbNullString
        If targetProductName.Exists(productCode) Then productName = CStr(targetProductName(productCode))
        If Len(productName) = 0 Then productName = NormalizeText(wsSource.Cells(r, sourceNameCol).Value)

        newRow(1) = navDate
        newRow(2) = productCode
        newRow(3) = productName
        newRow(4) = CDbl(navValue)
        newRow(5) = vbNullString
        newRow(6) = vbNullString
        newRow(7) = vbNullString
        newRow(8) = vbNullString
        newRow(9) = vbNullString
        newRow(10) = vbNullString

        If newRows.Exists(pk) Then duplicateSourceCount = duplicateSourceCount + 1
        newRows(pk) = newRow
        dbDict(pk) = True

ContinueSourceRow:
    Next r

    If newRows.Count > 0 Then
        Dim writeArr() As Variant
        ReDim writeArr(1 To newRows.Count, 1 To 10)

        Dim sortedKeys() As String
        sortedKeys = SortKeys(newRows.keys)

        Dim idx As Long
        Dim key As Variant
        Dim arr As Variant
        Dim c As Long

        For idx = LBound(sortedKeys) To UBound(sortedKeys)
            key = sortedKeys(idx)
            arr = newRows(key)
            For c = 1 To 10
                writeArr(idx + 1, c) = arr(c)
            Next c
        Next idx

        Dim writeStartRow As Long
        writeStartRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row + 1
        If writeStartRow < 2 Then writeStartRow = 2

        wsDB.Range(wsDB.Cells(writeStartRow, 1), _
                   wsDB.Cells(writeStartRow + newRows.Count - 1, 10)).Value = writeArr
        wsDB.Range(wsDB.Cells(writeStartRow, 1), _
                   wsDB.Cells(writeStartRow + newRows.Count - 1, 1)).NumberFormat = "yyyy-mm-dd"
        wsDB.Range(wsDB.Cells(writeStartRow, 4), _
                   wsDB.Cells(writeStartRow + newRows.Count - 1, 4)).NumberFormat = "0.0000"

        insertedCount = newRows.Count
    End If

    Dim finalLastRow As Long
    finalLastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    If finalLastRow >= 3 Then
        wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(finalLastRow, 10)).Sort _
            Key1:=wsDB.Cells(2, 1), Order1:=xlAscending, _
            Key2:=wsDB.Cells(2, 2), Order2:=xlAscending, _
            header:=xlNo
    End If

    MsgBox "绘图净值数据导入完成" & vbCrLf & vbCrLf & _
           "处理范围：" & folderPath & vbCrLf & _
           "内部sheet：" & INTERNAL_SOURCE_SHEET_NAME & vbCrLf & _
           "耗时：" & Format(Timer - t0, "0.00") & " 秒" & vbCrLf & vbCrLf & _
           "处理结果：" & vbCrLf & _
           "新增行数：" & insertedCount & vbCrLf & _
           "源数据产品编号未存在于目标表行数：" & skippedNoTargetProductCount & vbCrLf & _
           "目标已存在产品日期行数：" & skippedExistingCount & vbCrLf & _
           "源数据重复产品日期行数：" & duplicateSourceCount & vbCrLf & _
           "无效日期行数：" & invalidDateCount & vbCrLf & _
           "无效净值行数：" & invalidNavCount & vbCrLf & _
           "净值为0行数：" & zeroNavCount & vbCrLf & vbCrLf & _
           "注意事项：" & vbCrLf & _
           "未找到任何[" & SOURCE_FILE_DESC & "]文件，已改从内部sheet导入。", _
           vbInformation, "绘图净值数据导入"
    Exit Sub

InternalFail:
    MsgBox "绘图净值数据导入失败" & vbCrLf & vbCrLf & _
           "错误信息：未找到外部净值表，改从内部sheet导入时失败：" & Err.Description & vbCrLf & _
           "内部sheet：" & INTERNAL_SOURCE_SHEET_NAME, vbCritical, "绘图净值数据导入"
End Sub


'==============================================================
' 辅助函数: 构造主键 = yyyy-mm-dd|产品编号
'==============================================================
Private Function BuildKey(ByVal dateVal As Variant, ByVal codeVal As Variant) As String
    On Error Resume Next
    If IsEmpty(dateVal) Or IsNull(dateVal) Then Exit Function
    If Len(Trim(CStr(codeVal))) = 0 Then Exit Function
    
    Dim d As Date
    If IsDate(dateVal) Then
        d = CDate(dateVal)
        BuildKey = Format(d, "yyyy-mm-dd") & "|" & Trim(CStr(codeVal))
    Else
        ' 日期解析失败,用原始字符串
        BuildKey = Trim(CStr(dateVal)) & "|" & Trim(CStr(codeVal))
    End If
End Function


'==============================================================
' 辅助函数: 构造标准主键,支持yyyymmdd和yyyy-mm-dd日期
'==============================================================
Private Function BuildNormalizedKey(ByVal dateVal As Variant, ByVal codeVal As Variant) As String
    If IsEmpty(dateVal) Or IsNull(dateVal) Then Exit Function
    If Len(NormalizeText(codeVal)) = 0 Then Exit Function

    Dim d As Date
    If TryParseDateValue(dateVal, d) Then
        BuildNormalizedKey = Format$(d, "yyyy-mm-dd") & "|" & NormalizeText(codeVal)
    End If
End Function


'==============================================================
' 辅助函数: 构建表头索引
'==============================================================
Private Function BuildHeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim c As Long
    Dim headerText As String

    For c = 1 To lastCol
        headerText = NormalizeText(ws.Cells(headerRow, c).Value)
        If Len(headerText) > 0 Then
            If Not dict.Exists(headerText) Then dict.Add headerText, c
        End If
    Next c

    Set BuildHeaderMap = dict
End Function


'==============================================================
' 辅助函数: 获取必要字段列号
'==============================================================
Private Function RequireHeader(ByVal headerMap As Object, ByVal headerName As String, ByVal sheetName As String) As Long
    If Not headerMap Is Nothing Then
        If headerMap.Exists(headerName) Then
            RequireHeader = CLng(headerMap(headerName))
            Exit Function
        End If
    End If

    Err.Raise vbObjectError + 4511, , sheetName & " 缺少必要字段：" & headerName
End Function


'==============================================================
' 辅助函数: 解析Excel日期、yyyy-mm-dd文本和yyyymmdd数字/文本
'==============================================================
Private Function TryParseDateValue(ByVal value As Variant, ByRef parsedDate As Date) As Boolean
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    On Error GoTo InvalidDate

    Dim textValue As String
    textValue = NormalizeText(value)
    If Len(textValue) = 0 Then Exit Function

    If Len(textValue) >= 10 And IsDate(Left$(textValue, 10)) Then
        parsedDate = CDate(Left$(textValue, 10))
        TryParseDateValue = True
        Exit Function
    End If

    If Len(textValue) = 8 And IsNumeric(textValue) Then
        parsedDate = DateSerial(CInt(Left$(textValue, 4)), CInt(Mid$(textValue, 5, 2)), CInt(Right$(textValue, 2)))
        If Format$(parsedDate, "yyyymmdd") <> textValue Then Exit Function
        TryParseDateValue = True
        Exit Function
    End If

    If IsDate(value) Then
        parsedDate = CDate(value)
        TryParseDateValue = True
    End If
    Exit Function

InvalidDate:
    TryParseDateValue = False
End Function


'==============================================================
' 辅助函数: 文本规范化
'==============================================================
Private Function NormalizeText(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then
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
' 辅助函数: 工作表最后使用行
'==============================================================
Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = foundCell.Row
    End If
End Function


'==============================================================
' 辅助函数: 工作表最后使用列
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


'==============================================================
' 辅助函数: 对字符串数组进行升序排序(简单冒泡,文件量不大够用)
'==============================================================
Private Function SortKeys(ByVal keys As Variant) As String()
    Dim arr() As String
    Dim n As Long: n = UBound(keys) - LBound(keys) + 1
    ReDim arr(0 To n - 1)
    
    Dim i As Long
    For i = 0 To n - 1
        arr(i) = CStr(keys(i))
    Next i
    
    Dim j As Long
    Dim tmp As String
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If arr(j) > arr(j + 1) Then
                tmp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = tmp
            End If
        Next j
    Next i
    
    SortKeys = arr
End Function
'==============================================================
' 辅助函数: 查找源表数据起始行
' 规则:
' 1) 优先在A列查找“净值日期”，找到后从下一行开始
' 2) 如果找不到，则从A列第一个可识别日期开始
' 3) 如果仍找不到，返回0
'==============================================================
Private Function FindSourceDataStartRow(ByVal ws As Worksheet) As Long
    Dim lastA As Long
    lastA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastA < 1 Then Exit Function
    
    '--- 1) 优先查找“净值日期” ---
    Dim headerCell As Range
    Set headerCell = ws.Columns("A").Find( _
        What:="净值日期", _
        After:=ws.Cells(ws.Rows.Count, "A"), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False _
    )
    
    If Not headerCell Is Nothing Then
        If headerCell.Row < lastA Then
            FindSourceDataStartRow = headerCell.Row + 1
            Exit Function
        End If
    End If
    
    '--- 2) 找不到“净值日期”时，查找A列第一个日期 ---
    Dim r As Long
    For r = 1 To lastA
        If IsValidSourceDate(ws.Cells(r, "A").value) Then
            FindSourceDataStartRow = r
            Exit Function
        End If
    Next r
End Function


'==============================================================
' 辅助函数: 判断是否为有效日期
' 这里限制为2000-01-01至2099-12-31，避免把普通数字误判为日期
'==============================================================
Private Function IsValidSourceDate(ByVal v As Variant) As Boolean
    On Error GoTo BadValue
    
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Len(Trim(CStr(v))) = 0 Then Exit Function
    
    If IsDate(v) Then
        Dim d As Date
        d = CDate(v)
        
        If d >= DateSerial(2000, 1, 1) And d <= DateSerial(2099, 12, 31) Then
            IsValidSourceDate = True
        End If
    End If
    
    Exit Function

BadValue:
    IsValidSourceDate = False
End Function


