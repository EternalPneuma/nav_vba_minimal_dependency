Option Explicit

'==============================================================
' 模块: 生成产品图表
' 功能: 在最新的[产品净值汇总_yyyymmdd.xlsx]中,
'       为每个产品sheet生成4个chart(净值/收益率 × 红/蓝)
'       - 收益率chart跳过开头连续的0值
'       - X轴强制为日期轴,只显示首尾日期
'       - 净值Y轴强制从1开始,收益率Y轴强制从0开始
'       - 只输出末尾数据标签
'==============================================================

' 模板文件目录：与 xlsm 同目录下的 assets\chart_templates
Private Const TPL_DIR As String = "assets\chart_templates\"

Private Const TPL_NAV_RED As String = "净值图表_红.crtx"
Private Const TPL_NAV_BLUE As String = "净值图表_蓝.crtx"
Private Const TPL_YIELD_RED As String = "收益率图表_红.crtx"
Private Const TPL_YIELD_BLUE As String = "收益率图表_蓝.crtx"

' chart布局参数(单位:磅)
Private Const CHART_LEFT_COL As String = "P"
Private Const CHART_WIDTH As Single = 480
Private Const CHART_HEIGHT_RED As Single = 280   ' 红色模板图表高度
Private Const CHART_HEIGHT_BLUE As Single = 280  ' 蓝色模板图表高度
Private Const CHART_GAP As Single = 20
Private Const SHEET_DATA_SUMMARY As String = "数据摘要"

Public Sub Chart03_GenerateCharts()
    
    Dim t0 As Double: t0 = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim dbPath As String: dbPath = wbDB.Path & "\"
    
    '--- 1. 检查4个模板文件 ---
    Dim tplPaths(1 To 4) As String
    tplPaths(1) = dbPath & TPL_DIR & TPL_NAV_RED
    tplPaths(2) = dbPath & TPL_DIR & TPL_NAV_BLUE
    tplPaths(3) = dbPath & TPL_DIR & TPL_YIELD_RED
    tplPaths(4) = dbPath & TPL_DIR & TPL_YIELD_BLUE
    
    Dim i As Long, missing As String
    For i = 1 To 4
        If Dir(tplPaths(i)) = "" Then
            missing = missing & vbCrLf & "  - " & Mid(tplPaths(i), InStrRev(tplPaths(i), "\") + 1)
        End If
    Next i
    If Len(missing) > 0 Then
        MsgBox "产品图表生成无法继续" & vbCrLf & vbCrLf & _
               "错误信息：缺少模板文件，请检查：" & missing, vbCritical, "产品图表生成"
        GoTo CleanUp
    End If
    
    '--- 2. 查找最新的[产品净值汇总_yyyymmdd.xlsx] ---
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^产品净值汇总_(\d{8})\.xlsx$"
    
    Dim fileName As String, latestFile As String, latestKey As String
    fileName = Dir(dbPath & "产品净值汇总_*.xlsx")
    Do While Len(fileName) > 0
        If regex.Test(fileName) Then
            Dim matches As Object
            Set matches = regex.Execute(fileName)
            If matches(0).SubMatches(0) > latestKey Then
                latestKey = matches(0).SubMatches(0)
                latestFile = fileName
            End If
        End If
        fileName = Dir()
    Loop
    
    If Len(latestFile) = 0 Then
        MsgBox "产品图表生成无法继续" & vbCrLf & vbCrLf & _
               "错误信息：未找到[产品净值汇总_yyyymmdd.xlsx]文件，请先运行[输出产品汇总]。", vbExclamation, "产品图表生成"
        GoTo CleanUp
    End If
    
    Dim targetPath As String: targetPath = dbPath & latestFile
    
    '--- 3. 检查目标文件是否已打开 ---
    Dim wbTarget As Workbook
    Dim wasOpen As Boolean: wasOpen = False
    
    On Error Resume Next
    Set wbTarget = Workbooks(latestFile)
    On Error GoTo 0
    
    If wbTarget Is Nothing Then
        Set wbTarget = Workbooks.Open(targetPath)
    Else
        wasOpen = True
    End If
    
    '--- 4. 遍历每个sheet,生成4个chart ---
    Dim ws As Worksheet
    Dim processedCount As Long: processedCount = 0
    Dim errSheets As String: errSheets = ""
    
    For Each ws In wbTarget.Worksheets
        If ws.name = SHEET_DATA_SUMMARY Then GoTo NextSheet
        
        ' 检查sheet是否有有效数据
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If lastRow < 2 Then GoTo NextSheet
        
        ' 先清除该sheet上已有的chart
        Dim co As ChartObject
        For Each co In ws.ChartObjects
            co.Delete
        Next co
        
        ' 净值的数据范围: 从第2行(首个数据行)到末行
        Dim navStartRow As Long: navStartRow = 2
        Dim dateRngNav As Range, navRng As Range
        Set dateRngNav = ws.Range("A" & navStartRow & ":A" & lastRow)
        Set navRng = ws.Range("D" & navStartRow & ":D" & lastRow)
        
        ' 收益率的数据范围: 跳过开头连续的0
        Dim yieldStartRow As Long
        yieldStartRow = FindFirstNonZeroRow(ws, "G", 2, lastRow)
        
        Dim dateRngYield As Range, yieldRng As Range
        Dim hasYieldData As Boolean: hasYieldData = False
        If yieldStartRow > 0 And yieldStartRow <= lastRow Then
            Set dateRngYield = ws.Range("A" & yieldStartRow & ":A" & lastRow)
            Set yieldRng = ws.Range("G" & yieldStartRow & ":G" & lastRow)
            hasYieldData = True
        End If
        
        Dim prodShort As String: prodShort = ws.name
        
        ' 计算4个chart的位置(2列2行)
        Dim baseLeft As Single, baseTop As Single
        baseLeft = ws.Range(CHART_LEFT_COL & "1").Left
        baseTop = ws.Range(CHART_LEFT_COL & "1").Top
        
        On Error GoTo SheetErr
        
        ' Chart 1: 净值-红 (左上)
        CreateChart ws, dateRngNav, navRng, tplPaths(1), _
            "chart_净值_红", prodShort & "单位净值表现", _
            baseLeft, baseTop, False, 1#, CHART_HEIGHT_RED
        
        ' Chart 2: 净值-蓝 (左下)
        CreateChart ws, dateRngNav, navRng, tplPaths(2), _
            "chart_净值_蓝", prodShort & "单位净值表现", _
            baseLeft, baseTop + CHART_HEIGHT_RED + CHART_GAP, False, 1#, CHART_HEIGHT_BLUE
    
        ' Chart 3 & 4: 收益率(强制Y轴下限为0)
        If hasYieldData Then
            ' Chart 3: 收益率-红 (右上)
            CreateChart ws, dateRngYield, yieldRng, tplPaths(3), _
                "chart_收益率_红", prodShort & "30日年化收益率" & vbLf & "(单位:%)", _
                baseLeft + CHART_WIDTH + CHART_GAP, baseTop, True, 0#, CHART_HEIGHT_RED
            
            ' Chart 4: 收益率-蓝 (右下)
            CreateChart ws, dateRngYield, yieldRng, tplPaths(4), _
                "chart_收益率_蓝", prodShort & "30日年化收益率" & vbLf & "(单位:%)", _
                baseLeft + CHART_WIDTH + CHART_GAP, baseTop + CHART_HEIGHT_RED + CHART_GAP, True, 0#, CHART_HEIGHT_BLUE
         End If
        
        processedCount = processedCount + 1
        On Error GoTo 0
        GoTo NextSheet

SheetErr:
        errSheets = errSheets & ws.name & "(" & Err.Description & "), "
        Err.Clear
        On Error GoTo 0

NextSheet:
    Next ws
    
    '--- 5. 保存文件 ---
    wbTarget.Save
    If Not wasOpen Then wbTarget.Close SaveChanges:=False
    
    '--- 6. 汇总提示 ---
    Dim msg As String
    msg = "产品图表生成完成" & vbCrLf & vbCrLf & _
          "目标文件：" & latestFile & vbCrLf & _
          "耗时：" & Format(Timer - t0, "0.00") & " 秒" & vbCrLf & vbCrLf & _
          "处理结果：" & vbCrLf & _
          "处理sheet数：" & processedCount
    If Len(errSheets) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "注意事项：" & vbCrLf & _
              "以下sheet处理出错：" & vbCrLf & Left(errSheets, Len(errSheets) - 2)
    End If
    
    MsgBox msg, vbInformation, "产品图表生成"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


'==============================================================
' 辅助函数: 找到指定列从startRow到endRow中,第一个非0(且非空)的行号
' 如果全部为0或空,返回0
'==============================================================
Private Function FindFirstNonZeroRow(ByVal ws As Worksheet, ByVal col As String, _
                                      ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long
    Dim v As Variant
    For r = startRow To endRow
        v = ws.Range(col & r).value
        If Not IsEmpty(v) And IsNumeric(v) Then
            If CDbl(v) <> 0 Then
                FindFirstNonZeroRow = r
                Exit Function
            End If
        End If
    Next r
    FindFirstNonZeroRow = 0  ' 全0或全空
End Function

'==============================================================
' 辅助过程: 创建一个chart并应用模板
'==============================================================
Private Sub CreateChart(ByVal ws As Worksheet, _
                       ByVal xRange As Range, ByVal yRange As Range, _
                       ByVal tplPath As String, _
                       ByVal chartName As String, ByVal title As String, _
                       ByVal leftPos As Single, ByVal topPos As Single, _
                       Optional ByVal isYieldChart As Boolean = False, _
                       Optional ByVal fixedYMin As Variant, _
                       Optional ByVal chartHeight As Single = 280)
    
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=leftPos, Top:=topPos, _
                                  Width:=CHART_WIDTH, Height:=chartHeight)
    co.name = chartName
    
    Dim ch As Chart
    Set ch = co.Chart
    
    ' 清空默认系列
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    
    ' 添加系列
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.values = yRange
    s.XValues = xRange
    s.name = title
    
    ' 应用模板
    On Error Resume Next
    ch.ApplyChartTemplate tplPath
    On Error GoTo 0
    
    ' 重新拿系列引用(ApplyChartTemplate可能让原引用失效)
    Set s = ch.SeriesCollection(1)
    
    ' 设置标题
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    
    ' 标题文字右对齐(短行向右靠拢,长行视觉上保持居中)
    On Error Resume Next
    ch.ChartTitle.Format.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
    On Error GoTo 0
    
    ' --- 强制X轴为日期轴,只显示首尾日期 ---
    Dim firstDate As Date, lastDate As Date
    Dim dayDiff As Long
    
    On Error Resume Next
    firstDate = CDate(xRange.Cells(1, 1).value)
    lastDate = CDate(xRange.Cells(xRange.Rows.Count, 1).value)
    On Error GoTo 0
    
    If firstDate > 0 And lastDate > firstDate Then
        dayDiff = lastDate - firstDate
        
        On Error Resume Next
        With ch.Axes(xlCategory)
            .CategoryType = xlTimeScale
            .MinimumScale = CDbl(firstDate)
            .MaximumScale = CDbl(lastDate)
            .MajorUnit = dayDiff
            .MajorUnitScale = xlDays
            .MinorUnit = dayDiff
            .MinorUnitScale = xlDays
            .TickLabels.NumberFormat = "yyyy-mm-dd"
        End With
        On Error GoTo 0
    End If
    
    ' --- Y轴自适应,并强制下限 ---
    Dim yMin As Double, yMax As Double
    Dim hasValid As Boolean
    hasValid = GetMinMaxFromRange(yRange, yMin, yMax)
    
    If hasValid Then
        Dim padding As Double
        If yMax > yMin Then
            padding = (yMax - yMin) * 0.1
        ElseIf isYieldChart Then
            padding = 1
        Else
            padding = 0.001
        End If
        
        Dim finalMin As Double, finalMax As Double
        finalMin = yMin - padding
        finalMax = yMax + padding
        
        ' --- 对齐到档位 ---
        Dim stepUnitMin As Double
        Dim stepUnitMax As Double
        
        If isYieldChart Then
            stepUnitMin = 0.1
            stepUnitMax = 1
        Else
            stepUnitMin = 0.001
            stepUnitMax = 0.001
        End If
        
        finalMin = AlignDown(finalMin, stepUnitMin)
        finalMax = AlignUp(finalMax, stepUnitMax)
        
        ' --- 强制Y轴下限 ---
        ' 净值图固定为1；收益率图固定为0
        If Not IsMissing(fixedYMin) Then
            finalMin = CDbl(fixedYMin)
        End If
        
        If finalMax <= finalMin Then
            finalMax = finalMin + stepUnitMax
        End If
        
        On Error Resume Next
        With ch.Axes(xlValue)
            .MinimumScale = finalMin
            .MaximumScale = finalMax
        End With
        On Error GoTo 0
    End If
    
    ' --- 强制ChartArea背景为白色,无边框 ---
    On Error Resume Next
    With ch.ChartArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Solid
    End With
    With ch.ChartArea.Format.Line
        .Visible = msoFalse
    End With
    On Error GoTo 0
    
    ' --- 字体设置: 西文Times New Roman, 中文仿宋, 复杂文字Times New Roman ---
    On Error Resume Next
    ' 批量设置全图字体
    With ch.ChartArea.Format.TextFrame2.TextRange.Font
        .name = "Times New Roman"
        .NameFarEast = "仿宋"
        .NameComplexScript = "Times New Roman"
    End With
    
    ' 标题字体单独设置
    If ch.HasTitle Then
        With ch.ChartTitle.Format.TextFrame2.TextRange.Font
            .name = "Times New Roman"
            .NameFarEast = "仿宋"
            .NameComplexScript = "Times New Roman"
            .Bold = msoFalse
        End With
        ch.ChartTitle.Font.Bold = False
    End If
    
    ' X轴/Y轴刻度标签字体
    With ch.Axes(xlCategory).TickLabels.Font
        .name = "Times New Roman"
    End With
    With ch.Axes(xlValue).TickLabels.Font
        .name = "Times New Roman"
    End With
    On Error GoTo 0
    
    ' --- 数据标签: 只保留最后一个有效点 ---
    ApplyLastPointDataLabel s, isYieldChart
End Sub

'==============================================================
' 辅助过程: 只显示最后一个数据点标签
'   放在模板和字体设置之后执行,避免被模板覆盖
'==============================================================
Private Sub ApplyLastPointDataLabel(ByVal s As Series, ByVal isYieldChart As Boolean)
    On Error Resume Next
    
    Dim ptCount As Long
    ptCount = s.Points.Count
    If ptCount <= 0 Then Exit Sub
    
    s.ApplyDataLabels Type:=xlDataLabelsShowValue
    
    Dim p As Long
    For p = 1 To ptCount
        s.Points(p).HasDataLabel = False
    Next p
    
    s.Points(ptCount).HasDataLabel = True
    With s.Points(ptCount).DataLabel
        .ShowValue = True
        .ShowCategoryName = False
        .ShowSeriesName = False
        If isYieldChart Then
            .NumberFormat = "0.0000"
        Else
            .NumberFormat = "0.0000"
        End If
        .Font.Name = "Times New Roman"
        .Font.Size = 11
        .Font.Bold = False
        .Position = xlLabelPositionAbove
    End With
    
    On Error GoTo 0
End Sub
'==============================================================
' 辅助函数: 从Range中获取数值的最小值和最大值
' 忽略空单元格和非数值,返回是否找到至少一个有效值
'==============================================================
Private Function GetMinMaxFromRange(ByVal rng As Range, _
                                     ByRef outMin As Double, ByRef outMax As Double) As Boolean
    Dim arr As Variant
    Dim isSingleCell As Boolean
    
    ' 单格Range读出来不是数组,需要兼容
    If rng.Cells.Count = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = rng.value
    Else
        arr = rng.value
    End If
    
    Dim r As Long, c As Long
    Dim v As Variant
    Dim found As Boolean: found = False
    Dim minVal As Double, maxVal As Double
    
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            v = arr(r, c)
            If Not IsEmpty(v) And IsNumeric(v) Then
                Dim d As Double: d = CDbl(v)
                If Not found Then
                    minVal = d
                    maxVal = d
                    found = True
                Else
                    If d < minVal Then minVal = d
                    If d > maxVal Then maxVal = d
                End If
            End If
        Next c
    Next r
    
    If found Then
        outMin = minVal
        outMax = maxVal
    End If
    GetMinMaxFromRange = found
End Function
'==============================================================
' 辅助函数: 向下取整到指定档位(stepUnit的整数倍)
'   例: AlignDown(1.0014, 0.001) = 1.001
'       AlignDown(2.34, 0.1)     = 2.3
'       AlignDown(-0.05, 0.1)    = -0.1  (负数也向下=更负)
'==============================================================
Private Function AlignDown(ByVal v As Double, ByVal stepUnit As Double) As Double
    If stepUnit <= 0 Then
        AlignDown = v
        Exit Function
    End If
    AlignDown = Int(v / stepUnit) * stepUnit
End Function

'==============================================================
' 辅助函数: 向上取整到指定档位
'   例: AlignUp(1.0014, 0.001) = 1.002
'       AlignUp(2.34, 0.1)     = 2.4
'       AlignUp(2.30, 0.1)     = 2.3  (已对齐则不变)
'==============================================================
Private Function AlignUp(ByVal v As Double, ByVal stepUnit As Double) As Double
    If stepUnit <= 0 Then
        AlignUp = v
        Exit Function
    End If
    
    Dim q As Double
    q = v / stepUnit
    
    ' 用一个小容差判断"是否已经是整数倍",避免浮点误差导致已对齐的值被多推一档
    Const EPS As Double = 0.000000001
    If Abs(q - Int(q)) < EPS Then
        ' 已对齐,不动
        AlignUp = Int(q) * stepUnit
    Else
        AlignUp = (Int(q) + 1) * stepUnit
    End If
End Function


