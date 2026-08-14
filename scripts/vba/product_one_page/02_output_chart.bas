' OnePage02_GenerateCharts：为产品一页通输出工作簿生成双坐标轴折线图

Option Explicit

Private Const TEMP_FILE_SUFFIX As String = "-产品一页通-临时.xlsx"
Private Const TEMP_FILE_PATTERN As String = "*-产品一页通-临时.xlsx"
Private Const FINAL_FILE_SUFFIX As String = "-产品一页通.xlsx"

Private Const COL_DATE As String = "A"
Private Const COL_PRODUCT_CODE As String = "B"
Private Const COL_PRODUCT_SHORT As String = "C"
Private Const COL_UNIT_NAV As String = "D"
Private Const COL_ANNUALIZED_RETURN As String = "F"
Private Const COL_REITS_ANNUALIZED_RETURN As String = "H"

Private Const PRODUCT_OA4400 As String = "OA4400"
Private Const REITS_LABEL As String = "中证REITs全收益-期间年化收益率"

Private Const CHART_LEFT_COL As String = "J"
Private Const CHART_TOP_ROW As Long = 2
Private Const CHART_WIDTH As Single = 500
Private Const CHART_HEIGHT As Single = 175

Private Const FONT_NAME As String = "微软雅黑"

Private Const COLOR_NAV As String = "#C00000"
Private Const COLOR_PRODUCT_RETURN As String = "#B29A73"
Private Const COLOR_REITS_RETURN As String = "#BFBFBF"

Private Const EXPECTED_RIGHT_GAP As Double = 54#
Private Const RIGHT_GAP_TOLERANCE As Double = 3#
Private Const LABEL_RIGHT_TOLERANCE As Double = 2#

Public Sub OnePage02_GenerateCharts()
    Dim appCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    appCalc = Application.Calculation

    Dim wbTarget As Workbook
    Dim wasOpen As Boolean
    Dim currentStep As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    currentStep = "查找产品一页通工作簿"
    Dim targetFile As String
    Dim targetPath As String
    targetFile = FindLatestOnePageWorkbook()
    targetPath = ThisWorkbook.Path & Application.PathSeparator & targetFile

    currentStep = "打开产品一页通临时工作簿"
    On Error Resume Next
    Set wbTarget = Workbooks(targetFile)
    On Error GoTo CleanFail
    If Not wbTarget Is Nothing Then
        Err.Raise vbObjectError + 5202, , "请先关闭一页通临时工作簿：" & targetFile
    End If
    Set wbTarget = Workbooks.Open(targetPath)

    currentStep = "生成产品图表"
    Dim ws As Worksheet
    Dim processedCount As Long
    Dim skippedText As String
    Dim errText As String

    For Each ws In wbTarget.Worksheets
        On Error GoTo SheetFail
        If LastUsedRow(ws) < 2 Then GoTo NextSheet
        If Len(NormalizeText(ws.Range(COL_PRODUCT_CODE & "2").Value)) = 0 Then GoTo NextSheet

        DeleteExistingCharts ws
        CreateProductChart ws
        processedCount = processedCount + 1
        On Error GoTo CleanFail
        GoTo NextSheet

SheetFail:
        errText = errText & ws.Name & "：" & Err.Description & vbCrLf
        Err.Clear
        On Error GoTo CleanFail

NextSheet:
    Next ws

    If Len(errText) > 0 Then Err.Raise vbObjectError + 5203, , "以下产品图表生成失败：" & vbCrLf & errText
    If processedCount = 0 Then Err.Raise vbObjectError + 5204, , "临时工作簿中没有可生成图表的产品sheet。"

    currentStep = "第一次保存并重开图表工作簿"
    wbTarget.Save
    wbTarget.Close SaveChanges:=False
    Set wbTarget = Workbooks.Open(targetPath)

    currentStep = "重开后重新应用图表布局"
    For Each ws In wbTarget.Worksheets
        If LastUsedRow(ws) >= 2 And Len(NormalizeText(ws.Range(COL_PRODUCT_CODE & "2").Value)) > 0 Then
            DeleteExistingCharts ws
            CreateProductChart ws
        End If
    Next ws
    wbTarget.Save
    wbTarget.Close SaveChanges:=False

    currentStep = "最终重开并验证图表布局"
    Set wbTarget = Workbooks.Open(targetPath)
    For Each ws In wbTarget.Worksheets
        If LastUsedRow(ws) >= 2 And Len(NormalizeText(ws.Range(COL_PRODUCT_CODE & "2").Value)) > 0 Then
            ValidateProductChart ws
        End If
    Next ws
    wbTarget.Close SaveChanges:=False
    Set wbTarget = Nothing

    currentStep = "替换正式图表工作簿"
    Dim finalPath As String
    finalPath = ThisWorkbook.Path & Application.PathSeparator & Left$(targetFile, 8) & FINAL_FILE_SUFFIX
    ReplaceFileKeepingOldOnFailure targetPath, finalPath

    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    Dim finalMsg As String
    finalMsg = "产品一页通图表生成完成" & vbCrLf & vbCrLf & _
               "目标文件：" & Mid$(finalPath, Len(ThisWorkbook.Path) + 2) & vbCrLf & vbCrLf & _
               "处理结果：" & vbCrLf & _
               "处理sheet数：" & processedCount
    If Len(skippedText) > 0 Or Len(errText) > 0 Then
        finalMsg = finalMsg & vbCrLf & vbCrLf & "注意事项："
        If Len(skippedText) > 0 Then finalMsg = finalMsg & vbCrLf & "跳过明细：" & vbCrLf & skippedText
        If Len(errText) > 0 Then finalMsg = finalMsg & vbCrLf & "异常明细：" & vbCrLf & errText
    End If
    ReportPipeline_MsgBox finalMsg, vbInformation, "产品一页通"
    Exit Sub

CleanFail:
    Dim failNumber As Long
    Dim failDescription As String
    Dim failStep As String
    failNumber = Err.Number
    failDescription = Err.Description
    failStep = currentStep

    On Error Resume Next
    If Not wbTarget Is Nothing Then
        If Not wasOpen Then wbTarget.Close SaveChanges:=False
    End If
    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If Len(failDescription) = 0 Then failDescription = "未知错误"
    If Len(failStep) = 0 Then failStep = "未记录"

    ReportPipeline_MsgBox "产品一页通图表生成失败" & vbCrLf & vbCrLf & _
           "错误信息：" & failDescription & vbCrLf & _
           "错误号：" & failNumber & vbCrLf & _
           "步骤：" & failStep, vbCritical, "产品一页通"
End Sub

Private Function FindLatestOnePageWorkbook() As String
    Dim fileName As String
    Dim latestFile As String
    Dim latestKey As String
    Dim dateText As String

    fileName = Dir$(ThisWorkbook.Path & Application.PathSeparator & TEMP_FILE_PATTERN)
    Do While Len(fileName) > 0
        dateText = Left$(fileName, 8)
        If Len(dateText) = 8 And IsNumeric(dateText) Then
            If Right$(fileName, Len(TEMP_FILE_SUFFIX)) = TEMP_FILE_SUFFIX Then
                If dateText > latestKey Then
                    latestKey = dateText
                    latestFile = fileName
                End If
            End If
        End If
        fileName = Dir$()
    Loop

    If Len(latestFile) = 0 Then
        Err.Raise vbObjectError + 5201, , "未找到 yyyymmdd-产品一页通-临时.xlsx，请先运行 OnePage01_ExportChartData。"
    End If

    FindLatestOnePageWorkbook = latestFile
End Function

Private Sub CreateProductChart(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim productCode As String
    Dim productShort As String
    productCode = NormalizeText(ws.Range(COL_PRODUCT_CODE & "2").Value)
    productShort = NormalizeText(ws.Range(COL_PRODUCT_SHORT & "2").Value)
    If Len(productShort) = 0 Then productShort = ws.Name

    Dim dateRng As Range
    Dim navRng As Range
    Dim productReturnRng As Range
    Set dateRng = ws.Range(COL_DATE & "2:" & COL_DATE & lastRow)
    Set navRng = ws.Range(COL_UNIT_NAV & "2:" & COL_UNIT_NAV & lastRow)
    Set productReturnRng = ws.Range(COL_ANNUALIZED_RETURN & "2:" & COL_ANNUALIZED_RETURN & lastRow)

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=ws.Range(CHART_LEFT_COL & CHART_TOP_ROW).Left, _
                                  Top:=ws.Range(CHART_LEFT_COL & CHART_TOP_ROW).Top, _
                                  Width:=CHART_WIDTH, Height:=CHART_HEIGHT)
    co.Name = "chart_产品一页通"

    Dim ch As Chart
    Set ch = co.Chart
    ch.ChartType = xlLine
    ch.HasTitle = False
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionTop

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    Dim sNav As Series
    Set sNav = AddLineSeries(ch, dateRng, navRng, productShort & "-单位净值", xlPrimary, COLOR_NAV)

    Dim sProductReturn As Series
    Set sProductReturn = AddLineSeries(ch, dateRng, productReturnRng, productShort & "-成立以来年化收益率", xlSecondary, COLOR_PRODUCT_RETURN)

    Dim sReitsReturn As Series
    If StrComp(productCode, PRODUCT_OA4400, vbTextCompare) = 0 Then
        Set sReitsReturn = AddLineSeries(ch, dateRng, ws.Range(COL_REITS_ANNUALIZED_RETURN & "2:" & COL_REITS_ANNUALIZED_RETURN & lastRow), REITS_LABEL, xlSecondary, COLOR_REITS_RETURN)
    End If

    FormatChartArea ch
    FormatCategoryAxis ch, dateRng
    FormatValueAxes ch, navRng, productReturnRng, sReitsReturn, productCode
    FormatPlotArea ch

    ApplyLastPointDataLabel sNav, navRng, "0.0000", COLOR_NAV
    ApplyLastPointDataLabel sProductReturn, productReturnRng, "0.00%", COLOR_PRODUCT_RETURN
    If Not sReitsReturn Is Nothing Then ApplyLastPointDataLabel sReitsReturn, ws.Range(COL_REITS_ANNUALIZED_RETURN & "2:" & COL_REITS_ANNUALIZED_RETURN & lastRow), "0.00%", COLOR_REITS_RETURN

    ch.DisplayBlanksAs = xlInterpolated
End Sub

Private Function AddLineSeries(ByVal ch As Chart, ByVal xRange As Range, ByVal yRange As Range, ByVal seriesName As String, ByVal axisGroup As XlAxisGroup, ByVal colorHex As String) As Series
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.XValues = xRange
    s.Values = yRange
    s.Name = seriesName
    s.ChartType = xlLine
    s.AxisGroup = axisGroup

    ApplySeriesLineStyle s, ColorFromHex(colorHex)

    Set AddLineSeries = s
End Function

Private Sub ApplySeriesLineStyle(ByVal s As Series, ByVal lineColor As Long)
    On Error Resume Next
    With s.Border
        .LineStyle = xlContinuous
        .Color = lineColor
        .Weight = xlMedium
    End With
    s.MarkerStyle = xlMarkerStyleNone
    On Error GoTo 0
End Sub

Private Function ColorFromHex(ByVal colorHex As String) As Long
    Dim s As String
    s = Replace(colorHex, "#", vbNullString)
    If Len(s) <> 6 Then
        ColorFromHex = RGB(0, 0, 0)
        Exit Function
    End If

    ColorFromHex = RGB(CLng("&H" & Left$(s, 2)), CLng("&H" & Mid$(s, 3, 2)), CLng("&H" & Right$(s, 2)))
End Function

Private Sub FormatChartArea(ByVal ch As Chart)
    On Error Resume Next
    With ch.ChartArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Solid
    End With
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Line.Visible = msoFalse

    With ch.ChartArea.Format.TextFrame2.TextRange.Font
        .Name = FONT_NAME
        .NameFarEast = FONT_NAME
        .NameComplexScript = FONT_NAME
        .Bold = msoTrue
    End With
    ch.ChartArea.Font.Name = FONT_NAME
    ch.ChartArea.Font.Bold = True
    ch.Legend.Font.Name = FONT_NAME
    ch.Legend.Font.Size = 10
    ch.Legend.Font.Bold = True
    On Error GoTo 0
End Sub

Private Sub FormatCategoryAxis(ByVal ch As Chart, ByVal dateRng As Range)
    Dim firstDate As Date
    Dim lastDate As Date
    Dim dayDiff As Long

    On Error Resume Next
    firstDate = CDate(dateRng.Cells(1, 1).Value)
    lastDate = CDate(dateRng.Cells(dateRng.Rows.Count, 1).Value)
    On Error GoTo 0

    If firstDate = 0 Or lastDate <= firstDate Then Exit Sub

    dayDiff = DateDiff("d", firstDate, lastDate)
    If dayDiff <= 0 Then dayDiff = 1

    On Error Resume Next
    With ch.Axes(xlCategory)
        .CategoryType = xlTimeScale
        .MinimumScale = CDbl(firstDate)
        .MaximumScale = CDbl(lastDate)
        .MajorUnit = dayDiff
        .MajorUnitScale = xlDays
        .MinorUnit = dayDiff
        .MinorUnitScale = xlDays
        .TickLabels.NumberFormat = "yyyy年mm月dd日"
        .TickLabels.Font.Name = FONT_NAME
        .TickLabels.Font.Bold = True
    End With
    On Error GoTo 0
End Sub

Private Sub FormatPlotArea(ByVal ch As Chart)
    Const PLOT_LEFT As Double = 42
    Const PLOT_TOP As Double = 24
    Const PLOT_BOTTOM_PAD As Double = 32

    With ch.PlotArea
        .InsideLeft = PLOT_LEFT
        .InsideTop = PLOT_TOP
        .InsideWidth = ch.ChartArea.Width - PLOT_LEFT - EXPECTED_RIGHT_GAP
        .InsideHeight = ch.ChartArea.Height - PLOT_TOP - PLOT_BOTTOM_PAD
    End With
End Sub

Private Sub FormatValueAxes(ByVal ch As Chart, ByVal navRng As Range, ByVal productReturnRng As Range, ByVal reitsSeries As Series, ByVal productCode As String)
    Dim navMin As Double
    Dim navMax As Double
    If Not GetMinMaxFromRange(navRng, navMin, navMax) Then
        Err.Raise vbObjectError + 5211, , "单位净值列没有有效数值。"
    End If

    Dim navAxisMax As Double
    If StrComp(productCode, PRODUCT_OA4400, vbTextCompare) = 0 Then
        navAxisMax = AlignUp(navMax + (navMax - 1#), 0.01)
    Else
        navAxisMax = AlignUp(navMax, 0.01)
    End If
    If navAxisMax <= 1# Then navAxisMax = 1.01

    Dim retMin As Double
    Dim retMax As Double
    If Not GetMinMaxFromRange(productReturnRng, retMin, retMax) Then
        Err.Raise vbObjectError + 5212, , "成立以来年化收益率列没有有效数值。"
    End If

    If Not reitsSeries Is Nothing Then
        IncludeSeriesMinMax reitsSeries, retMin, retMax
    End If

    Dim retAxisMin As Double
    Dim retAxisMax As Double
    If StrComp(productCode, PRODUCT_OA4400, vbTextCompare) = 0 Then
        retAxisMin = AlignDown(retMin + retMin * 0.1, 0.01)
        retAxisMax = AlignUp(retMax + retMax * 0.1, 0.01)
    Else
        retAxisMin = 0#
        retAxisMax = AlignUp(retMax + retMax * 1#, 0.01)
    End If
    If retAxisMax <= retAxisMin Then retAxisMax = retAxisMin + 0.01

    With ch.Axes(xlValue, xlPrimary)
        .MinimumScale = 1#
        .MaximumScale = navAxisMax
        .TickLabels.NumberFormat = "0.0000"
        .TickLabels.Font.Name = FONT_NAME
        .TickLabels.Font.Bold = True
    End With

    With ch.Axes(xlValue, xlSecondary)
        .MinimumScale = retAxisMin
        .MaximumScale = retAxisMax
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabels.NumberFormat = ";;;"
        .TickLabels.Font.Name = FONT_NAME
        .TickLabels.Font.Bold = True
    End With
End Sub

Private Sub ApplyLastPointDataLabel(ByVal s As Series, ByVal yRange As Range, ByVal numberFormatText As String, ByVal colorHex As String)
    Dim ptCount As Long
    ptCount = s.Points.Count
    If ptCount <= 0 Then Exit Sub

    Dim lastValidPoint As Long
    lastValidPoint = LastNumericPointIndexFromRange(yRange)
    If lastValidPoint <= 0 Then Exit Sub
    If lastValidPoint > ptCount Then lastValidPoint = ptCount

    s.ApplyDataLabels Type:=xlDataLabelsShowValue

    Dim p As Long
    For p = 1 To ptCount
        s.Points(p).HasDataLabel = False
    Next p

    s.Points(lastValidPoint).HasDataLabel = True
    With s.Points(lastValidPoint).DataLabel
        .ShowValue = True
        .ShowCategoryName = False
        .ShowSeriesName = False
        .NumberFormat = numberFormatText
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = ColorFromHex(colorHex)
        .Position = xlLabelPositionRight
    End With
End Sub

Private Sub ValidateProductChart(ByVal ws As Worksheet)
    Dim co As ChartObject
    On Error Resume Next
    Set co = ws.ChartObjects("chart_产品一页通")
    On Error GoTo 0
    If co Is Nothing Then Err.Raise vbObjectError + 5220, , "产品sheet缺少一页通图表：" & ws.Name

    co.Activate
    DoEvents

    Dim ch As Chart
    Set ch = co.Chart
    Dim actualGap As Double
    actualGap = ch.ChartArea.Width - ch.PlotArea.InsideLeft - ch.PlotArea.InsideWidth
    If Abs(actualGap - EXPECTED_RIGHT_GAP) > RIGHT_GAP_TOLERANCE Then
        Err.Raise vbObjectError + 5221, , "图表右侧留白验证失败：" & ws.Name & _
                  "；期望=" & Format$(EXPECTED_RIGHT_GAP, "0.0") & _
                  "；实际=" & Format$(actualGap, "0.0")
    End If

    If ch.Axes(xlValue, xlSecondary).TickLabels.NumberFormat <> ";;;" Then
        Err.Raise vbObjectError + 5222, , "图表右轴刻度标签未隐藏：" & ws.Name
    End If

    Dim plotRight As Double
    plotRight = ch.PlotArea.InsideLeft + ch.PlotArea.InsideWidth
    Dim seriesItem As Series
    For Each seriesItem In ch.SeriesCollection
        ValidateLastDataLabel seriesItem, plotRight, ch.ChartArea.Width, ws.Name
    Next seriesItem
End Sub

Private Sub ValidateLastDataLabel(ByVal s As Series, ByVal plotRight As Double, _
                                  ByVal chartWidth As Double, ByVal sheetName As String)
    Dim pointIndex As Long
    For pointIndex = s.Points.Count To 1 Step -1
        If s.Points(pointIndex).HasDataLabel Then Exit For
    Next pointIndex
    If pointIndex = 0 Then Err.Raise vbObjectError + 5223, , "图表末端数据标签缺失：" & sheetName

    ' Excel 2007可能不支持DataLabel的几何属性；标签存在仍是强制要求，
    ' 只有支持几何读取的版本才执行边界检查。
    On Error GoTo GeometryUnavailable
    Dim label As DataLabel
    Set label = s.Points(pointIndex).DataLabel
    If label.Left < plotRight - RIGHT_GAP_TOLERANCE Then
        Err.Raise vbObjectError + 5224, , "图表数据标签未进入右侧留白区：" & sheetName
    End If
    If label.Left + label.Width > chartWidth + LABEL_RIGHT_TOLERANCE Then
        Err.Raise vbObjectError + 5225, , "图表数据标签超出右边界：" & sheetName
    End If
    Exit Sub

GeometryUnavailable:
    If Err.Number = 438 Then
        Err.Clear
        Exit Sub
    End If
    Dim errorNumber As Long
    Dim errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, , errorDescription
End Sub

Private Sub ReplaceFileKeepingOldOnFailure(ByVal tempPath As String, ByVal finalPath As String)
    Dim backupPath As String
    backupPath = finalPath & ".bak"
    On Error GoTo ReplaceFail

    If Len(Dir$(backupPath)) > 0 Then Kill backupPath
    If Len(Dir$(finalPath)) > 0 Then Name finalPath As backupPath
    Name tempPath As finalPath
    If Len(Dir$(backupPath)) > 0 Then Kill backupPath
    Exit Sub

ReplaceFail:
    Dim description As String
    description = Err.Description
    On Error Resume Next
    If Len(Dir$(finalPath)) = 0 And Len(Dir$(backupPath)) > 0 Then Name backupPath As finalPath
    On Error GoTo 0
    Err.Raise vbObjectError + 5226, , "替换正式图表工作簿失败：" & description
End Sub

Private Function LastNumericPointIndexFromRange(ByVal rng As Range) As Long
    Dim i As Long
    Dim v As Variant

    For i = rng.Rows.Count To 1 Step -1
        v = rng.Cells(i, 1).Value
        If Not IsError(v) And Not IsEmpty(v) And IsNumeric(v) Then
            LastNumericPointIndexFromRange = i
            Exit Function
        End If
    Next i
End Function

Private Sub DeleteExistingCharts(ByVal ws As Worksheet)
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co
End Sub

Private Function GetMinMaxFromRange(ByVal rng As Range, ByRef outMin As Double, ByRef outMax As Double) As Boolean
    Dim arr As Variant
    If rng.Cells.Count = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = rng.Value
    Else
        arr = rng.Value
    End If

    Dim r As Long
    Dim c As Long
    Dim v As Variant
    Dim d As Double
    Dim found As Boolean

    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            v = arr(r, c)
            If Not IsError(v) And Not IsEmpty(v) And IsNumeric(v) Then
                d = CDbl(v)
                If Not found Then
                    outMin = d
                    outMax = d
                    found = True
                Else
                    If d < outMin Then outMin = d
                    If d > outMax Then outMax = d
                End If
            End If
        Next c
    Next r

    GetMinMaxFromRange = found
End Function

Private Sub IncludeSeriesMinMax(ByVal s As Series, ByRef outMin As Double, ByRef outMax As Double)
    Dim valuesArr As Variant
    valuesArr = s.Values

    Dim v As Variant
    Dim d As Double
    If IsArray(valuesArr) Then
        Dim i As Long
        For i = LBound(valuesArr) To UBound(valuesArr)
            v = valuesArr(i)
            If Not IsError(v) And Not IsEmpty(v) And IsNumeric(v) Then
                d = CDbl(v)
                If d < outMin Then outMin = d
                If d > outMax Then outMax = d
            End If
        Next i
    ElseIf Not IsError(valuesArr) And Not IsEmpty(valuesArr) And IsNumeric(valuesArr) Then
        d = CDbl(valuesArr)
        If d < outMin Then outMin = d
        If d > outMax Then outMax = d
    End If
End Sub

Private Function AlignDown(ByVal v As Double, ByVal stepUnit As Double) As Double
    If stepUnit <= 0 Then
        AlignDown = v
    Else
        AlignDown = Int(v / stepUnit) * stepUnit
    End If
End Function

Private Function AlignUp(ByVal v As Double, ByVal stepUnit As Double) As Double
    If stepUnit <= 0 Then
        AlignUp = v
        Exit Function
    End If

    Dim q As Double
    q = v / stepUnit

    Const EPS As Double = 0.000000001
    If Abs(q - Int(q)) < EPS Then
        AlignUp = Int(q) * stepUnit
    Else
        AlignUp = (Int(q) + 1) * stepUnit
    End If
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = foundCell.Row
    End If
End Function

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
