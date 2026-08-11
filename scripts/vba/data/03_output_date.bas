' Data03_ExportProductReport：生成统一字段结构的 yyyyMMdd-上层产品分类表现.xlsx

Option Explicit

Private Const SHEET_PRODUCT_CATEGORY As String = "产品分类"
Private Const SHEET_NAV As String = "上层产品净值数据(181)"
Private Const SHEET_PRODUCT_INFO As String = "产品信息"

Private Const COL_SEQ As String = "序号"
Private Const COL_TRUST_CODE As String = "信托计划代码"
Private Const COL_SERIES As String = "系列"
Private Const COL_CATEGORY As String = "分类"
Private Const COL_PRODUCT_NAME As String = "产品名称"
Private Const COL_EXPORT_ENABLED As String = "是否导出"
Private Const COL_PREV_PREV_OPEN As String = "上上一开放日"
Private Const COL_PREV_OPEN As String = "上一开放日"
Private Const COL_BASELINE_DATE As String = "基准日期"
Private Const COL_NEXT_OPEN As String = "下一开放日"
Private Const COL_THEORETICAL_INTERVAL As String = "理论间隔"
Private Const COL_PREV_INTERVAL As String = "上次开放实际间隔"
Private Const COL_INTERVAL As String = "实际间隔"
Private Const COL_ELAPSED As String = "运作时间"
Private Const COL_PREV_PREV_NAV As String = "上上一开放日净值"
Private Const COL_PREV_NAV As String = "上一开放日净值"
Private Const COL_BASELINE_NAV As String = "基准日期净值"
Private Const COL_BENCHMARK_RATE As String = "基准收益率"
Private Const COL_PREV_PERIOD_ANNUAL As String = "上一周期年化"
Private Const COL_CURRENT_PERIOD_ANNUAL As String = "当前周期年化"
Private Const COL_7DAY_ANNUAL As String = "7日年化"
Private Const COL_28DAY_ANNUAL As String = "28日年化"
Private Const COL_INCEPTION_ANNUAL As String = "成立以来年化"
Private Const COL_NAV_DATE As String = "日期"
Private Const COL_UNIT_NAV As String = "单位净值"
Private Const COL_INCEPTION_DATE As String = "成立日"
Private Const DEFAULT_INCEPTION_NAV As Double = 1

Public Sub Data03_ExportProductReport()
    ExportUnifiedProductReport
End Sub

Private Sub ExportUnifiedProductReport()
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    Dim wbOutput As Workbook
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Dim configErrors As String
    If Not ReportConfig_Validate(configErrors) Then
        Err.Raise vbObjectError + 4001, , "报表配置静态预检未通过：" & vbCrLf & configErrors
    End If

    Dim categorySheetMap As Object
    Set categorySheetMap = ReportConfig_GetEnabledCategorySheetMap()
    If categorySheetMap.Count = 0 Then Err.Raise vbObjectError + 4002, , "没有启用的报表分类。"

    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Worksheets(SHEET_PRODUCT_CATEGORY)
    Dim sourceHeaders As Object
    Set sourceHeaders = BuildHeaderMap(wsSource, 1)
    RequireSourceHeaders sourceHeaders

    Dim navLookup As Object
    Dim inceptionLookup As Object
    Set navLookup = BuildNAVDateLookup()
    Set inceptionLookup = BuildInceptionDateLookup()

    Dim baselineDate As Date
    baselineDate = ReadBaselineDate(wsSource, sourceHeaders)

    Set wbOutput = CreateOutputWorkbook(categorySheetMap)
    WriteUnifiedHeaders wbOutput

    Dim rowCounters As Object
    Set rowCounters = CreateTextDictionary()
    Dim categoryKey As Variant
    For Each categoryKey In categorySheetMap.Keys
        rowCounters.Add CStr(categoryKey), 2
    Next categoryKey

    Dim exportedCount As Long
    Dim skippedCount As Long
    Dim missingMetricCount As Long
    WriteAllProducts wsSource, sourceHeaders, wbOutput, categorySheetMap, rowCounters, _
                     navLookup, inceptionLookup, exportedCount, skippedCount, missingMetricCount

    Dim outputPath As String
    Dim tempPath As String
    outputPath = ThisWorkbook.Path & Application.PathSeparator & _
                 Format$(baselineDate, "yyyymmdd") & "-上层产品分类表现.xlsx"
    tempPath = ThisWorkbook.Path & Application.PathSeparator & _
               Format$(baselineDate, "yyyymmdd") & "-上层产品分类表现.tmp.xlsx"

    If Len(Dir(tempPath)) > 0 Then Kill tempPath
    wbOutput.SaveAs FileName:=tempPath, FileFormat:=xlOpenXMLWorkbook
    wbOutput.Close SaveChanges:=False
    Set wbOutput = Nothing
    ReplaceFileSafely tempPath, outputPath

    RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation

    ReportPipeline_MsgBox "产品分类表现报告生成完成" & vbCrLf & vbCrLf & _
           "基准日期：" & Format$(baselineDate, "yyyy-mm-dd") & vbCrLf & _
           "输出产品：" & exportedCount & vbCrLf & _
           "停用或未配置产品：" & skippedCount & vbCrLf & _
           "存在空缺指标的产品：" & missingMetricCount & vbCrLf & vbCrLf & _
           "输出文件：" & vbCrLf & outputPath, vbInformation, "产品分类表现报告"
    Exit Sub

CleanFail:
    Dim errorDescription As String
    errorDescription = Err.Description
    On Error Resume Next
    If Not wbOutput Is Nothing Then wbOutput.Close SaveChanges:=False
    On Error GoTo 0
    RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation
    ReportPipeline_MsgBox "产品分类表现报告生成失败" & vbCrLf & vbCrLf & _
           "错误信息：" & errorDescription, vbCritical, "产品分类表现报告"
End Sub

Private Sub WriteAllProducts(ByVal wsSource As Worksheet, ByVal sourceHeaders As Object, _
                             ByVal wbOutput As Workbook, ByVal categorySheetMap As Object, _
                             ByVal rowCounters As Object, ByVal navLookup As Object, _
                             ByVal inceptionLookup As Object, ByRef exportedCount As Long, _
                             ByRef skippedCount As Long, ByRef missingMetricCount As Long)
    Dim lastRow As Long
    lastRow = LastUsedRow(wsSource)

    Dim r As Long
    For r = 2 To lastRow
        Dim trustCode As String
        trustCode = NormalizeText(SourceValue(wsSource, r, sourceHeaders, COL_TRUST_CODE))
        If Len(trustCode) = 0 Then GoTo ContinueRow
        If Not IsEnabledValue(SourceValue(wsSource, r, sourceHeaders, COL_EXPORT_ENABLED)) Then
            skippedCount = skippedCount + 1
            GoTo ContinueRow
        End If

        Dim categoryName As String
        categoryName = NormalizeText(SourceValue(wsSource, r, sourceHeaders, COL_CATEGORY))
        If Not categorySheetMap.Exists(categoryName) Then
            skippedCount = skippedCount + 1
            GoTo ContinueRow
        End If

        Dim wsOutput As Worksheet
        Dim outputRow As Long
        Set wsOutput = wbOutput.Worksheets(CStr(categorySheetMap(categoryName)))
        outputRow = CLng(rowCounters(categoryName))

        Dim prevPrevNav As Variant
        Dim prevNav As Variant
        Dim baselineNav As Variant
        Dim prevInterval As Variant
        Dim elapsedDays As Variant
        Dim theoreticalInterval As Variant
        Dim baselineDateValue As Variant

        prevPrevNav = SourceValue(wsSource, r, sourceHeaders, COL_PREV_PREV_NAV)
        prevNav = SourceValue(wsSource, r, sourceHeaders, COL_PREV_NAV)
        baselineNav = SourceValue(wsSource, r, sourceHeaders, COL_BASELINE_NAV)
        prevInterval = SourceValue(wsSource, r, sourceHeaders, COL_PREV_INTERVAL)
        elapsedDays = SourceValue(wsSource, r, sourceHeaders, COL_ELAPSED)
        theoreticalInterval = SourceValue(wsSource, r, sourceHeaders, COL_THEORETICAL_INTERVAL)
        baselineDateValue = SourceValue(wsSource, r, sourceHeaders, COL_BASELINE_DATE)

        Dim criticalDate As Date
        If Len(NormalizeText(SourceValue(wsSource, r, sourceHeaders, COL_PRODUCT_NAME))) = 0 Then
            Err.Raise vbObjectError + 4007, , "产品“" & trustCode & "”缺少产品名称。"
        End If
        If Not TryReadDate(baselineDateValue, criticalDate) Then
            Err.Raise vbObjectError + 4008, , "产品“" & trustCode & "”缺少有效基准日期。"
        End If
        Dim prevAnnual As Variant
        Dim currentAnnual As Variant
        Dim annual7 As Variant
        Dim annual28 As Variant
        Dim inceptionAnnual As Variant

        prevAnnual = CalculateSimpleAnnual(prevNav, prevPrevNav, prevInterval)
        currentAnnual = CalculateSimpleAnnual(baselineNav, prevNav, elapsedDays)

        If IsNumericValue(theoreticalInterval) And IsNumericValue(elapsedDays) Then
            If CDbl(theoreticalInterval) > 60 And CDbl(elapsedDays) < 7 And Not IsEmpty(prevAnnual) Then
                currentAnnual = prevAnnual
            End If
        End If

        annual7 = CalculateTrailingAnnual(navLookup, trustCode, baselineDateValue, baselineNav, 7)
        annual28 = CalculateTrailingAnnual(navLookup, trustCode, baselineDateValue, baselineNav, 28)
        inceptionAnnual = CalculateInceptionAnnual(inceptionLookup, trustCode, baselineDateValue, baselineNav)

        WriteOutputValue wsOutput, outputRow, 1, SourceValue(wsSource, r, sourceHeaders, COL_SEQ)
        WriteOutputValue wsOutput, outputRow, 2, trustCode
        WriteOutputValue wsOutput, outputRow, 3, SourceValue(wsSource, r, sourceHeaders, COL_SERIES)
        WriteOutputValue wsOutput, outputRow, 4, categoryName
        WriteOutputValue wsOutput, outputRow, 5, SourceValue(wsSource, r, sourceHeaders, COL_PRODUCT_NAME)
        WriteDateValue wsOutput, outputRow, 6, SourceValue(wsSource, r, sourceHeaders, COL_PREV_PREV_OPEN)
        WriteOutputValue wsOutput, outputRow, 7, prevPrevNav
        WriteDateValue wsOutput, outputRow, 8, SourceValue(wsSource, r, sourceHeaders, COL_PREV_OPEN)
        WriteOutputValue wsOutput, outputRow, 9, prevNav
        WriteDateValue wsOutput, outputRow, 10, baselineDateValue
        WriteOutputValue wsOutput, outputRow, 11, baselineNav
        WriteDateValue wsOutput, outputRow, 12, SourceValue(wsSource, r, sourceHeaders, COL_NEXT_OPEN)
        WriteOutputValue wsOutput, outputRow, 13, theoreticalInterval
        WriteOutputValue wsOutput, outputRow, 14, SourceValue(wsSource, r, sourceHeaders, COL_INTERVAL)
        WriteOutputValue wsOutput, outputRow, 15, prevInterval
        WriteOutputValue wsOutput, outputRow, 16, elapsedDays
        WriteOutputValue wsOutput, outputRow, 17, SourceValue(wsSource, r, sourceHeaders, COL_BENCHMARK_RATE)
        WritePercentValue wsOutput, outputRow, 18, prevAnnual
        WritePercentValue wsOutput, outputRow, 19, currentAnnual
        WritePercentValue wsOutput, outputRow, 20, annual7
        WritePercentValue wsOutput, outputRow, 21, annual28
        WritePercentValue wsOutput, outputRow, 22, inceptionAnnual

        If IsEmpty(prevAnnual) Or IsEmpty(currentAnnual) Or IsEmpty(annual7) Or _
           IsEmpty(annual28) Or IsEmpty(inceptionAnnual) Then
            missingMetricCount = missingMetricCount + 1
        End If

        rowCounters(categoryName) = outputRow + 1
        exportedCount = exportedCount + 1
ContinueRow:
    Next r

    Dim sheet As Worksheet
    For Each sheet In wbOutput.Worksheets
        sheet.Rows(1).Font.Bold = True
        sheet.Columns("A:V").AutoFit
        sheet.Columns("F:F").NumberFormat = "yyyy-mm-dd"
        sheet.Columns("H:H").NumberFormat = "yyyy-mm-dd"
        sheet.Columns("J:J").NumberFormat = "yyyy-mm-dd"
        sheet.Columns("L:L").NumberFormat = "yyyy-mm-dd"
        sheet.Columns("R:V").NumberFormat = "0.00%"
    Next sheet
End Sub

Private Function CreateOutputWorkbook(ByVal categorySheetMap As Object) As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add

    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop

    Dim categoryKey As Variant
    Dim isFirst As Boolean
    isFirst = True
    For Each categoryKey In categorySheetMap.Keys
        Dim ws As Worksheet
        If isFirst Then
            Set ws = wb.Worksheets(1)
            isFirst = False
        Else
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        End If
        ws.Name = CStr(categorySheetMap(categoryKey))
    Next categoryKey

    Set CreateOutputWorkbook = wb
End Function

Private Sub WriteUnifiedHeaders(ByVal wb As Workbook)
    Dim headers As Variant
    headers = Array(COL_SEQ, COL_TRUST_CODE, COL_SERIES, COL_CATEGORY, COL_PRODUCT_NAME, _
                    COL_PREV_PREV_OPEN, COL_PREV_PREV_NAV, COL_PREV_OPEN, COL_PREV_NAV, _
                    COL_BASELINE_DATE, COL_BASELINE_NAV, COL_NEXT_OPEN, COL_THEORETICAL_INTERVAL, _
                    COL_INTERVAL, COL_PREV_INTERVAL, COL_ELAPSED, COL_BENCHMARK_RATE, _
                    COL_PREV_PERIOD_ANNUAL, COL_CURRENT_PERIOD_ANNUAL, COL_7DAY_ANNUAL, _
                    COL_28DAY_ANNUAL, COL_INCEPTION_ANNUAL)

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim i As Long
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
    Next ws
End Sub

Private Function CalculateSimpleAnnual(ByVal endNav As Variant, ByVal startNav As Variant, _
                                       ByVal dayCount As Variant) As Variant
    If Not IsNumericValue(endNav) Or Not IsNumericValue(startNav) Or Not IsNumericValue(dayCount) Then
        CalculateSimpleAnnual = Empty
    ElseIf CDbl(startNav) = 0 Or CDbl(dayCount) <= 0 Then
        CalculateSimpleAnnual = Empty
    Else
        CalculateSimpleAnnual = (CDbl(endNav) / CDbl(startNav) - 1) * (365 / CDbl(dayCount))
    End If
End Function

Private Function CalculateTrailingAnnual(ByVal navLookup As Object, ByVal trustCode As String, _
                                         ByVal baselineDateValue As Variant, ByVal baselineNav As Variant, _
                                         ByVal trailingDays As Long) As Variant
    Dim baselineDate As Date
    If Not TryReadDate(baselineDateValue, baselineDate) Or Not IsNumericValue(baselineNav) Then
        CalculateTrailingAnnual = Empty
        Exit Function
    End If

    Dim priorNav As Variant
    priorNav = LookupNAV(navLookup, trustCode, DateAdd("d", -trailingDays, baselineDate))
    CalculateTrailingAnnual = CalculateSimpleAnnual(baselineNav, priorNav, trailingDays)
End Function

Private Function CalculateInceptionAnnual(ByVal inceptionLookup As Object, ByVal trustCode As String, _
                                          ByVal baselineDateValue As Variant, ByVal baselineNav As Variant) As Variant
    Dim baselineDate As Date
    If Not TryReadDate(baselineDateValue, baselineDate) Or Not IsNumericValue(baselineNav) Then
        CalculateInceptionAnnual = Empty
        Exit Function
    End If
    If Not inceptionLookup.Exists(trustCode) Then
        CalculateInceptionAnnual = Empty
        Exit Function
    End If

    Dim inceptionDate As Date
    inceptionDate = CDate(inceptionLookup(trustCode))
    Dim dayCount As Long
    dayCount = DateDiff("d", inceptionDate, baselineDate) + 1
    If dayCount <= 0 Or DEFAULT_INCEPTION_NAV = 0 Then
        CalculateInceptionAnnual = Empty
    Else
        CalculateInceptionAnnual = (CDbl(baselineNav) / DEFAULT_INCEPTION_NAV - 1) * (365 / dayCount)
    End If
End Function

Private Function BuildNAVDateLookup() As Object
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAV)
    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)
    RequireHeader headers, COL_TRUST_CODE, SHEET_NAV
    RequireHeader headers, COL_NAV_DATE, SHEET_NAV
    RequireHeader headers, COL_UNIT_NAV, SHEET_NAV

    Dim result As Object
    Set result = CreateTextDictionary()

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    Dim r As Long
    For r = 2 To lastRow
        Dim trustCode As String
        Dim navDate As Date
        Dim navValue As Variant
        trustCode = NormalizeText(ws.Cells(r, CLng(headers(COL_TRUST_CODE))).Value)
        navValue = ws.Cells(r, CLng(headers(COL_UNIT_NAV))).Value
        If Len(trustCode) = 0 Then GoTo ContinueRow
        If Not TryReadDate(ws.Cells(r, CLng(headers(COL_NAV_DATE))).Value, navDate) Then GoTo ContinueRow
        If Not IsNumericValue(navValue) Then GoTo ContinueRow

        Dim productDates As Object
        If Not result.Exists(trustCode) Then
            Set productDates = CreateTextDictionary()
            Set result.Item(trustCode) = productDates
        Else
            Set productDates = result.Item(trustCode)
        End If
        Dim dateKey As String
        dateKey = Format$(navDate, "yyyymmdd")
        If Not productDates.Exists(dateKey) Then productDates.Add dateKey, CDbl(navValue)
ContinueRow:
    Next r

    Set BuildNAVDateLookup = result
End Function

Private Function LookupNAV(ByVal navLookup As Object, ByVal trustCode As String, ByVal targetDate As Date) As Variant
    If Not navLookup.Exists(trustCode) Then
        LookupNAV = Empty
        Exit Function
    End If

    Dim productDates As Object
    Set productDates = navLookup.Item(trustCode)
    Dim dateKey As String
    dateKey = Format$(DateOnly(targetDate), "yyyymmdd")
    If productDates.Exists(dateKey) Then
        LookupNAV = productDates(dateKey)
    Else
        LookupNAV = Empty
    End If
End Function

Private Function BuildInceptionDateLookup() As Object
    Dim result As Object
    Set result = CreateTextDictionary()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PRODUCT_INFO)
    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)
    RequireHeader headers, COL_TRUST_CODE, SHEET_PRODUCT_INFO
    RequireHeader headers, COL_INCEPTION_DATE, SHEET_PRODUCT_INFO

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    Dim r As Long
    For r = 2 To lastRow
        Dim trustCode As String
        Dim inceptionDate As Date
        trustCode = NormalizeText(ws.Cells(r, CLng(headers(COL_TRUST_CODE))).Value)
        If Len(trustCode) > 0 And TryReadDate(ws.Cells(r, CLng(headers(COL_INCEPTION_DATE))).Value, inceptionDate) Then
            If Not result.Exists(trustCode) Then result.Add trustCode, inceptionDate
        End If
    Next r

    Set BuildInceptionDateLookup = result
End Function

Private Function ReadBaselineDate(ByVal ws As Worksheet, ByVal headers As Object) As Date
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim found As Boolean
    Dim result As Date
    Dim r As Long
    For r = 2 To lastRow
        Dim candidate As Date
        If TryReadDate(SourceValue(ws, r, headers, COL_BASELINE_DATE), candidate) Then
            If Not found Then
                result = candidate
                found = True
            ElseIf candidate <> result Then
                Err.Raise vbObjectError + 4003, , "产品分类中存在多个基准日期：" & _
                          Format$(result, "yyyy-mm-dd") & "、" & Format$(candidate, "yyyy-mm-dd")
            End If
        End If
    Next r

    If Not found Then Err.Raise vbObjectError + 4004, , "产品分类中没有有效基准日期。"
    ReadBaselineDate = result
End Function

Private Sub RequireSourceHeaders(ByVal headers As Object)
    Dim requiredHeaders As Variant
    requiredHeaders = Array(COL_SEQ, COL_TRUST_CODE, COL_SERIES, COL_CATEGORY, COL_PRODUCT_NAME, _
                            COL_EXPORT_ENABLED, COL_PREV_PREV_OPEN, COL_PREV_OPEN, COL_BASELINE_DATE, _
                            COL_NEXT_OPEN, COL_THEORETICAL_INTERVAL, COL_PREV_INTERVAL, COL_INTERVAL, _
                            COL_ELAPSED, COL_PREV_PREV_NAV, COL_PREV_NAV, COL_BASELINE_NAV, COL_BENCHMARK_RATE)
    Dim i As Long
    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        RequireHeader headers, CStr(requiredHeaders(i)), SHEET_PRODUCT_CATEGORY
    Next i
End Sub

Private Function SourceValue(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal headers As Object, _
                             ByVal headerName As String) As Variant
    If headers.Exists(headerName) Then
        SourceValue = ws.Cells(rowNumber, CLng(headers(headerName))).Value
    Else
        SourceValue = Empty
    End If
End Function

Private Sub WriteOutputValue(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal columnNumber As Long, _
                             ByVal value As Variant)
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Or Len(NormalizeText(value)) = 0 Then
        ws.Cells(rowNumber, columnNumber).ClearContents
    Else
        ws.Cells(rowNumber, columnNumber).Value = value
    End If
End Sub

Private Sub WriteDateValue(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal columnNumber As Long, _
                           ByVal value As Variant)
    Dim parsedDate As Date
    If TryReadDate(value, parsedDate) Then
        ws.Cells(rowNumber, columnNumber).Value = parsedDate
        ws.Cells(rowNumber, columnNumber).NumberFormat = "yyyy-mm-dd"
    Else
        ws.Cells(rowNumber, columnNumber).ClearContents
    End If
End Sub

Private Sub WritePercentValue(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal columnNumber As Long, _
                              ByVal value As Variant)
    If IsNumericValue(value) Then
        ws.Cells(rowNumber, columnNumber).Value = CDbl(value)
        ws.Cells(rowNumber, columnNumber).NumberFormat = "0.00%"
    Else
        ws.Cells(rowNumber, columnNumber).ClearContents
    End If
End Sub

Private Function IsNumericValue(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    IsNumericValue = IsNumeric(value)
End Function

Private Function IsEnabledValue(ByVal value As Variant) As Boolean
    Dim textValue As String
    textValue = UCase$(NormalizeText(value))
    IsEnabledValue = (textValue = "是" Or textValue = "Y" Or textValue = "YES" Or _
                      textValue = "1" Or textValue = "TRUE" Or textValue = "启用")
End Function

Private Sub ReplaceFileSafely(ByVal tempPath As String, ByVal outputPath As String)
    Dim backupPath As String
    backupPath = outputPath & ".previous"

    On Error GoTo ReplaceFail
    If Len(Dir(backupPath)) > 0 Then Kill backupPath
    If Len(Dir(outputPath)) > 0 Then Name outputPath As backupPath
    Name tempPath As outputPath
    If Len(Dir(backupPath)) > 0 Then Kill backupPath
    Exit Sub

ReplaceFail:
    Dim originalError As String
    originalError = Err.Description
    On Error Resume Next
    If Len(Dir(outputPath)) = 0 And Len(Dir(backupPath)) > 0 Then Name backupPath As outputPath
    On Error GoTo 0
    Err.Raise vbObjectError + 4005, , "替换输出文件失败：" & originalError
End Sub

Private Sub RestoreApplicationState(ByVal screenUpdating As Boolean, ByVal enableEvents As Boolean, _
                                    ByVal displayAlerts As Boolean, ByVal calculation As XlCalculation)
    Application.Calculation = calculation
    Application.DisplayAlerts = displayAlerts
    Application.EnableEvents = enableEvents
    Application.ScreenUpdating = screenUpdating
End Sub

Private Sub RequireHeader(ByVal headers As Object, ByVal headerName As String, ByVal sheetName As String)
    If Not headers.Exists(headerName) Then
        Err.Raise vbObjectError + 4006, , sheetName & "工作表缺少字段：" & headerName
    End If
End Sub

Private Function BuildHeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim result As Object
    Set result = CreateTextDictionary()

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)
    Dim c As Long
    For c = 1 To lastCol
        Dim headerText As String
        headerText = NormalizeText(ws.Cells(headerRow, c).Value)
        If Len(headerText) > 0 Then
            If Not result.Exists(headerText) Then result.Add headerText, c
        End If
    Next c

    Set BuildHeaderMap = result
End Function

Private Function CreateTextDictionary() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set CreateTextDictionary = result
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = foundCell.Row
    End If
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = foundCell.Column
    End If
End Function

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

Private Function DateOnly(ByVal value As Date) As Date
    DateOnly = DateSerial(Year(value), Month(value), Day(value))
End Function

Private Function TryReadDate(ByVal value As Variant, ByRef outDate As Date) As Boolean
    If IsDate(value) Then
        outDate = DateOnly(CDate(value))
        TryReadDate = True
        Exit Function
    End If

    Dim textValue As String
    textValue = NormalizeText(value)
    If Len(textValue) = 0 Then Exit Function
    If InStr(textValue, ".") > 0 Then textValue = Left$(textValue, InStr(textValue, ".") - 1)

    If Len(textValue) = 8 And IsNumeric(textValue) Then
        outDate = DateSerial(CInt(Left$(textValue, 4)), CInt(Mid$(textValue, 5, 2)), CInt(Right$(textValue, 2)))
        TryReadDate = True
        Exit Function
    End If

    textValue = Replace(Replace(textValue, ".", "-"), "/", "-")
    If IsDate(textValue) Then
        outDate = DateOnly(CDate(textValue))
        TryReadDate = True
    End If
End Function
