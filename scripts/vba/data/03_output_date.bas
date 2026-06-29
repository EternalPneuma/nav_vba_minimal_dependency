' Data03_ExportProductReport：生成 yyyyMMdd-上层产品分类表现.xlsx 的产品分类表现报告模块

Option Explicit

Private Const TARGET_SHEET_NAME As String = "产品分类"
Private Const NAV_SHEET_NAME As String = "上层产品净值数据(181)"
Private Const PRODUCT_INFO_SHEET_NAME As String = "产品信息"

Private Const COL_TRUST_CODE As String = "信托计划代码"
Private Const COL_NAV_DATE As String = "日期"
Private Const COL_UNIT_NAV As String = "单位净值"
Private Const COL_SEQ As String = "序号"
Private Const COL_CATEGORY As String = "分类"
Private Const COL_SERIES As String = "系列"
Private Const COL_PRODUCT_NAME As String = "产品名称"
Private Const COL_PREV_OPEN As String = "上一开放日"
Private Const COL_PREV_PREV_OPEN As String = "上上一开放日"
Private Const COL_NEXT_OPEN As String = "下一开放日"
Private Const COL_BASELINE_DATE As String = "基准日期"
Private Const COL_BASELINE_NAV As String = "基准日期净值"
Private Const COL_PREV_NAV As String = "上一开放日净值"
Private Const COL_PREV_PREV_NAV As String = "上上一开放日净值"
Private Const COL_THEORETICAL_INTERVAL As String = "理论间隔"
Private Const COL_INTERVAL As String = "实际间隔"
Private Const COL_PREV_INTERVAL As String = "上次开放实际间隔"
Private Const COL_ELAPSED As String = "运作时间"
Private Const COL_BENCHMARK_RATE As String = "基准收益率"
Private Const COL_INCEPTION_DATE As String = "成立日"

Private Const COL_CURRENT_PERIOD_ANNUAL As String = "当前周期年化"
Private Const COL_PREV_PERIOD_ANNUAL As String = "上一周期年化"
Private Const COL_7DAY_ANNUAL As String = "7日年化"
Private Const COL_28DAY_ANNUAL As String = "28日年化"
Private Const COL_INCEPTION_ANNUAL As String = "成立以来年化"
Private Const DEFAULT_INCEPTION_NAV As Double = 1

Private Const CAT_STABLE As String = "稳享长期限"
Private Const CAT_DIRECT As String = "直销"
Private Const CAT_BANK As String = "交行代销"
Private Const CAT_YUANRONG_ANXIANG As String = "圆融安享"

Public Sub Data03_ExportProductReport()
    CalculateClassReportCore
End Sub

' --- Backward-compatible aliases ---
Private Sub CalculateClassReportCore()
    Dim appCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    appCalc = Application.Calculation

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' 1. Build NAV lookup
    Dim navLookup As Object
    Set navLookup = BuildNAVDateLookup()

    Dim inceptionLookup As Object
    Set inceptionLookup = BuildInceptionDateLookup()

    ' 2. Create output workbook
    Dim baselineDate As Date
    Dim wbOutput As Workbook
    Set wbOutput = CreateOutputWorkbook(baselineDate)
    If wbOutput Is Nothing Then
        Err.Raise vbObjectError + 4001, , "输出工作簿创建失败"
    End If

    ' 3. Write headers
    WriteSheetHeaders wbOutput

    ' 4. Write common fields + filter by category
    Dim rowCounters As Object
    WriteCommonFields wbOutput, rowCounters

    ' 5. Category computations
    ComputeStableLongTerm wbOutput, inceptionLookup
    ComputeDirectSales wbOutput, navLookup, inceptionLookup
    ComputeBankAgent wbOutput, navLookup, inceptionLookup
    ComputeYuanRongAnXiang wbOutput, navLookup

    ' 6. Save as .xlsx
    Dim outputPath As String
    outputPath = ThisWorkbook.Path & Application.PathSeparator & Format$(baselineDate, "yyyymmdd") & "-上层产品分类表现.xlsx"

    Application.DisplayAlerts = False
    wbOutput.SaveAs FileName:=outputPath, FileFormat:=xlOpenXMLWorkbook
    wbOutput.Close SaveChanges:=False
    Application.DisplayAlerts = True

    ' 7. Restore settings + report
    Application.Calculation = appCalc
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts

    Dim msg As String
    msg = "产品分类表现报告生成完成" & vbCrLf & vbCrLf & _
          "基准日期：" & Format$(baselineDate, "yyyy-mm-dd") & vbCrLf & vbCrLf & _
          "处理结果：" & vbCrLf & _
          "稳享长期限：" & rowCounters(CAT_STABLE) - 2 & " 条" & vbCrLf & _
          "直销：" & rowCounters(CAT_DIRECT) - 2 & " 条" & vbCrLf & _
          "交行代销：" & rowCounters(CAT_BANK) - 2 & " 条" & vbCrLf & _
          "圆融安享：" & rowCounters(CAT_YUANRONG_ANXIANG) - 2 & " 条" & vbCrLf & vbCrLf & _
          "输出文件：" & vbCrLf & outputPath
    MsgBox msg, vbInformation, "产品分类表现报告"

    Exit Sub

CleanFail:
    Application.Calculation = appCalc
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    MsgBox "产品分类表现报告生成失败" & vbCrLf & vbCrLf & _
           "错误信息：" & Err.Description, vbCritical, "产品分类表现报告"
End Sub

' --- 通用字段写入（产品分类 → 输出工作表） ---

Private Sub WriteCommonFields(ByVal wbOutput As Workbook, ByRef rowCounters As Object)
    Dim srcWs As Worksheet
    Dim headerMap As Object
    Dim lastRow As Long
    Dim r As Long
    Dim outRow As Long
    Dim categoryText As String
    Dim targetSheetName As String
    Dim seqCol As Long
    Dim trustCodeCol As Long
    Dim seriesCol As Long
    Dim categoryCol As Long
    Dim productNameCol As Long
    Dim prevOpenCol As Long
    Dim prevNavCol As Long
    Dim baselineDateCol As Long
    Dim baselineNavCol As Long
    Dim nextOpenCol As Long
    Dim theoreticalIntervalCol As Long
    Dim intervalCol As Long
    Dim elapsedCol As Long
    Dim benchmarkRateCol As Long
    Dim normalizedStable As String
    Dim normalizedDirect As String
    Dim normalizedBank As String
    Dim normalizedYuanRongAnXiang As String
    Dim wsOutput As Worksheet

    If wbOutput Is Nothing Then Exit Sub

    On Error Resume Next
    Set srcWs = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    On Error GoTo 0
    If srcWs Is Nothing Then Exit Sub

    Set headerMap = BuildHeaderMap(srcWs, 1)
    If headerMap Is Nothing Then Exit Sub

    If Not headerMap.Exists(COL_SEQ) Then Exit Sub
    If Not headerMap.Exists(COL_TRUST_CODE) Then Exit Sub
    If Not headerMap.Exists(COL_SERIES) Then Exit Sub
    If Not headerMap.Exists(COL_CATEGORY) Then Exit Sub
    If Not headerMap.Exists(COL_PRODUCT_NAME) Then Exit Sub
    If Not headerMap.Exists(COL_PREV_OPEN) Then Exit Sub
    If Not headerMap.Exists(COL_PREV_NAV) Then Exit Sub
    If Not headerMap.Exists(COL_BASELINE_DATE) Then Exit Sub
    If Not headerMap.Exists(COL_BASELINE_NAV) Then Exit Sub
    If Not headerMap.Exists(COL_NEXT_OPEN) Then Exit Sub
    If Not headerMap.Exists(COL_THEORETICAL_INTERVAL) Then Exit Sub
    If Not headerMap.Exists(COL_INTERVAL) Then Exit Sub
    If Not headerMap.Exists(COL_ELAPSED) Then Exit Sub

    seqCol = CLng(headerMap(COL_SEQ))
    trustCodeCol = CLng(headerMap(COL_TRUST_CODE))
    seriesCol = CLng(headerMap(COL_SERIES))
    categoryCol = CLng(headerMap(COL_CATEGORY))
    productNameCol = CLng(headerMap(COL_PRODUCT_NAME))
    prevOpenCol = CLng(headerMap(COL_PREV_OPEN))
    prevNavCol = CLng(headerMap(COL_PREV_NAV))
    baselineDateCol = CLng(headerMap(COL_BASELINE_DATE))
    baselineNavCol = CLng(headerMap(COL_BASELINE_NAV))
    nextOpenCol = CLng(headerMap(COL_NEXT_OPEN))
    theoreticalIntervalCol = CLng(headerMap(COL_THEORETICAL_INTERVAL))
    intervalCol = CLng(headerMap(COL_INTERVAL))
    elapsedCol = CLng(headerMap(COL_ELAPSED))
    If headerMap.Exists(COL_BENCHMARK_RATE) Then benchmarkRateCol = CLng(headerMap(COL_BENCHMARK_RATE))

    If rowCounters Is Nothing Then
        Set rowCounters = CreateObject("Scripting.Dictionary")
        rowCounters.CompareMode = vbTextCompare
    End If

    normalizedStable = NormalizeText(CAT_STABLE)
    normalizedDirect = NormalizeText(CAT_DIRECT)
    normalizedBank = NormalizeText(CAT_BANK)
    normalizedYuanRongAnXiang = NormalizeText(CAT_YUANRONG_ANXIANG)

    If Not rowCounters.Exists(CAT_STABLE) Then rowCounters.Add CAT_STABLE, 2
    If Not rowCounters.Exists(CAT_DIRECT) Then rowCounters.Add CAT_DIRECT, 2
    If Not rowCounters.Exists(CAT_BANK) Then rowCounters.Add CAT_BANK, 2
    If Not rowCounters.Exists(CAT_YUANRONG_ANXIANG) Then rowCounters.Add CAT_YUANRONG_ANXIANG, 2

    lastRow = LastUsedRow(srcWs)

    For r = 2 To lastRow
        categoryText = NormalizeText(srcWs.Cells(r, categoryCol).Value)
        If Len(categoryText) = 0 Then GoTo ContinueRow

        targetSheetName = vbNullString
        If StrComp(categoryText, normalizedStable, vbTextCompare) = 0 Then
            targetSheetName = CAT_STABLE
        ElseIf StrComp(categoryText, normalizedDirect, vbTextCompare) = 0 Then
            targetSheetName = CAT_DIRECT
        ElseIf StrComp(categoryText, normalizedBank, vbTextCompare) = 0 Then
            targetSheetName = CAT_BANK
        ElseIf StrComp(categoryText, normalizedYuanRongAnXiang, vbTextCompare) = 0 Then
            targetSheetName = CAT_YUANRONG_ANXIANG
        Else
            GoTo ContinueRow
        End If

        If Not rowCounters.Exists(targetSheetName) Then rowCounters.Add targetSheetName, 2
        outRow = CLng(rowCounters(targetSheetName))

        Set wsOutput = wbOutput.Worksheets(targetSheetName)

        If targetSheetName = CAT_YUANRONG_ANXIANG Then
            wsOutput.Cells(outRow, 1).Value = srcWs.Cells(r, seqCol).Value
            wsOutput.Cells(outRow, 2).Value = srcWs.Cells(r, trustCodeCol).Value
            wsOutput.Cells(outRow, 3).Value = srcWs.Cells(r, seriesCol).Value
            wsOutput.Cells(outRow, 4).Value = srcWs.Cells(r, productNameCol).Value
            wsOutput.Cells(outRow, 5).Value = srcWs.Cells(r, baselineDateCol).Value
            wsOutput.Cells(outRow, 6).Value = srcWs.Cells(r, baselineNavCol).Value
            wsOutput.Cells(outRow, 7).Value = srcWs.Cells(r, nextOpenCol).Value
            wsOutput.Cells(outRow, 8).Value = srcWs.Cells(r, theoreticalIntervalCol).Value
            wsOutput.Cells(outRow, 9).Value = srcWs.Cells(r, intervalCol).Value

            wsOutput.Cells(outRow, 5).NumberFormat = "yyyy-mm-dd"
            wsOutput.Cells(outRow, 7).NumberFormat = "yyyy-mm-dd"
        Else
            wsOutput.Cells(outRow, 1).Value = srcWs.Cells(r, seqCol).Value
            wsOutput.Cells(outRow, 2).Value = srcWs.Cells(r, trustCodeCol).Value
            wsOutput.Cells(outRow, 3).Value = srcWs.Cells(r, seriesCol).Value
            wsOutput.Cells(outRow, 4).Value = srcWs.Cells(r, productNameCol).Value
            wsOutput.Cells(outRow, 5).Value = srcWs.Cells(r, prevOpenCol).Value
            wsOutput.Cells(outRow, 6).Value = srcWs.Cells(r, prevNavCol).Value
            wsOutput.Cells(outRow, 7).Value = srcWs.Cells(r, baselineDateCol).Value
            wsOutput.Cells(outRow, 8).Value = srcWs.Cells(r, baselineNavCol).Value
            wsOutput.Cells(outRow, 9).Value = srcWs.Cells(r, nextOpenCol).Value
            wsOutput.Cells(outRow, 10).Value = srcWs.Cells(r, theoreticalIntervalCol).Value
            wsOutput.Cells(outRow, 11).Value = srcWs.Cells(r, intervalCol).Value
            If targetSheetName = CAT_STABLE Then
                wsOutput.Cells(outRow, 12).Value = srcWs.Cells(r, elapsedCol).Value
            End If

            wsOutput.Cells(outRow, 5).NumberFormat = "yyyy-mm-dd"
            wsOutput.Cells(outRow, 7).NumberFormat = "yyyy-mm-dd"
            wsOutput.Cells(outRow, 9).NumberFormat = "yyyy-mm-dd"

            If (targetSheetName = CAT_DIRECT Or targetSheetName = CAT_BANK) And benchmarkRateCol > 0 Then
                wsOutput.Cells(outRow, 12).Value = srcWs.Cells(r, benchmarkRateCol).Value
            End If
        End If

        rowCounters(targetSheetName) = outRow + 1

ContinueRow:
    Next r
End Sub

' --- 交行代销：数据写入 + 7日/28日/双周期/成立以来年化计算 ---
Private Sub ComputeBankAgent(ByVal wbOutput As Workbook, ByVal navLookup As Object, ByVal inceptionLookup As Object)
    Dim wsOutput As Worksheet
    Dim srcWs As Worksheet
    Dim headerMap As Object
    Dim rowLookup As Object
    Dim lastRow As Long
    Dim r As Long
    Dim trustCode As String
    Dim sourceData As Variant
    Dim prevPrevOpen As Variant
    Dim prevPrevNav As Variant
    Dim elapsed As Variant
    Dim prevInterval As Variant
    Dim prevNav As Variant
    Dim baselineNav As Variant
    Dim baselineDate As Variant
    Dim targetDate7 As Date
    Dim targetDate28 As Date
    Dim nav7 As Variant
    Dim nav28 As Variant
    Dim rowKey As String

    If wbOutput Is Nothing Then Exit Sub

    On Error Resume Next
    Set wsOutput = wbOutput.Worksheets("交行代销")
    Set srcWs = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    On Error GoTo 0

    If wsOutput Is Nothing Then Exit Sub
    If srcWs Is Nothing Then Exit Sub

    Set headerMap = BuildHeaderMap(srcWs, 1)
    If headerMap Is Nothing Then Exit Sub
    If Not headerMap.Exists(COL_TRUST_CODE) Then Exit Sub
    If Not headerMap.Exists(COL_PREV_OPEN) Then Exit Sub
    If Not headerMap.Exists(COL_BASELINE_DATE) Then Exit Sub
    If Not headerMap.Exists(COL_NEXT_OPEN) Then Exit Sub
    If Not headerMap.Exists(COL_PREV_PREV_OPEN) Then Exit Sub
    If Not headerMap.Exists(COL_PREV_PREV_NAV) Then Exit Sub
    If Not headerMap.Exists(COL_ELAPSED) Then Exit Sub
    If Not headerMap.Exists(COL_PREV_INTERVAL) Then Exit Sub
    If Not headerMap.Exists(COL_PREV_NAV) Then Exit Sub
    If Not headerMap.Exists(COL_BASELINE_NAV) Then Exit Sub

    Set rowLookup = BuildBankAgentLookup(srcWs, headerMap)
    If rowLookup Is Nothing Then Exit Sub

    lastRow = LastUsedRow(wsOutput)

    For r = 2 To lastRow
        trustCode = NormalizeText(wsOutput.Cells(r, 2).Value)

        prevPrevOpen = Empty
        prevPrevNav = Empty
        elapsed = Empty
        prevInterval = Empty
        prevNav = Empty
        baselineNav = Empty
        baselineDate = wsOutput.Cells(r, 7).Value

        rowKey = BuildReportRowKey(trustCode, wsOutput.Cells(r, 5).Value, baselineDate, wsOutput.Cells(r, 9).Value)
        If Len(rowKey) > 0 Then
            If rowLookup.Exists(rowKey) Then
                sourceData = rowLookup(rowKey)
                prevPrevOpen = sourceData(0)
                prevPrevNav = sourceData(1)
                elapsed = sourceData(2)
                prevInterval = sourceData(3)
                prevNav = sourceData(4)
                baselineNav = sourceData(5)
            End If
        End If

        If IsEmpty(prevPrevOpen) Or IsError(prevPrevOpen) Then
            wsOutput.Cells(r, 13).ClearContents
        Else
            wsOutput.Cells(r, 13).Value = prevPrevOpen
            wsOutput.Cells(r, 13).NumberFormat = "yyyy-mm-dd"
        End If

        If IsEmpty(prevPrevNav) Or IsError(prevPrevNav) Then
            wsOutput.Cells(r, 14).ClearContents
        Else
            wsOutput.Cells(r, 14).Value = prevPrevNav
        End If

        If IsEmpty(elapsed) Or IsError(elapsed) Then
            wsOutput.Cells(r, 15).ClearContents
        Else
            wsOutput.Cells(r, 15).Value = elapsed
        End If

        wsOutput.Cells(r, 16).ClearContents
        If Not IsEmpty(prevPrevNav) And Not IsError(prevPrevNav) _
           And Not IsEmpty(prevNav) And Not IsError(prevNav) _
           And Not IsEmpty(prevInterval) And Not IsError(prevInterval) Then
            If CDbl(prevPrevNav) <> 0 And CDbl(prevInterval) > 0 Then
                wsOutput.Cells(r, 16).Value = (CDbl(prevNav) / CDbl(prevPrevNav) - 1) * (365 / CDbl(prevInterval))
                wsOutput.Cells(r, 16).NumberFormat = "0.00%"
            End If
        End If

        wsOutput.Cells(r, 17).ClearContents
        If Not IsEmpty(baselineNav) And Not IsError(baselineNav) _
           And Not IsEmpty(prevNav) And Not IsError(prevNav) _
           And Not IsEmpty(elapsed) And Not IsError(elapsed) Then
            If CDbl(prevNav) <> 0 And CDbl(elapsed) > 0 Then
                wsOutput.Cells(r, 17).Value = (CDbl(baselineNav) / CDbl(prevNav) - 1) * (365 / CDbl(elapsed))
                wsOutput.Cells(r, 17).NumberFormat = "0.00%"
            End If
        End If

        If Not IsEmpty(wsOutput.Cells(r, 10).Value) And Not IsError(wsOutput.Cells(r, 10).Value) _
           And IsNumeric(wsOutput.Cells(r, 10).Value) _
           And Not IsEmpty(elapsed) And Not IsError(elapsed) _
           And IsNumeric(elapsed) Then
            If CDbl(wsOutput.Cells(r, 10).Value) > 60 And CDbl(elapsed) < 7 Then
                wsOutput.Cells(r, 17).Value = wsOutput.Cells(r, 16).Value
                wsOutput.Cells(r, 17).NumberFormat = "0.00%"
            End If
        End If

        wsOutput.Cells(r, 18).ClearContents
        If Not IsEmpty(baselineDate) And IsDate(baselineDate) _
           And Not IsEmpty(baselineNav) And Not IsError(baselineNav) Then
            targetDate7 = DateAdd("d", -7, CDate(baselineDate))
            nav7 = LookupNAV(navLookup, trustCode, targetDate7)
            If Not IsEmpty(nav7) And Not IsError(nav7) Then
                If CDbl(nav7) <> 0 Then
                    wsOutput.Cells(r, 18).Value = (CDbl(baselineNav) / CDbl(nav7) - 1) * (365 / 7)
                    wsOutput.Cells(r, 18).NumberFormat = "0.00%"
                End If
            End If
        End If

        wsOutput.Cells(r, 19).ClearContents
        If Not IsEmpty(baselineDate) And IsDate(baselineDate) _
           And Not IsEmpty(baselineNav) And Not IsError(baselineNav) Then
            targetDate28 = DateAdd("d", -28, CDate(baselineDate))
            nav28 = LookupNAV(navLookup, trustCode, targetDate28)
            If Not IsEmpty(nav28) And Not IsError(nav28) Then
                If CDbl(nav28) <> 0 Then
                    wsOutput.Cells(r, 19).Value = (CDbl(baselineNav) / CDbl(nav28) - 1) * (365 / 28)
                    wsOutput.Cells(r, 19).NumberFormat = "0.00%"
                End If
            End If
        End If

        WriteInceptionAnnual wsOutput, r, 20, trustCode, baselineNav, baselineDate, inceptionLookup
    Next r
End Sub

Private Function BuildBankAgentLookup(ByVal srcWs As Worksheet, ByVal headerMap As Object) As Object
    Dim dict As Object
    Dim lastRow As Long
    Dim r As Long
    Dim trustCode As String
    Dim rowKey As String
    Dim rowData As Variant
    Dim trustCodeCol As Long
    Dim prevOpenCol As Long
    Dim baselineDateCol As Long
    Dim nextOpenCol As Long
    Dim prevPrevOpenCol As Long
    Dim prevPrevNavCol As Long
    Dim elapsedCol As Long
    Dim prevIntervalCol As Long
    Dim prevNavCol As Long
    Dim baselineNavCol As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    trustCodeCol = CLng(headerMap(COL_TRUST_CODE))
    prevOpenCol = CLng(headerMap(COL_PREV_OPEN))
    baselineDateCol = CLng(headerMap(COL_BASELINE_DATE))
    nextOpenCol = CLng(headerMap(COL_NEXT_OPEN))
    prevPrevOpenCol = CLng(headerMap(COL_PREV_PREV_OPEN))
    prevPrevNavCol = CLng(headerMap(COL_PREV_PREV_NAV))
    elapsedCol = CLng(headerMap(COL_ELAPSED))
    prevIntervalCol = CLng(headerMap(COL_PREV_INTERVAL))
    prevNavCol = CLng(headerMap(COL_PREV_NAV))
    baselineNavCol = CLng(headerMap(COL_BASELINE_NAV))

    lastRow = LastUsedRow(srcWs)

    For r = 2 To lastRow
        trustCode = NormalizeText(srcWs.Cells(r, trustCodeCol).Value)
        If Len(trustCode) = 0 Then GoTo ContinueRow
        rowKey = BuildReportRowKey(trustCode, srcWs.Cells(r, prevOpenCol).Value, srcWs.Cells(r, baselineDateCol).Value, srcWs.Cells(r, nextOpenCol).Value)
        If Len(rowKey) = 0 Then GoTo ContinueRow

        If Not dict.Exists(rowKey) Then
            rowData = Array( _
                srcWs.Cells(r, prevPrevOpenCol).Value, _
                srcWs.Cells(r, prevPrevNavCol).Value, _
                srcWs.Cells(r, elapsedCol).Value, _
                srcWs.Cells(r, prevIntervalCol).Value, _
                srcWs.Cells(r, prevNavCol).Value, _
                srcWs.Cells(r, baselineNavCol).Value)
            dict.Add rowKey, rowData
        End If
ContinueRow:
    Next r

    Set BuildBankAgentLookup = dict
End Function

Private Function BuildReportRowKey(ByVal trustCodeValue As Variant, ByVal prevOpenValue As Variant, _
                                   ByVal baselineDateValue As Variant, ByVal nextOpenValue As Variant) As String
    Dim trustCode As String
    trustCode = NormalizeText(trustCodeValue)
    If Len(trustCode) = 0 Then Exit Function

    BuildReportRowKey = trustCode & "|" & FormatKeyDate(prevOpenValue) & "|" & _
                        FormatKeyDate(baselineDateValue) & "|" & FormatKeyDate(nextOpenValue)
End Function

Private Function FormatKeyDate(ByVal value As Variant) As String
    Dim parsedDate As Date
    If TryReadDate(value, parsedDate) Then
        FormatKeyDate = Format$(parsedDate, "yyyymmdd")
    Else
        FormatKeyDate = NormalizeText(value)
    End If
End Function

Private Function TryReadDate(ByVal value As Variant, ByRef outDate As Date) As Boolean
    If IsDate(value) Then
        outDate = DateValue(CDate(value))
        TryReadDate = True
        Exit Function
    End If

    Dim textValue As String
    textValue = NormalizeText(value)
    If Len(textValue) = 8 And IsNumeric(textValue) Then
        outDate = DateSerial(CInt(Left$(textValue, 4)), CInt(Mid$(textValue, 5, 2)), CInt(Right$(textValue, 2)))
        TryReadDate = True
        Exit Function
    End If

    textValue = Replace(Replace(textValue, ".", "-"), "/", "-")
    If IsDate(textValue) Then
        outDate = DateValue(CDate(textValue))
        TryReadDate = True
    End If
End Function

Private Function ParseYYYYMMDD(ByVal value As Variant, ByRef outDate As Date) As Boolean
    If IsDate(value) Then
        ' Excel 日期序列号
        outDate = DateValue(CDate(value))
        ParseYYYYMMDD = True
        Exit Function
    End If

    Dim textValue As String
    textValue = Trim$(CStr(value))

    ' 去除可能的 .0 后缀（Excel 将整数存储为浮点）
    If InStr(textValue, ".") > 0 Then
        textValue = Left$(textValue, InStr(textValue, ".") - 1)
    End If

    If Len(textValue) = 8 And IsNumeric(textValue) Then
        Dim y As Long: y = CLng(Left$(textValue, 4))
        Dim m As Long: m = CLng(Mid$(textValue, 5, 2))
        Dim d As Long: d = CLng(Right$(textValue, 2))
        If m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
            outDate = DateSerial(y, m, d)
            ParseYYYYMMDD = True
            Exit Function
        End If
    End If

    ParseYYYYMMDD = False
End Function

Private Function BuildHeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim c As Long
    For c = 1 To lastCol
        Dim hdr As String
        hdr = NormalizeText(ws.Cells(headerRow, c).value)
        If Len(hdr) > 0 Then
            If Not dict.Exists(hdr) Then dict.Add hdr, c
        End If
    Next c

    Set BuildHeaderMap = dict
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

' --- NAV 查找基础设施 ---

Private Function BuildNAVDateLookup() As Object
    Dim outerDict As Object
    Set outerDict = CreateObject("Scripting.Dictionary")
    outerDict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(NAV_SHEET_NAME)

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(ws, 1)

    If Not headerMap.Exists(COL_NAV_DATE) Then Exit Function
    If Not headerMap.Exists(COL_TRUST_CODE) Then Exit Function
    If Not headerMap.Exists(COL_UNIT_NAV) Then Exit Function

    Dim dateCol As Long:      dateCol = CLng(headerMap(COL_NAV_DATE))
    Dim trustCodeCol As Long: trustCodeCol = CLng(headerMap(COL_TRUST_CODE))
    Dim navCol As Long:       navCol = CLng(headerMap(COL_UNIT_NAV))

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        Dim cellDate As Variant
        cellDate = ws.Cells(r, dateCol).Value
        If IsEmpty(cellDate) Or IsError(cellDate) Then GoTo ContinueNAV

        Dim parsedDate As Date
        If Not ParseYYYYMMDD(cellDate, parsedDate) Then GoTo ContinueNAV
        Dim dateKey As String
        dateKey = Format$(parsedDate, "yyyymmdd")

        Dim trustCode As Variant
        trustCode = ws.Cells(r, trustCodeCol).Value
        If IsEmpty(trustCode) Or IsError(trustCode) Then GoTo ContinueNAV

        Dim navValue As Variant
        navValue = ws.Cells(r, navCol).Value
        If IsError(navValue) Then GoTo ContinueNAV

        Dim codeKey As String
        codeKey = NormalizeText(trustCode)

        If Not outerDict.Exists(codeKey) Then
            Dim innerDict As Object
            Set innerDict = CreateObject("Scripting.Dictionary")
            innerDict.CompareMode = vbTextCompare
            outerDict.Add codeKey, innerDict
        End If

        ' 同一 (code, date) 保留第一条
        If Not outerDict(codeKey).Exists(dateKey) Then
            outerDict(codeKey).Add dateKey, navValue
        End If
ContinueNAV:
    Next r

    Set BuildNAVDateLookup = outerDict
End Function

Private Function LookupNAV(ByVal navLookup As Object, ByVal trustCodeValue As Variant, ByVal targetDate As Variant) As Variant
    If navLookup Is Nothing Then
        LookupNAV = Empty
        Exit Function
    End If

    If IsEmpty(trustCodeValue) Or IsError(trustCodeValue) Then
        LookupNAV = Empty
        Exit Function
    End If
    If IsEmpty(targetDate) Or IsError(targetDate) Then
        LookupNAV = Empty
        Exit Function
    End If

    Dim codeKey As String
    codeKey = NormalizeText(trustCodeValue)
    If Len(codeKey) = 0 Then
        LookupNAV = Empty
        Exit Function
    End If
    If Not navLookup.Exists(codeKey) Then
        LookupNAV = Empty
        Exit Function
    End If

    Dim dateKey As String
    On Error GoTo LookupFail
    dateKey = Format$(CDate(targetDate), "yyyymmdd")
    On Error GoTo 0
    If Not navLookup(codeKey).Exists(dateKey) Then
        LookupNAV = Empty
        Exit Function
    End If

    LookupNAV = navLookup(codeKey)(dateKey)
    Exit Function

LookupFail:
    On Error GoTo 0
    LookupNAV = Empty
End Function

Private Function BuildInceptionDateLookup() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PRODUCT_INFO_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set BuildInceptionDateLookup = dict
        Exit Function
    End If

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(ws, 1)
    If headerMap Is Nothing Then
        Set BuildInceptionDateLookup = dict
        Exit Function
    End If
    If Not headerMap.Exists(COL_TRUST_CODE) Then
        Set BuildInceptionDateLookup = dict
        Exit Function
    End If
    If Not headerMap.Exists(COL_INCEPTION_DATE) Then
        Set BuildInceptionDateLookup = dict
        Exit Function
    End If

    Dim trustCodeCol As Long
    Dim inceptionDateCol As Long
    trustCodeCol = CLng(headerMap(COL_TRUST_CODE))
    inceptionDateCol = CLng(headerMap(COL_INCEPTION_DATE))

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    Dim trustCode As String
    Dim inceptionDate As Date
    For r = 2 To lastRow
        trustCode = NormalizeText(ws.Cells(r, trustCodeCol).Value)
        If Len(trustCode) = 0 Then GoTo ContinueProductInfo
        If Not TryReadDate(ws.Cells(r, inceptionDateCol).Value, inceptionDate) Then GoTo ContinueProductInfo

        If Not dict.Exists(trustCode) Then dict.Add trustCode, inceptionDate
ContinueProductInfo:
    Next r

    Set BuildInceptionDateLookup = dict
End Function

Private Sub WriteInceptionAnnual(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal outputCol As Long, _
                                 ByVal trustCodeValue As Variant, ByVal baselineNav As Variant, _
                                 ByVal baselineDateValue As Variant, ByVal inceptionLookup As Object)
    ws.Cells(rowNumber, outputCol).ClearContents

    If inceptionLookup Is Nothing Then Exit Sub
    If IsEmpty(trustCodeValue) Or IsError(trustCodeValue) Then Exit Sub
    If IsEmpty(baselineNav) Or IsError(baselineNav) Then Exit Sub
    If IsEmpty(baselineDateValue) Or IsError(baselineDateValue) Then Exit Sub

    Dim trustCode As String
    trustCode = NormalizeText(trustCodeValue)
    If Len(trustCode) = 0 Then Exit Sub
    If Not inceptionLookup.Exists(trustCode) Then Exit Sub

    Dim baselineDate As Date
    If Not TryReadDate(baselineDateValue, baselineDate) Then Exit Sub

    Dim inceptionDate As Date
    inceptionDate = CDate(inceptionLookup(trustCode))

    Dim dayCount As Long
    dayCount = DateDiff("d", inceptionDate, baselineDate) + 1
    If dayCount <= 0 Then Exit Sub

    If DEFAULT_INCEPTION_NAV <> 0 Then
        ws.Cells(rowNumber, outputCol).Value = (CDbl(baselineNav) / DEFAULT_INCEPTION_NAV - 1) * (365 / dayCount)
        ws.Cells(rowNumber, outputCol).NumberFormat = "0.00%"
    End If
End Sub

' --- 输出工作簿创建与表头写入 ---

Private Function CreateOutputWorkbook(ByRef baselineDate As Date) As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim srcWs As Worksheet
    Dim headerMap As Object
    Dim lastRow As Long
    Dim baseCol As Long
    Dim r As Long
    Dim cellValue As Variant

    On Error GoTo CleanFail

    On Error Resume Next
    Set srcWs = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    On Error GoTo CleanFail

    If Not srcWs Is Nothing Then
        Set headerMap = BuildHeaderMap(srcWs, 1)
        If Not headerMap Is Nothing Then
            If headerMap.Exists(COL_BASELINE_DATE) Then
                baseCol = CLng(headerMap(COL_BASELINE_DATE))
                lastRow = LastUsedRow(srcWs)
                For r = 2 To lastRow
                    cellValue = srcWs.Cells(r, baseCol).Value
                    If Not IsEmpty(cellValue) And Not IsError(cellValue) Then
                        If IsDate(cellValue) Then
                            baselineDate = CDate(cellValue)
                            Exit For
                        End If
                    End If
                Next r
            End If
        End If
    End If

    Set wb = Workbooks.Add

    Application.DisplayAlerts = False
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True

    Set ws = wb.Worksheets(1)
    ws.Name = "稳享长期限"

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "直销"

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "交行代销"

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "圆融安享"

    Set CreateOutputWorkbook = wb
    Exit Function

CleanFail:
    Application.DisplayAlerts = True
    Set CreateOutputWorkbook = Nothing
End Function

Private Sub WriteSheetHeaders(ByVal wbOutput As Workbook)
    Dim ws As Worksheet
    Dim headers As Variant
    Dim i As Long

    If wbOutput Is Nothing Then Exit Sub

    Set ws = wbOutput.Worksheets("稳享长期限")
    headers = Array("序号", "信托计划代码", "系列", "产品名称", "上一开放日", "上一开放日净值", "基准日期", "基准日期净值", "下一开放日", "理论间隔", "实际间隔", "运作时间", "当前周期年化", "成立以来年化")
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit

    Set ws = wbOutput.Worksheets("直销")
    headers = Array("序号", "信托计划代码", "系列", "产品名称", "上一开放日", "上一开放日净值", "基准日期", "基准日期净值", "下一开放日", "理论间隔", "实际间隔", "基准收益率", "7日年化", "28日年化", "成立以来年化")
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit

    Set ws = wbOutput.Worksheets("交行代销")
    headers = Array("序号", "信托计划代码", "系列", "产品名称", "上一开放日", "上一开放日净值", "基准日期", "基准日期净值", "下一开放日", "理论间隔", "实际间隔", "基准收益率", "上上一开放日", "上上一开放日净值", "运作时间", "上一周期年化", "当前周期年化", "7日年化", "28日年化", "成立以来年化")
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit

    Set ws = wbOutput.Worksheets("圆融安享")
    headers = Array("序号", "信托计划代码", "系列", "产品名称", "基准日期", "基准日期净值", "下一开放日", "理论间隔", "实际间隔", "7日年化", "28日年化")
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
End Sub

' --- 稳享长期限：当前周期年化 + 成立以来年化计算 ---
Private Sub ComputeStableLongTerm(ByVal wbOutput As Workbook, ByVal inceptionLookup As Object)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim trustCode As String
    Dim baselineNav As Variant
    Dim prevNav As Variant
    Dim elapsedValue As Variant

    Set ws = wbOutput.Worksheets("稳享长期限")
    
    ' Output columns (1-based):
    ' Col 1=序号, 2=信托计划代码, 3=系列, 4=产品名称
    ' Col 5=上一开放日, 6=上一开放日净值
    ' Col 7=基准日期, 8=基准日期净值
    ' Col 9=下一开放日, 10=理论间隔, 11=实际间隔, 12=运作时间, 13=当前周期年化, 14=成立以来年化
    
    lastRow = LastUsedRow(ws)
    
    For r = 2 To lastRow
        trustCode = NormalizeText(ws.Cells(r, 2).Value)

        ws.Cells(r, 13).ClearContents
        baselineNav = ws.Cells(r, 8).Value
        prevNav = ws.Cells(r, 6).Value
        elapsedValue = ws.Cells(r, 12).Value
        
        If Not IsEmpty(baselineNav) And Not IsError(baselineNav) Then
            If Not IsEmpty(prevNav) And Not IsError(prevNav) Then
                If Not IsEmpty(elapsedValue) And Not IsError(elapsedValue) Then
                    If CDbl(prevNav) <> 0 And CLng(elapsedValue) > 0 Then
                        ws.Cells(r, 13).Value = (CDbl(baselineNav) / CDbl(prevNav) - 1) * (365 / CLng(elapsedValue))
                        ws.Cells(r, 13).NumberFormat = "0.00%"
                    End If
                End If
            End If
        End If

        WriteInceptionAnnual ws, r, 14, trustCode, baselineNav, ws.Cells(r, 7).Value, inceptionLookup
    Next r
End Sub

' --- 直销：7日/28日年化 + 成立以来年化计算 ---
Private Sub ComputeDirectSales(ByVal wbOutput As Workbook, ByVal navLookup As Object, ByVal inceptionLookup As Object)
    Dim ws As Worksheet
    Set ws = wbOutput.Worksheets("直销")
    
    ' Col 1=序号, 2=信托计划代码, 3=系列, 4=产品名称
    ' Col 5=上一开放日, 6=上一开放日净值
    ' Col 7=基准日期, 8=基准日期净值
    ' Col 9=下一开放日, 10=理论间隔, 11=实际间隔, 12=基准收益率, 13=7日年化, 14=28日年化, 15=成立以来年化
    
    Dim lastRow As Long, r As Long
    Dim trustCode As Variant, baselineDate As Variant, baselineNav As Variant
    Dim targetDate7 As Date, targetDate28 As Date
    Dim nav7 As Variant, nav28 As Variant
    
    lastRow = LastUsedRow(ws)
    
    For r = 2 To lastRow
        trustCode = ws.Cells(r, 2).Value
        baselineDate = ws.Cells(r, 7).Value
        baselineNav = ws.Cells(r, 8).Value
        
        If IsEmpty(trustCode) Or IsError(trustCode) Then GoTo ContinueDS
        If IsEmpty(baselineDate) Or Not IsDate(baselineDate) Then GoTo ContinueDS
        If IsEmpty(baselineNav) Or IsError(baselineNav) Then GoTo ContinueDS
        
        ' 7日年化
        targetDate7 = DateAdd("d", -7, CDate(baselineDate))
        nav7 = LookupNAV(navLookup, trustCode, targetDate7)
        If Not IsEmpty(nav7) And Not IsError(nav7) Then
            If CDbl(nav7) <> 0 Then
                ws.Cells(r, 13).Value = (CDbl(baselineNav) / CDbl(nav7) - 1) * (365 / 7)
                ws.Cells(r, 13).NumberFormat = "0.00%"
            End If
        End If
        
        ' 28日年化
        targetDate28 = DateAdd("d", -28, CDate(baselineDate))
        nav28 = LookupNAV(navLookup, trustCode, targetDate28)
        If Not IsEmpty(nav28) And Not IsError(nav28) Then
            If CDbl(nav28) <> 0 Then
                ws.Cells(r, 14).Value = (CDbl(baselineNav) / CDbl(nav28) - 1) * (365 / 28)
                ws.Cells(r, 14).NumberFormat = "0.00%"
            End If
        End If
        
        WriteInceptionAnnual ws, r, 15, trustCode, baselineNav, baselineDate, inceptionLookup
         
ContinueDS:
    Next r
End Sub

' --- 圆融安享：7日/28日年化计算 ---
Private Sub ComputeYuanRongAnXiang(ByVal wbOutput As Workbook, ByVal navLookup As Object)
    Dim ws As Worksheet
    Set ws = wbOutput.Worksheets("圆融安享")

    ' Col 1=序号, 2=信托计划代码, 3=系列, 4=产品名称
    ' Col 5=基准日期, 6=基准日期净值, 7=下一开放日, 8=理论间隔, 9=实际间隔, 10=7日年化, 11=28日年化

    Dim lastRow As Long, r As Long
    Dim trustCode As Variant, baselineDate As Variant, baselineNav As Variant
    Dim targetDate7 As Date, targetDate28 As Date
    Dim nav7 As Variant, nav28 As Variant

    lastRow = LastUsedRow(ws)

    For r = 2 To lastRow
        trustCode = ws.Cells(r, 2).Value
        baselineDate = ws.Cells(r, 5).Value
        baselineNav = ws.Cells(r, 6).Value

        If IsEmpty(trustCode) Or IsError(trustCode) Then GoTo ContinueYRA
        If IsEmpty(baselineDate) Or Not IsDate(baselineDate) Then GoTo ContinueYRA
        If IsEmpty(baselineNav) Or IsError(baselineNav) Then GoTo ContinueYRA

        ws.Cells(r, 10).ClearContents
        targetDate7 = DateAdd("d", -7, CDate(baselineDate))
        nav7 = LookupNAV(navLookup, trustCode, targetDate7)
        If Not IsEmpty(nav7) And Not IsError(nav7) Then
            If CDbl(nav7) <> 0 Then
                ws.Cells(r, 10).Value = (CDbl(baselineNav) / CDbl(nav7) - 1) * (365 / 7)
                ws.Cells(r, 10).NumberFormat = "0.00%"
            End If
        End If

        ws.Cells(r, 11).ClearContents
        targetDate28 = DateAdd("d", -28, CDate(baselineDate))
        nav28 = LookupNAV(navLookup, trustCode, targetDate28)
        If Not IsEmpty(nav28) And Not IsError(nav28) Then
            If CDbl(nav28) <> 0 Then
                ws.Cells(r, 11).Value = (CDbl(baselineNav) / CDbl(nav28) - 1) * (365 / 28)
                ws.Cells(r, 11).NumberFormat = "0.00%"
            End If
        End If

ContinueYRA:
    Next r
End Sub
