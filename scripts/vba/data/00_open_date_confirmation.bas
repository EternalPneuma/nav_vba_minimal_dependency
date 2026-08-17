' OpenDate：开放日推算、人工确认与台账写入

Option Explicit

Private Const SHEET_PRODUCT_CATEGORY As String = "产品分类"
Private Const SHEET_OPEN_DATES As String = "开放日"
Private Const SHEET_NAV As String = "上层产品净值数据(181)"
Private Const SHEET_PENDING As String = "开放日待确认"

Private Const COL_SEQ As String = "序号"
Private Const COL_TRUST_CODE As String = "信托计划代码"
Private Const COL_PRODUCT_NAME As String = "产品名称"
Private Const COL_INTERVAL As String = "理论间隔"
Private Const COL_OPEN_DATE As String = "开放日"
Private Const COL_NAV_DATE As String = "日期"
Private Const COL_EXPORT_ENABLED As String = "是否导出"

Private Const PENDING_COL_BASELINE As String = "基准日期"
Private Const PENDING_COL_ANCHOR As String = "最近已知开放日"
Private Const PENDING_COL_PROPOSED As String = "推算开放日"
Private Const PENDING_COL_STATUS As String = "确认状态"
Private Const PENDING_COL_NOTE As String = "说明"

Private Const STATUS_PENDING As String = "待确认"
Private Const STATUS_CONFIRM As String = "确认"
Private Const STATUS_REJECT As String = "拒绝"
Private Const STATUS_WRITTEN As String = "已写入"
Private Const STATUS_MANUAL As String = "待人工处理"
Private Const STATUS_RESOLVED As String = "已人工处理"

Private Const PRODUCT_CODE_6_MONTH_101 As String = "P83600"
Private Const PRODUCT_CODE_6_MONTH_102 As String = "P83800"

Public Function OpenDate_EnsureCurrentBaselineReady() As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    oldCalculation = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents

    On Error GoTo CleanFail
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim resultMessage As String
    OpenDate_EnsureCurrentBaselineReady = OpenDate_EnsureReady(GetBaselineDateFromNAV(), resultMessage)

    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation
    Exit Function

CleanFail:
    Dim errorNumber As Long
    Dim errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation
    Err.Raise errorNumber, , errorDescription
End Function

Public Function OpenDate_EnsureReady(ByVal baselineDate As Date, ByRef resultMessage As String) As Boolean
    Dim currentStage As String
    On Error GoTo EnsureFail

    currentStage = "创建待确认工作表"
    Dim wsPending As Worksheet
    Set wsPending = EnsurePendingSheet()

    currentStage = "写入已确认开放日"
    Dim appliedCount As Long
    Dim rejectedCount As Long
    ApplyReviewedProposals wsPending, appliedCount, rejectedCount

    currentStage = "生成开放日提议"
    Dim createdCount As Long
    Dim manualCount As Long
    CreateMissingProposals wsPending, baselineDate, createdCount, manualCount

    currentStage = "核对人工补录"
    ReconcileManuallyResolvedProposals wsPending, baselineDate

    currentStage = "统计待处理记录"
    Dim unresolvedCount As Long
    unresolvedCount = CountUnresolvedForBaseline(wsPending, baselineDate)

    resultMessage = "开放日确认状态" & vbCrLf & vbCrLf & _
                    "基准日期：" & Format$(baselineDate, "yyyy-mm-dd") & vbCrLf & _
                    "本次写入台账：" & appliedCount & vbCrLf & _
                    "本次新增提议：" & createdCount & vbCrLf & _
                    "缺少推算锚点：" & manualCount & vbCrLf & _
                    "仍待处理：" & unresolvedCount

    If rejectedCount > 0 Then
        resultMessage = resultMessage & vbCrLf & "已拒绝但尚未人工补充：" & rejectedCount
    End If

    If unresolvedCount > 0 Then
        resultMessage = resultMessage & vbCrLf & vbCrLf & _
                        "请在“" & SHEET_PENDING & "”工作表核对推算结果，将确认状态改为“确认”或“拒绝”。" & vbCrLf & _
                        "拒绝后还需在“" & SHEET_OPEN_DATES & "”中人工补充正确日期。"
        OpenDate_EnsureReady = False
    Else
        OpenDate_EnsureReady = True
    End If
    Exit Function

EnsureFail:
    Dim errorNumber As Long
    Dim errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, , currentStage & "失败：" & errorDescription
End Function

Public Sub OpenDate_RunStandalone()
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim baselineDate As Date
    baselineDate = GetBaselineDateFromNAV()

    Dim resultMessage As String
    Dim isReady As Boolean
    isReady = OpenDate_EnsureReady(baselineDate, resultMessage)

    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If isReady Then
        MsgBox resultMessage & vbCrLf & vbCrLf & "开放日台账已就绪。", vbInformation, "开放日补算"
    Else
        ThisWorkbook.Worksheets(SHEET_PENDING).Activate
        MsgBox resultMessage, vbExclamation, "开放日补算"
    End If
    Exit Sub

CleanFail:
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "开放日补算失败" & vbCrLf & vbCrLf & "错误信息：" & Err.Description, _
           vbCritical, "开放日补算"
End Sub

Private Sub ApplyReviewedProposals(ByVal wsPending As Worksheet, ByRef appliedCount As Long, _
                                   ByRef rejectedCount As Long)
    Dim currentStage As String
    On Error GoTo ApplyFail

    currentStage = "读取待确认表头"
    Dim pendingHeaders As Object
    Set pendingHeaders = BuildHeaderMap(wsPending, 1)

    currentStage = "读取开放日表头"
    Dim wsOpen As Worksheet
    Set wsOpen = ThisWorkbook.Worksheets(SHEET_OPEN_DATES)

    Dim openHeaders As Object
    Set openHeaders = BuildHeaderMap(wsOpen, 1)
    RequireHeader openHeaders, COL_SEQ, SHEET_OPEN_DATES
    RequireHeader openHeaders, COL_OPEN_DATE, SHEET_OPEN_DATES

    currentStage = "建立开放日索引"
    Dim openIndex As Object
    Set openIndex = BuildOpenDayIndex(wsOpen, openHeaders)

    currentStage = "处理确认状态"
    Dim lastRow As Long
    lastRow = LastUsedRow(wsPending)

    Dim r As Long
    For r = 2 To lastRow
        Dim statusText As String
        statusText = NormalizeText(wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_STATUS))).Value)

        If statusText = STATUS_CONFIRM Then
            Dim seqKey As String
            Dim proposedDate As Date
            seqKey = NormalizeText(wsPending.Cells(r, CLng(pendingHeaders(COL_SEQ))).Value)

            If Len(seqKey) = 0 Then
                wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_NOTE))).Value = "无法写入：缺少序号"
            ElseIf Not TryReadDate(wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_PROPOSED))).Value, proposedDate) Then
                wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_NOTE))).Value = "无法写入：推算开放日无效"
            Else
                If AppendOpenDayIfMissing(wsOpen, openHeaders, openIndex, seqKey, proposedDate) Then
                    appliedCount = appliedCount + 1
                End If
                wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_STATUS))).Value = STATUS_WRITTEN
                wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_NOTE))).Value = "已写入开放日台账"
            End If
        ElseIf statusText = STATUS_REJECT Then
            rejectedCount = rejectedCount + 1
            wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_NOTE))).Value = "已拒绝；请在开放日台账人工补充正确日期"
        ElseIf Len(statusText) > 0 And statusText <> STATUS_PENDING And _
               statusText <> STATUS_MANUAL And statusText <> STATUS_WRITTEN And _
               statusText <> STATUS_RESOLVED Then
            wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_NOTE))).Value = "确认状态无效，只允许确认、拒绝、待确认、待人工处理"
        End If
    Next r
    Exit Sub

ApplyFail:
    Dim errorNumber As Long
    Dim errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, , currentStage & "失败：" & errorDescription
End Sub

Private Sub CreateMissingProposals(ByVal wsPending As Worksheet, ByVal baselineDate As Date, _
                                   ByRef createdCount As Long, ByRef manualCount As Long)
    Dim wsProducts As Worksheet
    Dim wsOpen As Worksheet
    Set wsProducts = ThisWorkbook.Worksheets(SHEET_PRODUCT_CATEGORY)
    Set wsOpen = ThisWorkbook.Worksheets(SHEET_OPEN_DATES)

    Dim productHeaders As Object
    Dim openHeaders As Object
    Dim pendingHeaders As Object
    Set productHeaders = BuildHeaderMap(wsProducts, 1)
    Set openHeaders = BuildHeaderMap(wsOpen, 1)
    Set pendingHeaders = BuildHeaderMap(wsPending, 1)

    RequireHeader productHeaders, COL_SEQ, SHEET_PRODUCT_CATEGORY
    RequireHeader productHeaders, COL_TRUST_CODE, SHEET_PRODUCT_CATEGORY
    RequireHeader productHeaders, COL_PRODUCT_NAME, SHEET_PRODUCT_CATEGORY
    RequireHeader productHeaders, COL_INTERVAL, SHEET_PRODUCT_CATEGORY
    RequireHeader openHeaders, COL_SEQ, SHEET_OPEN_DATES
    RequireHeader openHeaders, COL_OPEN_DATE, SHEET_OPEN_DATES

    Dim openIndex As Object
    Set openIndex = BuildOpenDayIndex(wsOpen, openHeaders)

    Dim proposalKeys As Object
    Set proposalKeys = LoadProposalKeys(wsPending, pendingHeaders)

    Dim lastRow As Long
    lastRow = LastUsedRow(wsProducts)

    Dim r As Long
    For r = 2 To lastRow
        If Not ProductIsEnabled(wsProducts, r, productHeaders) Then GoTo ContinueProduct

        Dim seqKey As String
        Dim intervalText As String
        seqKey = NormalizeText(wsProducts.Cells(r, CLng(productHeaders(COL_SEQ))).Value)
        intervalText = NormalizeText(wsProducts.Cells(r, CLng(productHeaders(COL_INTERVAL))).Value)
        If Len(seqKey) = 0 Or Len(intervalText) = 0 Or Not IsNumeric(intervalText) Then GoTo ContinueProduct

        Dim intervalDays As Long
        Dim trustCode As String
        intervalDays = CLng(intervalText)
        trustCode = NormalizeText(wsProducts.Cells(r, CLng(productHeaders(COL_TRUST_CODE))).Value)
        If intervalDays <= 0 Or intervalDays = 1 Then GoTo ContinueProduct

        Dim nextKnownDate As Variant
        Dim anchorDate As Variant
        nextKnownDate = Empty
        anchorDate = Empty
        If openIndex.Exists(seqKey) Then
            Dim productOpenDates As Collection
            Set productOpenDates = openIndex(seqKey)
            nextKnownDate = FindNextDate(productOpenDates, baselineDate)
            anchorDate = FindAnchorDate(productOpenDates, baselineDate)
        End If
        If Not IsEmpty(nextKnownDate) Then GoTo ContinueProduct

        Dim proposedDate As Variant
        Dim noteText As String
        proposedDate = InferNextOpenDate(baselineDate, intervalDays, anchorDate, trustCode)

        If IsEmpty(proposedDate) Then
            noteText = "没有最近已知开放日，无法按理论间隔推算"
            If AddProposalIfMissing(wsPending, pendingHeaders, proposalKeys, wsProducts, r, productHeaders, _
                                    baselineDate, anchorDate, proposedDate, STATUS_MANUAL, noteText) Then
                createdCount = createdCount + 1
                manualCount = manualCount + 1
            End If
        Else
            noteText = BuildInferenceNote(intervalDays, trustCode)
            If AddProposalIfMissing(wsPending, pendingHeaders, proposalKeys, wsProducts, r, productHeaders, _
                                    baselineDate, anchorDate, proposedDate, STATUS_PENDING, noteText) Then
                createdCount = createdCount + 1
            End If
        End If
ContinueProduct:
    Next r

    wsPending.Columns.AutoFit
End Sub

Private Function AddProposalIfMissing(ByVal wsPending As Worksheet, ByVal pendingHeaders As Object, _
                                      ByVal proposalKeys As Object, ByVal wsProducts As Worksheet, _
                                      ByVal productRow As Long, ByVal productHeaders As Object, _
                                      ByVal baselineDate As Date, ByVal anchorDate As Variant, _
                                      ByVal proposedDate As Variant, ByVal statusText As String, _
                                      ByVal noteText As String) As Boolean
    Dim seqKey As String
    Dim proposalKey As String
    seqKey = NormalizeText(wsProducts.Cells(productRow, CLng(productHeaders(COL_SEQ))).Value)
    proposalKey = BuildProposalKey(seqKey, baselineDate)
    If proposalKeys.Exists(proposalKey) Then Exit Function

    Dim nextRow As Long
    nextRow = LastUsedRow(wsPending) + 1

    wsPending.Cells(nextRow, CLng(pendingHeaders(COL_SEQ))).Value = wsProducts.Cells(productRow, CLng(productHeaders(COL_SEQ))).Value
    wsPending.Cells(nextRow, CLng(pendingHeaders(COL_TRUST_CODE))).Value = wsProducts.Cells(productRow, CLng(productHeaders(COL_TRUST_CODE))).Value
    wsPending.Cells(nextRow, CLng(pendingHeaders(COL_PRODUCT_NAME))).Value = wsProducts.Cells(productRow, CLng(productHeaders(COL_PRODUCT_NAME))).Value
    wsPending.Cells(nextRow, CLng(pendingHeaders(COL_INTERVAL))).Value = wsProducts.Cells(productRow, CLng(productHeaders(COL_INTERVAL))).Value
    wsPending.Cells(nextRow, CLng(pendingHeaders(PENDING_COL_BASELINE))).Value = baselineDate
    wsPending.Cells(nextRow, CLng(pendingHeaders(PENDING_COL_BASELINE))).NumberFormat = "yyyy-mm-dd"

    If Not IsEmpty(anchorDate) Then
        wsPending.Cells(nextRow, CLng(pendingHeaders(PENDING_COL_ANCHOR))).Value = CDate(anchorDate)
        wsPending.Cells(nextRow, CLng(pendingHeaders(PENDING_COL_ANCHOR))).NumberFormat = "yyyy-mm-dd"
    End If
    If Not IsEmpty(proposedDate) Then
        wsPending.Cells(nextRow, CLng(pendingHeaders(PENDING_COL_PROPOSED))).Value = CDate(proposedDate)
        wsPending.Cells(nextRow, CLng(pendingHeaders(PENDING_COL_PROPOSED))).NumberFormat = "yyyy-mm-dd"
    End If

    wsPending.Cells(nextRow, CLng(pendingHeaders(PENDING_COL_STATUS))).Value = statusText
    wsPending.Cells(nextRow, CLng(pendingHeaders(PENDING_COL_NOTE))).Value = noteText
    proposalKeys.Add proposalKey, True
    AddProposalIfMissing = True
End Function

Private Function CountUnresolvedForBaseline(ByVal wsPending As Worksheet, ByVal baselineDate As Date) As Long
    Dim headers As Object
    Set headers = BuildHeaderMap(wsPending, 1)

    Dim lastRow As Long
    lastRow = LastUsedRow(wsPending)

    Dim r As Long
    For r = 2 To lastRow
        Dim rowBaseline As Date
        If TryReadDate(wsPending.Cells(r, CLng(headers(PENDING_COL_BASELINE))).Value, rowBaseline) Then
            If DateOnly(rowBaseline) = DateOnly(baselineDate) Then
                Dim statusText As String
                statusText = NormalizeText(wsPending.Cells(r, CLng(headers(PENDING_COL_STATUS))).Value)
                If statusText <> STATUS_WRITTEN And statusText <> STATUS_RESOLVED Then
                    CountUnresolvedForBaseline = CountUnresolvedForBaseline + 1
                End If
            End If
        End If
    Next r
End Function

Private Sub ReconcileManuallyResolvedProposals(ByVal wsPending As Worksheet, ByVal baselineDate As Date)
    Dim pendingHeaders As Object
    Set pendingHeaders = BuildHeaderMap(wsPending, 1)

    Dim wsOpen As Worksheet
    Set wsOpen = ThisWorkbook.Worksheets(SHEET_OPEN_DATES)
    Dim openHeaders As Object
    Set openHeaders = BuildHeaderMap(wsOpen, 1)
    Dim openIndex As Object
    Set openIndex = BuildOpenDayIndex(wsOpen, openHeaders)

    Dim lastRow As Long
    lastRow = LastUsedRow(wsPending)

    Dim r As Long
    For r = 2 To lastRow
        Dim rowBaseline As Date
        If Not TryReadDate(wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_BASELINE))).Value, rowBaseline) Then GoTo ContinueRow
        If DateOnly(rowBaseline) <> DateOnly(baselineDate) Then GoTo ContinueRow

        Dim statusText As String
        statusText = NormalizeText(wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_STATUS))).Value)
        If statusText = STATUS_WRITTEN Or statusText = STATUS_RESOLVED Then GoTo ContinueRow

        Dim seqKey As String
        seqKey = NormalizeText(wsPending.Cells(r, CLng(pendingHeaders(COL_SEQ))).Value)
        If openIndex.Exists(seqKey) Then
            Dim productOpenDates As Collection
            Set productOpenDates = openIndex(seqKey)
            If Not IsEmpty(FindNextDate(productOpenDates, baselineDate)) Then
                wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_STATUS))).Value = STATUS_RESOLVED
                wsPending.Cells(r, CLng(pendingHeaders(PENDING_COL_NOTE))).Value = "开放日台账已人工补充未来开放日"
            End If
        End If
ContinueRow:
    Next r
End Sub

Private Function InferNextOpenDate(ByVal baselineDate As Date, ByVal intervalDays As Long, _
                                   ByVal anchorDate As Variant, ByVal trustCode As String) As Variant
    If intervalDays = 7 Then
        InferNextOpenDate = GetNextWednesday(baselineDate)
    ElseIf intervalDays = 183 And trustCode = PRODUCT_CODE_6_MONTH_101 Then
        InferNextOpenDate = GetNextMonthEndWednesday(baselineDate)
    ElseIf intervalDays = 183 And trustCode = PRODUCT_CODE_6_MONTH_102 Then
        InferNextOpenDate = GetNextSecondWednesday(baselineDate)
    ElseIf IsEmpty(anchorDate) Then
        InferNextOpenDate = Empty
    Else
        Dim candidate As Date
        candidate = CDate(anchorDate)
        Do
            candidate = DateAdd("d", intervalDays, candidate)
        Loop While candidate <= DateOnly(baselineDate)
        InferNextOpenDate = candidate
    End If
End Function

Private Function FindNextDate(ByVal dates As Collection, ByVal baselineDate As Date) As Variant
    Dim item As Variant
    Dim result As Variant
    result = Empty
    For Each item In dates
        Dim currentDate As Date
        currentDate = DateOnly(CDate(item))
        If currentDate > DateOnly(baselineDate) Then
            If IsEmpty(result) Or currentDate < CDate(result) Then result = currentDate
        End If
    Next item
    FindNextDate = result
End Function

Private Function FindAnchorDate(ByVal dates As Collection, ByVal baselineDate As Date) As Variant
    Dim item As Variant
    Dim result As Variant
    result = Empty
    For Each item In dates
        Dim currentDate As Date
        currentDate = DateOnly(CDate(item))
        If currentDate <= DateOnly(baselineDate) Then
            If IsEmpty(result) Or currentDate > CDate(result) Then result = currentDate
        End If
    Next item
    FindAnchorDate = result
End Function

Private Function GetNextWednesday(ByVal baselineDate As Date) As Date
    Dim daysToAdd As Long
    daysToAdd = 3 - Weekday(baselineDate, vbMonday)
    If daysToAdd <= 0 Then daysToAdd = daysToAdd + 7
    GetNextWednesday = DateAdd("d", daysToAdd, DateOnly(baselineDate))
End Function

Private Function GetLastWednesdayOfMonth(ByVal targetDate As Date) As Date
    Dim lastDay As Date
    lastDay = DateSerial(Year(targetDate), Month(targetDate) + 1, 0)

    Dim daysBack As Long
    daysBack = Weekday(lastDay, vbMonday) - 3
    If daysBack < 0 Then daysBack = daysBack + 7
    GetLastWednesdayOfMonth = DateAdd("d", -daysBack, lastDay)
End Function

Private Function GetNextMonthEndWednesday(ByVal baselineDate As Date) As Date
    Dim currentMonthWednesday As Date
    currentMonthWednesday = GetLastWednesdayOfMonth(baselineDate)

    If currentMonthWednesday > DateOnly(baselineDate) Then
        GetNextMonthEndWednesday = currentMonthWednesday
    Else
        GetNextMonthEndWednesday = GetLastWednesdayOfMonth(DateAdd("m", 1, baselineDate))
    End If
End Function

Private Function GetSecondWednesdayOfMonth(ByVal targetDate As Date) As Date
    Dim firstDay As Date
    firstDay = DateSerial(Year(targetDate), Month(targetDate), 1)

    Dim daysToFirstWednesday As Long
    daysToFirstWednesday = 3 - Weekday(firstDay, vbMonday)
    If daysToFirstWednesday < 0 Then daysToFirstWednesday = daysToFirstWednesday + 7
    GetSecondWednesdayOfMonth = DateAdd("d", daysToFirstWednesday + 7, firstDay)
End Function

Private Function GetNextSecondWednesday(ByVal baselineDate As Date) As Date
    Dim currentMonthWednesday As Date
    currentMonthWednesday = GetSecondWednesdayOfMonth(baselineDate)

    If currentMonthWednesday > DateOnly(baselineDate) Then
        GetNextSecondWednesday = currentMonthWednesday
    Else
        GetNextSecondWednesday = GetSecondWednesdayOfMonth(DateAdd("m", 1, baselineDate))
    End If
End Function

Private Function BuildInferenceNote(ByVal intervalDays As Long, ByVal trustCode As String) As String
    If intervalDays = 7 Then
        BuildInferenceNote = "周开规则：取下一周周三"
    ElseIf intervalDays = 183 And trustCode = PRODUCT_CODE_6_MONTH_101 Then
        BuildInferenceNote = "6个月101特殊规则：取本月或下月最后一个周三；如遇节假日请顺延后确认"
    ElseIf intervalDays = 183 And trustCode = PRODUCT_CODE_6_MONTH_102 Then
        BuildInferenceNote = "6个月102特殊规则：取本月或下月第二个周三；如遇节假日请顺延后确认"
    Else
        BuildInferenceNote = "从最近已知开放日起按" & intervalDays & "个自然日滚动"
    End If
End Function

Private Function EnsurePendingSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PENDING)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_PENDING
        Dim headers As Variant
        headers = Array(COL_SEQ, COL_TRUST_CODE, COL_PRODUCT_NAME, COL_INTERVAL, _
                        PENDING_COL_BASELINE, PENDING_COL_ANCHOR, PENDING_COL_PROPOSED, _
                        PENDING_COL_STATUS, PENDING_COL_NOTE)
        Dim i As Long
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
        ws.Rows(1).Font.Bold = True
        ws.Cells.Font.Name = "微软雅黑"
        ws.Columns.AutoFit
    Else
        Dim headerMap As Object
        Set headerMap = BuildHeaderMap(ws, 1)
        RequireHeader headerMap, COL_SEQ, SHEET_PENDING
        RequireHeader headerMap, COL_TRUST_CODE, SHEET_PENDING
        RequireHeader headerMap, COL_PRODUCT_NAME, SHEET_PENDING
        RequireHeader headerMap, COL_INTERVAL, SHEET_PENDING
        RequireHeader headerMap, PENDING_COL_BASELINE, SHEET_PENDING
        RequireHeader headerMap, PENDING_COL_ANCHOR, SHEET_PENDING
        RequireHeader headerMap, PENDING_COL_PROPOSED, SHEET_PENDING
        RequireHeader headerMap, PENDING_COL_STATUS, SHEET_PENDING
        RequireHeader headerMap, PENDING_COL_NOTE, SHEET_PENDING
    End If

    ApplyPendingStatusValidation ws
    Set EnsurePendingSheet = ws
End Function

Private Sub ApplyPendingStatusValidation(ByVal ws As Worksheet)
    Dim reviewName As Name
    On Error Resume Next
    Set reviewName = ThisWorkbook.Names("OpenDateReviewValues")
    On Error GoTo 0
    If reviewName Is Nothing Then Exit Sub

    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)
    Dim targetRange As Range
    Set targetRange = ws.Range(ws.Cells(2, CLng(headers(PENDING_COL_STATUS))), _
                               ws.Cells(5000, CLng(headers(PENDING_COL_STATUS))))
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=OpenDateReviewValues"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With
End Sub

Private Function BuildOpenDayIndex(ByVal ws As Worksheet, ByVal headers As Object) As Object
    Dim currentStage As String
    On Error GoTo IndexFail
    currentStage = "创建字典"
    Dim result As Object
    Set result = CreateTextDictionary()

    currentStage = "读取最后一行"
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        Dim seqKey As String
        Dim openDate As Date
        currentStage = "第" & r & "行读取序号"
        seqKey = NormalizeText(ws.Cells(r, CLng(headers(COL_SEQ))).Value)
        If Len(seqKey) = 0 Then GoTo ContinueRow
        currentStage = "第" & r & "行读取开放日"
        If Not TryReadDate(ws.Cells(r, CLng(headers(COL_OPEN_DATE))).Value, openDate) Then GoTo ContinueRow

        Dim productOpenDates As Collection
        currentStage = "第" & r & "行取得日期集合"
        If Not result.Exists(seqKey) Then
            Set productOpenDates = New Collection
            result.Add seqKey, productOpenDates
        Else
            Set productOpenDates = result(seqKey)
        End If
        currentStage = "第" & r & "行去重"
        If Not DateExists(productOpenDates, openDate) Then productOpenDates.Add openDate
ContinueRow:
    Next r

    Set BuildOpenDayIndex = result
    Exit Function

IndexFail:
    Dim errorNumber As Long
    Dim errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, , currentStage & "失败：" & errorDescription
End Function

Private Function AppendOpenDayIfMissing(ByVal ws As Worksheet, ByVal headers As Object, _
                                        ByVal openIndex As Object, ByVal seqKey As String, _
                                        ByVal openDate As Date) As Boolean
    Dim productOpenDates As Collection
    If openIndex.Exists(seqKey) Then
        Set productOpenDates = openIndex(seqKey)
        If DateExists(productOpenDates, openDate) Then Exit Function
    End If

    Dim nextRow As Long
    nextRow = LastUsedRow(ws) + 1
    ws.Cells(nextRow, CLng(headers(COL_SEQ))).Value = seqKey
    ws.Cells(nextRow, CLng(headers(COL_OPEN_DATE))).Value = DateOnly(openDate)
    ws.Cells(nextRow, CLng(headers(COL_OPEN_DATE))).NumberFormat = "yyyy-mm-dd"

    If Not openIndex.Exists(seqKey) Then
        Set productOpenDates = New Collection
        Set openIndex.Item(seqKey) = productOpenDates
    Else
        Set productOpenDates = openIndex(seqKey)
    End If
    productOpenDates.Add DateOnly(openDate)
    AppendOpenDayIfMissing = True
End Function

Private Function DateExists(ByVal dates As Collection, ByVal targetDate As Date) As Boolean
    Dim item As Variant
    For Each item In dates
        If DateOnly(CDate(item)) = DateOnly(targetDate) Then
            DateExists = True
            Exit Function
        End If
    Next item
End Function

Private Function LoadProposalKeys(ByVal ws As Worksheet, ByVal headers As Object) As Object
    Dim result As Object
    Set result = CreateTextDictionary()

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        Dim seqKey As String
        Dim baselineDate As Date
        seqKey = NormalizeText(ws.Cells(r, CLng(headers(COL_SEQ))).Value)
        If Len(seqKey) > 0 And TryReadDate(ws.Cells(r, CLng(headers(PENDING_COL_BASELINE))).Value, baselineDate) Then
            Dim proposalKey As String
            proposalKey = BuildProposalKey(seqKey, baselineDate)
            If Not result.Exists(proposalKey) Then result.Add proposalKey, True
        End If
    Next r

    Set LoadProposalKeys = result
End Function

Private Function BuildProposalKey(ByVal seqKey As String, ByVal baselineDate As Date) As String
    BuildProposalKey = seqKey & "|" & Format$(DateOnly(baselineDate), "yyyymmdd")
End Function

Private Function ProductIsEnabled(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal headers As Object) As Boolean
    If Not headers.Exists(COL_EXPORT_ENABLED) Then
        ProductIsEnabled = True
        Exit Function
    End If

    Dim textValue As String
    textValue = UCase$(NormalizeText(ws.Cells(rowNumber, CLng(headers(COL_EXPORT_ENABLED))).Value))
    If Len(textValue) = 0 Then
        ProductIsEnabled = True
    Else
        ProductIsEnabled = (textValue = "是" Or textValue = "Y" Or textValue = "YES" Or _
                            textValue = "1" Or textValue = "TRUE" Or textValue = "启用")
    End If
End Function

Private Function GetBaselineDateFromNAV() As Date
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAV)

    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)
    RequireHeader headers, COL_NAV_DATE, SHEET_NAV

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim found As Boolean
    Dim maxDate As Date
    Dim r As Long
    For r = 2 To lastRow
        Dim parsedDate As Date
        If TryReadDate(ws.Cells(r, CLng(headers(COL_NAV_DATE))).Value, parsedDate) Then
            If Not found Or parsedDate > maxDate Then
                maxDate = parsedDate
                found = True
            End If
        End If
    Next r

    If Not found Then Err.Raise vbObjectError + 6101, , SHEET_NAV & "工作表中未找到有效日期"
    GetBaselineDateFromNAV = maxDate
End Function

Private Sub RequireHeader(ByVal headers As Object, ByVal headerName As String, ByVal sheetName As String)
    If Not headers.Exists(headerName) Then
        Err.Raise vbObjectError + 6102, , sheetName & "工作表缺少字段：" & headerName
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
    On Error GoTo RowFail
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = foundCell.Row
    End If
    Exit Function
RowFail:
    Err.Raise Err.Number, , "工作表=" & ws.Name & "；" & Err.Description
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
