Option Explicit

' 工具用途：按"产品分类"中的理论间隔，补充"开放日"工作表中缺失的下一开放日。
' 使用方式：将本文件内容复制到 上层产品净值数据库.xlsm 的标准模块中，运行 ToolFillNextOpenDateByInterval。
'
' 初稿逻辑：
'   1. 基准日期沿用 02_calculate_open_date：从"上层产品净值数据(181)"的"日期"列取最大日期。
'   2. 从"开放日"工作表中按"序号"查找 >= 基准日期的最小开放日。
'   3. 如果没有找到下一开放日，再按"产品分类"中的"理论间隔"推算，并追加写回"开放日"工作表。
'   4. 理论间隔=1 直接跳过；7=下一个周三；183=本月/下月最后一个周三；其他正数按天数推算。
' 注意：本工具不写入、不新增、不清理"产品分类"工作表中的任何输出列；产品分类结果统一由 02_calculate_open_date 负责刷新。

Private Const TARGET_SHEET_NAME As String = "产品分类"
Private Const SOURCE_SHEET_NAME As String = "开放日"
Private Const NAV_SHEET_NAME As String = "上层产品净值数据(181)"

Private Const COL_SEQ As String = "序号"
Private Const COL_OPEN_DATE As String = "开放日"
Private Const COL_THEORETICAL_INTERVAL As String = "理论间隔"
Private Const COL_NAV_DATE As String = "日期"

Public Sub Tool03_FillNextOpenDate()
    FillNextOpenDateByIntervalCore
End Sub


Private Sub FillNextOpenDateByIntervalCore()
    Dim appCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    appCalc = Application.Calculation

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)

    Dim targetHeader As Object
    Set targetHeader = BuildHeaderMap(wsTarget, 1)

    RequireHeader targetHeader, COL_SEQ, TARGET_SHEET_NAME
    RequireHeader targetHeader, COL_THEORETICAL_INTERVAL, TARGET_SHEET_NAME

    Dim seqCol As Long
    Dim intervalCol As Long
    seqCol = CLng(targetHeader(COL_SEQ))
    intervalCol = CLng(targetHeader(COL_THEORETICAL_INTERVAL))

    Dim baselineDate As Date
    baselineDate = GetBaselineDateFromNAV()

    Dim openDayIndex As Object
    Set openDayIndex = BuildOpenDayIndex()

    Dim wsOpenDay As Worksheet
    Set wsOpenDay = ThisWorkbook.Worksheets(SOURCE_SHEET_NAME)

    Dim openDayHeader As Object
    Set openDayHeader = BuildHeaderMap(wsOpenDay, 1)
    RequireHeader openDayHeader, COL_SEQ, SOURCE_SHEET_NAME
    RequireHeader openDayHeader, COL_OPEN_DATE, SOURCE_SHEET_NAME

    Dim openDaySeqCol As Long
    Dim openDayDateCol As Long
    openDaySeqCol = CLng(openDayHeader(COL_SEQ))
    openDayDateCol = CLng(openDayHeader(COL_OPEN_DATE))

    Dim lastRow As Long
    lastRow = LastUsedRow(wsTarget)

    Dim r As Long
    Dim existingOpenDayCount As Long
    Dim inferredCount As Long
    Dim appendedOpenDayCount As Long
    Dim skippedCount As Long

    For r = 2 To lastRow
        Dim seqKey As String
        seqKey = NormalizeText(wsTarget.Cells(r, seqCol).value)
        If Len(seqKey) = 0 Then GoTo ContinueRow

        Dim intervalText As String
        intervalText = NormalizeText(wsTarget.Cells(r, intervalCol).value)
        If intervalText = "1" Then
            skippedCount = skippedCount + 1
            GoTo ContinueRow
        End If

        Dim nextDate As Variant
        Dim currentOpenDate As Variant
        nextDate = Empty
        currentOpenDate = Empty

        If openDayIndex.Exists(seqKey) Then
            nextDate = FindNextOpenDate(openDayIndex(seqKey), baselineDate)
            currentOpenDate = FindPrevOpenDate(openDayIndex(seqKey), baselineDate)
        End If

        If Not IsEmpty(nextDate) Then
            existingOpenDayCount = existingOpenDayCount + 1
        Else
            Dim inferredDate As Variant
            inferredDate = InferNextOpenDateByInterval(baselineDate, wsTarget.Cells(r, intervalCol).value, currentOpenDate)

            If Not IsEmpty(inferredDate) Then
                inferredCount = inferredCount + 1

                If AppendOpenDayIfMissing(wsOpenDay, openDayIndex, openDaySeqCol, openDayDateCol, seqKey, CDate(inferredDate)) Then
                    appendedOpenDayCount = appendedOpenDayCount + 1
                End If
            Else
                skippedCount = skippedCount + 1
            End If
        End If

ContinueRow:
    Next r

    Application.Calculation = appCalc
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    MsgBox "下一开放日补算完成" & vbCrLf & vbCrLf & _
           "基准日期：" & Format$(baselineDate, "yyyy-mm-dd") & vbCrLf & vbCrLf & _
           "处理结果：" & vbCrLf & _
           "开放日表已有下一开放日：" & existingOpenDayCount & vbCrLf & _
           "按理论间隔推算：" & inferredCount & vbCrLf & _
           "追加开放日记录：" & appendedOpenDayCount & vbCrLf & _
           "产品分类写入：0" & vbCrLf & _
           "跳过：" & skippedCount, vbInformation, "下一开放日补算"
    Exit Sub

CleanFail:
    Application.Calculation = appCalc
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "下一开放日补算失败" & vbCrLf & vbCrLf & _
           "错误信息：" & Err.Description, vbCritical, "下一开放日补算"
End Sub

Private Function BuildOpenDayIndex() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SOURCE_SHEET_NAME)

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(ws, 1)
    RequireHeader headerMap, COL_SEQ, SOURCE_SHEET_NAME
    RequireHeader headerMap, COL_OPEN_DATE, SOURCE_SHEET_NAME

    Dim seqCol As Long
    Dim dateCol As Long
    seqCol = CLng(headerMap(COL_SEQ))
    dateCol = CLng(headerMap(COL_OPEN_DATE))

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        Dim seqKey As String
        seqKey = NormalizeText(ws.Cells(r, seqCol).value)
        If Len(seqKey) = 0 Then GoTo ContinueIdx

        Dim openDate As Date
        If Not TryReadDate(ws.Cells(r, dateCol).value, openDate) Then GoTo ContinueIdx

        If Not dict.Exists(seqKey) Then
            dict.Add seqKey, New Collection
        End If
        dict(seqKey).Add openDate

ContinueIdx:
    Next r

    Set BuildOpenDayIndex = dict
End Function

Private Function FindNextOpenDate(ByVal openDates As Collection, ByVal baselineDate As Date) As Variant
    Dim d As Variant
    Dim nextDate As Variant
    nextDate = Empty

    For Each d In openDates
        Dim dt As Date
        dt = CDate(d)
        If dt >= baselineDate Then
            If IsEmpty(nextDate) Or dt < CDate(nextDate) Then
                nextDate = dt
            End If
        End If
    Next d

    FindNextOpenDate = nextDate
End Function

Private Function FindPrevOpenDate(ByVal openDates As Collection, ByVal baselineDate As Date) As Variant
    Dim d As Variant
    Dim prevDate As Variant
    prevDate = Empty

    For Each d In openDates
        Dim dt As Date
        dt = CDate(d)
        If dt < baselineDate Then
            If IsEmpty(prevDate) Or dt > CDate(prevDate) Then
                prevDate = dt
            End If
        End If
    Next d

    FindPrevOpenDate = prevDate
End Function

Private Function InferNextOpenDateByInterval(ByVal baselineDate As Date, ByVal intervalValue As Variant, ByVal currentOpenDate As Variant) As Variant
    Dim intervalText As String
    intervalText = NormalizeText(intervalValue)
    If Len(intervalText) = 0 Or Not IsNumeric(intervalText) Then
        InferNextOpenDateByInterval = Empty
        Exit Function
    End If

    Dim intervalDays As Long
    intervalDays = CLng(intervalText)

    If intervalDays <= 0 Then
        InferNextOpenDateByInterval = Empty
    ElseIf intervalDays = 1 Then
        InferNextOpenDateByInterval = Empty
    ElseIf intervalDays = 7 Then
        InferNextOpenDateByInterval = GetNextWednesday(baselineDate)
    ElseIf intervalDays = 183 Then
        InferNextOpenDateByInterval = GetNextMonthEndWednesday(baselineDate)
    Else
        InferNextOpenDateByInterval = DateAdd("d", intervalDays, baselineDate)
    End If
End Function

Private Function GetNextWednesday(ByVal baselineDate As Date) As Date
    Dim daysToAdd As Long
    daysToAdd = 3 - Weekday(baselineDate, vbMonday)
    If daysToAdd <= 0 Then daysToAdd = daysToAdd + 7
    GetNextWednesday = DateAdd("d", daysToAdd, dateValue(baselineDate))
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

    If currentMonthWednesday >= dateValue(baselineDate) Then
        GetNextMonthEndWednesday = currentMonthWednesday
    Else
        GetNextMonthEndWednesday = GetLastWednesdayOfMonth(DateAdd("m", 1, baselineDate))
    End If
End Function

Private Function AppendOpenDayIfMissing(ByVal ws As Worksheet, ByVal openDayIndex As Object, _
                                        ByVal seqCol As Long, ByVal dateCol As Long, _
                                        ByVal seqKey As String, ByVal openDate As Date) As Boolean
    If OpenDayExists(openDayIndex, seqKey, openDate) Then
        AppendOpenDayIfMissing = False
        Exit Function
    End If

    Dim nextRow As Long
    nextRow = LastUsedRow(ws) + 1

    ws.Cells(nextRow, seqCol).value = seqKey
    ws.Cells(nextRow, dateCol).value = openDate
    ws.Cells(nextRow, dateCol).NumberFormat = "yyyy-mm-dd"

    If Not openDayIndex.Exists(seqKey) Then
        openDayIndex.Add seqKey, New Collection
    End If
    openDayIndex(seqKey).Add openDate

    AppendOpenDayIfMissing = True
End Function

Private Function OpenDayExists(ByVal openDayIndex As Object, ByVal seqKey As String, ByVal openDate As Date) As Boolean
    If Not openDayIndex.Exists(seqKey) Then Exit Function

    Dim d As Variant
    For Each d In openDayIndex(seqKey)
        If dateValue(CDate(d)) = dateValue(openDate) Then
            OpenDayExists = True
            Exit Function
        End If
    Next d
End Function

Private Function GetBaselineDateFromNAV() As Date
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(NAV_SHEET_NAME)

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(ws, 1)
    RequireHeader headerMap, COL_NAV_DATE, NAV_SHEET_NAME

    Dim dateCol As Long
    dateCol = CLng(headerMap(COL_NAV_DATE))

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim maxDate As Date
    Dim found As Boolean
    Dim r As Long

    For r = 2 To lastRow
        Dim parsedDate As Date
        If TryReadDate(ws.Cells(r, dateCol).value, parsedDate) Then
            If Not found Or parsedDate > maxDate Then
                maxDate = parsedDate
                found = True
            End If
        End If
    Next r

    If Not found Then
        Err.Raise vbObjectError + 5101, , NAV_SHEET_NAME & "工作表中未找到有效日期"
    End If

    GetBaselineDateFromNAV = maxDate
End Function

Private Sub RequireHeader(ByVal headerMap As Object, ByVal columnName As String, ByVal sheetName As String)
    If Not headerMap.Exists(columnName) Then
        Err.Raise vbObjectError + 5102, , sheetName & "工作表缺少字段：" & columnName
    End If
End Sub

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

Private Function TryReadDate(ByVal value As Variant, ByRef outDate As Date) As Boolean
    If IsDate(value) Then
        outDate = dateValue(CDate(value))
        TryReadDate = True
        Exit Function
    End If

    Dim textValue As String
    textValue = NormalizeText(value)
    If Len(textValue) = 0 Then Exit Function

    If InStr(textValue, ".") > 0 Then
        textValue = Left$(textValue, InStr(textValue, ".") - 1)
    End If

    If Len(textValue) = 8 And IsNumeric(textValue) Then
        outDate = DateSerial(CInt(Left$(textValue, 4)), CInt(Mid$(textValue, 5, 2)), CInt(Right$(textValue, 2)))
        TryReadDate = True
        Exit Function
    End If

    textValue = Replace(Replace(textValue, ".", "-"), "/", "-")
    If IsDate(textValue) Then
        outDate = dateValue(CDate(textValue))
        TryReadDate = True
    End If
End Function


