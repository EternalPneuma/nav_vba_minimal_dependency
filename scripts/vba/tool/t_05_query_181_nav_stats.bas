Option Explicit

'==============================================================
' 模块: 查询181净值记录统计
' 功能: 按“信托计划代码”查询“上层产品净值数据(181)”中的记录数、日期范围和遗漏天数。
'       可先选择代码单元格；取消选择后可手工输入代码。
'==============================================================

Private Const SOURCE_SHEET_NAME As String = "上层产品净值数据(181)"
Private Const FIELD_CODE As String = "信托计划代码"
Private Const FIELD_ACCOUNT_NAME As String = "套账名称"
Private Const HEADER_SCAN_ROWS As Long = 30

Public Sub Tool05_Query181NavStats()

    Dim oldCalculation As XlCalculation
    oldCalculation = Application.Calculation

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim wsData As Worksheet
    Set wsData = GetWorksheet(wb, SOURCE_SHEET_NAME)

    If wsData Is Nothing Then
        MsgBox "查询181净值记录无法继续" & vbCrLf & vbCrLf & _
               "错误信息：未找到工作表：" & SOURCE_SHEET_NAME, vbExclamation, "查询181净值记录"
        GoTo CleanUp
    End If

    Dim headerRow As Long, codeCol As Long, accountNameCol As Long, dateCol As Long
    headerRow = FindHeaderRow(wsData, FIELD_CODE, HEADER_SCAN_ROWS)
    If headerRow = 0 Then
        MsgBox "查询181净值记录无法继续" & vbCrLf & vbCrLf & _
               "错误信息：未在前 " & HEADER_SCAN_ROWS & " 行找到字段：" & FIELD_CODE, vbExclamation, "查询181净值记录"
        GoTo CleanUp
    End If

    codeCol = FindColumnByHeader(wsData, headerRow, FIELD_CODE)
    accountNameCol = FindColumnByHeader(wsData, headerRow, FIELD_ACCOUNT_NAME)
    dateCol = FindDateColumn(wsData, headerRow)

    If codeCol = 0 Or dateCol = 0 Then
        MsgBox "查询181净值记录无法继续" & vbCrLf & vbCrLf & _
               "错误信息：未找到必要字段。" & vbCrLf & _
               "必须存在：" & FIELD_CODE & "，以及 日期/净值日期/估值日期 之一。", _
               vbExclamation, "查询181净值记录"
        GoTo CleanUp
    End If

    Dim queryCode As String, selectedAccountName As String
    queryCode = PickOrInputTrustCode(wsData, codeCol, accountNameCol, headerRow, selectedAccountName)
    If Len(queryCode) = 0 Then GoTo CleanUp

    Dim lastRow As Long, lastCol As Long
    lastRow = LastUsedRow(wsData)
    lastCol = LastUsedColumn(wsData, headerRow)

    If lastRow <= headerRow Then
        MsgBox "查询181净值记录无需处理" & vbCrLf & vbCrLf & _
               "处理结果：" & vbCrLf & _
               SOURCE_SHEET_NAME & " 没有可查询的数据。", vbExclamation, "查询181净值记录"
        GoTo CleanUp
    End If

    Dim dataArr As Variant
    dataArr = wsData.Range(wsData.Cells(headerRow + 1, 1), wsData.Cells(lastRow, lastCol)).Value2

    Dim rowOffset As Long
    rowOffset = headerRow

    Dim recordCount As Long, invalidDateCount As Long
    Dim minDate As Date, maxDate As Date, currentDate As Date
    Dim hasDate As Boolean
    Dim uniqueDateDict As Object, accountNameDict As Object
    Set uniqueDateDict = CreateObject("Scripting.Dictionary")
    Set accountNameDict = CreateObject("Scripting.Dictionary")

    Dim i As Long, codeValue As String, dateKey As String, accountName As String
    For i = 1 To UBound(dataArr, 1)
        codeValue = NormalizeText(dataArr(i, codeCol))

        If StrComp(codeValue, queryCode, vbTextCompare) = 0 Then
            recordCount = recordCount + 1

            If accountNameCol > 0 Then
                accountName = NormalizeText(dataArr(i, accountNameCol))
                If Len(accountName) > 0 Then
                    If Not accountNameDict.Exists(accountName) Then accountNameDict.Add accountName, True
                End If
            End If

            If TryParseDate(dataArr(i, dateCol), currentDate) Then
                If Not hasDate Then
                    minDate = currentDate
                    maxDate = currentDate
                    hasDate = True
                Else
                    If currentDate < minDate Then minDate = currentDate
                    If currentDate > maxDate Then maxDate = currentDate
                End If

                dateKey = Format(currentDate, "yyyy-mm-dd")
                If Not uniqueDateDict.Exists(dateKey) Then uniqueDateDict.Add dateKey, True
            Else
                invalidDateCount = invalidDateCount + 1
            End If
        End If
    Next i

    If recordCount = 0 Then
        MsgBox "查询181净值记录无需处理" & vbCrLf & vbCrLf & _
               "处理结果：" & vbCrLf & _
               "未找到信托计划代码：" & queryCode, vbInformation, "查询181净值记录"
        GoTo CleanUp
    End If

    Dim totalDays As Long, missingDays As Long
    Dim missingDates As Collection
    Set missingDates = New Collection

    If hasDate Then
        totalDays = DateDiff("d", minDate, maxDate) + 1
        CollectMissingDates minDate, maxDate, uniqueDateDict, missingDates
        missingDays = missingDates.Count
    End If


    Dim minDateText As String, maxDateText As String
    If hasDate Then
        minDateText = Format(minDate, "yyyy-mm-dd")
        maxDateText = Format(maxDate, "yyyy-mm-dd")
    Else
        minDateText = "-"
        maxDateText = "-"
    End If

    MsgBox "查询181净值记录完成" & vbCrLf & vbCrLf & _
           "处理结果：" & vbCrLf & _
           "信托计划代码：" & queryCode & vbCrLf & _
           "套账名称：" & DisplayAccountNames(selectedAccountName, accountNameDict) & vbCrLf & _
           "记录数：" & recordCount & " 条" & vbCrLf & _
           "最早日期：" & minDateText & vbCrLf & _
           "最晚日期：" & maxDateText & vbCrLf & _
           "共计天数：" & totalDays & vbCrLf & _
           "遗漏天数：" & missingDays & vbCrLf & _
           BuildMissingDateText(missingDates), _
           vbInformation, "查询181净值记录"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "查询181净值记录失败" & vbCrLf & vbCrLf & _
           "错误信息：" & Err.Description, vbCritical, "查询181净值记录"
    Resume CleanUp
End Sub

Private Function PickOrInputTrustCode(ByVal ws As Worksheet, ByVal codeCol As Long, _
                                      ByVal accountNameCol As Long, ByVal headerRow As Long, _
                                      ByRef selectedAccountName As String) As String
    Dim selectedRange As Range

    On Error Resume Next
    Set selectedRange = Application.InputBox( _
        Prompt:="请选择一个“信托计划代码”单元格；如需手工输入，请点击“取消”。", _
        Title:="查询181净值记录", Type:=8)
    On Error GoTo 0

    If Not selectedRange Is Nothing Then
        If selectedRange.Worksheet.Name <> ws.Name Or selectedRange.Column <> codeCol Or selectedRange.Row <= headerRow Then
            MsgBox "查询181净值记录无法继续" & vbCrLf & vbCrLf & _
                   "错误信息：请选择工作表“" & ws.Name & "”中“" & FIELD_CODE & "”列的数据单元格。", _
                   vbExclamation, "查询181净值记录"
            Exit Function
        End If

        PickOrInputTrustCode = NormalizeText(selectedRange.Value2)
        If accountNameCol > 0 Then selectedAccountName = NormalizeText(ws.Cells(selectedRange.Row, accountNameCol).Value2)
        If Len(PickOrInputTrustCode) > 0 Then Exit Function
    End If

    PickOrInputTrustCode = NormalizeText(InputBox("请输入信托计划代码：", "查询181净值记录"))
End Function

Private Sub CollectMissingDates(ByVal minDate As Date, ByVal maxDate As Date, _
                                ByVal uniqueDateDict As Object, ByVal missingDates As Collection)
    Dim d As Date, key As String, dayOffset As Long
    For dayOffset = 0 To DateDiff("d", minDate, maxDate)
        d = DateAdd("d", dayOffset, minDate)
        key = Format(d, "yyyy-mm-dd")
        If Not uniqueDateDict.Exists(key) Then missingDates.Add d
    Next dayOffset
End Sub

Private Function BuildMissingDateText(ByVal missingDates As Collection) As String
    If missingDates Is Nothing Then
        BuildMissingDateText = "遗漏日期: 无"
        Exit Function
    End If

    If missingDates.Count = 0 Then
        BuildMissingDateText = "遗漏日期: 无"
        Exit Function
    End If

    Dim result As String, i As Long, displayCount As Long
    displayCount = WorksheetFunction.Min(missingDates.Count, 20)
    result = "遗漏日期: "

    For i = 1 To displayCount
        If i > 1 Then result = result & "、"
        result = result & Format(missingDates(i), "yyyy-mm-dd")
    Next i

    If missingDates.Count > displayCount Then
        result = result & " 等" & missingDates.Count & "天"
    End If

    BuildMissingDateText = result
End Function
Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal requiredHeader As String, ByVal scanRows As Long) As Long
    Dim r As Long, c As Long, lastCol As Long
    For r = 1 To WorksheetFunction.Min(scanRows, ws.Rows.Count)
        lastCol = LastUsedColumn(ws, r)
        For c = 1 To lastCol
            If StrComp(NormalizeText(ws.Cells(r, c).Value2), requiredHeader, vbTextCompare) = 0 Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function FindColumnByHeader(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = LastUsedColumn(ws, headerRow)
    For c = 1 To lastCol
        If StrComp(NormalizeText(ws.Cells(headerRow, c).Value2), headerName, vbTextCompare) = 0 Then
            FindColumnByHeader = c
            Exit Function
        End If
    Next c
End Function

Private Function FindDateColumn(ByVal ws As Worksheet, ByVal headerRow As Long) As Long
    Dim candidates As Variant
    candidates = Array("日期", "净值日期", "估值日期")

    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        FindDateColumn = FindColumnByHeader(ws, headerRow, CStr(candidates(i)))
        If FindDateColumn > 0 Then Exit Function
    Next i
End Function

Private Function TryParseDate(ByVal value As Variant, ByRef parsedDate As Date) As Boolean
    On Error GoTo InvalidDate
    If IsError(value) Then Exit Function
    If IsEmpty(value) Or Len(Trim(CStr(value))) = 0 Then Exit Function

    Dim dateText As String
    dateText = Trim(CStr(value))

    If Len(dateText) = 8 And Not (dateText Like "*[!0-9]*") Then
        parsedDate = DateSerial(CInt(Left(dateText, 4)), CInt(Mid(dateText, 5, 2)), CInt(Right(dateText, 2)))
        TryParseDate = True
    ElseIf IsNumeric(value) Then
        If CDbl(value) >= 1 And CDbl(value) <= 2958465 Then
            parsedDate = DateValue(CDate(CDbl(value)))
            TryParseDate = True
        End If
    ElseIf IsDate(value) Then
        parsedDate = DateValue(CDate(value))
        TryParseDate = True
    End If
    Exit Function

InvalidDate:
    TryParseDate = False
End Function

Private Function DisplayAccountNames(ByVal selectedAccountName As String, ByVal accountNameDict As Object) As String
    Dim result As String, key As Variant

    If Len(selectedAccountName) > 0 Then
        result = selectedAccountName
    ElseIf Not accountNameDict Is Nothing Then
        For Each key In accountNameDict.Keys
            If Len(result) > 0 Then result = result & "、"
            result = result & CStr(key)
        Next key
    End If

    If Len(result) = 0 Then result = "-"
    DisplayAccountNames = result
End Function

Private Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = lastCell.Row
    End If
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet, ByVal rowNumber As Long) As Long
    LastUsedColumn = ws.Cells(rowNumber, ws.Columns.Count).End(xlToLeft).Column
    If LastUsedColumn = 1 And Len(NormalizeText(ws.Cells(rowNumber, 1).Value2)) = 0 Then LastUsedColumn = 1
End Function

Private Function NormalizeText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    NormalizeText = Trim(CStr(value))
End Function
