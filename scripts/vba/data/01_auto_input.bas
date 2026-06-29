Option Explicit

' 模块用途：从本工作簿同级目录增量导入每日净值 Excel 到本工作簿的“上层产品净值数据(181)”工作表。
' 使用方式：将本文件内容复制到 上层产品净值数据库.xlsm 的标准模块中，运行 Data01_ImportNav181。

Private Const TARGET_SHEET_NAME As String = "上层产品净值数据(181)"
Private Const DATE_HEADER As String = "日期"
Private Const SOURCE_FILE_HEADER As String = "source_file"
Private Const SOURCE_SHEET_HEADER As String = "source_sheet"
Private Const MAX_HEADER_SCAN_ROWS As Long = 30
Private Const SOURCE_FILE_PREFIX As String = "HS-181_多账套净值查询_"
Private Const SOURCE_FILE_EXT As String = ".xlsx"

Public Sub Data01_ImportNav181()
    ImportNAVIncremental181Core
End Sub

' --- Backward-compatible aliases ---
Private Sub ImportNAVIncremental181Core()
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox("运行前请确认：" & vbCrLf & vbCrLf & _
                           "1. 已经完成 181_多账套净值查询。" & vbCrLf & _
                           "2. 查询结果已保存到当前数据库工作簿同级目录。" & vbCrLf & _
                           "3. 文件名保持系统默认格式：" & SOURCE_FILE_PREFIX & "yyyymmdd" & SOURCE_FILE_EXT & vbCrLf & vbCrLf & _
                           "确认后开始导入。", _
                           vbQuestion + vbYesNo + vbDefaultButton2, "净值数据增量导入")
    If confirmResult <> vbYes Then Exit Sub

    Dim appCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldAskToUpdateLinks As Boolean

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldAskToUpdateLinks = Application.AskToUpdateLinks
    appCalc = Application.Calculation

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual

    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)

    Dim sourceDir As String
    sourceDir = ThisWorkbook.Path
    If Len(sourceDir) = 0 Then
        Err.Raise vbObjectError + 1001, , "当前数据库工作簿尚未保存，请先保存后再运行导入。"
    End If
    If Dir(sourceDir, vbDirectory) = vbNullString Then
        Err.Raise vbObjectError + 1002, , "源目录不存在：" & sourceDir
    End If

    Dim targetHeaderMap As Object
    Set targetHeaderMap = BuildHeaderMap(wsTarget, 1)

    If Not targetHeaderMap.Exists(DATE_HEADER) Then
        Err.Raise vbObjectError + 1003, , "目标表缺少字段：" & DATE_HEADER
    End If

    Dim latestDate As Date
    Dim hasLatestDate As Boolean
    hasLatestDate = GetLatestTargetDate(wsTarget, CLng(targetHeaderMap(DATE_HEADER)), latestDate)

    Dim existingKeys As Object
    Set existingKeys = CreateObject("Scripting.Dictionary")
    existingKeys.CompareMode = vbTextCompare
    LoadExistingBusinessKeys wsTarget, targetHeaderMap, existingKeys

    Dim importedRows As Long
    Dim skippedOldRows As Long
    Dim skippedDuplicateRows As Long
    Dim skippedBlankRows As Long
    Dim failedFiles As String
    Dim processedFiles As Long

    ImportNewRowsFromFolder sourceDir, wsTarget, targetHeaderMap, existingKeys, hasLatestDate, latestDate, _
                            importedRows, skippedOldRows, skippedDuplicateRows, skippedBlankRows, processedFiles, failedFiles

    Dim postDuplicateRows As Long
    postDuplicateRows = CountDuplicateBusinessRows(wsTarget, targetHeaderMap)

    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    Dim latestText As String
    If hasLatestDate Then
        latestText = Format$(latestDate, "yyyy-mm-dd")
    Else
        latestText = "目标表暂无有效日期"
    End If

    Dim message As String
    message = "净值数据增量导入完成" & vbCrLf & vbCrLf & _
              "处理范围：" & sourceDir & vbCrLf & _
              "导入前最新日期：" & latestText & vbCrLf & vbCrLf & _
              "处理结果：" & vbCrLf & _
              "符合要求表格数：" & processedFiles & vbCrLf & _
              "处理文件数：" & processedFiles & vbCrLf & _
              "新增行数：" & importedRows & vbCrLf & _
              "跳过旧日期行数：" & skippedOldRows & vbCrLf & _
              "跳过重复行数：" & skippedDuplicateRows & vbCrLf & _
              "跳过空行数：" & skippedBlankRows & vbCrLf & _
              "导入后重复业务行数：" & postDuplicateRows
    If Len(failedFiles) > 0 Then
        message = message & vbCrLf & vbCrLf & "注意事项：" & vbCrLf & _
                  "读取失败文件/工作表：" & vbCrLf & failedFiles
    End If
    MsgBox message, vbInformation, "净值数据增量导入"
    Exit Sub

CleanFail:
    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "净值数据增量导入失败" & vbCrLf & vbCrLf & _
           "错误信息：" & Err.Description, vbCritical, "净值数据增量导入"
End Sub

Private Sub ImportNewRowsFromFolder(ByVal sourceDir As String, ByVal wsTarget As Worksheet, ByVal targetHeaderMap As Object, _
                                    ByVal existingKeys As Object, ByVal hasLatestDate As Boolean, ByVal latestDate As Date, _
                                    ByRef importedRows As Long, ByRef skippedOldRows As Long, ByRef skippedDuplicateRows As Long, _
                                    ByRef skippedBlankRows As Long, ByRef processedFiles As Long, ByRef failedFiles As String)
    Dim fileName As String
    fileName = Dir(sourceDir & Application.PathSeparator & SOURCE_FILE_PREFIX & "*" & SOURCE_FILE_EXT)

    Do While Len(fileName) > 0
        If IsMatchingSourceWorkbookName(fileName) Then
            processedFiles = processedFiles + 1
            ImportOneWorkbook sourceDir & Application.PathSeparator & fileName, fileName, wsTarget, targetHeaderMap, existingKeys, _
                              hasLatestDate, latestDate, importedRows, skippedOldRows, skippedDuplicateRows, skippedBlankRows, failedFiles
        End If
        fileName = Dir()
    Loop
End Sub

Private Sub ImportOneWorkbook(ByVal filePath As String, ByVal fileName As String, ByVal wsTarget As Worksheet, ByVal targetHeaderMap As Object, _
                              ByVal existingKeys As Object, ByVal hasLatestDate As Boolean, ByVal latestDate As Date, _
                              ByRef importedRows As Long, ByRef skippedOldRows As Long, ByRef skippedDuplicateRows As Long, _
                              ByRef skippedBlankRows As Long, ByRef failedFiles As String)
    Dim wbSource As Workbook
    On Error GoTo OpenFail
    Set wbSource = Workbooks.Open(fileName:=filePath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    On Error GoTo 0

    Dim wsSource As Worksheet
    For Each wsSource In wbSource.Worksheets
        On Error GoTo SheetFail
        ImportOneSheet wsSource, fileName, wsTarget, targetHeaderMap, existingKeys, hasLatestDate, latestDate, _
                       importedRows, skippedOldRows, skippedDuplicateRows, skippedBlankRows
        On Error GoTo 0
ContinueSheet:
    Next wsSource

    wbSource.Close SaveChanges:=False
    Exit Sub

SheetFail:
    failedFiles = failedFiles & fileName & " / " & wsSource.Name & "：" & Err.Description & vbCrLf
    Err.Clear
    On Error GoTo 0
    Resume ContinueSheet

OpenFail:
    failedFiles = failedFiles & fileName & "：" & Err.Description & vbCrLf
    Err.Clear
End Sub

Private Sub ImportOneSheet(ByVal wsSource As Worksheet, ByVal fileName As String, ByVal wsTarget As Worksheet, ByVal targetHeaderMap As Object, _
                           ByVal existingKeys As Object, ByVal hasLatestDate As Boolean, ByVal latestDate As Date, _
                           ByRef importedRows As Long, ByRef skippedOldRows As Long, ByRef skippedDuplicateRows As Long, _
                           ByRef skippedBlankRows As Long)
    Dim headerRow As Long
    headerRow = DetectHeaderRow(wsSource)

    Dim sourceHeaderMap As Object
    Set sourceHeaderMap = BuildHeaderMap(wsSource, headerRow)
    If Not sourceHeaderMap.Exists(DATE_HEADER) Then Exit Sub

    Dim lastRow As Long
    lastRow = LastUsedRow(wsSource)
    If lastRow <= headerRow Then Exit Sub

    Dim sourceDateCol As Long
    sourceDateCol = CLng(sourceHeaderMap(DATE_HEADER))

    Dim r As Long
    For r = headerRow + 1 To lastRow
        If IsRowEmpty(wsSource, r) Then
            skippedBlankRows = skippedBlankRows + 1
        Else
            Dim sourceDate As Date
            If TryReadDate(wsSource.Cells(r, sourceDateCol).value, sourceDate) Then
                If hasLatestDate And sourceDate <= latestDate Then
                    skippedOldRows = skippedOldRows + 1
                Else
                    Dim rowKey As String
                    rowKey = BuildSourceBusinessKey(wsSource, r, sourceHeaderMap, targetHeaderMap)
                    If existingKeys.Exists(rowKey) Then
                        skippedDuplicateRows = skippedDuplicateRows + 1
                    Else
                        AppendSourceRow wsSource, r, wsTarget, targetHeaderMap, sourceHeaderMap
                        existingKeys.Add rowKey, True
                        importedRows = importedRows + 1
                    End If
                End If
            Else
                skippedBlankRows = skippedBlankRows + 1
            End If
        End If
    Next r
End Sub

Private Sub AppendSourceRow(ByVal wsSource As Worksheet, ByVal sourceRow As Long, ByVal wsTarget As Worksheet, ByVal targetHeaderMap As Object, _
                            ByVal sourceHeaderMap As Object)
    Dim targetRow As Long
    targetRow = LastUsedRow(wsTarget) + 1

    Dim header As Variant
    For Each header In targetHeaderMap.Keys
        If sourceHeaderMap.Exists(CStr(header)) Then
            wsTarget.Cells(targetRow, CLng(targetHeaderMap(header))).value = wsSource.Cells(sourceRow, CLng(sourceHeaderMap(CStr(header)))).value
        End If
    Next header
End Sub

Private Function DetectHeaderRow(ByVal ws As Worksheet) As Long
    Dim maxRows As Long
    maxRows = WorksheetFunction.Min(MAX_HEADER_SCAN_ROWS, LastUsedRow(ws))
    If maxRows < 1 Then
        DetectHeaderRow = 1
        Exit Function
    End If

    Dim bestRow As Long
    Dim bestScore As Double
    bestRow = 1
    bestScore = -1#

    Dim r As Long
    For r = 1 To maxRows
        Dim nonBlank As Long
        nonBlank = CountNonBlankInRow(ws, r)
        If nonBlank >= 2 Then
            Dim uniqueCount As Long
            uniqueCount = CountUniqueTextInRow(ws, r)

            Dim score As Double
            score = CDbl(nonBlank) + CDbl(uniqueCount) / CDbl(nonBlank)
            If score > bestScore Then
                bestScore = score
                bestRow = r
            End If
        End If
    Next r

    DetectHeaderRow = bestRow
End Function

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

Private Function GetLatestTargetDate(ByVal ws As Worksheet, ByVal dateCol As Long, ByRef latestDate As Date) As Boolean
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        Dim currentDate As Date
        If TryReadDate(ws.Cells(r, dateCol).value, currentDate) Then
            If Not GetLatestTargetDate Or currentDate > latestDate Then
                latestDate = currentDate
                GetLatestTargetDate = True
            End If
        End If
    Next r
End Function

Private Sub LoadExistingBusinessKeys(ByVal ws As Worksheet, ByVal headerMap As Object, ByVal existingKeys As Object)
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        If Not IsRowEmpty(ws, r) Then
            Dim key As String
            key = BuildTargetBusinessKey(ws, r, headerMap)
            If Not existingKeys.Exists(key) Then existingKeys.Add key, True
        End If
    Next r
End Sub

Private Function CountDuplicateBusinessRows(ByVal ws As Worksheet, ByVal headerMap As Object) As Long
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        If Not IsRowEmpty(ws, r) Then
            Dim key As String
            key = BuildTargetBusinessKey(ws, r, headerMap)
            If dict.Exists(key) Then
                CountDuplicateBusinessRows = CountDuplicateBusinessRows + 1
            Else
                dict.Add key, True
            End If
        End If
    Next r
End Function

Private Function BuildTargetBusinessKey(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal headerMap As Object) As String
    Dim parts As Collection
    Set parts = New Collection

    Dim header As Variant
    For Each header In headerMap.Keys
        If IsBusinessHeader(CStr(header)) Then
            parts.Add CStr(header) & "=" & NormalizeCellForKey(ws.Cells(rowIndex, CLng(headerMap(header))).value)
        End If
    Next header

    BuildTargetBusinessKey = JoinCollection(parts, ChrW$(30))
End Function

Private Function BuildSourceBusinessKey(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal sourceHeaderMap As Object, ByVal targetHeaderMap As Object) As String
    Dim parts As Collection
    Set parts = New Collection

    Dim header As Variant
    For Each header In targetHeaderMap.Keys
        If IsBusinessHeader(CStr(header)) Then
            If sourceHeaderMap.Exists(CStr(header)) Then
                parts.Add CStr(header) & "=" & NormalizeCellForKey(ws.Cells(rowIndex, CLng(sourceHeaderMap(CStr(header)))).value)
            Else
                parts.Add CStr(header) & "="
            End If
        End If
    Next header

    BuildSourceBusinessKey = JoinCollection(parts, ChrW$(30))
End Function

Private Function IsBusinessHeader(ByVal headerText As String) As Boolean
    IsBusinessHeader = (StrComp(headerText, SOURCE_FILE_HEADER, vbTextCompare) <> 0 And _
                        StrComp(headerText, SOURCE_SHEET_HEADER, vbTextCompare) <> 0)
End Function

Private Function IsMatchingSourceWorkbookName(ByVal fileName As String) As Boolean
    If Left$(fileName, 2) = "~$" Then Exit Function
    If StrComp(Right$(fileName, Len(SOURCE_FILE_EXT)), SOURCE_FILE_EXT, vbTextCompare) <> 0 Then Exit Function
    If StrComp(Left$(fileName, Len(SOURCE_FILE_PREFIX)), SOURCE_FILE_PREFIX, vbTextCompare) <> 0 Then Exit Function

    Dim dateText As String
    dateText = Mid$(fileName, Len(SOURCE_FILE_PREFIX) + 1, 8)
    If Len(dateText) <> 8 Or Not IsNumeric(dateText) Then Exit Function

    Dim expectedNameLength As Long
    expectedNameLength = Len(SOURCE_FILE_PREFIX) + 8 + Len(SOURCE_FILE_EXT)
    If Len(fileName) <> expectedNameLength Then Exit Function

    IsMatchingSourceWorkbookName = True
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

Private Function NormalizeCellForKey(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Then
        NormalizeCellForKey = vbNullString
    ElseIf IsDate(value) Then
        NormalizeCellForKey = Format$(DateValue(CDate(value)), "yyyy-mm-dd")
    ElseIf IsNumeric(value) Then
        NormalizeCellForKey = Trim$(CStr(CDbl(value)))
    Else
        NormalizeCellForKey = NormalizeText(value)
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

Private Function IsRowEmpty(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    IsRowEmpty = (WorksheetFunction.CountA(ws.Rows(rowIndex)) = 0)
End Function

Private Function CountNonBlankInRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Long
    CountNonBlankInRow = WorksheetFunction.CountA(ws.Rows(rowIndex))
End Function

Private Function CountUniqueTextInRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Long
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim c As Long
    For c = 1 To lastCol
        Dim textValue As String
        textValue = NormalizeText(ws.Cells(rowIndex, c).value)
        If Len(textValue) > 0 Then
            If Not dict.Exists(textValue) Then dict.Add textValue, True
        End If
    Next c

    CountUniqueTextInRow = dict.Count
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

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = foundCell.Column
    End If
End Function

Private Function JoinCollection(ByVal values As Collection, ByVal delimiter As String) As String
    Dim result As String
    Dim i As Long
    For i = 1 To values.Count
        If i > 1 Then result = result & delimiter
        result = result & CStr(values(i))
    Next i
    JoinCollection = result
End Function
