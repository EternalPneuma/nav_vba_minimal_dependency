' OnePage00_CheckAndImportNavData：运行产品一页通数据导出前，检查并补充绘图净值数据

Option Explicit

Private Const SHEET_SOURCE_NAV As String = "上层产品净值数据(181)"
Private Const SHEET_TARGET_NAV As String = "绘图净值数据"

Private Const COL_SOURCE_PRODUCT_CODE As String = "信托计划代码"
Private Const COL_SOURCE_NAV_DATE As String = "日期"
Private Const COL_SOURCE_UNIT_NAV As String = "单位净值"
Private Const COL_SOURCE_ACCOUNT_NAME As String = "账套名称"

Private Const COL_TARGET_NAV_DATE As String = "净值日期"
Private Const COL_TARGET_PRODUCT_CODE As String = "产品编号"
Private Const COL_TARGET_PRODUCT_NAME As String = "产品名称"
Private Const COL_TARGET_NAV As String = "净值"

Public Sub OnePage00_CheckAndImportNavData()
    Dim appCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    appCalc = Application.Calculation

    Dim currentStep As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    currentStep = "读取源数据sheet"
    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Worksheets(SHEET_SOURCE_NAV)

    currentStep = "读取目标数据sheet"
    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Worksheets(SHEET_TARGET_NAV)

    currentStep = "定位源数据字段"
    Dim sourceHeaders As Object
    Set sourceHeaders = BuildHeaderMap(wsSource, 1)

    Dim sourceCodeCol As Long
    Dim sourceDateCol As Long
    Dim sourceNavCol As Long
    Dim sourceNameCol As Long

    sourceCodeCol = RequireHeader(sourceHeaders, COL_SOURCE_PRODUCT_CODE, SHEET_SOURCE_NAV)
    sourceDateCol = RequireHeader(sourceHeaders, COL_SOURCE_NAV_DATE, SHEET_SOURCE_NAV)
    sourceNavCol = RequireHeader(sourceHeaders, COL_SOURCE_UNIT_NAV, SHEET_SOURCE_NAV)
    sourceNameCol = FindHeader(sourceHeaders, COL_SOURCE_ACCOUNT_NAME)

    currentStep = "定位目标数据字段"
    Dim targetHeaders As Object
    Set targetHeaders = BuildHeaderMap(wsTarget, 1)

    Dim targetDateCol As Long
    Dim targetCodeCol As Long
    Dim targetNameCol As Long
    Dim targetNavCol As Long

    targetDateCol = RequireHeader(targetHeaders, COL_TARGET_NAV_DATE, SHEET_TARGET_NAV)
    targetCodeCol = RequireHeader(targetHeaders, COL_TARGET_PRODUCT_CODE, SHEET_TARGET_NAV)
    targetNameCol = RequireHeader(targetHeaders, COL_TARGET_PRODUCT_NAME, SHEET_TARGET_NAV)
    targetNavCol = RequireHeader(targetHeaders, COL_TARGET_NAV, SHEET_TARGET_NAV)

    currentStep = "统计目标数据已有产品日期"
    Dim targetLatestDate As Object
    Dim targetProductName As Object
    Dim existingKeys As Object
    Set targetLatestDate = CreateObject("Scripting.Dictionary")
    Set targetProductName = CreateObject("Scripting.Dictionary")
    Set existingKeys = CreateObject("Scripting.Dictionary")
    targetLatestDate.CompareMode = vbTextCompare
    targetProductName.CompareMode = vbTextCompare
    existingKeys.CompareMode = vbTextCompare

    BuildTargetIndex wsTarget, targetDateCol, targetCodeCol, targetNameCol, targetLatestDate, targetProductName, existingKeys
    If targetLatestDate.Count = 0 Then
        Err.Raise vbObjectError + 5401, , SHEET_TARGET_NAV & " 中没有可用于对比的产品编号和净值日期。"
    End If

    currentStep = "检查源数据新增记录"
    Dim pendingRows As Object
    Set pendingRows = CreateObject("Scripting.Dictionary")
    pendingRows.CompareMode = vbTextCompare

    Dim sourceLastRow As Long
    sourceLastRow = LastUsedRow(wsSource)

    Dim r As Long
    Dim productCode As String
    Dim sourceName As String
    Dim navDate As Date
    Dim navValue As Variant
    Dim pendingKey As String
    Dim rowArr(1 To 4) As Variant
    Dim invalidDateCount As Long
    Dim invalidNavCount As Long
    Dim zeroNavCount As Long
    Dim skippedNoTargetCount As Long
    Dim skippedExistingCount As Long
    Dim duplicateSourceCount As Long

    For r = 2 To sourceLastRow
        productCode = NormalizeText(wsSource.Cells(r, sourceCodeCol).Value)
        If Len(productCode) = 0 Then GoTo ContinueSourceRow

        If Not targetLatestDate.Exists(productCode) Then
            skippedNoTargetCount = skippedNoTargetCount + 1
            GoTo ContinueSourceRow
        End If

        If Not TryParseDateValue(wsSource.Cells(r, sourceDateCol).Value, navDate) Then
            invalidDateCount = invalidDateCount + 1
            GoTo ContinueSourceRow
        End If

        pendingKey = Format$(navDate, "yyyy-mm-dd") & "|" & productCode
        If existingKeys.Exists(pendingKey) Then
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

        sourceName = vbNullString
        If targetProductName.Exists(productCode) Then sourceName = CStr(targetProductName(productCode))
        If Len(sourceName) = 0 And sourceNameCol > 0 Then sourceName = NormalizeText(wsSource.Cells(r, sourceNameCol).Value)

        rowArr(1) = navDate
        rowArr(2) = productCode
        rowArr(3) = sourceName
        rowArr(4) = CDbl(navValue)

        If pendingRows.Exists(pendingKey) Then duplicateSourceCount = duplicateSourceCount + 1
        pendingRows(pendingKey) = rowArr

ContinueSourceRow:
    Next r

    Dim importedRows As Long
    Dim importedProducts As Long

    If pendingRows.Count > 0 Then
        currentStep = "写入绘图净值数据"
        importedRows = AppendPendingRows(wsTarget, pendingRows, targetDateCol, targetCodeCol, targetNameCol, targetNavCol)
        importedProducts = CountProducts(pendingRows)
    End If

    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    Dim msg As String
    msg = "产品一页通数据预检查完成" & vbCrLf & vbCrLf & _
          "处理结果：" & vbCrLf & _
          "新增导入行数：" & importedRows & vbCrLf & _
          "涉及产品数：" & importedProducts & vbCrLf & _
          "源数据无对应产品编号行数：" & skippedNoTargetCount & vbCrLf & _
          "目标已存在产品日期行数：" & skippedExistingCount & vbCrLf & _
          "源数据重复产品日期行数：" & duplicateSourceCount & vbCrLf & _
          "无效日期行数：" & invalidDateCount & vbCrLf & _
          "无效净值行数：" & invalidNavCount & vbCrLf & _
          "净值为0行数：" & zeroNavCount

    If importedRows = 0 Then
        msg = msg & vbCrLf & vbCrLf & "注意事项：" & vbCrLf & _
              "没有发现需要导入的新净值数据。"
    Else
        msg = msg & vbCrLf & vbCrLf & "注意事项：" & vbCrLf & _
              "请继续运行 OnePage01_ExportChartData。"
    End If

    MsgBox msg, vbInformation, "产品一页通"
    Exit Sub

CleanFail:
    Dim failNumber As Long
    Dim failDescription As String
    Dim failStep As String
    failNumber = Err.Number
    failDescription = Err.Description
    failStep = currentStep

    On Error Resume Next
    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0

    If Len(failDescription) = 0 Then failDescription = "未知错误"
    If Len(failStep) = 0 Then failStep = "未记录"

    MsgBox "产品一页通数据预检查失败" & vbCrLf & vbCrLf & _
           "错误信息：" & failDescription & vbCrLf & _
           "错误号：" & failNumber & vbCrLf & _
           "步骤：" & failStep, vbCritical, "产品一页通"
End Sub

Private Sub BuildTargetIndex(ByVal ws As Worksheet, ByVal dateCol As Long, ByVal codeCol As Long, ByVal nameCol As Long, ByVal latestDate As Object, ByVal productName As Object, ByVal existingKeys As Object)
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    Dim productCode As String
    Dim nameText As String
    Dim parsedDate As Date
    Dim key As String

    For r = 2 To lastRow
        productCode = NormalizeText(ws.Cells(r, codeCol).Value)
        If Len(productCode) = 0 Then GoTo ContinueRow
        If Not TryParseDateValue(ws.Cells(r, dateCol).Value, parsedDate) Then GoTo ContinueRow

        key = Format$(parsedDate, "yyyy-mm-dd") & "|" & productCode
        existingKeys(key) = True

        If Not latestDate.Exists(productCode) Then
            latestDate(productCode) = parsedDate
        ElseIf parsedDate > CDate(latestDate(productCode)) Then
            latestDate(productCode) = parsedDate
        End If

        nameText = NormalizeText(ws.Cells(r, nameCol).Value)
        If Len(nameText) > 0 Then productName(productCode) = nameText

ContinueRow:
    Next r
End Sub

Private Function AppendPendingRows(ByVal ws As Worksheet, ByVal pendingRows As Object, ByVal dateCol As Long, ByVal codeCol As Long, ByVal nameCol As Long, ByVal navCol As Long) As Long
    Dim nRows As Long
    nRows = pendingRows.Count
    If nRows = 0 Then Exit Function

    Dim rowsArr() As Variant
    ReDim rowsArr(1 To nRows, 1 To 5)

    Dim idx As Long
    Dim key As Variant
    Dim rowArr As Variant

    For Each key In pendingRows.Keys
        idx = idx + 1
        rowArr = pendingRows(key)
        rowsArr(idx, 1) = CDate(rowArr(1))
        rowsArr(idx, 2) = CStr(rowArr(2))
        rowsArr(idx, 3) = CStr(rowArr(3))
        rowsArr(idx, 4) = CDbl(rowArr(4))
        rowsArr(idx, 5) = Format$(CDate(rowArr(1)), "yyyy-mm-dd") & "|" & CStr(rowArr(2))
    Next key

    SortPendingRows rowsArr

    Dim firstWriteRow As Long
    firstWriteRow = LastUsedRow(ws) + 1

    Dim writeArr() As Variant
    ReDim writeArr(1 To nRows, 1 To 4)

    Dim i As Long
    For i = 1 To nRows
        writeArr(i, 1) = rowsArr(i, 1)
        writeArr(i, 2) = rowsArr(i, 2)
        writeArr(i, 3) = rowsArr(i, 3)
        writeArr(i, 4) = rowsArr(i, 4)
    Next i

    For i = 1 To nRows
        ws.Cells(firstWriteRow + i - 1, dateCol).Value = writeArr(i, 1)
        ws.Cells(firstWriteRow + i - 1, codeCol).Value = writeArr(i, 2)
        ws.Cells(firstWriteRow + i - 1, nameCol).Value = writeArr(i, 3)
        ws.Cells(firstWriteRow + i - 1, navCol).Value = writeArr(i, 4)
    Next i

    ws.Range(ws.Cells(firstWriteRow, dateCol), ws.Cells(firstWriteRow + nRows - 1, dateCol)).NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(firstWriteRow, navCol), ws.Cells(firstWriteRow + nRows - 1, navCol)).NumberFormat = "0.0000"

    AppendPendingRows = nRows
End Function

Private Function CountProducts(ByVal pendingRows As Object) As Long
    Dim productSet As Object
    Set productSet = CreateObject("Scripting.Dictionary")
    productSet.CompareMode = vbTextCompare

    Dim key As Variant
    Dim rowArr As Variant

    For Each key In pendingRows.Keys
        rowArr = pendingRows(key)
        productSet(CStr(rowArr(2))) = True
    Next key

    CountProducts = productSet.Count
End Function

Private Sub SortPendingRows(ByRef arr As Variant)
    Dim n As Long
    n = UBound(arr, 1)
    If n < 2 Then Exit Sub

    Dim nCols As Long
    nCols = UBound(arr, 2)

    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim tmp As Variant

    For i = 1 To n - 1
        For j = 1 To n - i
            If CStr(arr(j, 5)) > CStr(arr(j + 1, 5)) Then
                For c = 1 To nCols
                    tmp = arr(j, c)
                    arr(j, c) = arr(j + 1, c)
                    arr(j + 1, c) = tmp
                Next c
            End If
        Next j
    Next i
End Sub

Private Function RequireHeader(ByVal headerMap As Object, ByVal headerName As String, ByVal sheetName As String) As Long
    RequireHeader = FindHeader(headerMap, headerName)
    If RequireHeader = 0 Then
        Err.Raise vbObjectError + 5411, , sheetName & " 缺少必要字段：" & headerName
    End If
End Function

Private Function FindHeader(ByVal headerMap As Object, ByVal headerName As String) As Long
    If headerMap Is Nothing Then Exit Function
    If headerMap.Exists(headerName) Then FindHeader = CLng(headerMap(headerName))
End Function

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

Private Function TryParseDateValue(ByVal value As Variant, ByRef parsedDate As Date) As Boolean
    If IsError(value) Or IsEmpty(value) Then Exit Function
    On Error GoTo InvalidDate

    Dim textValue As String
    textValue = NormalizeText(value)

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
