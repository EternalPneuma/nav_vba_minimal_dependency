' OnePage01_ExportChartData：导出产品一页通图表数据，供后续画图脚本使用

Option Explicit

Private Const SHEET_NAV_DATA As String = "净值数据"
Private Const SHEET_PRODUCT_INFO As String = "产品信息"

Private Const COL_PRODUCT_CODE As String = "产品编号"
Private Const COL_TRUST_CODE As String = "信托计划代码"
Private Const COL_PRODUCT_SHORT As String = "产品简称"
Private Const COL_PRODUCT_NAME As String = "产品名称"
Private Const COL_PRODUCT_FULL_NAME As String = "产品全称"
Private Const COL_TRUST_NAME As String = "信托计划名称"
Private Const COL_INCEPTION_DATE As String = "成立日"
Private Const COL_NAV_DATE As String = "日期"
Private Const COL_NAV_DATE_ALT As String = "净值日期"
Private Const COL_UNIT_NAV As String = "单位净值"
Private Const COL_UNIT_NAV_ALT As String = "净值"
Private Const COL_REITS_CLOSE As String = "REITs收盘"
Private Const COL_REITS_ANNUALIZED_RETURN As String = "REITs期间年化收益率"

Private Const DEFAULT_INCEPTION_NAV As Double = 1#

Public Sub OnePage01_ExportChartData()
    Dim appCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    appCalc = Application.Calculation

    Dim wbOutput As Workbook
    Dim currentStep As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    currentStep = "读取净值数据sheet"
    Dim wsNav As Worksheet
    Set wsNav = FindWorksheetByNames(Array(SHEET_NAV_DATA, "绘图净值数据", "上层产品净值数据(181)"))
    If wsNav Is Nothing Then
        Err.Raise vbObjectError + 5101, , "未找到净值数据sheet：" & SHEET_NAV_DATA
    End If

    currentStep = "读取产品信息sheet"
    Dim wsProduct As Worksheet
    Set wsProduct = ThisWorkbook.Worksheets(SHEET_PRODUCT_INFO)

    Dim targetProducts As Variant
    targetProducts = Array("OA4400", "P83600", "P83800")

    currentStep = "读取产品信息：产品简称和成立日"
    Dim productInfo As Object
    Set productInfo = BuildProductInfoLookup(wsProduct, targetProducts)

    currentStep = "读取净值数据：目标产品历史净值"
    Dim navGroups As Object
    Dim maxNavDate As Date
    Set navGroups = BuildNavGroups(wsNav, targetProducts, maxNavDate)

    If maxNavDate = 0 Then
        Err.Raise vbObjectError + 5102, , "未在净值数据中找到 OA4400 / P83600 / P83800 的有效净值记录。"
    End If

    currentStep = "读取REITs全收益数据"
    Dim reitsLookup As Object
    Set reitsLookup = BuildReitsLookup(maxNavDate)

    currentStep = "创建输出工作簿"
    Set wbOutput = CreateCleanWorkbook()

    Dim usedSheetNames As Object
    Set usedSheetNames = CreateObject("Scripting.Dictionary")
    usedSheetNames.CompareMode = vbTextCompare

    Dim i As Long
    Dim productCode As String
    Dim exportedCount As Long
    Dim missingText As String
    Dim wsOut As Worksheet
    Dim info As Object
    Dim subDict As Object
    Dim sheetName As String

    For i = LBound(targetProducts) To UBound(targetProducts)
        productCode = CStr(targetProducts(i))

        If Not productInfo.Exists(productCode) Then
            missingText = missingText & productCode & "：产品信息缺少信托计划代码/产品编号、产品简称或成立日" & vbCrLf
            GoTo NextProduct
        End If
        If Not navGroups.Exists(productCode) Then
            missingText = missingText & productCode & "：净值数据缺少有效记录" & vbCrLf
            GoTo NextProduct
        End If

        Set info = productInfo(productCode)
        Set subDict = navGroups(productCode)

        sheetName = CleanSheetName(CStr(info("shortName")), usedSheetNames)
        usedSheetNames(sheetName) = 1

        If exportedCount = 0 Then
            Set wsOut = wbOutput.Worksheets(1)
            wsOut.Name = sheetName
        Else
            Set wsOut = wbOutput.Worksheets.Add(After:=wbOutput.Worksheets(wbOutput.Worksheets.Count))
            wsOut.Name = sheetName
        End If

        currentStep = "写入产品sheet：" & productCode
        WriteProductSheet wsOut, productCode, CStr(info("shortName")), CDate(info("inceptionDate")), subDict, reitsLookup
        exportedCount = exportedCount + 1

NextProduct:
    Next i

    If exportedCount = 0 Then
        Err.Raise vbObjectError + 5103, , "三个目标产品均未生成数据。" & vbCrLf & missingText
    End If

    currentStep = "保存输出工作簿"
    Dim outputPath As String
    outputPath = ThisWorkbook.Path & Application.PathSeparator & Format$(maxNavDate, "yyyymmdd") & "-产品一页通.xlsx"
    wbOutput.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    wbOutput.Close SaveChanges:=False
    Set wbOutput = Nothing

    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    Dim finalMsg As String
    finalMsg = "产品一页通数据导出完成" & vbCrLf & vbCrLf & _
               "处理结果：" & vbCrLf & _
               "导出产品数：" & exportedCount & vbCrLf & vbCrLf & _
               "输出文件：" & vbCrLf & outputPath
    If Len(missingText) > 0 Then
        finalMsg = finalMsg & vbCrLf & vbCrLf & "注意事项：" & vbCrLf & _
                   "未导出明细：" & vbCrLf & missingText
    End If
    MsgBox finalMsg, vbInformation, "产品一页通"
    Exit Sub

CleanFail:
    Dim failNumber As Long
    Dim failDescription As String
    Dim failStep As String
    failNumber = Err.Number
    failDescription = Err.Description
    failStep = currentStep

    On Error Resume Next
    If Not wbOutput Is Nothing Then wbOutput.Close SaveChanges:=False
    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If Len(failDescription) = 0 Then failDescription = "未知错误"
    If Len(failStep) = 0 Then failStep = "未记录"

    MsgBox "产品一页通数据导出失败" & vbCrLf & vbCrLf & _
           "错误信息：" & failDescription & vbCrLf & _
           "错误号：" & failNumber & vbCrLf & _
           "步骤：" & failStep, vbCritical, "产品一页通"
End Sub

Private Function BuildProductInfoLookup(ByVal wsProduct As Worksheet, ByVal targetProducts As Variant) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim targetSet As Object
    Set targetSet = BuildTargetSet(targetProducts)

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(wsProduct, 1)

    Dim codeCol As Long
    codeCol = FindFirstExistingHeader(headerMap, Array(COL_TRUST_CODE, COL_PRODUCT_CODE))
    Dim shortCol As Long
    shortCol = FindFirstExistingHeader(headerMap, Array(COL_PRODUCT_SHORT, COL_PRODUCT_NAME, COL_PRODUCT_FULL_NAME, COL_TRUST_NAME))
    Dim inceptionCol As Long
    inceptionCol = FindFirstExistingHeader(headerMap, Array(COL_INCEPTION_DATE))

    If codeCol = 0 Or shortCol = 0 Or inceptionCol = 0 Then
        Err.Raise vbObjectError + 5111, , "产品信息缺少必要字段：信托计划代码/产品编号、产品简称、成立日。"
    End If

    Dim lastRow As Long
    lastRow = LastUsedRow(wsProduct)

    Dim r As Long
    Dim productCode As String
    Dim shortName As String
    Dim inceptionValue As Variant
    Dim info As Object

    For r = 2 To lastRow
        productCode = NormalizeText(wsProduct.Cells(r, codeCol).Value)
        If Len(productCode) = 0 Then GoTo ContinueRow
        If Not targetSet.Exists(productCode) Then GoTo ContinueRow

        shortName = NormalizeText(wsProduct.Cells(r, shortCol).Value)
        inceptionValue = wsProduct.Cells(r, inceptionCol).Value
        If Len(shortName) = 0 Or Not IsDate(inceptionValue) Then GoTo ContinueRow

        Set info = CreateObject("Scripting.Dictionary")
        info.CompareMode = vbTextCompare
        info("shortName") = shortName
        info("inceptionDate") = CDate(inceptionValue)
        If result.Exists(productCode) Then result.Remove productCode
        result.Add productCode, info

ContinueRow:
    Next r

    Set BuildProductInfoLookup = result
End Function

Private Function BuildNavGroups(ByVal wsNav As Worksheet, ByVal targetProducts As Variant, ByRef maxNavDate As Date) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim targetSet As Object
    Set targetSet = BuildTargetSet(targetProducts)

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(wsNav, 1)

    Dim codeCol As Long
    codeCol = FindFirstExistingHeader(headerMap, Array(COL_PRODUCT_CODE, COL_TRUST_CODE))
    Dim dateCol As Long
    dateCol = FindFirstExistingHeader(headerMap, Array(COL_NAV_DATE, COL_NAV_DATE_ALT))
    Dim navCol As Long
    navCol = FindFirstExistingHeader(headerMap, Array(COL_UNIT_NAV, COL_UNIT_NAV_ALT))

    If codeCol = 0 Or dateCol = 0 Or navCol = 0 Then
        Err.Raise vbObjectError + 5121, , "净值数据缺少必要字段：产品编号、日期、单位净值。"
    End If

    Dim lastRow As Long
    lastRow = LastUsedRow(wsNav)

    Dim r As Long
    Dim productCode As String
    Dim navDateValue As Variant
    Dim navValue As Variant
    Dim pk As String
    Dim subDict As Object
    Dim rowArr(1 To 2) As Variant

    For r = 2 To lastRow
        productCode = NormalizeText(wsNav.Cells(r, codeCol).Value)
        If Len(productCode) = 0 Then GoTo ContinueRow
        If Not targetSet.Exists(productCode) Then GoTo ContinueRow

        navDateValue = wsNav.Cells(r, dateCol).Value
        navValue = wsNav.Cells(r, navCol).Value
        Dim parsedNavDate As Date
        If Not TryParseDateValue(navDateValue, parsedNavDate) Then GoTo ContinueRow
        If Not IsNumeric(navValue) Then GoTo ContinueRow

        If parsedNavDate > maxNavDate Then maxNavDate = parsedNavDate
        pk = Format$(parsedNavDate, "yyyy-mm-dd") & "|" & productCode

        If result.Exists(productCode) Then
            Set subDict = result(productCode)
        Else
            Set subDict = CreateObject("Scripting.Dictionary")
            subDict.CompareMode = vbTextCompare
            result.Add productCode, subDict
        End If

        rowArr(1) = parsedNavDate
        rowArr(2) = CDbl(navValue)
        subDict(pk) = rowArr

ContinueRow:
    Next r

    Set BuildNavGroups = result
End Function

Private Function BuildReitsLookup(ByVal maxNavDate As Date) As Object
    Dim reitsPath As String
    reitsPath = FindReitsDataPath(maxNavDate)

    Dim wbReits As Workbook
    Set wbReits = Workbooks.Open(Filename:=reitsPath, ReadOnly:=True)

    On Error GoTo CleanFail

    Dim wsReits As Worksheet
    Set wsReits = wbReits.Worksheets(1)

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(wsReits, 1)

    Dim dateCol As Long
    dateCol = FindFirstExistingHeader(headerMap, Array(COL_NAV_DATE, COL_NAV_DATE_ALT))
    Dim closeCol As Long
    closeCol = FindFirstExistingHeader(headerMap, Array(COL_REITS_CLOSE))
    Dim returnCol As Long
    returnCol = FindFirstExistingHeader(headerMap, Array(COL_REITS_ANNUALIZED_RETURN))

    If dateCol = 0 Or closeCol = 0 Or returnCol = 0 Then
        Err.Raise vbObjectError + 5132, , "REITs全收益数据缺少必要字段：日期、REITs收盘、REITs期间年化收益率。"
    End If

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = LastUsedRow(wsReits)

    Dim r As Long
    Dim parsedDate As Date
    Dim closeValue As Variant
    Dim returnValue As Variant
    Dim rowArr(1 To 2) As Variant
    Dim maxReitsDate As Date

    For r = 2 To lastRow
        If Not TryParseDateValue(wsReits.Cells(r, dateCol).Value, parsedDate) Then GoTo ContinueRow

        closeValue = wsReits.Cells(r, closeCol).Value
        returnValue = wsReits.Cells(r, returnCol).Value
        If Not IsNumeric(closeValue) Or Not IsNumeric(returnValue) Then GoTo ContinueRow

        rowArr(1) = CDbl(closeValue)
        rowArr(2) = CDbl(returnValue)
        result(Format$(parsedDate, "yyyy-mm-dd")) = rowArr
        If parsedDate > maxReitsDate Then maxReitsDate = parsedDate

ContinueRow:
    Next r

    If result.Count = 0 Then
        Err.Raise vbObjectError + 5134, , "REITs全收益数据文件没有可用记录：" & reitsPath
    End If
    If maxReitsDate < maxNavDate Then
        Err.Raise vbObjectError + 5135, , "REITs全收益数据日期早于基准日期。" & vbCrLf & _
                                      "REITs最大日期：" & Format$(maxReitsDate, "yyyy-mm-dd") & vbCrLf & _
                                      "基准日期：" & Format$(maxNavDate, "yyyy-mm-dd") & vbCrLf & _
                                      "文件：" & reitsPath
    End If

    wbReits.Close SaveChanges:=False
    Set BuildReitsLookup = result
    Exit Function

CleanFail:
    Dim failDescription As String
    failDescription = Err.Description
    On Error Resume Next
    wbReits.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 5133, , "读取REITs全收益数据失败：" & failDescription
End Function

Private Function FindReitsDataPath(ByVal baseDate As Date) As String
    Dim folderPath As String
    folderPath = ThisWorkbook.Path

    Dim fileName As String
    fileName = Dir$(folderPath & Application.PathSeparator & "*-中证REITs全收益.xlsx")

    Dim foundAny As Boolean
    Dim foundValidDate As Boolean
    Dim chosenPath As String
    Dim chosenDate As Date
    Dim latestOldDate As Date
    Dim candidateDate As Date
    Dim dateText As String

    Do While Len(fileName) > 0
        foundAny = True
        dateText = Left$(fileName, 8)
        If Len(dateText) = 8 And IsNumeric(dateText) Then
            If IsDate(Left$(dateText, 4) & "-" & Mid$(dateText, 5, 2) & "-" & Right$(dateText, 2)) Then
                foundValidDate = True
                candidateDate = DateSerial(CInt(Left$(dateText, 4)), CInt(Mid$(dateText, 5, 2)), CInt(Right$(dateText, 2)))
                If candidateDate >= baseDate Then
                    If chosenDate = 0 Or candidateDate < chosenDate Then
                        chosenDate = candidateDate
                        chosenPath = folderPath & Application.PathSeparator & fileName
                    End If
                ElseIf candidateDate > latestOldDate Then
                    latestOldDate = candidateDate
                End If
            End If
        End If

        fileName = Dir$()
    Loop

    If Len(chosenPath) > 0 Then
        FindReitsDataPath = chosenPath
        Exit Function
    End If

    If foundAny And foundValidDate Then
        Err.Raise vbObjectError + 5131, , "REITs全收益数据文件日期早于基准日期。" & vbCrLf & _
                                      "最新REITs文件日期：" & Format$(latestOldDate, "yyyy-mm-dd") & vbCrLf & _
                                      "基准日期：" & Format$(baseDate, "yyyy-mm-dd") & vbCrLf & _
                                      "请先生成不早于基准日期的 REITs 全收益文件。"
    ElseIf foundAny Then
        Err.Raise vbObjectError + 5131, , "REITs全收益数据文件命名不符合要求：" & vbCrLf & _
                                      folderPath & Application.PathSeparator & "yyyymmdd-中证REITs全收益.xlsx"
    Else
        Err.Raise vbObjectError + 5131, , "未找到REITs全收益数据文件：" & folderPath & Application.PathSeparator & "yyyymmdd-中证REITs全收益.xlsx"
    End If
End Function

Private Sub WriteProductSheet(ByVal wsOut As Worksheet, ByVal productCode As String, ByVal productShort As String, ByVal inceptionDate As Date, ByVal subDict As Object, ByVal reitsLookup As Object)
    wsOut.Cells(1, 1).Value = "日期"
    wsOut.Cells(1, 2).Value = "产品编号"
    wsOut.Cells(1, 3).Value = "产品简称"
    wsOut.Cells(1, 4).Value = "单位净值"
    wsOut.Cells(1, 5).Value = "成立以来天数"
    wsOut.Cells(1, 6).Value = "成立以来年化收益率"
    If StrComp(productCode, "OA4400", vbTextCompare) = 0 Then
        wsOut.Cells(1, 7).Value = COL_REITS_CLOSE
        wsOut.Cells(1, 8).Value = COL_REITS_ANNUALIZED_RETURN
    End If

    Dim nRows As Long
    nRows = subDict.Count
    If nRows = 0 Then Exit Sub

    Dim tmpArr() As Variant
    ReDim tmpArr(1 To nRows, 1 To 3)

    Dim idx As Long
    Dim pkKey As Variant
    Dim rowArr As Variant

    For Each pkKey In subDict.Keys
        idx = idx + 1
        rowArr = subDict(pkKey)
        tmpArr(idx, 1) = rowArr(1)
        tmpArr(idx, 2) = rowArr(2)
        tmpArr(idx, 3) = Format$(CDate(rowArr(1)), "yyyy-mm-dd")
    Next pkKey

    SortByCol tmpArr, 3
    SmoothDividendDays tmpArr, 2

    Dim writeArr() As Variant
    ReDim writeArr(1 To nRows, 1 To 6)

    Dim i As Long
    Dim elapsedDays As Long
    Dim unitNav As Double
    Dim reitsArr() As Variant
    Dim reitsRow As Variant
    If StrComp(productCode, "OA4400", vbTextCompare) = 0 Then ReDim reitsArr(1 To nRows, 1 To 2)

    For i = 1 To nRows
        unitNav = CDbl(tmpArr(i, 2))
        elapsedDays = DateDiff("d", inceptionDate, CDate(tmpArr(i, 1))) + 1

        writeArr(i, 1) = CDate(tmpArr(i, 1))
        writeArr(i, 2) = productCode
        writeArr(i, 3) = productShort
        writeArr(i, 4) = unitNav
        If elapsedDays > 0 Then
            writeArr(i, 5) = elapsedDays
            writeArr(i, 6) = (unitNav / DEFAULT_INCEPTION_NAV - 1#) / elapsedDays * 365#
        Else
            writeArr(i, 5) = vbNullString
            writeArr(i, 6) = vbNullString
        End If

        If StrComp(productCode, "OA4400", vbTextCompare) = 0 Then
            If Not reitsLookup Is Nothing Then
                If reitsLookup.Exists(CStr(tmpArr(i, 3))) Then
                    reitsRow = reitsLookup(CStr(tmpArr(i, 3)))
                    reitsArr(i, 1) = reitsRow(1)
                    reitsArr(i, 2) = reitsRow(2)
                End If
            End If
        End If
    Next i

    wsOut.Range("A2").Resize(nRows, 6).Value = writeArr
    If StrComp(productCode, "OA4400", vbTextCompare) = 0 Then
        wsOut.Range("G2").Resize(nRows, 2).Value = reitsArr
    End If
    wsOut.Columns("A").NumberFormat = "yyyy-mm-dd"
    wsOut.Columns("D").NumberFormat = "0.0000"
    wsOut.Columns("E").NumberFormat = "0"
    wsOut.Columns("F").NumberFormat = "0.00%"
    If StrComp(productCode, "OA4400", vbTextCompare) = 0 Then
        wsOut.Columns("G").NumberFormat = "0.0000"
        wsOut.Columns("H").NumberFormat = "0.00%"
        wsOut.Columns("A:H").AutoFit
    Else
        wsOut.Columns("A:F").AutoFit
    End If
End Sub

Private Sub SmoothDividendDays(ByRef arr As Variant, ByVal navCol As Long)
    Const K_RATIO As Double = 3#
    Const THRESHOLD As Double = 0.0005

    Dim n As Long
    n = UBound(arr, 1)
    If n < 3 Then Exit Sub

    Dim newNav() As Double
    ReDim newNav(1 To n)
    Dim isValid() As Boolean
    ReDim isValid(1 To n)

    Dim i As Long
    For i = 1 To n
        If IsNumeric(arr(i, navCol)) And Not IsEmpty(arr(i, navCol)) Then
            newNav(i) = CDbl(arr(i, navCol))
            isValid(i) = True
        Else
            isValid(i) = False
        End If
    Next i

    Dim prevV As Double
    Dim curV As Double
    Dim nextV As Double
    Dim expected As Double
    Dim jump As Double
    Dim baseline As Double
    Dim limit As Double

    For i = 2 To n - 1
        If Not isValid(i) Then GoTo NextPoint
        If Not isValid(i - 1) Then GoTo NextPoint
        If Not isValid(i + 1) Then GoTo NextPoint

        prevV = CDbl(arr(i - 1, navCol))
        curV = CDbl(arr(i, navCol))
        nextV = CDbl(arr(i + 1, navCol))

        expected = (prevV + nextV) / 2
        jump = Abs(curV - expected)
        baseline = Abs(nextV - prevV)

        limit = baseline * K_RATIO
        If prevV * THRESHOLD > limit Then limit = prevV * THRESHOLD

        If jump > limit Then
            newNav(i) = expected
        End If
NextPoint:
    Next i

    For i = 1 To n
        If isValid(i) Then
            arr(i, navCol) = newNav(i)
        End If
    Next i
End Sub

Private Function TryParseDateValue(ByVal value As Variant, ByRef parsedDate As Date) As Boolean
    If IsError(value) Or IsEmpty(value) Then Exit Function

    If IsDate(value) Then
        parsedDate = CDate(value)
        TryParseDateValue = True
        Exit Function
    End If

    Dim textValue As String
    textValue = NormalizeText(value)
    If Len(textValue) = 8 And IsNumeric(textValue) Then
        parsedDate = DateSerial(CInt(Left$(textValue, 4)), CInt(Mid$(textValue, 5, 2)), CInt(Right$(textValue, 2)))
        TryParseDateValue = True
    End If
End Function
Private Function CreateCleanWorkbook() As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Set CreateCleanWorkbook = wb
End Function

Private Function FindWorksheetByNames(ByVal candidates As Variant) As Worksheet
    Dim i As Long
    Dim ws As Worksheet

    For i = LBound(candidates) To UBound(candidates)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(candidates(i)))
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set FindWorksheetByNames = ws
            Exit Function
        End If
    Next i
End Function

Private Function BuildTargetSet(ByVal targetProducts As Variant) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim i As Long
    For i = LBound(targetProducts) To UBound(targetProducts)
        result(CStr(targetProducts(i))) = True
    Next i

    Set BuildTargetSet = result
End Function

Private Function CleanSheetName(ByVal rawName As String, ByVal usedDict As Object) As String
    Dim s As String
    s = rawName

    Dim badChars As Variant
    badChars = Array("\", "/", "?", "*", "[", "]", ":")

    Dim i As Long
    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, CStr(badChars(i)), "-")
    Next i

    Do While Left$(s, 1) = "'"
        s = Mid$(s, 2)
    Loop
    Do While Right$(s, 1) = "'"
        s = Left$(s, Len(s) - 1)
    Loop

    s = NormalizeText(s)
    If Len(s) = 0 Then s = "未命名"
    If Len(s) > 31 Then s = Left$(s, 31)

    Dim baseName As String
    baseName = s

    Dim suffix As Long
    suffix = 2

    Do While usedDict.Exists(s)
        Dim suffixText As String
        suffixText = "_" & suffix
        If Len(baseName) + Len(suffixText) > 31 Then
            s = Left$(baseName, 31 - Len(suffixText)) & suffixText
        Else
            s = baseName & suffixText
        End If
        suffix = suffix + 1
    Loop

    CleanSheetName = s
End Function

Private Sub SortByCol(ByRef arr As Variant, ByVal sortCol As Long)
    Dim n As Long
    n = UBound(arr, 1)

    Dim nCols As Long
    nCols = UBound(arr, 2)

    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim tmp As Variant

    For i = 1 To n - 1
        For j = 1 To n - i
            If CStr(arr(j, sortCol)) > CStr(arr(j + 1, sortCol)) Then
                For c = 1 To nCols
                    tmp = arr(j, c)
                    arr(j, c) = arr(j + 1, c)
                    arr(j + 1, c) = tmp
                Next c
            End If
        Next j
    Next i
End Sub

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

Private Function FindFirstExistingHeader(ByVal headerMap As Object, ByVal candidates As Variant) As Long
    If headerMap Is Nothing Then Exit Function

    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If headerMap.Exists(CStr(candidates(i))) Then
            FindFirstExistingHeader = CLng(headerMap(CStr(candidates(i))))
            Exit Function
        End If
    Next i
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
