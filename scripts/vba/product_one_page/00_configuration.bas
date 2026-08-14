' OnePage00：一页通独立配置、模板和数据预检

Option Explicit

Private Const SHEET_PRODUCT_INFO As String = "产品信息"
Private Const SHEET_DRAWING_NAV As String = "绘图净值数据"
Private Const SHEET_SOURCE_NAV As String = "上层产品净值数据(181)"

Private Const COL_TRUST_CODE As String = "信托计划代码"
Private Const COL_PRODUCT_CODE As String = "产品编号"
Private Const COL_PRODUCT_SHORT As String = "产品简称"
Private Const COL_PRODUCT_NAME As String = "产品名称"
Private Const COL_INCEPTION_DATE As String = "成立日"
Private Const COL_NAV_DATE As String = "净值日期"
Private Const COL_NAV_DATE_ALT As String = "日期"
Private Const COL_UNIT_NAV As String = "净值"
Private Const COL_UNIT_NAV_ALT As String = "单位净值"
Private Const COL_ASSET_NAV As String = "资产净值"

Private Const ANCHOR_CHART_0 As String = "chart_000"
Private Const ANCHOR_CHART_1 As String = "chart_001"
Private Const ANCHOR_CHART_2 As String = "chart_002"
Private Const ANCHOR_TABLE As String = "table_000"

Private Const MSO_FALSE As Long = 0

Public Sub OnePage00_RequireStaticReady()
    Dim validationErrors As String
    If Not OnePage_ValidateConfiguration(validationErrors) Then
        Err.Raise vbObjectError + 5501, , "一页通静态预检未通过：" & vbCrLf & validationErrors
    End If
End Sub

Public Sub OnePage00_RequireDataReady()
    OnePage_RequireDataReady
End Sub

Public Function OnePage_ValidateConfiguration(ByRef validationErrors As String) As Boolean
    validationErrors = vbNullString

    If Not ReportConfig_HasOnePageConfiguration() Then
        AppendError validationErrors, "报表配置缺少一页通版本配置，请先运行 Report00_MigrateConfiguration。"
        Exit Function
    End If

    Dim definitions As Collection
    Set definitions = ReportConfig_GetOnePageDefinitions()
    If definitions.Count = 0 Then
        AppendError validationErrors, "一页通版本配置没有数据。"
        Exit Function
    End If

    Dim requiredFields As Variant
    requiredFields = Array("版本名称", "是否启用", "输出顺序", "模板文件", "输出文件前缀", _
                           "基准产品代码", "顶层产品1代码", "顶层产品2代码")

    Dim seenNames As Object
    Dim seenOrders As Object
    Dim seenPrefixes As Object
    Set seenNames = CreateTextDictionary()
    Set seenOrders = CreateTextDictionary()
    Set seenPrefixes = CreateTextDictionary()

    Dim definition As Object
    Dim enabledCount As Long
    Dim commonBaselineCode As String
    Dim rowNumber As Long
    For Each definition In definitions
        rowNumber = rowNumber + 1
        If Not HasFields(definition, requiredFields, validationErrors) Then GoTo ContinueDefinition

        Dim versionName As String
        Dim enabledText As String
        Dim templateFile As String
        Dim outputPrefix As String
        Dim baselineCode As String
        Dim topCode1 As String
        Dim topCode2 As String
        versionName = DefinitionText(definition, "版本名称")
        enabledText = DefinitionText(definition, "是否启用")
        templateFile = DefinitionText(definition, "模板文件")
        outputPrefix = DefinitionText(definition, "输出文件前缀")
        baselineCode = DefinitionText(definition, "基准产品代码")
        topCode1 = DefinitionText(definition, "顶层产品1代码")
        topCode2 = DefinitionText(definition, "顶层产品2代码")

        If Len(versionName) = 0 Then
            AppendError validationErrors, "一页通配置第" & rowNumber & "行缺少版本名称。"
        ElseIf seenNames.Exists(versionName) Then
            AppendError validationErrors, "一页通版本名称重复：" & versionName
        Else
            seenNames(versionName) = True
        End If

        If Not IsRecognizedBoolean(enabledText) Then
            AppendError validationErrors, "一页通版本“" & versionName & "”的是否启用值无效：" & enabledText
            GoTo ContinueDefinition
        End If
        If Not IsEnabledText(enabledText) Then GoTo ContinueDefinition
        enabledCount = enabledCount + 1

        If Not IsPositiveLong(definition("输出顺序")) Then
            AppendError validationErrors, "一页通版本“" & versionName & "”的输出顺序必须为正整数。"
        ElseIf seenOrders.Exists(CStr(CLng(definition("输出顺序")))) Then
            AppendError validationErrors, "一页通输出顺序重复：" & CStr(CLng(definition("输出顺序")))
        Else
            seenOrders(CStr(CLng(definition("输出顺序")))) = True
        End If

        If Len(outputPrefix) = 0 Or InStr(outputPrefix, "\") > 0 Or InStr(outputPrefix, "/") > 0 Or InStr(outputPrefix, ":") > 0 Then
            AppendError validationErrors, "一页通版本“" & versionName & "”的输出文件前缀无效。"
        ElseIf seenPrefixes.Exists(outputPrefix) Then
            AppendError validationErrors, "一页通输出文件前缀重复：" & outputPrefix
        Else
            seenPrefixes(outputPrefix) = True
        End If

        If Not IsSafeRelativePath(templateFile) Then
            AppendError validationErrors, "一页通版本“" & versionName & "”的模板必须是工作簿目录下的安全相对路径：" & templateFile
        ElseIf Len(Dir$(ThisWorkbook.Path & Application.PathSeparator & templateFile)) = 0 Then
            AppendError validationErrors, "一页通版本“" & versionName & "”缺少模板：" & templateFile
        End If

        If Len(baselineCode) = 0 Or Len(topCode1) = 0 Or Len(topCode2) = 0 Then
            AppendError validationErrors, "一页通版本“" & versionName & "”的三个产品代码必须填写完整。"
        ElseIf StrComp(baselineCode, topCode1, vbTextCompare) = 0 Or _
               StrComp(baselineCode, topCode2, vbTextCompare) = 0 Or _
               StrComp(topCode1, topCode2, vbTextCompare) = 0 Then
            AppendError validationErrors, "一页通版本“" & versionName & "”的三个产品代码不能重复。"
        End If

        If Len(commonBaselineCode) = 0 Then
            commonBaselineCode = baselineCode
        ElseIf StrComp(commonBaselineCode, baselineCode, vbTextCompare) <> 0 Then
            AppendError validationErrors, "所有启用的一页通版本必须使用同一基准产品。"
        End If

ContinueDefinition:
    Next definition

    If enabledCount = 0 Then AppendError validationErrors, "一页通没有启用版本。"

    If Len(validationErrors) = 0 Then ValidateRequiredProductInfo definitions, validationErrors
    If Len(validationErrors) = 0 Then ValidateTemplates definitions, validationErrors

    OnePage_ValidateConfiguration = (Len(validationErrors) = 0)
End Function

Public Function OnePage_GetEnabledDefinitions() As Collection
    Dim source As Collection
    Set source = ReportConfig_GetOnePageDefinitions()

    Dim result As New Collection
    Dim definition As Object
    For Each definition In source
        If IsEnabledText(DefinitionText(definition, "是否启用")) Then InsertDefinitionByOrder result, definition
    Next definition

    Set OnePage_GetEnabledDefinitions = result
End Function

Public Function OnePage_GetTargetProductCodes() As Variant
    Dim result As Object
    Set result = CreateTextDictionary()

    Dim definitions As Collection
    Set definitions = OnePage_GetEnabledDefinitions()
    Dim definition As Object
    Dim fieldName As Variant
    For Each definition In definitions
        For Each fieldName In Array("基准产品代码", "顶层产品1代码", "顶层产品2代码")
            result(DefinitionText(definition, CStr(fieldName))) = True
        Next fieldName
    Next definition

    OnePage_GetTargetProductCodes = result.Keys
End Function

Public Function OnePage_GetBaselineCode() As String
    Dim definitions As Collection
    Set definitions = OnePage_GetEnabledDefinitions()
    If definitions.Count = 0 Then Err.Raise vbObjectError + 5502, , "一页通没有启用版本。"
    OnePage_GetBaselineCode = DefinitionText(definitions(1), "基准产品代码")
End Function

Public Function OnePage_GetBaselineDate() As Date
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DRAWING_NAV)
    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)

    Dim codeCol As Long
    Dim dateCol As Long
    Dim navCol As Long
    codeCol = FindHeader(headers, Array(COL_PRODUCT_CODE, COL_TRUST_CODE))
    dateCol = FindHeader(headers, Array(COL_NAV_DATE, COL_NAV_DATE_ALT))
    navCol = FindHeader(headers, Array(COL_UNIT_NAV, COL_UNIT_NAV_ALT))
    If codeCol = 0 Or dateCol = 0 Or navCol = 0 Then Err.Raise vbObjectError + 5503, , SHEET_DRAWING_NAV & "缺少产品代码、净值日期或净值字段。"

    Dim baselineCode As String
    baselineCode = OnePage_GetBaselineCode()
    Dim found As Boolean
    Dim r As Long
    Dim parsedDate As Date
    For r = 2 To LastUsedRow(ws)
        If StrComp(NormalizeText(ws.Cells(r, codeCol).Value), baselineCode, vbTextCompare) = 0 Then
            If TryReadDate(ws.Cells(r, dateCol).Value, parsedDate) And IsPositiveNumber(ws.Cells(r, navCol).Value) Then
                If Not found Or parsedDate > OnePage_GetBaselineDate Then
                    OnePage_GetBaselineDate = parsedDate
                    found = True
                End If
            End If
        End If
    Next r
    If Not found Then Err.Raise vbObjectError + 5504, , "基准产品没有有效绘图净值：" & baselineCode
End Function

Public Sub OnePage_RequireDataReady()
    Dim baseDate As Date
    baseDate = OnePage_GetBaselineDate()
    Dim targetCodes As Variant
    targetCodes = OnePage_GetTargetProductCodes()

    Dim wsDrawing As Worksheet
    Set wsDrawing = ThisWorkbook.Worksheets(SHEET_DRAWING_NAV)
    RequireCodesOnDate wsDrawing, targetCodes, baseDate, Array(COL_PRODUCT_CODE, COL_TRUST_CODE), _
                       Array(COL_NAV_DATE, COL_NAV_DATE_ALT), Array(COL_UNIT_NAV, COL_UNIT_NAV_ALT), "绘图净值"

    Dim topCodes As Object
    Set topCodes = CreateTextDictionary()
    Dim definitions As Collection
    Set definitions = OnePage_GetEnabledDefinitions()
    Dim definition As Object
    For Each definition In definitions
        topCodes(DefinitionText(definition, "顶层产品1代码")) = True
        topCodes(DefinitionText(definition, "顶层产品2代码")) = True
    Next definition

    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Worksheets(SHEET_SOURCE_NAV)
    RequireCodesOnDate wsSource, topCodes.Keys, baseDate, Array(COL_TRUST_CODE, COL_PRODUCT_CODE), _
                       Array(COL_NAV_DATE_ALT, COL_NAV_DATE), Array(COL_ASSET_NAV), "资产净值"
End Sub

Public Function OnePage_DefinitionText(ByVal definition As Object, ByVal fieldName As String) As String
    OnePage_DefinitionText = DefinitionText(definition, fieldName)
End Function

Private Sub ValidateRequiredProductInfo(ByVal definitions As Collection, ByRef validationErrors As String)
    Dim requiredCodes As Object
    Set requiredCodes = CreateTextDictionary()
    Dim definition As Object
    Dim fieldName As Variant
    For Each definition In definitions
        If IsEnabledText(DefinitionText(definition, "是否启用")) Then
            For Each fieldName In Array("基准产品代码", "顶层产品1代码", "顶层产品2代码")
                requiredCodes(DefinitionText(definition, CStr(fieldName))) = True
            Next fieldName
        End If
    Next definition

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PRODUCT_INFO)
    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)
    Dim codeCol As Long
    Dim shortCol As Long
    Dim inceptionCol As Long
    codeCol = FindHeader(headers, Array(COL_TRUST_CODE, COL_PRODUCT_CODE))
    shortCol = FindHeader(headers, Array(COL_PRODUCT_SHORT, COL_PRODUCT_NAME))
    inceptionCol = FindHeader(headers, Array(COL_INCEPTION_DATE))
    If codeCol = 0 Or shortCol = 0 Or inceptionCol = 0 Then
        AppendError validationErrors, SHEET_PRODUCT_INFO & "缺少信托计划代码、产品简称或成立日字段。"
        Exit Sub
    End If

    Dim found As Object
    Set found = CreateTextDictionary()
    Dim r As Long
    Dim code As String
    For r = 2 To LastUsedRow(ws)
        code = NormalizeText(ws.Cells(r, codeCol).Value)
        If requiredCodes.Exists(code) Then
            If Len(NormalizeText(ws.Cells(r, shortCol).Value)) > 0 And IsDate(ws.Cells(r, inceptionCol).Value) Then found(code) = True
        End If
    Next r

    Dim key As Variant
    For Each key In requiredCodes.Keys
        If Not found.Exists(CStr(key)) Then AppendError validationErrors, "产品信息缺少代码、简称或成立日：" & CStr(key)
    Next key
End Sub

Private Sub ValidateTemplates(ByVal definitions As Collection, ByRef validationErrors As String)
    Dim pptApp As Object
    Dim pptWasRunning As Boolean
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application")
    Else
        pptWasRunning = True
    End If
    On Error GoTo TemplateFail
    If pptApp Is Nothing Then Err.Raise vbObjectError + 5510, , "无法启动 PowerPoint。"

    Dim definition As Object
    For Each definition In definitions
        If IsEnabledText(DefinitionText(definition, "是否启用")) Then
            Dim templatePath As String
            templatePath = ThisWorkbook.Path & Application.PathSeparator & DefinitionText(definition, "模板文件")
            Dim pres As Object
            Set pres = pptApp.Presentations.Open(templatePath, MSO_FALSE, MSO_FALSE, MSO_FALSE)
            ValidateTemplateAnchor pres, ANCHOR_CHART_0, False, DefinitionText(definition, "版本名称")
            ValidateTemplateAnchor pres, ANCHOR_CHART_1, False, DefinitionText(definition, "版本名称")
            ValidateTemplateAnchor pres, ANCHOR_CHART_2, False, DefinitionText(definition, "版本名称")
            ValidateTemplateAnchor pres, ANCHOR_TABLE, True, DefinitionText(definition, "版本名称")
            pres.Close
            Set pres = Nothing
        End If
    Next definition

    If Not pptWasRunning Then pptApp.Quit
    Set pptApp = Nothing
    Exit Sub

TemplateFail:
    Dim description As String
    description = Err.Description
    On Error Resume Next
    If Not pres Is Nothing Then pres.Close
    If Not pptApp Is Nothing Then If Not pptWasRunning Then pptApp.Quit
    On Error GoTo 0
    AppendError validationErrors, "一页通模板预检失败：" & description
End Sub

Private Sub ValidateTemplateAnchor(ByVal pres As Object, ByVal anchorName As String, ByVal requireTable As Boolean, ByVal versionName As String)
    Dim foundShape As Object
    Dim foundCount As Long
    Dim sld As Object
    For Each sld In pres.Slides
        CountNamedShapes sld.Shapes, anchorName, foundCount, foundShape
    Next sld
    If foundCount <> 1 Then Err.Raise vbObjectError + 5511, , "版本“" & versionName & "”模板中的" & anchorName & "数量必须为1，实际为" & foundCount & "。"
    If requireTable Then
        On Error Resume Next
        Dim hasTable As Boolean
        hasTable = (foundShape.HasTable <> 0)
        On Error GoTo 0
        If Not hasTable Then Err.Raise vbObjectError + 5512, , "版本“" & versionName & "”模板中的" & anchorName & "不是表格。"
    End If
End Sub

Private Sub CountNamedShapes(ByVal shapesCollection As Object, ByVal targetName As String, ByRef foundCount As Long, ByRef foundShape As Object)
    Dim shp As Object
    For Each shp In shapesCollection
        If StrComp(shp.Name, targetName, vbTextCompare) = 0 Then
            foundCount = foundCount + 1
            Set foundShape = shp
        End If

        Dim groupItems As Object
        Set groupItems = Nothing
        On Error Resume Next
        Set groupItems = shp.GroupItems
        On Error GoTo 0
        If Not groupItems Is Nothing Then CountNamedShapes groupItems, targetName, foundCount, foundShape
    Next shp
End Sub

Private Sub RequireCodesOnDate(ByVal ws As Worksheet, ByVal codes As Variant, ByVal baseDate As Date, _
                               ByVal codeHeaders As Variant, ByVal dateHeaders As Variant, ByVal valueHeaders As Variant, _
                               ByVal valueDescription As String)
    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)
    Dim codeCol As Long
    Dim dateCol As Long
    Dim valueCol As Long
    codeCol = FindHeader(headers, codeHeaders)
    dateCol = FindHeader(headers, dateHeaders)
    valueCol = FindHeader(headers, valueHeaders)
    If codeCol = 0 Or dateCol = 0 Or valueCol = 0 Then Err.Raise vbObjectError + 5520, , ws.Name & "缺少" & valueDescription & "预检所需字段。"

    Dim required As Object
    Set required = CreateTextDictionary()
    Dim codeItem As Variant
    For Each codeItem In codes
        required(CStr(codeItem)) = False
    Next codeItem

    Dim r As Long
    Dim parsedDate As Date
    Dim code As String
    For r = 2 To LastUsedRow(ws)
        code = NormalizeText(ws.Cells(r, codeCol).Value)
        If required.Exists(code) Then
            If TryReadDate(ws.Cells(r, dateCol).Value, parsedDate) Then
                If parsedDate = baseDate And IsPositiveNumber(ws.Cells(r, valueCol).Value) Then required(code) = True
            End If
        End If
    Next r

    Dim missing As String
    For Each codeItem In required.Keys
        If Not CBool(required(codeItem)) Then missing = missing & CStr(codeItem) & "、"
    Next codeItem
    If Len(missing) > 0 Then
        missing = Left$(missing, Len(missing) - 1)
        Err.Raise vbObjectError + 5521, , Format$(baseDate, "yyyy-mm-dd") & "缺少有效" & valueDescription & "：" & missing
    End If
End Sub

Private Sub InsertDefinitionByOrder(ByVal target As Collection, ByVal definition As Object)
    Dim orderValue As Long
    orderValue = CLng(definition("输出顺序"))
    Dim i As Long
    For i = 1 To target.Count
        If orderValue < CLng(target(i)("输出顺序")) Then
            target.Add definition, Before:=i
            Exit Sub
        End If
    Next i
    target.Add definition
End Sub

Private Function HasFields(ByVal definition As Object, ByVal fields As Variant, ByRef validationErrors As String) As Boolean
    HasFields = True
    Dim fieldName As Variant
    For Each fieldName In fields
        If Not definition.Exists(CStr(fieldName)) Then
            AppendError validationErrors, "一页通版本配置缺少字段：" & CStr(fieldName)
            HasFields = False
        End If
    Next fieldName
End Function

Private Function DefinitionText(ByVal definition As Object, ByVal fieldName As String) As String
    If definition Is Nothing Then Exit Function
    If Not definition.Exists(fieldName) Then Exit Function
    DefinitionText = NormalizeText(definition(fieldName))
End Function

Private Function IsSafeRelativePath(ByVal pathText As String) As Boolean
    If Len(pathText) = 0 Then Exit Function
    If Left$(pathText, 1) = "\" Or Left$(pathText, 1) = "/" Then Exit Function
    If InStr(pathText, ":") > 0 Or InStr(pathText, "..") > 0 Then Exit Function
    IsSafeRelativePath = True
End Function

Private Function IsEnabledText(ByVal value As Variant) As Boolean
    Dim textValue As String
    textValue = UCase$(NormalizeText(value))
    IsEnabledText = (textValue = "是" Or textValue = "Y" Or textValue = "YES" Or textValue = "1" Or textValue = "TRUE" Or textValue = "启用")
End Function

Private Function IsRecognizedBoolean(ByVal value As Variant) As Boolean
    Dim textValue As String
    textValue = UCase$(NormalizeText(value))
    IsRecognizedBoolean = IsEnabledText(textValue) Or textValue = "否" Or textValue = "N" Or textValue = "NO" Or textValue = "0" Or textValue = "FALSE" Or textValue = "停用"
End Function

Private Function IsPositiveLong(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    If CDbl(value) <> Fix(CDbl(value)) Then Exit Function
    IsPositiveLong = (CLng(value) > 0)
End Function

Private Function IsPositiveNumber(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    IsPositiveNumber = (CDbl(value) > 0)
End Function

Private Function BuildHeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim result As Object
    Set result = CreateTextDictionary()
    Dim c As Long
    For c = 1 To LastUsedColumn(ws)
        Dim headerText As String
        headerText = NormalizeText(ws.Cells(headerRow, c).Value)
        If Len(headerText) > 0 Then If Not result.Exists(headerText) Then result(headerText) = c
    Next c
    Set BuildHeaderMap = result
End Function

Private Function FindHeader(ByVal headers As Object, ByVal candidates As Variant) As Long
    Dim candidate As Variant
    For Each candidate In candidates
        If headers.Exists(CStr(candidate)) Then
            FindHeader = CLng(headers(CStr(candidate)))
            Exit Function
        End If
    Next candidate
End Function

Private Function CreateTextDictionary() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set CreateTextDictionary = result
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then LastUsedRow = 1 Else LastUsedRow = foundCell.Row
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then LastUsedColumn = 1 Else LastUsedColumn = foundCell.Column
End Function

Private Function NormalizeText(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    Dim textValue As String
    textValue = CStr(value)
    textValue = Replace(textValue, ChrW$(12288), " ")
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    NormalizeText = WorksheetFunction.Trim(textValue)
End Function

Private Function TryReadDate(ByVal value As Variant, ByRef parsedDate As Date) As Boolean
    On Error GoTo InvalidDate
    If IsDate(value) Then
        parsedDate = DateSerial(Year(CDate(value)), Month(CDate(value)), Day(CDate(value)))
        TryReadDate = True
        Exit Function
    End If
    Dim textValue As String
    textValue = NormalizeText(value)
    If Len(textValue) = 8 And IsNumeric(textValue) Then
        parsedDate = DateSerial(CInt(Left$(textValue, 4)), CInt(Mid$(textValue, 5, 2)), CInt(Right$(textValue, 2)))
        TryReadDate = (Format$(parsedDate, "yyyymmdd") = textValue)
    End If
    Exit Function
InvalidDate:
    TryReadDate = False
End Function

Private Sub AppendError(ByRef errors As String, ByVal message As String)
    errors = errors & "· " & message & vbCrLf
End Sub
