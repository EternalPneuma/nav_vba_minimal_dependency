' Report00：报表配置迁移与静态预检

Option Explicit

Private Const SHEET_PRODUCT_CATEGORY As String = "产品分类"
Private Const SHEET_PRODUCT_INFO As String = "产品信息"
Private Const SHEET_REPORT_CONFIG As String = "报表配置"
Private Const SHEET_CONFIG_GUIDE As String = "配置说明"

Private Const TABLE_CATEGORIES As String = "tblReportCategories"
Private Const TABLE_GROUPS As String = "tblReportGroups"
Private Const TABLE_SCHEMES As String = "tblOutputSchemes"
Private Const TABLE_CHARTS As String = "tblReportCharts"

Private Const COL_EXPORT_ENABLED As String = "是否导出"
Private Const COL_CATEGORY As String = "分类"
Private Const COL_SERIES As String = "系列"
Private Const COL_TRUST_CODE As String = "信托计划代码"
Private Const COL_PRODUCT_NAME As String = "产品名称"
Private Const COL_SEQ As String = "序号"

Public Sub Report00_MigrateConfiguration()
    MigrateConfiguration True
End Sub

Public Sub Report00_MigrateConfigurationSilent()
    MigrateConfiguration False
End Sub

Private Sub MigrateConfiguration(ByVal showResult As Boolean)
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    EnsureProductExportColumn
    EnsureReportConfigurationSheet
    EnsureConfigurationGuideSheet

    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating

    Dim validationErrors As String
    If ReportConfig_Validate(validationErrors) Then
        If showResult Then MsgBox "报表配置迁移完成，静态预检通过。", vbInformation, "报表配置迁移"
    ElseIf showResult Then
        MsgBox "报表配置迁移完成，但静态预检发现问题：" & vbCrLf & vbCrLf & validationErrors, _
               vbExclamation, "报表配置迁移"
    End If
    Exit Sub

CleanFail:
    Dim errorDescription As String
    errorDescription = Err.Description
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    If showResult Then
        MsgBox "报表配置迁移失败" & vbCrLf & vbCrLf & "错误信息：" & errorDescription, _
               vbCritical, "报表配置迁移"
    Else
        Err.Raise vbObjectError + 6003, , errorDescription
    End If
End Sub

Public Sub Report00_ValidateConfiguration()
    Dim validationErrors As String
    If ReportConfig_Validate(validationErrors) Then
        MsgBox "报表配置静态预检通过。", vbInformation, "报表配置预检"
    Else
        MsgBox "报表配置静态预检未通过：" & vbCrLf & vbCrLf & validationErrors, _
               vbExclamation, "报表配置预检"
    End If
End Sub

Public Sub Report00_RequireValidConfiguration()
    Dim validationErrors As String
    If Not ReportConfig_Validate(validationErrors) Then
        Err.Raise vbObjectError + 6002, , "报表配置静态预检未通过：" & vbCrLf & validationErrors
    End If
End Sub

' 返回 Dictionary(产品归属 -> 工作表名称)，仅包含启用分类。
Public Function ReportConfig_GetEnabledCategorySheetMap() As Object
    Dim result As Object
    Set result = CreateTextDictionary()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_REPORT_CONFIG)
    Dim table As ListObject
    Set table = ws.ListObjects(TABLE_CATEGORIES)

    If Not table.DataBodyRange Is Nothing Then
        Dim categoryNames() As String
        Dim sheetNames() As String
        Dim sortOrders() As Long
        Dim count As Long
        Dim r As Long
        For r = 1 To table.DataBodyRange.Rows.Count
            If IsEnabledValue(TableValue(table, r, "是否启用")) Then
                Dim categoryName As String
                categoryName = TableText(table, r, "分类")
                If Len(categoryName) > 0 Then
                    count = count + 1
                    ReDim Preserve categoryNames(1 To count)
                    ReDim Preserve sheetNames(1 To count)
                    ReDim Preserve sortOrders(1 To count)
                    categoryNames(count) = categoryName
                    sheetNames(count) = TableText(table, r, "工作表名称")
                    sortOrders(count) = CLng(TableValue(table, r, "工作表顺序"))
                End If
            End If
        Next r

        Dim i As Long
        Dim j As Long
        For i = 1 To count - 1
            For j = i + 1 To count
                If sortOrders(i) > sortOrders(j) Then
                    Dim tempText As String
                    Dim tempOrder As Long
                    tempText = categoryNames(i): categoryNames(i) = categoryNames(j): categoryNames(j) = tempText
                    tempText = sheetNames(i): sheetNames(i) = sheetNames(j): sheetNames(j) = tempText
                    tempOrder = sortOrders(i): sortOrders(i) = sortOrders(j): sortOrders(j) = tempOrder
                End If
            Next j
        Next i

        For i = 1 To count
            result.Add categoryNames(i), sheetNames(i)
        Next i
    End If

    Set ReportConfig_GetEnabledCategorySheetMap = result
End Function

' 返回正式报表引用的信托计划代码集合。
Public Function ReportConfig_GetRequiredChartCodeSet() As Object
    Dim result As Object
    Set result = CreateTextDictionary()

    Dim categoryMap As Object
    Set categoryMap = ReportConfig_GetEnabledCategorySheetMap()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_REPORT_CONFIG)
    Dim table As ListObject
    Set table = ws.ListObjects(TABLE_CHARTS)

    If Not table.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To table.DataBodyRange.Rows.Count
            Dim categoryName As String
            Dim trustCode As String
            categoryName = TableText(table, r, "分类")
            trustCode = TableText(table, r, "信托计划代码")
            If categoryMap.Exists(categoryName) And Len(trustCode) > 0 Then result(trustCode) = True
        Next r
    End If

    Set ReportConfig_GetRequiredChartCodeSet = result
End Function

' 返回 Dictionary(信托计划代码 -> |红|、|蓝| 或 |红|蓝|)。
Public Function ReportConfig_GetChartThemeMap() As Object
    Dim result As Object
    Set result = CreateTextDictionary()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_REPORT_CONFIG)
    Dim categoryTable As ListObject
    Dim chartTable As ListObject
    Set categoryTable = ws.ListObjects(TABLE_CATEGORIES)
    Set chartTable = ws.ListObjects(TABLE_CHARTS)

    Dim categoryThemes As Object
    Set categoryThemes = CreateTextDictionary()
    If Not categoryTable.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To categoryTable.DataBodyRange.Rows.Count
            If IsEnabledValue(TableValue(categoryTable, r, "是否启用")) Then
                categoryThemes(TableText(categoryTable, r, "分类")) = TableText(categoryTable, r, "视觉主题")
            End If
        Next r
    End If

    If Not chartTable.DataBodyRange Is Nothing Then
        For r = 1 To chartTable.DataBodyRange.Rows.Count
            Dim categoryName As String
            Dim trustCode As String
            categoryName = TableText(chartTable, r, "分类")
            trustCode = TableText(chartTable, r, "信托计划代码")
            If categoryThemes.Exists(categoryName) And Len(trustCode) > 0 Then
                Dim themeToken As String
                themeToken = "|" & CStr(categoryThemes(categoryName)) & "|"
                If Not result.Exists(trustCode) Then result.Add trustCode, "|"
                If InStr(1, CStr(result(trustCode)), themeToken, vbTextCompare) = 0 Then
                    result(trustCode) = CStr(result(trustCode)) & CStr(categoryThemes(categoryName)) & "|"
                End If
            End If
        Next r
    End If

    Set ReportConfig_GetChartThemeMap = result
End Function

' 以下查询返回只读快照；每行是以配置表头为键的 Dictionary。
Public Function ReportConfig_GetCategoryDefinitions() As Collection
    Set ReportConfig_GetCategoryDefinitions = ReadConfigurationRows(TABLE_CATEGORIES)
End Function

Public Function ReportConfig_GetGroupDefinitions() As Collection
    Set ReportConfig_GetGroupDefinitions = ReadConfigurationRows(TABLE_GROUPS)
End Function

Public Function ReportConfig_GetSchemeDefinitions() As Collection
    Set ReportConfig_GetSchemeDefinitions = ReadConfigurationRows(TABLE_SCHEMES)
End Function

Public Function ReportConfig_GetChartDefinitions() As Collection
    Set ReportConfig_GetChartDefinitions = ReadConfigurationRows(TABLE_CHARTS)
End Function

Public Function ReportConfig_Validate(ByRef validationErrors As String) As Boolean
    validationErrors = vbNullString

    Dim wsCategory As Worksheet
    Dim wsInfo As Worksheet
    Dim wsConfig As Worksheet
    Set wsCategory = GetWorksheetIfExists(ThisWorkbook, SHEET_PRODUCT_CATEGORY)
    Set wsInfo = GetWorksheetIfExists(ThisWorkbook, SHEET_PRODUCT_INFO)
    Set wsConfig = GetWorksheetIfExists(ThisWorkbook, SHEET_REPORT_CONFIG)

    If wsCategory Is Nothing Then AppendValidationError validationErrors, "缺少工作表：" & SHEET_PRODUCT_CATEGORY
    If wsInfo Is Nothing Then AppendValidationError validationErrors, "缺少工作表：" & SHEET_PRODUCT_INFO
    If wsConfig Is Nothing Then AppendValidationError validationErrors, "缺少工作表：" & SHEET_REPORT_CONFIG & "（请先运行 Report00_MigrateConfiguration）"
    If wsCategory Is Nothing Or wsInfo Is Nothing Or wsConfig Is Nothing Then Exit Function

    Dim categoryTable As ListObject
    Dim groupTable As ListObject
    Dim schemeTable As ListObject
    Dim chartTable As ListObject
    Set categoryTable = GetTableIfExists(wsConfig, TABLE_CATEGORIES)
    Set groupTable = GetTableIfExists(wsConfig, TABLE_GROUPS)
    Set schemeTable = GetTableIfExists(wsConfig, TABLE_SCHEMES)
    Set chartTable = GetTableIfExists(wsConfig, TABLE_CHARTS)

    If categoryTable Is Nothing Then AppendValidationError validationErrors, "报表配置缺少表格：" & TABLE_CATEGORIES
    If groupTable Is Nothing Then AppendValidationError validationErrors, "报表配置缺少表格：" & TABLE_GROUPS
    If schemeTable Is Nothing Then AppendValidationError validationErrors, "报表配置缺少表格：" & TABLE_SCHEMES
    If chartTable Is Nothing Then AppendValidationError validationErrors, "报表配置缺少表格：" & TABLE_CHARTS
    If categoryTable Is Nothing Or groupTable Is Nothing Or schemeTable Is Nothing Or chartTable Is Nothing Then Exit Function

    RequireTableHeaders categoryTable, Array("分类", "是否启用", "工作表名称", "报表主标题", "页内标题", "视觉主题", "工作表顺序"), validationErrors
    RequireTableHeaders groupTable, Array("分类", "产品系列", "展示分组", "分组标题", "分组顺序", "输出字段方案"), validationErrors
    RequireTableHeaders schemeTable, Array("方案名称", "字段1", "字段1标题", "字段2", "字段2标题", "字段3", "字段3标题", "字段4", "字段4标题", "字段5", "字段5标题"), validationErrors
    RequireTableHeaders chartTable, Array("分类", "位置序号", "信托计划代码", "图表类型", "是否必需"), validationErrors
    If Len(validationErrors) > 0 Then Exit Function

    Dim categoryDefinitions As Object
    Dim sheetNames As Object
    Dim schemeDefinitions As Object
    Dim groupDefinitions As Object
    Dim productInfoCodes As Object
    Set categoryDefinitions = CreateTextDictionary()
    Set sheetNames = CreateTextDictionary()
    Set schemeDefinitions = CreateTextDictionary()
    Set groupDefinitions = CreateTextDictionary()
    Set productInfoCodes = LoadProductInfoCodes(wsInfo, validationErrors)

    ValidateCategoryTable categoryTable, categoryDefinitions, sheetNames, validationErrors
    ValidateSchemeTable schemeTable, schemeDefinitions, validationErrors
    ValidateGroupTable groupTable, categoryDefinitions, schemeDefinitions, groupDefinitions, validationErrors
    ValidateProductAssignments wsCategory, categoryDefinitions, groupDefinitions, validationErrors
    ValidateChartTable chartTable, categoryDefinitions, productInfoCodes, validationErrors

    ReportConfig_Validate = (Len(validationErrors) = 0)
End Function

Private Sub EnsureProductExportColumn()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PRODUCT_CATEGORY)

    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)

    Dim exportCol As Long
    If headers.Exists(COL_EXPORT_ENABLED) Then
        exportCol = CLng(headers(COL_EXPORT_ENABLED))
    Else
        exportCol = LastUsedColumn(ws) + 1
        ws.Cells(1, exportCol).Value = COL_EXPORT_ENABLED
        ws.Cells(1, exportCol).Font.Bold = True
    End If

    Dim seqCol As Long
    Dim codeCol As Long
    If headers.Exists(COL_SEQ) Then seqCol = CLng(headers(COL_SEQ))
    If headers.Exists(COL_TRUST_CODE) Then codeCol = CLng(headers(COL_TRUST_CODE))

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        If (seqCol > 0 And Len(NormalizeText(ws.Cells(r, seqCol).Value)) > 0) Or _
           (codeCol > 0 And Len(NormalizeText(ws.Cells(r, codeCol).Value)) > 0) Then
            If Len(NormalizeText(ws.Cells(r, exportCol).Value)) = 0 Then
                ws.Cells(r, exportCol).Value = "是"
            End If
        End If
    Next r

    ws.Columns(exportCol).AutoFit
End Sub

Private Sub EnsureReportConfigurationSheet()
    Dim ws As Worksheet
    Set ws = GetWorksheetIfExists(ThisWorkbook, SHEET_REPORT_CONFIG)
    If Not ws Is Nothing Then
        If TableExists(ws, TABLE_CATEGORIES) And _
           TableExists(ws, TABLE_GROUPS) And _
           TableExists(ws, TABLE_SCHEMES) And _
           TableExists(ws, TABLE_CHARTS) Then
            Exit Sub
        End If
        Err.Raise vbObjectError + 6001, , "已存在“报表配置”工作表，但配置表格不完整。为避免覆盖业务配置，请先人工检查或删除该工作表后重试。"
    End If

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SHEET_REPORT_CONFIG
    ws.Cells.Font.Name = "微软雅黑"

    WriteCategoryConfiguration ws, 1
    WriteGroupConfiguration ws, 10
    WriteSchemeConfiguration ws, 27
    WriteChartConfiguration ws, 38
    WriteValidationLists ws
    ApplyConfigurationValidation ws

    ws.Columns("A:K").AutoFit
    ws.Activate
    ws.Range("A1").Select
End Sub

Private Sub WriteCategoryConfiguration(ByVal ws As Worksheet, ByVal titleRow As Long)
    Dim headers As Variant
    Dim rows As Variant
    headers = Array("分类", "是否启用", "工作表名称", "报表主标题", "页内标题", "视觉主题", "工作表顺序")
    rows = Array( _
        Array("稳享长期限", "是", "稳享长期限", "", "热销产品", "红", 1), _
        Array("直销", "是", "直销", "", "热销产品", "红", 2), _
        Array("交行代销", "是", "交行代销", "", "热销产品", "红", 3), _
        Array("圆融安享", "是", "圆融安享", "交银国信·圆融安享汇益固收稳健系列信托计划" & vbLf & "历史到期产品收益情况", "", "蓝", 4), _
        Array("恒丰", "是", "恒丰", "", "热销产品", "红", 5))
    CreateConfigTable ws, titleRow, "一、分类配置", headers, rows, TABLE_CATEGORIES
End Sub

Private Sub WriteGroupConfiguration(ByVal ws As Worksheet, ByVal titleRow As Long)
    Dim headers As Variant
    Dim rows As Variant
    headers = Array("分类", "产品系列", "展示分组", "分组标题", "分组顺序", "输出字段方案")
    rows = Array( _
        Array("稳享长期限", "汇益稳享系列兑付业绩", "稳享长期限", "汇益稳享系列兑付业绩", 1, "长期限"), _
        Array("直销", "汇益稳健系列兑付业绩", "直销稳健", "汇益稳健系列兑付业绩", 1, "净值型"), _
        Array("直销", "汇益系列兑付业绩", "直销汇益", "汇益系列兑付业绩", 2, "报价型"), _
        Array("直销", "交鑫致远系列兑付业绩", "直销长期限", "交鑫致远系列兑付业绩", 3, "成立以来"), _
        Array("交行代销", "汇益稳健", "交行稳健", "汇益稳健系列业绩", 1, "净值型"), _
        Array("交行代销", "汇益", "交行汇益", "汇益系列业绩", 2, "报价型"), _
        Array("交行代销", "汇益稳享", "交行稳享", "汇益稳享系列业绩", 3, "双周期"), _
        Array("交行代销", "蓝色港湾", "蓝色港湾", "蓝色港湾系列业绩", 4, "长期限"), _
        Array("圆融安享", "日开", "圆融日开", "日开", 1, "净值型"), _
        Array("圆融安享", "周开", "圆融周开", "周开", 2, "净值型"), _
        Array("圆融安享", "月开", "圆融月开", "月开", 3, "净值型"), _
        Array("恒丰", "汇益稳享系列兑付业绩", "恒丰长期限", "汇益稳享系列兑付业绩", 1, "长期限"))
    CreateConfigTable ws, titleRow, "二、展示分组配置", headers, rows, TABLE_GROUPS
End Sub

Private Sub WriteSchemeConfiguration(ByVal ws As Worksheet, ByVal titleRow As Long)
    Dim headers As Variant
    Dim rows As Variant
    headers = Array("方案名称", "字段1", "字段1标题", "字段2", "字段2标题", "字段3", "字段3标题", "字段4", "字段4标题", "字段5", "字段5标题")
    rows = Array( _
        Array("长期限", "PRODUCT_NAME", "产品名称", "NEXT_OPEN", "下一开放日", "OPEN_FREQUENCY", "开放频率", "ELAPSED_DAYS", "当前周期运作天数", "CURRENT_PERIOD_ANNUAL", "当前周期年化"), _
        Array("净值型", "PRODUCT_NAME", "净值型产品名称", "NEXT_OPEN", "下一开放日", "OPEN_FREQUENCY", "开放频率", "ANNUAL_7D", "7日年化收益率", "ANNUAL_28D", "28日年化收益率"), _
        Array("报价型", "PRODUCT_NAME", "报价型产品名称", "NEXT_OPEN", "下一开放日", "OPEN_FREQUENCY", "开放频率", "BENCHMARK_RATE", "基准", "ANNUAL_28D", "28日年化收益率"), _
        Array("成立以来", "PRODUCT_NAME", "长期限产品名称", "NEXT_OPEN", "下一开放日", "OPEN_FREQUENCY", "开放频率", "", "", "INCEPTION_ANNUAL", "成立以来年化收益率"), _
        Array("双周期", "PRODUCT_NAME", "长期限产品名称", "NEXT_OPEN", "下一开放日", "PREV_PERIOD_ANNUAL", "上期年化收益率", "ELAPSED_DAYS", "当前周期运作天数", "CURRENT_PERIOD_ANNUAL", "当期年化收益率"))
    CreateConfigTable ws, titleRow, "三、输出字段方案", headers, rows, TABLE_SCHEMES
End Sub

Private Sub WriteChartConfiguration(ByVal ws As Worksheet, ByVal titleRow As Long)
    Dim headers As Variant
    Dim rows As Variant
    headers = Array("分类", "位置序号", "信托计划代码", "图表类型", "是否必需")
    rows = Array( _
        Array("稳享长期限", 1, "O22600", "组合图", "是"), _
        Array("稳享长期限", 2, "PG2800", "组合图", "是"), _
        Array("交行代销", 1, "O25900", "组合图", "是"), _
        Array("交行代销", 2, "O51900", "组合图", "是"), _
        Array("交行代销", 3, "OE0800", "组合图", "是"), _
        Array("直销", 1, "N73400", "组合图", "是"), _
        Array("直销", 2, "O46900", "组合图", "是"), _
        Array("直销", 3, "P83600", "组合图", "是"), _
        Array("圆融安享", 1, "OL0800", "组合图", "是"), _
        Array("圆融安享", 2, "OH1100", "组合图", "是"), _
        Array("圆融安享", 3, "OH8400", "组合图", "是"), _
        Array("恒丰", 1, "PE3300", "组合图", "是"), _
        Array("恒丰", 2, "PH3400", "组合图", "是"))
    CreateConfigTable ws, titleRow, "四、图表位置配置", headers, rows, TABLE_CHARTS
End Sub

Private Sub CreateConfigTable(ByVal ws As Worksheet, ByVal titleRow As Long, ByVal titleText As String, _
                              ByVal headers As Variant, ByVal rows As Variant, ByVal tableName As String)
    ws.Cells(titleRow, 1).Value = titleText
    ws.Cells(titleRow, 1).Font.Bold = True
    ws.Cells(titleRow, 1).Font.Size = 12

    Dim headerRow As Long
    headerRow = titleRow + 1

    Dim c As Long
    For c = LBound(headers) To UBound(headers)
        ws.Cells(headerRow, c + 1).Value = headers(c)
    Next c

    Dim r As Long
    Dim rowValues As Variant
    For r = LBound(rows) To UBound(rows)
        rowValues = rows(r)
        For c = LBound(rowValues) To UBound(rowValues)
            ws.Cells(headerRow + 1 + r - LBound(rows), c + 1).Value = rowValues(c)
        Next c
    Next r

    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = headerRow + 1 + UBound(rows) - LBound(rows)
    lastCol = UBound(headers) - LBound(headers) + 1

    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))

    Dim table As ListObject
    Set table = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    table.Name = tableName
    table.TableStyle = "TableStyleMedium2"
End Sub

Private Sub WriteValidationLists(ByVal ws As Worksheet)
    ws.Range("M1").Value = "是否值"
    ws.Range("M2").Value = "是"
    ws.Range("M3").Value = "否"
    ws.Range("N1").Value = "主题值"
    ws.Range("N2").Value = "红"
    ws.Range("N3").Value = "蓝"
    ws.Range("O1").Value = "指标值"

    Dim fields As Variant
    fields = Array("PRODUCT_NAME", "NEXT_OPEN", "OPEN_FREQUENCY", "BENCHMARK_RATE", _
                   "ELAPSED_DAYS", "PREV_PERIOD_ANNUAL", "CURRENT_PERIOD_ANNUAL", _
                   "ANNUAL_7D", "ANNUAL_28D", "INCEPTION_ANNUAL")
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        ws.Cells(i + 2, 15).Value = fields(i)
    Next i

    ws.Range("P1").Value = "图表类型值"
    ws.Range("P2").Value = "组合图"
    ws.Range("Q1").Value = "确认状态值"
    ws.Range("Q2").Value = "待确认"
    ws.Range("Q3").Value = "待人工处理"
    ws.Range("Q4").Value = "确认"
    ws.Range("Q5").Value = "拒绝"
    ws.Range("Q6").Value = "已写入"
    ws.Range("Q7").Value = "已人工处理"

    AddOrReplaceWorkbookName "ReportBooleanValues", "='" & ws.Name & "'!$M$2:$M$3"
    AddOrReplaceWorkbookName "ReportThemeValues", "='" & ws.Name & "'!$N$2:$N$3"
    AddOrReplaceWorkbookName "ReportMetricValues", "='" & ws.Name & "'!$O$2:$O$11"
    AddOrReplaceWorkbookName "ReportChartTypeValues", "='" & ws.Name & "'!$P$2:$P$2"
    AddOrReplaceWorkbookName "OpenDateReviewValues", "='" & ws.Name & "'!$Q$2:$Q$7"
    ws.Columns("M:Q").Hidden = True
End Sub

Private Sub ApplyConfigurationValidation(ByVal ws As Worksheet)
    ApplyListValidation ws.ListObjects(TABLE_CATEGORIES).ListColumns("是否启用").DataBodyRange, "=ReportBooleanValues"
    ApplyListValidation ws.ListObjects(TABLE_CATEGORIES).ListColumns("视觉主题").DataBodyRange, "=ReportThemeValues"
    ApplyListValidation ws.ListObjects(TABLE_CHARTS).ListColumns("图表类型").DataBodyRange, "=ReportChartTypeValues"
    ApplyListValidation ws.ListObjects(TABLE_CHARTS).ListColumns("是否必需").DataBodyRange, "=ReportBooleanValues"

    Dim schemeTable As ListObject
    Set schemeTable = ws.ListObjects(TABLE_SCHEMES)
    Dim slot As Long
    For slot = 1 To 5
        ApplyListValidation schemeTable.ListColumns("字段" & slot).DataBodyRange, "=ReportMetricValues"
    Next slot

    Dim productSheet As Worksheet
    Set productSheet = ThisWorkbook.Worksheets(SHEET_PRODUCT_CATEGORY)
    Dim headers As Object
    Set headers = BuildHeaderMap(productSheet, 1)
    Dim lastRow As Long
    lastRow = LastUsedRow(productSheet)
    If lastRow >= 2 And headers.Exists(COL_EXPORT_ENABLED) Then
        ApplyListValidation productSheet.Range(productSheet.Cells(2, CLng(headers(COL_EXPORT_ENABLED))), _
                                               productSheet.Cells(lastRow, CLng(headers(COL_EXPORT_ENABLED)))), _
                            "=ReportBooleanValues"
    End If
End Sub

Private Sub ApplyListValidation(ByVal targetRange As Range, ByVal formulaText As String)
    If targetRange Is Nothing Then Exit Sub
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formulaText
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With
End Sub

Private Sub AddOrReplaceWorkbookName(ByVal nameText As String, ByVal refersToText As String)
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nameText, RefersTo:=refersToText
End Sub

Private Sub EnsureConfigurationGuideSheet()
    Dim ws As Worksheet
    Set ws = GetWorksheetIfExists(ThisWorkbook, SHEET_CONFIG_GUIDE)
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SHEET_CONFIG_GUIDE
    ws.Cells.Font.Name = "微软雅黑"

    ws.Range("A1").Value = "产品净值周报配置说明"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    Dim guideLines As Variant
    guideLines = Array( _
        "1. 产品增减：在“产品分类”中维护“是否导出”，不要删除历史产品行。", _
        "2. 产品归属：每个产品只能有一个分类；新分类必须先在“报表配置”登记。", _
        "3. 展示分组：在报表配置中用“分类 + 产品系列”映射展示分组和输出字段方案。", _
        "4. 输出字段：只能从受控指标中选择，不要填写公式或自造指标代码。", _
        "5. 图表：使用信托计划代码和位置序号，不要填写产品名称或单元格地址。", _
        "6. 开放日：推算结果先进入“开放日待确认”，确认后才能写入开放日台账。", _
        "7. 正式输出：必需图表必须与报表基准日期一致；缺图时一键流程停止。", _
        "8. 修改配置后先运行 Report00_ValidateConfiguration，再运行正式流程。")

    Dim i As Long
    For i = LBound(guideLines) To UBound(guideLines)
        ws.Cells(i + 3, 1).Value = guideLines(i)
    Next i

    ws.Columns("A").ColumnWidth = 110
    ws.Range("A3:A" & UBound(guideLines) + 3).WrapText = True
    ws.Range("A3:A" & UBound(guideLines) + 3).Rows.AutoFit
End Sub

Private Sub ValidateCategoryTable(ByVal table As ListObject, ByVal categoryDefinitions As Object, _
                                  ByVal sheetNames As Object, ByRef validationErrors As String)
    If table.DataBodyRange Is Nothing Then
        AppendValidationError validationErrors, "分类配置没有数据。"
        Exit Sub
    End If

    Dim r As Long
    For r = 1 To table.DataBodyRange.Rows.Count
        Dim categoryName As String
        Dim enabledText As String
        Dim sheetName As String
        Dim themeName As String
        categoryName = TableText(table, r, "分类")
        enabledText = TableText(table, r, "是否启用")
        sheetName = TableText(table, r, "工作表名称")
        themeName = TableText(table, r, "视觉主题")

        If Len(categoryName) = 0 Then
            AppendValidationError validationErrors, "分类配置第" & r & "行缺少分类。"
        ElseIf categoryDefinitions.Exists(categoryName) Then
            AppendValidationError validationErrors, "分类配置重复：" & categoryName
        Else
            categoryDefinitions.Add categoryName, IsEnabledValue(enabledText)
        End If

        If Len(enabledText) > 0 And Not IsRecognizedBoolean(enabledText) Then
            AppendValidationError validationErrors, "分类“" & categoryName & "”的是否启用值无效：" & enabledText
        End If

        If IsEnabledValue(enabledText) Then
            If Len(sheetName) = 0 Then
                AppendValidationError validationErrors, "分类“" & categoryName & "”缺少工作表名称。"
            ElseIf Not IsValidWorksheetName(sheetName) Then
                AppendValidationError validationErrors, "分类“" & categoryName & "”的工作表名称无效：" & sheetName
            ElseIf sheetNames.Exists(sheetName) Then
                AppendValidationError validationErrors, "工作表名称重复：" & sheetName
            Else
                sheetNames.Add sheetName, True
            End If

            If themeName <> "红" And themeName <> "蓝" Then
                AppendValidationError validationErrors, "分类“" & categoryName & "”的视觉主题只允许“红”或“蓝”。"
            End If

            If Not IsPositiveLong(TableValue(table, r, "工作表顺序")) Then
                AppendValidationError validationErrors, "分类“" & categoryName & "”的工作表顺序必须为正整数。"
            End If
        End If
    Next r
End Sub

Private Sub ValidateSchemeTable(ByVal table As ListObject, ByVal schemeDefinitions As Object, _
                                ByRef validationErrors As String)
    If table.DataBodyRange Is Nothing Then
        AppendValidationError validationErrors, "输出字段方案没有数据。"
        Exit Sub
    End If

    Dim allowedFields As Object
    Set allowedFields = CreateTextDictionary()
    AddAllowedField allowedFields, "PRODUCT_NAME"
    AddAllowedField allowedFields, "NEXT_OPEN"
    AddAllowedField allowedFields, "OPEN_FREQUENCY"
    AddAllowedField allowedFields, "BENCHMARK_RATE"
    AddAllowedField allowedFields, "ELAPSED_DAYS"
    AddAllowedField allowedFields, "PREV_PERIOD_ANNUAL"
    AddAllowedField allowedFields, "CURRENT_PERIOD_ANNUAL"
    AddAllowedField allowedFields, "ANNUAL_7D"
    AddAllowedField allowedFields, "ANNUAL_28D"
    AddAllowedField allowedFields, "INCEPTION_ANNUAL"

    Dim r As Long
    For r = 1 To table.DataBodyRange.Rows.Count
        Dim schemeName As String
        schemeName = TableText(table, r, "方案名称")
        If Len(schemeName) = 0 Then
            AppendValidationError validationErrors, "输出字段方案第" & r & "行缺少方案名称。"
        ElseIf schemeDefinitions.Exists(schemeName) Then
            AppendValidationError validationErrors, "输出字段方案重复：" & schemeName
        Else
            schemeDefinitions.Add schemeName, True
        End If

        Dim seenFields As Object
        Set seenFields = CreateTextDictionary()

        Dim slot As Long
        For slot = 1 To 5
            Dim fieldName As String
            Dim fieldTitle As String
            fieldName = UCase$(TableText(table, r, "字段" & slot))
            fieldTitle = TableText(table, r, "字段" & slot & "标题")

            If Len(fieldName) > 0 Then
                If Not allowedFields.Exists(fieldName) Then
                    AppendValidationError validationErrors, "方案“" & schemeName & "”包含未知指标：" & fieldName
                ElseIf seenFields.Exists(fieldName) Then
                    AppendValidationError validationErrors, "方案“" & schemeName & "”重复使用指标：" & fieldName
                Else
                    seenFields.Add fieldName, True
                End If
                If Len(fieldTitle) = 0 Then
                    AppendValidationError validationErrors, "方案“" & schemeName & "”的字段" & slot & "缺少显示标题。"
                End If
            ElseIf Len(fieldTitle) > 0 Then
                AppendValidationError validationErrors, "方案“" & schemeName & "”的字段" & slot & "为空，但仍填写了显示标题。"
            End If
        Next slot

        If UCase$(TableText(table, r, "字段1")) <> "PRODUCT_NAME" Then
            AppendValidationError validationErrors, "方案“" & schemeName & "”的字段1必须是 PRODUCT_NAME。"
        End If
    Next r
End Sub

Private Sub ValidateGroupTable(ByVal table As ListObject, ByVal categoryDefinitions As Object, _
                               ByVal schemeDefinitions As Object, ByVal groupDefinitions As Object, _
                               ByRef validationErrors As String)
    If table.DataBodyRange Is Nothing Then
        AppendValidationError validationErrors, "展示分组配置没有数据。"
        Exit Sub
    End If

    Dim groupSchemes As Object
    Set groupSchemes = CreateTextDictionary()

    Dim r As Long
    For r = 1 To table.DataBodyRange.Rows.Count
        Dim categoryName As String
        Dim seriesName As String
        Dim groupName As String
        Dim schemeName As String
        Dim mapKey As String
        Dim groupKey As String

        categoryName = TableText(table, r, "分类")
        seriesName = TableText(table, r, "产品系列")
        groupName = TableText(table, r, "展示分组")
        schemeName = TableText(table, r, "输出字段方案")
        mapKey = categoryName & ChrW$(30) & seriesName
        groupKey = categoryName & ChrW$(30) & groupName

        If Len(categoryName) = 0 Or Len(seriesName) = 0 Or Len(groupName) = 0 Then
            AppendValidationError validationErrors, "展示分组配置第" & r & "行的分类、产品系列和展示分组均不能为空。"
        ElseIf groupDefinitions.Exists(mapKey) Then
            AppendValidationError validationErrors, "分类与产品系列映射重复：" & categoryName & " / " & seriesName
        Else
            groupDefinitions.Add mapKey, groupName
        End If

        If Not categoryDefinitions.Exists(categoryName) Then
            AppendValidationError validationErrors, "展示分组引用未知分类：" & categoryName
        End If
        If Not schemeDefinitions.Exists(schemeName) Then
            AppendValidationError validationErrors, "展示分组“" & groupName & "”引用未知输出字段方案：" & schemeName
        End If

        If groupSchemes.Exists(groupKey) Then
            If StrComp(CStr(groupSchemes(groupKey)), schemeName, vbTextCompare) <> 0 Then
                AppendValidationError validationErrors, "展示分组“" & categoryName & " / " & groupName & "”配置了多个输出字段方案。"
            End If
        Else
            groupSchemes.Add groupKey, schemeName
        End If

        If Not IsPositiveLong(TableValue(table, r, "分组顺序")) Then
            AppendValidationError validationErrors, "展示分组“" & categoryName & " / " & groupName & "”的分组顺序必须为正整数。"
        End If
    Next r
End Sub

Private Sub ValidateProductAssignments(ByVal ws As Worksheet, ByVal categoryDefinitions As Object, _
                                       ByVal groupDefinitions As Object, ByRef validationErrors As String)
    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)

    Dim requiredHeaders As Variant
    requiredHeaders = Array(COL_TRUST_CODE, COL_PRODUCT_NAME, COL_CATEGORY, COL_SERIES, COL_EXPORT_ENABLED)

    Dim i As Long
    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        If Not headers.Exists(CStr(requiredHeaders(i))) Then
            AppendValidationError validationErrors, SHEET_PRODUCT_CATEGORY & "缺少字段：" & CStr(requiredHeaders(i))
        End If
    Next i
    If Len(validationErrors) > 0 Then Exit Sub

    Dim codeCategories As Object
    Set codeCategories = CreateTextDictionary()

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        Dim code As String
        Dim categoryName As String
        Dim seriesName As String
        Dim exportText As String
        code = NormalizeText(ws.Cells(r, CLng(headers(COL_TRUST_CODE))).Value)
        If Len(code) = 0 Then GoTo ContinueRow

        categoryName = NormalizeText(ws.Cells(r, CLng(headers(COL_CATEGORY))).Value)
        seriesName = NormalizeText(ws.Cells(r, CLng(headers(COL_SERIES))).Value)
        exportText = NormalizeText(ws.Cells(r, CLng(headers(COL_EXPORT_ENABLED))).Value)

        If Len(exportText) > 0 And Not IsRecognizedBoolean(exportText) Then
            AppendValidationError validationErrors, "产品分类第" & r & "行的是否导出值无效：" & exportText
        End If

        If IsEnabledValue(exportText) Then
            If Len(NormalizeText(ws.Cells(r, CLng(headers(COL_PRODUCT_NAME))).Value)) = 0 Then
                AppendValidationError validationErrors, "产品“" & code & "”缺少产品名称。"
            End If
            If Not categoryDefinitions.Exists(categoryName) Then
                AppendValidationError validationErrors, "产品“" & code & "”引用未配置分类：" & categoryName
            ElseIf Not CBool(categoryDefinitions(categoryName)) Then
                AppendValidationError validationErrors, "产品“" & code & "”启用导出，但分类“" & categoryName & "”已停用。"
            End If

            Dim mapKey As String
            mapKey = categoryName & ChrW$(30) & seriesName
            If Not groupDefinitions.Exists(mapKey) Then
                AppendValidationError validationErrors, "产品“" & code & "”没有展示分组映射：" & categoryName & " / " & seriesName
            End If
        End If

        If codeCategories.Exists(code) Then
            If StrComp(CStr(codeCategories(code)), categoryName, vbTextCompare) <> 0 Then
                AppendValidationError validationErrors, "信托计划代码“" & code & "”存在多个产品归属：" & codeCategories(code) & "、" & categoryName
            End If
        Else
            codeCategories.Add code, categoryName
        End If
ContinueRow:
    Next r
End Sub

Private Sub ValidateChartTable(ByVal table As ListObject, ByVal categoryDefinitions As Object, _
                               ByVal productInfoCodes As Object, ByRef validationErrors As String)
    If table.DataBodyRange Is Nothing Then Exit Sub

    Dim placements As Object
    Set placements = CreateTextDictionary()

    Dim r As Long
    For r = 1 To table.DataBodyRange.Rows.Count
        Dim categoryName As String
        Dim code As String
        Dim chartType As String
        Dim requiredText As String
        categoryName = TableText(table, r, "分类")
        code = TableText(table, r, "信托计划代码")
        chartType = TableText(table, r, "图表类型")
        requiredText = TableText(table, r, "是否必需")

        If Not categoryDefinitions.Exists(categoryName) Then
            AppendValidationError validationErrors, "图表配置引用未知分类：" & categoryName
        End If
        If Len(code) = 0 Or Not productInfoCodes.Exists(code) Then
            AppendValidationError validationErrors, "图表配置引用未登记信托计划代码：" & code
        End If
        If chartType <> "组合图" Then
            AppendValidationError validationErrors, "图表配置目前只支持“组合图”：" & categoryName & " / " & code
        End If
        If Len(requiredText) > 0 And Not IsRecognizedBoolean(requiredText) Then
            AppendValidationError validationErrors, "图表“" & categoryName & " / " & code & "”的是否必需值无效：" & requiredText
        End If
        If Not IsPositiveLong(TableValue(table, r, "位置序号")) Then
            AppendValidationError validationErrors, "图表“" & categoryName & " / " & code & "”的位置序号必须为正整数。"
        Else
            Dim placementKey As String
            placementKey = categoryName & ChrW$(30) & CStr(CLng(TableValue(table, r, "位置序号")))
            If placements.Exists(placementKey) Then
                AppendValidationError validationErrors, "图表位置重复：" & categoryName & " / " & TableText(table, r, "位置序号")
            Else
                placements.Add placementKey, True
            End If
        End If
    Next r
End Sub

Private Function LoadProductInfoCodes(ByVal ws As Worksheet, ByRef validationErrors As String) As Object
    Dim result As Object
    Set result = CreateTextDictionary()

    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)
    If Not headers.Exists(COL_TRUST_CODE) Then
        AppendValidationError validationErrors, SHEET_PRODUCT_INFO & "缺少字段：" & COL_TRUST_CODE
        Set LoadProductInfoCodes = result
        Exit Function
    End If

    Dim codeCol As Long
    codeCol = CLng(headers(COL_TRUST_CODE))

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        Dim code As String
        code = NormalizeText(ws.Cells(r, codeCol).Value)
        If Len(code) > 0 Then
            If Not result.Exists(code) Then result.Add code, True
        End If
    Next r

    Set LoadProductInfoCodes = result
End Function

Private Sub RequireTableHeaders(ByVal table As ListObject, ByVal requiredHeaders As Variant, _
                                ByRef validationErrors As String)
    Dim i As Long
    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        If TableColumnIndex(table, CStr(requiredHeaders(i))) = 0 Then
            AppendValidationError validationErrors, "表格“" & table.Name & "”缺少字段：" & CStr(requiredHeaders(i))
        End If
    Next i
End Sub

Private Function ReadConfigurationRows(ByVal tableName As String) As Collection
    Dim result As New Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_REPORT_CONFIG)
    Dim table As ListObject
    Set table = ws.ListObjects(tableName)

    If Not table.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To table.DataBodyRange.Rows.Count
            Dim rowData As Object
            Set rowData = CreateTextDictionary()
            Dim c As Long
            For c = 1 To table.ListColumns.Count
                rowData(table.ListColumns(c).Name) = table.DataBodyRange.Cells(r, c).Value
            Next c
            result.Add rowData
        Next r
    End If

    Set ReadConfigurationRows = result
End Function

Private Function TableColumnIndex(ByVal table As ListObject, ByVal columnName As String) As Long
    Dim column As ListColumn
    For Each column In table.ListColumns
        If StrComp(NormalizeText(column.Name), columnName, vbTextCompare) = 0 Then
            TableColumnIndex = column.Index
            Exit Function
        End If
    Next column
End Function

Private Function TableValue(ByVal table As ListObject, ByVal bodyRow As Long, ByVal columnName As String) As Variant
    Dim columnIndex As Long
    columnIndex = TableColumnIndex(table, columnName)
    If columnIndex = 0 Or table.DataBodyRange Is Nothing Then
        TableValue = Empty
    Else
        TableValue = table.DataBodyRange.Cells(bodyRow, columnIndex).Value
    End If
End Function

Private Function TableText(ByVal table As ListObject, ByVal bodyRow As Long, ByVal columnName As String) As String
    TableText = NormalizeText(TableValue(table, bodyRow, columnName))
End Function

Private Function IsEnabledValue(ByVal value As Variant) As Boolean
    Dim textValue As String
    textValue = UCase$(NormalizeText(value))
    IsEnabledValue = (textValue = "是" Or textValue = "Y" Or textValue = "YES" Or _
                      textValue = "1" Or textValue = "TRUE" Or textValue = "启用")
End Function

Private Function IsRecognizedBoolean(ByVal value As Variant) As Boolean
    Dim textValue As String
    textValue = UCase$(NormalizeText(value))
    IsRecognizedBoolean = (textValue = "是" Or textValue = "Y" Or textValue = "YES" Or _
                           textValue = "1" Or textValue = "TRUE" Or textValue = "启用" Or _
                           textValue = "否" Or textValue = "N" Or textValue = "NO" Or _
                           textValue = "0" Or textValue = "FALSE" Or textValue = "停用")
End Function

Private Function IsPositiveLong(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    If CDbl(value) <> Fix(CDbl(value)) Then Exit Function
    IsPositiveLong = (CLng(value) > 0)
End Function

Private Function IsValidWorksheetName(ByVal sheetName As String) As Boolean
    If Len(sheetName) = 0 Or Len(sheetName) > 31 Then Exit Function

    Dim invalidChars As Variant
    invalidChars = Array("\", "/", "?", "*", "[", "]", ":")

    Dim i As Long
    For i = LBound(invalidChars) To UBound(invalidChars)
        If InStr(1, sheetName, CStr(invalidChars(i)), vbBinaryCompare) > 0 Then Exit Function
    Next i

    IsValidWorksheetName = True
End Function

Private Sub AddAllowedField(ByVal fields As Object, ByVal fieldName As String)
    If Not fields.Exists(fieldName) Then fields.Add fieldName, True
End Sub

Private Sub AppendValidationError(ByRef validationErrors As String, ByVal message As String)
    Const MAX_ERROR_TEXT As Long = 12000
    If Len(validationErrors) >= MAX_ERROR_TEXT Then Exit Sub
    validationErrors = validationErrors & "· " & message & vbCrLf
End Sub

Private Function CreateTextDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set CreateTextDictionary = dict
End Function

Private Function GetWorksheetIfExists(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetIfExists = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetTableIfExists(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set GetTableIfExists = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function TableExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    Dim table As ListObject
    Set table = GetTableIfExists(ws, tableName)
    TableExists = Not (table Is Nothing)
End Function

Private Function BuildHeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim dict As Object
    Set dict = CreateTextDictionary()

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim c As Long
    For c = 1 To lastCol
        Dim headerText As String
        headerText = NormalizeText(ws.Cells(headerRow, c).Value)
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
