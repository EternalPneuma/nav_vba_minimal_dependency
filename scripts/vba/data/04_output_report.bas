' Data04_ExportDisplayReport：按报表配置生成 yyyyMMdd-产品收益展示.xlsx

Option Explicit

Private Const SOURCE_FILE_SUFFIX As String = "-上层产品分类表现.xlsx"
Private Const OUTPUT_FILE_SUFFIX As String = "-产品收益展示.xlsx"
Private Const SHEET_PRODUCT_CATEGORY As String = "产品分类"
Private Const SHEET_PRODUCT_INFO As String = "产品信息"
Private Const ASSET_IMAGE_FOLDER As String = "assets\images"
Private Const TITLE_IMAGE_NAME As String = "title.png"
Private Const PRODUCT_CODE_6_MONTH_101 As String = "P83600"
Private Const PRODUCT_CODE_6_MONTH_102 As String = "P83800"

Private Const COL_BASELINE_DATE As String = "基准日期"
Private Const COL_SEQ As String = "序号"
Private Const COL_TRUST_CODE As String = "信托计划代码"
Private Const COL_SERIES As String = "系列"
Private Const COL_CATEGORY As String = "分类"
Private Const COL_PRODUCT_NAME As String = "产品名称"
Private Const COL_NEXT_OPEN As String = "下一开放日"
Private Const COL_THEORETICAL_INTERVAL As String = "理论间隔"
Private Const COL_ELAPSED As String = "运作时间"
Private Const COL_BENCHMARK_RATE As String = "基准收益率"
Private Const COL_PREV_PERIOD_ANNUAL As String = "上一周期年化"
Private Const COL_CURRENT_PERIOD_ANNUAL As String = "当前周期年化"
Private Const COL_7DAY_ANNUAL As String = "7日年化"
Private Const COL_28DAY_ANNUAL As String = "28日年化"
Private Const COL_INCEPTION_ANNUAL As String = "成立以来年化"
Private Const COL_PRODUCT_SHORT As String = "产品简称"

Private Const FONT_RED As String = "微软雅黑"
Private Const FONT_BLUE As String = "宋体"
Private Const COLOR_DARK_RED As Long = &H19198B
Private Const COLOR_DARK_RED_ALT As Long = &HA1370
Private Const COLOR_GOLD As Long = &H8BDCF4
Private Const COLOR_WHITE As Long = &HFFFFFF
Private Const COLOR_BLUE_TITLE As Long = &H7C3702
Private Const COLOR_BLUE_RECORD As Long = &HF8F0EC
Private Const COLOR_BLACK As Long = &H0
Private Const DECORATION_COLUMN_WIDTH As Double = 3
Private Const REPORT_START_COLUMN As Long = 2   ' B列；A/G为装饰列
Private Const REPORT_CONTENT_COLUMN_COUNT As Long = 5   ' B:F

Public Sub Data04_ExportDisplayReport()
    ExportConfiguredDisplayReport
End Sub

Private Sub ExportConfiguredDisplayReport()
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldAskToUpdateLinks As Boolean
    Dim oldCalculation As XlCalculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldAskToUpdateLinks = Application.AskToUpdateLinks
    oldCalculation = Application.Calculation

    Dim wbSource As Workbook
    Dim wbOutput As Workbook
    Dim currentStage As String
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual

    currentStage = "校验报表配置"
    Dim configErrors As String
    If Not ReportConfig_Validate(configErrors) Then
        Err.Raise vbObjectError + 4101, , "报表配置静态预检未通过：" & vbCrLf & configErrors
    End If

    Dim baselineDate As Date
    baselineDate = GetBaselineDateFromProductCategory()

    Dim sourcePath As String
    sourcePath = ThisWorkbook.Path & Application.PathSeparator & _
                 Format$(baselineDate, "yyyymmdd") & SOURCE_FILE_SUFFIX
    If Len(Dir(sourcePath)) = 0 Then Err.Raise vbObjectError + 4102, , "未找到分类表现文件：" & sourcePath

    currentStage = "打开分类表现文件"
    Set wbSource = Workbooks.Open(FileName:=sourcePath, ReadOnly:=True, UpdateLinks:=False)

    currentStage = "读取报表配置"
    Dim categoryDefinitions As Collection
    Dim groupDefinitions As Collection
    Dim schemeDefinitions As Collection
    Dim chartDefinitions As Collection
    Set categoryDefinitions = ReportConfig_GetCategoryDefinitions()
    Set groupDefinitions = ReportConfig_GetGroupDefinitions()
    Set schemeDefinitions = ReportConfig_GetSchemeDefinitions()
    Set chartDefinitions = ReportConfig_GetChartDefinitions()

    Dim enabledCategories As Collection
    Set enabledCategories = GetSortedEnabledCategories(categoryDefinitions)
    currentStage = "创建展示工作簿"
    Set wbOutput = CreateDisplayWorkbook(enabledCategories)

    currentStage = "建立配置索引"
    Dim schemesByName As Object
    Set schemesByName = IndexRowsByKey(schemeDefinitions, "方案名称")
    Dim productShortNames As Object
    Set productShortNames = BuildProductShortNameMap()

    Dim missingRequiredImages As String
    Dim categoryDefinition As Variant
    For Each categoryDefinition In enabledCategories
        Dim categoryName As String
        Dim sheetName As String
        Dim themeName As String
        categoryName = RowText(categoryDefinition, "分类")
        sheetName = RowText(categoryDefinition, "工作表名称")
        themeName = RowText(categoryDefinition, "视觉主题")

        Dim wsSource As Worksheet
        Dim wsTarget As Worksheet
        Set wsSource = wbSource.Worksheets(sheetName)
        Set wsTarget = wbOutput.Worksheets(sheetName)

        currentStage = "初始化分类：" & categoryName
        Dim nextRow As Long
        nextRow = InitializeCategoryPage(wsTarget, categoryDefinition)
        currentStage = "插入分类图表：" & categoryName
        nextRow = InsertConfiguredCharts(wsTarget, categoryName, themeName, nextRow, _
                                         chartDefinitions, productShortNames, baselineDate, missingRequiredImages)
        currentStage = "写入分类分组：" & categoryName
        nextRow = WriteConfiguredGroups(wsSource, wsTarget, categoryName, themeName, nextRow, _
                                        groupDefinitions, schemesByName)
        currentStage = "完成分类页：" & categoryName
        FinishCategoryPage wsTarget, themeName, nextRow, baselineDate
    Next categoryDefinition

    If Len(missingRequiredImages) > 0 Then
        Err.Raise vbObjectError + 4103, , "以下必需图表缺失或日期不一致：" & vbCrLf & missingRequiredImages
    End If

    Dim outputPath As String
    Dim tempPath As String
    outputPath = ThisWorkbook.Path & Application.PathSeparator & _
                 Format$(baselineDate, "yyyymmdd") & OUTPUT_FILE_SUFFIX
    tempPath = ThisWorkbook.Path & Application.PathSeparator & _
               Format$(baselineDate, "yyyymmdd") & "-产品收益展示.tmp.xlsx"

    currentStage = "保存展示报表"
    If Len(Dir(tempPath)) > 0 Then Kill tempPath
    wbOutput.SaveAs FileName:=tempPath, FileFormat:=xlOpenXMLWorkbook
    wbOutput.Close SaveChanges:=False
    Set wbOutput = Nothing
    wbSource.Close SaveChanges:=False
    Set wbSource = Nothing
    ReplaceFileSafely tempPath, outputPath

    RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, _
                            oldAskToUpdateLinks, oldCalculation
    RecordLastRunMessage "Data04 成功：" & outputPath

    ReportPipeline_MsgBox "展示报表生成完成" & vbCrLf & vbCrLf & _
           "基准日期：" & Format$(baselineDate, "yyyy-mm-dd") & vbCrLf & _
           "输出分类：" & enabledCategories.Count & vbCrLf & vbCrLf & _
           "输出文件：" & vbCrLf & outputPath, vbInformation, "展示报表"
    Exit Sub

CleanFail:
    Dim errorDescription As String
    errorDescription = currentStage & "失败：" & Err.Description
    On Error Resume Next
    If Not wbOutput Is Nothing Then wbOutput.Close SaveChanges:=False
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    On Error GoTo 0
    RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, _
                            oldAskToUpdateLinks, oldCalculation
    RecordLastRunMessage "Data04 失败：" & errorDescription
    ReportPipeline_MsgBox "展示报表生成失败" & vbCrLf & vbCrLf & _
           "错误信息：" & errorDescription, vbCritical, "展示报表"
End Sub

Private Function CreateDisplayWorkbook(ByVal categoryDefinitions As Collection) As Workbook
    Dim currentSheetName As String
    On Error GoTo CreateFail
    Dim wb As Workbook
    Set wb = Workbooks.Add

    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop

    Dim isFirst As Boolean
    isFirst = True
    Dim definition As Variant
    For Each definition In categoryDefinitions
        Dim ws As Worksheet
        If isFirst Then
            Set ws = wb.Worksheets(1)
            isFirst = False
        Else
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        End If
        currentSheetName = RowText(definition, "工作表名称")
        ws.Name = currentSheetName
    Next definition

    Set CreateDisplayWorkbook = wb
    Exit Function

CreateFail:
    Dim errorDescription As String
    Dim workbookName As String
    Dim existingSheetName As String
    errorDescription = Err.Description
    On Error Resume Next
    workbookName = wb.Name
    existingSheetName = ws.Name
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 4104, , "创建工作表“" & currentSheetName & "”失败；工作簿=" & _
              workbookName & "；当前工作表=" & existingSheetName & "：" & errorDescription
End Function

Private Function InitializeCategoryPage(ByVal ws As Worksheet, ByVal categoryDefinition As Object) As Long
    Dim themeName As String
    themeName = RowText(categoryDefinition, "视觉主题")
    ws.Cells.Clear

    ApplyBaseThemeLayout ws, themeName

    If themeName = "蓝" Then
        ws.Cells.Font.Name = FONT_BLUE
        With ws.Range("B1:F1")
            .Merge
            .Value = DisplayTextOrDefault(RowValue(categoryDefinition, "报表主标题"), RowText(categoryDefinition, "分类"))
            .Interior.Color = COLOR_BLUE_TITLE
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Name = FONT_BLUE
            .Font.Bold = True
            .Font.Size = 20
            .Font.Color = COLOR_WHITE
        End With
        ws.Range("A1:G1").Interior.Color = COLOR_BLUE_TITLE
        ws.Rows(1).RowHeight = 56
        InitializeCategoryPage = 2
    Else
        ws.Cells.Font.Name = FONT_RED
        AddRedTitleImage ws
        With ws.Range("B3:F3")
            .Merge
            .Value = DisplayTextOrDefault(RowValue(categoryDefinition, "页内标题"), "热销产品")
            .Interior.Color = COLOR_DARK_RED
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 24
            .Font.Color = COLOR_GOLD
        End With
        ws.Range("A3:G3").Interior.Color = COLOR_DARK_RED
        InitializeCategoryPage = 4
    End If
End Function

Private Function InsertConfiguredCharts(ByVal ws As Worksheet, ByVal categoryName As String, _
                                        ByVal themeName As String, ByVal firstRow As Long, _
                                        ByVal chartDefinitions As Collection, ByVal productShortNames As Object, _
                                        ByVal baselineDate As Date, ByRef missingRequiredImages As String) As Long
    Dim rows As Collection
    Set rows = GetSortedRows(chartDefinitions, "分类", categoryName, "位置序号")
    If rows.Count = 0 Then
        InsertConfiguredCharts = firstRow
        Exit Function
    End If

    Dim imageFolder As String
    imageFolder = ThisWorkbook.Path & Application.PathSeparator & _
                  "产品图表_" & Format$(baselineDate, "yyyymmdd") & Application.PathSeparator

    Dim currentRow As Long
    currentRow = firstRow

    Dim definition As Variant
    For Each definition In rows
        Dim trustCode As String
        Dim requiredImage As Boolean
        trustCode = RowText(definition, "信托计划代码")
        requiredImage = IsEnabledValue(RowValue(definition, "是否必需"))

        ' VBA的块内Dim仍是过程级变量；每轮必须显式清空，避免简称缺失时沿用上一位置的图片。
        Dim imageName As String
        Dim imagePath As String
        Dim missingReason As String
        imageName = vbNullString
        imagePath = vbNullString
        missingReason = vbNullString
        If productShortNames.Exists(trustCode) Then
            imageName = CStr(productShortNames(trustCode)) & "_" & themeName & ".png"
            imagePath = imageFolder & imageName
        Else
            missingReason = "产品信息缺少产品简称"
        End If

        Dim targetRange As Range
        Set targetRange = DisplayRangeForRow(ws, themeName, currentRow)
        targetRange.Interior.Color = COLOR_WHITE

        If Len(imagePath) = 0 Or Len(Dir(imagePath)) = 0 Then
            If Len(missingReason) = 0 Then missingReason = "图片文件缺失：" & imageName
            If requiredImage Then
                missingRequiredImages = missingRequiredImages & categoryName & " / " & trustCode & _
                                        " / " & missingReason & vbCrLf
            End If
            targetRange.Merge
            targetRange.Value = "图表缺失：" & trustCode
            targetRange.HorizontalAlignment = xlCenter
            targetRange.VerticalAlignment = xlCenter
            ws.Rows(currentRow).RowHeight = 80
        Else
            InsertPictureFitWidthAndSetRowHeight ws, imagePath, targetRange, currentRow
        End If
        currentRow = currentRow + 1
    Next definition

    InsertConfiguredCharts = currentRow
End Function

Private Function WriteConfiguredGroups(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet, _
                                       ByVal categoryName As String, ByVal themeName As String, _
                                       ByVal firstRow As Long, ByVal groupDefinitions As Collection, _
                                       ByVal schemesByName As Object) As Long
    Dim sourceHeaders As Object
    Set sourceHeaders = BuildHeaderMap(wsSource, 1)

    Dim groups As Collection
    Set groups = GetUniqueSortedGroups(groupDefinitions, categoryName)

    Dim currentRow As Long
    currentRow = firstRow

    Dim groupDefinition As Variant
    For Each groupDefinition In groups
        Dim schemeName As String
        schemeName = RowText(groupDefinition, "输出字段方案")
        If Not schemesByName.Exists(schemeName) Then
            Err.Raise vbObjectError + 4201, , "展示分组引用未知输出字段方案：" & schemeName
        End If

        Dim schemeDefinition As Object
        Set schemeDefinition = schemesByName(schemeName)

        Dim titleRow As Long
        Dim headerRow As Long
        Dim firstDataRow As Long
        titleRow = currentRow
        headerRow = titleRow + 1
        firstDataRow = headerRow + 1

        WriteGroupTitle wsTarget, themeName, titleRow, RowText(groupDefinition, "分组标题")
        WriteSchemeHeaders wsTarget, themeName, headerRow, schemeDefinition

        Dim sourceRows() As Long
        Dim sourceRowCount As Long
        CollectAndSortSourceRows wsSource, sourceHeaders, categoryName, _
                                 RowText(groupDefinition, "展示分组"), groupDefinitions, _
                                 sourceRows, sourceRowCount

        Dim dataIndex As Long
        For dataIndex = 1 To sourceRowCount
            WriteSchemeDataRow wsSource, sourceHeaders, sourceRows(dataIndex), wsTarget, _
                               firstDataRow + dataIndex - 1, themeName, schemeDefinition
        Next dataIndex

        Dim lastBlockRow As Long
        If sourceRowCount = 0 Then
            lastBlockRow = headerRow
        Else
            lastBlockRow = firstDataRow + sourceRowCount - 1
        End If
        FormatGroupBlock wsTarget, themeName, headerRow, lastBlockRow
        If themeName <> "蓝" And sourceRowCount > 0 Then
            ApplyRedFrequencyGroups wsSource, sourceHeaders, sourceRows, sourceRowCount, _
                                    wsTarget, firstDataRow, schemeDefinition
        End If
        currentRow = lastBlockRow + 1
    Next groupDefinition

    WriteConfiguredGroups = currentRow
End Function

Private Sub WriteGroupTitle(ByVal ws As Worksheet, ByVal themeName As String, ByVal rowNumber As Long, _
                            ByVal titleText As String)
    Dim targetRange As Range
    Set targetRange = DisplayRangeForRow(ws, themeName, rowNumber)
    targetRange.Merge
    targetRange.Value = titleText
    targetRange.HorizontalAlignment = xlCenter
    targetRange.VerticalAlignment = xlCenter
    targetRange.Font.Bold = True
    targetRange.Font.Size = IIf(themeName = "蓝", 16, 24)
    targetRange.Font.Color = IIf(themeName = "蓝", COLOR_WHITE, COLOR_GOLD)
    targetRange.Interior.Color = IIf(themeName = "蓝", COLOR_BLUE_TITLE, COLOR_DARK_RED)
End Sub

Private Sub WriteSchemeHeaders(ByVal ws As Worksheet, ByVal themeName As String, ByVal rowNumber As Long, _
                               ByVal scheme As Object)
    Dim startColumn As Long
    startColumn = DisplayStartColumn(themeName)

    Dim slot As Long
    For slot = 1 To 5
        ws.Cells(rowNumber, startColumn + slot - 1).Value = RowText(scheme, "字段" & slot & "标题")
    Next slot
End Sub

Private Sub WriteSchemeDataRow(ByVal wsSource As Worksheet, ByVal sourceHeaders As Object, _
                               ByVal sourceRow As Long, ByVal wsTarget As Worksheet, _
                               ByVal targetRow As Long, ByVal themeName As String, ByVal scheme As Object)
    Dim startColumn As Long
    startColumn = DisplayStartColumn(themeName)

    Dim slot As Long
    For slot = 1 To 5
        Dim fieldId As String
        fieldId = UCase$(RowText(scheme, "字段" & slot))
        WriteConfiguredField wsSource, sourceHeaders, sourceRow, _
                             wsTarget.Cells(targetRow, startColumn + slot - 1), fieldId
    Next slot
End Sub

Private Sub WriteConfiguredField(ByVal wsSource As Worksheet, ByVal headers As Object, _
                                 ByVal sourceRow As Long, ByVal targetCell As Range, _
                                 ByVal fieldId As String)
    If Len(fieldId) = 0 Then
        targetCell.ClearContents
        Exit Sub
    End If

    Dim sourceHeader As String
    Select Case fieldId
        Case "PRODUCT_NAME": sourceHeader = COL_PRODUCT_NAME
        Case "NEXT_OPEN": sourceHeader = COL_NEXT_OPEN
        Case "BENCHMARK_RATE": sourceHeader = COL_BENCHMARK_RATE
        Case "ELAPSED_DAYS": sourceHeader = COL_ELAPSED
        Case "PREV_PERIOD_ANNUAL": sourceHeader = COL_PREV_PERIOD_ANNUAL
        Case "CURRENT_PERIOD_ANNUAL": sourceHeader = COL_CURRENT_PERIOD_ANNUAL
        Case "ANNUAL_7D": sourceHeader = COL_7DAY_ANNUAL
        Case "ANNUAL_28D": sourceHeader = COL_28DAY_ANNUAL
        Case "INCEPTION_ANNUAL": sourceHeader = COL_INCEPTION_ANNUAL
        Case "OPEN_FREQUENCY"
            If headers.Exists(COL_THEORETICAL_INTERVAL) Then
                targetCell.Value = SourceOpenFrequency(wsSource, headers, sourceRow)
            Else
                targetCell.Value = "\"
            End If
            Exit Sub
        Case Else
            Err.Raise vbObjectError + 4202, , "未知展示指标：" & fieldId
    End Select

    If Not headers.Exists(sourceHeader) Then Err.Raise vbObjectError + 4203, , "分类表现缺少字段：" & sourceHeader

    Dim value As Variant
    value = wsSource.Cells(sourceRow, CLng(headers(sourceHeader))).Value
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Or Len(NormalizeText(value)) = 0 Then
        targetCell.Value = "\"
    Else
        targetCell.Value = value
        If fieldId = "NEXT_OPEN" Then
            targetCell.NumberFormat = "yyyy-mm-dd"
        ElseIf IsPercentField(fieldId) Then
            targetCell.NumberFormat = "0.00%"
        End If
    End If
End Sub

Private Sub FormatGroupBlock(ByVal ws As Worksheet, ByVal themeName As String, ByVal headerRow As Long, _
                             ByVal lastRow As Long)
    Dim startColumn As Long
    startColumn = DisplayStartColumn(themeName)

    Dim block As Range
    Set block = ws.Range(ws.Cells(headerRow, startColumn), ws.Cells(lastRow, startColumn + 4))
    With block
        .RowHeight = 22
        .Font.Name = IIf(themeName = "蓝", FONT_BLUE, FONT_RED)
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    With ws.Range(ws.Cells(headerRow, startColumn), ws.Cells(headerRow, startColumn + 4))
        .Interior.Color = IIf(themeName = "蓝", COLOR_BLUE_TITLE, COLOR_DARK_RED)
        .Font.Bold = True
        .Font.Color = IIf(themeName = "蓝", COLOR_WHITE, COLOR_GOLD)
    End With

    If lastRow > headerRow Then
        Dim dataRange As Range
        Set dataRange = ws.Range(ws.Cells(headerRow + 1, startColumn), ws.Cells(lastRow, startColumn + 4))
        If themeName = "蓝" Then
            dataRange.Interior.Color = COLOR_BLUE_RECORD
            dataRange.Font.Color = COLOR_BLACK
        Else
            dataRange.Interior.Color = COLOR_DARK_RED
            dataRange.Font.Color = COLOR_WHITE
        End If
        ws.Range(ws.Cells(headerRow + 1, startColumn), ws.Cells(lastRow, startColumn)).HorizontalAlignment = xlLeft
    End If

    ApplyWhiteBorder block
End Sub

Private Sub ApplyRedFrequencyGroups(ByVal wsSource As Worksheet, ByVal sourceHeaders As Object, _
                                    ByRef sourceRows() As Long, ByVal sourceRowCount As Long, _
                                    ByVal wsTarget As Worksheet, ByVal firstDataRow As Long, _
                                    ByVal scheme As Object)
    Dim currentStep As String
    On Error GoTo ApplyFail
    If sourceRowCount = 0 Then Exit Sub
    If Not sourceHeaders.Exists(COL_THEORETICAL_INTERVAL) Then Exit Sub

    Dim frequencySlot As Long
    frequencySlot = FindSchemeFieldSlot(scheme, "OPEN_FREQUENCY")

    Dim groupStartIndex As Long
    Dim groupEndIndex As Long
    Dim groupIndex As Long
    groupStartIndex = 1

    Do While groupStartIndex <= sourceRowCount
        currentStep = "读取第" & CStr(groupStartIndex) & "条开放频率"
        Dim frequencyText As String
        frequencyText = NormalizeText(SourceOpenFrequency(wsSource, sourceHeaders, _
                                                          sourceRows(groupStartIndex)))
        groupEndIndex = groupStartIndex
        ' VBA的And不会短路；先检查数组边界，再读取下一项。
        Do While groupEndIndex < sourceRowCount
            If NormalizeText(SourceOpenFrequency(wsSource, sourceHeaders, _
                                                 sourceRows(groupEndIndex + 1))) <> frequencyText Then Exit Do
            groupEndIndex = groupEndIndex + 1
        Loop

        Dim targetGroupStart As Long
        Dim targetGroupEnd As Long
        targetGroupStart = firstDataRow + groupStartIndex - 1
        targetGroupEnd = firstDataRow + groupEndIndex - 1

        Dim fillColor As Long
        If groupIndex Mod 2 = 0 Then
            fillColor = COLOR_DARK_RED
        Else
            fillColor = COLOR_DARK_RED_ALT
        End If
        wsTarget.Range(wsTarget.Cells(targetGroupStart, REPORT_START_COLUMN), _
                       wsTarget.Cells(targetGroupEnd, REPORT_START_COLUMN + _
                                      REPORT_CONTENT_COLUMN_COUNT - 1)).Interior.Color = fillColor

        ' 只有字段方案实际展示开放频率时才合并该列。
        If frequencySlot > 0 Then
            Dim frequencyColumn As Long
            frequencyColumn = REPORT_START_COLUMN + frequencySlot - 1
            Dim frequencyRange As Range
            Set frequencyRange = wsTarget.Range(wsTarget.Cells(targetGroupStart, frequencyColumn), _
                                                wsTarget.Cells(targetGroupEnd, frequencyColumn))
            If targetGroupEnd > targetGroupStart Then
                Dim oldDisplayAlerts As Boolean
                oldDisplayAlerts = Application.DisplayAlerts
                Application.DisplayAlerts = False
                frequencyRange.Merge
                Application.DisplayAlerts = oldDisplayAlerts
            End If
            frequencyRange.HorizontalAlignment = xlCenter
            frequencyRange.VerticalAlignment = xlCenter
            ApplyWhiteBorder frequencyRange
        End If

        groupStartIndex = groupEndIndex + 1
        groupIndex = groupIndex + 1
    Loop
    Exit Sub

ApplyFail:
    Dim errorDescription As String
    errorDescription = Err.Description
    Err.Raise vbObjectError + 4210, , "按开放频率设置红色分组（" & currentStep & "）失败：" & errorDescription
End Sub

Private Function FindSchemeFieldSlot(ByVal scheme As Object, ByVal fieldId As String) As Long
    Dim slot As Long
    For slot = 1 To REPORT_CONTENT_COLUMN_COUNT
        If StrComp(RowText(scheme, "字段" & slot), fieldId, vbTextCompare) = 0 Then
            FindSchemeFieldSlot = slot
            Exit Function
        End If
    Next slot
End Function

Private Sub FinishCategoryPage(ByVal ws As Worksheet, ByVal themeName As String, ByVal nextRow As Long, _
                               ByVal baselineDate As Date)
    Dim startColumn As Long
    startColumn = DisplayStartColumn(themeName)

    Dim noteRange As Range
    Set noteRange = ws.Range(ws.Cells(nextRow, startColumn), ws.Cells(nextRow, startColumn + 4))
    noteRange.Merge
    noteRange.Value = "*表中数据来源于托管人复核的产品净值数据，数据截至" & _
                      Format$(baselineDate, "yyyy-mm-dd") & "，仅供参考，产品有风险，投资需谨慎"
    noteRange.HorizontalAlignment = xlCenter
    noteRange.VerticalAlignment = xlCenter
    noteRange.Font.Size = 12

    Dim riskRange As Range
    Set riskRange = ws.Range(ws.Cells(nextRow + 1, startColumn), ws.Cells(nextRow + 1, startColumn + 4))
    riskRange.Merge
    riskRange.Value = "风险提示:本产品由交银国际信托有限公司发行与管理，交通银行股份有限公司作为代销机构不承担产品的投资、兑付责任；" & vbLf & _
                      "*请您认真阅读信托合同、产品说明书、风险申明书等法律文件，根据风险承受能力选择合适的产品；" & vbLf & _
                      "*信托计划不承诺保证本金不受损失或最低收益，过往业绩并不预示其未来表现，产品发行人管理的其他产品的业绩并不构成未来产品业绩表现的保证；" & vbLf & _
                      "*下表中信托产品的代销机构风险评级为3R-平衡型，该类产品的风险中等，所投资金存在一定亏损风险，收益或利益浮动且有一定波动。"
    riskRange.HorizontalAlignment = xlLeft
    riskRange.VerticalAlignment = xlCenter
    riskRange.WrapText = True
    riskRange.Font.Size = 10

    If themeName = "蓝" Then
        noteRange.Interior.Color = COLOR_BLUE_TITLE
        noteRange.Font.Color = COLOR_WHITE
        riskRange.Interior.Color = COLOR_BLUE_TITLE
        riskRange.Font.Color = COLOR_WHITE
    Else
        noteRange.Interior.Color = COLOR_DARK_RED_ALT
        noteRange.Font.Color = COLOR_GOLD
        riskRange.Interior.Color = COLOR_DARK_RED
        riskRange.Font.Color = COLOR_GOLD
    End If

    ApplyWhiteBorder noteRange
    riskRange.Borders.LineStyle = xlNone
    ws.Rows(nextRow + 1).AutoFit
    If ws.Rows(nextRow + 1).RowHeight < 70 Then ws.Rows(nextRow + 1).RowHeight = 70

    ' 先确定最终B:F列宽，再按最终宽度等比例重排图片。
    ws.Range(ws.Cells(1, REPORT_START_COLUMN), _
             ws.Cells(nextRow + 1, REPORT_START_COLUMN + REPORT_CONTENT_COLUMN_COUNT - 1)).Columns.AutoFit
    ApplyDecorationColumns ws, themeName, nextRow + 1
    RefitReportPictures ws
End Sub

Private Sub CollectAndSortSourceRows(ByVal wsSource As Worksheet, ByVal headers As Object, _
                                     ByVal categoryName As String, ByVal groupName As String, _
                                     ByVal groupDefinitions As Collection, ByRef rowIndexes() As Long, _
                                     ByRef rowCount As Long)
    ' rowIndexes/rowCount由多个展示分组复用；每次调用必须从空结果开始。
    ' VBA循环块没有独立变量作用域，不显式清空会把前一分组继续带入下一分组。
    rowCount = 0
    Erase rowIndexes

    Dim allowedSeries As Object
    Set allowedSeries = CreateTextDictionary()

    Dim definition As Variant
    For Each definition In groupDefinitions
        If StrComp(RowText(definition, "分类"), categoryName, vbTextCompare) = 0 And _
           StrComp(RowText(definition, "展示分组"), groupName, vbTextCompare) = 0 Then
            allowedSeries(RowText(definition, "产品系列")) = True
        End If
    Next definition

    Dim lastRow As Long
    lastRow = LastUsedRow(wsSource)
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(NormalizeText(wsSource.Cells(r, CLng(headers(COL_CATEGORY))).Value), categoryName, vbTextCompare) = 0 And _
           allowedSeries.Exists(NormalizeText(wsSource.Cells(r, CLng(headers(COL_SERIES))).Value)) Then
            rowCount = rowCount + 1
            ReDim Preserve rowIndexes(1 To rowCount)
            rowIndexes(rowCount) = r
        End If
    Next r

    Dim i As Long
    Dim j As Long
    For i = 1 To rowCount - 1
        For j = i + 1 To rowCount
            If CompareSourceRows(wsSource, headers, rowIndexes(i), rowIndexes(j)) > 0 Then
                Dim tempRow As Long
                tempRow = rowIndexes(i)
                rowIndexes(i) = rowIndexes(j)
                rowIndexes(j) = tempRow
            End If
        Next j
    Next i
End Sub

Private Function CompareSourceRows(ByVal ws As Worksheet, ByVal headers As Object, _
                                   ByVal leftRow As Long, ByVal rightRow As Long) As Long
    Dim leftInterval As Double
    Dim rightInterval As Double
    leftInterval = NumericSortValue(ws.Cells(leftRow, CLng(headers(COL_THEORETICAL_INTERVAL))).Value)
    rightInterval = NumericSortValue(ws.Cells(rightRow, CLng(headers(COL_THEORETICAL_INTERVAL))).Value)

    If leftInterval < rightInterval Then
        CompareSourceRows = -1
    ElseIf leftInterval > rightInterval Then
        CompareSourceRows = 1
    Else
        Dim leftSeq As Double
        Dim rightSeq As Double
        leftSeq = NumericSortValue(ws.Cells(leftRow, CLng(headers(COL_SEQ))).Value)
        rightSeq = NumericSortValue(ws.Cells(rightRow, CLng(headers(COL_SEQ))).Value)
        CompareSourceRows = Sgn(leftSeq - rightSeq)
    End If
End Function

Private Function GetSortedEnabledCategories(ByVal definitions As Collection) As Collection
    Dim filtered As New Collection
    Dim definition As Variant
    For Each definition In definitions
        If IsEnabledValue(RowValue(definition, "是否启用")) Then filtered.Add definition
    Next definition
    Set GetSortedEnabledCategories = SortRowsByNumericField(filtered, "工作表顺序")
End Function

Private Function GetUniqueSortedGroups(ByVal definitions As Collection, ByVal categoryName As String) As Collection
    Dim result As New Collection
    Dim seen As Object
    Set seen = CreateTextDictionary()

    Dim definition As Variant
    For Each definition In definitions
        If StrComp(RowText(definition, "分类"), categoryName, vbTextCompare) = 0 Then
            Dim groupName As String
            groupName = RowText(definition, "展示分组")
            If Not seen.Exists(groupName) Then
                seen.Add groupName, True
                result.Add definition
            End If
        End If
    Next definition

    Set GetUniqueSortedGroups = SortRowsByNumericField(result, "分组顺序")
End Function

Private Function GetSortedRows(ByVal definitions As Collection, ByVal filterField As String, _
                               ByVal filterValue As String, ByVal sortField As String) As Collection
    Dim result As New Collection
    Dim definition As Variant
    For Each definition In definitions
        If StrComp(RowText(definition, filterField), filterValue, vbTextCompare) = 0 Then result.Add definition
    Next definition
    Set GetSortedRows = SortRowsByNumericField(result, sortField)
End Function

Private Function SortRowsByNumericField(ByVal rows As Collection, ByVal fieldName As String) As Collection
    Dim result As New Collection
    Dim used() As Boolean
    If rows.Count = 0 Then
        Set SortRowsByNumericField = result
        Exit Function
    End If
    ReDim used(1 To rows.Count)

    Dim outputIndex As Long
    Dim bestIndex As Long
    Dim bestValue As Double
    Dim i As Long
    For outputIndex = 1 To rows.Count
        bestIndex = 0
        bestValue = 0
        For i = 1 To rows.Count
            If Not used(i) Then
                Dim candidateValue As Double
                candidateValue = NumericSortValue(RowValue(rows(i), fieldName))
                If bestIndex = 0 Or candidateValue < bestValue Then
                    bestIndex = i
                    bestValue = candidateValue
                End If
            End If
        Next i
        used(bestIndex) = True
        result.Add rows(bestIndex)
    Next outputIndex

    Set SortRowsByNumericField = result
End Function

Private Function IndexRowsByKey(ByVal rows As Collection, ByVal keyField As String) As Object
    Dim result As Object
    Set result = CreateTextDictionary()
    Dim row As Variant
    For Each row In rows
        Set result.Item(RowText(row, keyField)) = row
    Next row
    Set IndexRowsByKey = result
End Function

Private Function BuildProductShortNameMap() As Object
    Dim result As Object
    Set result = CreateTextDictionary()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PRODUCT_INFO)
    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)
    If Not headers.Exists(COL_TRUST_CODE) Or Not headers.Exists(COL_PRODUCT_SHORT) Then
        Err.Raise vbObjectError + 4204, , SHEET_PRODUCT_INFO & "缺少信托计划代码或产品简称。"
    End If

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    Dim r As Long
    For r = 2 To lastRow
        Dim trustCode As String
        Dim shortName As String
        trustCode = NormalizeText(ws.Cells(r, CLng(headers(COL_TRUST_CODE))).Value)
        shortName = NormalizeText(ws.Cells(r, CLng(headers(COL_PRODUCT_SHORT))).Value)
        If Len(trustCode) > 0 And Len(shortName) > 0 Then result(trustCode) = shortName
    Next r

    Set BuildProductShortNameMap = result
End Function

Private Function SourceOpenFrequency(ByVal wsSource As Worksheet, ByVal headers As Object, _
                                     ByVal sourceRow As Long) As Variant
    Dim trustCode As Variant
    trustCode = Empty
    If headers.Exists(COL_TRUST_CODE) Then
        trustCode = wsSource.Cells(sourceRow, CLng(headers(COL_TRUST_CODE))).Value
    End If

    SourceOpenFrequency = FormatOpenFrequency( _
        wsSource.Cells(sourceRow, CLng(headers(COL_THEORETICAL_INTERVAL))).Value, trustCode)
End Function

Private Function FormatOpenFrequency(ByVal value As Variant, ByVal trustCode As Variant) As Variant
    Dim textValue As String
    textValue = NormalizeText(value)
    If Len(textValue) = 0 Then
        FormatOpenFrequency = "\"
    ElseIf textValue = "63" Then
        FormatOpenFrequency = "每2个月"
    ElseIf textValue = "154" Then
        FormatOpenFrequency = "每5个月"
    ElseIf textValue = "183" And NormalizeText(trustCode) = PRODUCT_CODE_6_MONTH_101 Then
        FormatOpenFrequency = "每月最后的周三"
    ElseIf textValue = "183" And NormalizeText(trustCode) = PRODUCT_CODE_6_MONTH_102 Then
        FormatOpenFrequency = "每月第二个周三"
    ElseIf IsNumeric(textValue) Then
        FormatOpenFrequency = "每" & textValue & "天"
    Else
        FormatOpenFrequency = value
    End If
End Function

Private Function IsPercentField(ByVal fieldId As String) As Boolean
    IsPercentField = (fieldId = "BENCHMARK_RATE" Or fieldId = "PREV_PERIOD_ANNUAL" Or _
                      fieldId = "CURRENT_PERIOD_ANNUAL" Or fieldId = "ANNUAL_7D" Or _
                      fieldId = "ANNUAL_28D" Or fieldId = "INCEPTION_ANNUAL")
End Function

Private Function DisplayStartColumn(ByVal themeName As String) As Long
    DisplayStartColumn = REPORT_START_COLUMN
End Function

Private Function DisplayRangeForRow(ByVal ws As Worksheet, ByVal themeName As String, _
                                    ByVal rowNumber As Long) As Range
    Dim startColumn As Long
    startColumn = DisplayStartColumn(themeName)
    Set DisplayRangeForRow = ws.Range(ws.Cells(rowNumber, startColumn), _
                                      ws.Cells(rowNumber, startColumn + REPORT_CONTENT_COLUMN_COUNT - 1))
End Function

Private Sub ApplyBaseThemeLayout(ByVal ws As Worksheet, ByVal themeName As String)
    ws.Columns("A").ColumnWidth = DECORATION_COLUMN_WIDTH
    ws.Columns("G").ColumnWidth = DECORATION_COLUMN_WIDTH
End Sub

Private Sub ApplyDecorationColumns(ByVal ws As Worksheet, ByVal themeName As String, ByVal lastRow As Long)
    ws.Columns("A").ColumnWidth = DECORATION_COLUMN_WIDTH
    ws.Columns("G").ColumnWidth = DECORATION_COLUMN_WIDTH
    ws.Range("A1:A" & lastRow).Interior.Color = ThemeMainColor(themeName)
    ws.Range("G1:G" & lastRow).Interior.Color = ThemeMainColor(themeName)
End Sub

Private Function ThemeMainColor(ByVal themeName As String) As Long
    If themeName = "蓝" Then
        ThemeMainColor = COLOR_BLUE_TITLE
    Else
        ThemeMainColor = COLOR_DARK_RED
    End If
End Function

Private Sub AddRedTitleImage(ByVal ws As Worksheet)
    Dim imagePath As String
    imagePath = AssetImagePath(TITLE_IMAGE_NAME)
    If Len(Dir(imagePath)) = 0 Then Exit Sub

    Dim targetRange As Range
    Set targetRange = ws.Range("B1:F2")
    Dim shape As Shape
    Set shape = ws.Shapes.AddPicture(imagePath, msoFalse, msoTrue, targetRange.Left, targetRange.Top, -1, -1)
    shape.Name = "report_title"
    FitPictureToRangeWidth shape, targetRange
    ws.Rows(1).RowHeight = shape.Height / 2
    ws.Rows(2).RowHeight = shape.Height / 2
End Sub

Private Sub InsertPictureFitWidthAndSetRowHeight(ByVal ws As Worksheet, ByVal imagePath As String, _
                                                  ByVal targetRange As Range, ByVal rowNumber As Long)
    Dim shape As Shape
    Set shape = ws.Shapes.AddPicture(imagePath, msoFalse, msoTrue, _
                                     targetRange.Left, targetRange.Top, -1, -1)
    shape.Name = "report_chart_" & CStr(rowNumber)
    FitPictureToRangeWidth shape, targetRange

    Dim newHeight As Double
    newHeight = shape.Height + 2
    If newHeight > 409.5 Then newHeight = 409.5
    ws.Rows(rowNumber).RowHeight = newHeight
End Sub

Private Sub RefitReportPictures(ByVal ws As Worksheet)
    Dim shape As Shape
    On Error Resume Next
    Set shape = ws.Shapes("report_title")
    On Error GoTo 0
    If Not shape Is Nothing Then
        FitPictureToRangeWidth shape, ws.Range("B1:F2")
        ws.Rows(1).RowHeight = shape.Height / 2
        ws.Rows(2).RowHeight = shape.Height / 2
    End If

    For Each shape In ws.Shapes
        If Left$(shape.Name, Len("report_chart_")) = "report_chart_" Then
            Dim rowNumber As Long
            rowNumber = CLng(Mid$(shape.Name, Len("report_chart_") + 1))
            FitPictureToRangeWidth shape, DisplayRangeForRow(ws, vbNullString, rowNumber)
            Dim newHeight As Double
            newHeight = shape.Height + 2
            If newHeight > 409.5 Then newHeight = 409.5
            ws.Rows(rowNumber).RowHeight = newHeight
        End If
    Next shape
End Sub

Private Sub FitPictureToRangeWidth(ByVal shape As Shape, ByVal targetRange As Range)
    shape.Placement = xlMove
    shape.LockAspectRatio = msoTrue
    shape.Width = targetRange.Width
    shape.Left = targetRange.Left
    shape.Top = targetRange.Top
End Sub

Private Function AssetImagePath(ByVal imageName As String) As String
    AssetImagePath = ThisWorkbook.Path & Application.PathSeparator & _
                     Replace(ASSET_IMAGE_FOLDER, "\", Application.PathSeparator) & _
                     Application.PathSeparator & imageName
End Function

Private Sub ApplyWhiteBorder(ByVal targetRange As Range)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Color = COLOR_WHITE
        .Weight = xlThin
    End With
End Sub

Private Function DisplayTextOrDefault(ByVal value As Variant, ByVal defaultText As String) As String
    Dim textValue As String
    textValue = NormalizeText(value)
    If Len(textValue) = 0 Then
        DisplayTextOrDefault = defaultText
    Else
        DisplayTextOrDefault = CStr(value)
    End If
End Function

Private Function RowValue(ByVal row As Object, ByVal fieldName As String) As Variant
    If row.Exists(fieldName) Then
        RowValue = row(fieldName)
    Else
        RowValue = Empty
    End If
End Function

Private Function RowText(ByVal row As Object, ByVal fieldName As String) As String
    RowText = NormalizeText(RowValue(row, fieldName))
End Function

Private Function NumericSortValue(ByVal value As Variant) As Double
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then
        NumericSortValue = 1E+30
    ElseIf Not IsNumeric(value) Then
        NumericSortValue = 1E+30
    Else
        NumericSortValue = CDbl(value)
    End If
End Function

Private Function IsEnabledValue(ByVal value As Variant) As Boolean
    Dim textValue As String
    textValue = UCase$(NormalizeText(value))
    IsEnabledValue = (textValue = "是" Or textValue = "Y" Or textValue = "YES" Or _
                      textValue = "1" Or textValue = "TRUE" Or textValue = "启用")
End Function

Private Function CreateTextDictionary() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set CreateTextDictionary = result
End Function

Private Function GetBaselineDateFromProductCategory() As Date
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PRODUCT_CATEGORY)
    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)
    If Not headers.Exists(COL_BASELINE_DATE) Then
        Err.Raise vbObjectError + 4205, , SHEET_PRODUCT_CATEGORY & "缺少字段：" & COL_BASELINE_DATE
    End If

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    Dim r As Long
    For r = 2 To lastRow
        Dim parsedDate As Date
        If TryReadDate(ws.Cells(r, CLng(headers(COL_BASELINE_DATE))).Value, parsedDate) Then
            GetBaselineDateFromProductCategory = parsedDate
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 4206, , SHEET_PRODUCT_CATEGORY & "没有有效基准日期。"
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
    Err.Raise vbObjectError + 4207, , "替换输出文件失败：" & originalError
End Sub

Private Sub RecordLastRunMessage(ByVal messageText As String)
    On Error Resume Next
    With ThisWorkbook.Worksheets("配置说明")
        .Range("A20").Value = "最近一次展示报表运行结果"
        .Range("A21").Value = messageText
    End With
    On Error GoTo 0
End Sub

Private Sub RestoreApplicationState(ByVal screenUpdating As Boolean, ByVal enableEvents As Boolean, _
                                    ByVal displayAlerts As Boolean, ByVal askToUpdateLinks As Boolean, _
                                    ByVal calculation As XlCalculation)
    Application.Calculation = calculation
    Application.AskToUpdateLinks = askToUpdateLinks
    Application.DisplayAlerts = displayAlerts
    Application.EnableEvents = enableEvents
    Application.ScreenUpdating = screenUpdating
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
