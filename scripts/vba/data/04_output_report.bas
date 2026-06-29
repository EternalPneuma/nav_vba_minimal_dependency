' Data04_ExportDisplayReport：将 yyyyMMdd-上层产品分类表现.xlsx 整理为 yyyyMMdd-展示.xlsx 展示报表模块

Option Explicit

Private Const SOURCE_FILE_SUFFIX As String = "-上层产品分类表现.xlsx"
Private Const OUTPUT_FILE_SUFFIX As String = "-产品收益展示.xlsx"
Private Const TITLE_IMAGE_NAME As String = "title.png"

Private Const PRODUCT_CATEGORY_SHEET_NAME As String = "产品分类"
Private Const CAT_STABLE As String = "稳享长期限"
Private Const CAT_DIRECT As String = "直销"
Private Const CAT_BANK As String = "交行代销"
Private Const CAT_YUANRONG_ANXIANG As String = "圆融安享"

Private Const COL_BASELINE_DATE As String = "基准日期"

Private Const STABLE_COL_SEQ As Long = 1
Private Const DIRECT_COL_SEQ As Long = 1
Private Const DIRECT_COL_SERIES As Long = 3
Private Const DIRECT_COL_PRODUCT_NAME As Long = 4
Private Const DIRECT_COL_NEXT_OPEN As Long = 9
Private Const DIRECT_COL_THEORETICAL_INTERVAL As Long = 10
Private Const DIRECT_COL_BENCHMARK_RATE As Long = 12
Private Const DIRECT_COL_7DAY_ANNUAL As Long = 13
Private Const DIRECT_COL_28DAY_ANNUAL As Long = 14
Private Const DIRECT_COL_INCEPTION_ANNUAL As Long = 15

Private Const BANK_COL_SEQ As Long = 1
Private Const BANK_COL_SERIES As Long = 3
Private Const BANK_COL_PRODUCT_NAME As Long = 4
Private Const BANK_COL_NEXT_OPEN As Long = 9
Private Const BANK_COL_THEORETICAL_INTERVAL As Long = 10
Private Const BANK_COL_BENCHMARK_RATE As Long = 12
Private Const BANK_COL_ELAPSED As Long = 15
Private Const BANK_COL_PREV_PERIOD_ANNUAL As Long = 16
Private Const BANK_COL_CURRENT_PERIOD_ANNUAL As Long = 17
Private Const BANK_COL_7DAY_ANNUAL As Long = 18
Private Const BANK_COL_28DAY_ANNUAL As Long = 19

Private Const STABLE_COL_PRODUCT_NAME As Long = 4
Private Const STABLE_COL_NEXT_OPEN As Long = 9
Private Const STABLE_COL_THEORETICAL_INTERVAL As Long = 10
Private Const STABLE_COL_ELAPSED As Long = 12
Private Const STABLE_COL_CURRENT_ANNUAL As Long = 13

Private Const YRA_COL_SEQ As Long = 1
Private Const YRA_COL_SERIES As Long = 3
Private Const YRA_COL_PRODUCT_NAME As Long = 4
Private Const YRA_COL_NEXT_OPEN As Long = 7
Private Const YRA_COL_THEORETICAL_INTERVAL As Long = 8
Private Const YRA_COL_7DAY_ANNUAL As Long = 10
Private Const YRA_COL_28DAY_ANNUAL As Long = 11

Private Const FONT_NAME As String = "微软雅黑"
Private Const FONT_NAME_YUANRONG As String = "宋体"
Private Const COLOR_DARK_RED As Long = &H19198B
Private Const COLOR_DARK_RED_ALT As Long = &HA1370
Private Const COLOR_GOLD As Long = &H8BDCF4
Private Const COLOR_WHITE As Long = &HFFFFFF
Private Const COLOR_YUANRONG_TITLE As Long = &H7C3702
Private Const COLOR_YUANRONG_RECORD As Long = &HF8F0EC
Private Const COLOR_BLACK As Long = &H0

Public Sub Data04_ExportDisplayReport()
    OutputDisplayReportCore
End Sub

Private Sub OutputDisplayReportCore()
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

    Dim wbSource As Workbook
    Dim wbOutput As Workbook

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual

    Dim baselineDate As Date
    baselineDate = GetBaselineDateFromProductCategory()

    Dim sourcePath As String
    sourcePath = ThisWorkbook.Path & Application.PathSeparator & Format$(baselineDate, "yyyymmdd") & SOURCE_FILE_SUFFIX
    If Dir(sourcePath) = vbNullString Then
        Err.Raise vbObjectError + 4101, , "未找到上一步输出文件：" & sourcePath
    End If

    Set wbSource = Workbooks.Open(FileName:=sourcePath, ReadOnly:=True, UpdateLinks:=False)
    Set wbOutput = CreateDisplayWorkbook()

    WriteStableLongTermReport wbSource.Worksheets(CAT_STABLE), wbOutput.Worksheets(CAT_STABLE), baselineDate
    WriteDirectSalesReport wbSource.Worksheets(CAT_DIRECT), wbOutput.Worksheets(CAT_DIRECT), baselineDate
    WriteBankAgentReport wbSource.Worksheets(CAT_BANK), wbOutput.Worksheets(CAT_BANK), baselineDate
    WriteYuanRongAnXiangReport wbSource.Worksheets(CAT_YUANRONG_ANXIANG), wbOutput.Worksheets(CAT_YUANRONG_ANXIANG), baselineDate

    AddTitleImageIfExists wbOutput.Worksheets(CAT_DIRECT)
    AddTitleImageIfExists wbOutput.Worksheets(CAT_BANK)
    AddTitleImageIfExists wbOutput.Worksheets(CAT_STABLE)

    Dim imageNotice As String
    imageNotice = InsertChartImagesIntoDisplay(wbOutput, ThisWorkbook.Path & Application.PathSeparator, Format$(baselineDate, "yyyymmdd"))

    Dim outputPath As String
    outputPath = ThisWorkbook.Path & Application.PathSeparator & Format$(baselineDate, "yyyymmdd") & OUTPUT_FILE_SUFFIX

    wbOutput.SaveAs FileName:=outputPath, FileFormat:=xlOpenXMLWorkbook
    wbOutput.Close SaveChanges:=False
    wbSource.Close SaveChanges:=False

    Application.Calculation = appCalc
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    Dim finalMsg As String
    finalMsg = "展示报表生成完成" & vbCrLf & vbCrLf & _
               "基准日期：" & Format$(baselineDate, "yyyy-mm-dd") & vbCrLf & vbCrLf & _
               "输出文件：" & vbCrLf & outputPath
    If Len(imageNotice) > 0 Then
        finalMsg = finalMsg & vbCrLf & vbCrLf & "注意事项：" & vbCrLf & imageNotice
    End If
    MsgBox finalMsg, vbInformation, "展示报表"
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not wbOutput Is Nothing Then wbOutput.Close SaveChanges:=False
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    On Error GoTo 0

    Application.Calculation = appCalc
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    MsgBox "展示报表生成失败" & vbCrLf & vbCrLf & _
           "错误信息：" & Err.Description, vbCritical, "展示报表"
End Sub

Private Function CreateDisplayWorkbook() As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = Workbooks.Add

    Application.DisplayAlerts = False
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True

    Set ws = wb.Worksheets(1)
    ws.Name = CAT_STABLE
    ws.Cells.Font.Name = FONT_NAME

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = CAT_DIRECT
    ws.Cells.Font.Name = FONT_NAME

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = CAT_BANK
    ws.Cells.Font.Name = FONT_NAME

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = CAT_YUANRONG_ANXIANG
    ws.Cells.Font.Name = FONT_NAME_YUANRONG

    Set CreateDisplayWorkbook = wb
End Function

Private Sub WriteStableLongTermReport(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet, ByVal baselineDate As Date)
    Dim lastSourceRow As Long
    Dim sourceRow As Long
    Dim targetRow As Long

    wsTarget.Cells.ClearFormats
    wsTarget.Cells.Font.Name = FONT_NAME

    With wsTarget.Range("B3:F3")
        .Merge
        .Value = "热销产品"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 24
        .Font.Color = COLOR_GOLD
    End With

    With wsTarget.Range("B6:F6")
        .Merge
        .Value = "汇益稳享系列兑付业绩"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 24
        .Font.Color = COLOR_GOLD
    End With

    wsTarget.Range("B7").Value = "产品名称"
    wsTarget.Range("C7").Value = "下一开放日"
    wsTarget.Range("D7").Value = "开放频率"
    wsTarget.Range("E7").Value = "当前周期运作天数"
    wsTarget.Range("F7").Value = "当前周期年化"

    lastSourceRow = LastUsedRow(wsSource)
    targetRow = 8

    For sourceRow = 2 To lastSourceRow
        If Len(NormalizeText(wsSource.Cells(sourceRow, STABLE_COL_PRODUCT_NAME).Value)) > 0 Then
            wsTarget.Cells(targetRow, 1).Value = BuildSortKey(wsSource.Cells(sourceRow, STABLE_COL_SEQ).Value)
            wsTarget.Cells(targetRow, 2).Value = wsSource.Cells(sourceRow, STABLE_COL_PRODUCT_NAME).Value
            wsTarget.Cells(targetRow, 3).Value = wsSource.Cells(sourceRow, STABLE_COL_NEXT_OPEN).Value
            wsTarget.Cells(targetRow, 4).Value = FormatOpenFrequency(wsSource.Cells(sourceRow, STABLE_COL_THEORETICAL_INTERVAL).Value)
            wsTarget.Cells(targetRow, 5).Value = wsSource.Cells(sourceRow, STABLE_COL_ELAPSED).Value
            wsTarget.Cells(targetRow, 6).Value = wsSource.Cells(sourceRow, STABLE_COL_CURRENT_ANNUAL).Value
            wsTarget.Cells(targetRow, 7).Value = BuildSortKey(wsSource.Cells(sourceRow, STABLE_COL_THEORETICAL_INTERVAL).Value)

            wsTarget.Cells(targetRow, 3).NumberFormat = "yyyy-mm-dd"
            wsTarget.Cells(targetRow, 6).NumberFormat = "0.00%"
            targetRow = targetRow + 1
        End If
    Next sourceRow

    Dim lastDataRow As Long
    Dim noteRow As Long
    Dim riskRow As Long
    Dim fillEndRow As Long

    lastDataRow = targetRow - 1
    If lastDataRow < 7 Then lastDataRow = 7
    If lastDataRow >= 8 Then
        SortStableLongTermRows wsTarget, lastDataRow
        wsTarget.Range("A8:A" & lastDataRow).ClearContents
        wsTarget.Range("G8:G" & lastDataRow).ClearContents
    End If

    noteRow = lastDataRow + 1
    riskRow = noteRow + 1
    fillEndRow = riskRow + 1

    With wsTarget.Range("A3:G" & fillEndRow)
        .Interior.Color = COLOR_DARK_RED
    End With

    ApplyFrequencyGroupFillAndMerge wsTarget, 8, lastDataRow

    With wsTarget.Range("B7:F" & lastDataRow)
        .RowHeight = 20
        .Font.Name = FONT_NAME
        .Font.Size = 14
        .Font.Color = COLOR_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    If lastDataRow >= 7 Then
        wsTarget.Range("B7:B" & lastDataRow).HorizontalAlignment = xlLeft
    End If

    With wsTarget.Range("B7:F" & lastDataRow).Borders
        .LineStyle = xlContinuous
        .Color = COLOR_WHITE
        .Weight = xlThin
    End With

    With wsTarget.Range("B7:F7")
        .Interior.Color = COLOR_DARK_RED
        .Font.Bold = True
        .Font.Color = COLOR_GOLD
    End With

    With wsTarget.Range("B" & noteRow & ":F" & noteRow)
        .Merge
        .Value = "*表中数据来源于托管人复核的产品净值数据，数据截至" & Format$(baselineDate, "yyyy-mm-dd") & "，仅供参考，产品有风险，投资需谨慎"
        .Interior.Color = COLOR_DARK_RED_ALT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 12
        .Font.Color = COLOR_GOLD
    End With
    ApplyWhiteBorder wsTarget.Range("B" & noteRow & ":F" & noteRow)

    With wsTarget.Range("B" & riskRow & ":F" & riskRow)
        .Merge
        .Value = "风险提示:本产品由交银国际信托有限公司发行与管理，交通银行股份有限公司作为代销机构不承担产品的投资、兑付责任；" & vbLf & _
                 "*请您认真阅读信托合同、产品说明书、风险申明书等法律文件，根据风险承受能力选择合适的产品；" & vbLf & _
                 "*信托计划不承诺保证本金不受损失或最低收益，过往业绩并不预示其未来表现，产品发行人管理的其他产品的业绩并不构成未来产品业绩表现的保证；" & vbLf & _
                 "*下表中信托产品的代销机构风险评级为3R-平衡型，该类产品的风险中等，所投资金存在一定亏损风险，收益或利益浮动且有一定波动。"
        .Interior.Color = COLOR_DARK_RED
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 10
        .Font.Color = COLOR_GOLD
    End With
    wsTarget.Rows(riskRow).AutoFit
    If wsTarget.Rows(riskRow).RowHeight < 70 Then wsTarget.Rows(riskRow).RowHeight = 70

    wsTarget.Columns("A").ColumnWidth = 3
    wsTarget.Columns("G").ColumnWidth = 3
    wsTarget.Columns("B:F").AutoFit
End Sub

Private Sub SortStableLongTermRows(ByVal ws As Worksheet, ByVal lastDataRow As Long)
    If lastDataRow < 8 Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("G8:G" & lastDataRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range("A8:A" & lastDataRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A8:G" & lastDataRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Function BuildSortKey(ByVal value As Variant) As Variant
    If IsError(value) Or IsEmpty(value) Then
        BuildSortKey = vbNullString
        Exit Function
    End If

    Dim textValue As String
    textValue = NormalizeText(value)
    If Len(textValue) = 0 Then
        BuildSortKey = vbNullString
    ElseIf IsNumeric(textValue) Then
        BuildSortKey = CDbl(textValue)
    Else
        BuildSortKey = textValue
    End If
End Function

Private Function FormatOpenFrequency(ByVal value As Variant) As Variant
    If IsError(value) Or IsEmpty(value) Then
        FormatOpenFrequency = value
        Exit Function
    End If

    Dim textValue As String
    textValue = NormalizeText(value)
    If Len(textValue) = 0 Then
        FormatOpenFrequency = value
    ElseIf textValue = "63" Then
        FormatOpenFrequency = "每2个月"
    ElseIf textValue = "154" Then
        FormatOpenFrequency = "每5个月"
    ElseIf IsNumeric(textValue) Then
        FormatOpenFrequency = "每" & textValue & "天"
    Else
        FormatOpenFrequency = value
    End If
End Function

Private Sub ApplyWhiteBorder(ByVal targetRange As Range)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Color = COLOR_WHITE
        .Weight = xlThin
    End With
End Sub

Private Sub WriteDirectSalesReport(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet, ByVal baselineDate As Date)
    Dim currentRow As Long
    Dim noteRow As Long
    Dim riskRow As Long
    Dim fillEndRow As Long

    wsTarget.Cells.ClearFormats
    wsTarget.Cells.Font.Name = FONT_NAME

    With wsTarget.Range("B3:F3")
        .Merge
        .Value = "热销产品"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 24
        .Font.Color = COLOR_GOLD
    End With

    currentRow = 7
    currentRow = WriteDirectSection(wsSource, wsTarget, currentRow, "汇益稳健系列兑付业绩", "稳健", 1, baselineDate)
    currentRow = WriteDirectSection(wsSource, wsTarget, currentRow, "汇益系列兑付业绩", "汇益", 2, baselineDate)
    currentRow = WriteDirectSection(wsSource, wsTarget, currentRow, "交鑫致远系列兑付业绩", "交鑫致远", 3, baselineDate)

    noteRow = currentRow
    riskRow = noteRow + 1
    fillEndRow = riskRow + 1

    wsTarget.Range("A3:A" & fillEndRow).Interior.Color = COLOR_DARK_RED
    wsTarget.Range("G3:G" & fillEndRow).Interior.Color = COLOR_DARK_RED
    wsTarget.Range("B3:F6").Interior.Color = COLOR_DARK_RED
    wsTarget.Range("B" & fillEndRow & ":F" & fillEndRow).Interior.Color = COLOR_DARK_RED

    With wsTarget.Range("B" & noteRow & ":F" & noteRow)
        .Merge
        .Value = "*表中数据来源于托管人复核的产品净值数据，数据截至" & Format$(baselineDate, "yyyy-mm-dd") & "，仅供参考，产品有风险，投资需谨慎"
        .Interior.Color = COLOR_DARK_RED_ALT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 12
        .Font.Color = COLOR_GOLD
    End With
    ApplyWhiteBorder wsTarget.Range("B" & noteRow & ":F" & noteRow)

    With wsTarget.Range("B" & riskRow & ":F" & riskRow)
        .Merge
        .Value = "风险提示:本产品由交银国际信托有限公司发行与管理，交通银行股份有限公司作为代销机构不承担产品的投资、兑付责任；" & vbLf & _
                 "*请您认真阅读信托合同、产品说明书、风险申明书等法律文件，根据风险承受能力选择合适的产品；" & vbLf & _
                 "*信托计划不承诺保证本金不受损失或最低收益，过往业绩并不预示其未来表现，产品发行人管理的其他产品的业绩并不构成未来产品业绩表现的保证；" & vbLf & _
                 "*下表中信托产品的代销机构风险评级为3R-平衡型，该类产品的风险中等，所投资金存在一定亏损风险，收益或利益浮动且有一定波动。"
        .Interior.Color = COLOR_DARK_RED
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 10
        .Font.Color = COLOR_GOLD
    End With
    wsTarget.Rows(riskRow).AutoFit
    If wsTarget.Rows(riskRow).RowHeight < 70 Then wsTarget.Rows(riskRow).RowHeight = 70

    wsTarget.Columns("A").ColumnWidth = 3
    wsTarget.Columns("G").ColumnWidth = 3
    wsTarget.Columns("B:F").AutoFit
End Sub

Private Function WriteDirectSection(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet, _
                                    ByVal titleRow As Long, ByVal titleText As String, _
                                    ByVal seriesKeyword As String, ByVal sectionType As Long, _
                                    ByVal baselineDate As Date) As Long
    Dim headerRow As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim nextRow As Long
    Dim sourceRow As Long
    Dim lastSourceRow As Long
    Dim intervalValue As Variant

    With wsTarget.Range("B" & titleRow & ":F" & titleRow)
        .Merge
        .Value = titleText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 24
        .Font.Color = COLOR_GOLD
        .Interior.Color = COLOR_DARK_RED
    End With

    headerRow = titleRow + 1
    firstDataRow = headerRow + 1
    nextRow = firstDataRow

    WriteDirectSectionHeaders wsTarget, headerRow, sectionType

    lastSourceRow = LastUsedRow(wsSource)
    For sourceRow = 2 To lastSourceRow
        If IsDirectSeriesMatch(wsSource.Cells(sourceRow, DIRECT_COL_SERIES).Value, seriesKeyword, sectionType) Then
            wsTarget.Cells(nextRow, 1).Value = BuildSortKey(wsSource.Cells(sourceRow, DIRECT_COL_SEQ).Value)
            wsTarget.Cells(nextRow, 2).Value = wsSource.Cells(sourceRow, DIRECT_COL_PRODUCT_NAME).Value

            intervalValue = wsSource.Cells(sourceRow, DIRECT_COL_THEORETICAL_INTERVAL).Value
            WriteNextOpenDateCell wsTarget.Cells(nextRow, 3), wsSource.Cells(sourceRow, DIRECT_COL_NEXT_OPEN).Value, intervalValue, baselineDate
            If sectionType = 3 Then
                wsTarget.Cells(nextRow, 4).Value = "每月最后一个周三的工作日"
                wsTarget.Cells(nextRow, 7).Value = 9999
            Else
                wsTarget.Cells(nextRow, 4).Value = FormatOpenFrequency(intervalValue)
                wsTarget.Cells(nextRow, 7).Value = BuildSortKey(intervalValue)
            End If

            If sectionType = 1 Then
                wsTarget.Cells(nextRow, 5).Value = wsSource.Cells(sourceRow, DIRECT_COL_7DAY_ANNUAL).Value
                wsTarget.Cells(nextRow, 6).Value = wsSource.Cells(sourceRow, DIRECT_COL_28DAY_ANNUAL).Value
            ElseIf sectionType = 2 Then
                wsTarget.Cells(nextRow, 5).Value = wsSource.Cells(sourceRow, DIRECT_COL_BENCHMARK_RATE).Value
                wsTarget.Cells(nextRow, 6).Value = PickDirectCycleAnnual(wsSource, sourceRow, intervalValue)
            Else
                wsTarget.Cells(nextRow, 6).Value = wsSource.Cells(sourceRow, DIRECT_COL_INCEPTION_ANNUAL).Value
            End If

            wsTarget.Range("E" & nextRow & ":F" & nextRow).NumberFormat = "0.00%"
            nextRow = nextRow + 1
        End If
    Next sourceRow

    lastDataRow = nextRow - 1
    If lastDataRow >= firstDataRow Then
        SortDirectSectionRows wsTarget, firstDataRow, lastDataRow
        wsTarget.Range("A" & firstDataRow & ":A" & lastDataRow).ClearContents
        wsTarget.Range("G" & firstDataRow & ":G" & lastDataRow).ClearContents
        ApplyDirectSectionFrequencyFillAndMerge wsTarget, firstDataRow, lastDataRow, sectionType
    End If

    FormatDirectSectionBlock wsTarget, headerRow, IIf(lastDataRow >= firstDataRow, lastDataRow, headerRow), sectionType
    WriteDirectSection = IIf(lastDataRow >= firstDataRow, lastDataRow + 1, headerRow + 1)
End Function

Private Sub WriteDirectSectionHeaders(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal sectionType As Long)
    ws.Range("B" & headerRow).Value = IIf(sectionType = 1, "净值型产品名称", IIf(sectionType = 2, "报价型产品名称", "长期限产品名称"))
    ws.Range("C" & headerRow).Value = "下一开放日"
    ws.Range("D" & headerRow).Value = "开放频率"

    If sectionType = 1 Then
        ws.Range("E" & headerRow).Value = "7日年化收益率"
        ws.Range("F" & headerRow).Value = "28日年化收益率"
    ElseIf sectionType = 2 Then
        ws.Range("E" & headerRow).Value = "基准"
        ws.Range("F" & headerRow).Value = "7日/28日年化收益率"
    Else
        With ws.Range("D" & headerRow & ":E" & headerRow)
            .Merge
            .Value = "开放频率"
        End With
        ws.Range("F" & headerRow).Value = "成立以来年化收益率"
    End If
End Sub

Private Sub FormatDirectSectionBlock(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal sectionType As Long)
    With ws.Range("B" & headerRow & ":F" & lastRow)
        .RowHeight = 20
        .Font.Name = FONT_NAME
        .Font.Size = 14
        .Font.Color = COLOR_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    If lastRow >= headerRow Then
        ws.Range("B" & headerRow & ":B" & lastRow).HorizontalAlignment = xlLeft
    End If

    With ws.Range("B" & headerRow & ":F" & lastRow).Borders
        .LineStyle = xlContinuous
        .Color = COLOR_WHITE
        .Weight = xlThin
    End With

    With ws.Range("B" & headerRow & ":F" & headerRow)
        .Interior.Color = COLOR_DARK_RED
        .Font.Bold = True
        .Font.Color = COLOR_GOLD
    End With

    If sectionType = 3 Then
        With ws.Range("D" & headerRow & ":E" & headerRow)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
End Sub

Private Sub SortDirectSectionRows(ByVal ws As Worksheet, ByVal firstDataRow As Long, ByVal lastDataRow As Long)
    If lastDataRow < firstDataRow Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("G" & firstDataRow & ":G" & lastDataRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range("A" & firstDataRow & ":A" & lastDataRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A" & firstDataRow & ":G" & lastDataRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Sub ApplyDirectSectionFrequencyFillAndMerge(ByVal ws As Worksheet, ByVal firstDataRow As Long, ByVal lastDataRow As Long, ByVal sectionType As Long)
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim frequencyText As String
    Dim groupIndex As Long
    Dim fillColor As Long
    Dim oldDisplayAlerts As Boolean

    If lastDataRow < firstDataRow Then Exit Sub

    groupStart = firstDataRow
    groupIndex = 0

    Do While groupStart <= lastDataRow
        frequencyText = NormalizeText(ws.Cells(groupStart, 4).Value)
        groupEnd = groupStart

        Do While groupEnd + 1 <= lastDataRow And NormalizeText(ws.Cells(groupEnd + 1, 4).Value) = frequencyText
            groupEnd = groupEnd + 1
        Loop

        If groupIndex Mod 2 = 0 Then
            fillColor = COLOR_DARK_RED
        Else
            fillColor = COLOR_DARK_RED_ALT
        End If

        ws.Range("B" & groupStart & ":F" & groupEnd).Interior.Color = fillColor

        oldDisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        If sectionType = 3 Then
            ws.Range("D" & groupStart & ":E" & groupEnd).Merge
        ElseIf groupEnd > groupStart Then
            ws.Range("D" & groupStart & ":D" & groupEnd).Merge
        End If
        Application.DisplayAlerts = oldDisplayAlerts

        With ws.Range("D" & groupStart & ":D" & groupEnd)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        If sectionType = 3 Then
            With ws.Range("D" & groupStart & ":E" & groupEnd)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If

        groupStart = groupEnd + 1
        groupIndex = groupIndex + 1
    Loop
End Sub

Private Function IsDirectSeriesMatch(ByVal seriesValue As Variant, ByVal seriesKeyword As String, ByVal sectionType As Long) As Boolean
    Dim seriesText As String
    seriesText = NormalizeText(seriesValue)

    If Len(seriesText) = 0 Then Exit Function

    If sectionType = 1 Then
        IsDirectSeriesMatch = (InStr(1, seriesText, "稳健", vbTextCompare) > 0)
    ElseIf sectionType = 2 Then
        IsDirectSeriesMatch = (InStr(1, seriesText, "汇益", vbTextCompare) > 0 And InStr(1, seriesText, "稳健", vbTextCompare) = 0)
    Else
        IsDirectSeriesMatch = (InStr(1, seriesText, "交鑫致远", vbTextCompare) > 0)
    End If
End Function

Private Function PickDirectCycleAnnual(ByVal wsSource As Worksheet, ByVal sourceRow As Long, ByVal intervalValue As Variant) As Variant
    Dim intervalText As String
    intervalText = NormalizeText(intervalValue)

    If IsWeeklyInterval(intervalValue) Then
        PickDirectCycleAnnual = wsSource.Cells(sourceRow, DIRECT_COL_7DAY_ANNUAL).Value
    ElseIf intervalText = "28" Or InStr(1, intervalText, "四周", vbTextCompare) > 0 Then
        PickDirectCycleAnnual = wsSource.Cells(sourceRow, DIRECT_COL_28DAY_ANNUAL).Value
    Else
        PickDirectCycleAnnual = vbNullString
    End If
End Function

Private Sub WriteOpenDateCell(ByVal targetCell As Range, ByVal dateValue As Variant)
    Dim parsedDate As Date

    If TryReadDate(dateValue, parsedDate) Then
        targetCell.Value = parsedDate
        targetCell.NumberFormat = "yyyy-mm-dd"
    Else
        targetCell.Value = "\"
    End If
End Sub

Private Sub WriteNextOpenDateCell(ByVal targetCell As Range, ByVal sourceDateValue As Variant, _
                                  ByVal intervalValue As Variant, ByVal baselineDate As Date)
    If IsWeeklyInterval(intervalValue) Then
        targetCell.Value = GetNextWednesday(baselineDate)
        targetCell.NumberFormat = "yyyy-mm-dd"
    Else
        WriteOpenDateCell targetCell, sourceDateValue
    End If
End Sub

Private Function IsWeeklyInterval(ByVal intervalValue As Variant) As Boolean
    Dim intervalText As String
    intervalText = NormalizeText(intervalValue)
    IsWeeklyInterval = (intervalText = "7" Or intervalText = "7天" Or intervalText = "7日" Or _
                        intervalText = "每周" Or intervalText = "周开" Or _
                        InStr(1, intervalText, "周频", vbTextCompare) > 0)
End Function

Private Function GetNextWednesday(ByVal baselineDate As Date) As Date
    Dim daysToAdd As Long
    daysToAdd = 3 - Weekday(baselineDate, vbMonday)
    If daysToAdd <= 0 Then daysToAdd = daysToAdd + 7
    GetNextWednesday = DateAdd("d", daysToAdd, DateValue(baselineDate))
End Function

Private Sub WriteBankAgentReport(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet, ByVal baselineDate As Date)
    Dim currentRow As Long
    Dim noteRow As Long
    Dim riskRow As Long
    Dim fillEndRow As Long

    wsTarget.Cells.ClearFormats
    wsTarget.Cells.Font.Name = FONT_NAME

    With wsTarget.Range("B3:F3")
        .Merge
        .Value = "热销产品"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 24
        .Font.Color = COLOR_GOLD
    End With

    currentRow = 7
    currentRow = WriteBankSection(wsSource, wsTarget, currentRow, "汇益稳健系列业绩", "稳健", 1, baselineDate)
    currentRow = WriteBankSection(wsSource, wsTarget, currentRow, "汇益系列业绩", "汇益", 2, baselineDate)
    currentRow = WriteBankSection(wsSource, wsTarget, currentRow, "汇益稳享系列业绩", "稳享", 3, baselineDate)
    currentRow = WriteBankSection(wsSource, wsTarget, currentRow, "蓝色港湾系列业绩", "蓝色港湾", 4, baselineDate)

    noteRow = currentRow
    riskRow = noteRow + 1
    fillEndRow = riskRow + 1

    wsTarget.Range("A3:A" & fillEndRow).Interior.Color = COLOR_DARK_RED
    wsTarget.Range("G3:G" & fillEndRow).Interior.Color = COLOR_DARK_RED
    wsTarget.Range("B3:F6").Interior.Color = COLOR_DARK_RED
    wsTarget.Range("B" & fillEndRow & ":F" & fillEndRow).Interior.Color = COLOR_DARK_RED

    With wsTarget.Range("B" & noteRow & ":F" & noteRow)
        .Merge
        .Value = "*表中数据来源于托管人复核的产品净值数据，数据截至" & Format$(baselineDate, "yyyy-mm-dd") & "，仅供参考，产品有风险，投资需谨慎"
        .Interior.Color = COLOR_DARK_RED_ALT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 12
        .Font.Color = COLOR_GOLD
    End With
    ApplyWhiteBorder wsTarget.Range("B" & noteRow & ":F" & noteRow)

    With wsTarget.Range("B" & riskRow & ":F" & riskRow)
        .Merge
        .Value = "风险提示:本产品由交银国际信托有限公司发行与管理，交通银行股份有限公司作为代销机构不承担产品的投资、兑付责任；" & vbLf & _
                 "*请您认真阅读信托合同、产品说明书、风险申明书等法律文件，根据风险承受能力选择合适的产品；" & vbLf & _
                 "*信托计划不承诺保证本金不受损失或最低收益，过往业绩并不预示其未来表现，产品发行人管理的其他产品的业绩并不构成未来产品业绩表现的保证；" & vbLf & _
                 "*下表中信托产品的代销机构风险评级为3R-平衡型，该类产品的风险中等，所投资金存在一定亏损风险，收益或利益浮动且有一定波动。"
        .Interior.Color = COLOR_DARK_RED
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 10
        .Font.Color = COLOR_GOLD
    End With
    wsTarget.Rows(riskRow).AutoFit
    If wsTarget.Rows(riskRow).RowHeight < 70 Then wsTarget.Rows(riskRow).RowHeight = 70

    wsTarget.Columns("A").ColumnWidth = 3
    wsTarget.Columns("G").ColumnWidth = 3
    wsTarget.Columns("B:F").AutoFit
End Sub

Private Function WriteBankSection(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet, _
                                  ByVal titleRow As Long, ByVal titleText As String, _
                                  ByVal seriesKeyword As String, ByVal sectionType As Long, _
                                  ByVal baselineDate As Date) As Long
    Dim headerRow As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim nextRow As Long
    Dim sourceRow As Long
    Dim lastSourceRow As Long
    Dim intervalValue As Variant

    With wsTarget.Range("B" & titleRow & ":F" & titleRow)
        .Merge
        .Value = titleText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 24
        .Font.Color = COLOR_GOLD
        .Interior.Color = COLOR_DARK_RED
    End With

    headerRow = titleRow + 1
    firstDataRow = headerRow + 1
    nextRow = firstDataRow

    WriteBankSectionHeaders wsTarget, headerRow, sectionType

    lastSourceRow = LastUsedRow(wsSource)
    For sourceRow = 2 To lastSourceRow
        If IsBankSeriesMatch(wsSource.Cells(sourceRow, BANK_COL_SERIES).Value, sectionType) Then
            intervalValue = wsSource.Cells(sourceRow, BANK_COL_THEORETICAL_INTERVAL).Value

            wsTarget.Cells(nextRow, 1).Value = BuildSortKey(wsSource.Cells(sourceRow, BANK_COL_SEQ).Value)
            wsTarget.Cells(nextRow, 2).Value = wsSource.Cells(sourceRow, BANK_COL_PRODUCT_NAME).Value
            WriteNextOpenDateCell wsTarget.Cells(nextRow, 3), wsSource.Cells(sourceRow, BANK_COL_NEXT_OPEN).Value, intervalValue, baselineDate
            wsTarget.Cells(nextRow, 7).Value = BuildSortKey(intervalValue)

            If sectionType = 1 Then
                wsTarget.Cells(nextRow, 4).Value = FormatOpenFrequency(intervalValue)
                wsTarget.Cells(nextRow, 5).Value = wsSource.Cells(sourceRow, BANK_COL_7DAY_ANNUAL).Value
                wsTarget.Cells(nextRow, 6).Value = wsSource.Cells(sourceRow, BANK_COL_28DAY_ANNUAL).Value
                wsTarget.Range("E" & nextRow & ":F" & nextRow).NumberFormat = "0.00%"
            ElseIf sectionType = 2 Then
                wsTarget.Cells(nextRow, 4).Value = FormatOpenFrequency(intervalValue)
                WriteValueOrSlash wsTarget.Cells(nextRow, 5), wsSource.Cells(sourceRow, BANK_COL_BENCHMARK_RATE).Value
                wsTarget.Cells(nextRow, 6).Value = PickBankCycleAnnual(wsSource, sourceRow, intervalValue)
                wsTarget.Range("E" & nextRow & ":F" & nextRow).NumberFormat = "0.00%"
            ElseIf sectionType = 3 Then
                WriteValueOrSlash wsTarget.Cells(nextRow, 4), wsSource.Cells(sourceRow, BANK_COL_PREV_PERIOD_ANNUAL).Value
                wsTarget.Cells(nextRow, 5).Value = wsSource.Cells(sourceRow, BANK_COL_ELAPSED).Value
                wsTarget.Cells(nextRow, 6).Value = wsSource.Cells(sourceRow, BANK_COL_CURRENT_PERIOD_ANNUAL).Value
                wsTarget.Range("D" & nextRow).NumberFormat = "0.00%"
                wsTarget.Range("F" & nextRow).NumberFormat = "0.00%"
            Else
                wsTarget.Cells(nextRow, 4).Value = FormatOpenFrequency(intervalValue)
                wsTarget.Cells(nextRow, 5).Value = wsSource.Cells(sourceRow, BANK_COL_ELAPSED).Value
                wsTarget.Cells(nextRow, 6).Value = wsSource.Cells(sourceRow, BANK_COL_CURRENT_PERIOD_ANNUAL).Value
                wsTarget.Range("F" & nextRow).NumberFormat = "0.00%"
            End If

            nextRow = nextRow + 1
        End If
    Next sourceRow

    lastDataRow = nextRow - 1
    If lastDataRow >= firstDataRow Then
        SortBankSectionRows wsTarget, firstDataRow, lastDataRow
        ApplyBankSectionFrequencyFill wsTarget, firstDataRow, lastDataRow, sectionType
        wsTarget.Range("A" & firstDataRow & ":A" & lastDataRow).ClearContents
        wsTarget.Range("G" & firstDataRow & ":G" & lastDataRow).ClearContents
    End If

    FormatBankSectionBlock wsTarget, headerRow, IIf(lastDataRow >= firstDataRow, lastDataRow, headerRow)
    WriteBankSection = IIf(lastDataRow >= firstDataRow, lastDataRow + 1, headerRow + 1)
End Function

Private Sub WriteBankSectionHeaders(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal sectionType As Long)
    If sectionType = 1 Then
        ws.Range("B" & headerRow).Value = "净值型产品名称"
        ws.Range("C" & headerRow).Value = "下一开放日"
        ws.Range("D" & headerRow).Value = "开放频率"
        ws.Range("E" & headerRow).Value = "7日年化收益率"
        ws.Range("F" & headerRow).Value = "28日年化收益率"
    ElseIf sectionType = 2 Then
        ws.Range("B" & headerRow).Value = "报价型产品名称"
        ws.Range("C" & headerRow).Value = "下一开放日"
        ws.Range("D" & headerRow).Value = "开放频率"
        ws.Range("E" & headerRow).Value = "基准"
        ws.Range("F" & headerRow).Value = "7日/28日年化收益率"
    ElseIf sectionType = 3 Then
        ws.Range("B" & headerRow).Value = "长期限产品名称"
        ws.Range("C" & headerRow).Value = "下一开放日"
        ws.Range("D" & headerRow).Value = "上期年化收益率"
        ws.Range("E" & headerRow).Value = "当前周期运作天数"
        ws.Range("F" & headerRow).Value = "当期年化收益率"
    Else
        ws.Range("B" & headerRow).Value = "产品名称"
        ws.Range("C" & headerRow).Value = "下一开放日"
        ws.Range("D" & headerRow).Value = "开放频率"
        ws.Range("E" & headerRow).Value = "当前周期运作天数"
        ws.Range("F" & headerRow).Value = "当期年化收益率"
    End If
End Sub

Private Sub FormatBankSectionBlock(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long)
    With ws.Range("B" & headerRow & ":F" & lastRow)
        .RowHeight = 20
        .Font.Name = FONT_NAME
        .Font.Size = 14
        .Font.Color = COLOR_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    If lastRow >= headerRow Then
        ws.Range("B" & headerRow & ":B" & lastRow).HorizontalAlignment = xlLeft
    End If

    With ws.Range("B" & headerRow & ":F" & lastRow).Borders
        .LineStyle = xlContinuous
        .Color = COLOR_WHITE
        .Weight = xlThin
    End With

    With ws.Range("B" & headerRow & ":F" & headerRow)
        .Interior.Color = COLOR_DARK_RED
        .Font.Bold = True
        .Font.Color = COLOR_GOLD
    End With
End Sub

Private Sub SortBankSectionRows(ByVal ws As Worksheet, ByVal firstDataRow As Long, ByVal lastDataRow As Long)
    If lastDataRow < firstDataRow Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("G" & firstDataRow & ":G" & lastDataRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range("A" & firstDataRow & ":A" & lastDataRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A" & firstDataRow & ":G" & lastDataRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Sub ApplyBankSectionFrequencyFill(ByVal ws As Worksheet, ByVal firstDataRow As Long, ByVal lastDataRow As Long, ByVal sectionType As Long)
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim frequencyText As String
    Dim groupIndex As Long
    Dim fillColor As Long
    Dim oldDisplayAlerts As Boolean

    If lastDataRow < firstDataRow Then Exit Sub

    groupStart = firstDataRow
    groupIndex = 0

    Do While groupStart <= lastDataRow
        frequencyText = NormalizeText(ws.Cells(groupStart, 7).Value)
        groupEnd = groupStart

        Do While groupEnd + 1 <= lastDataRow And NormalizeText(ws.Cells(groupEnd + 1, 7).Value) = frequencyText
            groupEnd = groupEnd + 1
        Loop

        If groupIndex Mod 2 = 0 Then
            fillColor = COLOR_DARK_RED
        Else
            fillColor = COLOR_DARK_RED_ALT
        End If

        ws.Range("B" & groupStart & ":F" & groupEnd).Interior.Color = fillColor

        If (sectionType <= 2 Or sectionType = 4) And groupEnd > groupStart Then
            oldDisplayAlerts = Application.DisplayAlerts
            Application.DisplayAlerts = False
            ws.Range("D" & groupStart & ":D" & groupEnd).Merge
            Application.DisplayAlerts = oldDisplayAlerts
        End If

        If sectionType <= 2 Or sectionType = 4 Then
            With ws.Range("D" & groupStart & ":D" & groupEnd)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If

        groupStart = groupEnd + 1
        groupIndex = groupIndex + 1
    Loop
End Sub

Private Function IsBankSeriesMatch(ByVal seriesValue As Variant, ByVal sectionType As Long) As Boolean
    Dim seriesText As String
    seriesText = NormalizeText(seriesValue)
    If Len(seriesText) = 0 Then Exit Function

    If sectionType = 1 Then
        IsBankSeriesMatch = (InStr(1, seriesText, "稳健", vbTextCompare) > 0)
    ElseIf sectionType = 2 Then
        IsBankSeriesMatch = (InStr(1, seriesText, "汇益", vbTextCompare) > 0 _
                             And InStr(1, seriesText, "稳健", vbTextCompare) = 0 _
                             And InStr(1, seriesText, "稳享", vbTextCompare) = 0)
    ElseIf sectionType = 3 Then
        IsBankSeriesMatch = (InStr(1, seriesText, "稳享", vbTextCompare) > 0)
    Else
        IsBankSeriesMatch = (InStr(1, seriesText, "蓝色港湾", vbTextCompare) > 0)
    End If
End Function

Private Function PickBankCycleAnnual(ByVal wsSource As Worksheet, ByVal sourceRow As Long, ByVal intervalValue As Variant) As Variant
    Dim intervalText As String
    intervalText = NormalizeText(intervalValue)

    If IsWeeklyInterval(intervalValue) Then
        PickBankCycleAnnual = wsSource.Cells(sourceRow, BANK_COL_7DAY_ANNUAL).Value
    ElseIf intervalText = "28" Or InStr(1, intervalText, "四周", vbTextCompare) > 0 Then
        PickBankCycleAnnual = wsSource.Cells(sourceRow, BANK_COL_28DAY_ANNUAL).Value
    Else
        PickBankCycleAnnual = vbNullString
    End If
End Function

Private Sub WriteValueOrSlash(ByVal targetCell As Range, ByVal sourceValue As Variant)
    If IsError(sourceValue) Or IsEmpty(sourceValue) Or Len(NormalizeText(sourceValue)) = 0 Then
        targetCell.Value = "\"
    Else
        targetCell.Value = sourceValue
    End If
End Sub

Private Sub ApplyFrequencyGroupFillAndMerge(ByVal ws As Worksheet, ByVal firstDataRow As Long, ByVal lastDataRow As Long)
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim frequencyText As String
    Dim groupIndex As Long
    Dim fillColor As Long
    Dim oldDisplayAlerts As Boolean

    If lastDataRow < firstDataRow Then Exit Sub

    groupStart = firstDataRow
    groupIndex = 0

    Do While groupStart <= lastDataRow
        frequencyText = NormalizeText(ws.Cells(groupStart, 4).Value)
        groupEnd = groupStart

        Do While groupEnd + 1 <= lastDataRow And NormalizeText(ws.Cells(groupEnd + 1, 4).Value) = frequencyText
            groupEnd = groupEnd + 1
        Loop

        If groupIndex Mod 2 = 0 Then
            fillColor = COLOR_DARK_RED
        Else
            fillColor = COLOR_DARK_RED_ALT
        End If

        ws.Range("B" & groupStart & ":F" & groupEnd).Interior.Color = fillColor

        If groupEnd > groupStart Then
            oldDisplayAlerts = Application.DisplayAlerts
            Application.DisplayAlerts = False
            ws.Range("D" & groupStart & ":D" & groupEnd).Merge
            Application.DisplayAlerts = oldDisplayAlerts
        End If

        With ws.Range("D" & groupStart & ":D" & groupEnd)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        groupStart = groupEnd + 1
        groupIndex = groupIndex + 1
    Loop
End Sub

Private Sub WriteYuanRongAnXiangReport(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet, ByVal baselineDate As Date)
    Dim currentRow As Long
    Dim noteRow As Long
    Dim riskRow As Long

    wsTarget.Cells.ClearFormats
    wsTarget.Cells.ClearContents
    wsTarget.Cells.Font.Name = FONT_NAME_YUANRONG

    With wsTarget.Range("A1:E1")
        .Merge
        .Value = "交银国信·圆融安享汇益固收稳健系列信托计划" & vbLf & "历史到期产品收益情况"
        .Interior.Color = COLOR_YUANRONG_TITLE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Name = FONT_NAME_YUANRONG
        .Font.Bold = True
        .Font.Size = 20
        .Font.Color = COLOR_WHITE
    End With
    wsTarget.Rows(1).RowHeight = 56

    wsTarget.Range("A2:E2").Interior.Color = COLOR_WHITE
    wsTarget.Rows(2).RowHeight = 80

    currentRow = 3
    currentRow = WriteYuanRongAnXiangSection(wsSource, wsTarget, currentRow, "日开", baselineDate)

    wsTarget.Range("A13:E13").Interior.Color = COLOR_WHITE
    wsTarget.Rows(13).RowHeight = 80

    currentRow = 14
    currentRow = WriteYuanRongAnXiangSection(wsSource, wsTarget, currentRow, "周开", baselineDate)

    wsTarget.Range("A18:E18").Interior.Color = COLOR_WHITE
    wsTarget.Rows(18).RowHeight = 80

    currentRow = 19
    currentRow = WriteYuanRongAnXiangSection(wsSource, wsTarget, currentRow, "月开", baselineDate)

    noteRow = currentRow
    riskRow = noteRow + 1

    With wsTarget.Range("A" & noteRow & ":E" & noteRow)
        .Merge
        .Value = "*表中数据来源于托管人复核的产品净值数据，数据截至" & Format$(baselineDate, "yyyy-mm-dd") & "，仅供参考，产品有风险，投资需谨慎"
        ApplyYuanRongTitleFormat wsTarget.Range("A" & noteRow & ":E" & noteRow), 12, False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With wsTarget.Range("A" & riskRow & ":E" & riskRow)
        .Merge
        .Value = "风险提示:本产品由交银国际信托有限公司发行与管理，交通银行股份有限公司作为代销机构不承担产品的投资、兑付责任；" & vbLf & _
                 "*请您认真阅读信托合同、产品说明书、风险申明书等法律文件，根据风险承受能力选择合适的产品；" & vbLf & _
                 "*信托计划不承诺保证本金不受损失或最低收益，过往业绩并不预示其未来表现，产品发行人管理的其他产品的业绩并不构成未来产品业绩表现的保证；" & vbLf & _
                 "*下表中信托产品的代销机构风险评级为3R-平衡型，该类产品的风险中等，所投资金存在一定亏损风险，收益或利益浮动且有一定波动。"
        ApplyYuanRongTitleFormat wsTarget.Range("A" & riskRow & ":E" & riskRow), 9, False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    wsTarget.Rows(riskRow).AutoFit
    If wsTarget.Rows(riskRow).RowHeight < 70 Then wsTarget.Rows(riskRow).RowHeight = 70

    ApplyWhiteBorder wsTarget.Range("A1:E" & riskRow)
    wsTarget.Columns("A").ColumnWidth = 50
    wsTarget.Columns("B").ColumnWidth = 14
    wsTarget.Columns("C").ColumnWidth = 14
    wsTarget.Columns("D").ColumnWidth = 16
    wsTarget.Columns("E").ColumnWidth = 20
    wsTarget.Columns("F").Hidden = True
    AddYuanRongLogoIfExists wsTarget
End Sub

Private Function WriteYuanRongAnXiangSection(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet, _
                                             ByVal headerRow As Long, ByVal seriesName As String, _
                                             ByVal baselineDate As Date) As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim nextRow As Long
    Dim sourceRow As Long
    Dim lastSourceRow As Long
    Dim intervalValue As Variant

    WriteYuanRongAnXiangHeaders wsTarget, headerRow

    firstDataRow = headerRow + 1
    nextRow = firstDataRow

    lastSourceRow = LastUsedRow(wsSource)
    For sourceRow = 2 To lastSourceRow
        If StrComp(NormalizeText(wsSource.Cells(sourceRow, YRA_COL_SERIES).Value), seriesName, vbTextCompare) = 0 Then
            intervalValue = wsSource.Cells(sourceRow, YRA_COL_THEORETICAL_INTERVAL).Value

            wsTarget.Cells(nextRow, 1).Value = wsSource.Cells(sourceRow, YRA_COL_PRODUCT_NAME).Value
            WriteNextOpenDateCell wsTarget.Cells(nextRow, 2), wsSource.Cells(sourceRow, YRA_COL_NEXT_OPEN).Value, intervalValue, baselineDate
            wsTarget.Cells(nextRow, 3).Value = FormatOpenFrequency(intervalValue)
            wsTarget.Cells(nextRow, 4).Value = wsSource.Cells(sourceRow, YRA_COL_7DAY_ANNUAL).Value
            wsTarget.Cells(nextRow, 5).Value = wsSource.Cells(sourceRow, YRA_COL_28DAY_ANNUAL).Value
            wsTarget.Cells(nextRow, 6).Value = BuildSortKey(wsSource.Cells(sourceRow, YRA_COL_SEQ).Value)
            wsTarget.Range("D" & nextRow & ":E" & nextRow).NumberFormat = "0.00%"
            nextRow = nextRow + 1
        End If
    Next sourceRow

    lastDataRow = nextRow - 1
    If lastDataRow >= firstDataRow Then
        SortYuanRongAnXiangRows wsTarget, firstDataRow, lastDataRow
        wsTarget.Range("F" & firstDataRow & ":F" & lastDataRow).ClearContents
        ApplyYuanRongRecordFormat wsTarget.Range("A" & firstDataRow & ":E" & lastDataRow)
        ApplyYuanRongFrequencyMerge wsTarget, firstDataRow, lastDataRow
    End If

    FormatYuanRongSectionBlock wsTarget, headerRow, IIf(lastDataRow >= firstDataRow, lastDataRow, headerRow)
    WriteYuanRongAnXiangSection = IIf(lastDataRow >= firstDataRow, lastDataRow + 1, headerRow + 1)
End Function

Private Sub WriteYuanRongAnXiangHeaders(ByVal ws As Worksheet, ByVal headerRow As Long)
    ws.Range("A" & headerRow).Value = "产品名称"
    ws.Range("B" & headerRow).Value = "下一开放日"
    ws.Range("C" & headerRow).Value = "开放频率"
    ws.Range("D" & headerRow).Value = "7日年化收益率"
    ws.Range("E" & headerRow).Value = "28日年化收益率"
End Sub

Private Sub FormatYuanRongSectionBlock(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long)
    ApplyYuanRongTitleFormat ws.Range("A" & headerRow & ":E" & headerRow), 14, True

    With ws.Range("A" & headerRow & ":E" & lastRow)
        .RowHeight = 22
        .Font.Name = FONT_NAME_YUANRONG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    If lastRow > headerRow Then
        ws.Range("A" & (headerRow + 1) & ":A" & lastRow).HorizontalAlignment = xlLeft
    End If
    ApplyWhiteBorder ws.Range("A" & headerRow & ":E" & lastRow)
End Sub

Private Sub ApplyYuanRongTitleFormat(ByVal targetRange As Range, ByVal fontSize As Long, ByVal boldText As Boolean)
    With targetRange
        .Interior.Color = COLOR_YUANRONG_TITLE
        .Font.Name = FONT_NAME_YUANRONG
        .Font.Size = fontSize
        .Font.Bold = boldText
        .Font.Color = COLOR_WHITE
    End With
End Sub

Private Sub ApplyYuanRongRecordFormat(ByVal targetRange As Range)
    With targetRange
        .Interior.Color = COLOR_YUANRONG_RECORD
        .Font.Name = FONT_NAME_YUANRONG
        .Font.Size = 14
        .Font.Bold = False
        .Font.Color = COLOR_BLACK
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    targetRange.Columns(1).HorizontalAlignment = xlLeft
    ApplyWhiteBorder targetRange
End Sub

Private Sub ApplyYuanRongFrequencyMerge(ByVal ws As Worksheet, ByVal firstDataRow As Long, ByVal lastDataRow As Long)
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim frequencyText As String
    Dim oldDisplayAlerts As Boolean

    If lastDataRow < firstDataRow Then Exit Sub

    groupStart = firstDataRow
    Do While groupStart <= lastDataRow
        frequencyText = NormalizeText(ws.Cells(groupStart, 3).Value)
        groupEnd = groupStart

        Do While groupEnd + 1 <= lastDataRow And NormalizeText(ws.Cells(groupEnd + 1, 3).Value) = frequencyText
            groupEnd = groupEnd + 1
        Loop

        If groupEnd > groupStart Then
            oldDisplayAlerts = Application.DisplayAlerts
            Application.DisplayAlerts = False
            ws.Range("C" & groupStart & ":C" & groupEnd).Merge
            Application.DisplayAlerts = oldDisplayAlerts
        End If

        With ws.Range("C" & groupStart & ":C" & groupEnd)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        groupStart = groupEnd + 1
    Loop
End Sub

Private Sub AddYuanRongLogoIfExists(ByVal ws As Worksheet)
    Dim logoPath As String
    logoPath = ThisWorkbook.Path & Application.PathSeparator & "logo.png"
    If Dir(logoPath) = vbNullString Then Exit Sub

    Dim targetCell As Range
    Set targetCell = ws.Range("E1")

    DeletePicturesInRange ws, targetCell

    Dim shp As Shape
    Set shp = ws.Shapes.AddPicture( _
        FileName:=logoPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=targetCell.Left, _
        Top:=targetCell.Top, _
        Width:=-1, _
        Height:=-1)

    shp.Name = "圆融安享Logo"
    shp.LockAspectRatio = msoTrue
    shp.Width = targetCell.Width
    shp.Left = targetCell.Left + targetCell.Width - shp.Width
    shp.Top = targetCell.Top
    shp.Placement = xlMove
End Sub

Private Sub SortYuanRongAnXiangRows(ByVal ws As Worksheet, ByVal firstDataRow As Long, ByVal lastDataRow As Long)
    If lastDataRow < firstDataRow Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("F" & firstDataRow & ":F" & lastDataRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A" & firstDataRow & ":F" & lastDataRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Sub AddTitleImageIfExists(ByVal ws As Worksheet)
    Dim imagePath As String
    imagePath = ThisWorkbook.Path & Application.PathSeparator & TITLE_IMAGE_NAME
    If Dir(imagePath) = vbNullString Then Exit Sub

    Dim targetRange As Range
    Set targetRange = ws.Range("A1:G2")

    Dim shp As Shape
    Set shp = ws.Shapes.AddPicture( _
        FileName:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=targetRange.Left, _
        Top:=targetRange.Top, _
        Width:=-1, _
        Height:=-1)

    shp.LockAspectRatio = msoTrue
    shp.Width = targetRange.Width
    shp.Placement = xlFreeFloating

    Dim imageHeight As Double
    imageHeight = shp.Height

    ws.Rows(1).RowHeight = imageHeight / 2
    ws.Rows(2).RowHeight = imageHeight / 2

    Set targetRange = ws.Range("A1:G2")
    shp.Left = targetRange.Left
    shp.Top = targetRange.Top
    shp.Placement = xlMove
End Sub

Private Function InsertChartImagesIntoDisplay(ByVal wbOutput As Workbook, ByVal dbPath As String, ByVal baselineKey As String) As String
    Dim imgKey As String
    imgKey = FindLatestChartFolderKey(dbPath)

    If Len(imgKey) = 0 Then
        InsertChartImagesIntoDisplay = "【图表图片提醒】未找到 产品图表_yyyymmdd 文件夹，展示报表已保留图片空位。"
        Exit Function
    End If

    Dim imgFolder As String
    imgFolder = dbPath & "产品图表_" & imgKey & Application.PathSeparator

    If Dir(imgFolder, vbDirectory) = vbNullString Then
        InsertChartImagesIntoDisplay = "【图表图片提醒】未找到图片文件夹：" & imgFolder & "，展示报表已保留图片空位。"
        Exit Function
    End If

    Dim missingImages As String
    Dim insertedCount As Long

    InsertChartImageGroup wbOutput.Worksheets(CAT_STABLE), imgFolder, _
        Array("汇益稳享364天101号_红.png", _
              "汇益稳享728天108号_红.png"), _
        Array("B4:F4", _
              "B5:F5"), _
        insertedCount, missingImages

    InsertChartImageGroup wbOutput.Worksheets(CAT_BANK), imgFolder, _
        Array("汇益固收稳健7天2号_红.png", _
              "汇益固收稳健28天6号_红.png", _
              "汇益稳享91天3号_红.png"), _
        Array("B4:F4", _
              "B5:F5", _
              "B6:F6"), _
        insertedCount, missingImages

    InsertChartImageGroup wbOutput.Worksheets(CAT_DIRECT), imgFolder, _
        Array("汇益固收稳健日开101号_红.png", _
              "汇益固收稳健28天101号_红.png", _
              "交鑫致远6个月101号_红.png"), _
        Array("B4:F4", _
              "B5:F5", _
              "B6:F6"), _
        insertedCount, missingImages

    InsertChartImageGroup wbOutput.Worksheets(CAT_YUANRONG_ANXIANG), imgFolder, _
        Array("圆融安享汇益固收稳健日开8号_蓝.png", _
              "圆融安享汇益固收稳健7天2号_蓝.png", _
              "圆融安享汇益固收稳健28天1号_蓝.png"), _
        Array("A2:E2", _
              "A13:E13", _
              "A18:E18"), _
        insertedCount, missingImages

    Dim notice As String
    notice = "【图表图片提醒】使用图片日期：" & imgKey

    If Len(baselineKey) > 0 And imgKey <> baselineKey Then
        notice = notice & vbCrLf & "【重要】图片日期与展示报表基准日期不一致：图片日期 " & imgKey & "，基准日期 " & baselineKey & "。"
    End If

    notice = notice & vbCrLf & "已插入图表图片：" & insertedCount & " 张。"

    If Len(missingImages) > 0 Then
        notice = notice & vbCrLf & "以下图片未找到，相关空位已保留：" & vbCrLf & Left$(missingImages, Len(missingImages) - 2)
    End If

    InsertChartImagesIntoDisplay = notice
End Function

Private Sub InsertChartImageGroup(ByVal ws As Worksheet, ByVal imgFolder As String, ByVal imgNames As Variant, _
                                  ByVal targetAddresses As Variant, ByRef insertedCount As Long, ByRef missingImages As String)
    If UBound(imgNames) - LBound(imgNames) <> UBound(targetAddresses) - LBound(targetAddresses) Then
        missingImages = missingImages & ws.Name & "(图片数量与目标区域数量不一致), "
        Exit Sub
    End If

    Dim i As Long
    For i = LBound(imgNames) To UBound(imgNames)
        Dim imgPath As String
        imgPath = imgFolder & CStr(imgNames(i))

        Dim targetRng As Range
        Set targetRng = ws.Range(CStr(targetAddresses(i)))
        targetRng.Interior.Color = COLOR_WHITE

        If Dir(imgPath) = vbNullString Then
            missingImages = missingImages & ws.Name & "/" & CStr(imgNames(i)) & ", "
        Else
            DeletePicturesInRange ws, targetRng
            InsertPictureFitWidthAndSetRowHeight ws, imgPath, targetRng, targetRng.Row
            insertedCount = insertedCount + 1
        End If
    Next i
End Sub

Private Function FindLatestChartFolderKey(ByVal dbPath As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^产品图表_(\d{8})$"
    regex.IgnoreCase = True

    Dim folderName As String
    folderName = Dir(dbPath & "产品图表_*", vbDirectory)

    Dim latestKey As String
    Dim matches As Object

    Do While Len(folderName) > 0
        If folderName <> "." And folderName <> ".." Then
            If (GetAttr(dbPath & folderName) And vbDirectory) = vbDirectory Then
                If regex.Test(folderName) Then
                    Set matches = regex.Execute(folderName)
                    If matches(0).SubMatches(0) > latestKey Then
                        latestKey = matches(0).SubMatches(0)
                    End If
                End If
            End If
        End If
        folderName = Dir()
    Loop

    FindLatestChartFolderKey = latestKey
End Function

Private Sub DeletePicturesInRange(ByVal ws As Worksheet, ByVal targetRng As Range)
    Dim i As Long
    Dim shp As Shape

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            If ShapeOverlapsRange(shp, targetRng) Then
                shp.Delete
            End If
        End If
    Next i
End Sub

Private Function ShapeOverlapsRange(ByVal shp As Shape, ByVal rng As Range) As Boolean
    Dim shpLeft As Double, shpRight As Double
    Dim shpTop As Double, shpBottom As Double
    Dim rngLeft As Double, rngRight As Double
    Dim rngTop As Double, rngBottom As Double

    shpLeft = shp.Left
    shpRight = shp.Left + shp.Width
    shpTop = shp.Top
    shpBottom = shp.Top + shp.Height

    rngLeft = rng.Left
    rngRight = rng.Left + rng.Width
    rngTop = rng.Top
    rngBottom = rng.Top + rng.Height

    ShapeOverlapsRange = Not ( _
        shpRight < rngLeft Or _
        shpLeft > rngRight Or _
        shpBottom < rngTop Or _
        shpTop > rngBottom)
End Function

Private Sub InsertPictureFitWidthAndSetRowHeight(ByVal ws As Worksheet, ByVal imgPath As String, _
                                                  ByVal targetRng As Range, ByVal rowNum As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddPicture( _
        FileName:=imgPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=targetRng.Left, _
        Top:=targetRng.Top, _
        Width:=-1, _
        Height:=-1)

    shp.Name = "展示图_" & CStr(rowNum)
    shp.LockAspectRatio = msoTrue
    shp.Width = targetRng.Width
    shp.Placement = xlMove

    Dim picW As Double
    Dim picH As Double
    picW = shp.Width
    picH = shp.Height

    Dim newRowHeight As Double
    newRowHeight = picH + 2
    If newRowHeight > 409.5 Then newRowHeight = 409.5
    ws.Rows(rowNum).RowHeight = newRowHeight

    shp.LockAspectRatio = msoTrue
    shp.Left = targetRng.Left
    shp.Top = targetRng.Top
    shp.Width = picW
    shp.Height = picH
    shp.Placement = xlMove
End Sub

Private Function GetBaselineDateFromProductCategory() As Date
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(PRODUCT_CATEGORY_SHEET_NAME)

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(ws, 1)
    If headerMap Is Nothing Then Err.Raise vbObjectError + 4201, , "无法读取产品分类表头"
    If Not headerMap.Exists(COL_BASELINE_DATE) Then Err.Raise vbObjectError + 4202, , "产品分类缺少字段：" & COL_BASELINE_DATE

    Dim baselineCol As Long
    baselineCol = CLng(headerMap(COL_BASELINE_DATE))

    Dim lastRow As Long
    Dim r As Long
    Dim parsedDate As Date
    lastRow = LastUsedRow(ws)

    For r = 2 To lastRow
        If TryReadDate(ws.Cells(r, baselineCol).Value, parsedDate) Then
            GetBaselineDateFromProductCategory = parsedDate
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 4203, , "产品分类未找到可用基准日期"
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
        outDate = DateValue(CDate(value))
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
        outDate = DateValue(CDate(textValue))
        TryReadDate = True
    End If
End Function
