Option Explicit

'==============================================================
' Module: Weekly Recommendation Material
' Entry : Weekly02_GenerateReport
' Output: 推荐材料-yyyymmdd.xlsx
'==============================================================

Private Const SHEET_PRODUCT_ELEMENT As String = "产品要素"
Private Const SHEET_ASSET As String = "主要底层资产"
Private Const SHEET_PRODUCT_CATEGORY As String = "产品分类"
Private Const SHEET_NAV_DATA As String = "绘图净值数据"
Private Const SHEET_NAV_DATA_ALIAS As String = "绘制净值数据"

Private Const COLOR_BLUE As String = "#19449A"
Private Const FONT_NAME As String = "微软雅黑"

Private Const TITLE_FONT_SIZE As Long = 36
Private Const SUBTITLE_FONT_SIZE As Long = 20
Private Const NORMAL_FONT_SIZE As Long = 18
Private Const NORMAL_ROW_HEIGHT As Double = 30

Public Sub Weekly02_GenerateReport()
    Dim t0 As Double: t0 = Timer
    
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim wsProduct As Worksheet: Set wsProduct = RequireSheet(wbDB, SHEET_PRODUCT_ELEMENT)
    Dim wsAsset As Worksheet: Set wsAsset = RequireSheet(wbDB, SHEET_ASSET)
    Dim wsCategory As Worksheet: Set wsCategory = RequireSheet(wbDB, SHEET_PRODUCT_CATEGORY)
    Dim wsNav As Worksheet: Set wsNav = GetNavDataSheet(wbDB)
    
    Dim baseDate As Date
    baseDate = GetBaseDate(wsCategory)
    
    Dim wbOut As Workbook
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    
    Dim productHeaders As Object
    Set productHeaders = GetHeaderMap(wsProduct)
    RequireHeaders productHeaders, Array("信托计划代码", "产品简称", "信托项目名称", "成立日", _
                                        "受托人", "产品类型", "风险评级", "产品期限", _
                                        "申赎规则", "费率", "投资范围", "产品规模")
    
    Dim assetHeaders As Object
    Set assetHeaders = GetHeaderMap(wsAsset)
    RequireHeaders assetHeaders, Array("信托计划代码", "层级", "节点名称", "节点说明", "排序")
    
    Dim navHeaders As Object
    Set navHeaders = GetHeaderMap(wsNav)
    RequireHeaders navHeaders, Array("净值日期", "产品编号", "净值")
    
    Dim lastProductRow As Long
    lastProductRow = wsProduct.Cells(wsProduct.Rows.Count, HeaderColumn(productHeaders, "信托计划代码")).End(xlUp).Row
    
    Dim productCount As Long: productCount = 0
    Dim r As Long
    For r = 2 To lastProductRow
        Dim productCode As String
        productCode = Trim(CStr(wsProduct.Cells(r, HeaderColumn(productHeaders, "信托计划代码")).Value))
        
        Dim productShort As String
        productShort = Trim(CStr(wsProduct.Cells(r, HeaderColumn(productHeaders, "产品简称")).Value))
        
        If Len(productCode) > 0 And Len(productShort) > 0 Then
            productCount = productCount + 1
            
            Dim wsOut As Worksheet
            If productCount = 1 Then
                Set wsOut = wbOut.Worksheets(1)
            Else
                Set wsOut = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
            End If
            
            wsOut.Name = UniqueSheetName(wbOut, SafeSheetName(productShort), wsOut)
            BuildOneProductSheet wsOut, wsProduct, productHeaders, r, wsAsset, assetHeaders, _
                                 wsNav, navHeaders, productCode, productShort, baseDate
        End If
    Next r
    
    If productCount = 0 Then
        Err.Raise vbObjectError + 101, , "产品要素中未找到可生成的产品。"
    End If
    
    Dim outputPath As String
    outputPath = wbDB.Path & "\推荐材料-" & Format(baseDate, "yyyymmdd") & ".xlsx"
    
    wbOut.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    
    MsgBox "推荐材料生成完成" & vbCrLf & vbCrLf & _
           "耗时：" & Format(Timer - t0, "0.0") & " 秒" & vbCrLf & vbCrLf & _
           "处理结果：" & vbCrLf & _
           "产品数量：" & productCount & vbCrLf & vbCrLf & _
           "输出文件：" & vbCrLf & outputPath, vbInformation, "推荐材料生成"

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Exit Sub

EH:
    MsgBox "推荐材料生成失败" & vbCrLf & vbCrLf & _
           "错误信息：" & Err.Description, vbExclamation, "推荐材料生成"
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    On Error GoTo 0
    Resume CleanExit
End Sub


Private Sub BuildOneProductSheet(ByVal wsOut As Worksheet, _
                                 ByVal wsProduct As Worksheet, _
                                 ByVal productHeaders As Object, _
                                 ByVal productRow As Long, _
                                 ByVal wsAsset As Worksheet, _
                                 ByVal assetHeaders As Object, _
                                 ByVal wsNav As Worksheet, _
                                 ByVal navHeaders As Object, _
                                 ByVal productCode As String, _
                                 ByVal productShort As String, _
                                 ByVal baseDate As Date)
    PrepareSheetLayout wsOut
    
    Dim curRow As Long: curRow = 2
    
    BuildTopTitle wsOut, curRow, productShort
    curRow = curRow + 1
    
    BuildBaseInfo wsOut, curRow, baseDate
    curRow = curRow + 1
    
    BuildTeamInfo wsOut, curRow
    curRow = curRow + 1
    
    BuildSubtitle wsOut, curRow, "产品要素"
    curRow = curRow + 1
    
    curRow = BuildProductElementRows(wsOut, curRow, wsProduct, productHeaders, productRow)
    
    BuildSubtitle wsOut, curRow, "主要底层资产"
    curRow = curRow + 1
    
    Dim flowRow As Long: flowRow = curRow
    wsOut.Rows(flowRow).RowHeight = 170
    wsOut.Range("B" & flowRow & ":E" & flowRow).Merge
    ApplyNormalStyle wsOut.Range("B" & flowRow & ":E" & flowRow)
    DrawAssetFlow wsOut, wsAsset, assetHeaders, productCode, productShort, flowRow
    curRow = curRow + 1
    
    curRow = BuildAssetDescriptionRows(wsOut, curRow, wsAsset, assetHeaders, productCode)
    
    BuildSubtitle wsOut, curRow, "产品业绩展示"
    curRow = curRow + 1
    
    curRow = BuildPerformanceRows(wsOut, curRow, wsProduct, productHeaders, productRow, _
                                  wsNav, navHeaders, productCode, productShort, baseDate)
    
    Dim chartRow As Long: chartRow = curRow
    wsOut.Range("B" & chartRow & ":E" & chartRow).Merge
    ApplyNormalStyle wsOut.Range("B" & chartRow & ":E" & chartRow)
    InsertChartImage wsOut, productShort, chartRow
    
    Dim blankRow As Long: blankRow = chartRow + 1
    wsOut.Rows(blankRow).RowHeight = NORMAL_ROW_HEIGHT
    wsOut.Range("B" & blankRow & ":E" & blankRow).Merge
    ApplyNormalStyle wsOut.Range("B" & blankRow & ":E" & blankRow)
    wsOut.Range("B" & blankRow & ":E" & blankRow).Borders.LineStyle = xlNone
    
    ApplyFinalBorders wsOut, chartRow
    SetupPage wsOut
End Sub

Private Sub PrepareSheetLayout(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Cells.Font.Name = FONT_NAME
    ws.Cells.Font.Size = NORMAL_FONT_SIZE
    ws.Cells.VerticalAlignment = xlCenter
    ws.Cells.HorizontalAlignment = xlCenter
    
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 45
    ws.Columns("D").ColumnWidth = 45
    ws.Columns("E").ColumnWidth = 45
    ws.Columns("F").ColumnWidth = 5
    ws.Rows(1).RowHeight = NORMAL_ROW_HEIGHT
    
    ActiveWindow.DisplayGridlines = False
End Sub

Private Sub BuildTopTitle(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal productShort As String)
    Dim titleRange As Range
    Set titleRange = ws.Range("B" & rowNo & ":E" & rowNo)
    titleRange.Merge
    titleRange.Value = ""
    
    Dim bgPath As String
    bgPath = ThisWorkbook.Path & "\background.png"
    
    Dim pic As Shape
    If FileExists(bgPath) Then
        Set pic = ws.Shapes.AddPicture(bgPath, msoFalse, msoTrue, 0, 0, -1, -1)
        pic.LockAspectRatio = msoTrue
        If pic.Width > titleRange.Width Then pic.Width = titleRange.Width
        ws.Rows(rowNo).RowHeight = pic.Height
        pic.Left = titleRange.Left + (titleRange.Width - pic.Width) / 2
        pic.Top = titleRange.Top + (titleRange.Height - pic.Height) / 2
        pic.Placement = xlMove
    Else
        ws.Rows(rowNo).RowHeight = 80
    End If
    
    Dim tb As Shape
    Set tb = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, titleRange.Left, titleRange.Top, titleRange.Width, titleRange.Height)
    With tb
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = productShort
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextFrame2.TextRange.Font
            .NameFarEast = FONT_NAME
            .Name = FONT_NAME
            .Size = TITLE_FONT_SIZE
            .Bold = msoTrue
            .Fill.ForeColor.RGB = HexToLong(COLOR_BLUE)
        End With
        .Placement = xlMove
    End With
    
    InsertTitleLogo ws, rowNo
End Sub

Private Sub BuildBaseInfo(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal baseDate As Date)
    With ws.Range("B" & rowNo & ":E" & rowNo)
        .Merge
        .Value = "数据截止至" & Format(baseDate, "yyyy年m月d日")
        ApplyNormalStyle .Cells
    End With
    ws.Rows(rowNo).RowHeight = NORMAL_ROW_HEIGHT
End Sub

Private Sub BuildTeamInfo(ByVal ws As Worksheet, ByVal rowNo As Long)
    With ws.Range("B" & rowNo & ":E" & rowNo)
        .Merge
        .Value = "资产管理事业部-固定收益业务团队"
        ApplyNormalStyle .Cells
        .HorizontalAlignment = xlRight
    End With
    ws.Rows(rowNo).RowHeight = NORMAL_ROW_HEIGHT
End Sub

Private Sub InsertTitleLogo(ByVal ws As Worksheet, ByVal rowNo As Long)
    Dim logoPath As String
    logoPath = ThisWorkbook.Path & "\logo_white.png"
    If Not FileExists(logoPath) Then Exit Sub
    
    Dim logoCell As Range
    Set logoCell = ws.Range("E1")
    
    Dim titleRange As Range
    Set titleRange = ws.Range("B" & rowNo & ":E" & rowNo)
    
    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture(logoPath, msoFalse, msoTrue, 0, 0, -1, -1)
    With pic
        .LockAspectRatio = msoTrue
        .Width = logoCell.Width * 2 / 3
        .Left = logoCell.Left + logoCell.Width - .Width
        .Top = titleRange.Top
        .Placement = xlMove
    End With
End Sub

Private Sub BuildSubtitle(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal titleText As String)
    With ws.Range("B" & rowNo & ":E" & rowNo)
        .Merge
        .Value = titleText
        .Font.Name = FONT_NAME
        .Font.Size = SUBTITLE_FONT_SIZE
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = HexToLong(COLOR_BLUE)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(rowNo).RowHeight = NORMAL_ROW_HEIGHT
End Sub

Private Function BuildProductElementRows(ByVal ws As Worksheet, _
                                         ByVal startRow As Long, _
                                         ByVal wsProduct As Worksheet, _
                                         ByVal headers As Object, _
                                         ByVal productRow As Long) As Long
    Dim labels As Variant
    Dim fields As Variant
    labels = Array("产品名称", "受托人", "产品类型", "风险评级", "产品期限", _
                   "申赎规则", "费率", "投资范围", "产品规模", "产品形态")
    fields = Array("信托项目名称", "受托人", "产品类型", "风险评级", "产品期限", _
                   "申赎规则", "费率", "投资范围", "产品规模", "产品形态")
    
    Dim rowNo As Long: rowNo = startRow
    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        If headers.Exists(CStr(fields(i))) Then
            Dim valueText As String
            valueText = Trim(CStr(wsProduct.Cells(productRow, headers(CStr(fields(i)))).Value))
            
            If CStr(fields(i)) <> "产品形态" Or Len(valueText) > 0 Then
                WriteLabelValueRow ws, rowNo, CStr(labels(i)), valueText
                If CStr(fields(i)) = "风险评级" Then
                    ws.Range("C" & rowNo & ":E" & rowNo).Font.Color = RGB(255, 0, 0)
                End If
                rowNo = rowNo + 1
            End If
        End If
    Next i
    
    BuildProductElementRows = rowNo
End Function

Private Function BuildAssetDescriptionRows(ByVal ws As Worksheet, _
                                           ByVal startRow As Long, _
                                           ByVal wsAsset As Worksheet, _
                                           ByVal headers As Object, _
                                           ByVal productCode As String) As Long
    Dim rows As Variant
    rows = GetLevel2AssetRows(wsAsset, headers, productCode)
    
    Dim rowNo As Long: rowNo = startRow
    Dim i As Long
    If IsEmpty(rows) Then
        WriteLabelValueRow ws, rowNo, "-", "-"
        rowNo = rowNo + 1
    Else
        For i = LBound(rows, 1) To UBound(rows, 1)
            WriteLabelValueRow ws, rowNo, CStr(rows(i, 1)), CStr(rows(i, 2))
            rowNo = rowNo + 1
        Next i
    End If
    
    BuildAssetDescriptionRows = rowNo
End Function

Private Function BuildPerformanceRows(ByVal ws As Worksheet, _
                                      ByVal startRow As Long, _
                                      ByVal wsProduct As Worksheet, _
                                      ByVal productHeaders As Object, _
                                      ByVal productRow As Long, _
                                      ByVal wsNav As Worksheet, _
                                      ByVal navHeaders As Object, _
                                      ByVal productCode As String, _
                                      ByVal productShort As String, _
                                      ByVal baseDate As Date) As Long
    Dim rate7 As Variant, rate30 As Variant, rateSince As Variant
    ReadPerformanceMetrics wsProduct, productHeaders, productRow, wsNav, navHeaders, _
                           productCode, baseDate, rate7, rate30, rateSince
    
    ws.Range("B" & startRow & ":E" & startRow).Value = Array("产品名称", "近7日年化收益率", "近30日年化收益率", "成立以来年化收益率")
    ws.Range("B" & (startRow + 1) & ":E" & (startRow + 1)).Value = Array(productShort, RateText(rate7), RateText(rate30), RateText(rateSince))
    
    ApplyNormalStyle ws.Range("B" & startRow & ":E" & (startRow + 1))
    ws.Rows(startRow & ":" & (startRow + 1)).RowHeight = NORMAL_ROW_HEIGHT
    BuildPerformanceRows = startRow + 2
End Function

Private Sub WriteLabelValueRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal labelText As String, ByVal valueText As String)
    ws.Range("B" & rowNo).Value = labelText
    ws.Range("C" & rowNo & ":E" & rowNo).Merge
    ws.Range("C" & rowNo).Value = valueText
    ApplyNormalStyle ws.Range("B" & rowNo & ":E" & rowNo)
    ws.Range("C" & rowNo).HorizontalAlignment = xlCenter
    ws.Rows(rowNo).RowHeight = NORMAL_ROW_HEIGHT
End Sub

Private Sub DrawAssetFlow(ByVal wsOut As Worksheet, _
                          ByVal wsAsset As Worksheet, _
                          ByVal headers As Object, _
                          ByVal productCode As String, _
                          ByVal productShort As String, _
                          ByVal flowRow As Long)
    Dim rows As Variant
    rows = GetLevel2AssetRows(wsAsset, headers, productCode)
    
    Dim area As Range
    Set area = wsOut.Range("B" & flowRow & ":E" & flowRow)
    
    Dim rootShape As Shape
    Set rootShape = AddNodeShape(wsOut, productShort, area.Left + (area.Width - 250) / 2, area.Top + 10, 250, 46, 20)
    
    If IsEmpty(rows) Then Exit Sub
    
    Dim n As Long
    n = UBound(rows, 1) - LBound(rows, 1) + 1
    
    Dim childW As Double: childW = 155
    Dim childH As Double: childH = 42
    Dim childTop As Double: childTop = area.Top + area.Height - childH - 12
    Dim busY As Double: busY = area.Top + 88
    Dim rootCenterX As Double: rootCenterX = rootShape.Left + rootShape.Width / 2
    
    AddFlowLine wsOut, rootCenterX, rootShape.Top + rootShape.Height, rootCenterX, busY, False
    
    Dim firstChildCenterX As Double
    Dim lastChildCenterX As Double
    firstChildCenterX = area.Left + (0.5 * area.Width / n)
    lastChildCenterX = area.Left + ((n - 0.5) * area.Width / n)
    AddFlowLine wsOut, firstChildCenterX, busY, lastChildCenterX, busY, False
    
    Dim i As Long
    For i = LBound(rows, 1) To UBound(rows, 1)
        Dim childLeft As Double
        childLeft = area.Left + ((i - LBound(rows, 1) + 0.5) * area.Width / n) - childW / 2
        
        Dim childShape As Shape
        Set childShape = AddNodeShape(wsOut, CStr(rows(i, 1)), childLeft, childTop, childW, childH, 17)
        
        AddFlowLine wsOut, childShape.Left + childShape.Width / 2, busY, _
                    childShape.Left + childShape.Width / 2, childShape.Top - 2, True
    Next i
End Sub

Private Function AddNodeShape(ByVal ws As Worksheet, _
                              ByVal textValue As String, _
                              ByVal leftPos As Double, _
                              ByVal topPos As Double, _
                              ByVal widthValue As Double, _
                              ByVal heightValue As Double, _
                              ByVal fontSize As Long) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthValue, heightValue)
    With shp
        .Fill.ForeColor.RGB = HexToLong(COLOR_BLUE)
        .Line.ForeColor.RGB = HexToLong(COLOR_BLUE)
        .TextFrame2.TextRange.Text = textValue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextFrame2.TextRange.Font
            .NameFarEast = FONT_NAME
            .Name = FONT_NAME
            .Size = fontSize
            .Bold = msoTrue
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
        .Placement = xlMove
    End With
    Set AddNodeShape = shp
End Function

Private Sub AddFlowLine(ByVal ws As Worksheet, _
                        ByVal x1 As Double, _
                        ByVal y1 As Double, _
                        ByVal x2 As Double, _
                        ByVal y2 As Double, _
                        ByVal hasArrow As Boolean)
    Dim lineShape As Shape
    Set lineShape = ws.Shapes.AddLine(x1, y1, x2, y2)
    With lineShape
        .Line.ForeColor.RGB = HexToLong(COLOR_BLUE)
        .Line.Weight = 2.25
        If hasArrow Then .Line.EndArrowheadStyle = msoArrowheadTriangle
        .Placement = xlMove
    End With
End Sub

Private Sub InsertChartImage(ByVal ws As Worksheet, ByVal productShort As String, ByVal rowNo As Long)
    Dim chartPath As String
    chartPath = FindLatestChartImage(productShort)
    
    Dim area As Range
    Set area = ws.Range("B" & rowNo & ":E" & rowNo)
    
    If Len(chartPath) = 0 Then
        ws.Rows(rowNo).RowHeight = 90
        area.Value = "未找到图表图片，请先运行图表图片导出步骤。"
        area.Font.Color = RGB(128, 128, 128)
        Exit Sub
    End If
    
    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture(chartPath, msoFalse, msoTrue, 0, 0, -1, -1)
    pic.LockAspectRatio = msoTrue
    
    pic.Width = area.Width
    ws.Rows(rowNo).RowHeight = pic.Height
    Set area = ws.Range("B" & rowNo & ":E" & rowNo)
    
    pic.Left = area.Left + (area.Width - pic.Width) / 2
    pic.Top = area.Top + (area.Height - pic.Height) / 2
    pic.Placement = xlMove
End Sub

Private Function FindLatestChartImage(ByVal productShort As String) As String
    Dim rootPath As String
    rootPath = ThisWorkbook.Path & "\"
    
    Dim folderName As String
    Dim latestKey As String
    Dim latestFolder As String
    
    folderName = Dir(rootPath & "产品图表_*", vbDirectory)
    Do While Len(folderName) > 0
        If folderName <> "." And folderName <> ".." Then
            If (GetAttr(rootPath & folderName) And vbDirectory) = vbDirectory Then
                Dim key As String
                key = Replace(folderName, "产品图表_", "")
                If Len(key) = 8 And key > latestKey Then
                    latestKey = key
                    latestFolder = rootPath & folderName & "\"
                End If
            End If
        End If
        folderName = Dir()
    Loop
    
    If Len(latestFolder) = 0 Then Exit Function
    
    Dim candidate As String
    candidate = latestFolder & productShort & "_蓝.png"
    If FileExists(candidate) Then
        FindLatestChartImage = candidate
        Exit Function
    End If
    
    candidate = latestFolder & productShort & "_红.png"
    If FileExists(candidate) Then
        FindLatestChartImage = candidate
    End If
End Function

Private Sub ReadPerformanceMetrics(ByVal wsProduct As Worksheet, _
                                   ByVal productHeaders As Object, _
                                   ByVal productRow As Long, _
                                   ByVal wsNav As Worksheet, _
                                   ByVal navHeaders As Object, _
                                   ByVal productCode As String, _
                                   ByVal baseDate As Date, _
                                   ByRef rate7 As Variant, _
                                   ByRef rate30 As Variant, _
                                   ByRef rateSince As Variant)
    rate7 = Empty
    rate30 = Empty
    rateSince = Empty
    
    Dim dateCol As Long: dateCol = HeaderColumn(navHeaders, "净值日期")
    Dim codeCol As Long: codeCol = HeaderColumn(navHeaders, "产品编号")
    Dim navCol As Long: navCol = HeaderColumn(navHeaders, "净值")
    ' 收集该产品全部净值数据点
    Dim navDates() As Date
    Dim navValues() As Double
    Dim dataCount As Long: dataCount = 0

    Dim lastRow As Long
    lastRow = wsNav.Cells(wsNav.Rows.Count, codeCol).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If Trim(CStr(wsNav.Cells(r, codeCol).Value)) = productCode Then
            If IsDate(wsNav.Cells(r, dateCol).Value) Then
                Dim navDate As Date
                navDate = CDate(wsNav.Cells(r, dateCol).Value)

                If navDate <= baseDate And IsNumeric(wsNav.Cells(r, navCol).Value) Then
                    Dim navValue As Double
                    navValue = CDbl(wsNav.Cells(r, navCol).Value)
                    If navValue > 0 Then
                        dataCount = dataCount + 1
                        ReDim Preserve navDates(1 To dataCount)
                        ReDim Preserve navValues(1 To dataCount)
                        navDates(dataCount) = navDate
                        navValues(dataCount) = navValue
                    End If
                End If
            End If
        End If
    Next r

    If dataCount = 0 Then Exit Sub

    ' 按日期降序排序（冒泡，数据量很小）
    Dim i As Long, j As Long
    For i = 1 To dataCount - 1
        For j = i + 1 To dataCount
            If navDates(j) > navDates(i) Then
                Dim tmpDate As Date: tmpDate = navDates(i)
                navDates(i) = navDates(j)
                navDates(j) = tmpDate
                Dim tmpNav As Double: tmpNav = navValues(i)
                navValues(i) = navValues(j)
                navValues(j) = tmpNav
            End If
        Next j
    Next i

    Dim latestNav As Double: latestNav = navValues(1)
    Dim latestDate As Date: latestDate = navDates(1)

    ' 从净值计算 7日/30日年化（单利公式，365天年）
    rate7 = CalcAnnualYieldFromNavSeries(latestNav, latestDate, navDates, navValues, dataCount, 7, 3)
    rate30 = CalcAnnualYieldFromNavSeries(latestNav, latestDate, navDates, navValues, dataCount, 30, 3)

    ' 成立以来年化（单利，沿用原逻辑）
    Dim inceptionCol As Long
    inceptionCol = HeaderColumn(productHeaders, "成立日")

    If IsDate(wsProduct.Cells(productRow, inceptionCol).Value) And latestNav > 0 Then
        Dim inceptionDate As Date
        inceptionDate = CDate(wsProduct.Cells(productRow, inceptionCol).Value)

        Dim daysCount As Long
        daysCount = DateDiff("d", inceptionDate, latestDate) + 1
        If daysCount > 0 Then
            rateSince = (latestNav - 1) * 365# / daysCount
        End If
    End If
End Sub


Private Function CalcAnnualYieldFromNavSeries(ByVal currentNav As Double, ByVal currentDate As Date, _
                                                ByRef navDates() As Date, ByRef navValues() As Double, _
                                                ByVal dataCount As Long, _
                                                ByVal targetDays As Long, ByVal maxBacktrack As Long) As Variant
    ' 单利公式: (V_t / V_base - 1) × (365 / 实际间隔)
    ' 匹配规则: 先找 -targetDays, 失败向前再尝试 maxBacktrack 次
    Dim offset As Long
    Dim i As Long
    Dim found As Boolean: found = False
    Dim pastNav As Double
    Dim actualDays As Long

    For offset = 0 To maxBacktrack
        Dim searchDate As Date
        searchDate = DateAdd("d", -(targetDays + offset), currentDate)

        For i = 2 To dataCount
            If navDates(i) = searchDate Then
                pastNav = navValues(i)
                actualDays = targetDays + offset
                found = True
                Exit For
            End If
        Next i

        If found Then Exit For
    Next offset

    If found And currentNav > 0 And pastNav > 0 Then
        CalcAnnualYieldFromNavSeries = (currentNav / pastNav - 1) * 365# / actualDays
    Else
        CalcAnnualYieldFromNavSeries = Empty
    End If
End Function

Private Function RateText(ByVal rateValue As Variant) As String
    If IsEmpty(rateValue) Or Not IsNumeric(rateValue) Then
        RateText = "-"
    Else
        RateText = Format(CDbl(rateValue), "0.0000%")
    End If
End Function

Private Function GetLevel2AssetRows(ByVal wsAsset As Worksheet, ByVal headers As Object, ByVal productCode As String) As Variant
    Dim codeCol As Long: codeCol = HeaderColumn(headers, "信托计划代码")
    Dim levelCol As Long: levelCol = HeaderColumn(headers, "层级")
    Dim nameCol As Long: nameCol = HeaderColumn(headers, "节点名称")
    Dim descCol As Long: descCol = HeaderColumn(headers, "节点说明")
    Dim sortCol As Long: sortCol = HeaderColumn(headers, "排序")
    
    Dim lastRow As Long
    lastRow = wsAsset.Cells(wsAsset.Rows.Count, codeCol).End(xlUp).Row
    
    Dim temp() As Variant
    Dim count As Long: count = 0
    
    Dim r As Long
    For r = 2 To lastRow
        If Trim(CStr(wsAsset.Cells(r, codeCol).Value)) = productCode Then
            If CLng(Val(wsAsset.Cells(r, levelCol).Value)) = 2 Then
                count = count + 1
                ReDim Preserve temp(1 To 3, 1 To count)
                temp(1, count) = Trim(CStr(wsAsset.Cells(r, nameCol).Value))
                temp(2, count) = Trim(CStr(wsAsset.Cells(r, descCol).Value))
                temp(3, count) = CLng(Val(wsAsset.Cells(r, sortCol).Value))
            End If
        End If
    Next r
    
    If count = 0 Then Exit Function
    
    Dim i As Long, j As Long
    For i = 1 To count - 1
        For j = i + 1 To count
            If CLng(temp(3, i)) > CLng(temp(3, j)) Then
                SwapAsset temp, i, j
            End If
        Next j
    Next i
    
    Dim result() As Variant
    ReDim result(1 To count, 1 To 2)
    For i = 1 To count
        result(i, 1) = temp(1, i)
        result(i, 2) = temp(2, i)
    Next i
    
    GetLevel2AssetRows = result
End Function

Private Sub SwapAsset(ByRef arr As Variant, ByVal i As Long, ByVal j As Long)
    Dim k As Long
    For k = 1 To 3
        Dim v As Variant
        v = arr(k, i)
        arr(k, i) = arr(k, j)
        arr(k, j) = v
    Next k
End Sub

Private Function GetBaseDate(ByVal wsCategory As Worksheet) As Date
    Dim headers As Object
    Set headers = GetHeaderMap(wsCategory)
    RequireHeaders headers, Array("基准日期")
    
    Dim baseCol As Long
    baseCol = HeaderColumn(headers, "基准日期")
    
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long
    lastRow = wsCategory.Cells(wsCategory.Rows.Count, baseCol).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastRow
        If IsDate(wsCategory.Cells(r, baseCol).Value) Then
            Dim d As Date
            d = CDate(wsCategory.Cells(r, baseCol).Value)
            seen(Format(d, "yyyy-mm-dd")) = d
        End If
    Next r
    
    If seen.Count = 0 Then
        Err.Raise vbObjectError + 201, , "产品分类中未找到基准日期。"
    End If
    
    If seen.Count > 1 Then
        Err.Raise vbObjectError + 202, , "产品分类中的基准日期不一致，请先检查后再生成推荐材料。"
    End If
    
    Dim baseItems As Variant
    baseItems = seen.Items
    GetBaseDate = CDate(baseItems(0))
End Function

Private Sub ApplyNormalStyle(ByVal rng As Range)
    With rng
        .Font.Name = FONT_NAME
        .Font.Size = NORMAL_FONT_SIZE
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .ShrinkToFit = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
    End With
End Sub

Private Sub ApplyFinalBorders(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim contentRange As Range
    Set contentRange = ws.Range("B2:E" & lastRow)
    With contentRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    
    With contentRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(0, 0, 0)
    End With
    With contentRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(0, 0, 0)
    End With
    With contentRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(0, 0, 0)
    End With
    With contentRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(0, 0, 0)
    End With
End Sub

Private Sub SetupPage(ByVal ws As Worksheet)
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
    End With
End Sub

Private Function GetHeaderMap(ByVal ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim c As Long
    For c = 1 To lastCol
        Dim key As String
        key = Trim(CStr(ws.Cells(1, c).Value))
        If Len(key) > 0 Then dict(key) = c
    Next c
    
    Set GetHeaderMap = dict
End Function

Private Sub RequireHeaders(ByVal headers As Object, ByVal names As Variant)
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Not headers.Exists(CStr(names(i))) Then
            Err.Raise vbObjectError + 301, , "缺少必需字段: " & CStr(names(i))
        End If
    Next i
End Sub

Private Function HeaderColumn(ByVal headers As Object, ByVal headerName As String) As Long
    If Not headers.Exists(headerName) Then
        Err.Raise vbObjectError + 302, , "缺少字段: " & headerName
    End If
    HeaderColumn = CLng(headers(headerName))
End Function

Private Function RequireSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set RequireSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If RequireSheet Is Nothing Then
        Err.Raise vbObjectError + 401, , "未找到 sheet: " & sheetName
    End If
End Function

Private Function GetNavDataSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetNavDataSheet = wb.Worksheets(SHEET_NAV_DATA)
    If GetNavDataSheet Is Nothing Then Set GetNavDataSheet = wb.Worksheets(SHEET_NAV_DATA_ALIAS)
    On Error GoTo 0
    
    If GetNavDataSheet Is Nothing Then
        Err.Raise vbObjectError + 402, , "未找到 sheet: " & SHEET_NAV_DATA & " 或 " & SHEET_NAV_DATA_ALIAS
    End If
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir(filePath)) > 0)
End Function

Private Function HexToLong(ByVal hexValue As String) As Long
    Dim v As String
    v = Replace(hexValue, "#", "")
    HexToLong = RGB(CLng("&H" & Mid$(v, 1, 2)), CLng("&H" & Mid$(v, 3, 2)), CLng("&H" & Mid$(v, 5, 2)))
End Function

Private Function SafeSheetName(ByVal rawName As String) As String
    Dim s As String
    s = Trim(rawName)
    s = Replace(s, ":", " ")
    s = Replace(s, "\", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, "?", " ")
    s = Replace(s, "*", " ")
    s = Replace(s, "[", " ")
    s = Replace(s, "]", " ")
    If Len(s) = 0 Then s = "推荐材料"
    If Len(s) > 31 Then s = Left$(s, 31)
    SafeSheetName = s
End Function

Private Function UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String, ByVal currentSheet As Worksheet) As String
    Dim candidate As String
    candidate = baseName
    
    Dim i As Long: i = 1
    Do While SheetNameExists(wb, candidate, currentSheet)
        i = i + 1
        candidate = Left$(baseName, 28) & "_" & CStr(i)
    Loop
    
    UniqueSheetName = candidate
End Function

Private Function SheetNameExists(ByVal wb As Workbook, ByVal sheetName As String, ByVal currentSheet As Worksheet) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not ws Is currentSheet Then
            If ws.Name = sheetName Then
                SheetNameExists = True
                Exit Function
            End If
        End If
    Next ws
End Function
