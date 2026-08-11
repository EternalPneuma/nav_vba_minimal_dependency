' OnePage03_ExportPptPdf：将产品一页通图表嵌入 PPT，并导出 PDF

Option Explicit

Private Const TARGET_FILE_SUFFIX As String = "-产品一页通.xlsx"
Private Const TARGET_FILE_PATTERN As String = "*-产品一页通.xlsx"

Private Const PPT_TEMPLATE_FILE As String = "assets\产品一页通-交鑫致远-模板.pptx"
Private Const OUTPUT_FILE_PREFIX As String = "产品一页通-交鑫致远-"
Private Const PPTX_EXTENSION As String = ".pptx"
Private Const PDF_EXTENSION As String = ".pdf"

Private Const CHART_OBJECT_NAME As String = "chart_产品一页通"
Private Const CHART_ANCHOR_NAME As String = "chart_产品一页通"

Private Const SLIDE1_TABLE_NAME As String = "表格 17"
Private Const SLIDE1_NAV_TABLE_ROW As Long = 4
Private Const SLIDE1_NAV_TABLE_COL As Long = 2
Private Const SHEET_SOURCE_NAV As String = "上层产品净值数据(181)"
Private Const COL_SOURCE_CODE As Long = 9
Private Const COL_SOURCE_DATE As Long = 2
Private Const COL_SOURCE_ASSET_NAV As Long = 3

Private Const PRODUCT_CODE_101 As String = "P83600"
Private Const PRODUCT_CODE_102 As String = "P83800"

Private Const COL_PRODUCT_CODE As String = "B"

Private Const PP_SAVE_AS_OPEN_XML_PRESENTATION As Long = 24
Private Const PP_SAVE_AS_PDF As Long = 32
Private Const PP_PASTE_OLE_OBJECT As Long = 10
Private Const MSO_FALSE As Long = 0
Private Const MSO_LINKED_PICTURE As Long = 11
Private Const MSO_PICTURE As Long = 13

Public Sub OnePage03_ExportPptPdf()
    Dim appCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    appCalc = Application.Calculation

    Dim wbCharts As Workbook
    Dim wasOpen As Boolean
    Dim pptApp As Object
    Dim pptWasRunning As Boolean
    Dim pres As Object
    Dim currentStep As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    currentStep = "查找产品一页通工作簿"
    Dim chartWorkbookFile As String
    Dim chartWorkbookPath As String
    Dim outputDateText As String
    chartWorkbookFile = FindLatestOnePageWorkbook()
    chartWorkbookPath = ThisWorkbook.Path & Application.PathSeparator & chartWorkbookFile
    outputDateText = Left$(chartWorkbookFile, 8)

    currentStep = "打开产品一页通工作簿"
    On Error Resume Next
    Set wbCharts = Workbooks(chartWorkbookFile)
    On Error GoTo CleanFail
    If wbCharts Is Nothing Then
        Set wbCharts = Workbooks.Open(chartWorkbookPath)
    Else
        wasOpen = True
    End If

    Dim targetProducts As Variant
    targetProducts = Array("OA4400", "P83600", "P83800")

    currentStep = "读取产品一页通PPT模板"
    Dim templateFile As String
    Dim templatePath As String
    templateFile = FindPptTemplate()
    templatePath = ThisWorkbook.Path & Application.PathSeparator & templateFile

    currentStep = "启动PowerPoint"
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application")
    Else
        pptWasRunning = True
    End If
    On Error GoTo CleanFail
    If pptApp Is Nothing Then
        Err.Raise vbObjectError + 5301, , "无法启动 PowerPoint，请确认本机已安装 PowerPoint。"
    End If

    currentStep = "打开PPT模板"
    Set pres = pptApp.Presentations.Open(templatePath, MSO_FALSE, MSO_FALSE, MSO_FALSE)

    currentStep = "替换PPT中的三个图表"
    Dim anchorShape As Object
    Dim sourceChart As ChartObject
    Dim i As Long
    For i = LBound(targetProducts) To UBound(targetProducts)
        Set anchorShape = FindShapeByName(pres, CHART_ANCHOR_NAME & "_" & CStr(targetProducts(i)))
        If anchorShape Is Nothing Then
            Err.Raise vbObjectError + 5302, , "PPT模板中未找到图表对象：" & CHART_ANCHOR_NAME & "_" & CStr(targetProducts(i))
        End If

        Set sourceChart = FindProductChartByCode(wbCharts, CStr(targetProducts(i)))
        ReplaceShapeWithEmbeddedChart anchorShape.Parent, anchorShape, sourceChart, CHART_ANCHOR_NAME & "_" & CStr(targetProducts(i))
    Next i

    currentStep = "更新幻灯片1中的资产净值表格"
    UpdateSlide1NavTable pres, wbCharts

    currentStep = "另存PPT并导出PDF"
    Dim outputPptPath As String
    Dim outputPdfPath As String
    outputPptPath = BuildDatedPptOutputPath(outputDateText)
    outputPdfPath = Left$(outputPptPath, Len(outputPptPath) - Len(PPTX_EXTENSION)) & PDF_EXTENSION

    pres.SaveAs outputPptPath, PP_SAVE_AS_OPEN_XML_PRESENTATION
    pres.SaveAs outputPdfPath, PP_SAVE_AS_PDF
    pres.Close
    Set pres = Nothing

    If Not pptWasRunning Then pptApp.Quit
    Set pptApp = Nothing

    If Not wasOpen Then wbCharts.Close SaveChanges:=False
    Set wbCharts = Nothing

    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    ReportPipeline_MsgBox "产品一页通PPT/PDF导出完成" & vbCrLf & vbCrLf & _
           "处理结果：" & vbCrLf & _
           "图表工作簿：" & chartWorkbookFile & vbCrLf & vbCrLf & _
           "输出文件：" & vbCrLf & _
           "PPT输出：" & outputPptPath & vbCrLf & _
           "PDF输出：" & outputPdfPath, vbInformation, "产品一页通"
    Exit Sub

CleanFail:
    Dim failNumber As Long
    Dim failDescription As String
    Dim failStep As String
    failNumber = Err.Number
    failDescription = Err.Description
    failStep = currentStep

    On Error Resume Next
    If Not pres Is Nothing Then pres.Close
    If Not pptApp Is Nothing Then
        If Not pptWasRunning Then pptApp.Quit
    End If
    If Not wbCharts Is Nothing Then
        If Not wasOpen Then wbCharts.Close SaveChanges:=False
    End If

    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If Len(failDescription) = 0 Then failDescription = "未知错误"
    If Len(failStep) = 0 Then failStep = "未记录"

    ReportPipeline_MsgBox "产品一页通PPT/PDF导出失败" & vbCrLf & vbCrLf & _
           "错误信息：" & failDescription & vbCrLf & _
           "错误号：" & failNumber & vbCrLf & _
           "步骤：" & failStep, vbCritical, "产品一页通"
End Sub

Private Function FindLatestOnePageWorkbook() As String
    Dim fileName As String
    Dim latestFile As String
    Dim latestKey As String
    Dim dateText As String

    fileName = Dir$(ThisWorkbook.Path & Application.PathSeparator & TARGET_FILE_PATTERN)
    Do While Len(fileName) > 0
        dateText = Left$(fileName, 8)
        If Len(dateText) = 8 And IsNumeric(dateText) Then
            If Right$(fileName, Len(TARGET_FILE_SUFFIX)) = TARGET_FILE_SUFFIX Then
                If dateText > latestKey Then
                    latestKey = dateText
                    latestFile = fileName
                End If
            End If
        End If
        fileName = Dir$()
    Loop

    If Len(latestFile) = 0 Then
        Err.Raise vbObjectError + 5311, , "未找到 yyyymmdd-产品一页通.xlsx，请先运行 OnePage01_ExportChartData 和 OnePage02_GenerateCharts。"
    End If

    FindLatestOnePageWorkbook = latestFile
End Function

Private Function FindPptTemplate() As String
    If Len(Dir$(ThisWorkbook.Path & Application.PathSeparator & PPT_TEMPLATE_FILE)) = 0 Then
        Err.Raise vbObjectError + 5312, , "未找到产品一页通PPT模板：" & PPT_TEMPLATE_FILE
    End If

    FindPptTemplate = PPT_TEMPLATE_FILE
End Function

Private Function FindProductChartByCode(ByVal wbCharts As Workbook, ByVal productCode As String) As ChartObject
    Dim ws As Worksheet
    Set ws = FindProductWorksheet(wbCharts, productCode)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 5321, , "产品一页通工作簿中未找到产品sheet：" & productCode
    End If

    Dim co As ChartObject
    Set co = FindProductChartObject(ws)
    If co Is Nothing Then
        Err.Raise vbObjectError + 5322, , "产品sheet中未找到图表：" & ws.Name
    End If

    Set FindProductChartByCode = co
End Function

Private Function FindProductWorksheet(ByVal wbCharts As Workbook, ByVal productCode As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wbCharts.Worksheets
        If StrComp(NormalizeText(ws.Range(COL_PRODUCT_CODE & "2").Value), productCode, vbTextCompare) = 0 Then
            Set FindProductWorksheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function FindProductChartObject(ByVal ws As Worksheet) As ChartObject
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        If StrComp(co.Name, CHART_OBJECT_NAME, vbTextCompare) = 0 Then
            Set FindProductChartObject = co
            Exit Function
        End If
    Next co

    If ws.ChartObjects.Count > 0 Then Set FindProductChartObject = ws.ChartObjects(1)
End Function

Private Function FindShapeByName(ByVal pres As Object, ByVal shapeName As String) As Object
    Dim sld As Object
    Dim foundShape As Object

    For Each sld In pres.Slides
        Set foundShape = FindShapeByNameInShapes(sld.Shapes, shapeName)
        If Not foundShape Is Nothing Then
            Set FindShapeByName = foundShape
            Exit Function
        End If
    Next sld
End Function

Private Function FindShapeByNameInShapes(ByVal shapesCollection As Object, ByVal shapeName As String) As Object
    Dim shp As Object
    Dim foundShape As Object

    For Each shp In shapesCollection
        If StrComp(shp.Name, shapeName, vbTextCompare) = 0 Then
            Set FindShapeByNameInShapes = shp
            Exit Function
        End If

        Set foundShape = FindShapeByNameInGroup(shp, shapeName)
        If Not foundShape Is Nothing Then
            Set FindShapeByNameInShapes = foundShape
            Exit Function
        End If
    Next shp
End Function

Private Function FindShapeByNameInGroup(ByVal shp As Object, ByVal shapeName As String) As Object
    Dim groupItems As Object
    On Error Resume Next
    Set groupItems = shp.GroupItems
    If Err.Number <> 0 Then
        Err.Clear
        Set groupItems = Nothing
    End If
    On Error GoTo 0

    If Not groupItems Is Nothing Then
        Set FindShapeByNameInGroup = FindShapeByNameInShapes(groupItems, shapeName)
    End If
End Function

Private Sub ReplaceShapeWithEmbeddedChart(ByVal sld As Object, ByVal oldShape As Object, ByVal sourceChart As ChartObject, ByVal newShapeName As String)
    Dim leftPos As Single
    Dim topPos As Single
    Dim widthVal As Single
    Dim heightVal As Single
    Dim rotationVal As Single

    leftPos = oldShape.Left
    topPos = oldShape.Top
    widthVal = oldShape.Width
    heightVal = oldShape.Height
    rotationVal = oldShape.Rotation

    Dim pastedShapes As Object
    Set pastedShapes = PasteChartWithoutRasterizing(sld, sourceChart, newShapeName)

    If pastedShapes Is Nothing Then
        Err.Raise vbObjectError + 5331, , "图表嵌入PPT失败：" & newShapeName
    End If
    If pastedShapes.Count = 0 Then
        Err.Raise vbObjectError + 5331, , "图表嵌入PPT失败：" & newShapeName
    End If

    Dim newShape As Object
    Set newShape = pastedShapes(1)
    With newShape
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        On Error Resume Next
        .Rotation = rotationVal
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        .Name = newShapeName & "_embedded"
    End With

    oldShape.Delete

    newShape.Name = newShapeName
End Sub

Private Function PasteChartWithoutRasterizing(ByVal sld As Object, ByVal sourceChart As ChartObject, ByVal chartName As String) As Object
    Dim pastedShapes As Object
    Dim pasteErrDescription As String

    sourceChart.Copy

    On Error Resume Next
    Set pastedShapes = sld.Shapes.PasteSpecial(DataType:=PP_PASTE_OLE_OBJECT, Link:=MSO_FALSE)
    If Err.Number <> 0 Then
        pasteErrDescription = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    If pastedShapes Is Nothing Then
        sourceChart.Copy
        On Error Resume Next
        Set pastedShapes = sld.Shapes.Paste
        If Err.Number <> 0 Then
            If Len(pasteErrDescription) > 0 Then pasteErrDescription = pasteErrDescription & "；"
            pasteErrDescription = pasteErrDescription & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If

    If pastedShapes Is Nothing Then
        Err.Raise vbObjectError + 5331, , "图表嵌入PPT失败：" & chartName & "；" & pasteErrDescription
    End If
    If pastedShapes.Count = 0 Then
        Err.Raise vbObjectError + 5331, , "图表嵌入PPT失败：" & chartName
    End If

    If IsRasterPictureShape(pastedShapes(1)) Then
        pastedShapes(1).Delete
        Err.Raise vbObjectError + 5332, , "图表粘贴结果仍为图片，已取消导出：" & chartName
    End If

    Set PasteChartWithoutRasterizing = pastedShapes
End Function

Private Function IsRasterPictureShape(ByVal shp As Object) As Boolean
    On Error Resume Next
    IsRasterPictureShape = (shp.Type = MSO_PICTURE Or shp.Type = MSO_LINKED_PICTURE)
    On Error GoTo 0
End Function

Private Function BuildDatedPptOutputPath(ByVal outputDateText As String) As String
    BuildDatedPptOutputPath = ThisWorkbook.Path & Application.PathSeparator & _
                              OUTPUT_FILE_PREFIX & outputDateText & PPTX_EXTENSION
End Function

Private Sub UpdateSlide1NavTable(ByVal pres As Object, ByVal wbCharts As Workbook)
    Dim tblShp As Object
    Set tblShp = FindTableShapeByName(pres, SLIDE1_TABLE_NAME)
    If tblShp Is Nothing Then
        Err.Raise vbObjectError + 5340, , "未在PPT第1页找到可更新表格：" & SLIDE1_TABLE_NAME & _
            "；需要至少" & SLIDE1_NAV_TABLE_ROW & "行" & SLIDE1_NAV_TABLE_COL & "列。" & _
            DescribeSlideTableShapes(pres.Slides(1).Shapes)
    End If

    Dim tbl As Object
    Set tbl = tblShp.Table

    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Worksheets(SHEET_SOURCE_NAV)

    Dim lastRow As Long
    lastRow = LastUsedRow(wsSource)

    Dim latestDate101 As Long
    Dim latestDate102 As Long
    Dim latestNav101 As Double
    Dim latestNav102 As Double

    Dim i As Long
    Dim code As String
    Dim cellDate As Variant
    Dim cellNav As Variant

    For i = 2 To lastRow
        code = NormalizeText(wsSource.Cells(i, COL_SOURCE_CODE).Value)

        If Len(code) = 0 Then GoTo NextRow
        If StrComp(code, PRODUCT_CODE_101, vbTextCompare) <> 0 And StrComp(code, PRODUCT_CODE_102, vbTextCompare) <> 0 Then GoTo NextRow

        cellDate = wsSource.Cells(i, COL_SOURCE_DATE).Value
        cellNav = wsSource.Cells(i, COL_SOURCE_ASSET_NAV).Value

        If Not IsNumeric(cellDate) Then GoTo NextRow
        If Not IsNumeric(cellNav) Then GoTo NextRow
        If CDbl(cellNav) <= 0 Then GoTo NextRow

        Dim d As Long
        d = CLng(cellDate)

        If StrComp(code, PRODUCT_CODE_101, vbTextCompare) = 0 Then
            If d > latestDate101 Then
                latestDate101 = d
                latestNav101 = CDbl(cellNav)
            End If
        ElseIf StrComp(code, PRODUCT_CODE_102, vbTextCompare) = 0 Then
            If d > latestDate102 Then
                latestDate102 = d
                latestNav102 = CDbl(cellNav)
            End If
        End If

NextRow:
    Next i

    If latestDate101 = 0 And latestDate102 = 0 Then
        Err.Raise vbObjectError + 5343, , "未在源数据中找到" & PRODUCT_CODE_101 & "或" & PRODUCT_CODE_102 & "的资产净值数据。"
    End If

    Dim totalNav As Double
    totalNav = latestNav101 + latestNav102

    Dim navText As String
    navText = "净值合计" & Format$(totalNav / 100000000#, "0.00") & "亿元"

    tbl.Cell(SLIDE1_NAV_TABLE_ROW, SLIDE1_NAV_TABLE_COL).Shape.TextFrame.TextRange.Text = navText
End Sub

Private Function FindTableShapeByName(ByVal pres As Object, ByVal tableName As String) As Object
    Dim shp As Object
    Set shp = FindShapeByName(pres, tableName)
    If Not shp Is Nothing Then
        If TableShapeCanUpdate(shp, SLIDE1_NAV_TABLE_ROW, SLIDE1_NAV_TABLE_COL) Then
            Set FindTableShapeByName = shp
            Exit Function
        End If
    End If

    Set FindTableShapeByName = FindUpdateableTableShapeInShapes(pres.Slides(1).Shapes, SLIDE1_NAV_TABLE_ROW, SLIDE1_NAV_TABLE_COL)
End Function

Private Function FindUpdateableTableShapeInShapes(ByVal shapesCollection As Object, ByVal minRows As Long, ByVal minCols As Long) As Object
    Dim shp As Object
    Dim foundShape As Object

    For Each shp In shapesCollection
        If TableShapeCanUpdate(shp, minRows, minCols) Then
            Set FindUpdateableTableShapeInShapes = shp
            Exit Function
        End If

        Set foundShape = FindUpdateableTableShapeInGroup(shp, minRows, minCols)
        If Not foundShape Is Nothing Then
            Set FindUpdateableTableShapeInShapes = foundShape
            Exit Function
        End If
    Next shp
End Function

Private Function FindUpdateableTableShapeInGroup(ByVal shp As Object, ByVal minRows As Long, ByVal minCols As Long) As Object
    Dim groupItems As Object
    On Error Resume Next
    Set groupItems = shp.GroupItems
    If Err.Number <> 0 Then
        Err.Clear
        Set groupItems = Nothing
    End If
    On Error GoTo 0

    If Not groupItems Is Nothing Then
        Set FindUpdateableTableShapeInGroup = FindUpdateableTableShapeInShapes(groupItems, minRows, minCols)
    End If
End Function

Private Function TableShapeCanUpdate(ByVal shp As Object, ByVal minRows As Long, ByVal minCols As Long) As Boolean
    On Error Resume Next
    TableShapeCanUpdate = (shp.HasTable <> 0 And shp.Table.Rows.Count >= minRows And shp.Table.Columns.Count >= minCols)
    If Err.Number <> 0 Then
        Err.Clear
        TableShapeCanUpdate = False
    End If
    On Error GoTo 0
End Function

Private Function ShapeHasTable(ByVal shp As Object) As Boolean
    On Error Resume Next
    ShapeHasTable = (shp.HasTable <> 0)
    If Err.Number <> 0 Then
        Err.Clear
        ShapeHasTable = False
    End If
    On Error GoTo 0
End Function

Private Function DescribeSlideTableShapes(ByVal shapesCollection As Object) As String
    Dim tableSummary As String
    tableSummary = CollectTableShapeDescriptions(shapesCollection)
    If Len(tableSummary) = 0 Then
        DescribeSlideTableShapes = " 当前第1页未识别到表格。"
    Else
        DescribeSlideTableShapes = " 当前第1页表格：" & tableSummary
    End If
End Function

Private Function CollectTableShapeDescriptions(ByVal shapesCollection As Object) As String
    Dim shp As Object
    Dim part As String
    Dim childPart As String

    For Each shp In shapesCollection
        If ShapeHasTable(shp) Then
            part = TableShapeDescription(shp)
            If Len(part) > 0 Then
                If Len(CollectTableShapeDescriptions) > 0 Then CollectTableShapeDescriptions = CollectTableShapeDescriptions & "；"
                CollectTableShapeDescriptions = CollectTableShapeDescriptions & part
            End If
        End If

        childPart = CollectTableDescriptionsInGroup(shp)
        If Len(childPart) > 0 Then
            If Len(CollectTableShapeDescriptions) > 0 Then CollectTableShapeDescriptions = CollectTableShapeDescriptions & "；"
            CollectTableShapeDescriptions = CollectTableShapeDescriptions & childPart
        End If
    Next shp
End Function

Private Function CollectTableDescriptionsInGroup(ByVal shp As Object) As String
    Dim groupItems As Object
    On Error Resume Next
    Set groupItems = shp.GroupItems
    If Err.Number <> 0 Then
        Err.Clear
        Set groupItems = Nothing
    End If
    On Error GoTo 0

    If Not groupItems Is Nothing Then
        CollectTableDescriptionsInGroup = CollectTableShapeDescriptions(groupItems)
    End If
End Function

Private Function TableShapeDescription(ByVal shp As Object) As String
    On Error Resume Next
    TableShapeDescription = shp.Name & "(" & shp.Table.Rows.Count & "行x" & shp.Table.Columns.Count & "列)"
    If Err.Number <> 0 Then
        Err.Clear
        TableShapeDescription = vbNullString
    End If
    On Error GoTo 0
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
