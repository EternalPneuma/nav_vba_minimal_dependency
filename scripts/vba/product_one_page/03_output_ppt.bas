' OnePage03：按一页通版本配置生成两份PPT，不生成PDF

Option Explicit

Private Const TARGET_FILE_SUFFIX As String = "-产品一页通.xlsx"
Private Const TARGET_FILE_PATTERN As String = "*-产品一页通.xlsx"
Private Const PPTX_EXTENSION As String = ".pptx"

Private Const CHART_OBJECT_NAME As String = "chart_产品一页通"
Private Const ANCHOR_CHART_0 As String = "chart_000"
Private Const ANCHOR_CHART_1 As String = "chart_001"
Private Const ANCHOR_CHART_2 As String = "chart_002"
Private Const ANCHOR_TABLE As String = "table_000"

Private Const SLIDE1_NAV_TABLE_ROW As Long = 4
Private Const SLIDE1_NAV_TABLE_COL As Long = 2
Private Const SHEET_SOURCE_NAV As String = "上层产品净值数据(181)"
Private Const COL_SOURCE_CODE As String = "信托计划代码"
Private Const COL_SOURCE_DATE As String = "日期"
Private Const COL_SOURCE_ASSET_NAV As String = "资产净值"
Private Const COL_PRODUCT_CODE As String = "B"

Private Const PP_SAVE_AS_OPEN_XML_PRESENTATION As Long = 24
Private Const PP_PASTE_OLE_OBJECT As Long = 10
Private Const MSO_FALSE As Long = 0
Private Const MSO_EMBEDDED_OLE_OBJECT As Long = 7

Public Sub OnePage03_ExportPpt()
    Dim appCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    appCalc = Application.Calculation

    Dim wbCharts As Workbook
    Dim chartWorkbookWasOpen As Boolean
    Dim pptApp As Object
    Dim pptWasRunning As Boolean
    Dim pres As Object
    Dim currentStep As String
    Dim tempPaths As New Collection
    Dim finalPaths As New Collection

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    currentStep = "执行一页通静态预检"
    OnePage00_RequireStaticReady

    currentStep = "执行一页通数据预检"
    OnePage_RequireDataReady
    Dim baselineDate As Date
    baselineDate = OnePage_GetBaselineDate()
    Dim outputDateText As String
    outputDateText = Format$(baselineDate, "yyyymmdd")

    currentStep = "查找产品一页通图表工作簿"
    Dim chartWorkbookFile As String
    Dim chartWorkbookPath As String
    chartWorkbookFile = FindLatestOnePageWorkbook(outputDateText)
    chartWorkbookPath = ThisWorkbook.Path & Application.PathSeparator & chartWorkbookFile

    currentStep = "打开产品一页通图表工作簿"
    On Error Resume Next
    Set wbCharts = Workbooks(chartWorkbookFile)
    On Error GoTo CleanFail
    If wbCharts Is Nothing Then
        Set wbCharts = Workbooks.Open(chartWorkbookPath, ReadOnly:=True)
    Else
        chartWorkbookWasOpen = True
    End If

    currentStep = "启动PowerPoint"
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application")
    Else
        pptWasRunning = True
    End If
    On Error GoTo CleanFail
    If pptApp Is Nothing Then Err.Raise vbObjectError + 5301, , "无法启动 PowerPoint，请确认本机已安装 PowerPoint。"

    Dim definitions As Collection
    Set definitions = OnePage_GetEnabledDefinitions()
    Dim definition As Object
    For Each definition In definitions
        Dim versionName As String
        versionName = OnePage_DefinitionText(definition, "版本名称")

        currentStep = "生成一页通版本：" & versionName
        Dim templatePath As String
        templatePath = ThisWorkbook.Path & Application.PathSeparator & OnePage_DefinitionText(definition, "模板文件")
        Set pres = pptApp.Presentations.Open(templatePath, MSO_FALSE, MSO_FALSE, MSO_FALSE)

        ReplaceConfiguredChart pres, wbCharts, ANCHOR_CHART_0, OnePage_DefinitionText(definition, "基准产品代码")
        ReplaceConfiguredChart pres, wbCharts, ANCHOR_CHART_1, OnePage_DefinitionText(definition, "顶层产品1代码")
        ReplaceConfiguredChart pres, wbCharts, ANCHOR_CHART_2, OnePage_DefinitionText(definition, "顶层产品2代码")
        UpdateVersionNavTable pres, definition, baselineDate

        Dim finalPath As String
        Dim tempPath As String
        finalPath = ThisWorkbook.Path & Application.PathSeparator & _
                    OnePage_DefinitionText(definition, "输出文件前缀") & outputDateText & PPTX_EXTENSION
        tempPath = ThisWorkbook.Path & Application.PathSeparator & "~一页通临时-" & _
                   Format$(definition("输出顺序"), "000") & "-" & outputDateText & PPTX_EXTENSION
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath

        pres.SaveAs tempPath, PP_SAVE_AS_OPEN_XML_PRESENTATION
        pres.Close
        Set pres = Nothing
        tempPaths.Add tempPath
        finalPaths.Add finalPath
    Next definition

    currentStep = "重开并验证PPT中的OLE图表"
    Dim i As Long
    For i = 1 To tempPaths.Count
        Set pres = pptApp.Presentations.Open(CStr(tempPaths(i)), MSO_FALSE, MSO_FALSE, MSO_FALSE)
        ValidateEmbeddedChart pres, ANCHOR_CHART_0
        ValidateEmbeddedChart pres, ANCHOR_CHART_1
        ValidateEmbeddedChart pres, ANCHOR_CHART_2
        ValidateNamedTable pres, ANCHOR_TABLE
        pres.Close
        Set pres = Nothing
    Next i

    currentStep = "同时替换正式PPT"
    CommitOutputFiles tempPaths, finalPaths

    If Not pptWasRunning Then pptApp.Quit
    Set pptApp = Nothing
    If Not chartWorkbookWasOpen Then wbCharts.Close SaveChanges:=False
    Set wbCharts = Nothing

    RestoreApplicationState appCalc, oldDisplayAlerts, oldEnableEvents, oldScreenUpdating

    Dim outputSummary As String
    For i = 1 To finalPaths.Count
        outputSummary = outputSummary & CStr(finalPaths(i)) & vbCrLf
    Next i
    ReportPipeline_MsgBox "产品一页通PPT导出完成" & vbCrLf & vbCrLf & _
           "基准日期：" & Format$(baselineDate, "yyyy-mm-dd") & vbCrLf & _
           "图表工作簿：" & chartWorkbookFile & vbCrLf & vbCrLf & _
           "输出文件：" & vbCrLf & outputSummary, vbInformation, "产品一页通"
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
    If Not pptApp Is Nothing Then If Not pptWasRunning Then pptApp.Quit
    If Not wbCharts Is Nothing Then If Not chartWorkbookWasOpen Then wbCharts.Close SaveChanges:=False
    DeleteFiles tempPaths
    RestoreApplicationState appCalc, oldDisplayAlerts, oldEnableEvents, oldScreenUpdating
    On Error GoTo 0

    If Len(failDescription) = 0 Then failDescription = "未知错误"
    If Len(failStep) = 0 Then failStep = "未记录"
    ReportPipeline_MsgBox "产品一页通PPT导出失败" & vbCrLf & vbCrLf & _
           "错误信息：" & failDescription & vbCrLf & _
           "错误号：" & failNumber & vbCrLf & _
           "步骤：" & failStep, vbCritical, "产品一页通"
End Sub

' 兼容旧宏入口；不再生成PDF。
Public Sub OnePage03_ExportPptPdf()
    OnePage03_ExportPpt
End Sub

Private Sub ReplaceConfiguredChart(ByVal pres As Object, ByVal wbCharts As Workbook, _
                                   ByVal anchorName As String, ByVal productCode As String)
    Dim anchorShape As Object
    Set anchorShape = FindUniqueShapeByName(pres, anchorName)
    Dim sourceChart As ChartObject
    Set sourceChart = FindProductChartByCode(wbCharts, productCode)
    ReplaceShapeWithEmbeddedChart anchorShape.Parent, anchorShape, sourceChart, anchorName
End Sub

Private Function FindLatestOnePageWorkbook(ByVal requiredDateText As String) As String
    Dim expectedFile As String
    expectedFile = requiredDateText & TARGET_FILE_SUFFIX
    If Len(Dir$(ThisWorkbook.Path & Application.PathSeparator & expectedFile)) = 0 Then
        Err.Raise vbObjectError + 5311, , "未找到与基准日期一致的产品一页通图表工作簿：" & expectedFile
    End If
    FindLatestOnePageWorkbook = expectedFile
End Function

Private Function FindProductChartByCode(ByVal wbCharts As Workbook, ByVal productCode As String) As ChartObject
    Dim ws As Worksheet
    For Each ws In wbCharts.Worksheets
        If StrComp(NormalizeText(ws.Range(COL_PRODUCT_CODE & "2").Value), productCode, vbTextCompare) = 0 Then
            Dim co As ChartObject
            On Error Resume Next
            Set co = ws.ChartObjects(CHART_OBJECT_NAME)
            On Error GoTo 0
            If co Is Nothing Then Err.Raise vbObjectError + 5322, , "产品sheet中未找到图表：" & ws.Name
            Set FindProductChartByCode = co
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 5321, , "产品一页通工作簿中未找到产品sheet：" & productCode
End Function

Private Sub ReplaceShapeWithEmbeddedChart(ByVal sld As Object, ByVal oldShape As Object, _
                                          ByVal sourceChart As ChartObject, ByVal newShapeName As String)
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
    Dim pasteDescription As String
    Dim attempt As Long
    For attempt = 1 To 3
        Set pastedShapes = Nothing
        sourceChart.Parent.Parent.Activate
        sourceChart.Parent.Activate
        sourceChart.Activate
        sourceChart.Chart.Refresh
        DoEvents
        sourceChart.Copy
        DoEvents

        On Error Resume Next
        Set pastedShapes = sld.Shapes.PasteSpecial(DataType:=PP_PASTE_OLE_OBJECT, Link:=MSO_FALSE)
        If pastedShapes Is Nothing And attempt = 3 Then Set pastedShapes = sld.Shapes.Paste
        If Err.Number <> 0 Then pasteDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If Not pastedShapes Is Nothing Then Exit For
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next attempt

    If pastedShapes Is Nothing Then Err.Raise vbObjectError + 5331, , "图表OLE嵌入失败：" & newShapeName & "；" & pasteDescription
    If pastedShapes.Count = 0 Then Err.Raise vbObjectError + 5332, , "图表OLE粘贴结果为空：" & newShapeName

    Dim newShape As Object
    Set newShape = pastedShapes(1)
    If newShape.Type <> MSO_EMBEDDED_OLE_OBJECT Then
        newShape.Delete
        Err.Raise vbObjectError + 5333, , "图表粘贴结果不是嵌入OLE对象：" & newShapeName
    End If

    With newShape
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        On Error Resume Next
        .Rotation = rotationVal
        Err.Clear
        On Error GoTo 0
        .Name = newShapeName & "_embedded"
    End With
    oldShape.Delete
    newShape.Name = newShapeName
End Sub

Private Sub UpdateVersionNavTable(ByVal pres As Object, ByVal definition As Object, ByVal baselineDate As Date)
    Dim tableShape As Object
    Set tableShape = FindUniqueShapeByName(pres, ANCHOR_TABLE)
    If tableShape.HasTable = 0 Then Err.Raise vbObjectError + 5340, , ANCHOR_TABLE & "不是表格。"
    If tableShape.Table.Rows.Count < SLIDE1_NAV_TABLE_ROW Or tableShape.Table.Columns.Count < SLIDE1_NAV_TABLE_COL Then
        Err.Raise vbObjectError + 5341, , ANCHOR_TABLE & "尺寸不足，无法更新净值合计。"
    End If

    Dim code1 As String
    Dim code2 As String
    code1 = OnePage_DefinitionText(definition, "顶层产品1代码")
    code2 = OnePage_DefinitionText(definition, "顶层产品2代码")
    Dim totalNav As Double
    totalNav = FindAssetNavOnDate(code1, baselineDate) + FindAssetNavOnDate(code2, baselineDate)
    tableShape.Table.Cell(SLIDE1_NAV_TABLE_ROW, SLIDE1_NAV_TABLE_COL).Shape.TextFrame.TextRange.Text = _
        "净值合计" & Format$(totalNav / 100000000#, "0.00") & "亿元"
End Sub

Private Function FindAssetNavOnDate(ByVal productCode As String, ByVal baselineDate As Date) As Double
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SOURCE_NAV)
    Dim headers As Object
    Set headers = BuildHeaderMap(ws, 1)
    If Not headers.Exists(COL_SOURCE_CODE) Or Not headers.Exists(COL_SOURCE_DATE) Or Not headers.Exists(COL_SOURCE_ASSET_NAV) Then
        Err.Raise vbObjectError + 5342, , SHEET_SOURCE_NAV & "缺少信托计划代码、日期或资产净值字段。"
    End If

    Dim r As Long
    Dim parsedDate As Date
    For r = 2 To LastUsedRow(ws)
        If StrComp(NormalizeText(ws.Cells(r, CLng(headers(COL_SOURCE_CODE))).Value), productCode, vbTextCompare) = 0 Then
            If TryReadDate(ws.Cells(r, CLng(headers(COL_SOURCE_DATE))).Value, parsedDate) Then
                If parsedDate = baselineDate And IsNumeric(ws.Cells(r, CLng(headers(COL_SOURCE_ASSET_NAV))).Value) Then
                    If CDbl(ws.Cells(r, CLng(headers(COL_SOURCE_ASSET_NAV))).Value) > 0 Then
                        FindAssetNavOnDate = CDbl(ws.Cells(r, CLng(headers(COL_SOURCE_ASSET_NAV))).Value)
                        Exit Function
                    End If
                End If
            End If
        End If
    Next r
    Err.Raise vbObjectError + 5343, , Format$(baselineDate, "yyyy-mm-dd") & "缺少有效资产净值：" & productCode
End Function

Private Sub ValidateEmbeddedChart(ByVal pres As Object, ByVal shapeName As String)
    Dim shp As Object
    Set shp = FindUniqueShapeByName(pres, shapeName)
    If shp.Type <> MSO_EMBEDDED_OLE_OBJECT Then Err.Raise vbObjectError + 5350, , "PPT对象不是嵌入OLE图表：" & shapeName
End Sub

Private Sub ValidateNamedTable(ByVal pres As Object, ByVal shapeName As String)
    Dim shp As Object
    Set shp = FindUniqueShapeByName(pres, shapeName)
    If shp.HasTable = 0 Then Err.Raise vbObjectError + 5351, , "PPT对象不是表格：" & shapeName
End Sub

Private Function FindUniqueShapeByName(ByVal pres As Object, ByVal shapeName As String) As Object
    Dim foundShape As Object
    Dim foundCount As Long
    Dim sld As Object
    For Each sld In pres.Slides
        CountNamedShapes sld.Shapes, shapeName, foundCount, foundShape
    Next sld
    If foundCount <> 1 Then Err.Raise vbObjectError + 5352, , "PPT中的" & shapeName & "数量必须为1，实际为" & foundCount & "。"
    Set FindUniqueShapeByName = foundShape
End Function

Private Sub CountNamedShapes(ByVal shapesCollection As Object, ByVal targetName As String, _
                             ByRef foundCount As Long, ByRef foundShape As Object)
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

Private Sub CommitOutputFiles(ByVal tempPaths As Collection, ByVal finalPaths As Collection)
    Dim backupPaths As New Collection
    Dim i As Long
    On Error GoTo CommitFail

    For i = 1 To finalPaths.Count
        Dim backupPath As String
        backupPath = CStr(finalPaths(i)) & ".bak"
        If Len(Dir$(backupPath)) > 0 Then Kill backupPath
        If Len(Dir$(CStr(finalPaths(i)))) > 0 Then Name CStr(finalPaths(i)) As backupPath
        backupPaths.Add backupPath
    Next i

    For i = 1 To tempPaths.Count
        Name CStr(tempPaths(i)) As CStr(finalPaths(i))
    Next i

    DeleteFiles backupPaths
    Exit Sub

CommitFail:
    Dim description As String
    description = Err.Description
    On Error Resume Next
    For i = 1 To finalPaths.Count
        If Len(Dir$(CStr(finalPaths(i)))) > 0 Then Kill CStr(finalPaths(i))
    Next i
    For i = 1 To backupPaths.Count
        If Len(Dir$(CStr(backupPaths(i)))) > 0 Then Name CStr(backupPaths(i)) As CStr(finalPaths(i))
    Next i
    On Error GoTo 0
    Err.Raise vbObjectError + 5360, , "同时替换正式PPT失败：" & description
End Sub

Private Sub DeleteFiles(ByVal paths As Collection)
    Dim item As Variant
    For Each item In paths
        If Len(Dir$(CStr(item))) > 0 Then Kill CStr(item)
    Next item
End Sub

Private Sub RestoreApplicationState(ByVal calculationMode As XlCalculation, ByVal displayAlerts As Boolean, _
                                    ByVal enableEvents As Boolean, ByVal screenUpdating As Boolean)
    Application.Calculation = calculationMode
    Application.DisplayAlerts = displayAlerts
    Application.EnableEvents = enableEvents
    Application.ScreenUpdating = screenUpdating
End Sub

Private Function BuildHeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Dim c As Long
    For c = 1 To LastUsedColumn(ws)
        Dim headerText As String
        headerText = NormalizeText(ws.Cells(headerRow, c).Value)
        If Len(headerText) > 0 Then If Not result.Exists(headerText) Then result(headerText) = c
    Next c
    Set BuildHeaderMap = result
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
