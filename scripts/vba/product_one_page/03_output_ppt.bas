' OnePage03_ExportPptPdf：将产品一页通图表导入 PPT，并导出 PDF

Option Explicit

Private Const TARGET_FILE_SUFFIX As String = "-产品一页通.xlsx"
Private Const TARGET_FILE_PATTERN As String = "*-产品一页通.xlsx"

Private Const PPT_TEMPLATE_FILE As String = "产品一页通-交鑫致远-模板.pptx"
Private Const OUTPUT_FILE_PREFIX As String = "产品一页通-交鑫致远-"
Private Const PPTX_EXTENSION As String = ".pptx"
Private Const PDF_EXTENSION As String = ".pdf"

Private Const CHART_OBJECT_NAME As String = "chart_产品一页通"
Private Const CHART_ANCHOR_NAME As String = "chart_产品一页通"

Private Const COL_PRODUCT_CODE As String = "B"

Private Const PP_SAVE_AS_OPEN_XML_PRESENTATION As Long = 24
Private Const PP_SAVE_AS_PDF As Long = 32
Private Const MSO_FALSE As Long = 0
Private Const MSO_TRUE As Long = -1

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
    Dim tempFolder As String
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

    currentStep = "导出临时图表图片"
    tempFolder = CreateTempImageFolder()

    Dim targetProducts As Variant
    targetProducts = Array("OA4400", "P83600", "P83800")

    Dim imagePaths(0 To 2) As String
    Dim i As Long
    For i = LBound(targetProducts) To UBound(targetProducts)
        imagePaths(i) = ExportProductChartImage(wbCharts, CStr(targetProducts(i)), tempFolder)
    Next i

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
    For i = LBound(targetProducts) To UBound(targetProducts)
        Set anchorShape = FindShapeByName(pres, CHART_ANCHOR_NAME & "_" & CStr(targetProducts(i)))
        If anchorShape Is Nothing Then
            Err.Raise vbObjectError + 5302, , "PPT模板中未找到图表对象：" & CHART_ANCHOR_NAME & "_" & CStr(targetProducts(i))
        End If
        ReplaceShapeWithPicture anchorShape.Parent, anchorShape, imagePaths(i), CHART_ANCHOR_NAME & "_" & CStr(targetProducts(i))
    Next i

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
    CleanTempImageFolder tempFolder

    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    MsgBox "产品一页通PPT/PDF导出完成" & vbCrLf & vbCrLf & _
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
    If Len(tempFolder) > 0 Then CleanTempImageFolder tempFolder

    Application.Calculation = appCalc
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If Len(failDescription) = 0 Then failDescription = "未知错误"
    If Len(failStep) = 0 Then failStep = "未记录"

    MsgBox "产品一页通PPT/PDF导出失败" & vbCrLf & vbCrLf & _
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

Private Function ExportProductChartImage(ByVal wbCharts As Workbook, ByVal productCode As String, ByVal tempFolder As String) As String
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

    Dim imagePath As String
    imagePath = tempFolder & Application.PathSeparator & productCode & ".png"
    If Not co.Chart.Export(Filename:=imagePath, FilterName:="PNG") Then
        Err.Raise vbObjectError + 5323, , "图表图片导出失败：" & productCode
    End If

    ExportProductChartImage = imagePath
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
    Dim shp As Object

    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            If StrComp(shp.Name, shapeName, vbTextCompare) = 0 Then
                Set FindShapeByName = shp
                Exit Function
            End If
        Next shp
    Next sld
End Function

Private Sub ReplaceShapeWithPicture(ByVal sld As Object, ByVal oldShape As Object, ByVal imagePath As String, ByVal newShapeName As String)
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

    oldShape.Delete

    Dim newShape As Object
    Set newShape = sld.Shapes.AddPicture(imagePath, MSO_FALSE, MSO_TRUE, leftPos, topPos, widthVal, heightVal)
    newShape.Name = newShapeName
    newShape.Rotation = rotationVal
End Sub

Private Function BuildDatedPptOutputPath(ByVal outputDateText As String) As String
    BuildDatedPptOutputPath = ThisWorkbook.Path & Application.PathSeparator & _
                              OUTPUT_FILE_PREFIX & outputDateText & PPTX_EXTENSION
End Function

Private Function CreateTempImageFolder() As String
    Dim baseFolder As String
    baseFolder = Environ$("TEMP")
    If Len(baseFolder) = 0 Then baseFolder = ThisWorkbook.Path

    Dim folderPath As String
    folderPath = baseFolder & Application.PathSeparator & "OnePageCharts_" & Format$(Now, "yyyymmdd_hhnnss")
    MkDir folderPath
    CreateTempImageFolder = folderPath
End Function

Private Sub CleanTempImageFolder(ByVal folderPath As String)
    On Error Resume Next
    Dim fileName As String
    fileName = Dir$(folderPath & Application.PathSeparator & "*.*")
    Do While Len(fileName) > 0
        Kill folderPath & Application.PathSeparator & fileName
        fileName = Dir$()
    Loop
    RmDir folderPath
    On Error GoTo 0
End Sub

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
