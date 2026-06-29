Option Explicit

Private Const SHEET_DATA_SUMMARY As String = "数据摘要"

'==============================================================
' 模块: 导出拼接图表
' 功能: 导出每个sheet的4个chart, 红色/蓝色均横向拼接, 单图存raw子文件夹
'==============================================================
  
Public Sub Chart04_ExportImages()
  
    Dim t0 As Double: t0 = Timer
  
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
  
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim dbPath As String: dbPath = wbDB.Path & "\"
  
    '--- 1. 查找最新的[产品净值汇总_yyyymmdd.xlsx] ---
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^产品净值汇总_(\d{8})\.xlsx$"
  
    Dim fileName As String, latestFile As String, latestKey As String
    fileName = Dir(dbPath & "产品净值汇总_*.xlsx")
    Do While Len(fileName) > 0
        If regex.Test(fileName) Then
            Dim matches As Object
            Set matches = regex.Execute(fileName)
            If matches(0).SubMatches(0) > latestKey Then
                latestKey = matches(0).SubMatches(0)
                latestFile = fileName
            End If
        End If
        fileName = Dir()
    Loop
      
    If Len(latestFile) = 0 Then
        MsgBox "产品图片导出无法继续" & vbCrLf & vbCrLf & _
               "错误信息：未找到[产品净值汇总_yyyymmdd.xlsx]文件，请先运行[生成产品图表]。", vbExclamation, "产品图片导出"
        GoTo CleanUp
    End If

    Dim targetPath As String: targetPath = dbPath & latestFile

    '--- 2. 创建输出文件夹结构 ---
    Dim outFolder As String, rawFolder As String
    outFolder = dbPath & "产品图表_" & latestKey & "\"
    rawFolder = outFolder & "raw\"

    If Dir(outFolder, vbDirectory) = "" Then MkDir outFolder
    If Dir(rawFolder, vbDirectory) = "" Then MkDir rawFolder

    '--- 3. 打开/获取目标工作簿 ---
    Dim wbTarget As Workbook
    Dim wasOpen As Boolean: wasOpen = False
  
    On Error Resume Next
    Set wbTarget = Workbooks(latestFile)
    On Error GoTo 0
  
    If wbTarget Is Nothing Then
        Set wbTarget = Workbooks.Open(targetPath)
    Else
        wasOpen = True
    End If

    '--- 4. 遍历每个sheet ---
    Dim ws As Worksheet
    Dim processedCount As Long: processedCount = 0
    Dim errMsg As String: errMsg = ""
  
    Dim chartNames As Variant
    chartNames = Array("chart_净值_红", "chart_收益率_红", "chart_净值_蓝", "chart_收益率_蓝")
  
    For Each ws In wbTarget.Worksheets
        If ws.Name = SHEET_DATA_SUMMARY Then GoTo NextSheet
      
        Dim safeName As String: safeName = ws.Name
      
        '--- 4.1 导出4个chart到raw文件夹 ---
        Dim rawFiles(1 To 4) As String
        Dim i As Long
      
        ' 临时开启屏幕刷新,让chart能正常渲染
        Application.ScreenUpdating = True
      
        For i = 1 To 4
            rawFiles(i) = rawFolder & safeName & "_" & Mid(CStr(chartNames(i - 1)), 7) & ".png"
      
            If ChartExists(ws, CStr(chartNames(i - 1))) Then
                Dim exportOK As Boolean
                exportOK = ExportChartWithRetry( _
                    ws.ChartObjects(CStr(chartNames(i - 1))), _
                    rawFiles(i), 3)  ' 最多重试3次
      
                If Not exportOK Then
                    errMsg = errMsg & ws.Name & "/" & CStr(chartNames(i - 1)) & "(导出空白,已重试), "
                    rawFiles(i) = ""
                End If
            Else
                rawFiles(i) = ""
            End If
        Next i
      
        Application.ScreenUpdating = False
      
        '--- 4.2 拼接红色组 (左右拼接) ---
        Dim outRed As String: outRed = outFolder & safeName & "_红.png"
        If Len(rawFiles(1)) > 0 And FileExists(rawFiles(1)) Then
            If Len(rawFiles(2)) > 0 And FileExists(rawFiles(2)) Then
                If Not MergeTwoImages(wbTarget, rawFiles(1), rawFiles(2), outRed) Then
                    errMsg = errMsg & ws.Name & "(红色拼接失败), "
                End If
            Else
                ' 收益率缺失,直接复制净值图
                FileCopy rawFiles(1), outRed
            End If
        End If
      
        '--- 4.3 拼接蓝色组 (左右拼接) ---
        Dim outBlue As String: outBlue = outFolder & safeName & "_蓝.png"
        If Len(rawFiles(3)) > 0 And FileExists(rawFiles(3)) Then
            If Len(rawFiles(4)) > 0 And FileExists(rawFiles(4)) Then
                If Not MergeTwoImages(wbTarget, rawFiles(3), rawFiles(4), outBlue) Then
                    errMsg = errMsg & ws.Name & "(蓝色拼接失败), "
                End If
            Else
                ' 收益率缺失,直接复制净值图
                FileCopy rawFiles(3), outBlue
            End If
        End If
  
        processedCount = processedCount + 1
NextSheet:
    Next ws

    '--- 5. 关闭目标文件 ---
    If Not wasOpen Then wbTarget.Close SaveChanges:=False

    '--- 6. 汇总提示 ---
    Dim msg As String
    msg = "产品图片导出完成" & vbCrLf & vbCrLf & _
          "耗时：" & Format(Timer - t0, "0.00") & " 秒" & vbCrLf & vbCrLf & _
          "处理结果：" & vbCrLf & _
          "处理产品数：" & processedCount & vbCrLf & vbCrLf & _
          "输出文件：" & vbCrLf & _
          "拼接图输出：" & outFolder & vbCrLf & _
          "原图输出：" & rawFolder
    If Len(errMsg) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "注意事项：" & vbCrLf & _
              "异常：" & vbCrLf & Left(errMsg, Len(errMsg) - 2)
    End If
    MsgBox msg, vbInformation, "产品图片导出"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

'==============================================================
' 辅助函数: 检查chart是否存在
'==============================================================
Private Function ChartExists(ByVal ws As Worksheet, ByVal chartName As String) As Boolean
    Dim co As ChartObject
    On Error Resume Next
    Set co = ws.ChartObjects(chartName)
    On Error GoTo 0
    ChartExists = Not (co Is Nothing)
End Function

'==============================================================
' 辅助函数: 检查文件是否存在
'==============================================================
Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir(filePath)) > 0)
End Function

'==============================================================
' 辅助函数: 横向拼接两张图片(返回成功/失败) - 供红色图表使用
'   实现: 创建临时chart -> 设尺寸 -> 把两张图作为Shape添加到chart内 -> 导出
'==============================================================
Private Function MergeTwoImages(ByVal wb As Workbook, _
                                 ByVal leftPath As String, _
                                 ByVal rightPath As String, _
                                 ByVal outputPath As String) As Boolean

    On Error GoTo ErrHandler

    ' 1. 先用临时Shape读取两张图的真实尺寸
    Dim tmpSheet As Worksheet
    Set tmpSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    tmpSheet.Name = "_tmp_" & Format(Timer * 1000, "0")

    Dim probeLeft As Shape, probeRight As Shape
    Set probeLeft = tmpSheet.Shapes.AddPicture( _
        fileName:=leftPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=0, Top:=0, Width:=-1, Height:=-1)
    Set probeRight = tmpSheet.Shapes.AddPicture( _
        fileName:=rightPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=0, Top:=0, Width:=-1, Height:=-1)

    Dim leftW As Single: leftW = probeLeft.Width
    Dim leftH As Single: leftH = probeLeft.Height
    Dim rightW As Single: rightW = probeRight.Width
    Dim rightH As Single: rightH = probeRight.Height

    Dim gap As Single: gap = 30   '设置左右间距为30磅
    Dim totalW As Single: totalW = leftW + rightW + gap
    Dim totalH As Single
    totalH = IIf(leftH > rightH, leftH, rightH)

    ' 删除探测用的图片
    probeLeft.Delete
    probeRight.Delete

    ' 2. 创建chart容器, 尺寸=总宽×最大高
    Dim co As ChartObject
    Set co = tmpSheet.ChartObjects.Add( _
        Left:=0, Top:=0, Width:=totalW, Height:=totalH)

    ' chart白色背景,无边框,无标题,无系列
    Dim ch As Chart: Set ch = co.Chart

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ch.HasTitle = False

    With ch.ChartArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Solid
    End With
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Fill.Visible = msoFalse

    ' 隐藏坐标轴(可能模板没设但默认会有)
    On Error Resume Next
    ch.Axes(xlCategory).Delete
    ch.Axes(xlValue).Delete
    On Error GoTo 0

    ' 3. 直接在chart内部添加两张图作为shape
    Dim shpLeft As Shape, shpRight As Shape
    Set shpLeft = ch.Shapes.AddPicture( _
        fileName:=leftPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=0, Top:=0, Width:=leftW, Height:=leftH)

    Set shpRight = ch.Shapes.AddPicture( _
        fileName:=rightPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=leftW + gap, Top:=0, Width:=rightW, Height:=rightH)

    ' 4. 导出
    ch.Export fileName:=outputPath, FilterName:="PNG"

    ' 5. 清理临时sheet
    Application.DisplayAlerts = False
    tmpSheet.Delete
    Application.DisplayAlerts = True

    MergeTwoImages = True
    Exit Function

ErrHandler:
    Debug.Print "MergeTwoImages 错误: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not tmpSheet Is Nothing Then
        Application.DisplayAlerts = False
        tmpSheet.Delete
        Application.DisplayAlerts = True
    End If
    MergeTwoImages = False
End Function

'==============================================================
' 辅助函数: 导出chart,带渲染等待和重试机制
'   - 激活chart触发渲染
'   - DoEvents让消息队列跑完
'   - 检查输出文件大小,过小则视为空白,重试
'==============================================================
Private Function ExportChartWithRetry(ByVal co As ChartObject, _
                                       ByVal outputPath As String, _
                                       ByVal maxRetries As Long) As Boolean
    Const MIN_VALID_SIZE As Long = 3000  ' 字节,小于此视为空白图

    Dim attempt As Long
    Dim ws As Worksheet: Set ws = co.Parent

    For attempt = 1 To maxRetries
        ' 1. 激活sheet和chart,触发渲染
        On Error Resume Next
        ws.Activate
        co.Activate           ' 激活ChartObject容器
        co.Chart.Refresh      ' 强制刷新
        DoEvents              ' 让系统处理消息队列
        DoEvents

        ' 2. 短暂等待让chart完成绘制
        ' 等待时长随重试次数递增: 0.5s, 1s, 1.5s
        Application.Wait Now + TimeSerial(0, 0, 0) + (attempt * 0.5 / 86400)
        DoEvents

        ' 3. 取消选中(避免影响后续操作)
        ws.Range("A1").Select
        DoEvents

        ' 4. 导出
        ' 删除已存在的旧文件,避免缓存
        If Dir(outputPath) <> "" Then Kill outputPath

        co.Chart.Export fileName:=outputPath, FilterName:="PNG"

        If Err.Number <> 0 Then
            Debug.Print "导出错误(attempt " & attempt & "): " & Err.Description
            Err.Clear
            On Error GoTo 0
            GoTo NextAttempt
        End If
        On Error GoTo 0

        ' 5. 验证文件大小
        Dim fSize As Long
        fSize = FileLen(outputPath)

        If fSize >= MIN_VALID_SIZE Then
            ExportChartWithRetry = True
            Exit Function
        End If

        Debug.Print "图片过小(" & fSize & " bytes), sheet=" & ws.Name & _
                    " chart=" & co.Name & " attempt=" & attempt

NextAttempt:
    Next attempt

    ExportChartWithRetry = False
End Function
