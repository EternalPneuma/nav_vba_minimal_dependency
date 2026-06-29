Option Explicit

' UserForm 名称：frmOperationPanel
' 说明：为 chart、data、weekly_recommendation、product_one_page 与 tool 提供简单操作面板。
' 说明：本窗体会在 Initialize 中动态生成控件，不需要手工绘制按钮。

Private mButtonHandlers As Collection
Private mStatusLabel As MSForms.Label
Private mMainFrame As MSForms.Frame
Private mRecommendationFrame As MSForms.Frame
Private mOnePageFrame As MSForms.Frame
Private mToolFrame As MSForms.Frame
Private mGuideFrame As MSForms.Frame

Private Sub UserForm_Initialize()
    Me.Caption = "上层产品净值自动化操作面板"
    Me.Width = 1120
    Me.Height = 700
    Me.BackColor = RGB(248, 248, 248)

    Set mButtonHandlers = New Collection

    AddTitleLabel "titleMain", "上层产品净值自动化操作面板", 18, 14, 1080, 24, 12, True
    AddTitleLabel "titleDesc", "推荐优先使用“产品业绩展示”一键流程；单项按钮用于补跑、排查或手工维护。", 18, 42, 1080, 18, 9, False

    Set mMainFrame = AddSection("frameMainWorkflow", "产品业绩展示", 18, 76, 600, 255)
    AddActionButton mMainFrame, "btnMainAll", "一键生成展示文件（推荐）", 16, 28, 560, 30, Array("Data01_ImportNav181", "Data02_CalculateOpenDate", "Data03_ExportProductReport", "Chart01_ImportNavData", "Chart02_ExportProductSummary", "Chart03_GenerateCharts", "Chart04_ExportImages", "Data04_ExportDisplayReport")
    AddActionButton mMainFrame, "btnMainStep1", "1. 导入净值数据(181)", 16, 70, 174, 28, Array("Data01_ImportNav181")
    AddActionButton mMainFrame, "btnMainStep2", "2. 测算开放日", 207, 70, 174, 28, Array("Data02_CalculateOpenDate")
    AddActionButton mMainFrame, "btnMainStep3", "3. 输出分类表现", 398, 70, 174, 28, Array("Data03_ExportProductReport")
    AddActionButton mMainFrame, "btnMainStep4", "4. 导入绘图净值", 16, 112, 174, 28, Array("Chart01_ImportNavData")
    AddActionButton mMainFrame, "btnMainStep5", "5. 输出产品汇总", 207, 112, 174, 28, Array("Chart02_ExportProductSummary")
    AddActionButton mMainFrame, "btnMainStep6", "6. 生成产品图表", 398, 112, 174, 28, Array("Chart03_GenerateCharts")
    AddActionButton mMainFrame, "btnMainStep7", "7. 导出产品图片", 16, 154, 174, 28, Array("Chart04_ExportImages")
    AddActionButton mMainFrame, "btnMainStep8", "8. 输出展示报表", 207, 154, 365, 28, Array("Data04_ExportDisplayReport")
    AddInfoLabel mMainFrame, "lblMainNote", "顺序：导入 181 净值 → 测算开放日 → 输出分类表现 → 生成图表图片 → 输出展示报表。", 16, 202, 560, 32

    Set mGuideFrame = AddSection("frameGuide", "操作指南", 640, 76, 440, 335)
    AddInfoLabel mGuideFrame, "lblGuide", _
        "一、产品业绩展示" & vbCrLf & _
        "涉及产品：稳享长期限、直销、交通银行、江苏银行" & vbCrLf & _
        "数据准备：内网系统“自定义查询”搜索“181”，下载固收部产品净值文件并放到同级目录。" & vbCrLf & _
        "    - 注: 一个 Excel 工作簿放置一天的数据，命名规范为[HS-181_多账套净值查询_yyyymmdd.xlsx]。" & vbCrLf & _
        "手工维护：如需更新基准收益率，请维护“产品分类”sheet。" & vbCrLf & vbCrLf & _
        "二、一页通" & vbCrLf & _
        "涉及产品：汇鑫1号、交鑫致远" & vbCrLf & _
        "说明：主流程“导入绘图净值”已覆盖 0.补充净值；一键按钮默认只运行 1-3。" & vbCrLf & _
        "数据准备：REITs 数据可用 fetch_reits_total_return.py 获取，或直接维护对应 Excel。" & vbCrLf & vbCrLf & _
        "三、推荐材料" & vbCrLf & _
        "涉及产品：汇益5号、汇益稳健日开101号、汇益稳健28天101号" & vbCrLf & _
        "说明：通过依赖数据计算规模和收益率，再生成材料。" & vbCrLf & _
        "手工维护：工作表[产品要素][主要底层资产][底层资产对应关系]" & vbCrLf & vbCrLf & _
        "四、运行前检查" & vbCrLf & _
        "请先保存当前工作簿，确认源文件位于同级目录，且输出文件未被打开占用。", _
        16, 28, 405, 285, 8

    Set mOnePageFrame = AddSection("frameOnePage", "一页通", 18, 348, 290, 150)
    AddActionButton mOnePageFrame, "btnOnePageAll", "一键运行一页通流程", 16, 28, 250, 30, Array("OnePage01_ExportChartData", "OnePage02_GenerateCharts", "OnePage03_ExportPptPdf")
    AddActionButton mOnePageFrame, "btnOnePageStep0", "0. 补充净值", 16, 72, 115, 28, Array("OnePage00_CheckAndImportNavData")
    AddActionButton mOnePageFrame, "btnOnePageStep1", "1. 导出数据", 151, 72, 115, 28, Array("OnePage01_ExportChartData")
    AddActionButton mOnePageFrame, "btnOnePageStep2", "2. 生成图表", 16, 108, 115, 28, Array("OnePage02_GenerateCharts")
    AddActionButton mOnePageFrame, "btnOnePageStep3", "3. 导出PPT/PDF", 151, 108, 115, 28, Array("OnePage03_ExportPptPdf")

    Set mRecommendationFrame = AddSection("frameRecommendation", "推荐材料", 328, 348, 290, 150)
    AddActionButton mRecommendationFrame, "btnRecommendationAll", "一键运行推荐材料流程", 16, 28, 250, 30, Array("Weekly01_UpdateDependencies", "Weekly02_GenerateReport")
    AddActionButton mRecommendationFrame, "btnRecommendationStep1", "1. 更新依赖", 16, 72, 115, 28, Array("Weekly01_UpdateDependencies")
    AddActionButton mRecommendationFrame, "btnRecommendationStep2", "2. 生成材料", 151, 72, 115, 28, Array("Weekly02_GenerateReport")

    Set mToolFrame = AddSection("frameTool", "维护工具", 640, 430, 440, 181)
    AddActionButton mToolFrame, "btnToolCleanData", "1. 绘图去重", 16, 28, 190, 28, Array("Tool01_CleanDuplicateData")
    AddActionButton mToolFrame, "btnToolDeleteData", "2. 删除产品", 226, 28, 190, 28, Array("Tool02_DeleteByProductId")
    AddActionButton mToolFrame, "btnToolFillOpenDate", "3. 补开放日", 16, 68, 190, 28, Array("Tool03_FillNextOpenDate")
    AddActionButton mToolFrame, "btnToolCheckNavData", "4. 核对净值", 226, 68, 190, 28, Array("Tool04_CheckNavData")
    AddActionButton mToolFrame, "btnToolQuery181Stats", "5. 查询181", 16, 108, 190, 28, Array("Tool05_Query181NavStats")
    AddActionButton mToolFrame, "btnToolDeleteEmptyRows", "6. 删除空行", 226, 108, 190, 28, Array("Tool06_DeleteEmptyRows")

    AddActionButton Me, "btnClose", "关闭面板", 984, 628, 96, 30, Array("__close__")

    Set mStatusLabel = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With mStatusLabel
        .Left = 18
        .Top = 628
        .Width = 940
        .Height = 54
        .Caption = "状态：等待操作。"
        .BackStyle = fmBackStyleTransparent
        .ForeColor = RGB(70, 70, 70)
        .Font.Name = "微软雅黑"
        .Font.Size = 9
        .WordWrap = True
    End With
End Sub

Public Sub RunPanelAction(ByVal actionTitle As String, ByVal macroNames As Variant)
    If IsArray(macroNames) Then
        If UBound(macroNames) >= LBound(macroNames) Then
            If CStr(macroNames(LBound(macroNames))) = "__close__" Then
                Unload Me
                Exit Sub
            End If
        End If
    End If

    Dim confirmText As String
    confirmText = "确认运行：" & actionTitle & "？" & vbCrLf & vbCrLf & _
                  "运行前请确认：" & vbCrLf & _
                  "1. 当前工作簿已保存。" & vbCrLf & _
                  "2. 相关源文件已经放在同级目录。" & vbCrLf & _
                  "3. 需要的输出文件未被其他程序占用。"
    If MsgBox(confirmText, vbQuestion + vbYesNo + vbDefaultButton2, "操作面板") <> vbYes Then Exit Sub

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    On Error GoTo RunFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim i As Long
    For i = LBound(macroNames) To UBound(macroNames)
        SetStatus "正在运行：" & CStr(macroNames(i))
        DoEvents
        RunMacroByName CStr(macroNames(i))
    Next i

    SetStatus "完成：" & actionTitle

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Exit Sub

RunFail:
    SetStatus "失败：" & actionTitle & vbCrLf & Err.Description
    MsgBox "操作面板运行失败" & vbCrLf & vbCrLf & _
           "任务名称：" & actionTitle & vbCrLf & _
           "错误信息：" & Err.Description, vbExclamation, "操作面板"
    Resume CleanExit
End Sub

Private Sub RunMacroByName(ByVal macroName As String)
    ' VBE 中所有模块通常都在同一层级；这里按过程名直接调用即可。
    Application.Run "'" & ThisWorkbook.Name & "'!" & macroName
End Sub

Private Sub SetStatus(ByVal statusText As String)
    If Not mStatusLabel Is Nothing Then
        mStatusLabel.Caption = "状态：" & statusText
    End If
End Sub

Private Function AddSection(ByVal controlName As String, _
                            ByVal captionText As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal controlWidth As Single, _
                            ByVal controlHeight As Single) As MSForms.Frame
    Dim frameCtl As MSForms.Frame
    Set frameCtl = Me.Controls.Add("Forms.Frame.1", controlName, True)
    With frameCtl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .Font.Name = "微软雅黑"
        .Font.Size = 9
        .ForeColor = RGB(50, 50, 50)
    End With
    Set AddSection = frameCtl
End Function

Private Sub AddTitleLabel(ByVal controlName As String, _
                          ByVal captionText As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal controlWidth As Single, _
                          ByVal controlHeight As Single, _
                          ByVal fontSize As Single, _
                          ByVal isBold As Boolean)
    Dim labelCtl As MSForms.Label
    Set labelCtl = Me.Controls.Add("Forms.Label.1", controlName, True)
    With labelCtl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .BackStyle = fmBackStyleTransparent
        .ForeColor = RGB(35, 35, 35)
        .Font.Name = "微软雅黑"
        .Font.Size = fontSize
        .Font.Bold = isBold
    End With
End Sub

Private Sub AddInfoLabel(ByVal parentControl As Object, _
                         ByVal controlName As String, _
                         ByVal captionText As String, _
                         ByVal leftPos As Single, _
                         ByVal topPos As Single, _
                         ByVal controlWidth As Single, _
                         ByVal controlHeight As Single, _
                         Optional ByVal fontSize As Single = 9)
    Dim labelCtl As MSForms.Label
    Set labelCtl = parentControl.Controls.Add("Forms.Label.1", controlName, True)
    With labelCtl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .BackStyle = fmBackStyleTransparent
        .ForeColor = RGB(70, 70, 70)
        .Font.Name = "微软雅黑"
        .Font.Size = fontSize
        .WordWrap = True
    End With
End Sub

Private Sub AddActionButton(ByVal parentControl As Object, _
                            ByVal controlName As String, _
                            ByVal captionText As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal controlWidth As Single, _
                            ByVal controlHeight As Single, _
                            ByVal macroNames As Variant)
    Dim buttonCtl As MSForms.CommandButton
    Set buttonCtl = parentControl.Controls.Add("Forms.CommandButton.1", controlName, True)
    With buttonCtl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .Font.Name = "微软雅黑"
        .Font.Size = 9
        .TakeFocusOnClick = False
    End With

    Dim buttonHandler As clsOperationPanelButton
    Set buttonHandler = New clsOperationPanelButton
    buttonHandler.Init buttonCtl, Me, captionText, macroNames
    mButtonHandlers.Add buttonHandler
End Sub
