Option Explicit

' UserForm 名称：frmOperationPanel
' 说明：固定尺寸操作面板。主按钮运行一键流程，右侧箭头展开可直接运行的单项任务。
' 说明：所有控件在 Initialize 中动态生成，不需要手工绘制。

Private Const PANEL_WIDTH As Single = 600
Private Const PANEL_HEIGHT As Single = 420
Private Const MENU_MAIN As String = "main"
Private Const MENU_ONE_PAGE As String = "onepage"
Private Const MENU_RECOMMENDATION As String = "recommendation"
Private Const MENU_TOOL As String = "tool"

Private mButtonHandlers As Collection
Private mAllButtons As Collection
Private mMenuFrames As Collection
Private mMenuButtonGroups As Collection
Private mOpenMenuKey As String
Private mStatusLabel As MSForms.Label
Private mDataDateLabel As MSForms.Label
Private mIsRunning As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "上层产品净值自动化操作面板"
    Me.Width = PANEL_WIDTH
    Me.Height = PANEL_HEIGHT
    Me.StartUpPosition = 1
    Me.BackColor = RGB(248, 248, 248)

    Set mButtonHandlers = New Collection
    Set mAllButtons = New Collection
    Set mMenuFrames = New Collection
    Set mMenuButtonGroups = New Collection

    AddTitleLabel "titleMain", "上层产品净值自动化操作面板", 14, 10, 570, 22, 11, True
    AddTitleLabel "titleDesc", "点击主按钮运行完整流程；点击 ▼ 可单独运行子流程。", 14, 34, 570, 16, 9, False

    AddSplitButton "btnMainAll", "btnMainMenu", "一键生成产品业绩展示（推荐优先）", _
                   14, 58, 570, 31, _
                   Array("Report00_RequireValidConfiguration", "Data01_ImportNav181", "Data02_CalculateOpenDate", "Data03_ExportProductReport", "Chart01_ImportNavData", "Chart02_ExportProductSummary", "Chart03_GenerateCharts", "Chart04_ExportImages", "Data04_ExportDisplayReport"), _
                   MENU_MAIN, "system", "按标准顺序更新净值、开放日、图表和展示报表。"

    Set mDataDateLabel = AddInfoLabel(Me, "lblDataDate", "当前净值截止：正在读取……", 18, 94, 562, 17, 9)

    AddSplitButton "btnOnePageAll", "btnOnePageMenu", "一键运行一页通", _
                   14, 118, 570, 31, _
                   Array("OnePage01_ExportChartData", "OnePage02_GenerateCharts", "OnePage03_ExportPptPdf"), _
                   MENU_ONE_PAGE, "system", "生成一页通数据、图表和 PPT/PDF。"

    AddSplitButton "btnRecommendationAll", "btnRecommendationMenu", "一键生成推荐材料", _
                   14, 158, 570, 31, _
                   Array("Weekly01_UpdateDependencies", "Weekly02_GenerateReport"), _
                   MENU_RECOMMENDATION, "system", "更新依赖数据后生成推荐材料。"

    AddSplitButton "btnTool", "btnToolMenu", "维护工具", _
                   14, 211, 570, 29, Array("__toggle_only__"), _
                   MENU_TOOL, "secondary", "展开数据维护、检查查询和报表配置工具。"

    AddActionButton Me, "btnGuide", "操作指南", 14, 251, 112, 27, Array("__guide__"), "secondary", _
                    "查看数据准备、维护位置和运行前检查。"
    AddInfoLabel Me, "lblRunTip", "运行前请保存工作簿，并关闭已打开的输出文件。", 140, 253, 318, 24, 9
    AddActionButton Me, "btnClose", "关闭面板", 472, 251, 112, 27, Array("__close__"), "secondary", _
                    "关闭操作面板。"

    Set mStatusLabel = AddInfoLabel(Me, "lblStatus", "等待：请选择操作。", 14, 293, 570, 82, 9)
    With mStatusLabel
        .BorderStyle = fmBorderStyleSingle
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(242, 242, 242)
        .ForeColor = RGB(90, 90, 90)
        .Caption = "等待：请选择操作。"
    End With

    BuildMainMenu
    BuildOnePageMenu
    BuildRecommendationMenu
    BuildToolMenu
    RefreshDataDate
End Sub

Private Sub AddSplitButton(ByVal mainName As String, ByVal arrowName As String, _
                           ByVal captionText As String, ByVal leftPos As Single, _
                           ByVal topPos As Single, ByVal totalWidth As Single, _
                           ByVal controlHeight As Single, ByVal macroNames As Variant, _
                           ByVal menuKey As String, ByVal buttonStyle As String, _
                           ByVal tipText As String)
    Const ARROW_WIDTH As Single = 32

    AddActionButton Me, mainName, captionText, leftPos, topPos, totalWidth - ARROW_WIDTH, _
                    controlHeight, macroNames, buttonStyle, tipText
    AddActionButton Me, arrowName, "▼", leftPos + totalWidth - ARROW_WIDTH, topPos, _
                    ARROW_WIDTH, controlHeight, Array("__toggle__" & menuKey), buttonStyle, _
                    "展开或收起“" & captionText & "”的单项任务。"
End Sub

Private Sub BuildMainMenu()
    Dim menuFrame As MSForms.Frame
    Dim menuButtons As Collection
    Set menuButtons = New Collection
    Set menuFrame = AddPopupFrame("popupMain", 14, 90, 570, 220)

    AddMenuButton menuFrame, menuButtons, "menuMainValidate", "预检报表配置", 7, 7, 554, 20, _
                  Array("Report00_ValidateConfiguration"), "normal", "只检查配置，不生成报表。"
    AddSeparator menuFrame, "sepMain", 7, 32, 554

    AddMenuButton menuFrame, menuButtons, "menuMain1", "1. 导入净值数据（181）", 7, 39, 554, 20, Array("Data01_ImportNav181"), "normal", "导入同级目录中的 181 净值文件。"
    AddMenuButton menuFrame, menuButtons, "menuMain2", "2. 测算开放日", 7, 61, 554, 20, Array("Data02_CalculateOpenDate"), "normal", "测算开放日；待确认日期需人工核对。"
    AddMenuButton menuFrame, menuButtons, "menuMain3", "3. 输出分类表现", 7, 83, 554, 20, Array("Data03_ExportProductReport"), "normal", "生成分类表现中间报表。"
    AddMenuButton menuFrame, menuButtons, "menuMain4", "4. 导入绘图净值", 7, 105, 554, 20, Array("Chart01_ImportNavData"), "normal", "更新绘图净值数据。"
    AddMenuButton menuFrame, menuButtons, "menuMain5", "5. 输出产品汇总", 7, 127, 554, 20, Array("Chart02_ExportProductSummary"), "normal", "生成产品净值汇总文件。"
    AddMenuButton menuFrame, menuButtons, "menuMain6", "6. 生成产品图表", 7, 149, 554, 20, Array("Chart03_GenerateCharts"), "normal", "按配置生成产品图表。"
    AddMenuButton menuFrame, menuButtons, "menuMain7", "7. 导出产品图片", 7, 171, 554, 20, Array("Chart04_ExportImages"), "normal", "导出产品组合图图片。"
    AddMenuButton menuFrame, menuButtons, "menuMain8", "8. 输出展示报表", 7, 193, 554, 20, Array("Data04_ExportDisplayReport"), "normal", "生成最终展示报表。"

    RegisterMenu MENU_MAIN, menuFrame, menuButtons
End Sub

Private Sub BuildOnePageMenu()
    Dim menuFrame As MSForms.Frame
    Dim menuButtons As Collection
    Set menuButtons = New Collection
    Set menuFrame = AddPopupFrame("popupOnePage", 14, 149, 570, 122)

    AddMenuButton menuFrame, menuButtons, "menuOnePage0", "0. 补充净值（可选）", 7, 7, 554, 20, Array("OnePage00_CheckAndImportNavData"), "normal", "从工作簿源数据补充一页通产品的绘图净值。"
    AddSeparator menuFrame, "sepOnePage", 7, 32, 554
    AddMenuButton menuFrame, menuButtons, "menuOnePage1", "1. 导出数据", 7, 39, 554, 20, Array("OnePage01_ExportChartData"), "normal", "导出一页通图表数据。"
    AddMenuButton menuFrame, menuButtons, "menuOnePage2", "2. 生成图表", 7, 61, 554, 20, Array("OnePage02_GenerateCharts"), "normal", "生成一页通图表。"
    AddMenuButton menuFrame, menuButtons, "menuOnePage3", "3. 导出 PPT/PDF", 7, 83, 554, 20, Array("OnePage03_ExportPptPdf"), "normal", "生成一页通 PPT 和 PDF。"

    RegisterMenu MENU_ONE_PAGE, menuFrame, menuButtons
End Sub

Private Sub BuildRecommendationMenu()
    Dim menuFrame As MSForms.Frame
    Dim menuButtons As Collection
    Set menuButtons = New Collection
    Set menuFrame = AddPopupFrame("popupRecommendation", 14, 189, 570, 57)

    AddMenuButton menuFrame, menuButtons, "menuRecommendation1", "1. 更新依赖", 7, 7, 554, 20, Array("Weekly01_UpdateDependencies"), "normal", "更新推荐材料所需的规模和收益率数据。"
    AddMenuButton menuFrame, menuButtons, "menuRecommendation2", "2. 生成材料", 7, 29, 554, 20, Array("Weekly02_GenerateReport"), "normal", "使用已更新的依赖数据生成推荐材料。"

    RegisterMenu MENU_RECOMMENDATION, menuFrame, menuButtons
End Sub

Private Sub BuildToolMenu()
    Dim menuFrame As MSForms.Frame
    Dim menuButtons As Collection
    Set menuButtons = New Collection
    Set menuFrame = AddPopupFrame("popupTool", 14, 63, 570, 148)

    AddMenuHeading menuFrame, "headingData", "数据维护", 8, 5, 550
    AddMenuButton menuFrame, menuButtons, "menuToolClean", "绘图去重", 7, 20, 273, 20, Array("Tool01_CleanDuplicateData"), "danger", "删除重复和缺少主键的绘图净值记录，重复时保留最后导入记录。"
    AddMenuButton menuFrame, menuButtons, "menuToolDelete", "删除产品", 287, 20, 273, 20, Array("Tool02_DeleteByProductId"), "danger", "按产品编号永久删除绘图净值记录。"
    AddMenuButton menuFrame, menuButtons, "menuToolOpenDate", "补开放日", 7, 42, 273, 20, Array("Tool03_FillNextOpenDate"), "normal", "单独运行开放日推算与人工确认流程。"
    AddMenuButton menuFrame, menuButtons, "menuToolEmptyRows", "删除空行", 287, 42, 273, 20, Array("Tool06_DeleteEmptyRows"), "danger", "按指定列删除当前工作表中的空行。"

    AddMenuHeading menuFrame, "headingCheck", "检查与查询", 8, 67, 550
    AddMenuButton menuFrame, menuButtons, "menuToolCheck", "核对净值", 7, 82, 273, 20, Array("Tool04_CheckNavData"), "normal", "检查产品净值数据完整性。"
    AddMenuButton menuFrame, menuButtons, "menuToolQuery", "查询 181", 287, 82, 273, 20, Array("Tool05_Query181NavStats"), "normal", "查询 181 净值文件统计信息。"

    AddMenuHeading menuFrame, "headingConfig", "报表配置", 8, 107, 550
    AddMenuButton menuFrame, menuButtons, "menuToolMigrate", "初始化报表配置", 7, 122, 273, 20, Array("Report00_MigrateConfiguration"), "normal", "创建缺失的报表配置结构；不会覆盖完整的现有配置。"
    AddMenuButton menuFrame, menuButtons, "menuToolValidate", "预检报表配置", 287, 122, 273, 20, Array("Report00_ValidateConfiguration"), "normal", "只检查配置，不生成报表。"

    RegisterMenu MENU_TOOL, menuFrame, menuButtons
End Sub

Private Sub RegisterMenu(ByVal menuKey As String, ByVal menuFrame As MSForms.Frame, ByVal menuButtons As Collection)
    menuFrame.Visible = False
    mMenuFrames.Add menuFrame, menuKey
    mMenuButtonGroups.Add menuButtons, menuKey
End Sub

Private Function AddPopupFrame(ByVal controlName As String, ByVal leftPos As Single, _
                               ByVal topPos As Single, ByVal controlWidth As Single, _
                               ByVal controlHeight As Single) As MSForms.Frame
    Dim frameCtl As MSForms.Frame
    Set frameCtl = Me.Controls.Add("Forms.Frame.1", controlName, True)
    With frameCtl
        .Caption = vbNullString
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .BackColor = RGB(250, 250, 250)
        .BorderStyle = fmBorderStyleSingle
        .Visible = False
    End With
    Set AddPopupFrame = frameCtl
End Function

Private Sub AddMenuButton(ByVal menuFrame As MSForms.Frame, ByVal menuButtons As Collection, _
                          ByVal controlName As String, ByVal captionText As String, _
                          ByVal leftPos As Single, ByVal topPos As Single, _
                          ByVal controlWidth As Single, ByVal controlHeight As Single, _
                          ByVal macroNames As Variant, ByVal buttonStyle As String, _
                          ByVal tipText As String)
    Dim buttonCtl As MSForms.CommandButton
    Set buttonCtl = AddActionButton(menuFrame, controlName, captionText, leftPos, topPos, _
                                    controlWidth, controlHeight, macroNames, buttonStyle, tipText)
    menuButtons.Add buttonCtl
End Sub

Private Sub AddMenuHeading(ByVal menuFrame As MSForms.Frame, ByVal controlName As String, _
                           ByVal captionText As String, ByVal leftPos As Single, _
                           ByVal topPos As Single, ByVal controlWidth As Single)
    Dim labelCtl As MSForms.Label
    Set labelCtl = menuFrame.Controls.Add("Forms.Label.1", controlName, True)
    With labelCtl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = 13
        .BackStyle = fmBackStyleTransparent
        .ForeColor = RGB(100, 100, 100)
        .Font.Name = "微软雅黑"
        .Font.Size = 8
        .Font.Bold = True
    End With
End Sub

Private Sub AddSeparator(ByVal menuFrame As MSForms.Frame, ByVal controlName As String, _
                         ByVal leftPos As Single, ByVal topPos As Single, _
                         ByVal controlWidth As Single)
    Dim labelCtl As MSForms.Label
    Set labelCtl = menuFrame.Controls.Add("Forms.Label.1", controlName, True)
    With labelCtl
        .Caption = vbNullString
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = 1
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(210, 210, 210)
    End With
End Sub

Public Sub RunPanelAction(ByVal actionTitle As String, ByVal macroNames As Variant)
    Dim commandText As String
    commandText = CStr(macroNames(LBound(macroNames)))

    If Left$(commandText, 10) = "__toggle__" Then
        ToggleMenu Mid$(commandText, 11)
        Exit Sub
    ElseIf commandText = "__toggle_only__" Then
        ToggleMenu MENU_TOOL
        Exit Sub
    End If

    HideAllMenus

    If commandText = "__close__" Then
        Unload Me
        Exit Sub
    ElseIf commandText = "__guide__" Then
        ShowOperationGuide
        Exit Sub
    End If

    Dim runOnePagePrep As Boolean
    If Not PrepareOnePageIfNeeded(macroNames, runOnePagePrep) Then Exit Sub
    If Not ConfirmPanelAction(actionTitle, macroNames) Then Exit Sub

    Dim isBatchMode As Boolean
    isBatchMode = (UBound(macroNames) > LBound(macroNames))
    If isBatchMode Then ReportPipeline_BeginBatch

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    Dim currentStepTitle As String
    mIsRunning = True
    SetPanelEnabled False

    On Error GoTo RunFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    If runOnePagePrep Then
        currentStepTitle = "补充一页通净值"
        SetStatus "运行中：准备，" & currentStepTitle, "running"
        DoEvents
        RunMacroByName "OnePage00_CheckAndImportNavData"
        If isBatchMode Then ReportPipeline_ThrowIfFailed
    End If

    Dim i As Long
    Dim stepNumber As Long
    Dim stepCount As Long
    stepCount = BusinessStepCount(macroNames)

    For i = LBound(macroNames) To UBound(macroNames)
        currentStepTitle = MacroDisplayName(CStr(macroNames(i)))
        If CStr(macroNames(i)) = "Report00_RequireValidConfiguration" Then
            SetStatus "运行中：准备，预检报表配置", "running"
        Else
            stepNumber = BusinessStepNumber(macroNames, i)
            SetStatus "运行中：步骤 " & stepNumber & "/" & stepCount & "，" & currentStepTitle, "running"
        End If
        DoEvents
        RunMacroByName CStr(macroNames(i))
        If isBatchMode Then ReportPipeline_ThrowIfFailed
    Next i

    SetStatus "完成：" & actionTitle, "success"
    RefreshDataDate
    If isBatchMode Then VBA.MsgBox actionTitle & "完成。", vbInformation, "操作面板"

CleanExit:
    If isBatchMode Then ReportPipeline_EndBatch
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    SetPanelEnabled True
    mIsRunning = False
    Exit Sub

RunFail:
    Dim errorDescription As String
    Dim adviceText As String
    errorDescription = Err.Description
    adviceText = FailureAdvice(errorDescription)
    SetStatus "失败：" & currentStepTitle & vbCrLf & errorDescription, "failure"
    MsgBox "操作面板运行失败" & vbCrLf & vbCrLf & _
           "流程：" & actionTitle & vbCrLf & _
           "失败步骤：" & currentStepTitle & vbCrLf & _
           "错误信息：" & errorDescription & vbCrLf & vbCrLf & _
           "处理建议：" & adviceText, vbExclamation, "操作面板"
    Resume CleanExit
End Sub

Private Function ConfirmPanelAction(ByVal actionTitle As String, ByVal macroNames As Variant) As Boolean
    Dim macroName As String
    macroName = CStr(macroNames(LBound(macroNames)))

    Dim promptText As String
    If macroName = "Tool01_CleanDuplicateData" Then
        promptText = "确认运行“绘图去重”？" & vbCrLf & vbCrLf & _
                     "该操作会按净值日期和产品编号去重，保留最后导入的记录，" & _
                     "并删除缺少主键的无效记录。请先保存工作簿。"
        ConfirmPanelAction = (MsgBox(promptText, vbYesNo + vbExclamation + vbDefaultButton2, "绘图去重") = vbYes)
        Exit Function
    End If

    If macroName = "Report00_MigrateConfiguration" Then
        promptText = "确认初始化报表配置？" & vbCrLf & vbCrLf & _
                     "该操作会创建缺失的配置结构；完整的现有配置不会被覆盖。"
        ConfirmPanelAction = (MsgBox(promptText, vbYesNo + vbQuestion + vbDefaultButton2, "初始化报表配置") = vbYes)
        Exit Function
    End If

    If UBound(macroNames) > LBound(macroNames) Then
        promptText = "确认运行“" & actionTitle & "”？" & vbCrLf & vbCrLf & _
                     "运行前请保存当前工作簿，确认源文件位于同级目录，" & _
                     "并关闭已打开的输出文件。"
        ConfirmPanelAction = (MsgBox(promptText, vbYesNo + vbQuestion + vbDefaultButton2, "操作面板") = vbYes)
        Exit Function
    End If

    ConfirmPanelAction = True
End Function

Private Function PrepareOnePageIfNeeded(ByVal macroNames As Variant, ByRef runPrep As Boolean) As Boolean
    PrepareOnePageIfNeeded = True
    runPrep = False

    If UBound(macroNames) <= LBound(macroNames) Then Exit Function
    If CStr(macroNames(LBound(macroNames))) <> "OnePage01_ExportChartData" Then Exit Function

    Dim sourceDate As Double
    Dim targetDate As Double
    sourceDate = GetLatestSheetDate("上层产品净值数据(181)", "日期")
    targetDate = GetLatestSheetDate("绘图净值数据", "净值日期")
    If sourceDate <= 0 Or targetDate <= 0 Or sourceDate <= targetDate Then Exit Function

    Dim response As VbMsgBoxResult
    response = MsgBox("工作簿中的源净值日期晚于绘图净值日期。" & vbCrLf & vbCrLf & _
                      "源净值截止：" & Format$(CDate(sourceDate), "yyyy-mm-dd") & vbCrLf & _
                      "绘图净值截止：" & Format$(CDate(targetDate), "yyyy-mm-dd") & vbCrLf & vbCrLf & _
                      "是否先运行“补充净值”，然后继续一页通流程？" & vbCrLf & _
                      "选择“否”将直接继续；选择“取消”将停止。", _
                      vbYesNoCancel + vbExclamation + vbDefaultButton1, "一页通数据检查")

    If response = vbCancel Then
        PrepareOnePageIfNeeded = False
    ElseIf response = vbYes Then
        runPrep = True
    End If
End Function

Private Sub ToggleMenu(ByVal menuKey As String)
    If mIsRunning Then Exit Sub

    If mOpenMenuKey = menuKey Then
        HideAllMenus
        Exit Sub
    End If

    HideAllMenus

    Dim menuFrame As MSForms.Frame
    Set menuFrame = mMenuFrames(menuKey)
    menuFrame.Visible = True
    menuFrame.ZOrder 0
    mOpenMenuKey = menuKey
    SetArrowCaption menuKey, "▲"

    Dim menuButtons As Collection
    Set menuButtons = mMenuButtonGroups(menuKey)
    If menuButtons.Count > 0 Then menuButtons(1).SetFocus
End Sub

Private Sub HideAllMenus()
    Dim menuKeys As Variant
    menuKeys = Array(MENU_MAIN, MENU_ONE_PAGE, MENU_RECOMMENDATION, MENU_TOOL)

    Dim i As Long
    Dim menuFrame As MSForms.Frame
    For i = LBound(menuKeys) To UBound(menuKeys)
        On Error Resume Next
        Set menuFrame = mMenuFrames(CStr(menuKeys(i)))
        If Not menuFrame Is Nothing Then menuFrame.Visible = False
        SetArrowCaption CStr(menuKeys(i)), "▼"
        Set menuFrame = Nothing
        On Error GoTo 0
    Next i
    mOpenMenuKey = vbNullString
End Sub

Private Sub SetArrowCaption(ByVal menuKey As String, ByVal captionText As String)
    Dim arrowName As String
    Select Case menuKey
        Case MENU_MAIN: arrowName = "btnMainMenu"
        Case MENU_ONE_PAGE: arrowName = "btnOnePageMenu"
        Case MENU_RECOMMENDATION: arrowName = "btnRecommendationMenu"
        Case MENU_TOOL: arrowName = "btnToolMenu"
    End Select
    If Len(arrowName) > 0 Then Me.Controls(arrowName).Caption = captionText
End Sub

Public Function HandlePanelKey(ByVal sourceButton As Object, ByVal keyCode As Long) As Boolean
    If mIsRunning Then
        HandlePanelKey = True
        Exit Function
    End If

    If keyCode = vbKeyEscape Then
        If Len(mOpenMenuKey) > 0 Then
            HideAllMenus
        Else
            Unload Me
        End If
        HandlePanelKey = True
        Exit Function
    End If

    If Len(mOpenMenuKey) = 0 Then Exit Function
    If keyCode <> vbKeyUp And keyCode <> vbKeyDown Then Exit Function

    Dim menuButtons As Collection
    Set menuButtons = mMenuButtonGroups(mOpenMenuKey)

    Dim i As Long
    Dim nextIndex As Long
    For i = 1 To menuButtons.Count
        If menuButtons(i).Name = sourceButton.Name Then
            If keyCode = vbKeyDown Then
                nextIndex = i + 1
                If nextIndex > menuButtons.Count Then nextIndex = 1
            Else
                nextIndex = i - 1
                If nextIndex < 1 Then nextIndex = menuButtons.Count
            End If
            menuButtons(nextIndex).SetFocus
            HandlePanelKey = True
            Exit Function
        End If
    Next i
End Function

Private Sub SetPanelEnabled(ByVal isEnabled As Boolean)
    Dim buttonCtl As Object
    For Each buttonCtl In mAllButtons
        buttonCtl.Enabled = isEnabled
    Next buttonCtl
End Sub

Private Sub SetStatus(ByVal statusText As String, ByVal statusKind As String)
    If mStatusLabel Is Nothing Then Exit Sub

    mStatusLabel.Caption = statusText
    Select Case statusKind
        Case "running"
            mStatusLabel.ForeColor = RGB(31, 78, 121)
            mStatusLabel.BackColor = RGB(226, 239, 252)
        Case "success"
            mStatusLabel.ForeColor = RGB(46, 110, 54)
            mStatusLabel.BackColor = RGB(232, 245, 233)
        Case "failure"
            mStatusLabel.ForeColor = RGB(156, 35, 35)
            mStatusLabel.BackColor = RGB(253, 235, 235)
        Case Else
            mStatusLabel.ForeColor = RGB(90, 90, 90)
            mStatusLabel.BackColor = RGB(242, 242, 242)
    End Select
End Sub

Private Sub RefreshDataDate()
    Dim latestDate As Double
    latestDate = GetLatestSheetDate("上层产品净值数据(181)", "日期")
    If latestDate > 0 Then
        mDataDateLabel.Caption = "当前净值截止：" & Format$(CDate(latestDate), "yyyy-mm-dd")
    Else
        mDataDateLabel.Caption = "当前净值截止：未识别"
    End If
End Sub

Private Function GetLatestSheetDate(ByVal sheetName As String, ByVal headerText As String) As Double
    On Error GoTo DateUnavailable

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim dateCol As Long
    Dim c As Long
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) = headerText Then
            dateCol = c
            Exit For
        End If
    Next c
    If dateCol = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row

    Dim r As Long
    Dim cellValue As Variant
    Dim dateValue As Double
    For r = 2 To lastRow
        cellValue = ws.Cells(r, dateCol).Value
        If IsDate(cellValue) Then
            dateValue = CDbl(CDate(cellValue))
            If dateValue > GetLatestSheetDate Then GetLatestSheetDate = dateValue
        End If
    Next r
    Exit Function

DateUnavailable:
    GetLatestSheetDate = 0
End Function

Private Function BusinessStepCount(ByVal macroNames As Variant) As Long
    BusinessStepCount = UBound(macroNames) - LBound(macroNames) + 1
    If CStr(macroNames(LBound(macroNames))) = "Report00_RequireValidConfiguration" Then
        BusinessStepCount = BusinessStepCount - 1
    End If
End Function

Private Function BusinessStepNumber(ByVal macroNames As Variant, ByVal itemIndex As Long) As Long
    BusinessStepNumber = itemIndex - LBound(macroNames) + 1
    If CStr(macroNames(LBound(macroNames))) = "Report00_RequireValidConfiguration" Then
        BusinessStepNumber = BusinessStepNumber - 1
    End If
End Function

Private Function MacroDisplayName(ByVal macroName As String) As String
    Select Case macroName
        Case "Report00_RequireValidConfiguration", "Report00_ValidateConfiguration": MacroDisplayName = "预检报表配置"
        Case "Report00_MigrateConfiguration": MacroDisplayName = "初始化报表配置"
        Case "Data01_ImportNav181": MacroDisplayName = "导入净值数据（181）"
        Case "Data02_CalculateOpenDate": MacroDisplayName = "测算开放日"
        Case "Data03_ExportProductReport": MacroDisplayName = "输出分类表现"
        Case "Chart01_ImportNavData": MacroDisplayName = "导入绘图净值"
        Case "Chart02_ExportProductSummary": MacroDisplayName = "输出产品汇总"
        Case "Chart03_GenerateCharts": MacroDisplayName = "生成产品图表"
        Case "Chart04_ExportImages": MacroDisplayName = "导出产品图片"
        Case "Data04_ExportDisplayReport": MacroDisplayName = "输出展示报表"
        Case "OnePage00_CheckAndImportNavData": MacroDisplayName = "补充净值"
        Case "OnePage01_ExportChartData": MacroDisplayName = "导出一页通数据"
        Case "OnePage02_GenerateCharts": MacroDisplayName = "生成一页通图表"
        Case "OnePage03_ExportPptPdf": MacroDisplayName = "导出一页通 PPT/PDF"
        Case "Weekly01_UpdateDependencies": MacroDisplayName = "更新推荐材料依赖"
        Case "Weekly02_GenerateReport": MacroDisplayName = "生成推荐材料"
        Case "Tool01_CleanDuplicateData": MacroDisplayName = "绘图去重"
        Case "Tool02_DeleteByProductId": MacroDisplayName = "删除产品"
        Case "Tool03_FillNextOpenDate": MacroDisplayName = "补开放日"
        Case "Tool04_CheckNavData": MacroDisplayName = "核对净值"
        Case "Tool05_Query181NavStats": MacroDisplayName = "查询 181"
        Case "Tool06_DeleteEmptyRows": MacroDisplayName = "删除空行"
        Case Else: MacroDisplayName = macroName
    End Select
End Function

Private Function FailureAdvice(ByVal errorDescription As String) As String
    If InStr(1, errorDescription, "开放日待确认", vbTextCompare) > 0 Then
        FailureAdvice = "核对“开放日待确认”工作表，确认后重新运行完整流程。"
    ElseIf InStr(1, errorDescription, "配置", vbTextCompare) > 0 Then
        FailureAdvice = "运行“预检报表配置”，按提示修正后重新运行。"
    ElseIf InStr(1, errorDescription, "占用", vbTextCompare) > 0 Then
        FailureAdvice = "关闭已打开的输出文件后重新运行。"
    Else
        FailureAdvice = "检查源文件和工作表数据；修正问题后重新运行完整流程。"
    End If
End Function

Private Sub ShowOperationGuide()
    Dim guideText As String
    guideText = "一、产品业绩展示" & vbCrLf & _
                "推荐优先运行，用于导入最新 181 净值并生成分类表现、图表和展示报表。" & vbCrLf & _
                "产品归属和基准收益率维护“产品分类”；分组、字段和图表维护“报表配置”。" & vbCrLf & _
                "出现“开放日待确认”时，核对确认后再次运行一键流程。" & vbCrLf & vbCrLf & _
                "二、一页通" & vbCrLf & _
                "可以独立启动，但可能复用已有净值数据；建议先更新产品业绩展示。" & vbCrLf & _
                "“补充净值”是可选前置任务，一键流程默认运行步骤 1-3。" & vbCrLf & vbCrLf & _
                "三、推荐材料" & vbCrLf & _
                "先更新依赖，再生成材料。" & vbCrLf & vbCrLf & _
                "四、运行前检查" & vbCrLf & _
                "保存当前工作簿，确认源文件位于同级目录，并关闭已打开的输出文件。"

    MsgBox guideText, vbInformation, "操作指南"
End Sub

Private Sub RunMacroByName(ByVal macroName As String)
    Application.Run "'" & ThisWorkbook.Name & "'!" & macroName
End Sub

Private Function AddActionButton(ByVal parentControl As Object, _
                                 ByVal controlName As String, _
                                 ByVal captionText As String, _
                                 ByVal leftPos As Single, _
                                 ByVal topPos As Single, _
                                 ByVal controlWidth As Single, _
                                 ByVal controlHeight As Single, _
                                 ByVal macroNames As Variant, _
                                 Optional ByVal buttonStyle As String = "normal", _
                                 Optional ByVal tipText As String = vbNullString) As MSForms.CommandButton
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
        .TakeFocusOnClick = True
        .TabStop = True
        .Default = False
        .Cancel = False
        .ControlTipText = tipText
    End With
    ApplyButtonStyle buttonCtl, buttonStyle

    Dim buttonHandler As clsOperationPanelButton
    Set buttonHandler = New clsOperationPanelButton
    buttonHandler.Init buttonCtl, Me, captionText, macroNames
    mButtonHandlers.Add buttonHandler
    mAllButtons.Add buttonCtl
    Set AddActionButton = buttonCtl
End Function

Private Sub ApplyButtonStyle(ByVal buttonCtl As MSForms.CommandButton, ByVal buttonStyle As String)
    Select Case buttonStyle
        Case "system"
            ' 保留 MSForms/Windows 的默认按钮外观。
        Case "primary"
            buttonCtl.BackColor = RGB(168, 28, 45)
            buttonCtl.ForeColor = RGB(255, 255, 255)
            buttonCtl.Font.Bold = True
        Case "danger"
            buttonCtl.BackColor = RGB(252, 226, 229)
            buttonCtl.ForeColor = RGB(145, 30, 42)
        Case "secondary"
            buttonCtl.BackColor = RGB(235, 235, 235)
            buttonCtl.ForeColor = RGB(55, 55, 55)
        Case Else
            buttonCtl.BackColor = RGB(255, 255, 255)
            buttonCtl.ForeColor = RGB(45, 45, 45)
    End Select
End Sub

Private Sub AddTitleLabel(ByVal controlName As String, ByVal captionText As String, _
                          ByVal leftPos As Single, ByVal topPos As Single, _
                          ByVal controlWidth As Single, ByVal controlHeight As Single, _
                          ByVal fontSize As Single, ByVal isBold As Boolean)
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

Private Function AddInfoLabel(ByVal parentControl As Object, ByVal controlName As String, _
                              ByVal captionText As String, ByVal leftPos As Single, _
                              ByVal topPos As Single, ByVal controlWidth As Single, _
                              ByVal controlHeight As Single, _
                              Optional ByVal fontSize As Single = 9) As MSForms.Label
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
    Set AddInfoLabel = labelCtl
End Function

Private Sub UserForm_Click()
    If Not mIsRunning Then HideAllMenus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mIsRunning Then
        Cancel = True
        MsgBox "任务正在运行，完成或失败后才能关闭操作面板。", vbInformation, "操作面板"
    End If
End Sub

Private Sub UserForm_Terminate()
    OperationPanel_Released
End Sub
