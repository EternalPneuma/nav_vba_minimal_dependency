Option Explicit

' 模块用途：打开 VBA 操作面板。
' 使用方式：
' 1. 在 VBE 中插入标准模块，复制本文件内容。
' 2. 插入 UserForm，名称改为 frmOperationPanel，复制 00_operation_panel_form 内容。
' 3. 插入类模块，名称改为 clsOperationPanelButton，复制 00_operation_panel_button 内容。
' 4. 运行 ShowOperationPanel。

Public Sub AAA启动控制台()
    ' 使用运行时名称避免在尚未导入 UserForm 时触发“变量未定义”的编译错误。
    ' 使用 sync-vba.ps1 同步时，脚本会创建名为 frmOperationPanel 的 UserForm。
    Dim panel As Object

    On Error GoTo FormNotInstalled
    Set panel = VBA.UserForms.Add("frmOperationPanel")
    panel.Show vbModeless
    Exit Sub

FormNotInstalled:
    MsgBox "操作面板无法打开" & vbCrLf & vbCrLf & _
           "错误信息：未找到操作面板窗体 frmOperationPanel。" & vbCrLf & _
           "请先运行 sync-vba.ps1 同步 optional_panel 模块组。", _
           vbExclamation, "操作面板"
End Sub
