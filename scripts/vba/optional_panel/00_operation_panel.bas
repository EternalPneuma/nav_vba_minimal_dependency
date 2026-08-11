Option Explicit

' 模块用途：打开 VBA 操作面板，并保证同一工作簿只存在一个面板实例。

Private mOperationPanel As Object

Public Sub AAA启动控制台()
    On Error GoTo FormNotInstalled

    If Not mOperationPanel Is Nothing Then
        On Error Resume Next
        If mOperationPanel.Visible Then
            mOperationPanel.Show vbModeless
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo FormNotInstalled
        Set mOperationPanel = Nothing
    End If

    Set mOperationPanel = VBA.UserForms.Add("frmOperationPanel")
    mOperationPanel.Show vbModeless
    Exit Sub

FormNotInstalled:
    Set mOperationPanel = Nothing
    MsgBox "操作面板无法打开" & vbCrLf & vbCrLf & _
           "错误信息：未找到操作面板窗体 frmOperationPanel。" & vbCrLf & _
           "请先运行 sync-vba.ps1 同步 optional_panel 模块组。", _
           vbExclamation, "操作面板"
End Sub

Public Sub OperationPanel_Released()
    Set mOperationPanel = Nothing
End Sub
