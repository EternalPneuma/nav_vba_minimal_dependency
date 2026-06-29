Option Explicit

' 类模块名称：clsOperationPanelButton
' 用途：承接操作面板动态按钮的 Click 事件。

Private WithEvents mButton As MSForms.CommandButton
Private mOwnerForm As Object
Private mMacroNames As Variant
Private mActionTitle As String

Public Sub Init(ByVal targetButton As MSForms.CommandButton, _
                ByVal ownerForm As Object, _
                ByVal actionTitle As String, _
                ByVal macroNames As Variant)
    Set mButton = targetButton
    Set mOwnerForm = ownerForm
    mActionTitle = actionTitle
    mMacroNames = macroNames
End Sub

Private Sub mButton_Click()
    mOwnerForm.RunPanelAction mActionTitle, mMacroNames
End Sub

