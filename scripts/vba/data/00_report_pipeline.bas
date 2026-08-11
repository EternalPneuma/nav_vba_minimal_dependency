' ReportPipeline：一键流程中的统一交互与失败传播

Option Explicit

Private mBatchMode As Boolean
Private mFailureMessage As String

Public Sub ReportPipeline_BeginBatch()
    mBatchMode = True
    mFailureMessage = vbNullString
End Sub

Public Sub ReportPipeline_EndBatch()
    mBatchMode = False
    mFailureMessage = vbNullString
End Sub

Public Function ReportPipeline_IsBatchMode() As Boolean
    ReportPipeline_IsBatchMode = mBatchMode
End Function

Public Function ReportPipeline_MsgBox(ByVal prompt As String, _
                                      Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                                      Optional ByVal title As String = vbNullString) As VbMsgBoxResult
    If Not mBatchMode Then
        ReportPipeline_MsgBox = VBA.MsgBox(prompt, buttons, title)
        Exit Function
    End If

    If IsPipelineFailureMessage(prompt, buttons, title) Then
        If Len(mFailureMessage) = 0 Then
            mFailureMessage = title & vbCrLf & prompt
        End If
    End If

    If (buttons And vbYesNo) = vbYesNo Then
        ReportPipeline_MsgBox = vbYes
    Else
        ReportPipeline_MsgBox = vbOK
    End If
End Function

Public Sub ReportPipeline_ThrowIfFailed()
    If Len(mFailureMessage) = 0 Then Exit Sub

    Dim failureMessage As String
    failureMessage = mFailureMessage
    mFailureMessage = vbNullString
    Err.Raise vbObjectError + 6301, , failureMessage
End Sub

Private Function IsPipelineFailureMessage(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle, _
                                          ByVal title As String) As Boolean
    If InStr(1, title, "开放日待确认", vbTextCompare) > 0 Then
        IsPipelineFailureMessage = True
    ElseIf (buttons And vbCritical) = vbCritical Then
        IsPipelineFailureMessage = True
    ElseIf InStr(1, prompt, "无法继续", vbTextCompare) > 0 Then
        IsPipelineFailureMessage = True
    ElseIf Left$(prompt, 20) Like "*失败*" Then
        IsPipelineFailureMessage = True
    End If
End Function
