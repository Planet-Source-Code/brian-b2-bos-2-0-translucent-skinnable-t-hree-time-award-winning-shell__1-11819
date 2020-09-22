Attribute VB_Name = "modCustomApplets"
Public ToolTipDisplayed As Boolean

Public Sub MsgBoxEX(text As String, Optional caption As String = "Information", Optional sound As SoundTypes = sndNone)
    frmInfo.DisplayInfo text, caption, sound
End Sub

Public Function ChoiceBoxEX(text As String, Optional caption As String = "Question") As Boolean
    frmChoiceBox.lblInfo.caption = text
    frmChoiceBox.lblTitle.caption = caption
    PlaySkinSound "Question"
    frmChoiceBox.Show vbModal
    ChoiceBoxEX = frmChoiceBox.Response
End Function

Public Sub ToolTipEX(text As String, Top, Left)
    If ToolTipDisplayed = True Then HideToolTip
    DoEvents
    DoEvents
    If ToolTipDisplayed = False Then
        frmToolTip.DisplayTip text, Val(Top), Val(Left)
        ToolTipDisplayed = True
    End If
End Sub

Public Sub HideToolTip()
    If ToolTipDisplayed = True Then
        SetWindowPos frmToolTip.hWND, 0, 0, 0, 0, 0, SWP_HIDEWINDOW Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
        ToolTipDisplayed = False
    End If
End Sub

