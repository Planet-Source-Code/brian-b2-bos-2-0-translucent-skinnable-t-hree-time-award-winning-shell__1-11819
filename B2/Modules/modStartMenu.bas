Attribute VB_Name = "modStartMenu"
Public StartMenuShown As Boolean
Public MenuShown(0 To 7) As Boolean
Sub HideMenus()
    If MenuShown(7) Then frmFolderMenu.HideMenu: MenuShown(7) = False
    If MenuShown(6) Then frmFolderMenu.HideMenu: MenuShown(6) = False
    If MenuShown(5) Then frmFolderMenu.HideMenu: MenuShown(5) = False
    If MenuShown(4) Then frmSettingsMenu.HideMe: MenuShown(4) = False
    If MenuShown(3) Then frmFindMenu.HideMe: MenuShown(3) = False
    If MenuShown(2) Then frmHelpMenu.HideMe: MenuShown(2) = False
    If MenuShown(0) Then frmShutdownMenu.HideMe: MenuShown(0) = False
End Sub

Sub ToggleStartMenu()
If StartMenuShown Then
    StartMenuShown = False
    frmTaskbar.StartButtonOff
    frmStartMenu.HideMe
    HideMenus
Else
    StartMenuShown = True
    frmTaskbar.HideTaskbarTips
    DoEvents
    frmTaskbar.StartButtonOn
    frmStartMenu.Display
    PlaySkinSound "menuopen"
End If
End Sub
