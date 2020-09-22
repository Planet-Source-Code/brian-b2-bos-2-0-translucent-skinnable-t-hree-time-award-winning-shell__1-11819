Attribute VB_Name = "modDesktopIconPopup"
Public DesktopIconsShown As Boolean

Sub ToggleDesktopIcons()
If DesktopIconsShown Then
    DesktopIconsShown = False
    frmTaskbar.DesktopIconOff
    frmFolderMenu.HideMenu
Else
    DesktopIconsShown = True
    frmTaskbar.DesktopIconOn
    frmFolderMenu.DrawMenu GetDesktopPath, frmTaskbar.StartButtonWidth * Screen.TwipsPerPixelX, (ScreenHeight - frmTaskbar.TaskbarHeight) * Screen.TwipsPerPixelY
    PlaySkinSound "menuopen"
End If
End Sub

