Attribute VB_Name = "modSkinCode"
Private SkinINIPath As String
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Sub SetTranslucencyLevel(val As Integer)
    SaveSetting "B2", "User", "TranslucencyLevel", Str(val)
End Sub

Function TranslucencyLevel() As Integer
    TranslucencyLevel = val(GetSetting("B2", "User", "TranslucencyLevel", "60"))
End Function

Sub SetMenuTranslucencyLevel(val As Integer)
    SaveSetting "B2", "User", "MenuTranslucencyLevel", Str(val)
End Sub

Function MenuTranslucencyLevel() As Integer
    MenuTranslucencyLevel = val(GetSetting("B2", "User", "MenuTranslucencyLevel", "60"))
End Function

Function GetSkinName() As String
    GetSkinName = GetSetting("B2", "Skin", "SkinName", "B2 Default")
End Function

Function GetTranslucent()
    GetTranslucent = (GetSetting("B2", "Skin", "Translucent", "true") = "true")
End Function

Function GetTaskbarButtonWidth() As Integer 'Loads the width of the taskbar program button
    GetTaskbarButtonWidth = val(GetSetting("B2", "Skin", "TaskbarButtonWidth", "200"))
End Function

Function GetMyComputerWidth() As Integer 'Loads the width of the my computer button
    GetMyComputerWidth = val(GetSetting("B2", "Skin", "MyComputerWidth", "30"))
End Function

Function GetTaskbarHeight() As Integer 'Loads the height of the taskbar
    GetTaskbarHeight = val(GetSetting("B2", "Skin", "TaskbarHeight", "30"))
End Function

Function GetStartButtonWidth() As Integer 'Loads the width of the start button
    GetStartButtonWidth = val(GetSetting("B2", "Skin", "StartButtonWidth", "30"))
End Function

Function GetDesktopButtonWidth()
    GetDesktopButtonWidth = val(GetSetting("B2", "Skin", "DesktopButtonWidth", "30"))
End Function

Function GetSkinImage(name As String) As String
    GetSkinImage = App.path & "\skins\" & GetSkinName & "\" & name
End Function

Function GetStartMenuWidth() As Integer
    GetStartMenuWidth = val(GetSetting("B2", "Skin", "StartMenuWidth", "220"))
End Function

Function GetStartMenuHeight() As Integer
    GetStartMenuHeight = val(GetSetting("B2", "Skin", "StartMenuHeight", "320"))
End Function

Function GetMenuShadow() As Boolean
    GetMenuShadow = (GetSetting("B2", "Skin", "MenuShadow", "true") = "true")
End Function

Function GetMenuTranslucency() As Boolean
    GetMenuTranslucency = (GetSetting("B2", "Skin", "MenuTranslucent", "true") = "true")
End Function

Function GetMenuItemHeight() As Integer
    GetMenuItemHeight = val(GetSetting("B2", "Skin", "MenuItemHeight", "20"))
End Function

Function GetShutdownMenuWidth() As Integer
    GetShutdownMenuWidth = val(GetSetting("B2", "Skin", "MenuWidth", "175"))
End Function

Function GetNormalColor() As Long
    GetNormalColor = val(GetSetting("B2", "Skin", "NormalColor", "&H00000000&"))
End Function

Function GetOverColor() As Long
    GetOverColor = val(GetSetting("B2", "Skin", "OverColor", "&H00000000&"))
End Function

Function PlaySkinSound(name As String)
    sndPlaySound App.path & "\skins\" & GetSkinName & "\sounds\" & name & ".wav", 1
End Function
Sub ChangeSkin(SkinName As String)
    SkinINIPath = App.path & "\skins\" & SkinName & "\skin.ini"
    SaveSetting "B2", "Skin", "SkinName", SkinName
    SaveSetting "B2", "Skin", "MenuShadow", LCase(ReadINI("startmenu", "alphashadow", SkinINIPath))
    SaveSetting "B2", "Skin", "MenuTranslucent", LCase(ReadINI("menus", "translucent", SkinINIPath))
    SaveSetting "B2", "Skin", "NormalColor", ReadINI("menus", "normalcolor", SkinINIPath)
    SaveSetting "B2", "Skin", "OverColor", ReadINI("menus", "overcolor", SkinINIPath)
    SaveSetting "B2", "Skin", "Translucent", ReadINI("taskbar", "translucent", SkinINIPath)

    Unload frmTaskbar
    Unload frmStartMenu
    Unload frmFolderMenu
    Unload frmShutdownMenu
    Unload frmSettingsMenu
    Unload frmFindMenu
    Unload frmHelpMenu
    
    frmTaskbar.Show
End Sub
