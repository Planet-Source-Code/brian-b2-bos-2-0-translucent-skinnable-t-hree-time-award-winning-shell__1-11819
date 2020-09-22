Attribute VB_Name = "modSystemColors"
Declare Function GetSysColor Lib "User32" (ByVal nIndex) As Long
Declare Sub SetSysColors Lib "User32" (ByVal nChanges%, lpSysColor%, lpColorValues&)
Dim NewColors(24) As Long
Dim IndexArray(24) As Integer

Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_3DDKSHADOW = 21
Public Const COLOR_3DFACE = COLOR_BTNFACE
Public Const COLOR_3DHILIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_INFOBK = 24
Public Const COLOR_INFOTEXT = 23

Sub RandomColors()
        SetSysColors 1, COLOR_BTNHIGHLIGHT, RGB(250, 250, 250)
        SetSysColors 1, COLOR_BTNSHADOW, RGB(180, 180, 180)
        SetSysColors 1, COLOR_3DLIGHT, vbBlack
        SetSysColors 1, COLOR_3DDKSHADOW, vbBlack
        SetSysColors 1, COLOR_WINDOWFRAME, vbBlack
        SetSysColors 1, COLOR_BTNFACE, RGB(220, 220, 220)
        SetSysColors 1, COLOR_MENU, RGB(220, 220, 220)
End Sub
