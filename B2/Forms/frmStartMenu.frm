VERSION 5.00
Begin VB.Form frmStartMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrExpand 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4020
      Top             =   2160
   End
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      Height          =   4875
      Left            =   1200
      ScaleHeight     =   4815
      ScaleWidth      =   5175
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   5235
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   7
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   6
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   5
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   4
      Left            =   3315
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   3
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   2
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   1
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   7
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3300
      TabIndex        =   7
      Top             =   0
      Width           =   3300
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   6
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3300
      TabIndex        =   6
      Top             =   600
      Width           =   3300
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   5
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3300
      TabIndex        =   5
      Top             =   1200
      Width           =   3300
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   4
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3300
      TabIndex        =   4
      Top             =   1800
      Width           =   3300
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   3
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3300
      TabIndex        =   3
      Top             =   2400
      Width           =   3300
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   2
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3300
      TabIndex        =   2
      Top             =   3000
      Width           =   3300
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   1
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3300
      TabIndex        =   1
      Top             =   3600
      Width           =   3300
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3300
      TabIndex        =   0
      Top             =   4200
      Width           =   3300
   End
End
Attribute VB_Name = "frmStartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartMenuHeight As Integer, StartMenuWidth As Integer
Public CorrectWidth As Integer, CorrectHeight As Integer
Private CurrentIndex As Integer, OldIndex As Integer
Private ExpandIndex As Integer
Private finalwidth As Integer, finalheight As Integer, finalleft As Integer, showindex As Integer
Implements ICcrpTimerNotify
Private tmrShow As ccrpTimer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If CurrentIndex = 7 Then CurrentIndex = 0 Else CurrentIndex = CurrentIndex + 1
                HideMenus
                SelectImage (CurrentIndex)
                If OldIndex <> -1 Then UnselectImage (OldIndex)
                OldIndex = CurrentIndex
        Case vbKeyDown
            If CurrentIndex = 0 Then CurrentIndex = 7 Else CurrentIndex = CurrentIndex - 1
                HideMenus
                SelectImage (CurrentIndex)
                If OldIndex <> -1 Then UnselectImage (OldIndex)
                OldIndex = CurrentIndex
        Case vbKeyRight
                ExpandMenu CurrentIndex, True
        Case vbKeyEscape
            HideMe
    End Select
End Sub

Private Sub Form_Load()
    Set tmrShow = New ccrpTimer
    LoadSkinSettings
    LoadSkinImages
    
    OldIndex = -1
    CurrentIndex = -1
End Sub

Sub Display()
    picDesktopCapture.Width = GetStartMenuWidth + 10
    picDesktopCapture.Height = GetStartMenuHeight + 10
    If GetMenuShadow Then
        BltDesktop frmTaskbar.CurLeft, frmTaskbar.CurTop - StartMenuHeight + 10, picDesktopCapture
        CreateAlphaShadow StartMenuWidth, StartMenuHeight
        SetWindowPos Me.hWND, HWND_TOPMOST, frmTaskbar.CurLeft, frmTaskbar.CurTop - StartMenuHeight + 10, StartMenuWidth, StartMenuHeight, 0
    Else
        BltDesktop 0, ScreenHeight - StartMenuHeight - frmTaskbar.TaskbarHeight, picDesktopCapture
        SetWindowPos Me.hWND, HWND_TOPMOST, 0, frmTaskbar.CurTop - Me.ScaleHeight, StartMenuWidth, StartMenuHeight, 0
    End If
    For i = 0 To 7
        'picItem(i).Top = (CorrectHeight / 8) * (7 - i)
        picItem(i).Cls
        AlphaBlending picItem(i).HDC, 0, 0, StartMenuWidth, CorrectHeight / 8, picDesktopCapture.HDC, 0, (7 - i) * (CorrectHeight / 8), StartMenuWidth, CorrectHeight / 8, MenuTranslucencyLevel
    Next
    
    finalheight = Me.Height
    finalwidth = Me.Width
    Me.Width = 0
    Me.Height = 0
    SetWindowPos Me.hWND, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    Me.Refresh
    Set tmrShow.Notify = Me
    showindex = 0
    tmrShow.Interval = 5
    tmrShow.Stats.Frequency = 5
    tmrShow.Enabled = True
End Sub

Sub CreateAlphaShadow(wid As Integer, hei As Integer)
    Me.Cls
    For i = 1 To 5
        'Draw the right shadow
        AlphaBlending Me.HDC, wid - 5 + i, 5, 1, hei - 5, picDesktopCapture.HDC, wid - 5 + i, 5, 1, hei - 5, i * 40 + 60
        'The right corner
        AlphaBlending Me.HDC, wid - 5, i + 5, 5, 1, picDesktopCapture.HDC, wid - 5, i + 5, 5, 1, (5 - i) * 40 + 60
        'The upper right corner
        BitBlt Me.HDC, wid - 5, 0, 5, 6, picDesktopCapture.HDC, wid - 5, 0, vbSrcCopy
        'The bottom section
        AlphaBlending Me.HDC, 5, hei - 5 + i, wid - 5, 1, picDesktopCapture.HDC, 5, hei - 5 + i, wid - 5, 1, i * 40 + 60
        'The bottom corner
        AlphaBlending Me.HDC, 5 + i, hei - 5, 1, 5, picDesktopCapture.HDC, 5 + i, hei - 5, 1, 1, (5 - i) * 40 + 60
        'The left bottom corner
        BitBlt Me.HDC, 0, hei - 5, 6, 5, picDesktopCapture.HDC, 0, hei - 5, vbSrcCopy
    Next
End Sub

Sub HideMe()
    tmrExpand.Enabled = False
    tmrShow.Enabled = False
    If OldIndex <> -1 Then UnselectImage (OldIndex)
    If CurrentIndex <> -1 Then UnselectImage (CurrentIndex)
    CurrentIndex = -1
    OldIndex = -1
    Me.Hide
End Sub

Sub LoadSkinSettings()
    If GetMenuShadow Then
        StartMenuWidth = GetStartMenuWidth + 5
        StartMenuHeight = GetStartMenuHeight + 5
    Else
        StartMenuWidth = GetStartMenuWidth
        StartMenuHeight = GetStartMenuHeight
    End If
    CorrectWidth = GetStartMenuWidth
    CorrectHeight = GetStartMenuHeight
End Sub

Sub LoadSkinImages()
    For i = 0 To 7
        picItem(i).Picture = LoadPicture(GetSkinImage("Start Menu\Main Menu\Normal\Image" & i & ".bmp"))
        picItemOver(i).Picture = LoadPicture(GetSkinImage("Start Menu\Main Menu\Over\Image" & i & ".bmp"))
    Next
End Sub

Sub UnselectImage(Index As Integer)
    picItem(Index).Cls
    AlphaBlending picItem(Index).HDC, 0, 0, StartMenuWidth, CorrectHeight / 8, picDesktopCapture.HDC, 0, (7 - Index) * (CorrectHeight / 8), StartMenuWidth, CorrectHeight / 8, MenuTranslucencyLevel
    picItem(Index).Refresh
    OldIndex = -1
End Sub

Sub SelectImage(Index As Integer)
    If OldIndex <> Index Then
        BitBlt picItem(Index).HDC, 0, 0, StartMenuWidth, CorrectHeight / 8, picItemOver(Index).HDC, 0, 0, vbSrcCopy
        AlphaBlending picItem(Index).HDC, 0, 0, StartMenuWidth, CorrectHeight / 8, picDesktopCapture.HDC, 0, (7 - Index) * (CorrectHeight / 8), StartMenuWidth, CorrectHeight / 8, MenuTranslucencyLevel
        picItem(Index).Refresh
        If Index = 1 Then PlaySkinSound "menuitemhover"
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToggleStartMenu
End Sub



Private Sub ICcrpTimerNotify_Timer(ByVal Milliseconds As Long)
    showindex = showindex + 1
    Me.Height = finalheight * (showindex / 25)
    Me.Width = finalwidth * (showindex / 25)
    If showindex = 25 Then
        tmrShow.Enabled = False
        SetWindowPos Me.hWND, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    End If
End Sub

Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 1
            ToggleStartMenu
            frmRun.Show
    End Select
End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If OldIndex <> -1 And OldIndex <> Index Then UnselectImage (OldIndex): HideMenus
    SelectImage (Index)
        
    ExpandIndex = Index
    tmrExpand.Enabled = True
    
    OldIndex = Index
End Sub

Sub ExpandMenu(Index As Integer, Optional Keystroke As Boolean = False)
    If CurrentIndex <> Index Or Keystroke Then
        Select Case Index
            Case 7
                frmFolderMenu.DrawMenu GetStartMenuPath, (CorrectWidth - 5) * Screen.TwipsPerPixelX, (ScreenHeight - StartMenuHeight) * Screen.TwipsPerPixelY
                MenuShown(7) = True
            Case 6
                frmFolderMenu.DrawMenu GetFavoritesPath, (CorrectWidth - 5) * Screen.TwipsPerPixelY, (ScreenHeight - StartMenuHeight + (CorrectHeight / 8)) * Screen.TwipsPerPixelY
                MenuShown(6) = True
            Case 5
                frmFolderMenu.DrawMenu GetRecent, (CorrectWidth - 5) * Screen.TwipsPerPixelY, (ScreenHeight - StartMenuHeight + ((CorrectHeight / 8) * 2)) * Screen.TwipsPerPixelY
                MenuShown(5) = True
            Case 4
                frmSettingsMenu.Display CorrectWidth + (Me.Left / Screen.TwipsPerPixelX) - 5, Me.Top / Screen.TwipsPerPixelY + picItem(4).Top
                MenuShown(4) = True
            Case 3
                frmFindMenu.Display CorrectWidth + (Me.Left / Screen.TwipsPerPixelX) - 5, Me.Top / Screen.TwipsPerPixelY + picItem(3).Top
                MenuShown(3) = True
            Case 2
                frmHelpMenu.Display CorrectWidth + (Me.Left / Screen.TwipsPerPixelX) - 5, Me.Top / Screen.TwipsPerPixelY + picItem(2).Top
                MenuShown(2) = True
            Case 0
                frmShutdownMenu.Display CorrectWidth + (Me.Left / Screen.TwipsPerPixelX) - 5, (Me.Top / Screen.TwipsPerPixelY) + picItem(0).Top - 40
                MenuShown(0) = True
        End Select
        If Index <> 1 Then PlaySkinSound "menuopen"
    End If
End Sub

Private Sub tmrExpand_Timer()
    ExpandMenu (ExpandIndex)
    tmrExpand.Enabled = False
    CurrentIndex = ExpandIndex
End Sub


