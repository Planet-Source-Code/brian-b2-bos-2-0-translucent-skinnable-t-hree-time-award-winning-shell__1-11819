VERSION 5.00
Begin VB.Form frmShutdownMenu 
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
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   4500
      ScaleHeight     =   3375
      ScaleWidth      =   5235
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   5235
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
      TabIndex        =   7
      Top             =   0
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
      TabIndex        =   6
      Top             =   600
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
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   3315
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   3
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   3300
      TabIndex        =   3
      Top             =   0
      Width           =   3300
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   2
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   3300
      TabIndex        =   2
      Top             =   600
      Width           =   3300
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   3300
      TabIndex        =   1
      Top             =   1200
      Width           =   3300
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   3300
      TabIndex        =   0
      Top             =   1800
      Width           =   3300
   End
End
Attribute VB_Name = "frmShutdownMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MenuHeight As Integer, MenuWidth As Integer
Public CorrectWidth As Integer, CorrectHeight As Integer
Private CurrentIndex As Integer, OldIndex As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If CurrentIndex = 3 Then CurrentIndex = 0 Else CurrentIndex = CurrentIndex + 1
                SelectImage (CurrentIndex)
                If OldIndex <> -1 Then UnselectImage (OldIndex)
                OldIndex = CurrentIndex
        Case vbKeyDown
            If CurrentIndex = 0 Then CurrentIndex = 3 Else CurrentIndex = CurrentIndex - 1
                SelectImage (CurrentIndex)
                If OldIndex <> -1 Then UnselectImage (OldIndex)
                OldIndex = CurrentIndex
        Case vbKeyEscape
            HideMe
    End Select
End Sub

Private Sub Form_Load()
    LoadSkinSettings
    LoadSkinImages
    
    OldIndex = -1
    CurrentIndex = -1
End Sub

Sub Display(X As Integer, Y As Integer)
    BltDesktop X, Y, picDesktopCapture
    
    For i = 0 To 3
        picItem(i).Top = GetMenuItemHeight * (3 - i)
        picItem(i).Height = GetMenuItemHeight
        picItem(i).Cls
        AlphaBlending picItem(i).HDC, 0, 0, MenuWidth, MenuHeight / 4, picDesktopCapture.HDC, 0, (3 - i) * (MenuHeight / 4), MenuWidth, MenuHeight / 4, MenuTranslucencyLevel
    Next
    
    SetWindowPos Me.hWND, HWND_TOPMOST, X, Y, MenuWidth, MenuHeight, 0
    Me.Show
    Me.Refresh
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
    If OldIndex <> -1 Then UnselectImage (OldIndex)
    If CurrentIndex <> -1 Then UnselectImage (CurrentIndex)
    CurrentIndex = -1
    OldIndex = -1
    Me.Hide
End Sub

Sub LoadSkinSettings()
    MenuWidth = GetShutdownMenuWidth
    MenuHeight = GetMenuItemHeight * 4
    picDesktopCapture.Width = MenuWidth
    picDesktopCapture.Height = MenuHeight
End Sub

Sub LoadSkinImages()
    For i = 0 To 3
        picItem(i).Picture = LoadPicture(GetSkinImage("Start Menu\Shutdown Menu\Normal\Image" & i & ".bmp"))
        picItemOver(i).Picture = LoadPicture(GetSkinImage("Start Menu\Shutdown Menu\Over\Image" & i & ".bmp"))
    Next
End Sub

Sub UnselectImage(Index As Integer)
    picItem(Index).Cls
    AlphaBlending picItem(Index).HDC, 0, 0, MenuWidth, MenuHeight / 4, picDesktopCapture.HDC, 0, (3 - Index) * (MenuHeight / 4), MenuWidth, MenuHeight / 4, MenuTranslucencyLevel
    picItem(Index).Refresh
    OldIndex = -1
End Sub

Sub SelectImage(Index As Integer)
    BitBlt picItem(Index).HDC, 0, 0, MenuWidth, MenuHeight / 4, picItemOver(Index).HDC, 0, 0, vbSrcCopy
    AlphaBlending picItem(Index).HDC, 0, 0, MenuWidth, MenuHeight / 4, picDesktopCapture.HDC, 0, (3 - Index) * (MenuHeight / 4), MenuWidth, MenuHeight / 4, MenuTranslucencyLevel
    picItem(Index).Refresh
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToggleStartMenu
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If OldIndex <> -1 Then UnselectImage (OldIndex)
    If CurrentIndex <> -1 Then UnselectImage (CurrentIndex)
    CurrentIndex = -1
    OldIndex = -1
End Sub

Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToggleStartMenu
    Select Case Index
        Case 0
            If ChoiceBoxEX("Do you really want to reset your shell to Explorer and reboot the computer?", "Reboot with Explorer") Then
                Debug.Print "<< Add shell reset functionallity >>"
                Shutdown SM_Reboot
            End If
        Case 1
            If ChoiceBoxEX("Do you really want to log off?", "Log off") Then
                Shutdown SM_Logoff
            End If
        Case 2
            If ChoiceBoxEX("Do you really want to restart the computer?", "Restart") Then
                Shutdown SM_Reboot
            End If
        Case 3
            If ChoiceBoxEX("Do you really want to shut down the computer?", "Shut down") Then
                Shutdown SM_Shutdown
            End If
    End Select
End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If OldIndex <> -1 And OldIndex <> Index Then UnselectImage (OldIndex)
    If CurrentIndex <> Index Then
        SelectImage (Index)
    End If
    OldIndex = Index
    CurrentIndex = Index
End Sub

