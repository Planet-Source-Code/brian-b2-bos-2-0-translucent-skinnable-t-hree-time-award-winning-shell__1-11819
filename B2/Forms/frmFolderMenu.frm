VERSION 5.00
Begin VB.Form frmFolderMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1740
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   16
      Top             =   1380
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTitlebar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   13
      Top             =   180
      Width           =   4635
      Begin VB.PictureBox picScrollDown 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3060
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picScrollUp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2820
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrActivePreview 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   840
      Top             =   960
   End
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1740
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   11
      Top             =   1380
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrPopup 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1740
      Top             =   1380
   End
   Begin VB.PictureBox picOver 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   3795
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   9
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   9
      Top             =   2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picFile 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   8
      Left            =   -780
      ScaleHeight     =   255
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   -300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   7
      Left            =   -2700
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   6
      Left            =   -2700
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   -2700
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   -2700
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   -2700
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   -2700
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   -1860
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   -1380
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmFolderMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PxH As Integer
Dim PxW As Integer
Dim Contents As Variant
Dim TxtWidth As Integer
Dim MaxWidth As Integer
Dim CurrentPath As String
Dim CurrentIndex As Integer
Dim OldIndex As Integer
Dim BorderSize As Integer
Dim PopupIndex As Integer
Private CurrentPopup As Integer
Dim CurrentFormIndex As Integer
Private ParentIndex As Integer
Private finalwidth As Integer, finalheight As Integer, finalleft As Integer, showindex As Integer
Implements ICcrpTimerNotify
Private tmrShow As ccrpTimer
Private curScroll As Integer
Dim PreviewIndex As Integer
Dim DragNow As Boolean, TX As Integer, TY As Integer

Function GetFormIndex() As Integer
    GetFormIndex = CurrentFormIndex
End Function

Public Sub DrawMenu(path As String, Optional X As Integer = -1, Optional Y As Integer = -1, Optional ParentInt As Integer = -1)
    maxitems = -1
    Set tmrShow = New ccrpTimer
    ParentIndex = ParentInt
    CurrentFormIndex = Forms.Count
    CurrentIndex = -1
    OldIndex = -1
    
    PopupIndex = -1
    CurrentPopup = -1
    
    CurrentPath = AddASlash(path)
    
    Contents = ListFolderItems(CurrentPath)
    Me.Left = frmTaskbar.picDesktopButton.Left * Screen.TwipsPerPixelX
    
    MaxWidth = picTitlebar.TextWidth(GetFileName(path)) + 20
    For i = 0 To UBound(Contents)
        TxtWidth = picFile(0).TextWidth(RemoveExtention(Contents(i)))
        If TxtWidth > MaxWidth Then MaxWidth = TxtWidth
    Next
    LoadSkinImages
    Me.Width = (MaxWidth + BorderSize + 30) * Screen.TwipsPerPixelX
    Me.Height = ((UBound(Contents) + 1) * 16 + BorderSize * 2 + 16) * Screen.TwipsPerPixelY
    
    picTitlebar.Width = Me.ScaleWidth - (BorderSize * 2)
    picTitlebar.Top = BorderSize
    picTitlebar.Left = BorderSize
    DrawFileIcon path, picTitlebar.HDC, Icon_Small, 2, 0
    picTitlebar.CurrentX = 25
    picTitlebar.CurrentY = 1
    picTitlebar.Print GetFileName(path)
    
    If ParentIndex = -1 Then
        Me.Top = Y - Me.Height
    Else
        Me.Top = Y
    End If
    
    If Me.Height > Screen.Height Then
        Me.Top = 0
        Me.Height = Screen.Height
        maxitems = Int((ScreenHeight - BorderSize * 2) / 16) - 1
        
        picScrollDown.Visible = True
        picScrollDown.Left = Me.ScaleWidth - BorderSize - 32

        picScrollUp.Visible = True
        picScrollUp.Left = Me.ScaleWidth - BorderSize - 16
        picScrollUp.Enabled = False
    ElseIf Me.Height + Me.Top > Screen.Height Then
        Me.Top = Screen.Height - Me.Height
    End If
    
    Me.Left = X
            
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
    
    DrawBackground
    If GetMenuTranslucency Then AlphaBlending picTitlebar.HDC, 0, 0, Me.ScaleWidth - BorderSize * 2, 16, picDesktopCapture.HDC, BorderSize, BorderSize, Me.ScaleWidth - BorderSize * 2, 16, MenuTranslucencyLevel
    If GetMenuTranslucency Then AlphaBlending picScrollUp.HDC, 0, 0, 16, 16, picDesktopCapture.HDC, picScrollUp.Left + BorderSize, BorderSize, 16, 16, MenuTranslucencyLevel
    If GetMenuTranslucency Then AlphaBlending picScrollDown.HDC, 0, 0, 16, 16, picDesktopCapture.HDC, picScrollDown.Left + BorderSize, BorderSize, 16, 16, MenuTranslucencyLevel
    If maxitems = -1 Then
        If picFile.UBound > UBound(Contents) + 1 Then
            For i = UBound(Contents) To picFile.UBound - 1
                picFile(i).Visible = False
            Next
        Else
            For i = picFile.UBound + 1 To UBound(Contents)
                    Load picFile(i)
            Next
        End If
    Else
        If picFile.UBound > maxitems + 1 Then
            For i = UBound(Contents) To maxitems - 1
                picFile(i).Visible = False
            Next
        Else
            For i = picFile.UBound + 1 To maxitems
                    Load picFile(i)
            Next
        End If
    End If
    
    DisplayItem (0)
    If maxitems = -1 Then
        For i = 1 To UBound(Contents)
            DisplayItem (i)
        Next
    Else
        For i = 1 To maxitems
            DisplayItem (i)
        Next
    End If
    

    
    Me.Refresh
    If UBound(Contents) < 20 Then
        finalwidth = Me.Width
        finalheight = Me.Height
        finalleft = Me.Left
        Me.Width = 0
        'Me.Height = 0
        SetWindowPos Me.hWND, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        Me.Refresh
        Set tmrShow.Notify = Me
        showindex = 0
        tmrShow.Interval = 5
        tmrShow.Stats.Frequency = 5
        tmrShow.Enabled = True
    Else
        SetWindowPos Me.hWND, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        Me.Refresh
    End If
End Sub


Sub LoadSkinImages()
    picBorder(0).Picture = LoadPicture(GetSkinImage("Common Menu\TopBorder.bmp"))
    picBorder(1).Picture = LoadPicture(GetSkinImage("Common Menu\RightBorder.bmp"))
    picBorder(2).Picture = LoadPicture(GetSkinImage("Common Menu\BottomBorder.bmp"))
    picBorder(3).Picture = LoadPicture(GetSkinImage("Common Menu\LeftBorder.bmp"))
    picBorder(4).Picture = LoadPicture(GetSkinImage("Common Menu\TopLeft.bmp"))
    picBorder(5).Picture = LoadPicture(GetSkinImage("Common Menu\TopRight.bmp"))
    picBorder(6).Picture = LoadPicture(GetSkinImage("Common Menu\BottomRight.bmp"))
    picBorder(7).Picture = LoadPicture(GetSkinImage("Common Menu\BottomLeft.bmp"))
    picBorder(8).Picture = LoadPicture(GetSkinImage("Common Menu\MiddleStretch.bmp"))
    picBorder(9).Picture = LoadPicture(GetSkinImage("Common Menu\MiddleOver.bmp"))
    picTemp.Picture = LoadPicture(GetSkinImage("Common Menu\Titlebar.bmp"))
    StretchBlt picTitlebar.HDC, 0, 0, picTitlebar.Width, 16, picTemp.HDC, 0, 0, picTemp.Width, 16, vbSrcCopy
    picScrollUp.Picture = LoadPicture(GetSkinImage("Common Menu\Arrows\UpArrow.bmp"))
    picScrollDown.Picture = LoadPicture(GetSkinImage("Common Menu\Arrows\DownArrow.bmp"))
    BorderSize = picBorder(0).Height
End Sub

Sub DrawBackground()
    PxH = Me.Height / Screen.TwipsPerPixelY
    PxW = Me.Width / Screen.TwipsPerPixelX
    
    Me.Cls
    
    'Draw the background
    TilePic 8, 1, 1, picBorder(8).Width, picBorder(8).Height, PxW - BorderSize * 2, PxH - 2
    
    'Draw the border
    StretchPic 0, BorderSize, 0, picBorder(0).Width, BorderSize, PxW - BorderSize * 2, BorderSize
    StretchPic 1, PxW - BorderSize, BorderSize, BorderSize, picBorder(1).Height, BorderSize, PxH - (BorderSize * 2)
    StretchPic 2, BorderSize, PxH - BorderSize, picBorder(2).Width, BorderSize, PxW - (BorderSize * 2), BorderSize
    StretchPic 3, 0, BorderSize, BorderSize, picBorder(3).Height, BorderSize, PxH - (BorderSize * 2)

    'Draw the corners
    CopyPic 4, 0, 0, BorderSize, BorderSize
    CopyPic 5, PxW - BorderSize, 0, BorderSize, BorderSize
    CopyPic 6, PxW - BorderSize, PxH - BorderSize, BorderSize, BorderSize
    CopyPic 7, 0, PxH - BorderSize, BorderSize, BorderSize
    
    'Draw the "over" image
    picOver.Width = PxW + 10
    StretchBlt picOver.HDC, 0, 0, PxW, 17, picBorder(9).HDC, 0, 0, 198, picBorder(9).Height, vbSrcCopy
    
    picDesktopCapture.Width = Me.ScaleWidth
    picDesktopCapture.Height = Me.ScaleHeight
    BltDesktop Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, picDesktopCapture
    If GetMenuTranslucency Then AlphaBlending Me.HDC, 0, 0, Me.Width, Me.Height, picDesktopCapture.HDC, 0, 0, Me.Width, Me.Height, MenuTranslucencyLevel
End Sub

Sub CopyPic(Index As Integer, X As Integer, Y As Integer, Width As Integer, Height As Integer)
    BitBlt Me.HDC, X, Y, Width, Height, picBorder(Index).HDC, 0, 0, vbSrcCopy
End Sub

Sub StretchPic(Index As Integer, X As Integer, Y As Integer, Width As Integer, Height As Integer, NewWidth As Integer, NewHeight As Integer)
    StretchBlt Me.HDC, X, Y, NewWidth, NewHeight, picBorder(Index).HDC, 0, 0, Width, Height, vbSrcCopy
End Sub

Sub TilePic(Index As Integer, X As Integer, Y As Integer, Width As Integer, Height As Integer, NewWidth As Integer, NewHeight As Integer)
For i = 0 To Int(NewHeight / Height) - 1
    StretchBlt Me.HDC, X, Y + (i * Height), NewWidth, Height, picBorder(Index).HDC, 0, 0, Width, Height, vbSrcCopy
Next
If Int(NewHeight / Height) * Height < NewHeight Then
    StretchBlt Me.HDC, X, Int(NewHeight / Height) * Height + Y, NewWidth, NewHeight - Int(NewHeight / Height) * Height, picBorder(Index).HDC, 0, 0, Width, Height, vbSrcCopy
End If
End Sub

Function ListFolderItems(ByVal path As String) As Variant
    'returns an array of directory names
    On Error Resume Next
    Dim Count, Items(), i, ItemName ' Declare variables.
    Dim Count2, Folders()
    ItemName = Dir(path, vbDirectory Or vbArchive Or vbSystem Or vbReadOnly) ' Get first directory name.
    Count = 0
    Count2 = 0
    
    Do While Not ItemName = ""
        'A file or directory name was returned
        If Not ItemName = "." And Not ItemName = ".." Then
            If IsDir(path & ItemName) Then
                ReDim Preserve Folders(Count2 + 1)
                Folders(Count2) = ItemName ' Add directory name to array
                Count2 = Count2 + 1
            Else
                ReDim Preserve Items(Count + 1)
                Items(Count) = ItemName ' Add directory name to array
                Count = Count + 1
            End If
        End If
        ItemName = Dir ' Get another item name
    Loop
    
    If Count <> 0 Then ReDim Preserve Items(Count - 1) Else ReDim Items(0)
    If Count2 <> 0 Then ReDim Preserve Folders(Count2 - 1) Else ReDim Folders(0)
    ListFolderItems = JoinArrays(FastSort(Folders), FastSort(Items))
End Function

Function JoinArrays(Array1 As Variant, Array2 As Variant) As Variant
    Dim tmparray()
    ReDim tmparray(UBound(Array1) + UBound(Array2) + 1)
    If UBound(Array1) = 0 Then
        JoinArrays = Array2
        Exit Function
    End If

    If UBound(Array2) = 0 Then
        JoinArrays = Array1
        Exit Function
    End If
    
    For i = 0 To UBound(Array1)
        tmparray(i) = Array1(i)
    Next
    
    For i = 0 To UBound(Array2)
        tmparray(i + UBound(Array1) + 1) = Array2(i)
    Next
    JoinArrays = tmparray
End Function

Sub DisplayItem(Index As Integer, Optional Over As Boolean = False)
    If Index + curScroll > UBound(Contents) Then
        picFile(Index).Visible = False
        Exit Sub
    End If
    
    picFile(Index).Top = 16 * Index + BorderSize + 16
    picFile(Index).Width = MaxWidth + 30 - BorderSize
    picFile(Index).Left = BorderSize
    picFile(Index).Visible = True
    If Over Then
        picFile(Index).ForeColor = GetOverColor
        BitBlt picFile(Index).HDC, 0, 0, PxW, 16, picOver.HDC, 0, 0, vbSrcCopy
        If GetMenuTranslucency Then AlphaBlending picFile(Index).HDC, 0, 0, PxW, 16, picDesktopCapture.HDC, 1, BorderSize + (Index + 1) * 16, PxW, 16, MenuTranslucencyLevel
    Else
        picFile(Index).ForeColor = GetNormalColor
        picFile(Index).Cls
        BitBlt picFile(Index).HDC, 0, 0, PxW, 16, Me.HDC, 1, BorderSize + (Index + 1) * 16, vbSrcCopy
    End If
    picFile(Index).CurrentX = 20
    picFile(Index).CurrentY = 0
    picFile(Index).Print RemoveExtention(Contents(Index + curScroll))
    DrawFileIcon CurrentPath & Contents(Index + curScroll), picFile(Index).HDC, Icon_Small
End Sub


Private Property Get ccrpTimer_Enabled() As Boolean

End Property

Private Property Let ccrpTimer_Enabled(ByVal RHS As Boolean)

End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            HideMenu
            If ParentIndex <> -1 Then Forms(ParentIndex).PopupClosed
    End Select
End Sub

Sub PopupClosed()
    CurrentPopup = -1
End Sub

Private Sub Form_Load()
    Dim tmrShow As New ccrpTimer
    Set tmrShow.Notify = Me
    tmrShow.Interval = 10
End Sub

Private Sub ICcrpCountdownNotify_Tick(ByVal TimeRemaining As Long)
    
End Sub

Private Sub ICcrpTimerNotify_Timer(ByVal Milliseconds As Long)
    showindex = showindex + 1
    Me.Width = finalwidth * (showindex / 25)
    Me.Refresh
    If showindex = 25 Then
        tmrShow.Enabled = False
        SetWindowPos Me.hWND, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    End If
End Sub

Private Sub picFile_Click(Index As Integer)
    ShellFile (CurrentPath & Contents(Index + curScroll))
    If DesktopIconsShown Then ToggleDesktopIcons Else ToggleStartMenu
End Sub

Private Sub picFile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If OldIndex <> -1 And OldIndex <> Index Then
    Out (OldIndex)
End If

If Index <> -1 And CurrentIndex <> Index Then
    Over (Index)
    If IsDir(CurrentPath & Contents(Index + curScroll)) Then
        HideSubmenus
        PopupIndex = Index
        tmrPopup.Enabled = True
        peviewindex = -1
    Else
        HideSubmenus
        CurrentPopup = -1
        PopupIndex = -1
        PreviewIndex = Index
        tmrPopup.Enabled = True
        PlaySkinSound "menuitemhover"
    End If
    OldIndex = Index
    CurrentIndex = Index
End If
End Sub

Sub HideSubmenus()
        If CurrentPopup <> -1 Then
            Forms(CurrentFormIndex + 1).HideMenu
            CurrentPopup = -1
        End If
End Sub

Sub Over(Index As Integer)
    picFile(Index).Cls
    DisplayItem Index, True
End Sub

Sub Out(Index As Integer)
    On Error Resume Next
    picFile(Index).Cls
    DisplayItem Index, False
End Sub

Private Function AddASlash(path As String)
    If Right(path, 1) = "\" Then AddASlash = path Else AddASlash = path & "\"
End Function

Private Function RemoveExtention(ByVal file As String) As String
    file = StrReverse(file)
    pos = InStr(file, ".")
    file = Right(file, Len(file) - pos)
    file = StrReverse(file)
    RemoveExtention = file
End Function

Public Function HideMenu()
    tmrShow.Enabled = False
    If CurrentPopup <> -1 Then Forms(CurrentFormIndex + 1).HideMenu
    If CurrentIndex <> -1 Then Out (CurrentIndex)
    If OldIndex <> -1 Then Out (OldIndex)
    Unload Me
End Function

Private Function IsDir(ByVal file As String) As Boolean
    IsDir = (GetAttr(file) And vbDirectory)
End Function



Private Sub picScrollDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picScrollDown.Picture = LoadPicture(GetSkinImage("Common Menu\Arrows\DownArrowDown.bmp"))
    If GetMenuTranslucency Then AlphaBlending picScrollDown.HDC, 0, 0, 16, 16, picDesktopCapture.HDC, picScrollDown.Left + BorderSize, BorderSize, 16, 16, MenuTranslucencyLevel
End Sub

Private Sub picScrollDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideSubmenus
    Out (CurrentIndex)
End Sub

Private Sub picScrollDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picScrollDown.Picture = LoadPicture(GetSkinImage("Common Menu\Arrows\DownArrow.bmp"))
    If GetMenuTranslucency Then AlphaBlending picScrollDown.HDC, 0, 0, 16, 16, picDesktopCapture.HDC, picScrollDown.Left + BorderSize, BorderSize, 16, 16, MenuTranslucencyLevel
    picScrollDown.Refresh
    ScrollDown
End Sub

Private Sub picScrollUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picScrollUp.Picture = LoadPicture(GetSkinImage("Common Menu\Arrows\UpArrowDown.bmp"))
    If GetMenuTranslucency Then AlphaBlending picScrollUp.HDC, 0, 0, 16, 16, picDesktopCapture.HDC, picScrollUp.Left + BorderSize, BorderSize, 16, 16, MenuTranslucencyLevel
End Sub

Private Sub picScrollUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideSubmenus
    Out (CurrentIndex)
End Sub

Private Sub picScrollUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picScrollUp.Picture = LoadPicture(GetSkinImage("Common Menu\Arrows\UpArrow.bmp"))
    If GetMenuTranslucency Then AlphaBlending picScrollUp.HDC, 0, 0, 16, 16, picDesktopCapture.HDC, picScrollUp.Left + BorderSize, BorderSize, 16, 16, MenuTranslucencyLevel
    picScrollUp.Refresh
    ScrollUp
End Sub

Private Sub picTitlebar_DblClick()
If DesktopIconsShown Then ToggleDesktopIcons
If StartMenuShown Then ToggleStartMenu
ShellFile CurrentPath
End Sub

Private Sub picTitlebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNow = True
    If CurrentFormIndex <> Forms.Count - 1 Then Forms(Forms.Count - 1).SetFocus: DragNow = False
    TX = X
    TY = Y
End Sub

Private Sub picTitlebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideSubmenus
    Out (CurrentIndex)
    If DragNow Then
        If X - TX < -15 Or X - TX > 15 Or Y - TY < -15 Or Y - TY > 15 Then
            DragNow = False
            frmFolderWindow.DrawMenu CurrentPath, Me.Left + X, Me.Top + Y, 0
            frmFolderWindow.DragMe TX, TY
            If StartMenuShown Then ToggleStartMenu Else ToggleDesktopIcons
        End If
    End If
End Sub

Private Sub picTitlebar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNow = False
End Sub

Private Sub tmrPopup_Timer()
    If PopupIndex <> -1 Then
        Dim f As New frmFolderMenu
        f.DrawMenu CurrentPath & Contents(PopupIndex + curScroll), Me.Left + Me.Width - 80, Me.Top + (PopupIndex * 16 + 1) * Screen.TwipsPerPixelY, CurrentFormIndex
        CurrentPopup = PopupIndex
        PlaySkinSound "menuopen"
    End If
'    If PreviewIndex <> -1 Then
'        ToolTipEX "", (Me.Top / Screen.TwipsPerPixelY) + PopupIndex * 16, (Me.Left + Me.Width) / Screen.TwipsPerPixelX, CurrentPath & Contents(PreviewIndex + curScroll)
'    Else
'        HideToolTip
'    End If
    tmrPopup.Enabled = False
End Sub

Sub ScrollDown()
    curScroll = curScroll + picFile.Count
    If curScroll + UBound(Contents) >= UBound(Contents) Then picScrollDown.Enabled = False
    For i = 0 To picFile.Count - 1
        DisplayItem (i)
    Next
    picScrollUp.Enabled = True
End Sub

Sub ScrollUp()
    curScroll = curScroll - picFile.Count
    If curScroll <= 0 Then picScrollUp.Enabled = False
    For i = 0 To picFile.Count - 1
        DisplayItem (i)
    Next
    picScrollDown.Enabled = True
End Sub
