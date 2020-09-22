VERSION 5.00
Begin VB.Form frmTaskbar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5280
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   9915
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   661
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   10
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   11
      Left            =   2220
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   21
      Top             =   4380
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picMyComputerButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   420
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   20
      Top             =   0
      Width           =   450
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   9
      Left            =   300
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   19
      Top             =   4140
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picResize 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   9795
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5280
      ScaleWidth      =   120
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picGripper 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   9675
      MousePointer    =   5  'Size
      ScaleHeight     =   5280
      ScaleWidth      =   120
      TabIndex        =   17
      Top             =   0
      Width           =   120
   End
   Begin VB.PictureBox picQuickstart 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1800
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   16
      Top             =   0
      Width           =   5280
      Begin VB.Image imgIcon 
         Height          =   255
         Index           =   0
         Left            =   60
         OLEDropMode     =   1  'Manual
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4500
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrCheckFocus 
      Interval        =   100
      Left            =   4920
      Top             =   2400
   End
   Begin VB.Timer tmrUpdateTime 
      Interval        =   5000
      Left            =   4320
      Top             =   2400
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   8
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   7
      Left            =   840
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   6
      Left            =   780
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   11
      Top             =   3660
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picTray 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   8175
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   10
      Top             =   0
      Width           =   1500
      Begin VB.Image imgTrayIcon 
         Height          =   255
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   255
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3:00 pm"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   14
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   300
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picProgram 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   2160
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   7
      Top             =   0
      Width           =   2055
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   300
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Top             =   120
         Width           =   1215
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picDesktopButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   900
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   5
      Top             =   0
      Width           =   870
   End
   Begin VB.PictureBox picStartButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   4
      Top             =   0
      Width           =   450
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   2580
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   1380
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   795
      Width           =   1215
   End
End
Attribute VB_Name = "frmTaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  'Copyright (C) 2000 BSoft
'Developed by Brian and Florian
'B2 is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

'Declerations for the "App-Tray"
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" _
        (pDicDesc As IconType, riid As CLSIdType, ByVal fOwn As Long, _
        lpUnk As Object) As Long
        
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias _
        "SHGetFileInfoA" (ByVal pszPath As String, ByVal _
        dwFileAttributes As Long, psfi As ShellFileInfoType, ByVal _
        cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type IconType
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Private Type CLSIdType
    ID(16) As Byte
End Type

Private Type ShellFileInfoType
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Const Large = &H100
Const Small = &H101

'Taskbar Settings (loaded from skin)
Public TaskbarHeight As Integer 'Height of the taskbar (in pixels)
Public StartButtonWidth As Integer 'Width of the start button (in pixels)
Public DesktopButtonWidth As Integer 'Width of the "desktop popup" button
Public TaskbarButtonWidth As Integer
Public MyComputerWidth As Integer

'Taskbar Settings (determined automaticlly)
Public TaskbarWidth As Integer 'Width of the taskbar (in pixels)
Public TaskbarButtons As Integer
Public MaxTaskLength As Integer

'Program list variables
Dim ButtonVisible() As Boolean
Dim ButtonDown() As Boolean
Dim ButtonCaption() As String
Dim AppTrayPath As String 'Path for the "quickstart" menu
Dim AppTrayCurrentFilename As String, Number1 As Integer, Index1 As Integer
Dim AppTrayOverIndex As Integer
Dim ProgName As String

Dim ProgramOverIndex As Integer

Dim StartButtonOver As Boolean
Dim DesktopButtonOver As Boolean
Dim ClockOver As Boolean
Dim GripperOver As Boolean
Dim ResizeOver As Boolean

Private TX As Integer, TY As Integer, DragNow As Boolean
Public CurTop As Integer, CurLeft As Integer, OldWidth As Integer
Private ResizeNow As Boolean
Private PrevWidth As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then UpdateIcons
End Sub

Private Sub Form_Load()
    AppTrayPath = App.path & "\quickstart\"   'Set Path to the Quickstart programs
    LoadSkinSettings 'Get taskbar settings determined by the skin
    GetTaskbarSettings 'Get taskbar settings not determined by the skin
    LoadSkinImages 'Load the pictures for the taskbar
    
    picDesktopCapture.Width = ScreenWidth: picDesktopCapture.Height = ScreenHeight
    BltDesktop 0, ScreenHeight - TaskbarHeight, picDesktopCapture
    
    ApplyTaskbarSettings 'Apply the taskbar settings
    ResizeDesktop
End Sub


Sub GetTaskbarSettings()
    TaskbarWidth = ScreenWidth
    TaskbarButtons = Int((TaskbarWidth - DesktopButtonWidth - MyComputerWidth - StartButtonWidth - picTray.Width - picQuickstart.Width) / TaskbarButtonWidth)
    MatchFonts
    For i = 10 To 100
        If picTaskbarImage(0).TextWidth(String(i, " ")) > TaskbarWidth - 20 Then Exit For
    Next
    MaxTaskLength = i
    ReDim ButtonVisible(TaskbarButtons)
    ReDim ButtonDown(TaskbarButtons)
    ReDim ButtonCaption(TaskbarButtons)
End Sub

Sub MatchFonts()
    picTaskbarImage(0).FontName = lblCaption(0).FontName
    picTaskbarImage(0).FontSize = lblCaption(0).FontSize
    picTaskbarImage(0).FontBold = lblCaption(0).FontBold
End Sub

Sub LoadSkinSettings()
    TaskbarHeight = GetTaskbarHeight
    StartButtonWidth = GetStartButtonWidth
    DesktopButtonWidth = GetDesktopButtonWidth
    TaskbarButtonWidth = GetTaskbarButtonWidth
    MyComputerWidth = GetMyComputerWidth
End Sub

Sub LoadSkinImages()
    picProgram(0).Picture = LoadPicture(GetSkinImage("Taskbar\Program.bmp")) 'Load the picture for the program button
    picTaskbarImage(0).Picture = LoadPicture(GetSkinImage("Taskbar\TaskbarBG.bmp")) 'Load the picture for the taskbar background
    picTaskbarImage(1).Picture = LoadPicture(GetSkinImage("Taskbar\startbutton.bmp")) 'Load the picture for the start button
    picTaskbarImage(2).Picture = LoadPicture(GetSkinImage("Taskbar\startbuttondown.bmp")) 'Load the picture for the pressed start button
    picTaskbarImage(3).Picture = LoadPicture(GetSkinImage("Taskbar\desktop.bmp")) 'Load the icon for the desktop popup
    picTaskbarImage(4).Picture = LoadPicture(GetSkinImage("Taskbar\desktopdown.bmp")) 'Load the icon for the pressed dekstop popup
    picTaskbarImage(5).Picture = LoadPicture(GetSkinImage("Taskbar\ProgramDown.bmp")) 'Load the picture for the pressed program button
    picTaskbarImage(6).Picture = LoadPicture(GetSkinImage("Taskbar\SystemTray\Left.bmp")) 'Load the picture for the left side of the tray
    picTaskbarImage(7).Picture = LoadPicture(GetSkinImage("Taskbar\SystemTray\Center.bmp")) 'Load the picture to be streatched across the middle of the tray
    picTaskbarImage(8).Picture = LoadPicture(GetSkinImage("Taskbar\SystemTray\Right.bmp")) 'Load the picture for the right side of the tray
    picTaskbarImage(9).Picture = LoadPicture(GetSkinImage("Taskbar\gripper.bmp"))  'Load the picture for the "gripper"
    picTaskbarImage(10).Picture = LoadPicture(GetSkinImage("Taskbar\MyComputer.bmp")) 'Load the picture for the pressed my computer button
    picTaskbarImage(11).Picture = LoadPicture(GetSkinImage("Taskbar\MyComputerDown.bmp")) 'Load the icon for the pressed my computer button
End Sub

Sub ApplyTaskbarSettings()
    CurTop = ScreenHeight - TaskbarHeight
    CurLeft = 0
    OldWidth = 800
    
    SetWindowPos Me.hWND, -1, 0, ScreenHeight - TaskbarHeight, ScreenWidth, TaskbarHeight, 0  'Move the window into place and bring it to the top
    
    StretchBlt Me.HDC, 0, 0, ScreenWidth, TaskbarHeight, picTaskbarImage(0).HDC, 0, 0, picTaskbarImage(0).Width, TaskbarHeight, vbSrcCopy
    
    picStartButton.Height = TaskbarHeight 'Set the height of the start button
    picStartButton.Width = StartButtonWidth 'Set the width of the start button

    picDesktopButton.Height = TaskbarHeight 'Set the height of the desktop popup icon
    picDesktopButton.Left = StartButtonWidth
    picDesktopButton.Width = DesktopButtonWidth 'Set the width of the desktop popup icon
    
    picMyComputerButton.Height = TaskbarHeight
    picMyComputerButton.Left = StartButtonWidth + DesktopButtonWidth
    picMyComputerButton.Width = MyComputerWidth
    
    picQuickstart.Left = MyComputerWidth + DesktopButtonWidth + StartButtonWidth
    
    UpdateIcons
    
    lblCaption(0).Width = TaskbarButtonWidth - 20
    picProgram(0).Width = TaskbarButtonWidth
    picProgram(0).Left = picQuickstart.Left + picQuickstart.Width + 2
    ButtonVisible(0) = True
    For i = 1 To TaskbarButtons
        Load picProgram(i)
        picProgram(i).Left = picProgram(i - 1).Left + picProgram(i - 1).Width
        picProgram(i).Visible = True
        picProgram(i).Width = TaskbarButtonWidth
        ButtonVisible(i) = True
        Load lblCaption(i)
        Set lblCaption(i).Container = picProgram(i)
        lblCaption(i).Width = TaskbarButtonWidth - 20
        lblCaption(i).Visible = True
    Next
    UpdateTranslucency False
    UpdateTaskbar
    UpdateTime
End Sub

Sub UpdateIcons()
AppTrayOverIndex = -1
If AppTrayPath = "" Then
picQuickstart.Width = 0
picQuickstart.Visible = False
Exit Sub
End If
AppTrayCurrentFilename = Dir(AppTrayPath, vbNormal)   'Get first file
If AppTrayCurrentFilename <> "" Then
'    If Number1 > 0 Then
'        For n = 1 To Number1
'            Unload imgIcon(n)
'            picQuickstart.Picture = LoadPicture()
'        Next n
'    End If

Number1 = -1
Do While AppTrayCurrentFilename <> ""
   'Ignore actual and higher directory
   If AppTrayCurrentFilename <> "." And AppTrayCurrentFilename <> ".." Then
      'Be sure that AppTrayCurrentFilename is not a directory
      If (GetAttr(AppTrayPath & AppTrayCurrentFilename)) <> vbDirectory Then
      Number1 = Number1 + 1
        If Number1 > 0 Then
        If Number1 > imgIcon.UBound Then
            Load imgIcon(Number1)
            imgIcon(Number1).Left = imgIcon(Number1 - 1).Left + imgIcon(Number1 - 1).Width + 3
            imgIcon(Number1).Picture = LoadIcon(Small)
            imgIcon(Number1).Tag = AppTrayCurrentFilename
            imgIcon(Number1).Visible = True
        End If
        Else
        imgIcon(0).Picture = LoadIcon(Small)
        imgIcon(0).Tag = AppTrayCurrentFilename
        End If
      End If
   End If
   AppTrayCurrentFilename = Dir   'Get next file
Loop
picQuickstart.Width = imgIcon(Number1).Left + imgIcon(Number1).Width + 2

'Set background (needed for updating icons)
picQuickstart.Cls
BitBlt picQuickstart.HDC, 0, 0, 5, TaskbarHeight, picTaskbarImage(6).HDC, 0, 0, vbSrcCopy
StretchBlt picQuickstart.HDC, 5, 0, picQuickstart.Width, TaskbarHeight, picTaskbarImage(7).HDC, 0, 0, 1, TaskbarHeight, vbSrcCopy
BitBlt picQuickstart.HDC, picQuickstart.Width - 5, 0, 5, TaskbarHeight, picTaskbarImage(8).HDC, 0, 0, vbSrcCopy
If GetTranslucent Then AlphaBlending picQuickstart.HDC, 0, 0, picQuickstart.Width, TaskbarHeight, picDesktopCapture.HDC, picQuickstart.Left, 0, picQuickstart.Width, TaskbarHeight, TranslucencyLevel
Else
'If there are no files present don't show the "App-Tray"
picQuickstart.Width = 0
picQuickstart.Visible = False
End If
'Set the positions for the program buttons
For n = 0 To picProgram.Count - 1
If n = 0 Then
picProgram(n).Left = picQuickstart.Left + picQuickstart.Width + 2
Else
picProgram(n).Left = picProgram(n - 1).Left + picProgram(n - 1).Width
End If
Next n
End Sub

'Get the icons for the "App-Tray
Private Function LoadIcon(Size&) As IPictureDisp
  Dim Result&, file$, Slash$
  Dim Unkown As IUnknown
  Dim Icon As IconType
  Dim CLSID As CLSIdType
  Dim ShellInfo As ShellFileInfoType
    
    file = AppTrayPath & AppTrayCurrentFilename
    Call SHGetFileInfo(file, 0, ShellInfo, Len(ShellInfo), Size)
 
    Icon.cbSize = Len(Icon)
    Icon.picType = vbPicTypeIcon
    Icon.hIcon = ShellInfo.hIcon
    CLSID.ID(8) = &HC0
    CLSID.ID(15) = &H46
    Result = OleCreatePictureIndirect(Icon, CLSID, 1, Unkown)
    
    Set LoadIcon = Unkown
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideTaskbarTips
End Sub

'Open program in the "App-Tray"
Private Sub imgIcon_DblClick(Index As Integer)

End Sub

Private Sub imgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    imgIcon(Index).Top = imgIcon(Index).Top + 1
    imgIcon(Index).Left = imgIcon(Index).Left + 1
    HideToolTip
Else
    CurrentSelectedTaskbarItem = Index
    HideToolTip
    PopupMenu frmInfo.mnuTaskbarIconMenu, , , , frmInfo.mnuOpen
End If
End Sub

Private Sub imgIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ProgramOverIndex <> -1 Then ProgramOverIndex = -1: HideToolTip
    If AppTrayOverIndex <> Index Then
        StartButtonOver = False
        HideTaskbarTips
        AppTrayOverIndex = Index
        ProgName = imgIcon(Index).Tag
        If Right(ProgName, 4) = ".lnk" Then
            ToolTipEX Left(ProgName, Len(ProgName) - 4), ScreenHeight - TaskbarHeight - 20, StartButtonWidth + DesktopButtonWidth + 16 * Index
        Else
            ToolTipEX ProgName, ScreenHeight - TaskbarHeight - 20, StartButtonWidth + DesktopButtonWidth + 16 * Index
        End If
    End If
End Sub

Private Sub imgIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    ShellExecute Me.hWND, "open", AppTrayPath & imgIcon(Index).Tag, "", "", 1
    imgIcon(Index).Top = imgIcon(Index).Top - 1
    imgIcon(Index).Left = imgIcon(Index).Left - 1
End If
End Sub

Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ActivateWindow (modWindowAPICalls.WindowID(Index))
        UpdateTaskbar
    ElseIf Button = vbRightButton Then
        HideTaskbarTips
        PopupMenu frmInfo.mnuTaskMenu
    End If
End Sub

Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If AppTrayOverIndex <> -1 Then AppTrayOverIndex = -1: HideToolTip
    If ProgramOverIndex = -1 Then HideTaskbarTips
    If ProgramOverIndex <> Index Then
        HideToolTip
        ProgramOverIndex = Index
        ToolTipEX ButtonCaption(Index), CurTop - 24, Index * TaskbarButtonWidth + StartButtonWidth + DesktopButtonWidth + picQuickstart.Width + CurLeft
    End If
End Sub

Private Sub lblCaption_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmInfo.DisplayInfo "You cannot drag an item onto a taskbar button. However, if you drag an item over a taskbar button, the program that that button represents will come to the front.", "Taskbar drag problem", sndError
End Sub

Private Sub lblCaption_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    lblCaption_MouseDown Index, 0, 0, 0, 0
End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ClockOver = False Then
        ToolTipEX Now, ScreenHeight - TaskbarHeight - 24, ScreenWidth - 150
        ClockOver = True
    End If
End Sub

Private Sub picDesktopButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If StartMenuShown Then
        ToggleStartMenu
    End If
    HideTaskbarTips
    ToggleDesktopIcons
End Sub

Private Sub picDesktopButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DesktopButtonOver = False Then HideTaskbarTips: DesktopButtonOver = True
    If ToolTipDisplayed = False And DesktopIconsShown = False Then
        ToolTipEX "Click here to display a list of your desktop icons", ScreenHeight - TaskbarHeight - 24, StartButtonWidth
    End If
End Sub

Private Sub picGripper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideTaskbarTips
    TX = X
    TY = Y
    DragNow = True
End Sub

Private Sub picGripper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragNow Then
    CurTop = CurTop + ((Y - TY) / Screen.TwipsPerPixelY)
    CurLeft = CurLeft + ((X - TX) / Screen.TwipsPerPixelX)
    If CurTop > ScreenHeight - TaskbarHeight - 15 Then
        CurTop = ScreenHeight - TaskbarHeight
        CurLeft = 0
        TaskbarWidth = ScreenWidth
        If picResize.Visible = True Then
            picResize.Visible = False
            UpdateWidth 0, 0, False
            PlaySkinSound "TaskbarDock"
        End If
    Else
        TaskbarWidth = OldWidth
        If picResize.Visible = False Then
            picResize.Visible = True
            picResize.Left = ScreenWidth
            UpdateWidth 0, 0, 0
            PlaySkinSound "TaskbarDock"
        End If
    End If
    
    SetWindowPos Me.hWND, HWND_TOPMOST, CurLeft, CurTop, TaskbarWidth, TaskbarHeight, 0
Else
    If GripperOver = False Then
        ToolTipEX "Drag here to move the B2 Taskbar", CurTop - 24, CurLeft + ScaleWidth - 220
        GripperOver = True
    End If
End If
End Sub

Private Sub picGripper_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNow = False
    If GetTranslucent Then
        UpdateTranslucency True
    Else
        UpdateTranslucency False
    End If
End Sub

Public Sub UpdateTranslucency(Optional ReCapture As Boolean = True)
    PrevWidth = TaskbarWidth
    If ReCapture Then
        SetWindowPos Me.hWND, HWND_TOPMOST, 0, -100, 0, 0, SWP_NOSIZE 'Hide me
        DoEvents
        BltDesktop CurLeft, CurTop, picDesktopCapture 'Capture desktop
        SetWindowPos Me.hWND, HWND_TOPMOST, CurLeft, CurTop, 0, 0, SWP_NOSIZE 'Show me
    End If
    
    UpdateWidth 0, 0, False 'Fix program button translucency
    StartButtonOff 'Fix start button translucency
    DesktopIconOff 'Fix desktop popup translucency
    MyComputerOff 'Fix My Computer translucency
    'Fix form translucency
        Me.Cls
        StretchBlt Me.HDC, 0, 0, ScreenWidth, TaskbarHeight, picTaskbarImage(0).HDC, 0, 0, 1, TaskbarHeight, vbSrcCopy
        If GetTranslucent Then AlphaBlending Me.HDC, 0, 0, ScreenWidth - CurLeft, TaskbarHeight, picDesktopCapture.HDC, 0, 0, ScreenWidth - CurLeft, TaskbarHeight, TranslucencyLevel
    ResizeTray 100 'Fix tray translucency
    UpdateIcons 'Update Quikclaunch tray translucency
    'Update Gripper & ResizeBar Translucency
        picGripper.Cls
        BitBlt picGripper.HDC, 0, 0, 8, TaskbarHeight, picTaskbarImage(9).HDC, 0, 0, vbSrcCopy
        If picResize.Visible Then
            picResize.Cls
            BitBlt picResize.HDC, 0, 0, 8, TaskbarHeight, picTaskbarImage(9).HDC, 0, 0, vbSrcCopy
            If GetTranslucent Then
                AlphaBlending picGripper.HDC, 0, 0, 8, TaskbarHeight, picDesktopCapture.HDC, TaskbarWidth - 16, 0, 8, TaskbarHeight, TranslucencyLevel
                AlphaBlending picResize.HDC, 0, 0, 8, TaskbarHeight, picDesktopCapture.HDC, TaskbarWidth - 8, 0, 8, TaskbarHeight, TranslucencyLevel
            End If
        Else
            If GetTranslucent Then AlphaBlending picGripper.HDC, 0, 0, 8, TaskbarHeight, picDesktopCapture.HDC, TaskbarWidth - 8, 0, 8, TaskbarHeight, TranslucencyLevel
        End If
End Sub

Private Sub picMyComputerButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MyComputerOn
End Sub

Private Sub picMyComputerButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MyComputerOff
    Shell "explorer.exe ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", vbNormalFocus
End Sub

Private Sub picProgram_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCaption_MouseDown Index, 0, 0, 0, 0
End Sub

Private Sub picProgram_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCaption_MouseMove Index, 0, 0, 0, 0
End Sub

Private Sub picProgram_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmInfo.DisplayInfo "You cannot drag an item onto a taskbar button. However, if you drag an item over a taskbar button, the program that that button represents will come to the front.", "Taskbar drag problem"
End Sub

Private Sub picProgram_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    lblCaption_MouseDown Index, 0, 0, 0, 0
End Sub

Private Sub picQuickstart_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFFiles) Then
        For i = 1 To Data.Files.Count
            If GetAttr(Data.Files(i)) = vbDirectory Then
                Debug.Print "Directory " & Data.Files(i)
            Else
                FileCopy Data.Files(i), App.path & "\QuickStart\" & GetFileName(Data.Files(i))
                Debug.Print Data.Files(i) & " > " & App.path & "\QuickStart\" & GetFileName(Data.Files(i))
            End If
        Next
    End If
End Sub

Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResizeNow = True
    HideTaskbarTips
End Sub

Private Sub picResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ResizeNow Then
        UpdateWidth (X - TX) / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY, True
    Else
        If ResizeOver = False Then
            ToolTipEX "Drag here to resize the B2 Taskbar", CurTop - 24, CurLeft + ScaleWidth - 210
            ResizeOver = True
        End If
    End If
    UpdateTranslucency False
End Sub

Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResizeNow = False
End Sub

Private Sub picStartButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DesktopIconsShown Then
        ToggleDesktopIcons
    End If
    ToggleStartMenu
End Sub

Sub StartButtonOn()
    BitBlt picStartButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picTaskbarImage(2).HDC, 0, 0, vbSrcCopy 'Copy the image for the start button
    If GetTranslucent Then AlphaBlending picStartButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picDesktopCapture.HDC, 0, 0, StartButtonWidth, TaskbarHeight, TranslucencyLevel
    picStartButton.Refresh
End Sub

Sub StartButtonOff()
    BitBlt picStartButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picTaskbarImage(1).HDC, 0, 0, vbSrcCopy 'Copy the image for the start button
    If GetTranslucent Then AlphaBlending picStartButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picDesktopCapture.HDC, 0, 0, StartButtonWidth, TaskbarHeight, TranslucencyLevel
    picStartButton.Refresh
End Sub

Sub DesktopIconOn()
    BitBlt picDesktopButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picTaskbarImage(4).HDC, 0, 0, vbSrcCopy 'Copy the image for the start button
    If GetTranslucent Then AlphaBlending picDesktopButton.HDC, 0, 0, DesktopButtonWidth, TaskbarHeight, picDesktopCapture.HDC, StartButtonWidth, 0, DesktopButtonWidth, TaskbarHeight, TranslucencyLevel
    picDesktopButton.Refresh
End Sub

Sub DesktopIconOff()
    BitBlt picDesktopButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picTaskbarImage(3).HDC, 0, 0, vbSrcCopy 'Copy the image for the start button
    If GetTranslucent Then AlphaBlending picDesktopButton.HDC, 0, 0, DesktopButtonWidth, TaskbarHeight, picDesktopCapture.HDC, StartButtonWidth, 0, DesktopButtonWidth, TaskbarHeight, TranslucencyLevel
    picDesktopButton.Refresh
End Sub

Private Sub picStartButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If StartButtonOver = False Then HideTaskbarTips: StartButtonOver = True
    If StartMenuShown = False And ToolTipDisplayed = False Then
        ToolTipEX "Click here to display the B2 Start Menu", ScreenHeight - TaskbarHeight - 24, 0
    End If
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HideTaskbarTips
End Sub

Private Sub tmrCheckFocus_Timer()
    Dim WindowName As Variant, WindowHwnd As Variant, AWI As Integer
    If StartMenuShown And NotOnStartMenu Then ToggleStartMenu
    If DesktopIconsShown And NotInDesktopIcons Then ToggleDesktopIcons
    UpdateTaskbar
End Sub

Function NotOnStartMenu() As Boolean
    a = GetForegroundWindow
    For i = frmFolderMenu.GetFormIndex To Forms.Count - 1
        If Forms(i).hWND = a Then
            NotOnStartMenu = False
            Exit Function
        End If
    Next
    If a = frmStartMenu.hWND Or a = frmShutdownMenu.hWND Or a = frmSettingsMenu.hWND Or a = frmHelpMenu.hWND Or a = frmFindMenu.hWND Then
        NotOnStartMenu = False
    Else
        NotOnStartMenu = True
    End If
End Function

Function NotInDesktopIcons() As Boolean
    a = GetForegroundWindow
    For i = frmFolderMenu.GetFormIndex To Forms.Count - 1
        If Forms(i).hWND = a Then
            NotInDesktopIcons = False
            Exit Function
        End If
    Next
End Function

Sub UpdateTaskbar()
    Programs = modWindowAPICalls.GetWindows
    For i = 0 To TaskbarButtons
        If i >= UBound(modWindowAPICalls.WindowID) Then
            If ButtonVisible(i) Then DisableButton (i)
        Else
            If ButtonVisible(i) = False Then EnableButton (i)
            If ButtonCaption(i) <> modWindowAPICalls.WindowName(i) Then
                UpdateButtonCaption (i)
                picProgram(i).Cls
                If ButtonDown(i) Then MakeButtonDown (i) Else UpdateButtonIcon (i): MakeTaskbarButtonTranslucent (i)
            End If
            If i = modWindowAPICalls.AWI Then
                If ButtonDown(i) = False Then MakeButtonDown (i)
            Else
                If ButtonDown(i) Then MakeButtonUp (i)
            End If
        End If
    Next
End Sub

Sub DisableButton(Index As Integer)
    picProgram(Index).Visible = False
    ButtonVisible(Index) = False
End Sub

Sub EnableButton(Index As Integer)
    picProgram(Index).Visible = True
    ButtonVisible(Index) = True
    picProgram(Index).Cls
    MakeTaskbarButtonTranslucent (Index)
End Sub

Sub MakeButtonDown(Index As Integer)
    ButtonDown(Index) = True
    BitBlt picProgram(Index).HDC, 0, 0, TaskbarButtonWidth, TaskbarHeight, picTaskbarImage(5).HDC, 0, 0, vbSrcCopy
    UpdateButtonIcon (Index)
    picProgram(Index).Refresh
    lblCaption(Index).Top = 9
    lblCaption(Index).Left = 21
    MakeTaskbarButtonTranslucent (Index)
End Sub

Sub MakeButtonUp(Index As Integer)
    ButtonDown(Index) = False
    picProgram(Index).Cls
    UpdateButtonIcon (Index)
    lblCaption(Index).Top = 8
    lblCaption(Index).Left = 20
    UpdateButtonIcon (Index)
    MakeTaskbarButtonTranslucent (Index)
End Sub

Sub MakeTaskbarButtonTranslucent(Index As Integer)
    If GetTranslucent Then AlphaBlending picProgram(Index).HDC, 0, 0, TaskbarButtonWidth, TaskbarHeight, picDesktopCapture.HDC, picProgram(Index).Left, 0, TaskbarButtonWidth, TaskbarHeight, TranslucencyLevel
    picProgram(Index).Refresh
End Sub

Sub UpdateButtonCaption(Index As Integer)
    cap = modWindowAPICalls.WindowName(Index)
    ButtonCaption(Index) = cap
    If picProgram(Index).TextWidth(cap) > picProgram(Index).Width - 80 Then
        cap = Left(cap, MaxTaskLength - 3) & "..."
    End If
    lblCaption(Index).caption = cap
End Sub

Sub UpdateButtonIcon(Index As Integer)
Down = ButtonDown(Index)
If Down Then
    DrawIcon picProgram(Index).HDC, modWindowAPICalls.WindowID(Index), 6, 7
Else
    DrawIcon picProgram(Index).HDC, modWindowAPICalls.WindowID(Index), 5, 6
End If
End Sub

Sub DrawIcon(HDC As Long, hWND As Long, X As Integer, Y As Integer, Optional largesize As Boolean = False)
        ico = GetIcon(hWND)
        If largesize Then
        
        Else
            DrawIconEx HDC, X, Y, ico, 16, 16, 0, 0, DI_NORMAL
        End If
End Sub

Sub ResizeTray(NewWidth As Integer)
    picTray.Cls
    picTray.Width = NewWidth
    BitBlt picTray.HDC, 0, 0, 5, TaskbarHeight, picTaskbarImage(6).HDC, 0, 0, vbSrcCopy
    StretchBlt picTray.HDC, 5, 0, NewWidth - 2, TaskbarHeight, picTaskbarImage(7).HDC, 0, 0, 1, TaskbarHeight, vbSrcCopy
    BitBlt picTray.HDC, NewWidth - 2, 0, 5, TaskbarHeight, picTaskbarImage(8).HDC, 0, 0, vbSrcCopy
    'StretchBlt picTray.HDC, NewWidth, 0, 3, TaskbarHeight, picTaskbarImage(0).HDC, 0, 0, 1, TaskbarHeight, vbSrcCopy
    lblTime.Left = NewWidth - lblTime.Width - 6
    If GetTranslucent Then AlphaBlending picTray.HDC, 0, 0, NewWidth, TaskbarHeight, picDesktopCapture.HDC, TaskbarWidth - NewWidth - IIf(picResize.Visible, picGripper.Width * 2, picGripper.Width), 0, NewWidth, TaskbarHeight, TranslucencyLevel
    picTray.Refresh
End Sub

Sub UpdateTime()
    lblTime.caption = NiceTime(True)
End Sub

Function NiceTime(ampm As Boolean)
If ampm Then
    a = Hour(Now)
    If a > 11 Then a = a - 12: strampm = "pm" Else strampm = "am"
    If a = 0 Then a = 12
    NiceTime = a & ":" & Format(Minute(Now), "00") & " " & strampm
Else
    NiceTime = Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00")
End If
End Function

Private Sub tmrUpdateTime_Timer()
    UpdateTime
    TrayWidth = 100
    lblTime.Left = TrayWidth - lblTime.Width - 6
End Sub

Sub HideTaskbarTips()
    If AppTrayOverIndex <> -1 Then HideToolTip: AppTrayOverIndex = -1
    If ProgramOverIndex <> -1 Then HideToolTip: ProgramOverIndex = -1
    If StartButtonOver = True Then HideToolTip: StartButtonOver = False
    If DesktopButtonOver = True Then HideToolTip: DesktopButtonOver = False
    If ClockOver = True Then HideToolTip: ClockOver = False
    If GripperOver = True Then HideToolTip: GripperOver = False
    If ResizeOver = True Then HideToolTip: ResizeOver = False
End Sub

Sub ResizeDesktop()
    Dim r As RECT
    Call SystemParametersInfo(SPI_GETWORKAREA, 0, r, SPIF_SENDCHANGE)
    r.Bottom = ScreenHeight - TaskbarHeight
    Call SystemParametersInfo(SPI_SETWORKAREA, 0, r, SPIF_SENDCHANGE)
End Sub

Sub UpdateWidth(X, Y, ChangeOldWidth As Boolean)
    TaskbarWidth = TaskbarWidth + X
    If TaskbarWidth < 630 Then TaskbarWidth = 630
    If ChangeOldWidth Then
        SetWindowPos Me.hWND, HWND_TOPMOST, 0, 0, TaskbarWidth, TaskbarHeight, SWP_NOMOVE
        OldWidth = TaskbarWidth
    End If
    oldnumbuttons = TaskbarButtons
    Debug.Print oldnumbuttons
    TaskbarButtons = Int((TaskbarWidth - DesktopButtonWidth - StartButtonWidth - picTray.Width - MyComputerWidth - picQuickstart.Width - (picGripper.Width * 2)) / TaskbarButtonWidth) - 1
    If TaskbarButtons < oldnumbuttons Then
        For i = oldnumbuttons To TaskbarButtons Step -1
            ButtonVisible(i) = False
            picProgram(i).Visible = False
        Next
    End If
    For i = 0 To TaskbarButtons
       ButtonCaption(i) = ""
    Next
    UpdateTaskbar
    Me.Refresh
End Sub

Sub MyComputerOn()
    picMyComputerButton.Cls
    BitBlt picMyComputerButton.HDC, 0, 0, MyComputerWidth, TaskbarHeight, picTaskbarImage(11).HDC, 0, 0, vbSrcCopy
    If GetTranslucent Then AlphaBlending picMyComputerButton.HDC, 0, 0, MyComputerWidth, TaskbarHeight, picDesktopCapture.HDC, StartButtonWidth + DesktopButtonWidth, 0, MyComputerWidth, TaskbarHeight, TranslucencyLevel
End Sub

Sub MyComputerOff()
    picMyComputerButton.Cls
    BitBlt picMyComputerButton.HDC, 0, 0, MyComputerWidth, TaskbarHeight, picTaskbarImage(10).HDC, 0, 0, vbSrcCopy
    If GetTranslucent Then AlphaBlending picMyComputerButton.HDC, 0, 0, MyComputerWidth, TaskbarHeight, picDesktopCapture.HDC, StartButtonWidth + DesktopButtonWidth, 0, MyComputerWidth, TaskbarHeight, TranslucencyLevel
End Sub
