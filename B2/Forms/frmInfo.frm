VERSION 5.00
Begin VB.Form frmInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1530
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5460
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTitleBar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   5535
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      Begin VB.Label lblTitle 
         BackColor       =   &H00000000&
         Caption         =   "[Box Title]"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   5355
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4380
      TabIndex        =   1
      Top             =   540
      Width           =   915
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   180
      Picture         =   "frmInfo.frx":0000
      Top             =   660
      Width           =   480
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[Box Label]"
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Top             =   420
      Width           =   3375
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuTaskbarIconMenu 
      Caption         =   "mnuTaskbarIconMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuTemp 
         Caption         =   "Temporary Menu"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuSendTo 
         Caption         =   "Copy To"
         Begin VB.Menu mnuSendToDesktop 
            Caption         =   "Desktop"
         End
         Begin VB.Menu mnuSendToStartMenu 
            Caption         =   "Start Menu"
         End
      End
   End
   Begin VB.Menu mnuTaskMenu 
      Caption         =   "mnuTaskMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuTemporaryMenu 
         Caption         =   "Temporary Menu"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSeperator2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuMenuItemMenu 
      Caption         =   "mnuMenuItemMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuTemp2 
         Caption         =   "Temporary Menu"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenItem 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DragNow As Boolean, DragX As Integer, DragY As Integer
Dim OKOver As Boolean

Public Enum SoundTypes
    sndNone
    sndError
    sndSuccess
    sndInformation
End Enum

Public Sub DisplayInfo(text As String, Optional caption As String = "Information", Optional sound As SoundTypes = sndNone)
    lblInfo.caption = text
    lblTitle.caption = caption
    Me.Show
    If sound = sndError Then PlaySkinSound "Error"
    If sound = sndSuccess Then PlaySkinSound "Success"
End Sub

Private Sub cmdOK_Click()
    Me.Hide
    HideToolTip
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If OKOver = False Then
        OKOver = True
        ToolTipEX "Click here to close the information box.", (Me.Top / Screen.TwipsPerPixelY) + cmdOK.Top - 23, (Me.Left / Screen.TwipsPerPixelX) + cmdOK.Left + 2
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideToolTip
    OKOver = False
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNow = True
    DragX = X
    DragY = Y
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragNow Then
        Me.Top = Me.Top + Y - DragY
        Me.Left = Me.Left + X - DragX
    End If
End Sub

Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNow = False
End Sub

Private Sub mnuDelete_Click()
    If ChoiceBoxEX("Do you really want to permenently delete " & frmTaskbar.imgIcon(CurrentSelectedTaskbarItem).Tag & "? The item will be not be kept in the recycle bin.", "Delete File") Then
        Kill App.path & "\quickstart\" & frmTaskbar.imgIcon(CurrentSelectedTaskbarItem).Tag
    End If
End Sub

Private Sub mnuRename_Click()
    newName = InputBox("<TEMP DIALOG>" & vbCrLf & "Please enter a new name for " & frmTaskbar.imgIcon(CurrentSelectedTaskbarItem).Tag & ":")
    If newName <> "" Then
        FileCopy App.path & "\quickstart\" & frmTaskbar.imgIcon(CurrentSelectedTaskbarItem).Tag, App.path & "\quickstart\" & newName
        Kill App.path & "\quickstart\" & frmTaskbar.imgIcon(CurrentSelectedTaskbarItem).Tag
    End If
End Sub

Private Sub mnuSendToDesktop_Click()
    FileCopy App.path & "\quickstart\" & frmTaskbar.imgIcon(CurrentSelectedTaskbarItem).Tag, GetDesktopPath & "\" & frmTaskbar.imgIcon(CurrentSelectedTaskbarItem).Tag
    MsgBoxEX "File copied", "Success", sndSuccess
End Sub

Private Sub mnuSendToStartMenu_Click()
    FileCopy App.path & "\quickstart\" & frmTaskbar.imgIcon(CurrentSelectedTaskbarItem).Tag, GetDesktopPath & "\" & frmTaskbar.imgIcon(CurrentSelectedTaskbarItem).Tag
End Sub
