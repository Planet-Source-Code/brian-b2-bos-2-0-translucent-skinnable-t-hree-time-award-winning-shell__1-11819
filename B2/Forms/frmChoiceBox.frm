VERSION 5.00
Begin VB.Form frmChoiceBox 
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
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
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
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   180
      Picture         =   "frmChoiceBox.frx":0000
      Top             =   660
      Width           =   480
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[Box Label]"
      Height          =   975
      Left            =   780
      TabIndex        =   0
      Top             =   420
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmChoiceBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DragNow As Boolean, DragX As Integer, DragY As Integer
Public Response As Boolean

Public Sub DisplayInfo(Text As String, Optional caption As String = "Information")
    lblInfo.caption = Text
    lblTitle.caption = caption
    Me.Show
    PlaySkinSound "Question"
End Sub

Private Sub cmdYes_Click()
    Response = True
    Me.Hide
    HideToolTip
End Sub

Private Sub cmdYes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ToolTipDisplayed = False Then
        ToolTipEX "Click here to choose yes and close the question box.", (Me.Top / Screen.TwipsPerPixelY) + cmdYes.Top - 23, (Me.Left / Screen.TwipsPerPixelX) + cmdYes.Left + 2 + cmdYes.Width
    End If
End Sub

Private Sub cmdNo_Click()
    Response = False
    Me.Hide
    HideToolTip
End Sub

Private Sub cmdNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ToolTipDisplayed = False Then
        ToolTipEX "Click here to choose no and close the question box.", (Me.Top / Screen.TwipsPerPixelY) + cmdNo.Top - 23, (Me.Left / Screen.TwipsPerPixelX) + cmdNo.Left + 2 + cmdNo.Width
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideToolTip
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

