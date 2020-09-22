VERSION 5.00
Begin VB.Form frmRun 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1515
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTitlebar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   0
      Width           =   5595
      Begin VB.CommandButton cmdClose 
         Height          =   120
         Left            =   5160
         TabIndex        =   4
         Top             =   120
         Width           =   135
      End
      Begin VB.Label lblTitlebar 
         BackColor       =   &H00000000&
         Caption         =   "Run..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   4755
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Default         =   -1  'True
      Height          =   375
      Left            =   4260
      TabIndex        =   1
      Top             =   420
      Width           =   1035
   End
   Begin VB.TextBox txtFilename 
      Appearance      =   0  'Flat
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
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   3795
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmRun.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type a URL or the name of a program, folder, or document and B2 will open it."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   900
      TabIndex        =   5
      Top             =   420
      Width           =   3135
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()
    SaveSetting "B2", "Run", "LastPath", txtFilename.text
    If ShellFile(txtFilename.text) = 2 Then
        MsgBoxEX "The file that you entered does not exist. If you are entering a website address, be sure to include the http:// if the web address does not include ""www"".", "File not found - B2 Error 1", sndError
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    txtFilename.text = GetSetting("B2", "Run", "LastPath", "")
    txtFilename.SelStart = 0
    txtFilename.SelLength = Len(txtFilename.text)
End Sub
