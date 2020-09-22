VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5715
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTitleBar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7335
      TabIndex        =   35
      Top             =   0
      Width           =   7335
      Begin VB.Label lblTitlebar 
         BackColor       =   &H00000000&
         Caption         =   "B2 Settings"
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
         TabIndex        =   36
         Top             =   60
         Width           =   7095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6060
      TabIndex        =   13
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4860
      TabIndex        =   12
      Top             =   5280
      Width           =   1095
   End
   Begin MSComctlLib.TreeView treSettings 
      Height          =   4695
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8281
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlTreeIcons"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlTreeIcons 
      Left            =   2100
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":13F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picContent 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   2
      Left            =   2760
      ScaleHeight     =   4635
      ScaleWidth      =   4455
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   3615
         TabIndex        =   34
         Top             =   3600
         Width           =   3615
      End
      Begin VB.PictureBox picTaskbarTrans 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   3615
         TabIndex        =   33
         Top             =   3000
         Width           =   3615
      End
      Begin MSComctlLib.Slider sldTranslucency 
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   900
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider sldMenuTranslucency 
         Height          =   255
         Left            =   960
         TabIndex        =   29
         Top             =   2100
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickFrequency   =   16
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Translucent"
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
         Index           =   16
         Left            =   1620
         TabIndex        =   32
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Transparent"
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
         Index           =   15
         Left            =   3240
         TabIndex        =   31
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Opaque"
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
         Index           =   14
         Left            =   240
         TabIndex        =   30
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Menu translucency level:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   180
         TabIndex        =   28
         Top             =   1740
         Width           =   2835
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Translucent"
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
         Index           =   12
         Left            =   1620
         TabIndex        =   27
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Transparent"
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
         Index           =   11
         Left            =   3240
         TabIndex        =   26
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Opaque"
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
         Index           =   10
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Takbar translucency level:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   24
         Top             =   600
         Width           =   2835
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000014&
         Index           =   5
         X1              =   60
         X2              =   4380
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   4
         X1              =   60
         X2              =   4380
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Translucency Settings"
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
         Index           =   2
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   2175
      End
   End
   Begin VB.PictureBox picContent 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   0
      Left            =   2760
      ScaleHeight     =   4635
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label lblExplaination 
         Alignment       =   2  'Center
         Caption         =   "Please select a category to configure the appearence of B2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   2040
         Width           =   4275
      End
   End
   Begin VB.PictureBox picContent 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   3
      Left            =   2760
      ScaleHeight     =   4635
      ScaleWidth      =   4455
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label lblExplaination 
         Alignment       =   2  'Center
         Caption         =   "Please select a category to configure the behavior of B2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   60
         TabIndex        =   8
         Top             =   1980
         Width           =   4275
      End
   End
   Begin VB.PictureBox picContent 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   6
      Left            =   2760
      ScaleHeight     =   4635
      ScaleWidth      =   4455
      TabIndex        =   45
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   3615
         TabIndex        =   46
         Top             =   780
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "This feature will be included in B2 1.1"
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
         Left            =   480
         TabIndex        =   48
         Top             =   1980
         Width           =   3555
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000014&
         Index           =   11
         X1              =   60
         X2              =   4380
         Y1              =   310
         Y2              =   310
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   10
         X1              =   60
         X2              =   4380
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lblExplaination 
         Caption         =   "AutoUpdate"
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
         Index           =   21
         Left            =   60
         TabIndex        =   47
         Top             =   60
         Width           =   4275
      End
   End
   Begin VB.PictureBox picContent 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   4
      Left            =   2760
      ScaleHeight     =   4635
      ScaleWidth      =   4455
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   3615
         TabIndex        =   40
         Top             =   780
         Width           =   3615
         Begin VB.OptionButton optAnimateMenus 
            Caption         =   "Skin"
            Height          =   255
            Index           =   2
            Left            =   1380
            TabIndex        =   41
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optAnimateMenus 
            Caption         =   "Off"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   42
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAnimateMenus 
            Caption         =   "On"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   43
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Animate menus:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   60
         TabIndex        =   44
         Top             =   480
         Width           =   2835
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Menu Settings"
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
         Index           =   4
         Left            =   60
         TabIndex        =   10
         Top             =   60
         Width           =   4275
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   60
         X2              =   4380
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   60
         X2              =   4380
         Y1              =   310
         Y2              =   310
      End
   End
   Begin VB.PictureBox picContent 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   5
      Left            =   2760
      ScaleHeight     =   4635
      ScaleWidth      =   4455
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label lblExplaination 
         Alignment       =   2  'Center
         Caption         =   "No B2Apps were included in this release of B2"
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
         Index           =   20
         Left            =   120
         TabIndex        =   49
         Top             =   2400
         Width           =   4275
      End
      Begin VB.Label lblExplaination 
         Caption         =   "B2Apps"
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
         Index           =   22
         Left            =   60
         TabIndex        =   39
         Top             =   60
         Width           =   4275
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   9
         X1              =   60
         X2              =   4380
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000014&
         Index           =   8
         X1              =   60
         X2              =   4380
         Y1              =   315
         Y2              =   315
      End
   End
   Begin VB.PictureBox picContent 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   1
      Left            =   2760
      ScaleHeight     =   4635
      ScaleWidth      =   4455
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4020
         TabIndex        =   37
         Top             =   420
         Width           =   315
      End
      Begin VB.PictureBox picSkinPreviewContainer 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   4215
         TabIndex        =   20
         Top             =   900
         Width           =   4215
         Begin VB.PictureBox picSkinPreview 
            AutoSize        =   -1  'True
            Height          =   1875
            Left            =   420
            ScaleHeight     =   1815
            ScaleWidth      =   3135
            TabIndex        =   21
            Top             =   120
            Width           =   3195
         End
      End
      Begin VB.ComboBox cmbSkinNames 
         Height          =   315
         ItemData        =   "frmSettings.frx":174C
         Left            =   120
         List            =   "frmSettings.frx":174E
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   420
         Width           =   3855
      End
      Begin VB.Label lblSkinDescription 
         Caption         =   "<Skin Description>"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   1200
         TabIndex        =   19
         Top             =   3660
         Width           =   3075
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Description:"
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
         Index           =   7
         Left            =   90
         TabIndex        =   18
         Top             =   3660
         Width           =   1095
      End
      Begin VB.Label lblSkinAuthor 
         Caption         =   "<Skin Author>"
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
         Left            =   1200
         TabIndex        =   17
         Top             =   3420
         Width           =   3075
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Author:"
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
         Index           =   6
         Left            =   480
         TabIndex        =   16
         Top             =   3420
         Width           =   735
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Skin Name:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label lblSkinName 
         Caption         =   "<Skin Name>"
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
         Left            =   1200
         TabIndex        =   14
         Top             =   3180
         Width           =   3135
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   60
         X2              =   4380
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   60
         X2              =   4380
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lblExplaination 
         Caption         =   "Skin Settings"
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
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   4275
      End
   End
   Begin VB.Label lblExplaination 
      Alignment       =   2  'Center
      Caption         =   "Please select a category to configure B2."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   2880
      TabIndex        =   22
      Top             =   2400
      Width           =   4275
   End
   Begin VB.Line linSeperator 
      BorderColor     =   &H80000010&
      Index           =   7
      X1              =   60
      X2              =   7200
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line linSeperator 
      BorderColor     =   &H80000014&
      Index           =   6
      X1              =   60
      X2              =   7200
      Y1              =   5175
      Y2              =   5175
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SkinNames As Variant
Dim SkinINIPath As String

Dim TX As Integer, TY As Integer, movenow As Boolean

Private Sub cmbSkinNames_Click()
    SkinINIPath = App.path & "\skins\" & cmbSkinNames.List(cmbSkinNames.ListIndex) & "\skin.ini"
    SkinPicturePath = App.path & "\skins\" & cmbSkinNames.List(cmbSkinNames.ListIndex) & "\preview.bmp"
    lblSkinName.caption = ReadINI("skin", "longname", SkinINIPath)
    lblSkinAuthor.caption = ReadINI("skin", "author", SkinINIPath)
    lblSkinDescription.caption = ReadINI("skin", "description", SkinINIPath)
    picSkinPreview.Picture = LoadPicture(SkinPicturePath)
    picSkinPreview.Left = (picSkinPreviewContainer.Width - picSkinPreview.Width) / 2
    picSkinPreview.Top = (picSkinPreviewContainer.Height - picSkinPreview.Height) / 2
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCustomColors_Click()
    RandomColors
End Sub

Private Sub cmdOK_Click()
    If cmbSkinNames.List(cmbSkinNames.ListIndex) <> GetSkinName Then
        ChangeSkin (cmbSkinNames.List(cmbSkinNames.ListIndex))
    End If
    If sldTranslucency.Value > 140 Or sldMenuTranslucency.Value > 140 Then
        If ChoiceBoxEX("Warning! One of your translucency settings has been set abnormally high. You may not be able to see your taskbar or menus. Do you want to continue?", "Translucency Level High") = False Then
            Exit Sub
        End If
    End If
    SetTranslucencyLevel sldTranslucency.Value
    SetMenuTranslucencyLevel sldMenuTranslucency.Value
    Unload Me
End Sub



Private Sub Form_Load()
    SetupTree
    
    SkinNames = ListFolderItems(App.path & "\skins\")
    For i = 0 To UBound(SkinNames)
        cmbSkinNames.AddItem SkinNames(i)
        If SkinNames(i) = GetSkinName Then cmbSkinNames.ListIndex = i
    Next
    sldTranslucency.Value = TranslucencyLevel
    sldMenuTranslucency.Value = MenuTranslucencyLevel
End Sub

Private Sub SetupTree()
    treSettings.Nodes.Add , tvwFirst, "Appearence", "Appearence", 1, 1
        treSettings.Nodes.Add 1, tvwChild, "Skin", "Skin", 2, 2
        treSettings.Nodes.Add 1, tvwChild, "Translucency", "Translucency", 3, 3
    treSettings.Nodes.Add , tvwFirst, "Behavior", "Behavior", 4, 4
        treSettings.Nodes.Add 4, tvwChild, "Menus", "Menus", 5, 5
    treSettings.Nodes.Add , tvwFirst, "B2Apps", "B2Apps", 6, 6
    treSettings.Nodes.Add , tvwFirst, "AutoUpdate", "AutoUpdate", 7, 7
End Sub


Private Sub lblTitlebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    movenow = True
    TX = X
    TY = Y
End Sub

Private Sub lblTitlebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If movenow Then
        Me.Top = Me.Top + Y - TY
        Me.Left = Me.Left + X - TX
    End If
End Sub

Private Sub lblTitlebar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    movenow = False
End Sub

Private Sub treSettings_NodeClick(ByVal Node As MSComctlLib.Node)
    picContent(Node.Index - 1).Visible = True
    For i = 0 To picContent.Count - 1
        If i <> Node.Index - 1 Then picContent(i).Visible = False
    Next
End Sub

Function ListFolderItems(ByVal path As String) As Variant
    'returns an array of directory names
    On Error Resume Next
    Dim Count, Items(), i, ItemName ' Declare variables.
    ItemName = Dir(path, vbDirectory Or vbArchive Or vbSystem Or vbReadOnly) ' Get first directory name.
    Count = 0

    Do While Not ItemName = ""
        'A file or directory name was returned
        If Not ItemName = "." And Not ItemName = ".." Then
            ReDim Preserve Items(Count + 1)
            Items(Count) = ItemName ' Add directory name to array
            Count = Count + 1
        End If
        ItemName = Dir ' Get another item name
    Loop
    ReDim Preserve Items(Count - 1)
    If Count = 0 Then
        ListFolderItems = -1
    Else
        ListFolderItems = Items
    End If
End Function
