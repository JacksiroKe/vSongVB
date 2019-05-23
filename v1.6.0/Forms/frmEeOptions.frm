VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEeOptions 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vSongBook Settings"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10935
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEeOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEeOptions.frx":146B7
   ScaleHeight     =   7305
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   9840
      Top             =   120
   End
   Begin VB.Frame fraSettings1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   60
      TabIndex        =   22
      Top             =   720
      Width           =   10815
      Begin vSongBook.XPButton cmdSave1 
         Height          =   615
         Left            =   3500
         TabIndex        =   40
         Top             =   5640
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   1085
         TX              =   "Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmEeOptions.frx":1C794
      End
      Begin VB.Frame fraSavedWell1 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         TabIndex        =   33
         Top             =   200
         Visible         =   0   'False
         Width           =   10095
         Begin VB.Shape Shape1 
            Height          =   1095
            Left            =   0
            Top             =   0
            Width           =   10095
         End
         Begin VB.Label lblSavedWell1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Your changes have been saved successfully!"
            Height          =   615
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   9255
         End
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         Height          =   1335
         Left            =   720
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   3270
         Width           =   8415
      End
      Begin VB.ComboBox cmbLanguage 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         ItemData        =   "frmEeOptions.frx":1C7B0
         Left            =   4800
         List            =   "frmEeOptions.frx":1C7C9
         TabIndex        =   24
         Text            =   "English"
         Top             =   1440
         Width           =   4335
      End
      Begin VB.CheckBox chkTabletMode 
         BackColor       =   &H80000016&
         Caption         =   "Tablet Mode"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   600
         TabIndex        =   23
         Top             =   240
         Width           =   9375
      End
      Begin VB.TextBox txtLangFile 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   720
         MultiLine       =   -1  'True
         TabIndex        =   32
         Text            =   "frmEeOptions.frx":1C811
         Top             =   3360
         Width           =   8415
      End
      Begin VB.Label lblRemaining 
         BackColor       =   &H80000016&
         Caption         =   "Characters remaining"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   31
         Top             =   4680
         Width           =   3495
      End
      Begin VB.Label lblRemaineth 
         BackColor       =   &H80000016&
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   30
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label lblYourNameDesc 
         BackColor       =   &H80000016&
         Caption         =   " e.g Brother Jack Siro or Kisii Evening Light Church"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   29
         Top             =   2790
         Width           =   8295
      End
      Begin VB.Label lblYourName 
         BackColor       =   &H80000016&
         Caption         =   "Your Name/ Name of Your Church:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   28
         Top             =   2400
         Width           =   6495
      End
      Begin VB.Label lblPreffered 
         BackColor       =   &H80000016&
         Caption         =   "Preffered Language:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   27
         Top             =   1470
         Width           =   4095
      End
      Begin VB.Label lblTabletMode 
         BackColor       =   &H80000016&
         Caption         =   "Tablet mode is when you are using touch screen input"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   26
         Top             =   750
         Width           =   9375
      End
   End
   Begin VB.Frame fraSettings2 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   60
      TabIndex        =   9
      Top             =   720
      Width           =   10815
      Begin vSongBook.XPButton cmdSave2 
         Height          =   615
         Left            =   3500
         TabIndex        =   41
         Top             =   5640
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   1085
         TX              =   "Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmEeOptions.frx":1C817
      End
      Begin VB.Frame fraSavedWell2 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1095
         Left            =   360
         TabIndex        =   35
         Top             =   200
         Visible         =   0   'False
         Width           =   10095
         Begin VB.Shape Shape2 
            Height          =   1095
            Left            =   0
            Top             =   0
            Width           =   10095
         End
         Begin VB.Label lblSavedWell2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Your changes have been saved successfully!"
            Height          =   615
            Left            =   240
            TabIndex        =   36
            Top             =   240
            Width           =   9495
         End
      End
      Begin VB.ComboBox cmbPreviewFont 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         ItemData        =   "frmEeOptions.frx":1C833
         Left            =   5760
         List            =   "frmEeOptions.frx":1C861
         TabIndex        =   15
         Text            =   "Trebuchet MS"
         Top             =   840
         Width           =   4695
      End
      Begin VB.ComboBox cmbProjectionFont 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         ItemData        =   "frmEeOptions.frx":1C91C
         Left            =   5760
         List            =   "frmEeOptions.frx":1C94A
         TabIndex        =   14
         Text            =   "Tahoma"
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Frame fraSampleText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sample Text:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   360
         TabIndex        =   12
         Top             =   2880
         Width           =   10095
         Begin VB.Label lblSampleText 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "AaBbYyZz"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2055
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   9735
         End
      End
      Begin VB.TextBox txtPreviewFont 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   11
         Text            =   "20"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtProjectionFont 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   10
         Text            =   "50"
         Top             =   1560
         Width           =   735
      End
      Begin MSComctlLib.Slider sldProjectionFont 
         Height          =   495
         Left            =   5760
         TabIndex        =   16
         Top             =   1560
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   10
         SmallChange     =   5
         Min             =   10
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComctlLib.Slider sldPreviewFont 
         DragIcon        =   "frmEeOptions.frx":1CA05
         Height          =   555
         Left            =   5760
         TabIndex        =   17
         Top             =   120
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   979
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   10
         SmallChange     =   5
         Min             =   10
         Max             =   80
         SelStart        =   30
         TickFrequency   =   5
         Value           =   30
         TextPosition    =   1
      End
      Begin VB.Label lblPreviewFS 
         BackColor       =   &H80000016&
         Caption         =   "Song Preview Font Size:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblPreviewFT 
         BackColor       =   &H80000016&
         Caption         =   "Song Preview Font Type:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   4935
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   10440
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   10440
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblProjectionFS 
         BackColor       =   &H80000016&
         Caption         =   "Song Projection Font Size:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   1560
         Width           =   4935
      End
      Begin VB.Label lblProjectionFT 
         BackColor       =   &H80000016&
         Caption         =   "Song Projection Font Type:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   2280
         Width           =   5175
      End
   End
   Begin VB.Frame fraSettings3 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6495
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   10815
      Begin VB.Label themeEight 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "White and Orange Theme"
         ForeColor       =   &H000040C0&
         Height          =   1695
         Left            =   8385
         TabIndex        =   1
         Top             =   4005
         Width           =   1950
      End
      Begin VB.Label themeSeven 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orange and White Theme"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   5775
         TabIndex        =   2
         Top             =   4005
         Width           =   1950
      End
      Begin VB.Label themeSix 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "White and Green Theme"
         ForeColor       =   &H00008000&
         Height          =   1695
         Left            =   3180
         TabIndex        =   3
         Top             =   4005
         Width           =   1950
      End
      Begin VB.Label themeFive 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Green and White Theme"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   585
         TabIndex        =   4
         Top             =   4005
         Width           =   1950
      End
      Begin VB.Label themeFour 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "White and Blue Theme"
         ForeColor       =   &H00C00000&
         Height          =   1695
         Left            =   8385
         TabIndex        =   5
         Top             =   900
         Width           =   1950
      End
      Begin VB.Label themeThree 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Blue and White Theme"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   5775
         TabIndex        =   6
         Top             =   900
         Width           =   1950
      End
      Begin VB.Label themeTwo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "White and Black Theme"
         Height          =   1695
         Left            =   3180
         TabIndex        =   7
         Top             =   900
         Width           =   1950
      End
      Begin VB.Label themeOne 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Black and White Theme"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   585
         TabIndex        =   8
         Top             =   900
         Width           =   1950
      End
      Begin VB.Shape shapeOne 
         BorderColor     =   &H00000000&
         BorderWidth     =   5
         Height          =   2205
         Left            =   360
         Shape           =   1  'Square
         Top             =   600
         Width           =   2400
      End
      Begin VB.Shape shapeTwo 
         BorderWidth     =   5
         Height          =   2205
         Left            =   2955
         Shape           =   1  'Square
         Top             =   600
         Width           =   2400
      End
      Begin VB.Shape shapeThree 
         BorderWidth     =   5
         Height          =   2205
         Left            =   5565
         Shape           =   1  'Square
         Top             =   600
         Width           =   2400
      End
      Begin VB.Shape shapeFive 
         BorderWidth     =   5
         Height          =   2205
         Left            =   360
         Shape           =   1  'Square
         Top             =   3750
         Width           =   2400
      End
      Begin VB.Shape shapeSix 
         BorderWidth     =   5
         Height          =   2205
         Left            =   2955
         Shape           =   1  'Square
         Top             =   3750
         Width           =   2400
      End
      Begin VB.Shape shapeFour 
         BorderWidth     =   5
         Height          =   2205
         Left            =   8160
         Shape           =   1  'Square
         Top             =   600
         Width           =   2400
      End
      Begin VB.Shape shapeSeven 
         BorderWidth     =   5
         Height          =   2205
         Left            =   5565
         Shape           =   1  'Square
         Top             =   3750
         Width           =   2400
      End
      Begin VB.Shape shapeEight 
         BorderWidth     =   5
         Height          =   2205
         Left            =   8160
         Shape           =   1  'Square
         Top             =   3750
         Width           =   2400
      End
      Begin VB.Line Line4 
         X1              =   405
         X2              =   10605
         Y1              =   3180
         Y2              =   3180
      End
   End
   Begin vSongBook.XPButton cmdTab3 
      Height          =   735
      Left            =   6050
      TabIndex        =   39
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1296
      TX              =   "Theme"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmEeOptions.frx":1D8CF
   End
   Begin vSongBook.XPButton cmdTab2 
      Height          =   735
      Left            =   3050
      TabIndex        =   38
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1296
      TX              =   "Display"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmEeOptions.frx":1D8EB
   End
   Begin vSongBook.XPButton cmdTab1 
      Height          =   735
      Left            =   50
      TabIndex        =   37
      Top             =   50
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1296
      TX              =   "General"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmEeOptions.frx":1D907
   End
End
Attribute VB_Name = "frmEeOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim SavedThis As Boolean, lRegion As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMinimize_Click()
    Me.WindowState = vbMinimized
End Sub


Private Sub cmbLanguage_Click()
    readLangTexts
End Sub

Private Sub cmdTab1_Click()
    cmdTabOne
End Sub

Private Sub cmdTab2_Click()
    cmdTabTwo
End Sub

Private Sub cmdTab3_Click()
    cmdTabThree
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    If AppSettings("tablet_mode") = "0" Then
        chkTabletMode.Value = vbUnchecked
    ElseIf AppSettings("tablet_mode") = "1" Then
        chkTabletMode.Value = vbChecked
    End If
    
    cmbLanguage.Text = AppSettings("preffered_lang")
    txtUserName.Text = AppSettings("user_name")
    sldPreviewFont.Value = AppSettings("preview_font_size")
    cmbPreviewFont.Text = MyFontType(AppSettings("preview_font_type"))
    txtPreviewFont.Text = sldPreviewFont.Value
    sldProjectionFont.Value = AppSettings("projection_font_size")
    cmbProjectionFont.Text = MyFontType(AppSettings("projection_font_type"))
    txtProjectionFont.Text = sldProjectionFont.Value
    lblSampleText.FontSize = 35
    lblSampleText.fontname = "Trebuchet MS"
    SavedVsbTheme (AppSettings("preffered_theme"))
    readLangTexts
End Sub


Private Sub readLangTexts()
   Open App.Path & "\Langs\" & cmbLanguage.Text & ".txt" For Input As #1
   txtLangFile.Text = Input$(LOF(1), #1)

   Me.Caption = getLangString(txtLangFile.Text, 42)
   cmdTab1.Caption = getLangString(txtLangFile.Text, 43)
   chkTabletMode.Caption = getLangString(txtLangFile.Text, 44)
   lblTabletMode.Caption = getLangString(txtLangFile.Text, 45)
   lblPreffered.Caption = getLangString(txtLangFile.Text, 7)
   lblYourName.Caption = getLangString(txtLangFile.Text, 13)
   lblYourNameDesc.Caption = getLangString(txtLangFile.Text, 14)
   lblRemaining.Caption = getLangString(txtLangFile.Text, 15)
   cmdSave1.Caption = getLangString(txtLangFile.Text, 4)
   cmdTab2.Caption = getLangString(txtLangFile.Text, 46)
   lblPreviewFS.Caption = getLangString(txtLangFile.Text, 47)
   lblPreviewFT.Caption = getLangString(txtLangFile.Text, 48)
   lblProjectionFS.Caption = getLangString(txtLangFile.Text, 49)
   lblProjectionFT.Caption = getLangString(txtLangFile.Text, 50)
   lblSavedWell1.Caption = getLangString(txtLangFile.Text, 89)
   lblSavedWell2.Caption = getLangString(txtLangFile.Text, 89)
   fraSampleText.Caption = getLangString(txtLangFile.Text, 51)
   lblSampleText.Caption = getLangString(txtLangFile.Text, 52)
   cmdSave2.Caption = getLangString(txtLangFile.Text, 4)
   cmdTab3.Caption = getLangString(txtLangFile.Text, 53)
   themeOne.Caption = getLangString(txtLangFile.Text, 54)
   themeTwo.Caption = getLangString(txtLangFile.Text, 55)
   themeThree.Caption = getLangString(txtLangFile.Text, 56)
   themeFour.Caption = getLangString(txtLangFile.Text, 57)
   themeFive.Caption = getLangString(txtLangFile.Text, 58)
   themeSix.Caption = getLangString(txtLangFile.Text, 59)
   themeSeven.Caption = getLangString(txtLangFile.Text, 60)
   themeEight.Caption = getLangString(txtLangFile.Text, 61)
   
   Close #1
End Sub

Private Function getLangString(ByVal sDataText As String, ByVal nLineNum As Long) As String
    Dim sText As String, nI As Long, nJ As Long, sTemp As String
    On Error GoTo ErrHandler
    sText = ""
    nI = 1
    nJ = 1
    sTemp = ""
    While (nI <= Len(sDataText))
        Select Case Mid(sDataText, nI, 1)
            Case vbCr
                If (nJ = nLineNum) Then
                    sText = sTemp
                End If
            Case vbLf
                nJ = nJ + 1
                sTemp = ""
            Case Else
                sTemp = sTemp & Mid(sDataText, nI, 1)
        End Select
        nI = nI + 1
    Wend
    If (nJ = nLineNum) Then
        sText = sTemp
    End If
    getLangString = sText

    Exit Function

ErrHandler:
    getLangString = ""
End Function

Private Sub minForm_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub tmrTimer_Timer()
    fraSavedWell1.Visible = False
    fraSavedWell2.Visible = False
    tmrTimer.Enabled = False
End Sub

Private Sub txtUserName_Change()
    lblRemaineth.Caption = 50 - Len(txtUserName.Text)
    If txtUserName.Text = "" Or txtUserName.Text = "null" Then
        cmdSave1.Enabled = False
    Else
        cmdSave1.Enabled = True
    End If
End Sub

Private Function SaveSettings(option_title, option_cont) As Boolean
    Set Rs = New ADODB.Recordset
        Rs.Open "Select * from app_options where option_title ='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
        Rs!option_content = option_cont
        Rs.Update
        Rs.Close
        SaveSettings = True
End Function

Private Function AppSettings(option_title) As String
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from app_options WHERE option_title='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
    AppSettings = Rs!option_content
    Rs.Close
End Function

Private Sub cmdSave1_Click()
    fraSavedWell1.Visible = True
    tmrTimer.Enabled = True
    If chkTabletMode.Value = vbChecked Then
       SavedThis = SaveSettings("tablet_mode", "1")
    ElseIf chkTabletMode.Value = vbUnchecked Then
        SavedThis = SaveSettings("tablet_mode", "0")
    End If
    
    SavedThis = SaveSettings("preffered_lang", cmbLanguage.Text)
    SavedThis = SaveSettings("user_name", txtUserName.Text)
End Sub

Private Sub sldPreviewFont_Click()
    lblSampleText.FontSize = sldPreviewFont.Value
    lblSampleText.fontname = cmbPreviewFont.Text
End Sub

Private Sub sldPreviewFont_Scroll()
    txtPreviewFont.Text = sldPreviewFont.Value
    lblSampleText.FontSize = sldPreviewFont.Value
    lblSampleText.fontname = cmbPreviewFont.Text
End Sub

Private Sub sldProjectionFont_Click()
    lblSampleText.FontSize = sldProjectionFont.Value
    lblSampleText.fontname = cmbProjectionFont.Text
End Sub

Private Sub sldProjectionFont_Scroll()
    txtProjectionFont.Text = sldProjectionFont.Value
    lblSampleText.FontSize = sldProjectionFont.Value
    lblSampleText.fontname = cmbProjectionFont.Text
End Sub

Private Sub cmbPreviewFont_Change()
    lblSampleText.FontSize = sldPreviewFont.Value
    lblSampleText.fontname = cmbPreviewFont.Text
End Sub

Private Sub cmbProjectionFont_Change()
    lblSampleText.FontSize = sldProjectionFont.Value
    lblSampleText.fontname = cmbProjectionFont.Text
End Sub

Private Sub cmbPreviewFont_Click()
    lblSampleText.FontSize = sldPreviewFont.Value
    lblSampleText.fontname = cmbPreviewFont.Text
End Sub

Private Sub cmbProjectionFont_Click()
    lblSampleText.FontSize = sldProjectionFont.Value
    lblSampleText.fontname = cmbProjectionFont.Text
End Sub

Private Sub cmdSave2_Click()
    fraSavedWell2.Visible = True
    tmrTimer.Enabled = True
    SavedThis = SaveSettings("preview_font_size", txtPreviewFont.Text)
    SavedThis = SaveSettings("preview_font_type", MyFontName(cmbPreviewFont.Text))
    SavedThis = SaveSettings("projection_font_size", txtProjectionFont.Text)
    SavedThis = SaveSettings("projection_font_type", MyFontName(cmbProjectionFont.Text))
End Sub

Private Sub themeOne_Click()
    SavedThis = SaveSettings("preffered_theme", themeOne_Value)
End Sub

Private Sub themeTwo_Click()
    SavedThis = SaveSettings("preffered_theme", themeTwo_Value)
End Sub

Private Sub themeThree_Click()
    SavedThis = SaveSettings("preffered_theme", themeThree_Value)
End Sub

Private Sub themeFour_Click()
    SavedThis = SaveSettings("preffered_theme", themeFour_Value)
End Sub

Private Sub themeFive_Click()
    SavedThis = SaveSettings("preffered_theme", themeFive_Value)
End Sub

Private Sub themeSix_Click()
    SavedThis = SaveSettings("preffered_theme", themeSix_Value)
End Sub

Private Sub themeSeven_Click()
    SavedThis = SaveSettings("preffered_theme", themeSeven_Value)
End Sub

Private Sub themeEight_Click()
    SavedThis = SaveSettings("preffered_theme", themeEight_Value)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCcHome.Enabled = True
    frmCcHome.readLangTexts
End Sub


