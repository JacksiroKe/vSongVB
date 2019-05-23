VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppSplash2 
   BorderStyle     =   0  'None
   Caption         =   "Welcome to vSongBook"
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10545
   Icon            =   "frmAppSplash2.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmAppSplash2.frx":146B7
   ScaleHeight     =   5730
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
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
      Left            =   9360
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrTimer 
      Interval        =   5
      Left            =   360
      Top             =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "v 1.6.0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblUsedBy 
      BackStyle       =   0  'Transparent
      Caption         =   "Currently Being Used By:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1800
      TabIndex        =   0
      Top             =   3600
      Width           =   7335
   End
End
Attribute VB_Name = "frmAppSplash2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim saved As Boolean, ddate As Date, ddiff As Integer, lRegion As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
 On Error GoTo ErrorHandler
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    lblUsedBy.Caption = AppSettings("user_name")
    saveMyDate
    Exit Sub
ErrorHandler:
MsgBox "Unable to locate database files. Either you obtain a fresh copy of vSongBook or contact the developer", vbExclamation, "vSongBook unexpected error"
Unload Me
End Sub

Private Function AppSettings(option_title) As String
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from app_options WHERE option_title='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
    AppSettings = Rs!option_content
End Function

Private Function SaveSettings(option_title, option_cont) As Boolean
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from app_options where option_title ='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
    Rs!option_content = option_cont
    Rs.Update
    Rs.Close
    SaveSettings = True
End Function

Private Sub saveMyDate()
    If AppSettings("install_date") = "0" Then
        ddate = DateValue(Now)
        saved = SaveSettings("install_date", ddate)
    End If
    
End Sub

Private Sub tmrTimer_Timerx()
    ProgressBar1.Value = ProgressBar1.Value + 1
    If lblUsedBy.Caption = "null" Then
        con.Close
        'frmCcHome.Show
        frmAppStart.Show
        Unload Me
    Else
        frmCcHome.Show
        Unload Me
    End If
End Sub

Private Sub tmrTimer_Timer()
    On Error GoTo ErrorHandler:
        ProgressBar1.Value = ProgressBar1.Value + 1
Exit Sub
ErrorHandler:
    If Err.Number = 380 Then
        If lblUsedBy.Caption = "null" Then
            con.Close
            'frmCcHome.Show
            frmAppStart.Show
            Unload Me
        Else
            frmCcHome.Show
            Unload Me
        End If
    End If
End Sub
