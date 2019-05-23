VERSION 5.00
Begin VB.Form frmAppSplash1 
   BorderStyle     =   0  'None
   Caption         =   "vSongBook"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   Icon            =   "frmAppSplash1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAppSplash1.frx":146B7
   ScaleHeight     =   8895
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   10200
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrTimer 
      Interval        =   3000
      Left            =   240
      Top             =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "v 1.6.0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7680
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblDateRemainder 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   8895
   End
   Begin VB.Label lblUsedBy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Currently Being Used By:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   1560
      TabIndex        =   0
      Top             =   4080
      Width           =   8895
   End
End
Attribute VB_Name = "frmAppSplash1"
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

Private Sub Trans_Parent()
On Error GoTo formErrHandler
    If FileExist(App.Path & "\Res\Splash.bmp") Then
        lRegion = TransparentForm(App.Path & "\Res\Splash.bmp")
        Call SetWindowRgn(Me.hWnd, lRegion, True)
    End If
formErrHandler:
     'Do Nothing
End Sub

Private Sub Form_Load()
     'tmrTimer1.Enabled = False
    Trans_Parent
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

Private Sub showDateRemainder()
    If AppSettings("install_date") = "0" Then
        ddate = DateValue(Now) 'ddate = TimeValue(Now)
        saved = SaveSettings("install_date", ddate)
    Else
       ddiff = DateDiff("d", AppSettings("install_date"), Now)
       If ddiff >= 14 Then
            lblDateRemainder.Caption = "God bless you. I hope you enjoyed Evaluating vSongBook. Install it afresh or contact the developer to learn more."
        Else
            lblDateRemainder.Caption = "God bless you. You are using the Evaluation copy of vSongBook. You have " & 14 - ddiff & " more days to go!"
        
        End If
    End If
    
End Sub

Private Sub saveMyDate()
    If AppSettings("install_date") = "0" Then
        ddate = DateValue(Now)
        saved = SaveSettings("install_date", ddate)
    End If
    
End Sub

Private Sub tmrTimer_Timer()
    tmrTimer.Enabled = False
    If lblUsedBy.Caption = "null" Then
        con.Close
        frmAppStart.Show
        Unload Me
    Else
        frmCcHome.Show
        Unload Me
    End If
End Sub
