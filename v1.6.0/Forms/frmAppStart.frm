VERSION 5.00
Begin VB.Form frmAppStart 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome to vSongBook"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8925
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
   Icon            =   "frmAppStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   595
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   240
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   8415
   End
   Begin VB.ComboBox cmbLanguage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   645
      ItemData        =   "frmAppStart.frx":146B7
      Left            =   6000
      List            =   "frmAppStart.frx":146D0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtLangFile 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmAppStart.frx":14718
      Top             =   2040
      Width           =   8415
   End
   Begin vSongBook.XPButton cmdSave 
      Height          =   735
      Left            =   3000
      TabIndex        =   8
      Top             =   4200
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1296
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmAppStart.frx":1471E
   End
   Begin VB.Label lblRemaining 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remaining characters"
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
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   584
      X2              =   8
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Label lblRemaineth 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3600
      TabIndex        =   5
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label lblYourNameDesc 
      BackColor       =   &H00FFFFFF&
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
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   8295
   End
   Begin VB.Label lblYourName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Name/ Name of Your Church:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   6495
   End
   Begin VB.Label lblPrefferedLang 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Preffered Language:"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmAppStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim SavedThis As Boolean
Dim fLang As Integer, tLang As String, nlang As String, iLang As Long

Private Sub cmbLanguage_Click()
    readLangTexts
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    cmbLanguage.Text = AppSettings("preffered_lang")
    txtUserName.Text = AppSettings("user_name")
    readLangTexts
    Exit Sub
ErrorHandler:
MsgBox "Unable to locate database files. Either you obtain a fresh copy of vSongBook or contact the developer", vbExclamation, "vSongBook unexpected error"
Unload Me
End Sub

Private Sub readLangTexts()
   Open App.Path & "\Langs\" & cmbLanguage.Text & ".txt" For Input As #1
   txtLangFile.Text = Input$(LOF(1), #1)

   Me.Caption = getLangString(txtLangFile.Text, 12)
   lblPrefferedLang.Caption = getLangString(txtLangFile.Text, 7)
   lblYourName.Caption = getLangString(txtLangFile.Text, 13)
   lblYourNameDesc.Caption = getLangString(txtLangFile.Text, 14)
   lblRemaining.Caption = getLangString(txtLangFile.Text, 15)
   cmdSave.Caption = getLangString(txtLangFile.Text, 4)
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

Private Function OpenLang(mystr) As String
    Dim tempStr() As String
    tLang = App.Path & "\Langs\" & AppSettings("preffered_lang") & ".txt"
    fLang = FreeFile
    Open tLang For Input As #fLang
    While Not EOF(fLang)
        Line Input #fLang, nlang
        If InStr(mystr, mystr) = 1 Then
            tempStr() = Split(nlang, "=>")
            OpenLang = Trim(tempStr(1))
        End If
    Wend
    Close #1
End Function

Private Sub SetAppLangStr()
    Me.Caption = OpenLang("form_first")
    lblPrefferedLang.Caption = OpenLang("your_preffered_language")
End Sub

Private Sub cmdSave_Click()
    SavedThis = SaveSettings("user_name", txtUserName.Text)
    SavedThis = SaveSettings("preffered_lang", cmbLanguage.Text)
    con.Close
    frmAppSplash2.Show
    Unload Me
End Sub

Private Sub txtUserName_Change()
    lblRemaineth.Caption = 50 - Len(txtUserName.Text)
    If txtUserName.Text = "" Or txtUserName.Text = "null" Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub

Private Function SaveSettings(option_title, option_cont) As Boolean
On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
        Rs.Open "Select * from app_options where option_title ='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
        Rs!option_content = option_cont
        Rs.Update
        Rs.Close
        SaveSettings = True
        Exit Function
ErrorHandler:
 MsgBox "Unable to save changes. Either you obtain a fresh copy of vSongBook or contact the developer", vbExclamation, "vSongBook unexpected error"
End Function

Private Function AppSettings(option_title) As String
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from app_options WHERE option_title='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
    AppSettings = Rs!option_content
    Rs.Close
End Function

