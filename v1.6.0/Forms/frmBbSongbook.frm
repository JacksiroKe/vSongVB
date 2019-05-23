VERSION 5.00
Begin VB.Form frmBbSongbook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Songbook Management | vSongBook"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10260
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
   Icon            =   "frmBbSongbook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin vSongBook.XPButton cmdSaveNow 
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
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
      MICON           =   "frmBbSongbook.frx":146B7
   End
   Begin VB.TextBox txtName 
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
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.ListBox lstSongBook 
      Appearance      =   0  'Flat
      Height          =   5280
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   9855
   End
   Begin VB.TextBox txtLangFile 
      Height          =   2895
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmBbSongbook.frx":146D3
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Label lblSongBookName 
      Caption         =   "SongBook Name:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      DrawMode        =   1  'Blackness
      Height          =   5535
      Left            =   120
      Top             =   720
      Width           =   10095
   End
End
Attribute VB_Name = "frmBbSongbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset


Private Sub Form_Unload(Cancel As Integer)
    frmCcHome.Enabled = True
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    readLangTexts
    Load_AllSongBook
End Sub

Private Sub readLangTexts()
   Open App.Path & "\Langs\" & frmCcHome.AppSettings("preffered_lang") & ".txt" For Input As #1
   txtLangFile.Text = Input$(LOF(1), #1)

   Me.Caption = getLangString(txtLangFile.Text, 12)
   Me.Caption = getLangString(txtLangFile.Text, 92) & " | vSongBook"
   lblSongBookName.Caption = getLangString(txtLangFile.Text, 93)
   cmdSaveNow.Caption = getLangString(txtLangFile.Text, 94)
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

Private Sub cmdSaveNow_Click()
    If txtName.Text = "" Then
        txtName.BackColor = &HFF&
        txtName.SetFocus
        Exit Sub
    Else
        On Error GoTo ErrorHandler
        Set Rs = New ADODB.Recordset
        Rs.Open "Select * from song_book", con, adOpenKeyset, adLockOptimistic
        Rs.AddNew
        Rs!sb_title = txtName.Text
        Rs.Update
        txtName.Text = ""
        Load_AllSongBook
        frmCcHome.ReinitializeSongbook
        Exit Sub
ErrorHandler:
    MsgBox Err.Description & " No. " & Err.Number
    End If
    
End Sub

Private Sub Load_AllSongBook()
    lstSongBook.Clear
    Dim str As String
    On Error GoTo ErrorHandler
     Set Rs = New ADODB.Recordset
        Rs.Open "Select * from song_book", con, adOpenKeyset, adLockOptimistic
        Do Until Rs.EOF
            lstSongBook.AddItem Rs!sb_title
            Rs.MoveNext
        Loop
        Rs.Close
        Exit Sub
ErrorHandler:
    MsgBox Err.Description & " No. " & Err.Number
End Sub


Private Sub txtName_Change()
    If txtName.Text = "" Then
        cmdSaveNow.Enabled = False
    Else
        cmdSaveNow.Enabled = True
    End If
End Sub
