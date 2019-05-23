VERSION 5.00
Begin VB.Form frmFfNewsong 
   Caption         =   "NewSong"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFfNewsong.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   13605
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstSongs 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3135
   End
   Begin VB.Frame fraCommands 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   5500
      TabIndex        =   2
      Top             =   0
      Width           =   7935
      Begin vSongBook.XPButton cmdSaveClose 
         Height          =   600
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   1058
         TX              =   "Save and &Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmFfNewsong.frx":146B7
      End
      Begin vSongBook.XPButton cmdSaveAdd 
         Height          =   600
         Left            =   4000
         TabIndex        =   7
         Top             =   120
         Width           =   3800
         _ExtentX        =   6694
         _ExtentY        =   1058
         TX              =   "Save &ONLY"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmFfNewsong.frx":146D3
      End
   End
   Begin VB.ComboBox cmbSongBook 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmFfNewsong.frx":146EF
      Left            =   120
      List            =   "frmFfNewsong.frx":1470B
      TabIndex        =   3
      Text            =   "Songs Of Worship"
      Top             =   120
      Width           =   4695
   End
   Begin VB.TextBox txtSongTitle 
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
      Left            =   3360
      TabIndex        =   0
      Top             =   840
      Width           =   10095
   End
   Begin VB.TextBox txtSongCont 
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
      Height          =   6525
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1440
      Width           =   10095
   End
   Begin VB.TextBox txtLangFile 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmFfNewsong.frx":147AB
      Top             =   2040
      Width           =   6495
   End
End
Attribute VB_Name = "frmFfNewsong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim SavedThis As Boolean
Dim mysong As String, searched As String, lRegion As Long


Private Sub cmbSongBook_Click()
    Me.Caption = "Add a New Song to " & cmbSongBook.Text & " | vSongBook"
    lstSongs.Clear
    openSongBook
End Sub

Private Sub cmdSaveClose_Click()
    SaveNewSong
    frmCcHome.cmbSongBook_Click
    cmbSongBook_Click
    Unload Me
End Sub

Private Sub cmdSaveAdd_Click()
    SaveNewSong
    txtSongTitle.Text = ""
    txtSongCont.Text = ""
    txtSongTitle.SetFocus
    frmCcHome.cmbSongBook_Click
    cmbSongBook_Click
End Sub

Private Sub Form_Load()
    txtSongCont.FontSize = frmCcHome.AppSettings("preview_font_size")
    txtSongCont.fontname = MyFontType(frmCcHome.AppSettings("preview_font_type"))
    
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    readLangTexts
    
    openSongBook
    
End Sub

Private Sub readLangTexts()
   Open App.Path & "\Langs\" & AppSettings("preffered_lang") & ".txt" For Input As #1
   txtLangFile.Text = Input$(LOF(1), #1)

   Me.Caption = getLangString(txtLangFile.Text, 79) & cmbSongBook.Text & " | vSongBook"
   cmdSaveAdd.Caption = getLangString(txtLangFile.Text, 80)
   cmdSaveClose.Caption = getLangString(txtLangFile.Text, 75)
   txtSongTitle.Text = getLangString(txtLangFile.Text, 81)
   txtSongCont.Text = getLangString(txtLangFile.Text, 82)
    
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

Private Sub Form_Resize()
    If (frmFfNewsong.Height > 7000 And frmFfNewsong.Width > 7000) Then
        fraCommands.left = frmFfNewsong.Width - 8500
        txtSongTitle.Width = frmFfNewsong.Width - 3700
        txtSongCont.Height = frmFfNewsong.Height - 2500
        txtSongCont.Width = frmFfNewsong.Width - 3700
        lstSongs.Height = frmFfNewsong.Height - 1600
    Else
        Exit Sub
    End If
    
End Sub

Private Sub openSongBook()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_content WHERE song_book = '" & cmbSongBook.Text & "'", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstSongs.AddItem Convert_Text_Min(Rs!song_title)
        Rs.MoveNext
    Loop
    Me.Caption = "Add a New Song " & Chr$(34) & "Song " & lstSongs.ListCount + 1 & "#" & Chr$(34) & " to " & cmbSongBook.Text & " | vSongBook"
End Sub

Private Sub txtSongCont_GotFocus()
    Static bSet As Boolean
    If Not bSet Then
        txtSongCont.Text = ""
        bSet = True
    End If
End Sub

Private Sub txtSongTitle_Change()
    If txtSongTitle.Text = "" Or txtSongTitle.Text = "Song Title" Then
        cmdSaveAdd.Enabled = False
        cmdSaveClose.Enabled = False
    Else
        cmdSaveAdd.Enabled = True
        cmdSaveClose.Enabled = True
    End If
End Sub

Private Sub txtSongTitle_GotFocus()
    Static bSet As Boolean
    If Not bSet Then
        txtSongTitle.Text = ""
        bSet = True
    End If
    
End Sub

Private Sub SaveNewSong()
    Dim songno As Integer, songtitle As String, songcont As String
    songno = lstSongs.ListCount + 1
    songtitle = songno & "# " & Convert_Text_Rvs(txtSongTitle.Text)
    songcont = Convert_Text_Rvs(txtSongCont.Text)
    
    On Error GoTo ErrorHandler
        
        Set Rs = New ADODB.Recordset
        Rs.Open "Select * from song_content WHERE song_book = '" & cmbSongBook.Text & "'", con, adOpenKeyset, adLockOptimistic
        Rs!song_no = songno
        Rs!song_book = cmbSongBook.Text
        Rs!song_title = songtitle
        Rs!song_content = songcont & " $ $ " & songtitle
        Rs.AddNew
        Rs.Close
        Exit Sub
ErrorHandler:
     MsgBox "Unable to add your song. Please try again", vbExclamation, "vSongBook"
End Sub

Public Function AppSettings(option_title) As String
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from app_options WHERE option_title='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
    AppSettings = Rs!option_content
    Rs.Close
End Function

Private Sub Form_Unload(Cancel As Integer)
    frmCcHome.Enabled = True
End Sub

