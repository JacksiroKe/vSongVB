VERSION 5.00
Begin VB.Form frmFfEditSong 
   Caption         =   "Edit a Song"
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
   Icon            =   "frmFfEditSong.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   13605
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   4
      Top             =   1440
      Width           =   10095
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
      TabIndex        =   3
      Top             =   840
      Width           =   10095
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
      ItemData        =   "frmFfEditSong.frx":146B7
      Left            =   120
      List            =   "frmFfEditSong.frx":146D3
      TabIndex        =   2
      Text            =   "Songs Of Worship"
      Top             =   120
      Width           =   4695
   End
   Begin VB.Frame fraCommands 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   5700
      TabIndex        =   1
      Top             =   0
      Width           =   7935
      Begin vSongBook.XPButton cmdSaveClose 
         Height          =   600
         Left            =   120
         TabIndex        =   7
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
         MICON           =   "frmFfEditSong.frx":14773
      End
      Begin vSongBook.XPButton cmdSaveOnly 
         Height          =   600
         Left            =   4000
         TabIndex        =   8
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
         MICON           =   "frmFfEditSong.frx":1478F
      End
   End
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
      ItemData        =   "frmFfEditSong.frx":147AB
      Left            =   120
      List            =   "frmFfEditSong.frx":147AD
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtLangFile 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmFfEditSong.frx":147AF
      Top             =   1560
      Width           =   9495
   End
   Begin VB.Label lblSongID 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmFfEditSong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim SavedThis As Boolean, songid As Integer, lRegion As Long
Dim mysongtitle As String, mysongcont As String, searched As String


Private Sub cmbSongBook_Click()
    lstSongs.Clear
    openSongBook
End Sub

Private Sub cmdSaveClose_Click()
    EditThisSong
    frmCcHome.cmbSongBook_Click
    cmbSongBook_Click
    Unload Me
End Sub

Private Sub cmdSaveOnly_Click()
    EditThisSong
    frmCcHome.cmbSongBook_Click
    cmbSongBook_Click
End Sub

Private Sub Form_Load()
    cmbSongBook.Text = frmCcHome.cmbSongBook.Text
    songid = frmCcHome.lblSongid.Caption
    
    txtSongCont.FontSize = frmCcHome.AppSettings("preview_font_size")
    txtSongCont.fontname = MyFontType(frmCcHome.AppSettings("preview_font_type"))

    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    openSongBook
    ChooseThisSong
    readLangTexts
End Sub

Private Sub readLangTexts()
   Open App.Path & "\Langs\" & AppSettings("preffered_lang") & ".txt" For Input As #1
   txtLangFile.Text = Input$(LOF(1), #1)

   Me.Caption = getLangString(txtLangFile.Text, 74) & cmbSongBook.Text & " | vSongBook"
   cmdSaveClose.Caption = getLangString(txtLangFile.Text, 75)
   cmdSaveOnly.Caption = getLangString(txtLangFile.Text, 76)
    
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
Private Sub ChooseThisSong()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_content WHERE songid=" & songid, con, adOpenKeyset, adLockOptimistic
    songid = Rs!songid
    mysongtitle = Replace(Rs!song_title, "+", Chr$(34))
    mysongtitle = Replace(mysongtitle, "^", "'")
    txtSongTitle.Text = Replace(mysongtitle, "$", vbNewLine)
    
    mysongcont = Replace(Rs!song_content, "+", Chr$(34))
    mysongcont = Replace(mysongcont, "^", "'")
    txtSongCont.Text = Replace(mysongcont, "$", vbNewLine)
    Me.Caption = "Editting Song: " & mysongtitle & " | vSongBook"
    
End Sub

Private Sub Form_Resize()
    If (frmFfNewsong.Height > 7000 And frmFfNewsong.Width > 7000) Then
        fraCommands.left = frmFfEditSong.Width - 8500
        txtSongTitle.Width = frmFfEditSong.Width - 3700
        txtSongCont.Height = frmFfEditSong.Height - 2500
        txtSongCont.Width = frmFfEditSong.Width - 3700
        lstSongs.Height = frmFfEditSong.Height - 1600
    Else
        Exit Sub
    End If
    
End Sub

Private Sub openSongBook()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_content WHERE song_book = '" & cmbSongBook.Text & "'", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstSongs.AddItem Rs!song_title
        Rs.MoveNext
    Loop
End Sub

Private Sub txtSongTitle_Change()
    If txtSongTitle.Text = "" Or txtSongTitle.Text = "Song Title" Then
        cmdSaveOnly.Enabled = False
        cmdSaveClose.Enabled = False
    Else
        cmdSaveOnly.Enabled = True
        cmdSaveClose.Enabled = True
    End If
End Sub

Private Sub EditThisSong()
    Dim songno As Integer, songtitle As String, songcont As String
    songtitle = Convert_Text_Rvs(txtSongTitle.Text)
    songcont = Convert_Text_Rvs(txtSongCont.Text)
    
    On Error GoTo ErrorHandler
        Set Rs = New ADODB.Recordset
        Rs.Open "Select * from song_content WHERE songid=" & songid, con, adOpenKeyset, adLockOptimistic
    Rs!song_no = songno
        Rs!song_book = cmbSongBook.Text
        Rs!song_title = songtitle
        Rs!song_content = songcont
        Rs.Update
        Rs.Close
        Exit Sub
ErrorHandler:
     MsgBox "Unable to edit your song. Please try again", vbExclamation, "vSongBook"
End Sub

Private Sub lstSongs_Click()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_content WHERE song_title = '" & lstSongs.Text & "'", con, adOpenKeyset, adLockOptimistic
    songid = Rs!songid
    txtSongTitle.Text = Convert_Text_Max(Rs!song_title)
    txtSongCont.Text = Convert_Text_Max(Rs!song_content)
    Me.Caption = "Editting Song: " & mysongtitle & " | vSongBook"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCcHome.Enabled = True
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

