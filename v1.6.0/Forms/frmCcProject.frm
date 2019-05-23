VERSION 5.00
Begin VB.Form frmCcProject 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Projection Mode"
   ClientHeight    =   8820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   Icon            =   "frmCcProject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   380
      Left            =   7080
      Picture         =   "frmCcProject.frx":146B7
      ScaleHeight     =   375
      ScaleWidth      =   945
      TabIndex        =   11
      Top             =   720
      Width           =   950
   End
   Begin VB.PictureBox cmdMinus 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4320
      Picture         =   "frmCcProject.frx":16627
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   10
      Top             =   7800
      Width           =   650
   End
   Begin VB.PictureBox cmdPlay 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   600
      Picture         =   "frmCcProject.frx":1692C
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   9
      Top             =   480
      Width           =   650
   End
   Begin VB.PictureBox cmdFont 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5760
      Picture         =   "frmCcProject.frx":16C95
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   8
      Top             =   7800
      Width           =   650
   End
   Begin VB.PictureBox cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5040
      Picture         =   "frmCcProject.frx":17078
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   7
      Top             =   7800
      Width           =   650
   End
   Begin VB.PictureBox cmdKala 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6480
      Picture         =   "frmCcProject.frx":173CF
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   6
      Top             =   7800
      Width           =   650
   End
   Begin VB.PictureBox cmdPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   7200
      Picture         =   "frmCcProject.frx":1779E
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   5
      Top             =   7800
      Width           =   650
   End
   Begin VB.PictureBox cmdNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   7920
      Picture         =   "frmCcProject.frx":17B9B
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   4
      Top             =   7800
      Width           =   650
   End
   Begin VB.Timer tmrTimerErr 
      Left            =   0
      Top             =   2160
   End
   Begin VB.Timer tmrTimerTime 
      Left            =   0
      Top             =   1560
   End
   Begin VB.Timer tmrTimerSong 
      Left            =   0
      Top             =   960
   End
   Begin VB.TextBox cmdLine 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblSongText 
      BackStyle       =   0  'Transparent
      Caption         =   "SongContent"
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
      Height          =   6135
      Left            =   1140
      TabIndex        =   3
      Top             =   1343
      Width           =   7815
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "Kisii Church"
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
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   7800
      Width           =   4335
   End
   Begin VB.Label lblSongTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SongTitle"
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
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.Line lineDown 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   1080
      X2              =   8760
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line lineTop 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   840
      X2              =   8520
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmCcProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Dim SavedThis As Boolean, content() As String
Dim cur_stanza As Integer, all_stanza As Integer, song_colour As Integer, song_fonttype As Integer
Dim title As String, mytitle As String, mysong As String, songid As String

Dim songScroll As Boolean, songTime As Integer, song_fontsize As Integer
Dim oldTime As Date, newTime As Date, diff As Date

Private Sub cmdClose_Click()
    'Unload frmCcHome
    Unload Me
End Sub

Private Sub cmdClose_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    Projection_Form_Resize
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    songid = frmCcHome.lblSongid.Caption
    songScroll = False
    ShowOrHide (AppSettings("tablet_mode"))
    
    song_fontsize = CInt(AppSettings("projection_font_size"))
    song_fonttype = CInt(AppSettings("projection_font_type"))
    lblUserName.Caption = AppSettings("user_name")
    lblSongText.FontSize = song_fontsize
    AppSettings ("projection_font_size")
    song_colour = AppSettings("preffered_theme")
    
    lblSongText.fontname = MyFontType(song_fonttype)
    SetProjectionTheme (AppSettings("preffered_theme"))
    SongForProjection
    
End Sub

Private Sub SongForProjection()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_content WHERE songid= " & songid & "", con, adOpenKeyset, adLockOptimistic
    lblSongTitle.Caption = Convert_Text_Min(Rs!song_title)
    mysong = Convert_Text_Min(Rs!song_content)
    
    content() = Split(mysong, "$ $")
    lblSongText.Caption = Replace(content(0), "$", vbNewLine)
    cur_stanza = 0
    all_stanza = UBound(content) + 1
    cmdPrev.Visible = False
 
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
End Function

Private Sub cmdFont_Click()
    song_fonttype = song_fonttype + 1
    Select Case song_fonttype
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
            SavedThis = SaveSettings("projection_font_type", song_fonttype)
        Case Else
            song_fonttype = 1
            SavedThis = SaveSettings("projection_font_type", song_fonttype)
    End Select
        lblSongText.fontname = MyFontType(song_fonttype)
End Sub

Private Sub cmdsong_colour_Click()
    song_colour = song_colour + 1
    Select Case song_colour
        Case 1, 2, 3, 4, 5, 6, 7, 8
            SavedThis = SaveSettings("preffered_theme", song_colour)
        Case Else
            song_colour = 1
            SavedThis = SaveSettings("preffered_theme", song_colour)
    End Select
        SetProjectionTheme (song_colour)
End Sub

Private Sub cmdAdd_Click()
    song_fontsize = song_fontsize + 2
    If song_fontsize < 100 Then
        SavedThis = SaveSettings("projection_font_size", song_fontsize)
        lblSongText.FontSize = song_fontsize
    End If
End Sub

Private Sub cmdMinus_Click()
    song_fontsize = song_fontsize - 2
    If song_fontsize > 10 Then
        SavedThis = SaveSettings("projection_font_size", song_fontsize)
        lblSongText.FontSize = song_fontsize
    End If
End Sub

Private Sub cmdNext_Click()
    If songScroll = True Then
        tmrTimerSong.Enabled = True
        tmrTimerTime.Enabled = False
    End If
    
    On Error GoTo ErrorHandler
        lblSongText.Caption = Replace(content(cur_stanza + 1), "$", vbNewLine)
        cur_stanza = cur_stanza + 1
        
        If cur_stanza + 1 = all_stanza Then
            cmdNext.Visible = False
            tmrTimerSong.Enabled = False
            songScroll = False
        Else
            cmdNext.Visible = True
        End If
        cmdPrev.Visible = True
        
        Exit Sub
ErrorHandler:
        tmrTimerSong.Enabled = False
        songScroll = False
        lineDown.BorderColor = &HFF&
        tmrTimerErr.Enabled = True
End Sub

Private Sub cmdPrev_Click()
 
    On Error GoTo ErrorHandler4
        lblSongText.Caption = Replace(content(cur_stanza - 1), "$", vbNewLine)
        cur_stanza = cur_stanza - 1
        
        If cur_stanza = 0 Then
            cmdPrev.Visible = False
        Else
            cmdPrev.Visible = True
        End If
        cmdNext.Visible = True
                
        Exit Sub
ErrorHandler4:
        lineTop.BorderColor = &HFF&
        tmrTimerErr.Enabled = True
End Sub

Private Sub showFirstStanza()
 
    tmrTimerSong.Enabled = False
    songScroll = False
    On Error GoTo ErrorHandler5
        lblSongText.Caption = Replace(content(0), "$", vbNewLine)
        cur_stanza = 0
        cmdPrev.Visible = False
        cmdNext.Visible = True
        Exit Sub
ErrorHandler5:
        lineTop.BorderColor = &HFF&
        tmrTimerErr.Enabled = True
End Sub

Private Sub showLastStanza()
 
    On Error GoTo ErrorHandler8
        lblSongText.Caption = Replace(content(all_stanza - 1), "$", vbNewLine)
        cur_stanza = all_stanza - 1
        cmdPrev.Visible = True
        cmdNext.Visible = False
        Exit Sub
ErrorHandler8:
        lineDown.BorderColor = &HFF&
        tmrTimerErr.Enabled = True
End Sub

Private Sub lblCanvas_Click()
    cmdLine.SetFocus
End Sub

Private Sub lblSongText_Click()
    cmdLine.SetFocus
End Sub

Private Sub tmrTimerErr_Timer()
    SetProjectionTheme (AppSettings("preffered_theme"))
End Sub

Private Sub cmdPlay_Click()
 
    If tmrTimerSong.Enabled = False Then
        songScroll = True
        tmrTimerTime.Enabled = True
        oldTime = Time
    ElseIf tmrTimerSong.Enabled = True Then
        tmrTimerSong.Enabled = False
    End If
End Sub

Private Sub tmrTimerTime_Timer()
 
    Static x As Long
    Static zz$, ss$
    x = x + 1
    newTime = Time
    diff = DateDiff("s", oldTime, newTime)
    songTime = Format((diff - ((diff \ 60) * 60)), "0")
    tmrTimerSong.Interval = songTime * 1000
End Sub

Private Sub tmrTimerSong_Timer()
 
    On Error GoTo ErrorHandler3
        lblSongText.Caption = Replace(content(cur_stanza + 1), "$", vbNewLine)
        cur_stanza = cur_stanza + 1
        
        If cur_stanza + 1 = all_stanza Then
            cmdNext.Visible = False
        Else
            cmdNext.Visible = True
        End If
        cmdPrev.Visible = True
        
        Exit Sub
ErrorHandler3:
        lineDown.BorderColor = &HFF&
        tmrTimerErr.Enabled = True
        tmrTimerSong.Enabled = False
End Sub

Private Sub cmdline_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload frmCcHome
        Unload Me
    End If
End Sub

Private Sub cmdline_KeyDown(KeyCode As Integer, Shift As Integer)
    'Space Key
    If KeyCode = 32 Then
        cmdPlay_Click
    End If
    
     'Page Up Key
    If KeyCode = 33 Then
        showFirstStanza
    End If
    
    'Page Down Key
    If KeyCode = 34 Then
        showLastStanza
    End If
    
    'End Key
    If KeyCode = 35 Then
        showLastStanza
    End If
    
    'Home Key
    If KeyCode = 36 Then
        showFirstStanza
    End If
    
    'Left key
    If KeyCode = 37 Then
        cmdMinus_Click
    End If
    'Up Key
    If KeyCode = 38 Then
        cmdPrev_Click
    End If
    
    'Right Arrow
    If KeyCode = 39 Then
        cmdAdd_Click
    End If
    
    'Down Arrow
    If KeyCode = 40 Then
        cmdNext_Click
    End If
    
    'Key C
    If KeyCode = 67 Then
        cmdFont_Click
    End If
        
    'Key V
    If KeyCode = 86 Then
        cmdFont_Click
    End If
    
    'Key X
    If KeyCode = 88 Then
        cmdsong_colour_Click
    End If
    
    'Key Z
    If KeyCode = 90 Then
        cmdsong_colour_Click
    End If
        
    'Add Key
    If KeyCode = 107 Then
        cmdAdd_Click
    End If
    
    'Subtract Key
    If KeyCode = 109 Then
        cmdMinus_Click
    End If
End Sub

Private Sub trmListen_Timer()
    cmdLine.SetFocus
End Sub



