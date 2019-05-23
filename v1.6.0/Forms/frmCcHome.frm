VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00EBB85E-60A4-4EB4-A8A3-E451747B2506}#1.0#0"; "TABSMATA.OCX"
Begin VB.MDIForm frmCcHome 
   AutoShowChildren=   0   'False
   BackColor       =   &H000080FF&
   Caption         =   "vSongBook"
   ClientHeight    =   7770
   ClientLeft      =   2775
   ClientTop       =   1605
   ClientWidth     =   13845
   Icon            =   "frmCcHome.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin AsMdiTabs.TabSmata TabSmata 
      Left            =   5040
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      Style           =   0
   End
   Begin VB.PictureBox panelRight 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   13725
      ScaleHeight     =   6660
      ScaleWidth      =   120
      TabIndex        =   8
      Top             =   735
      Width           =   120
   End
   Begin VB.PictureBox panelTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   13845
      TabIndex        =   2
      Top             =   0
      Width           =   13845
      Begin VB.PictureBox cmdSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   520
         Left            =   4560
         Picture         =   "frmCcHome.frx":146B7
         ScaleHeight     =   495
         ScaleWidth      =   465
         TabIndex        =   16
         Top             =   120
         Width           =   495
      End
      Begin VB.Frame fraOptions 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   520
         Left            =   11160
         TabIndex        =   11
         Top             =   120
         Width           =   2535
         Begin VB.PictureBox cmdSettings 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   400
            Left            =   1920
            Picture         =   "frmCcHome.frx":1658B
            ScaleHeight     =   407.755
            ScaleMode       =   0  'User
            ScaleWidth      =   407.755
            TabIndex        =   15
            Top             =   70
            Width           =   400
         End
         Begin VB.PictureBox cmdProject 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   400
            Left            =   120
            Picture         =   "frmCcHome.frx":16A0B
            ScaleHeight     =   407.755
            ScaleMode       =   0  'User
            ScaleWidth      =   407.755
            TabIndex        =   14
            Top             =   70
            Width           =   400
         End
         Begin VB.PictureBox cmdEdit 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   400
            Left            =   720
            Picture         =   "frmCcHome.frx":18865
            ScaleHeight     =   407.755
            ScaleMode       =   0  'User
            ScaleWidth      =   407.755
            TabIndex        =   13
            Top             =   70
            Width           =   400
         End
         Begin VB.PictureBox cmdCompose 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   400
            Left            =   1320
            Picture         =   "frmCcHome.frx":1A6D1
            ScaleHeight     =   407.755
            ScaleMode       =   0  'User
            ScaleWidth      =   407.755
            TabIndex        =   12
            Top             =   70
            Width           =   400
         End
         Begin VB.Shape shpToolBar 
            Height          =   520
            Left            =   0
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   5040
         TabIndex        =   6
         Top             =   120
         Width           =   6015
      End
      Begin VB.ComboBox cmbSongBook 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "frmCcHome.frx":1C560
         Left            =   120
         List            =   "frmCcHome.frx":1C562
         TabIndex        =   4
         Text            =   "Songs Of Worship"
         Top             =   120
         Width           =   4335
      End
      Begin VB.PictureBox cmdSongBooks 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   5
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox cmdSongBook 
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
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   4275
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label lblSongid 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox panelLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   0
      ScaleHeight     =   6660
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   735
      Width           =   4935
      Begin VB.CheckBox chkSearch 
         Appearance      =   0  'Flat
         Caption         =   "Search in all songbooks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   240
         TabIndex        =   7
         Top             =   0
         Width           =   4575
      End
      Begin VB.ListBox lstSongs 
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
         Height          =   6915
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox txtLangFile 
         Height          =   6015
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "frmCcHome.frx":1C564
         Top             =   720
         Width           =   4335
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7395
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   18759
            Text            =   "vSongBook 4PC"
            TextSave        =   "vSongBook 4PC"
            Object.ToolTipText     =   "Turn Your Computer Into a SongBook"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "9/30/2017"
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "2:39 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
      EndProperty
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuTabNew 
         Caption         =   "&New Preview Tab"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuTabClose 
         Caption         =   "&Close Current Tab"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuTabsClose 
         Caption         =   "&Close All Tabs"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAppExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuVsongbook 
      Caption         =   "&vSongBook"
      Begin VB.Menu mnuProjection 
         Caption         =   "&Project Song"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEditSong 
         Caption         =   "&Edit Song"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuNewSong 
         Caption         =   "&New Song"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuvSongs 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSongBooks 
         Caption         =   "&SongBooks"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuStyle 
      Caption         =   "&Styles"
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style 1"
         Index           =   0
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style 2"
         Index           =   1
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style 3"
         Index           =   2
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style 4"
         Index           =   3
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style 5"
         Index           =   4
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style 6"
         Index           =   5
      End
   End
   Begin VB.Menu mnuNav 
      Caption         =   "&Navigation-Style"
      Begin VB.Menu mnuNavi 
         Caption         =   "Scroll Buttons"
         Index           =   0
      End
      Begin VB.Menu mnuNavi 
         Caption         =   "Dropdown Button"
         Index           =   1
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help Desk"
      Begin VB.Menu mnuHowWorks 
         Caption         =   "&How it Works"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuHelpDesk 
         Caption         =   "&Help Desk"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "&Language"
      Begin VB.Menu mnuAfrikaans 
         Caption         =   "&Afrikaans"
      End
      Begin VB.Menu mnuChichewa 
         Caption         =   "&Chichewa"
      End
      Begin VB.Menu mnuEnglish 
         Caption         =   "&English"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFrench 
         Caption         =   "&French"
      End
      Begin VB.Menu mnuPortuguese 
         Caption         =   "&Portuguese"
      End
      Begin VB.Menu mnuSpanish 
         Caption         =   "&Spanish"
      End
      Begin VB.Menu mnuSwahili 
         Caption         =   "&Swahili"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmCcHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim Songbook_found As Boolean, Songs_found As Boolean, SavedThis As Boolean
Dim mytitle As String, mysong As String, searched As String, song_id As Integer
Dim MDIForms As Integer, MyAnswer As VbMsgBoxResult, MsgBoxTitle As String, MsgBoxText As String
Dim fLang As Integer, tLang As String, nlang As String, iLang As Long

Private Sub checkClosing()
    If MDIForms > 2 Then
        MsgBoxTitle = "Check closing vSongBook?"
        MsgBoxText = "You are about to close " & MDIForms & " tabs. Are you sure you want to continue?"
        MyAnswer = MsgBox(MsgBoxText, vbOKCancel, MsgBoxTitle)
        Select Case MyAnswer
            Case vbOK
                Unload Me
            Case vbCancel
                'MsgBox "You've pressed Cancel"
        End Select
    End If
End Sub

Private Sub chkSearch_Click()
    If txtSearch.Text = "" Then
    
    Else
        If chkSearch.Value = vbChecked Then
            lstSongs.Clear
            normalSearch
        ElseIf chkSearch.Value = vbUnchecked Then
            lstSongs.Clear
            strictSearch
        End If
    End If
End Sub

Public Sub show_Projection_Window()
    If lstSongs.Text = "" Then
    
    Else
        frmCcProject.Show
    End If
End Sub

Public Sub show_Settings_Window()
    Me.Enabled = False
    frmEeOptions.Show , Me
End Sub

Public Sub show_Compose_Window()
    Me.Enabled = False
    frmFfNewsong.Show , Me
End Sub

Public Sub show_Edit_Window()
    Me.Enabled = False
    frmFfEditSong.Show , Me
End Sub

Public Sub show_Help_Window()
    Me.Enabled = False
    frmDdHelp.Show , Me
End Sub

Public Sub show_Infor_Window()
    Me.Enabled = False
    frmDdInfo.Show , Me
End Sub

Public Sub show_SongBook_Window()
    Me.Enabled = False
    frmBbSongbook.Show , Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sndplaysound (App.Path & "\Tools\close.wav"), 1
End Sub

Private Sub cmdCompose_Click()
    show_Compose_Window
End Sub

Private Sub cmdEdit_Click()
    show_Edit_Window
End Sub

Private Sub cmdProject_Click()
    show_Projection_Window
End Sub

Private Sub cmdSearch_Click()
    txtSearch.SetFocus
End Sub

Private Sub cmdSettings_Click()
    show_Settings_Window
End Sub

Private Sub cmdSongBook_Click()
    show_SongBook_Window
End Sub

Private Sub cmdSongBooks_Click()
    show_SongBook_Window
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'checkClosing
    'Cancel = 1
    Unload Me
End Sub

Private Sub mnuAfrikaans_Click()
    SavedThis = SaveSettings("preffered_lang", "Afrikaans")
    mnuAfrikaans.Checked = True
    mnuChichewa.Checked = False
    mnuEnglish.Checked = False
    mnuFrench.Checked = False
    mnuPortuguese.Checked = False
    mnuSpanish.Checked = False
    mnuSwahili.Checked = False
    readLangTexts
End Sub

Private Sub mnuChichewa_Click()
    SavedThis = SaveSettings("preffered_lang", "Chichewa")
    mnuAfrikaans.Checked = False
    mnuChichewa.Checked = True
    mnuEnglish.Checked = False
    mnuFrench.Checked = False
    mnuPortuguese.Checked = False
    mnuSpanish.Checked = False
    mnuSwahili.Checked = False
    readLangTexts
End Sub

Private Sub mnuEnglish_Click()
    SavedThis = SaveSettings("preffered_lang", "English")
    mnuAfrikaans.Checked = False
    mnuChichewa.Checked = False
    mnuEnglish.Checked = True
    mnuFrench.Checked = False
    mnuPortuguese.Checked = False
    mnuSpanish.Checked = False
    mnuSwahili.Checked = False
    readLangTexts
End Sub

Private Sub mnuFrench_Click()
    SavedThis = SaveSettings("preffered_lang", "French")
    mnuAfrikaans.Checked = False
    mnuChichewa.Checked = False
    mnuEnglish.Checked = False
    mnuFrench.Checked = True
    mnuPortuguese.Checked = False
    mnuSpanish.Checked = False
    mnuSwahili.Checked = False
    readLangTexts
End Sub

Private Sub mnuPortuguese_Click()
    SavedThis = SaveSettings("preffered_lang", "Portuguese")
    mnuAfrikaans.Checked = False
    mnuChichewa.Checked = False
    mnuEnglish.Checked = False
    mnuFrench.Checked = False
    mnuPortuguese.Checked = True
    mnuSpanish.Checked = False
    mnuSwahili.Checked = False
    readLangTexts
End Sub

Private Sub mnuSpanish_Click()
    SavedThis = SaveSettings("preffered_lang", "Spanish")
    mnuAfrikaans.Checked = False
    mnuChichewa.Checked = False
    mnuEnglish.Checked = False
    mnuFrench.Checked = False
    mnuPortuguese.Checked = False
    mnuSpanish.Checked = True
    mnuSwahili.Checked = False
    readLangTexts
End Sub

Private Sub mnuSwahili_Click()
    SavedThis = SaveSettings("preffered_lang", "Swahili")
    mnuAfrikaans.Checked = False
    mnuChichewa.Checked = False
    mnuEnglish.Checked = False
    mnuFrench.Checked = False
    mnuPortuguese.Checked = False
    mnuSpanish.Checked = False
    mnuSwahili.Checked = True
    readLangTexts
End Sub

Private Sub mnuTabsClose_Click()
    Do Until ActiveForm Is Nothing
        Unload ActiveForm
    Loop
    MDIForms = 0
End Sub

Private Sub TabSmata_ColorChanged(NewColor As stdole.OLE_COLOR)
    '--Here you can assign any control color --> NewColor
    '  for eg,
    '  Picture1.BackColor = NewColor
    '  Where NewColor is a one of the color generated for OneNote style
End Sub

Private Sub TabSmata_DropdownButtonClick()
    PopupMenu mnuWindow
End Sub


Private Sub MDIForm_Load()
    TabSmata.DrawIcons = Not TabSmata.DrawIcons
    'mnuStyles_Click (4)
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    
    MDIForms = 0
    checkStyles
    checkDefaultLang
    Songbook_found = False
    If AppSettings("last_window_startup") = 0 Then
        Me.WindowState = vbNormal
        Me.Height = AppSettings("last_window_height")
        Me.Width = AppSettings("last_window_width")
    ElseIf AppSettings("last_window_startup") = 1 Then
        Me.WindowState = vbMaximized
    End If
    InitializeSongBook
    
End Sub

Private Sub checkStyles()
    Dim navis As Integer, tabbs As Integer
    navis = AppSettings("navi_style")
    tabbs = AppSettings("tab_style")
    
    mnuStyles(tabbs).Checked = True
    TabSmata.Style = tabbs
    mnuNavi(navis).Checked = True
    TabSmata.NavigationStyle = navis
End Sub

Private Sub checkDefaultLang()
    Dim preflang As String
    preflang = AppSettings("preffered_lang")
    
    If preflang = "Afrikaans" Then
        mnuAfrikaans.Checked = True
    ElseIf preflang = "Chichewa" Then
        mnuChichewa.Checked = True
    ElseIf preflang = "English" Then
        mnuEnglish.Checked = True
    ElseIf preflang = "French" Then
        mnuFrench.Checked = True
    ElseIf preflang = "Portuguese" Then
        mnuPortuguese.Checked = True
    ElseIf preflang = "Spanish" Then
        mnuSpanish.Checked = True
    ElseIf preflang = "Swahili" Then
        mnuSwahili.Checked = True
    End If
    readLangTexts
End Sub

Private Sub MDIForm_Resize()
    If (frmCcHome.Height > 7000 And frmCcHome.Width > 7000) Then
        lstSongs.Height = frmCcHome.Height - 2380
        txtSearch.Width = frmCcHome.Width - 8100
        fraOptions.left = txtSearch.left + txtSearch.Width + 120
    Else
        Exit Sub
    End If
    
    If Me.WindowState = vbMaximized Then
        SavedThis = SaveSettings("last_window_startup", "1")
    ElseIf Me.WindowState = vbNormal Then
        SavedThis = SaveSettings("last_window_startup", "0")
        SavedThis = SaveSettings("last_window_height", frmCcHome.Height)
        SavedThis = SaveSettings("last_window_width", frmCcHome.Width)
    End If
    
End Sub

Private Sub mnuAppExit_Click()
    checkClosing
End Sub

Private Sub mnuEditSong_Click()
    show_Edit_Window
End Sub

Private Sub mnuHelpDesk_Click()
    show_Help_Window
End Sub

Private Sub mnuHowWorks_Click()
    show_Infor_Window
End Sub

Private Sub mnuNewSong_Click()
    show_Compose_Window
End Sub

Private Sub mnuProjection_Click()
    show_Projection_Window
End Sub

Private Sub mnuSettings_Click()
    show_Settings_Window
End Sub

Private Sub mnuSongBooks_Click()
    show_SongBook_Window
End Sub

Private Sub mnuTabNew_Click()
   Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_content WHERE songid = " & lblSongid.Caption & "", con, adOpenKeyset, adLockOptimistic
    
    Dim pre_view As New frmCcSong
    pre_view.Caption = Rs!song_title
    pre_view.Show
    MDIForms = MDIForms + 1
End Sub

Private Sub mnuTabClose_Click()
    If ActiveForm Is Nothing Then Exit Sub
    Unload ActiveForm
    MDIForms = MDIForms - 1
End Sub

Private Sub mnuFocusrect_Click()
    'mnuFocusrect.Checked = Not mnuFocusrect.Checked
    TabSmata.ShowFocusRect = Not TabSmata.ShowFocusRect
End Sub

Private Sub mnuHelpAbout_Click()
    'TabSmata.About
End Sub


Private Sub mnuNavi_Click(Index As Integer)
Dim i As Long
    For i = 0 To 1
        mnuNavi(i).Checked = False
    Next i
    SavedThis = SaveSettings("navi_style", Index)
    mnuNavi(Index).Checked = True
    TabSmata.NavigationStyle = Index
End Sub

Private Sub mnuPopupClose_Click()
    'mnuFileClose_Click
End Sub

Private Sub TabSmata_TabBarClick(Button As Integer, X As Long, Y As Long)
    Debug.Print "TabBarClick (" & Button & ", " & X & ", " & Y & ")"
End Sub

Private Sub TabSmata_TabClick(TabHwnd As Long, Button As Integer, X As Long, Y As Long)
    Debug.Print "TabClick (" & TabHwnd & ", " & Button & ", " & X & ", " & Y & ")"
    If Button = vbRightButton Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub mnuStyles_Click(Index As Integer)
Dim i As Long
    For i = 0 To 5
        mnuStyles(i).Checked = False
    Next i
    SavedThis = SaveSettings("tab_style", Index)
    mnuStyles(Index).Checked = True
    TabSmata.Style = Index
End Sub

Public Sub InitializeSongBook()
    songbook_list
    If Songbook_found Then
        cmdSongBook.Visible = False
        cmbSongBook.Enabled = True
        chkSearch.Enabled = True
        lstSongs.Enabled = True
        txtSearch.Enabled = True
        cmdCompose.Enabled = True
        cmdEdit.Enabled = True
        cmdProject.Enabled = True
        
        If AppSettings("show_help_first") = "0" Then
            show_Infor_Window
            SavedThis = SaveSettings("show_help_first", "1")
        End If
        
    Else
        cmdSongBook.Visible = True
        cmbSongBook.Enabled = False
        chkSearch.Enabled = False
        lstSongs.Enabled = False
        txtSearch.Enabled = False
        cmdCompose.Enabled = False
        cmdEdit.Enabled = False
        cmdProject.Enabled = False
        fraOptions.Enabled = False
    End If
End Sub

Public Sub ReinitializeSongbook()
cmbSongBook.Clear

    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_book WHERE sb_enabled =0", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        cmbSongBook.AddItem Rs!sb_title
        Rs.MoveNext
    Loop
    
    Rs.Close
    Songbook_found = True
    openSongBook
    cmdSongBook.Visible = False
    cmbSongBook.Enabled = True
    chkSearch.Enabled = True
    lstSongs.Enabled = True
    txtSearch.Enabled = True
    cmdCompose.Enabled = True
    cmdEdit.Enabled = True
    cmdProject.Enabled = True
    
    On Error Resume Next
        cmbSongBook.ListIndex = 0
        Me.Caption = cmbSongBook.Text & " songs - vSongBook | " & AppSettings("user_name")
    
End Sub

Public Sub cmbSongBook_Click()
    lstSongs.Clear
    openSongBook
End Sub

Private Sub openSongBook()
lstSongs.Clear
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_content WHERE song_book = '" & cmbSongBook.Text & "' ORDER BY songid", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstSongs.AddItem Convert_Text_Min(Rs!song_title)
        Rs.MoveNext
    Loop
    
    On Error Resume Next
        lstSongs.ListIndex = 0
    
    Me.Caption = lstSongs.ListCount & " " & cmbSongBook.Text & " songs - vSongBook | " & AppSettings("user_name")
    
    If lstSongs.Text = "" Then
        cmdEdit.Enabled = False
        cmdProject.Enabled = False
    Else
        cmdEdit.Enabled = True
        cmdProject.Enabled = True
    End If
End Sub

Private Sub songbook_list()
On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_book where sb_enabled =0", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        cmbSongBook.AddItem Rs!sb_title
        Rs.MoveNext
    Loop
    Songbook_found = True
        cmbSongBook.ListIndex = 0
    Me.Caption = cmbSongBook.Text & " songs - vSongBook | " & AppSettings("user_name")
    
   Exit Sub
ErrorHandler:
Songbook_found = False
'MsgBox Err.Description & " No. " & Err.Number
    MsgBox "Unable to find songbooks in the database. You might have to add songbooks and later their " & _
    "various songs in them before starting to use vSongBook" & vbNewLine & _
    "1. Click on the " & Chr$(34) & "Add a SongBook" & Chr$(34) & _
    " button to add songbook first." & vbNewLine & _
    "2. Later click the " & Chr$(34) & "Add New Song" & Chr$(34) & " Button on the Toolbar at the bottom", vbInformation, "vSongBook"
   'frmBbSongbook.Show
End Sub

Public Function SaveSettings(pref_title, pref_content) As Boolean
    'On Error Resume Next
        Set Rs = New ADODB.Recordset
        Rs.Open "Select * from app_options where option_title ='" & pref_title & "'", con, adOpenKeyset, adLockOptimistic
        Rs!option_content = pref_content
        Rs.Update
        SaveSettings = True
End Function

Public Function AppSettings(option_title) As String
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from app_options WHERE option_title='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
    AppSettings = Rs!option_content
    Rs.Close
End Function

Private Sub lstSongs_Click()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_content WHERE song_title = '" & lstSongs.Text & "'", con, adOpenKeyset, adLockOptimistic
    lblSongid.Caption = Rs!songid
    Dim pre_view As New frmCcSong
    pre_view.Caption = Rs!song_title
    pre_view.Show
    MDIForms = MDIForms + 1
End Sub

Private Sub txtSearch_Change()
    If chkSearch.Value = vbChecked Then
        lstSongs.Clear
        normalSearch
    ElseIf chkSearch.Value = vbUnchecked Then
        lstSongs.Clear
        strictSearch
    End If
End Sub

Private Sub strictSearch()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_content WHERE song_book = '" & cmbSongBook.Text & "' AND song_content LIKE '%" & txtSearch.Text & "%'", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstSongs.AddItem Rs!song_title
        Rs.MoveNext
    Loop
    'lstSongs.ListIndex = 0
    Me.Caption = lstSongs.ListCount & " songs found with " & Chr$(34) & txtSearch.Text & Chr$(34) & " - vSongBook | " & AppSettings("user_name")
    
End Sub


Private Sub normalSearch()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_book WHERE song_content LIKE '%" & txtSearch.Text & "%'", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstSongs.AddItem Rs!song_title
        Rs.MoveNext
    Loop
    'lstSongs.ListIndex = 0
    Me.Caption = lstSongs.ListCount & " songs found with " & Chr$(34) & txtSearch.Text & Chr$(34) & " - vSongBook | " & AppSettings("user_name")
    
End Sub

Private Sub txtSearch_GotFocus()
    Static bSet As Boolean
    If Not bSet Then
        txtSearch.Text = ""
        bSet = True
    End If
End Sub


Public Sub readLangTexts()
   Open App.Path & "\Langs\" & AppSettings("preffered_lang") & ".txt" For Input As #1
   txtLangFile.Text = Input$(LOF(1), #1)

   Me.Caption = getLangString(txtLangFile.Text, 2)
   mnuFileTop.Caption = getLangString(txtLangFile.Text, 18)
   mnuTabNew.Caption = getLangString(txtLangFile.Text, 19)
   mnuTabClose.Caption = getLangString(txtLangFile.Text, 20)
   mnuTabsClose.Caption = getLangString(txtLangFile.Text, 21)
   mnuAppExit.Caption = getLangString(txtLangFile.Text, 22)
   mnuVsongbook.Caption = getLangString(txtLangFile.Text, 2)
   mnuProjection.Caption = getLangString(txtLangFile.Text, 23)
   mnuEditSong.Caption = getLangString(txtLangFile.Text, 24)
   mnuNewSong.Caption = getLangString(txtLangFile.Text, 25)
   mnuSongBooks.Caption = getLangString(txtLangFile.Text, 26)
   mnuSettings.Caption = getLangString(txtLangFile.Text, 27)
   mnuStyle.Caption = getLangString(txtLangFile.Text, 28)
   'mnuStyles.Caption = getLangString(txtLangFile.Text, 29)
   mnuNav.Caption = getLangString(txtLangFile.Text, 30)
   'mnuNavi.Caption = getLangString(txtLangFile.Text, 31)
   mnuWindow.Caption = getLangString(txtLangFile.Text, 33)
   mnuHelpTop.Caption = getLangString(txtLangFile.Text, 35)
   mnuHowWorks.Caption = getLangString(txtLangFile.Text, 34)
   mnuHelpDesk.Caption = getLangString(txtLangFile.Text, 35)
   mnuLanguage.Caption = getLangString(txtLangFile.Text, 96)
   mnuPopup.Caption = getLangString(txtLangFile.Text, 36)
   mnuPopupClose.Caption = getLangString(txtLangFile.Text, 37)
   
   'cmdSongBook.Caption = getLangString(txtLangFile.Text, 38)
   chkSearch.Caption = getLangString(txtLangFile.Text, 39)
   
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

