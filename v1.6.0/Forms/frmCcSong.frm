VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCcSong 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Untitled - SongView"
   ClientHeight    =   7545
   ClientLeft      =   5940
   ClientTop       =   4410
   ClientWidth     =   12570
   ControlBox      =   0   'False
   Icon            =   "frmCcSong.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   12570
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSongCont 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   12495
   End
   Begin VB.PictureBox picToolBar 
      Align           =   2  'Align Bottom
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   12570
      TabIndex        =   0
      Top             =   6810
      Width           =   12570
      Begin vSongBook.XPButton cmdFontType 
         Height          =   615
         Left            =   5150
         TabIndex        =   7
         Top             =   50
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   1085
         TX              =   "Font Type"
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
         MICON           =   "frmCcSong.frx":146B7
      End
      Begin vSongBook.XPButton cmdEdit 
         Height          =   615
         Left            =   2600
         TabIndex        =   6
         Top             =   50
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   1085
         TX              =   "Edit"
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
         MICON           =   "frmCcSong.frx":146D3
      End
      Begin vSongBook.XPButton cmdProject 
         Height          =   615
         Left            =   50
         TabIndex        =   5
         Top             =   50
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   1085
         TX              =   "Project"
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
         MICON           =   "frmCcSong.frx":146EF
      End
      Begin MSComctlLib.Slider sldFontSize 
         Height          =   555
         Left            =   7920
         TabIndex        =   4
         Top             =   120
         Width           =   4155
         _ExtentX        =   7329
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
      End
      Begin VB.Label lblSongID 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6240
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox txtLangFile 
      Height          =   2895
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmCcSong.frx":1470B
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frmCcSong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim mysong As String, cur_songid As String
Dim fontype As Integer, SavedThis As Boolean

Private Sub cmd_Edit_Click()
    frmCcHome.lblSongid.Caption = lblSongid.Caption
    frmCcHome.show_Edit_Window
End Sub

Private Sub cmd_Project_Click()
    frmCcHome.lblSongid.Caption = lblSongid.Caption
    frmCcHome.show_Projection_Window
End Sub

Private Sub cmdEdit_Click()
    frmCcHome.lblSongid.Caption = lblSongid.Caption
    frmCcHome.show_Edit_Window
End Sub

Private Sub cmdProject_Click()
    frmCcHome.lblSongid.Caption = lblSongid.Caption
    frmCcHome.show_Projection_Window
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    readLangTexts
    
    cur_songid = frmCcHome.lblSongid.Caption
    txtSongCont.FontSize = frmCcHome.AppSettings("preview_font_size")
    sldFontSize.Value = frmCcHome.AppSettings("preview_font_size")
    txtSongCont.fontname = MyFontType(frmCcHome.AppSettings("preview_font_type"))
    
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from song_content WHERE songid =" & cur_songid & "", con, adOpenKeyset, adLockOptimistic
    Me.Caption = ConvertSong(Rs!song_title)
    txtSongCont.Text = ConvertSong(Rs!song_content)
    lblSongid.Caption = Rs!songid
End Sub

Private Sub readLangTexts()
   Open App.Path & "\Langs\" & frmCcHome.AppSettings("preffered_lang") & ".txt" For Input As #1
   txtLangFile.Text = Input$(LOF(1), #1)

   cmdProject.Caption = getLangString(txtLangFile.Text, 85)
   cmdEdit.Caption = getLangString(txtLangFile.Text, 86)
   cmdFontType.Caption = getLangString(txtLangFile.Text, 87)
   
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

Public Function ConvertSong(songtext) As String
    mysong = Replace(songtext, "+", Chr$(34))
    mysong = Replace(mysong, "^", "'")
    mysong = Replace(mysong, "$", vbNewLine)
    ConvertSong = mysong
End Function

Private Sub Form_Resize()
On Error GoTo formErrHandler
    txtSongCont.left = 100
    txtSongCont.Width = Me.Width - 300
    txtSongCont.Top = 0
    txtSongCont.Height = Me.Height - 1400
    'cmdProject.left = 100
    'cmdProject.Width = picToolBar.Width / 5
    'cmdEdit.Width = cmdProject.Width
    'cmdFontType.Width = cmdProject.Width
    'sldFontSize.Width = cmdProject.Width * 2
    'cmdEdit.left = cmdProject.left + cmdProject.Width + 100
    'cmdFontType.left = cmdEdit.left + cmdEdit.Width + 100
    'sldFontSize.left = cmdFontType.left + cmdFontType.Width + 100
formErrHandler:
   'jkjk
End Sub

Private Sub sldFontSize_Scroll()
    txtSongCont.FontSize = sldFontSize.Value
    SavedThis = frmCcHome.SaveSettings("preview_font_size", sldFontSize.Value)
End Sub

Private Sub cmdFontType_Click()
    fontype = fontype + 1
    Select Case fontype
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
            SavedThis = frmCcHome.SaveSettings("preview_font_type", fontype)
        Case Else
            fontype = 1
            SavedThis = frmCcHome.SaveSettings("preview_font_type", fontype)
    End Select
        txtSongCont.fontname = MyFontType(fontype)
End Sub

