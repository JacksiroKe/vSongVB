VERSION 5.00
Begin VB.Form frmDdInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "How vSongBook Works"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDdInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTab1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   2840
      Width           =   10900
      Begin VB.ListBox lstHowStuff_i 
         Appearance      =   0  'Flat
         Height          =   3705
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   10580
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
         Height          =   1935
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmDdInfo.frx":146B7
         Top             =   720
         Width           =   5775
      End
      Begin VB.Shape shpTab1 
         Height          =   3975
         Left            =   0
         Top             =   0
         Width           =   10815
      End
   End
   Begin VB.Frame fraTab2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4065
      Left            =   120
      TabIndex        =   4
      Top             =   2840
      Width           =   10900
      Begin VB.ListBox lstHowStuff_ii 
         Appearance      =   0  'Flat
         Height          =   3705
         ItemData        =   "frmDdInfo.frx":146BD
         Left            =   120
         List            =   "frmDdInfo.frx":146BF
         TabIndex        =   5
         Top             =   120
         Width           =   10580
      End
      Begin VB.Shape shpTab2 
         Height          =   3975
         Left            =   0
         Top             =   0
         Width           =   10815
      End
   End
   Begin vSongBook.XPButton cmdTab1 
      Height          =   705
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1244
      TX              =   "Pojection Mode"
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
      MICON           =   "frmDdInfo.frx":146C1
   End
   Begin vSongBook.XPButton cmdTab2 
      Height          =   705
      Left            =   4440
      TabIndex        =   8
      Top             =   2280
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1244
      TX              =   "Searching Mode"
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
      MICON           =   "frmDdInfo.frx":146DD
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   120
      Top             =   840
      Width           =   10815
   End
   Begin VB.Label lblExplanation 
      Caption         =   $"frmDdInfo.frx":146F9
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   10575
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   10920
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblHowWorks 
      Alignment       =   2  'Center
      Caption         =   "How vSongBook Works"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmDdInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim lRegion As Long

Private Sub cmdTab1_Click()
    cmdTab1.Top = 2160
    cmdTab2.Top = 2220
    fraTab1.Visible = True
    fraTab2.Visible = False
End Sub

Private Sub cmdTab2_Click()
    cmdTab1.Top = 2220
    cmdTab2.Top = 2160
    fraTab1.Visible = False
    fraTab2.Visible = True
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    readLangTexts
    
    'SEARCHING MODE
    lstHowStuff_i.AddItem " 1. " & Chr$(34) & "Ctrl + A" & Chr$(34) & " - New Preview Tab"
    lstHowStuff_i.AddItem " 2. " & Chr$(34) & "Ctrl + W" & Chr$(34) & " - Close Current Tab"
    lstHowStuff_i.AddItem " 3. " & Chr$(34) & "Ctrl + Q" & Chr$(34) & " - Close All Tabs"
    lstHowStuff_i.AddItem " 4. " & Chr$(34) & "Ctrl + X" & Chr$(34) & " - Exit"
    lstHowStuff_i.AddItem " 5. " & Chr$(34) & "F5" & Chr$(34) & " - Project Song"
    lstHowStuff_i.AddItem " 6. " & Chr$(34) & "F6" & Chr$(34) & " - Edit Song"
    lstHowStuff_i.AddItem " 7. " & Chr$(34) & "Ctrl + N" & Chr$(34) & " - New Song"
    lstHowStuff_i.AddItem " 8. " & Chr$(34) & "F8" & Chr$(34) & " - Manage SongBooks"
    lstHowStuff_i.AddItem " 9. " & Chr$(34) & "F9" & Chr$(34) & " - vSongBook Settings"
    lstHowStuff_i.AddItem " 10. " & Chr$(34) & "F2" & Chr$(34) & " - How it Works"
    lstHowStuff_i.AddItem " 11. " & Chr$(34) & "F1" & Chr$(34) & " - Help Desk"
    
    'PROJECTION MODE
    lstHowStuff_ii.AddItem " 1. Letter " & Chr$(34) & "Z" & Chr$(34) & " - Change Projection Theme"
    lstHowStuff_ii.AddItem " 2. Letter " & Chr$(34) & "X" & Chr$(34) & " - Change Projection Theme"
    lstHowStuff_ii.AddItem " 3. Letter " & Chr$(34) & "C" & Chr$(34) & " - Change Font Type"
    lstHowStuff_ii.AddItem " 4. Letter " & Chr$(34) & "V" & Chr$(34) & " - Change Font Type"
    lstHowStuff_ii.AddItem " 5. Key " & Chr$(34) & "-" & Chr$(34) & " - Decrease Font Size"
    lstHowStuff_ii.AddItem " 6. Key " & Chr$(34) & "+" & Chr$(34) & " - Increase Font Size"
    lstHowStuff_ii.AddItem " 7. " & Chr$(34) & "Up" & Chr$(34) & " Arrow - Go Previous Song Stanza"
    lstHowStuff_ii.AddItem " 8. " & Chr$(34) & "Down" & Chr$(34) & " Arrow - Go to Next Song Stanza"
    'lstHowStuff_ii.AddItem " 9. " & Chr$(34) & "Left" & Chr$(34) & " Arrow - Go to Previous Song"
    'lstHowStuff_ii.AddItem " 10. " & Chr$(34) & "Right" & Chr$(34) & " Arrow - Go to Next Song"
    lstHowStuff_ii.AddItem " 9. " & Chr$(34) & "Left" & Chr$(34) & " Arrow - Reduce Font Size"
    lstHowStuff_ii.AddItem " 10. " & Chr$(34) & "Right" & Chr$(34) & " Arrow - Increase Font Size"
    lstHowStuff_ii.AddItem " 11. " & Chr$(34) & "PgUp" & Chr$(34) & " Key - Go to First Song Stanza"
    lstHowStuff_ii.AddItem " 12. " & Chr$(34) & "PgDn" & Chr$(34) & " Key - Go to Last Song Stanza"
    lstHowStuff_ii.AddItem " 13. " & Chr$(34) & "Home" & Chr$(34) & " Key - Go to First Song Stanza"
    lstHowStuff_ii.AddItem " 14. " & Chr$(34) & "End" & Chr$(34) & " Key - Go to Last Song Stanza"
    lstHowStuff_ii.AddItem " 15. " & Chr$(34) & "Esc" & Chr$(34) & " Key - Close Projection Mode"
    
End Sub

Private Sub readLangTexts()
   Open App.Path & "\Langs\" & AppSettings("preffered_lang") & ".txt" For Input As #1
   txtLangFile.Text = Input$(LOF(1), #1)

   Me.Caption = getLangString(txtLangFile.Text, 68)
   lblHowWorks.Caption = getLangString(txtLangFile.Text, 68)
   lblExplanation.Caption = getLangString(txtLangFile.Text, 69)
   cmdTab1.Caption = getLangString(txtLangFile.Text, 70)
   cmdTab2.Caption = getLangString(txtLangFile.Text, 71)
    
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

Public Function AppSettings(option_title) As String
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from app_options WHERE option_title='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
    AppSettings = Rs!option_content
    Rs.Close
End Function

Private Sub Form_Unload(Cancel As Integer)
    frmCcHome.Enabled = True
End Sub

