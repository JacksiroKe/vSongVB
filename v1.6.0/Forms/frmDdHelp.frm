VERSION 5.00
Begin VB.Form frmDdHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "vSongBook HelpDesk"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDdHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCommunication 
      Appearance      =   0  'Flat
      Height          =   3705
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   10695
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
      Height          =   2175
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmDdHelp.frx":146B7
      Top             =   2880
      Width           =   5415
   End
   Begin VB.Shape Shape2 
      Height          =   3975
      Left            =   120
      Top             =   2280
      Width           =   10935
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   120
      Top             =   960
      Width           =   10935
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   11040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "vSongBook Help Desk"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   10695
   End
   Begin VB.Label lblExplanation 
      Caption         =   $"frmDdHelp.frx":146BD
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   10695
   End
End
Attribute VB_Name = "frmDdHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim lRegion As Long

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\vSongBook.mdb;"
    con.Open
    readLangTexts
    
    lstCommunication.AddItem " 1. Call/Text/WhatsApp - +254 711 474 787"
    lstCommunication.AddItem " 2. Call/Text/WhatsApp - +254 731 973 180"
    lstCommunication.AddItem " 3. Email Address - jaksiro@gmail.com"
    lstCommunication.AddItem " 4. Website - www.jacksiro.wordpress.com"
    lstCommunication.AddItem " 5. Linkedin.com/sillajackson - Jackson Siro"
    lstCommunication.AddItem " 6. Instagram - @jaksiro - Jack Siro"
    lstCommunication.AddItem " 7. Facebook.com/jaksiro - Jack Siro"
    lstCommunication.AddItem " 8. Twitter - @jaksiro - Jack Siro"
    
End Sub

Private Sub readLangTexts()
   Open App.Path & "\Langs\" & AppSettings("preffered_lang") & ".txt" For Input As #1
   txtLangFile.Text = Input$(LOF(1), #1)

   Me.Caption = getLangString(txtLangFile.Text, 64)
   lblExplanation.Caption = getLangString(txtLangFile.Text, 65)
    
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
