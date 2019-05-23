VERSION 5.00
Begin VB.Form frmUpdate 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox panelLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   0
      ScaleHeight     =   5445
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.TextBox txtLangFile 
         Height          =   6855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmUpdate.frx":0000
         Top             =   360
         Width           =   4695
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
         TabIndex        =   2
         Top             =   480
         Width           =   4695
      End
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
         Left            =   360
         TabIndex        =   1
         Top             =   -120
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
