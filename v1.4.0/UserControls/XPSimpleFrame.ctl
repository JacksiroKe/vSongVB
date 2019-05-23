VERSION 5.00
Begin VB.UserControl XPSimpleFrame 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ControlContainer=   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   4770
   ToolboxBitmap   =   "XPSimpleFrame.ctx":0000
   Begin VB.Shape Shape1 
      BorderColor     =   &H00B99D7F&
      Height          =   540
      Left            =   675
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "XPSimpleFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'       --------------------------------
'       ------------- By ---------------
'       ----- Mohammad Ali Sohrabi -----
'       ------ ali6236@yahoo.com -------
'       ------- !!!Freeware!!! ---------
'       --------------------------------

Private Sub UserControl_Resize()
    Shape1.Move 0, 0, ScaleWidth, ScaleHeight
    AutoSizeContained = True
End Sub
Property Get AutoSizeContained() As Boolean
    AutoSizeContained = False
End Property
Property Let AutoSizeContained(NewVal As Boolean)
On Error Resume Next
    Dim i As Object
    Set i = UserControl.ContainedControls(0)
    If i Is Nothing Then Exit Property
    i.BorderStyle = 0
    i.Appearance = 0
    i.Move 45, 45, ScaleWidth - 90, ScaleHeight - 90
    If i.Height <> ScaleHeight - 90 Then
        UserControl.Height = i.Height + 90
    End If
End Property

