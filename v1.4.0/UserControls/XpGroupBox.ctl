VERSION 5.00
Begin VB.UserControl XPGroupBox 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   ControlContainer=   -1  'True
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   199
   Begin VB.Image img 
      Height          =   300
      Left            =   840
      Picture         =   "XPGroupBox.ctx":0000
      Top             =   480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "XPGroupBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Teh Ming Han (teh_minghan@hotmail.com)
'The default colour should be RGB(240, 232, 224)
'change it when form loads

Dim m_Font As Font
Dim m_BackColor As OLE_COLOR

Const m_def_BackColor = vbYellow

Event Click()
Attribute Click.VB_UserMemId = -600
Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Private Sub lbl_Change()
    UserControl_Resize
End Sub

Private Sub lbl_Click()
    RaiseEvent Click
End Sub

Private Sub lbl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_Initialize()
    UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
    Caption = Ambient.DisplayName
    Enabled = True
    Set Font = UserControl.Ambient.Font
    BackColor = m_def_BackColor
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
End Sub

Private Sub UserControl_Resize()
    UserControl.ScaleMode = 3
    UserControl.Cls
    UserControl.MaskColor = Hex(ECE9D8)
    Dim brx, bry, bw, bh, bly As Integer
    brx = UserControl.ScaleWidth - 3
    bry = UserControl.ScaleHeight - 3
    bw = UserControl.ScaleWidth - 6
    bh = UserControl.ScaleHeight - 6 - (lbl.Height \ 2)
    bly = lbl.Height \ 2
    lbl.Top = 0
    lbl.Left = 15
    UserControl.PaintPicture img.Picture, 0, bly, 3, 3, 0, 0, 3, 3
    UserControl.PaintPicture img.Picture, brx, bly, 3, 3, 19, 0, 3, 3
    UserControl.PaintPicture img.Picture, brx, bry, 3, 3, 19, 18, 3, 3
    UserControl.PaintPicture img.Picture, 0, bry, 3, 3, 0, 17, 3, 3
    UserControl.PaintPicture img.Picture, 3, bly, bw, 1, 3, 0, 16, 1
    UserControl.PaintPicture img.Picture, brx + 2, bly + 3, 1, bh, 21, 3, 1, 14
    UserControl.PaintPicture img.Picture, 3, bry + 2, bw, 1, 3, 19, 16, 1
    UserControl.PaintPicture img.Picture, 0, bly + 3, 1, bh, 0, 3, 1, 14
    UserControl.MaskColor = Hex(ECE9D8)
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    If Enabled = True Then lbl.ForeColor = RGB(0, 70, 213) Else: lbl.ForeColor = RGB(161, 161, 146)
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lbl.Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
End Sub

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    Set UserControl.Font = vNewFont
    Set lbl.Font = m_Font
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
    lbl.BackColor = m_BackColor
    UserControl_Resize
End Property

Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
    lbl.Caption() = vNewCaption
    Call UserControl_Resize
    PropertyChanged "Caption"
End Property
