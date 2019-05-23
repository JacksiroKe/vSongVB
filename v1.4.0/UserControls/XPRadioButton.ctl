VERSION 5.00
Begin VB.UserControl XPRadioButton 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   FillStyle       =   0  'Solid
   MaskColor       =   &H00000000&
   ScaleHeight     =   1335
   ScaleWidth      =   1095
   ToolboxBitmap   =   "XPRadioButton.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   360
   End
   Begin VB.Image img 
      Height          =   195
      Index           =   7
      Left            =   840
      Picture         =   "XPRadioButton.ctx":0312
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image img 
      Height          =   195
      Index           =   6
      Left            =   600
      Picture         =   "XPRadioButton.ctx":0382
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image img 
      Height          =   195
      Index           =   5
      Left            =   360
      Picture         =   "XPRadioButton.ctx":04DF
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image img 
      Height          =   195
      Index           =   4
      Left            =   120
      Picture         =   "XPRadioButton.ctx":0648
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image img 
      Height          =   195
      Index           =   3
      Left            =   840
      Picture         =   "XPRadioButton.ctx":079F
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image img 
      Height          =   195
      Index           =   2
      Left            =   600
      Picture         =   "XPRadioButton.ctx":07FA
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image img 
      Height          =   195
      Index           =   1
      Left            =   360
      Picture         =   "XPRadioButton.ctx":08DF
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image img 
      Height          =   195
      Index           =   0
      Left            =   120
      Picture         =   "XPRadioButton.ctx":0A3B
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image p 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   120
      Top             =   120
      Width           =   195
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "XPRadioButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Teh Ming Han (teh_minghan@hotmail.com)

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long

Private Type POINT_API
    x As Long
    Y As Long
End Type

Dim m_Font As Font
Dim m_Value As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR

Const m_def_Value = False
Const m_def_BackColor = vbButtonFace
Const m_def_ForeColor = vbBlack

Event Click()
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseOut()
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
    Call UserControl_Click
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, x, Y)
End Sub

Private Sub p_Click()
    UserControl_Click
End Sub

Private Sub p_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub p_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call UserControl_MouseMove(Button, Shift, p.Left, p.Top)
End Sub

Private Sub p_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, x, Y)
End Sub

Private Sub Timer1_Timer()
    Dim pnt As POINT_API
    UserControl.ScaleMode = 3
    GetCursorPos pnt
    ScreenToClient UserControl.hwnd, pnt

    If pnt.x < UserControl.ScaleLeft Or _
       pnt.Y < UserControl.ScaleTop Or _
       pnt.x > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
       pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
       
        Timer1.Enabled = False
        RaiseEvent MouseOut
        define_pic
    End If
End Sub

Private Sub UserControl_Click()
    Dim rd As Object
    RaiseEvent Click
    
    For Each rd In UserControl.Parent
        If TypeOf rd Is XPRadioButton Then
            rd.Value = False
        End If
    Next rd
    Value = True
    define_pic
End Sub

Private Sub UserControl_Initialize()
    UserControl_Resize
    define_pic
    UserControl.BackColor = m_BackColor
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    If Enabled = False Then
        enabled_pic
    Else: define_pic
    End If
    If Enabled = True Then lbl.ForeColor = m_ForeColor Else lbl.ForeColor = RGB(161, 161, 146)
End Property

Private Sub UserControl_InitProperties()
    Caption = Ambient.DisplayName
    Enabled = True
    Value = False
    Set Font = UserControl.Ambient.Font
    BackColor = m_def_BackColor
    ForeColor = m_def_ForeColor
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Enabled = True Then
        If Value = True Then
            p.Picture = img(6).Picture
        ElseIf Value = False Then
            p.Picture = img(2).Picture
        End If
    End If
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Timer1.Enabled = True
    If x >= 0 And Y >= 0 And _
       x <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
        RaiseEvent MouseMove(Button, Shift, x, Y)
        If Button = vbLeftButton Then
            If Enabled = True Then
                If Value = True Then
                    p.Picture = img(6).Picture
                ElseIf Value = False Then
                    p.Picture = img(2).Picture
                End If
            End If
        Else
            If Enabled = True Then
                If Value = True Then
                    p.Picture = img(5).Picture
                ElseIf Value = False Then
                    p.Picture = img(1).Picture
                End If
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("Value", m_def_Value)
    Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
End Sub

Private Sub UserControl_Resize()
    UserControl.ScaleMode = 1
    p.Height = 240
    p.Width = 240
    p.Left = 0
    p.Top = (UserControl.Height - p.Height) \ 2
    lbl.Top = (UserControl.Height - lbl.Height) \ 2
    lbl.Left = 480
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Caption", lbl.Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
End Sub

Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal vNewValue As Boolean)
    m_Value = vNewValue
    define_pic
    PropertyChanged "Value"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
    lbl.Caption() = vNewCaption
    Call UserControl_Resize
    PropertyChanged "Caption"
End Property

Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    Set UserControl.Font = vNewFont
    Set lbl.Font = m_Font
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Private Function define_pic()
    If Enabled = True Then
        If Value = True Then
            p.Picture = img(4).Picture
        ElseIf Value = False Then
            p.Picture = img(0).Picture
        End If
    Else: enabled_pic
    End If
End Function

Private Function enabled_pic()
    If Value = True Then
        p.Picture = img(7).Picture
    ElseIf Value = False Then
        p.Picture = img(3).Picture
    End If
End Function

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    If Enabled = True Then lbl.ForeColor = m_ForeColor Else lbl.ForeColor = RGB(161, 161, 146)
End Property
