VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl XPTopButtons 
   CanGetFocus     =   0   'False
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1665
   FillStyle       =   0  'Solid
   MaskColor       =   &H00000000&
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   111
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   720
   End
   Begin PicClip.PictureClip pc 
      Index           =   0
      Left            =   360
      Top             =   0
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   4
      Picture         =   "XPTopButtons.ctx":0000
   End
   Begin PicClip.PictureClip pc 
      Index           =   1
      Left            =   360
      Top             =   360
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   4
      Picture         =   "XPTopButtons.ctx":14FE
   End
   Begin PicClip.PictureClip pc 
      Index           =   2
      Left            =   360
      Top             =   720
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   4
      Picture         =   "XPTopButtons.ctx":29FC
   End
   Begin PicClip.PictureClip pc 
      Index           =   3
      Left            =   360
      Top             =   1080
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   4
      Picture         =   "XPTopButtons.ctx":3EFA
   End
End
Attribute VB_Name = "XPTopButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Teh Ming Han (teh_minghan@hotmail.com)

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long

Dim s As Integer

Private Type POINT_API
    x As Long
    Y As Long
End Type

Public Enum typebutton
    CloseB = 0
    MaxB = 1
    MinB = 2
    RestoreB = 3
End Enum

Dim type_value As typebutton
Const type_def_value = typebutton.CloseB

Event Click()
Attribute Click.VB_UserMemId = -600
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Private Sub Timer1_Timer()
    Dim pnt As POINT_API
    GetCursorPos pnt
    ScreenToClient UserControl.hwnd, pnt

    If pnt.x < UserControl.ScaleLeft Or _
       pnt.Y < UserControl.ScaleTop Or _
       pnt.x > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
       pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        Timer1.Enabled = False
        RaiseEvent MouseOut
        If UserControl.Enabled = True Then UserControl.Picture = pc(s).GraphicCell(0)
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
Exit Sub
End Sub

Private Sub UserControl_Initialize()
    typevalue_pic
    UserControl.Picture = pc(s).GraphicCell(0)
    UserControl.Height = 21
    UserControl.Width = 21
End Sub

Private Sub UserControl_InitProperties()
    Value = type_def_value
    Enabled = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UserControl.Picture = pc(s).GraphicCell(2)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Timer1.Enabled = True
    If x >= 0 And Y >= 0 And _
       x <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
        RaiseEvent MouseMove(Button, Shift, x, Y)
        If Button = vbLeftButton Then
            UserControl.Picture = pc(s).GraphicCell(2)
        Else: UserControl.Picture = pc(s).GraphicCell(1)
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UserControl.Picture = pc(s).GraphicCell(0)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    type_value = PropBag.ReadProperty("Value", type_def_value)
    Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 315
    UserControl.Width = 315
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    If UserControl.Enabled = True Then
        UserControl.Picture = pc(s).GraphicCell(0)
    ElseIf UserControl.Enabled = False Then
        UserControl.Picture = pc(s).GraphicCell(3)
    End If
End Property

Private Sub UserControl_Show()
    typevalue_pic
End Sub

Private Sub UserControl_Terminate()
    typevalue_pic
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", type_value, type_def_value)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Value() As typebutton
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
    Value = type_value
End Property

Public Property Let Value(ByVal vNewValue As typebutton)
    type_value = vNewValue
    PropertyChanged "Value"
    typevalue_pic
End Property

Private Sub typevalue_pic()
    If Value = CloseB Then
        s = 0
    ElseIf Value = MaxB Then
        s = 1
    ElseIf Value = MinB Then
        s = 2
    ElseIf Value = RestoreB Then
        s = 3
    End If
    If Enabled = True Then UserControl.Picture = pc(s).GraphicCell(0) Else UserControl.Picture = pc(s).GraphicCell(3)
End Sub
