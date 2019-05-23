VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl XPHelp 
   CanGetFocus     =   0   'False
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   ScaleHeight     =   435
   ScaleWidth      =   1380
   ToolboxBitmap   =   "XPHelp.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   0
   End
   Begin PicClip.PictureClip pc1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1667
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   3
      Picture         =   "XPHelp.ctx":0312
   End
End
Attribute VB_Name = "XPHelp"
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
        If UserControl.Enabled = True Then UserControl.Picture = pc1.GraphicCell(0)
    End If
End Sub

Private Sub UserControl_Click()
    UserControl.Picture = pc1.GraphicCell(0)
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    UserControl.Picture = pc1.GraphicCell(0)
    UserControl.Height = 21
    UserControl.Width = 21
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UserControl.Picture = pc1.GraphicCell(2)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Timer1.Enabled = True
    If x >= 0 And Y >= 0 And _
       x <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
       RaiseEvent MouseMove(Button, Shift, x, Y)
       If Button = vbLeftButton Then
           UserControl.Picture = pc1.GraphicCell(2)
       Else:     UserControl.Picture = pc1.GraphicCell(1)
       End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 315
    UserControl.Width = 315
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub
