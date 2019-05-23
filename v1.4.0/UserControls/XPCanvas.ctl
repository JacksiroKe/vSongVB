VERSION 5.00
Begin VB.UserControl XPCanvas 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   ControlContainer=   -1  'True
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   ToolboxBitmap   =   "XPCanvas.ctx":0000
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1680
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1080
      Top             =   3240
   End
   Begin VB.PictureBox pictop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   410
      TabIndex        =   3
      Top             =   0
      Width           =   6150
      Begin VB.Label lblcaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Windows XP Controls"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   390
         TabIndex        =   4
         Top             =   120
         Width           =   1965
      End
      Begin VB.Image imgicon 
         Height          =   240
         Left            =   90
         Picture         =   "XPCanvas.ctx":0312
         Stretch         =   -1  'True
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgresize 
         Height          =   135
         Index           =   6
         Left            =   6000
         MousePointer    =   6  'Size NE SW
         Top             =   0
         Width           =   135
      End
      Begin VB.Image imgresize 
         Height          =   135
         Index           =   7
         Left            =   0
         MousePointer    =   8  'Size NW SE
         Top             =   0
         Width           =   135
      End
      Begin VB.Image imgresize 
         Height          =   135
         Index           =   2
         Left            =   120
         MousePointer    =   7  'Size N S
         Top             =   0
         Width           =   5895
      End
      Begin VB.Label lblshadow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Windows XP Controls"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   480
         TabIndex        =   5
         Top             =   120
         Width           =   1965
      End
   End
   Begin VB.PictureBox picbottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   410
      TabIndex        =   2
      Top             =   4845
      Width           =   6150
      Begin VB.Image imgresize 
         Height          =   135
         Index           =   5
         Left            =   0
         MousePointer    =   6  'Size NE SW
         Top             =   0
         Width           =   135
      End
      Begin VB.Image imgresize 
         Height          =   135
         Index           =   4
         Left            =   6000
         MousePointer    =   8  'Size NW SE
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.PictureBox picleft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4410
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   294
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      Top             =   435
      Width           =   135
   End
   Begin VB.PictureBox picright 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4410
      Left            =   6015
      MousePointer    =   9  'Size W E
      ScaleHeight     =   294
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   0
      Top             =   435
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   720
      Picture         =   "XPCanvas.ctx":089C
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgtop 
      Height          =   435
      Left            =   1320
      Top             =   960
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image imgleft 
      Height          =   465
      Left            =   600
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgright 
      Height          =   465
      Left            =   4800
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgbottom 
      Height          =   75
      Left            =   2280
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgtopmax 
      Height          =   435
      Left            =   1320
      Top             =   1560
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image imgleft1 
      Height          =   465
      Left            =   840
      Picture         =   "XPCanvas.ctx":0E26
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgtop1 
      Height          =   435
      Left            =   2400
      Picture         =   "XPCanvas.ctx":1360
      Top             =   960
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image imgright1 
      Height          =   465
      Left            =   4560
      Picture         =   "XPCanvas.ctx":1F56
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgtopmax1 
      Height          =   435
      Left            =   2400
      Picture         =   "XPCanvas.ctx":2490
      Top             =   1560
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image imgbottom1 
      Height          =   75
      Left            =   2280
      Picture         =   "XPCanvas.ctx":3086
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgtop2 
      Height          =   435
      Left            =   3480
      Picture         =   "XPCanvas.ctx":35CC
      Top             =   960
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image imgtopmax2 
      Height          =   435
      Left            =   3480
      Picture         =   "XPCanvas.ctx":41C2
      Top             =   1560
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image imgleft2 
      Height          =   465
      Left            =   1080
      Picture         =   "XPCanvas.ctx":4DB8
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgbottom2 
      Height          =   75
      Left            =   2280
      Picture         =   "XPCanvas.ctx":52F2
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgright2 
      Height          =   465
      Left            =   4320
      Picture         =   "XPCanvas.ctx":5838
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "XPCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/----------------------------------------------------\
'/Descriptions: Creation of Windows XP windows and    \
'/              controls in Visual Basic              \
'/Created by: Teh Ming Han (teh_minghan@hotmail.com)  \
'/Special thanks: Chris Yates (cyates@neo.rr.com)     \
'/                for trans_colour module             \
'/                                                    \
'/REMEMBER TO VOTE!                                   \
'/                                                    \
'/If you use this code in your program please give me \
'/credit and e-mail me (teh_minghan@hotmail.com) and  \
'/tell me about your program.                         \
'/------Hope you find it useful!---------2001---------\
'/----------------------------------------------------\

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Private Type PointAPI
    x As Long
    Y As Long
End Type

Dim oldcp As PointAPI
Dim newcp As PointAPI
Dim m_Icon As Picture
Dim FixedSingle As Boolean

Event Click()
Attribute Click.VB_UserMemId = -600
Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Event Resize()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Private Sub lblcaption_DblClick()
    If Fixed_Single = False Then If UserControl.Parent.WindowState = 0 Then UserControl.Parent.WindowState = 2 Else UserControl.Parent.WindowState = 0
End Sub

Private Sub pictop_DblClick()
    If Fixed_Single = False Then If UserControl.Parent.WindowState = 0 Then UserControl.Parent.WindowState = 2 Else UserControl.Parent.WindowState = 0
End Sub

Private Sub pictop_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If UserControl.Parent.WindowState = 0 Then
        ReleaseCapture
        SendMessage UserControl.Parent.hwnd, &HA1, 2, 0&
    End If
    DoEvents
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If Screen.ActiveForm.hwnd <> UserControl.Parent.hwnd Then
        lost_f
        UserControl_Resize
        Timer1.Enabled = False
    End If
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    If Screen.ActiveForm.hwnd = UserControl.Parent.hwnd Then
        got_f
        UserControl_Resize
        Timer2.Enabled = False
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
    Timer1.Enabled = False
    got_f
    UserControl_Resize
End Sub

Private Sub UserControl_ExitFocus()
    Timer1.Enabled = True
    Timer2.Enabled = True
End Sub

Private Sub UserControl_GotFocus()
    Timer1.Enabled = False
    Timer2.Enabled = False
    got_f
    UserControl_Resize
End Sub

Private Sub UserControl_Initialize()
    got_f
    UserControl_Resize
End Sub

Private Sub imgresize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Fixed_Single = False Then GetCursorPos oldcp
End Sub

Private Sub imgresize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Fixed_Single = False Then
        GetCursorPos newcp
        ResizeForm UserControl.Parent, oldcp, newcp, Index
    End If
End Sub

Private Sub UserControl_InitProperties()
    Caption = Ambient.DisplayName
    Fixed_Single = False
    Set Icon = Image1.Picture
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
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Set Icon = PropBag.ReadProperty("Icon", Image1.Picture)
    Fixed_Single = PropBag.ReadProperty("Fixed_Single", False)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.Parent.ScaleMode = 3
    UserControl.AutoRedraw = True
    UserControl.BackColor = RGB(236, 233, 216)
    UserControl.ScaleMode = 3
    'Ease drawings so transparancy can function correctly
    UserControl.Cls
    'Sets fixed data
    picleft.Width = 5
    picright.Width = 5
    pictop.Height = 29
    picbottom.Height = 5
    'Algin to fixed other data
    picleft.Align = 3
    picright.Align = 4
    pictop.Align = 1
    picbottom.Align = 2
    'Draws left and right
    picleft.PaintPicture imgleft.Picture, 0, 0, picleft.Width, picleft.Height, 0, 0, imgleft.Width, imgleft.Height
    picright.PaintPicture imgright.Picture, 0, 0, picright.Width, picright.Height, 0, 0, imgright.Width, imgright.Height

    If UserControl.Parent.WindowState = 0 Then
        'Normal with round edge
        pictop.PaintPicture imgtop.Picture, 0, 0, 5, pictop.Height, 0, 0, 5, imgright.Height
        pictop.PaintPicture imgtop.Picture, 5, 0, pictop.Width - 10, pictop.Height, 5, 0, imgtop.Width - 10, imgtop.Height
        pictop.PaintPicture imgtop.Picture, pictop.Width - 5, 0, 5, pictop.Height, imgtop.Width - 5, 0, 5, imgtop.Height
    ElseIf UserControl.Parent.WindowState = 2 Then
        'Maximized with sharp edge (different picture)
        pictop.PaintPicture imgtopmax.Picture, 0, 0, 5, pictop.Height, 0, 0, 5, imgright.Height
        pictop.PaintPicture imgtopmax.Picture, 5, 0, pictop.Width - 10, pictop.Height, 5, 0, imgtop.Width - 10, imgtop.Height
        pictop.PaintPicture imgtopmax.Picture, pictop.Width - 5, 0, 5, pictop.Height, imgtop.Width - 5, 0, 5, imgtop.Height
    End If
    'Bottom remains the same
    picbottom.PaintPicture imgbottom.Picture, 0, 0, picbottom.Width, picbottom.Height, 0, 0, imgbottom.Width, imgbottom.Height
    lblcaption.Top = 6
    lblcaption.Left = 26
    lblshadow.Top = 8
    lblshadow.Left = 28
    
    imgresize(6).Top = 0
    imgresize(6).Left = UserControl.Parent.ScaleWidth - 9
    imgresize(4).Top = picbottom.ScaleHeight - 9
    imgresize(4).Left = picbottom.ScaleWidth - 9
    imgresize(2).Width = pictop.ScaleWidth - 18
    imgresize(2).Left = 9

    If Fixed_Single = False Then
        If UserControl.Parent.WindowState = 0 Then
            picleft.MousePointer = vbSizeWE
            picright.MousePointer = vbSizeWE
            picbottom.MousePointer = vbSizeNS
            imgresize(6).MousePointer = vbSizeNESW
            imgresize(2).MousePointer = vbSizeNS
            imgresize(4).MousePointer = vbSizeNWSE
            imgresize(5).MousePointer = vbSizeNESW
            imgresize(7).MousePointer = vbSizeNWSE
        ElseIf UserControl.Parent.WindowState = 2 Then
            picleft.MousePointer = vbDefault
            picright.MousePointer = vbDefault
            picbottom.MousePointer = vbDefault
            imgresize(6).MousePointer = vbDefault
            imgresize(2).MousePointer = vbDefault
            imgresize(4).MousePointer = vbDefault
            imgresize(5).MousePointer = vbDefault
            imgresize(7).MousePointer = vbDefault
        End If
    ElseIf Fixed_Single = True Then
        picleft.MousePointer = vbDefault
        picright.MousePointer = vbDefault
        picbottom.MousePointer = vbDefault
        imgresize(6).MousePointer = vbDefault
        imgresize(2).MousePointer = vbDefault
        imgresize(4).MousePointer = vbDefault
        imgresize(5).MousePointer = vbDefault
        imgresize(7).MousePointer = vbDefault
    End If
    
    DoEvents
    RaiseEvent Resize
End Sub

Private Sub lost_f()
    imgleft.Picture = imgleft2.Picture
    imgright.Picture = imgright2.Picture
    imgtop.Picture = imgtop2.Picture
    imgtopmax.Picture = imgtopmax2.Picture
    imgbottom.Picture = imgbottom2.Picture
End Sub

Private Sub got_f()
    imgleft.Picture = imgleft1.Picture
    imgright.Picture = imgright1.Picture
    imgtop.Picture = imgtop1.Picture
    imgtopmax.Picture = imgtopmax1.Picture
    imgbottom.Picture = imgbottom1.Picture
End Sub

Private Sub lblcaption_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If UserControl.Parent.WindowState = 0 Then
        ReleaseCapture
        SendMessage UserControl.Parent.hwnd, &HA1, 2, 0&
    End If
End Sub

Private Sub picbottom_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Fixed_Single = False Then GetCursorPos oldcp
End Sub

Private Sub picbottom_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Fixed_Single = False Then
        GetCursorPos newcp
        ResizeForm UserControl.Parent, oldcp, newcp, 3
    End If
End Sub

Private Sub picleft_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Fixed_Single = False Then GetCursorPos oldcp
End Sub

Private Sub picleft_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Fixed_Single = False Then
        GetCursorPos newcp
        ResizeForm UserControl.Parent, oldcp, newcp, 0
    End If
End Sub

Private Sub picright_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Fixed_Single = False Then GetCursorPos oldcp
End Sub

Private Sub ResizeForm(frm As Form, oldcp As PointAPI, newcp As PointAPI, ResizeMode As Integer)
    On Error Resume Next
' Oldcp: Old cursor position (MouseDown)
' Newcp: New cursor position (MouseUp)
' ResizeMode:   0 - Left side
'               1 - Right side
'               2 - Top side
'               3 - Bottom side
'               4 - Bottom right corner
'               5 - Bottom left corner
'               6 - Top right corner
'               7 - Top left corner
    Dim DifferenceX
    Dim DifferenceY
    DifferenceX = (newcp.x - oldcp.x) * Screen.TwipsPerPixelX
    DifferenceY = (newcp.Y - oldcp.Y) * Screen.TwipsPerPixelY
    
    Select Case ResizeMode
    Case 0
        frm.Move frm.Left + DifferenceX, frm.Top, frm.Width - DifferenceX, frm.Height
    Case 1
        frm.Move frm.Left, frm.Top, frm.Width + DifferenceX, frm.Height
    Case 2
        frm.Move frm.Left, frm.Top + DifferenceY, frm.Width, frm.Height - DifferenceY
    Case 3
        frm.Move frm.Left, frm.Top, frm.Width, frm.Height + DifferenceY
    Case 4
        frm.Move frm.Left, frm.Top, frm.Width + DifferenceX, frm.Height + DifferenceY
    Case 5
        frm.Move frm.Left + DifferenceX, frm.Top, frm.Width - DifferenceX, frm.Height + DifferenceY
    Case 6
        frm.Move frm.Left, frm.Top + DifferenceY, frm.Width + DifferenceX, frm.Height - DifferenceY
    Case 7
        frm.Move frm.Left + DifferenceX, frm.Top + DifferenceY, frm.Width - DifferenceX, frm.Height - DifferenceY
    End Select
    
End Sub

Private Sub picright_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Fixed_Single = False Then
        GetCursorPos newcp
        ResizeForm UserControl.Parent, oldcp, newcp, 1
    End If
End Sub

Public Sub make_trans(frm As Form)
    frm.Cls
    frm.ScaleMode = 3
    
    If frm.WindowState = 0 Then
        frm.PaintPicture imgtop.Picture, 0, 0, 5, pictop.Height, 0, 0, 5, imgtop.Height
        frm.PaintPicture imgtop.Picture, pictop.Width - 5, 0, 5, pictop.Height, imgtop.Width - 5, 0, 5, imgtop.Height
    End If
    
    AutoFormShape frm, RGB(255, 0, 255)
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = lblcaption.Caption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
    lblcaption.Caption() = vNewCaption
    lblshadow.Caption() = vNewCaption
    UserControl.Parent.Caption = vNewCaption
    Call UserControl_Resize
    PropertyChanged "Caption"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lblcaption.Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Icon", m_Icon, Image1.Picture)
    Call PropBag.WriteProperty("Fixed_Single", FixedSingle, False)
End Sub

Public Property Get Icon() As Picture
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set m_Icon = New_Icon
    imgicon.Picture = New_Icon
    PropertyChanged "Icon"
End Property

Public Sub SetFocus()
    Timer1.Enabled = False
    Timer2.Enabled = False
    got_f
    UserControl_Resize
End Sub

Public Property Get Fixed_Single() As Boolean
    Fixed_Single = FixedSingle
End Property

Public Property Let Fixed_Single(ByVal vNewValue As Boolean)
    FixedSingle = vNewValue
    UserControl_Resize
End Property

Public Sub AlwaysOnTop(FrmID As Form, OnTop As Boolean)
    ' This sub uses an argument to determine whether
    ' to make the specified form always on top or not
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

    If OnTop Then
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub
