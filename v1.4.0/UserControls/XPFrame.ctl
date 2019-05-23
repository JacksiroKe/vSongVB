VERSION 5.00
Begin VB.UserControl XPFrame 
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ControlContainer=   -1  'True
   ForeColor       =   &H00D54600&
   ScaleHeight     =   1545
   ScaleWidth      =   3795
   ToolboxBitmap   =   "XPFrame.ctx":0000
End
Attribute VB_Name = "XPFrame"
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

Option Explicit
Private m_Caption As String
Private m_BorderColor As OLE_COLOR
Private m_TextColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_RightToLeft As Boolean
Private m_Align As enAlign
Private m_Picture As StdPicture

Private GradientBackground As Boolean
Private Const m_TopColor As Long = &HBFD0D0
Private Const m_BottomColor As Long = &HD54600

Public Enum enAlign
    [Left Align] = 0
    [Right Align] = 1
    [Center Align] = 2
End Enum

Private Sub UserControl_InitProperties()
    m_Caption = UserControl.Name
    m_RightToLeft = False
    m_BorderColor = &HBFD0D0
    m_TextColor = &HD54600
    m_BackColor = &H8000000F
    m_Align = [Left Align]
End Sub

Private Sub UserControl_Paint()
    Redraw
End Sub

Private Sub Redraw()
Static SubIsDoing As Boolean
Dim Wid As Long, Hei As Long
Dim LineWidth As Long, LineHeight As Long
Dim ChamWid As Long, ChamHei As Long
Dim Px1 As Long, Px2 As Long
Dim Py1 As Long, Py2 As Long
Dim StartY As Long
Dim lngTextX As Long, lngTextWidth As Long

If SubIsDoing Then Exit Sub
SubIsDoing = True
    'Background Picture
    UserControl.BackColor = m_BackColor
    Set UserControl.Picture = m_Picture
    'Clear control for redrawing
    Cls
    'Values & Variables ...
    Px1 = TwipX(1): Px2 = 2 * Px1 'Px1=1 Pixel in twips
    Py1 = TwipY(1): Py2 = 2 * Py1
    Wid = ScaleWidth ' = Usercontrol.Width
    Hei = ScaleHeight ' = Usercontrol.Height
    ChamWid = TwipX(2)
    ChamHei = TwipX(2)
    LineWidth = Wid - ChamWid
    LineHeight = Hei - ChamHei
    StartY = TextHeight(m_Caption) / 2
    lngTextWidth = TextWidth(m_Caption)
    'Calculate the left pixel of text rect.
    Select Case m_Align
    Case [Left Align]
        lngTextX = ChamWid + (5 * Px1)
    Case [Right Align]
        lngTextX = Wid - lngTextWidth - (ChamWid + (5 * Px1))
    Case [Center Align]
        lngTextX = (Wid - lngTextWidth) \ 2
    End Select
    
    ForeColor = m_BorderColor
    '###### Draw Vertical & Horizontal Lines #######
    If Len(m_Caption) = 0 Then
        Line (ChamWid, StartY)-(LineWidth, StartY) ' ----- Upper line
    Else
        Line (ChamWid, StartY)-(lngTextX - Px1, StartY) '--- upperline (left of text)
        Line (lngTextX + lngTextWidth + Px1, StartY)-(LineWidth, StartY) '---- Upper line (right of text)
    End If
    Line (Wid - Px1, ChamHei + StartY)-(Wid - Px1, LineHeight) '||||| Right line
    Line (0, ChamHei + StartY)-(0, LineHeight) '|||| Left Line
    Line (ChamWid, Hei - Py1)-(LineWidth, Hei - Py1) ' ----- Botton Line
    
    '######### Draw Corner Pixels ############
    ' Top Left Corner
    PSet (Px2, Py1 + StartY)
    PSet (Py1, Py2 + StartY)
    PSet (Py1, Py1 + StartY)
    ' Top Right Corner
    PSet (LineWidth - Px1, Px1 + StartY)
    PSet (LineWidth, Px1 + StartY)
    PSet (LineWidth, Px2 + StartY)
    ' Bottom Left Corner
    PSet (Px1, LineHeight - Px1)
    PSet (Px2, LineHeight)
    PSet (Px1, LineHeight)
    ' Bottom Right Corner
    PSet (LineWidth, LineHeight)
    PSet (LineWidth - Px1, LineHeight)
    PSet (LineWidth, LineHeight - Px1)
    
    '############# Draw Text! ###############
    ForeColor = m_TextColor 'Set color of text
    If Len(m_Caption) <> 0 Then
        'Write text
        CurrentX = lngTextX
        CurrentY = 0
        Print m_Caption
    End If
    SubIsDoing = False
End Sub

Private Function TwipX(lngPixel As Long) As Long
    'Convert Pixel to Twips
    TwipX = ScaleX(lngPixel, vbPixels, vbTwips)
End Function
Private Function TwipY(lngPixel As Long) As Long
    'Convert Pixel to Twips
    TwipY = ScaleY(lngPixel, vbPixels, vbTwips)
End Function


Property Let RightToLeft(NewVal As Boolean)
On Error Resume Next
    m_RightToLeft = NewVal
    UserControl.RightToLeft = NewVal
    PropertyChanged "RightToLeft"
    Redraw
End Property
Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property

Property Get Align() As enAlign
    Align = m_Align
End Property
Property Let Align(NewVal As enAlign)
    m_Align = NewVal
    PropertyChanged "Align"
    Redraw
End Property

Property Let Caption(NewVal As String)
    m_Caption = NewVal
    PropertyChanged "Caption"
    Redraw
End Property
Property Get Caption() As String
    Caption = m_Caption
End Property

Property Set Font(NewVal As Font)
    Set UserControl.Font = NewVal
    PropertyChanged "Font"
    Redraw
End Property
Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Property Let FontItalic(NewVal As Boolean)
    UserControl.FontItalic = NewVal
    PropertyChanged "FontItalic"
    Redraw
End Property
Property Get FontItalic() As Boolean
    FontItalic = UserControl.FontItalic
End Property

Property Let FontBold(NewVal As Boolean)
    UserControl.FontBold = NewVal
    PropertyChanged "FontBold"
    Redraw
End Property
Property Get FontBold() As Boolean
    FontBold = UserControl.FontBold
End Property

Property Let FontName(NewVal As String)
    UserControl.FontName = NewVal
    PropertyChanged "FontName"
    Redraw
End Property
Property Get FontName() As String
    FontName = UserControl.FontName
End Property

Property Let FontSize(NewVal As Long)
    UserControl.FontSize = NewVal
    PropertyChanged "FontSize"
    Redraw
End Property
Property Get FontSize() As Long
    FontSize = UserControl.FontSize
End Property

Property Let BorderColor(NewVal As OLE_COLOR)
    m_BorderColor = NewVal
    PropertyChanged "BorderColor"
    Redraw
End Property
Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Property Let TextColor(NewVal As OLE_COLOR)
    m_TextColor = NewVal
    PropertyChanged "TextColor"
    Redraw
End Property
Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

Property Let BackColor(NewVal As OLE_COLOR)
    m_BackColor = NewVal
    PropertyChanged "BackColor"
    Redraw
End Property
Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property
Property Set Picture(NewPic As StdPicture)
    Set m_Picture = NewPic
    PropertyChanged "Picture"
    Redraw
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    PropBag.WriteProperty "Caption", m_Caption, vbNullString
    PropBag.WriteProperty "RightToLeft", m_RightToLeft, False
    PropBag.WriteProperty "Font", UserControl.Font
    PropBag.WriteProperty "BorderColor", m_BorderColor, &HBFD0D0
    PropBag.WriteProperty "TextColor", m_TextColor, &HD54600
    PropBag.WriteProperty "BackColor", m_BackColor, &H8000000F
    PropBag.WriteProperty "FontName", UserControl.FontName
    PropBag.WriteProperty "FontSize", UserControl.FontSize
    PropBag.WriteProperty "FontBold", UserControl.FontBold
    PropBag.WriteProperty "FontItalic", UserControl.FontItalic
    PropBag.WriteProperty "Align", m_Align, 0
    PropBag.WriteProperty "Picture", m_Picture, Nothing
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    m_Caption = PropBag.ReadProperty("Caption", vbNullString)
    m_RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    Set UserControl.Font = PropBag.ReadProperty("Font")
    m_BorderColor = PropBag.ReadProperty("BorderColor", &HBFD0D0)
    m_TextColor = PropBag.ReadProperty("TextColor", &HD54600)
    m_BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.FontName = PropBag.ReadProperty("FontName")
    UserControl.FontSize = PropBag.ReadProperty("FontSize")
    UserControl.FontBold = PropBag.ReadProperty("FontBold")
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic")
    m_Align = PropBag.ReadProperty("Align", 0)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

