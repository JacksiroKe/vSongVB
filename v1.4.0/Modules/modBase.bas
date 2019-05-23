Attribute VB_Name = "modBase"
Option Explicit

Sub Main()
    'Program_Start
End Sub

Sub DoNothing()
    
End Sub

Sub Log_Write(LogMsg As String)
   Debug.Print LogMsg
End Sub

Function LOG_FILENAME() As String
   LOG_FILENAME = App.Path & "\vsb_error_log.txt"
End Function

Function File_Exists(FileName As String) _
   As Boolean
   
   On Error GoTo NotFound:
   
   Dim i As Long
   i = GetAttr(FileName)
   
   If i And vbDirectory Then
      Exit Function
   End If
   
   File_Exists = True

Exit Function
NotFound:
   File_Exists = False
End Function

Function MyFontType(fontint) As String
    If fontint = 1 Then
        MyFontType = "ALGERIAN"
    ElseIf fontint = 2 Then
        MyFontType = "Arial"
    ElseIf fontint = 3 Then
        MyFontType = "Calibri"
    ElseIf fontint = 4 Then
        MyFontType = "Century Gothic"
    ElseIf fontint = 5 Then
        MyFontType = "Comic Sans MS"
    ElseIf fontint = 6 Then
        MyFontType = "Corbel"
    ElseIf fontint = 7 Then
        MyFontType = "Courier New"
    ElseIf fontint = 8 Then
        MyFontType = "Old English Text MT"
    ElseIf fontint = 9 Then
        MyFontType = "Palatino Linotype"
    ElseIf fontint = 10 Then
        MyFontType = "Tahoma"
    ElseIf fontint = 11 Then
        MyFontType = "Tempus Sans ITC"
    ElseIf fontint = 12 Then
        MyFontType = "Times New Roman"
    ElseIf fontint = 13 Then
        MyFontType = "Trebuchet MS"
    ElseIf fontint = 14 Then
        MyFontType = "Verdana"
    End If
End Function

Function MyFontName(fontname) As Integer
    If fontname = "ALGERIAN" Then
        MyFontName = 1
    ElseIf fontname = "Arial" Then
        MyFontName = 2
    ElseIf fontname = "Calibri" Then
        MyFontName = 3
    ElseIf fontname = "Century Gothic" Then
        MyFontName = 4
    ElseIf fontname = "Comic Sans MS" Then
        MyFontName = 5
    ElseIf fontname = "Corbel" Then
        MyFontName = 6
    ElseIf fontname = "Courier New" Then
        MyFontName = 7
    ElseIf fontname = "Old English Text MT" Then
        MyFontName = 8
    ElseIf fontname = "Palatino Linotype" Then
        MyFontName = 9
    ElseIf fontname = "Tahoma" Then
        MyFontName = 10
    ElseIf fontname = "Tempus Sans ITC" Then
        MyFontName = 11
    ElseIf fontname = "Times New Roman" Then
        MyFontName = 12
    ElseIf fontname = "Trebuchet MS" Then
        MyFontName = 13
    ElseIf fontname = "Verdana" Then
        MyFontName = 14
    End If
End Function

Function Convert_Text_Min(my_text) As String
    my_text = Replace(my_text, "+", Chr$(34))
    my_text = Replace(my_text, "^", "'")
    Convert_Text_Min = my_text
End Function

Function Convert_Text_Max(my_text) As String
    my_text = Replace(my_text, "+", Chr$(34))
    my_text = Replace(my_text, "^", "'")
    my_text = Replace(my_text, "$", vbNewLine)
    Convert_Text_Max = my_text
End Function

Function Convert_Text_Rvs(my_text) As String
    my_text = Replace(my_text, Chr$(34), "+")
    my_text = Replace(my_text, "'", "^")
    my_text = Replace(my_text, vbNewLine, " $ ")
    my_text = Replace(my_text, "  ", " ")
    Convert_Text_Rvs = my_text
End Function

