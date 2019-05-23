Attribute VB_Name = "modTransparent"
Option Explicit

Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function TransparentForm Lib "Res\SkinForm.dll" (ByVal sPathFile As String) As Long

Public Function FileExist(ByVal sFileName As String) As Boolean
    Dim HFile As Long
        
    On Error GoTo ErrFileExit
    
    HFile = FreeFile()
    Open sFileName For Input As #HFile
    Close #HFile
    
    FileExist = True
    Exit Function
ErrFileExit:
    FileExist = False
End Function



