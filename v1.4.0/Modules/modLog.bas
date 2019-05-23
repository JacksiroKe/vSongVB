Attribute VB_Name = "modLog"
Option Explicit

Sub Log_Write(LogMsg As String)
   Debug.Print LogMsg
End Sub

Function LOG_FILENAME() As String
   LOG_FILENAME = App.Path & "\js_error_log.txt"
End Function


