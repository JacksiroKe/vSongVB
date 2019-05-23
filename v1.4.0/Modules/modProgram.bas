Attribute VB_Name = "modProgram"
Option Explicit

Private Const AUTHOR As String = "Jackson Siro"

Sub Program_Start()
   DoEvents
   Log_Write "start"
   frmAaa.Show: DoEvents
   
   If Database_Connect = DB_LOAD_FAIL Then
      Log_Write "data error"
      MsgBox "Unable to connect to the database. Either you obtain a fresh copy of vSongBook or contact the developer", vbExclamation, "vSongBook unexpected error"
      Program_End
   Else
      Log_Write "loaded"
      'frmAaa.Status_Update "listing"
      'frmCcHome.Show
      ' frmSearch.Show
      'frmAaa.Hide
      'Memory_Clear
   End If
   
End Sub

Sub Program_End()
   Log_Write "end"
   Unload frmAa1
   Unload frmAaa
   Unload frmCcHome
   Unload frmCcProject
   Unload frmDdHelp
   Unload frmDdInfo
   Unload frmEeOptions
   Unload frmFfEditSong
   Unload frmFfNewsong
   Database_Close
   End
End Sub

Function Program_Title() As String
   Program_Title = "vSongBook4PC"
End Function

Function Program_Version() As String
   Program_Version = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Function

Function Program_Author() As String
   Program_Author = AUTHOR
End Function

