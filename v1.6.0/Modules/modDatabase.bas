Attribute VB_Name = "modDatabase"
Option Explicit

Public Const DB_LOAD_FAIL As Long = -1
Public Const DB_LOAD_SUCCESS As Long = 0
Dim con As New ADODB.Connection, Rs As New ADODB.Recordset
Dim saved As Boolean, ddate As Date, ddiff As Integer

Function Db_FileName() As String
   Db_FileName = App.Path & "\Tools\vSongBook.mdb"
End Function

Sub Database_Open()
 On Error GoTo ErrorHandler
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + Db_FileName + ";"
    con.Open
    Exit Sub
ErrorHandler:
MsgBox "Unable to open database files. Either you obtain a fresh copy of vSongBook or contact the developer", vbExclamation, "vSongBook unexpected error"
End Sub

Function Database_Connect() As Long
   Log_Write "Begin Load"
   If Not File_Exists(Db_FileName) Then
      Log_Write "Connect Database, no file: " & Db_FileName
      Database_Connect = DB_LOAD_FAIL
   Else
      Database_Open
      Database_Connect = DB_LOAD_SUCCESS
   End If
   Log_Write "End Load"
End Function

Sub Database_Close()
    con.Close
        
End Sub
