Attribute VB_Name = "modGlobals"
Global SocketCount As Integer
Global Connections As Integer

Global UserIndex As Integer 'User index on user list
Global UserCount As Long
Global User(1 To 100) As String 'User ID and Name

Global ServerbotOn As Boolean 'Is serverbot turned On or Off

Global ServerbotDB As New Collection 'Serverbots Database
Global UserDB  As New Collection 'User database

Sub AddLog(LogData As String, LogType As String)
LogType = LCase$(LogType)

'Color code the log data
If LogType = "user" Then
    frmMain.Log.Document.body.innerHTML = frmMain.Log.Document.body.innerHTML & "<font face=Verdana color=lightgreen><small>" & LogData & "</small></font><br>"
ElseIf LogType = "server" Then
    frmMain.Log.Document.body.innerHTML = frmMain.Log.Document.body.innerHTML & "<font face=Verdana color=lightblue><small>" & LogData & "</small></font><br>"
ElseIf LogType = "connection" Then
    frmMain.Log.Document.body.innerHTML = frmMain.Log.Document.body.innerHTML & "<font face=Verdana color=maroon><small>" & LogData & "</small></font><br>"
End If

frmMain.Log.Document.body.scrolltop = CLng(Len(frmMain.Log.Document.body.innerHTML)) * 1000
End Sub
