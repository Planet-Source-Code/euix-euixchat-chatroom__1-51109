Attribute VB_Name = "modSocket"
Sub EndServer()
On Error Resume Next
Dim Count As Integer

For Count = 0 To frmMain.lstUsers.ListCount
    frmMain.lstUsers.RemoveItem Count
Next Count

End Sub
