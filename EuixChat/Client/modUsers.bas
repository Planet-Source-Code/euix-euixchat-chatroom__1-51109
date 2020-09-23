Attribute VB_Name = "modUsers"
Sub ChangeIcon(Username As String, IconID As Integer)
'Changes the icon of the user based on status
    
    Dim UserLoop As Integer
    For UserLoop = 0 To frmMain.lstUsers.Nodes.Count
        
        If frmMain.lstUsers.Nodes.Item(UserLoop).Text = Username Then
            frmMain.lstUsers.Nodes.Item(UserLoop).Image = IconID
            Exit Sub
        End If
        
    Next UserLoop
    
End Sub
