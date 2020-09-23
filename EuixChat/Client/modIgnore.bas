Attribute VB_Name = "modIgnore"
Function CheckIgnore(strUsername As String) As Boolean
'Checks a user to see if he's being ignored.

Dim IgnoreLoop As Integer
For IgnoreLoop = 1 To IgnoreDB.Count
    
    If IgnoreDB.Item(IgnoreLoop) = LCase$(strUsername) Then
        CheckIgnore = True
        Exit Function
    End If
    
Next IgnoreLoop

'Couldnt find username in list
CheckIgnore = False
End Function

Sub LoadIgnoreList()
On Error Resume Next

Dim listUsername As String
Open App.Path & "\Ignore.dat" For Input As #1

    Do
        Input #1, listUsername
        IgnoreDB.Add listUsername
    DoEvents
    Loop Until EOF(1) = True

Close #1
End Sub

Sub SaveIgnoreList()

Dim IgnoreCount As Integer
Open App.Path & "\Ignore.dat" For Output As #1
    
    For IgnoreCount = 1 To IgnoreDB.Count
        Print #1, IgnoreDB.Item(IgnoreCount)
    Next IgnoreCount

Close #1
End Sub

Sub IgnoreUser(strUsername As String)
    
    If CheckIgnore(LCase$(strUsername)) = True Then
        MsgBox "This user is already being ignored!", vbInformation, "Ignore User"
        Exit Sub
    End If
    
    IgnoreDB.Add LCase$(strUsername)
    Call SendData("userlist", " ")
End Sub

Sub UnIgnoreUser(strUsername As String)
Dim IgnoreLoop As Integer

    If CheckIgnore(LCase$(strUsername)) = False Then
        MsgBox "This user is already unignored!", vbInformation, "UnIgnore User"
        Exit Sub
    End If
    
    For IgnoreLoop = 1 To IgnoreDB.Count
        
        If IgnoreDB.Item(IgnoreLoop) = LCase$(strUsername) Then
            IgnoreDB.Remove IgnoreLoop
            Call SendData("userlist", " ")
            Exit Sub
        End If
        
    Next IgnoreLoop
    
End Sub
