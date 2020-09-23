Attribute VB_Name = "modServerbot"
'This is the module for serverbot, serverbot will welcome users,
'and respond to users.

Sub ServerbotSpeak(Word As String, User As String)
Dim WordList As Integer
Dim WordItem() As String

For WordList = 1 To ServerbotDB.Count
WordItem = Split(ServerbotDB.Item(WordList), Chr(156))
    
    'Special messages
    If InStr(1, Word, "time") <> 0 Then 'Serverbot says time
        Call SendDataAll("chat", "Serverbot" & Chr(156) & "Verdana" & Chr(156) & "#000000" & Chr(156) & "Server time is " & CStr(Time) & ".")
        Exit Sub
    Else
    
    If InStr(1, LCase$(Word), WordItem(0)) <> 0 Then
       
    'Replaces USER placeholder with the user speaking
        If InStr(1, WordItem(1), "[USER]") > 0 Then
            WordItem(1) = Replace(WordItem(1), "[USER]", User)
        End If
        
        If InStr(1, WordItem(1), "/comma") > 0 Then
            WordItem(1) = Replace(WordItem(1), "/comma", ",")
        End If
    
            Call SendDataAll("chat", "Serverbot" & Chr(156) & "Verdana" & Chr(156) & "#000000" & Chr(156) & WordItem(1))
            Exit Sub
        End If
    End If

Next WordList

End Sub

Sub LoadServerbotDB()
Dim WordData As String

'Open the user database file
Open App.Path & "/Database/Serverbot.db" For Input As #1
        
    Do Until EOF(1) = True
        Input #1, WordData 'Get username from file
        ServerbotDB.Add WordData 'Add user to database
    DoEvents
    Loop
    
Close #1
End Sub

Sub SaveServerbotDB()
Dim WordLoop As Integer

'Save the user database file
Kill App.Path & "/Database/Serverbot.db" 'Delete existing database
Open App.Path & "/Database/Serverbot.db" For Output As #1

    For WordLoop = 1 To ServerbotDB.Count
        Print #1, ServerbotDB.Item(WordLoop)
    Next WordLoop

Close #1

End Sub

