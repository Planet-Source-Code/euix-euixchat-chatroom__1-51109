Attribute VB_Name = "modData"
Sub SpliceData(Index As Integer, Data As String)
Dim DataLoop As Integer 'Loop of data received
Dim DataCarriage() As String 'End of Data
Dim DataArray() As String 'Split the data into parts for parsing

DataCarriage = Split(Data, Chr(170)) 'Split data by ª

'Loop through all the data
For DataLoop = 0 To UBound(DataCarriage)
If DataCarriage(DataLoop) = "" Then Exit Sub

    'Get the controller and messages
    DataArray = Split(DataCarriage(DataLoop), Chr(171))
    Call ParseData(Index, DataArray(0), DataArray(1)) 'Parse data
Next DataLoop

End Sub

Sub ParseData(Index As Integer, Controller As String, Message As String)
Dim UserData() As String

'Check the controller, than execute it's commands
If Controller = "vercheck" Then 'Version Check
    Dim VerInfo() As String
    VerInfo = Split(Message, Chr(156))
        
        If VerInfo(0) = CStr(App.Major) Then
        
            If VerInfo(1) = CStr(App.Minor) Then
                Exit Sub
            Else
                frmMain.sckServer(Index).Close
                Call AddLog("Socket #" & CStr(Index) & " disconnected because of invalid VerID.", "user")
            End If
            
        Else
            frmMain.sckServer(Index).Close
        End If

'ABOVE: If Version check is incorrect, then the connected user
'gets disconnected

ElseIf Controller = "login" Then
UserData = Split(Message, Chr(156))

    If CheckUser(LCase$(UserData(0))) = True Then
        
        'Check to see if the user is already logged in.
        If CheckLogin(LCase$(UserData(0))) = True Then
            Call SendData(Index, "feedback", "loggedin") 'User logged in, send him a message
        Else 'User wasn't logged in
            
            'Check the users password
            If CheckPassword(LCase$(UserData(0)), LCase$(UserData(1))) = True Then
                
                'Check the status to determine if he is an admin or what
                If CheckStatus(LCase$(UserData(0))) = "normal" Then
                    User(Index) = UserData(0) 'Set User ID and Name
                    frmMain.lstUsers.AddItem User(Index)
                    Call SendData(Index, "feedback", "login")
                    Call SendUserList
                    Call SendDataAll("serverchat", User(Index) & " has logged in.") 'Send notification to all sockets
                    
                ElseIf CheckStatus(LCase$(UserData(0))) = "admin" Then
                    User(Index) = UserData(0) 'Set User ID and Name
                    frmMain.lstUsers.AddItem User(Index)
                    Call SendData(Index, "feedback", "adminlogin")
                    Call SendUserList
                    Call SendDataAll("serverchat", User(Index) & " has logged in.") 'Send notification to all sockets
                
                ElseIf CheckStatus(LCase$(UserData(0))) = "specialadmin" Then
                    User(Index) = UserData(0) 'Set user ID and name
                    frmMain.lstUsers.AddItem User(Index)
                    Call SendData(Index, "feedback", "adminloginspecial")
                    Call SendUserList
                    Call SendDataAll("serverchat", User(Index) & " has logged in.")
                    
                ElseIf CheckStatus(LCase$(UserData(0))) = "banned" Then
                'Account is banned, so let them know :)
                    Call SendData(Index, "feedback", "banned")
                End If
                
            Else 'Password was incorrect
                Call SendData(Index, "feedback", "passwrong")
            End If
            
        End If
        
    Else
        Call SendData(Index, "feedback", "notregistered")
    End If

ElseIf Controller = "register" Then 'User is registering...
UserData = Split(Message, Chr(156))
    
    If CheckUser(LCase$(UserData(0))) = True Then
        Call SendData(Index, "feedback", "registered")
    Else 'Being registering the username
        UserDB.Add LCase$(UserData(0)) & ";" & LCase$(UserData(1) & ";" & "normal") 'Add user into the database
        Call SaveUserDB 'Save the database
        Call SendData(Index, "feedback", "registercomplete")
    End If
    
ElseIf Controller = "chat" Then
Dim ChatData() As String 'Get chat data
ChatData = Split(Message, Chr(156))

    Call SendDataAll("chat", Message) 'Send to all users

'Serverbot chat commanding
 If ServerbotOn = True Then
    Call ServerbotSpeak(ChatData(3), ChatData(0))
 End If
 
'Here is private messaging commands...
ElseIf Controller = "pm" Then 'Open new pm window
Dim PMData() As String
    PMData = Split(Message, Chr(156))
    Call SendData(GetUserIndex(PMData(0)), "pm", PMData(1))
    
ElseIf Controller = "pmchat" Then 'Pm chat data
Dim PMChatData() As String
    PMChatData = Split(Message, Chr(156))
    Call SendData(GetUserIndex(PMChatData(0)), "pmchat", PMChatData(1) & Chr(156) & PMChatData(2))
   
ElseIf Controller = "pmignore" Then 'User is being ignored
    Call SendData(GetUserIndex(Message), "feedback", "pmignore")

ElseIf Controller = "pmaccept" Then 'PM accepted
    Dim PMAccept() As String
    PMAccept = Split(Message, Chr(156))
    Call SendData(GetUserIndex(PMAccept(0)), "pmaccept", PMAccept(1))

ElseIf Controller = "pmclose" Then 'PM closed

    Dim ClosePM() As String
    ClosePM = Split(Message, Chr(156))
    Call SendData(GetUserIndex(ClosePM(0)), "pmclose", ClosePM(1))
    
ElseIf Controller = "userlist" Then 'Refresh list
    Call SendUserList
    
'Here is moderator commands...
ElseIf Controller = "modkick" Then
Dim KickData() As String
KickData = Split(Message, Chr(156))

    Call Kick(KickData(0), KickData(1)) 'Kicks a user from the chatroom
    
ElseIf Controller = "modclear" Then
    Call SendDataAll("clear", " ") 'Clears chatroom for all users

ElseIf Controller = "modusermessage" Then 'Admin sending user a message
Dim modUser() As String
    modUser = Split(Message, Chr(156))
    Call SendData(GetUserIndex(modUser(0)), "usermessage", modUser(1) & Chr(156) & modUser(2))

ElseIf Controller = "specialAdmin" Then
    Call ChangeStatus(LCase$(Message), "admin")
    
ElseIf Controller = "specialNormal" Then
    Call ChangeStatus(LCase$(Message), "normal")
    
ElseIf Controller = "specialBan" Then
    Call ChangeStatus(LCase$(Message), "banned")

ElseIf Controller = "specialChat" Then
    Call SendDataAll("serverchat", Message)
End If

End Sub

Sub SendData(Index As Integer, Controller As String, Message As String)
On Error Resume Next

frmMain.sckServer(Index).SendData Controller & Chr(171) & Message & Chr(170)
DoEvents
End Sub

Sub SendDataAll(Controller As String, Message As String)
On Error Resume Next
Dim UserLoop As Integer

'Loop through all of the users
For UserLoop = 0 To SocketCount Step 1

    If frmMain.sckServer(UserLoop).State = sckConnected Then
        frmMain.sckServer(UserLoop).SendData Controller & Chr(171) & Message & Chr(170)
        DoEvents
    End If
    
Next UserLoop

End Sub
