Attribute VB_Name = "modData"
Function CheckPMInstance(Username As String) As Boolean
Dim PMLoop As Integer

For PMLoop = 1 To PMCount

    If PMWindow(PMLoop).WindowName = Username Then
        CheckPMInstance = True
        Exit Function
    End If
    
Next PMLoop

'Couldnt find the PM.. guess window isn't already open
CheckPMInstance = False
End Function

Sub SpliceData(Data As String)
Dim DataLoop As Integer 'Loop of data received
Dim DataCarriage() As String 'End of Data
Dim DataArray() As String 'Split the data into parts for parsing

DataCarriage = Split(Data, Chr(170)) 'Split data by ª

'Loop through all the data
For DataLoop = 0 To UBound(DataCarriage)
If DataCarriage(DataLoop) = "" Then Exit Sub

    'Get the controller and messages
    DataArray = Split(DataCarriage(DataLoop), Chr(171))
    Call ParseData(DataArray(0), DataArray(1)) 'Parse data
Next DataLoop

End Sub

Sub ParseData(Controller As String, Message As String)
Dim ChatInfo() As String
Dim UserCount As Integer

If Controller = "feedback" Then

    'Parse Feedback messages
    If Message = "login" Then
        frmLogin.Hide
        frmMain.Show
    ElseIf Message = "adminlogin" Then 'User is admin, give him rights
        frmLogin.Hide
        frmMain.Show
        frmMain.menuAdmin.Visible = True
    ElseIf Message = "adminloginspecial" Then 'Special Admin, give him rights
        frmLogin.Hide
        frmMain.Show
        frmMain.menuAdmin.Visible = True
        frmMain.modSpecial.Visible = True
    ElseIf Message = "notregistered" Then
        MsgBox "The username you entered isn't registered!" & vbCrLf & "Please register it before logging in.", vbInformation, "Not Registered"
        frmLogin.lblLogin.Enabled = True
        frmLogin.lblRegister.Enabled = True
        Exit Sub
    ElseIf Message = "registered" Then
        MsgBox "The username you entered is already registered!" & vbCrLf & "Please choose another username.", vbInformation, "Username Registered"
        frmRegister.Show
        Exit Sub
    ElseIf Message = "registercomplete" Then
        MsgBox "Your username was successfully registered!", vbInformation, "Registration Complete"
        Unload frmRegister
        frmLogin.Show
        Exit Sub
    ElseIf Message = "loggedin" Then
        MsgBox "The username you entered is already logged in!", vbInformation, "User Logged In"
        frmLogin.lblLogin.Enabled = True
        Exit Sub
    ElseIf Message = "passwrong" Then 'Password is incorrect
        MsgBox "The password you entered was incorrect!", vbInformation, "Password Incorrect"
        frmLogin.lblLogin.Enabled = True
        Exit Sub
    ElseIf Message = "banned" Then 'Account is banned
        MsgBox "This account is currently banned!" & vbCrLf & "Contact admin if you beleive this is a mistake.", vbInformation, "Account Banned"
        End
    ElseIf Message = "iplogin" Then 'IP Already connected
        MsgBox "Your IP Address is already logged into TKChat!", vbInformation, "User Logged In"
        End
    ElseIf Message = "pmignore" Then
        MsgBox "This user is ignoring you!", vbInformation, "Ignore"
        Exit Sub
    End If
    
ElseIf Controller = "chat" Then
ChatInfo = Split(Message, Chr(156))

'Replace < and > with the string equivalents
ChatInfo(3) = Replace(ChatInfo(3), Chr(60), "&lt;")
ChatInfo(3) = Replace(ChatInfo(3), Chr(62), "&gt;")

'Replace image code, so we can make images in the chatroom
If InStr(1, ChatInfo(3), "[img=") > 0 Then
    ChatInfo(3) = Replace(ChatInfo(3), "[img=", "<img src=")
    ChatInfo(3) = Replace(ChatInfo(3), " /img]", ">")
End If

'Replace link code, so we can make links in the chatroom
Dim Link() As String, LinkData() As String

If InStr(1, ChatInfo(3), "[link=") > 0 Then
    Link = Split(ChatInfo(3), "[link=")
    LinkData = Split(Link(1), " /link]")
    
    ChatInfo(3) = Replace(ChatInfo(3), "[link=", "<a href=" & Chr(34) & LinkData(0) & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">")
    ChatInfo(3) = Replace(ChatInfo(3), " /link]", "</a>")
End If

'Check for emoticons
If EmoteStatus = True Then
    Dim EmoteLoop As Integer
    Dim Emotes() As String
    Emotes = Split(Emoticons, " ") 'Split emoticons var by spaces

'Loop through emoticons, then replace the emoticons in text with images.
For EmoteLoop = 0 To UBound(Emotes)

    If InStr(1, ChatInfo(3), Emotes(EmoteLoop)) > 0 Then
        ChatInfo(3) = Replace(ChatInfo(3), Emotes(EmoteLoop), EmoteFilename(Emotes(EmoteLoop)), , , vbTextCompare)
    End If
      
Next EmoteLoop
End If

    If Username = ChatInfo(0) Then
        Call AddHTML(frmMain.Chat, "<font face=Verdana color=red>" & ChatInfo(0) & ":</font> <font face=" & ChatInfo(1) & " color=" & ChatInfo(2) & ">" & ChatInfo(3) & "</font>")
    Else
        Call AddHTML(frmMain.Chat, "<font face=Verdana color=blue>" & ChatInfo(0) & ":</font> <font face=" & ChatInfo(1) & " color=" & ChatInfo(2) & ">" & ChatInfo(3) & "</font>")
    End If
    
ElseIf Controller = "userlist" Then 'Parse user list
Dim UserList() As String
Dim UserType() As String
Dim IconNum As Integer

UserList = Split(Message, Chr(156)) 'Split the user list data
frmMain.lstUsers.Nodes.Clear 'Clear the tree list

'Make the admin and user nodes
frmMain.lstUsers.Nodes.Add , , , "Admin"
frmMain.lstUsers.Nodes(1).Bold = True
frmMain.lstUsers.Nodes(1).Expanded = True

frmMain.lstUsers.Nodes.Add , , , "Users"
frmMain.lstUsers.Nodes(2).Bold = True
frmMain.lstUsers.Nodes(2).Expanded = True

    For UserCount = 0 To UBound(UserList)
    'Split the username to get the usertype
    UserType = Split(UserList(UserCount), ".")
    
        If UserType(0) = "" Then Exit For 'Exit if the username is nothing
        Dim IgnoreUsername As String
        IgnoreUsername = UserType(0)
            
            If CheckIgnore(LCase$(IgnoreUsername)) = True Then
                IconNum = 2
            Else
                IconNum = 1
            End If
        
        'Check user's type to determine where he's placed.
        If UserType(1) = "normal" Then
            frmMain.lstUsers.Nodes.Add 2, tvwChild, , UserType(0), IconNum
        ElseIf UserType(1) = "admin" Then 'Moderator
            frmMain.lstUsers.Nodes.Add 1, tvwChild, , UserType(0), IconNum
        ElseIf UserType(1) = "specialadmin" Then 'Special Moderator
            frmMain.lstUsers.Nodes.Add 1, tvwChild, , UserType(0), IconNum
        End If
        
    Next UserCount
       
ElseIf Controller = "globalmess" Then 'Global message
    MsgBox Message, vbInformation, "Server Message"
    
ElseIf Controller = "serverchat" Then 'Server message
    Call AddHTML(frmMain.Chat, "<font face=Verdana color=lightgreen>" & Message)
    
ElseIf Controller = "shutdown" Then 'Server is shutting down
    MsgBox Message, vbInformation, "Server Shutdown"
    frmMain.sckClient.Close
    
ElseIf Controller = "clear" Then 'Clear the chatroom
    Call ClearHTML(frmMain.Chat)
    
ElseIf Controller = "usermessage" Then 'User Message
Dim modUser() As String
    modUser = Split(Message, Chr(156))
    MsgBox modUser(1), vbExclamation, "Admin Message from: " & modUser(0)

ElseIf Controller = "kick" Then 'User was kicked! Haha
    MsgBox "You have been kicked from the server!" & vbCrLf & "Reason: " & Message, vbInformation, "Kicked"
    End
    
'Private messaging commands
ElseIf Controller = "pm" Then 'Opens new PM window
    
    If CheckIgnore(Message) = True Then
        Call SendData("pmignore", Message)
        Exit Sub
    Else
        Call SendData("pmaccept", Message & Chr(156) & Username)
        Call NewPM(Message)
    End If
    
ElseIf Controller = "pmchat" Then 'Chat with user in PM
Dim PMMessage() As String
Dim PMLoop As Integer
PMMessage = Split(Message, Chr(156))

PMMessage(1) = Replace(PMMessage(1), Chr(60), "&lt;")
PMMessage(1) = Replace(PMMessage(1), Chr(62), "&gt;")

    For PMLoop = 1 To PMCount
    
        If PMWindow(PMLoop).WindowName = LCase$(PMMessage(0)) Then
            Call AddHTML(PMWindow(PMLoop).Chat, "<font face=Verdana color=blue>" & PMMessage(0) & ": </font><font face=Verdana color=#000000>" & PMMessage(1) & "</font>")
        Else
        End If
        
    Next PMLoop

ElseIf Controller = "pmaccept" Then 'PM Accepted
    Call NewPM(Message)
    
ElseIf Controller = "pmclose" Then 'PM Window was closed
Dim PMWinLoop As Integer

    For PMWinLoop = 1 To PMCount
    
        If PMWindow(PMWinLoop).WindowName = LCase$(Message) Then
            Unload PMWindow(PMWinLoop)
            Exit Sub
        End If
        
    Next PMWinLoop
    
End If

End Sub

Sub SendData(Controller As String, Message As String)
'Sends data to the server in the protocol
frmMain.sckClient.SendData Controller & Chr(171) & Message & Chr(170)
End Sub

Sub LoadSettings()
On Error Resume Next

Dim Setting1 As String
Dim Setting2 As String
Dim Setting3 As String

Open App.Path & "\Settings.dat" For Input As #1
    Input #1, Setting1
    FontIndex = CInt(Setting1)
    Input #1, Setting2
    FontColor = CLng(Setting2)
    Input #1, Setting3
    
    If Setting3 = "On" Then
        EmoteStatus = True
    ElseIf Setting3 = "Off" Then
        EmoteStatus = False
    End If
    
Close #1

End Sub

Sub SaveSettings()
On Error Resume Next

Open App.Path & "\Settings.dat" For Output As #1
    Print #1, CStr(FontIndex)
    Print #1, CStr(FontColor)
    
    If EmoteStatus = True Then
        Print #1, "On"
    Else
        Print #1, "Off"
    End If
    
Close #1

End Sub
