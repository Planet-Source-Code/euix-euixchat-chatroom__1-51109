Attribute VB_Name = "modUsers"


Function GetUserIndex(Username As String) As Integer
'Gets a users socket number

Dim UserLoop As Integer
For UserLoop = 1 To UBound(User)

    If User(UserLoop) = Username Then
        GetUserIndex = UserLoop
        Exit Function
    End If
    
Next UserLoop

End Function

Function CheckLogin(Username As String) As Boolean
'Checks to see if a user is logged in already.

Dim UserLoop As Integer
For UserLoop = 1 To UBound(User)
    
    If LCase$(User(UserLoop)) = Username Then 'Yup, user is logged in
        CheckLogin = True
        Exit Function
    End If
    
Next UserLoop

'Guess the user isn't logged in, return false
CheckLogin = False
End Function

Function CheckPassword(Username As String, Password As String) As Boolean
'Checks a users password to see if it's correct

Dim UserLoop As Integer
For UserLoop = 1 To UserDB.Count
UserData = Split(UserDB.Item(UserLoop), ";")

    If UserData(0) = Username Then 'Found user, check password
        If UserData(1) = Password Then 'Password is correct!
            CheckPassword = True
            Exit Function
        Else 'Password is incorrect!
            CheckPassword = False
            Exit Function
        End If
    End If
    
Next UserLoop

End Function


Function CheckStatus(Username As String) As String
Dim UserLoop As Integer
Dim UserData() As String

For UserLoop = 1 To UserDB.Count
UserData = Split(UserDB.Item(UserLoop), ";")

    If UserData(0) = Username Then
        CheckStatus = UserData(2)
        Exit Function
    End If
    
Next UserLoop

End Function

Sub Kick(Username As String, Reason As String)
'Kicks a user from the server

Dim UserLoop As Integer

For UserLoop = 1 To UBound(User)

    If User(UserLoop) = Username Then
        Logout User(UserLoop)
        Call AddLog("User " & User(UserLoop) & " was kicked.", "user")
        Call SendDataAll("serverchat", User(UserLoop) & " was kicked from the server.")
        Call SendData(UserLoop, "kick", Reason)
        Call SendUserList   'Refresh user list
        User(UserLoop) = ""
        Exit Sub
    End If
    
Next UserLoop

End Sub

Sub Logout(Username As String)
Dim Count As Integer

For Count = 0 To frmMain.lstUsers.ListCount

    If frmMain.lstUsers.List(Count) = Username Then
        frmMain.lstUsers.RemoveItem (Count)
        Exit Sub
    End If
    
Next Count

End Sub
Sub SendUserList()
Dim UserCount As Integer
Dim Users As Integer
Dim UserType As String 'Admin or Normal?
Dim UserList As String

Users = frmMain.lstUsers.ListCount
For UserCount = 0 To Users
UserType = CheckStatus(LCase$(frmMain.lstUsers.List(UserCount)))

    If UserCount = Users Then
        UserList = UserList & frmMain.lstUsers.List(UserCount) & "." & UserType
        Exit For
    End If
    
UserList = UserList & frmMain.lstUsers.List(UserCount) & "." & UserType & Chr(156)
Next UserCount

Call SendDataAll("userlist", UserList)
End Sub

Sub LoadUserDB()
Dim UserData As String

'Open the user database file
Open App.Path & "/Database/Users.db" For Input As #1
        
    Do Until EOF(1) = True
        Input #1, UserData 'Get username from file
        UserDB.Add UserData 'Add user to database
    DoEvents
    Loop
    
Close #1
End Sub

Sub SaveUserDB()
Dim UserLoop As Integer

'Save the user database file
Kill App.Path & "/Database/Users.db" 'Delete existing database
Open App.Path & "/Database/Users.db" For Output As #1

    For UserLoop = 1 To UserDB.Count
        Print #1, UserDB.Item(UserLoop)
    Next UserLoop

Close #1

End Sub

Function CheckUser(Username As String) As Boolean
'Checks to see if a user is registered

Dim UserLoop As Integer
Dim UserData() As String

For UserLoop = 1 To UserDB.Count
UserData = Split(UserDB.Item(UserLoop), ";")

    If UserData(0) = Username Then
        CheckUser = True
        Exit Function
    End If
    
Next UserLoop

'We couldn't find the user, so return false
CheckUser = False
End Function

Sub ChangeStatus(Username As String, Status As String)
Dim UserLoop As Integer
Dim UserData() As String

For UserLoop = 1 To UserDB.Count
UserData = Split(UserDB.Item(UserLoop), ";")

    If UserData(0) = Username Then
        UserDB.Remove UserLoop
        UserDB.Add UserData(0) & ";" & UserData(1) & ";" & Status
        Exit Sub
    End If

Next UserLoop

End Sub
