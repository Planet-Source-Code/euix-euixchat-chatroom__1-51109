VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EuixChat Server"
   ClientHeight    =   3675
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKick 
      Caption         =   "Kick User"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser Log 
      Height          =   3255
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   5741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer tmrStatus 
      Interval        =   500
      Left            =   1200
      Top             =   480
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3420
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   6297
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "1:33 PM"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   720
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   15000
   End
   Begin VB.ListBox lstUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Users:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   555
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu fileStart 
         Caption         =   "&Start Server"
      End
      Begin VB.Menu fileShutdown 
         Caption         =   "&Shutdown Server"
      End
   End
   Begin VB.Menu menuUsers 
      Caption         =   "&Users"
      Begin VB.Menu usersEdit 
         Caption         =   "&Edit Registered Users"
      End
   End
   Begin VB.Menu menuServerbot 
      Caption         =   "&Serverbot"
      Begin VB.Menu serverbotStatus 
         Caption         =   "&Turn Serverbot Off"
      End
      Begin VB.Menu serverbotWordList 
         Caption         =   "&Edit Word List"
      End
   End
   Begin VB.Menu menuCommands 
      Caption         =   "&Commands"
      Begin VB.Menu cmdSend 
         Caption         =   "&Send Global Message"
      End
      Begin VB.Menu cmdChatMess 
         Caption         =   "&Send Chat Message"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChatMess_Click()
Dim Message As String
Message = InputBox("Enter the Chat Message to Send", "Send Global Message")

    If Message = "" Then Exit Sub
    Call SendDataAll("serverchat", Message)
End Sub

Private Sub cmdKick_Click()
Dim KickReason As String

KickReason = InputBox("Enter a reason for kicking user", "Kick Reason")
If KickReason = "" Then Exit Sub

    Call Kick(lstUsers.List(UserIndex), KickReason)
End Sub

Private Sub cmdSend_Click()
Dim Message As String
Message = InputBox("Enter the Global Message to Send", "Send Global Message")

    If Message = "" Then Exit Sub
    Call SendDataAll("globalmess", Message)
End Sub

Private Sub fileShutdown_Click()
frmShutdown.Show
End Sub

Private Sub fileStart_Click()

'Check the caption and Start/Stop the Server

If fileStart.Caption = "&Start Server" Then
    sckServer(0).Listen 'Listen for connections
    Call AddLog("Started Server - " & CStr(Time), "Server")
    fileStart.Caption = "&Stop Server" 'Set caption so we can stop the server
    Exit Sub 'Exit so we don't cause an endless loop
ElseIf fileStart.Caption = "&Stop Server" Then
    sckServer(0).Close 'Stop listening
    Call EndServer
    Call AddLog("Stopped Server - " & CStr(Time), "Server")
    fileStart.Caption = "&Start Server" 'Set caption so we can start it up
    Exit Sub 'Exit so we don't cause an endless loop
End If

End Sub

Private Sub Form_Load()

'Load the user database
Call LoadUserDB

'Load the serverbot database and turn serverbot on
Call LoadServerbotDB
ServerbotOn = True

'Setup the server log
Log.Navigate "about:blank"

'Set the connections to zero
Status.Panels(2).Text = "0 Connections"

Do While Log.ReadyState <> READYSTATE_COMPLETE
    DoEvents
Loop
    DoEvents
    
Log.Document.body.Style.border = "1px solid black"
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstUsers_Click()
UserIndex = lstUsers.ListIndex
End Sub

Private Sub sckServer_Close(Index As Integer)

If User(Index) <> "" Then
    Logout User(Index) 'Remove user from list
    Call AddLog("User " & User(Index) & " disconnected.", "user")
    Call SendDataAll("serverchat", User(Index) & " has logged out.") 'Send notification to all sockets
    Call SendUserList 'Refresh the user list
    User(Index) = "" 'Delete the username
End If

End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

'Increase the socket count
SocketCount = SocketCount + 1
'Increase the number of connections
Connections = Connections + 1
'Create a new socket
Load sckServer(SocketCount)
'Accept the request with the new socket
sckServer(SocketCount).Accept requestID

Status.Panels(2).Text = Connections & " Connections"
Call AddLog("Connection Request - ID# " & CStr(requestID), "Connection")

End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String 'Data recevied from clients
sckServer(Index).GetData strData

'Splice the data so we can parse it
Call SpliceData(Index, strData)
End Sub

Private Sub serverbotStatus_Click()

If serverbotStatus.Caption = "&Turn Serverbot On" Then
    serverbotStatus.Caption = "&Turn Serverbot Off"
    ServerbotOn = True
    Exit Sub
ElseIf serverbotStatus.Caption = "&Turn Serverbot Off" Then
    serverbotStatus.Caption = "&Turn Serverbot On"
    ServerbotOn = False
    Exit Sub
End If
    
End Sub

Private Sub serverbotWordList_Click()
frmServerbotWords.Show
End Sub

Private Sub tmrStatus_Timer()
On Error Resume Next

'Checks the current state of the server socket and sets the Status
'Bar accordingly.

If sckServer(0).State = sckListening Then
    Status.Panels(1).Text = "Listening for Connections..."
ElseIf sckServer(0).State = sckClosed Then
    Status.Panels(1).Text = "Server not running"
End If

End Sub

Private Sub usersEdit_Click()
frmUsers.Show
End Sub
