VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EuixChat"
   ClientHeight    =   6630
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9555
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
   ScaleHeight     =   6630
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picEmoticons 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   4320
      ScaleHeight     =   945
      ScaleWidth      =   2025
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
      Begin VB.Image Emote 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   6
         Left            =   1080
         Picture         =   "frmMain.frx":1042
         Top             =   480
         Width           =   300
      End
      Begin VB.Image Emote 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   600
         Picture         =   "frmMain.frx":1308
         Top             =   480
         Width           =   300
      End
      Begin VB.Image Emote 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   120
         Picture         =   "frmMain.frx":15CE
         Top             =   480
         Width           =   300
      End
      Begin VB.Image Emote 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   1560
         Picture         =   "frmMain.frx":1894
         Top             =   120
         Width           =   300
      End
      Begin VB.Image Emote 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   1080
         Picture         =   "frmMain.frx":1B64
         Top             =   120
         Width           =   300
      End
      Begin VB.Image Emote 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":1E28
         Stretch         =   -1  'True
         Top             =   120
         Width           =   300
      End
      Begin VB.Image Emote 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":2277
         Top             =   120
         Width           =   300
      End
   End
   Begin MSComctlLib.ImageList UserList 
      Left            =   240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":253E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3590
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Toolbar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      ScaleHeight     =   375
      ScaleWidth      =   1575
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
      Begin VB.PictureBox Tool 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   1200
         Picture         =   "frmMain.frx":392A
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   9
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox Tool 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmMain.frx":3BF1
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox Tool 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "frmMain.frx":3F7B
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox Tool 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":4305
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   0
         Width           =   240
      End
   End
   Begin SHDocVwCtl.WebBrowser Chat 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      ExtentX         =   12091
      ExtentY         =   9763
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
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   6240
      Width           =   975
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   120
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "207.6.66.147"
      RemotePort      =   15000
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   120
      MaxLength       =   200
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   6240
      Width           =   5655
   End
   Begin MSComctlLib.TreeView lstUsers 
      Height          =   6135
      Left            =   7080
      TabIndex        =   8
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   10821
      _Version        =   393217
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblUsers 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Users:"
      Height          =   195
      Left            =   7080
      TabIndex        =   0
      Top             =   120
      Width           =   555
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu fileDis 
         Caption         =   "&Disconnect"
      End
   End
   Begin VB.Menu menuOptions 
      Caption         =   "&Options"
      Begin VB.Menu optionsFont 
         Caption         =   "&Font Settings"
      End
      Begin VB.Menu optionsEmote 
         Caption         =   "&Turn Emoticons Off"
      End
   End
   Begin VB.Menu menuAdmin 
      Caption         =   "&Moderator Commands"
      Visible         =   0   'False
      Begin VB.Menu cmdKick 
         Caption         =   "&Kick User"
      End
      Begin VB.Menu modSendMess 
         Caption         =   "&Send User Message"
      End
      Begin VB.Menu cmdClear 
         Caption         =   "&Clear Chatroom"
      End
      Begin VB.Menu modSpecial 
         Caption         =   "S&pecial"
         Visible         =   0   'False
         Begin VB.Menu specialMakeAdmin 
            Caption         =   "&Make User Admin"
         End
         Begin VB.Menu specialMakeNormal 
            Caption         =   "&Make User Normal"
         End
         Begin VB.Menu specialUserBan 
            Caption         =   "&Ban User"
         End
         Begin VB.Menu specialChat 
            Caption         =   "&Send Chat Message"
         End
      End
   End
   Begin VB.Menu menuUserlist 
      Caption         =   "USERLIST"
      Visible         =   0   'False
      Begin VB.Menu userlistIgnore 
         Caption         =   "&Ignore User"
      End
      Begin VB.Menu userlistUnIgnore 
         Caption         =   "&UnIgnore User"
      End
      Begin VB.Menu userlistPM 
         Caption         =   "&Private Message"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu helpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()
    Call SendData("modclear", " ")
End Sub

Private Sub cmdKick_Click()
Dim KickReason As String

If frmMain.lstUsers.SelectedItem.Text = "" Then Exit Sub
KickReason = InputBox("Enter a reason for kicking user", "Kick Reason")
If KickReason = "" Then Exit Sub

    Call SendData("modkick", frmMain.lstUsers.SelectedItem.Text & Chr(156) & KickReason)
End Sub

Private Sub cmdSend_Click()

If Len(txtText.Text) = 0 Or Replace(txtText, " ", "") = "" Then
    MsgBox "You must enter a message to send!", vbInformation, "Invalid Message"
    Exit Sub
End If

Call SendData("chat", Username & Chr(156) & Screen.Fonts(FontIndex) & Chr(156) & ConvertHex(FontColor) & Chr(156) & txtText.Text)

txtText.Text = ""
txtText.SetFocus
End Sub

Private Sub Emote_Click(Index As Integer)

    If Index = 0 Then 'Smile
        txtText.SelText = ":)"
    ElseIf Index = 1 Then 'Sad
        txtText.SelText = ":("
    ElseIf Index = 2 Then 'Grin
        txtText.SelText = ":D"
    ElseIf Index = 3 Then 'Eh?
        txtText.SelText = "o_O"
    ElseIf Index = 4 Then 'Oh
        txtText.SelText = ":O"
    ElseIf Index = 5 Then 'Tounge
        txtText.SelText = ":P"
    ElseIf Index = 6 Then 'Wink
        txtText.SelText = ";)"
    End If
    
End Sub

Private Sub Emote_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Emote(Index).BorderStyle = 1
End Sub

Private Sub fileDis_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

'Set the image list for the user list.
lstUsers.ImageList = UserList

Call ClearHTML(Chat)  'Setup the Chat Window
Call LoadSettings 'Load settings
Call SetEmoticons 'Set the emoticons variable

If EmoteStatus = False Then
    optionsEmote.Caption = "T&urn Emoticons On"
ElseIf EmoteStatus = True Then
    optionsEmote.Caption = "&Turn Emoticons Off"
End If

MsgBox CStr(IgnoreDB.Item(1))
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call SaveIgnoreList
End

End Sub

Private Sub helpAbout_Click()
MsgBox "TKChat v" & App.Major & "." & App.Minor & " programmed by Euix" & vbCrLf & "http://euix.geahost.com", vbInformation, "About"
End Sub

Private Sub lstUsers_DblClick()

If lstUsers.SelectedItem.Text = "Admin" Then
    Exit Sub
ElseIf lstUsers.SelectedItem.Text = "Users" Then
    Exit Sub
ElseIf lstUsers.SelectedItem.Text = Username Then
    MsgBox "You can't Private Message yourself!", vbInformation, "Private Message"
    Exit Sub
End If

If lstUsers.SelectedItem.Text <> "" Then

    If CheckPMInstance(LCase$(lstUsers.SelectedItem.Text)) = False Then
        Call SendData("pm", lstUsers.SelectedItem.Text & Chr(156) & Username)
    Else
        MsgBox "You already have a Private Message window open!", vbInformation, "Private Message"
        Exit Sub
    End If
End If

End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        PopupMenu menuUserlist
    End If
    
End Sub

Private Sub modSendMess_Click()
Dim Message As String

    If frmMain.lstUsers.SelectedItem.Text = "" Then Exit Sub
    Message = InputBox("Enter a message to send to " & frmMain.lstUsers.SelectedItem.Text & ".", "Send Message")
    If Message = "" Then Exit Sub
    
    Call SendData("modusermessage", lstUsers.SelectedItem.Text & Chr(156) & Username & Chr(156) & Message)
    
End Sub

Private Sub optionsEmote_Click()

If optionsEmote.Caption = "&Turn Emoticons Off" Then
    optionsEmote.Caption = "T&urn Emoticons On"
    EmoteStatus = False
    Exit Sub
ElseIf optionsEmote.Caption = "T&urn Emoticons On" Then
    optionsEmote.Caption = "T&urn Emoticons Off"
    EmoteStatus = True
    Exit Sub
End If

End Sub

Private Sub optionsFont_Click()
frmFont.Show
End Sub



Private Sub picEmoticons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim EmoteCount As Integer
    
    For EmoteCount = 0 To Emote.UBound
        Emote(EmoteCount).BorderStyle = 0
    Next EmoteCount
    
End Sub

Private Sub sckClient_Close()
Unload Me
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
Dim strData As String 'Data received
sckClient.GetData strData 'Put data received into variable

'Splice the data
Call SpliceData(strData)

End Sub

Private Sub specialChat_Click()
Dim Message As String

    Message = InputBox("Enter a message to send to the chatroom.", "Send Message")
    If Message = "" Then Exit Sub
    
    Call SendData("specialChat", Message)

End Sub

Private Sub specialMakeAdmin_Click()

If lstUsers.SelectedItem.Text = "" Then Exit Sub
Call SendData("specialAdmin", lstUsers.SelectedItem.Text)

End Sub

Private Sub specialMakeNormal_Click()

If lstUsers.SelectedItem.Text = "" Then Exit Sub
Call SendData("specialNormal", lstUsers.SelectedItem.Text)

End Sub

Private Sub specialUserBan_Click()
Dim BanReason As String

If lstUsers.SelectedItem.Text = "" Then Exit Sub
BanReason = InputBox("Enter a reason for banning user", "Ban Reason")
If BanReason = "" Then Exit Sub

Call SendData("specialBan", lstUsers.SelectedItem.Text)
Call SendData("modkick", lstUsers.SelectedItem.Text & Chr(156) & BanReason)
End Sub

Private Sub Tool_Click(Index As Integer)
Dim Link As String, Picture As String

If Index = 0 Then 'Font options
    frmFont.Show
ElseIf Index = 1 Then 'Insert Link
    On Error Resume Next
    Link = InputBox("Enter a URL to be sent as a hyperlink", "Enter a URL")
    If Link = "" Then Exit Sub
    
    'Input the link command
    txtText.Text = txtText.Text & "[link=" & Link & " /link]"
ElseIf Index = 2 Then 'Insert Picture
    On Error Resume Next
    Picture = InputBox("Enter a URL to be sent as a picture", "Enter a URL")
    If Picture = "" Then Exit Sub
    
    'Input the picture command
    txtText.Text = txtText.Text & "[img=" & Picture & " /img]"
ElseIf Index = 3 Then
    
    If picEmoticons.Visible = False Then
        picEmoticons.Visible = True
        Exit Sub
    ElseIf picEmoticons.Visible = True Then
        picEmoticons.Visible = False
        Exit Sub
    End If
    
End If

End Sub

Private Sub userProfile_Click()
    
    'Gets a users profile
    If lstUsers.SelectedItem.Text = "" Then Exit Sub
    Call SendData("profile", lstUsers.SelectedItem.Text)
    
End Sub


Private Sub userlistIgnore_Click()
On Error Resume Next

    If lstUsers.SelectedItem.Text = "" Then Exit Sub
    Call IgnoreUser(lstUsers.SelectedItem.Text)
    
End Sub

Private Sub userlistPM_Click()
On Error Resume Next

'If lstUsers.SelectedItem.Text = "Admin" Then
'    Exit Sub
'ElseIf lstUsers.SelectedItem.Text = "Users" Then
'    Exit Sub
'ElseIf lstUsers.SelectedItem.Text = Username Then
'    MsgBox "You can't Private Message yourself!", vbInformation, "Private Message"
'    Exit Sub
'End If

If lstUsers.SelectedItem.Text <> "" Then
    Call NewPM(lstUsers.SelectedItem.Text)
    Call SendData("pm", lstUsers.SelectedItem.Text & Chr(156) & Username)
End If

End Sub

Private Sub userlistUnIgnore_Click()
On Error Resume Next

    If lstUsers.SelectedItem.Text = "" Then Exit Sub
    Call UnIgnoreUser(lstUsers.SelectedItem.Text)
End Sub
