VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registered Users"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete User"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Change Status"
      Height          =   2415
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton cmdSpecial 
         Caption         =   "Special"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdmin 
         Caption         =   "Admin"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdBanned 
         Caption         =   "Banned"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdNormal 
         Caption         =   "Normal"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.ListBox lstStatus 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox lstUsers 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registered Users:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdmin_Click()
ChangeStatus lstUsers.List(lstUsers.ListIndex), "admin"
lstStatus.List(lstStatus.ListIndex) = "admin"
Call SaveUserDB
End Sub

Private Sub cmdBanned_Click()
ChangeStatus lstUsers.List(lstUsers.ListIndex), "banned"
lstStatus.List(lstStatus.ListIndex) = "banned"
Call SaveUserDB
End Sub

Private Sub cmdDelete_Click()
Dim UserLoop As Integer
Dim UserData() As String
Dim Reason As Integer

Reason = MsgBox("Are you sure you wish to delete the user?" & vbCrLf & "It will permantly remove him from the database!", vbYesNo, "Delete User")

If Reason = vbYes Then
    For UserLoop = 1 To UserDB.Count
    UserData = Split(UserDB.Item(UserLoop), ";")

        If UserData(0) = LCase$(lstUsers.List(lstUsers.ListIndex)) Then
            UserDB.Remove UserLoop
            Call SaveUserDB
        
            lstStatus.RemoveItem lstUsers.ListIndex
            lstUsers.RemoveItem lstUsers.ListIndex
        
            Exit Sub
        End If

    Next UserLoop
    
ElseIf Reason = vbNo Then
    Exit Sub
End If

End Sub


Private Sub cmdNormal_Click()
ChangeStatus lstUsers.List(lstUsers.ListIndex), "normal"
lstStatus.List(lstStatus.ListIndex) = "normal"
Call SaveUserDB
End Sub

Private Sub cmdSpecial_Click()

ChangeStatus lstUsers.List(lstUsers.ListIndex), "specialadmin"
lstStatus.List(lstStatus.ListIndex) = "special"
Call SaveUserDB

End Sub

Private Sub Form_Load()
Dim UserLoop As Integer
Dim UserData() As String

'Load the users from the database
For UserLoop = 1 To UserDB.Count
    UserData = Split(UserDB.Item(UserLoop), ";")
    lstUsers.AddItem UserData(0)
    
    If UserData(2) = "specialadmin" Then
        lstStatus.AddItem "special"
        GoTo NextLoop
    End If
    
    lstStatus.AddItem UserData(2)
    
NextLoop:
Next UserLoop

End Sub

Private Sub lstStatus_Click()
lstUsers.ListIndex = lstStatus.ListIndex
End Sub

Private Sub lstUsers_Click()
lstStatus.ListIndex = lstUsers.ListIndex
End Sub
