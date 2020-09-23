VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmPrivateMessage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Private Messenge - ()"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5160
   Icon            =   "frmPrivateMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIgnore 
      Caption         =   " Ignore User"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Picture         =   "frmPrivateMessage.frx":1042
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   3975
   End
   Begin SHDocVwCtl.WebBrowser Chat 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4935
      ExtentX         =   8705
      ExtentY         =   5530
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
      Location        =   ""
   End
End
Attribute VB_Name = "frmPrivateMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PMUser As String
Public WindowName As String

Private Sub cmdIgnore_Click()
    Call IgnoreUser(WindowName)
    Unload Me
End Sub

Private Sub cmdSend_Click()

If Len(txtSend.Text) = 0 Or Replace(txtSend, " ", "") = "" Then
    MsgBox "You must enter a message to send!", vbInformation, "Invalid Message"
    Exit Sub
End If

Call Replace(txtSend.Text, Chr(60), "&lt;")
Call Replace(txtSend.Text, Chr(62), "&gt;")
Call AddHTML(Me.Chat, "<font face=Verdana color=red>" & Username & ": </font><font face=Verdana color=black>" & txtSend.Text)
Call SendData("pmchat", PMUser & Chr(156) & Username & Chr(156) & txtSend.Text)

txtSend.Text = ""
txtSend.SetFocus
End Sub

Private Sub Form_Load()
    Me.Caption = "Private Message - (" & PMUser & ")"
    Call ClearHTML(Chat) 'Clear, and set borders
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SendData("pmclose", Me.PMUser & Chr(156) & Username)
PMCount = PMCount - 1
End Sub

Private Sub lblNote_Click()

End Sub
