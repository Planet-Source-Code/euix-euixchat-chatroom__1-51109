VERSION 5.00
Begin VB.Form frmBanned 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Banned IPs"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove IP Address"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ListBox lstBanned 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP Addresses:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmBanned"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
Dim BanLoop As Integer

For BanLoop = 1 To BannedDB.Count

    If BannedDB.Item(BanLoop) = lstBanned.List(lstBanned.ListIndex) Then
        BannedDB.Remove BanLoop
    End If
    
Next BanLoop

lstBanned.RemoveItem lstBanned.ListIndex
End Sub

Private Sub Form_Load()
Dim BanLoop As Integer

For BanLoop = 1 To BannedDB.Count
    lstBanned.AddItem BannedDB.Item(BanLoop)
Next BanLoop

End Sub
