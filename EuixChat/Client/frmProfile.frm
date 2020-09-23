VERSION 5.00
Begin VB.Form frmProfile 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4365
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
   ScaleHeight     =   2070
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   -120
      TabIndex        =   8
      Top             =   1560
      Width           =   4695
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E9A00C&
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtWebsite 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtQoute 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qoute:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Website:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   750
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public usrNickName As String
Public usrName As String
Public userEmail As String
Public userWebsite As String
Public userQoute As String

Private Sub Form_Load()

'Set all the profile object's captions
Me.Caption = usrNickName & "'s Profile"
txtName.Text = usrName
txtEmail.Text = userEmail
txtWebsite.Text = userWebsite
txtQoute.Text = userQoute

End Sub

Private Sub lblClose_Click()
usrNickName = ""
usrName = ""
userEmail = ""
userWebsite = ""
userQoute = ""

Unload Me
End Sub
