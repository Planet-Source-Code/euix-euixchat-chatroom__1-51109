VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Toolkit Chat Login"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmLogin.frx":1042
   ScaleHeight     =   1395
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUsername 
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
      Left            =   1200
      MaxLength       =   26
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.Frame fraContainer 
      Height          =   735
      Left            =   -120
      TabIndex        =   2
      Top             =   840
      Width           =   5055
      Begin VB.Label lblRegister 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Register"
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
         Left            =   2520
         MouseIcon       =   "frmLogin.frx":13CC
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblLogin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
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
         Left            =   720
         MouseIcon       =   "frmLogin.frx":2096
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label lblUsername 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   885
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'Check version
Call SendData("vercheck", CStr(App.Major) & Chr(156) & CStr(App.Minor))
Call LoadIgnoreList

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lblLogin_Click()

'Login to TKChat
lblLogin.Enabled = False
Username = txtUsername.Text
Call SendData("login", Username & Chr(156) & txtPassword.Text) 'Send Login Info
End Sub

Private Sub lblRegister_Click()
On Error Resume Next
frmRegister.Show
End Sub
