VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Register Username"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
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
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraContainer 
      Height          =   735
      Left            =   -120
      TabIndex        =   4
      Top             =   960
      Width           =   5175
      Begin VB.Label cmdRegister 
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
         Left            =   840
         MouseIcon       =   "frmRegister.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label cmdCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
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
         MouseIcon       =   "frmRegister.frx":0CCA
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmdRegister_Click()
On Error Resume Next
Call SendData("register", txtUsername.Text & Chr(156) & txtPassword.Text)
End Sub

