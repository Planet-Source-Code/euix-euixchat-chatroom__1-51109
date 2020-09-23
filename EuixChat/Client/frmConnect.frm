VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   80
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picConnect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      Picture         =   "frmConnect.frx":0000
      ScaleHeight     =   1215
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmConnect.Show 'Show the form before connecting

frmMain.sckClient.Connect
Do Until frmMain.sckClient.State = sckConnected

If frmMain.sckClient.State = sckError Then
    lblStatus.Caption = "Connection Failed!"
    MsgBox "Toolkit Chat was unable to connect to the Server!", vbInformation, "Unable to Connect"
    End
End If

lblStatus.Caption = "Connecting..."
DoEvents
Loop

lblStatus.Caption = "Connected."
frmLogin.Show 'Show the login form
Me.Hide
End Sub

