VERSION 5.00
Begin VB.Form frmShutdown 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shutdown Server"
   ClientHeight    =   1230
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
   Icon            =   "frmShutdown.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   -120
      TabIndex        =   2
      Top             =   720
      Width           =   5055
      Begin VB.Label lblShutdown 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shutdown"
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
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         Enabled         =   0   'False
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
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtShutdown 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a reason to shutdown the server:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub lblCancel_Click()
Unload Me
End Sub

Private Sub lblShutdown_Click()
Dim Reason As Integer

ReasonForShutdown:
If Replace(txtShutdown.Text, " ", "") = "" Then
    Reason = MsgBox("Are you sure you wish to give no reason for shutdown?", vbYesNo, "Shutdown Reason")
Else 'Shutdown the server
    Call SendDataAll("shutdown", txtShutdown.Text)
    End
End If

If Reason = vbYes Then 'Shutdown the server
    Call SendDataAll("shutdown", txtShutdown.Text) 'Send shutdown data to all clients.
    End 'End all server processes, and exit.
ElseIf Reason = vbNo Then 'Don't shutdown the server
    GoTo ReasonForShutdown
End If

End Sub
