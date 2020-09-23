VERSION 5.00
Begin VB.Form frmEditWord 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Word Action"
   ClientHeight    =   2235
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
   Icon            =   "frmEditWord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   -120
      TabIndex        =   4
      Top             =   1680
      Width           =   5055
      Begin VB.Label cmdCancel 
         Alignment       =   2  'Center
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
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label cmdSave 
         Alignment       =   2  'Center
         Caption         =   "Save Word Action"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Serverbot's Reaction"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
      Begin VB.TextBox txtReaction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Label lblWord 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Word:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "frmEditWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ServerbotWord As String
Public ServerbotReaction As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim WordLoop As Integer
Dim WordData() As String

For WordLoop = 1 To ServerbotDB.Count
    WordData = Split(ServerbotDB.Item(WordLoop), Chr(156))
    
    If WordData(0) = Me.ServerbotWord Then
        ServerbotDB.Remove WordLoop
        ServerbotDB.Add Me.ServerbotWord & Chr(156) & Me.txtReaction.Text
        Call SaveServerbotDB
        Unload Me
        Exit Sub
    End If
    
Next WordLoop

End Sub

Private Sub Form_Load()
lblWord.Caption = ServerbotWord
txtReaction.Text = ServerbotReaction
End Sub

Private Sub lbWord_Click()

End Sub

