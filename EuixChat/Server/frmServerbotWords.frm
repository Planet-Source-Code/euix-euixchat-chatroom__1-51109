VERSION 5.00
Begin VB.Form frmServerbotWords 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Serverbot Word List"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServerbotWords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Word"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Word"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Word Action"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ListBox lstWords 
      Appearance      =   0  'Flat
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Words:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "This edits serverbot's database to make him say certain things when a user says one of his keywords."
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmServerbotWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub LoadWordList()
Dim WordLoop As Integer
Dim WordData() As String

lstWords.Clear
For WordLoop = 1 To ServerbotDB.Count
    WordData = Split(ServerbotDB.Item(WordLoop), Chr(156))
    lstWords.AddItem WordData(0)
Next WordLoop

End Sub

Private Sub cmdAdd_Click()
Dim Word As String
Word = InputBox("Enter a word or statement for serverbot to look for!", "Add Word")
If Word = "" Then Exit Sub

lstWords.AddItem LCase$(Word)
ServerbotDB.Add LCase$(Word) & Chr(156) & ""
Call SaveServerbotDB
End Sub

Private Sub cmdDelete_Click()
Dim WordLoop As Integer
Dim WordData() As String

For WordLoop = 1 To ServerbotDB.Count
    WordData = Split(ServerbotDB.Item(WordLoop), Chr(156))
    
    If WordData(0) = lstWords.List(lstWords.ListIndex) Then
        ServerbotDB.Remove WordLoop
        lstWords.RemoveItem lstWords.ListIndex
        Exit Sub
    End If
    
Next WordLoop

Call SaveServerbotDB
End Sub

Private Sub cmdEdit_Click()
Dim WordLoop As Integer
Dim WordData() As String

For WordLoop = 1 To ServerbotDB.Count
    WordData = Split(ServerbotDB.Item(WordLoop), Chr(156))
    
    If WordData(0) = lstWords.List(lstWords.ListIndex) Then
        frmEditWord.ServerbotWord = lstWords.List(lstWords.ListIndex)
        frmEditWord.ServerbotReaction = WordData(1)
        frmEditWord.Show
        Exit Sub
    End If
    
Next WordLoop
End Sub

Private Sub Form_Load()
Call LoadWordList
End Sub
