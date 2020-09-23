VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFont 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Font Settings"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3360
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
   ScaleHeight     =   1305
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      ScaleHeight     =   225
      ScaleWidth      =   2385
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Save Settings"
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
      MouseIcon       =   "frmFont.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Color:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   540
   End
   Begin VB.Label lblFont 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   435
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Dim FontLoop As Integer

    'Get list of fonts in font combo
    For FontLoop = 0 To Screen.FontCount - 1
        cmbFont.AddItem Screen.Fonts(FontLoop)
    Next FontLoop

'Get already set values...
cmbFont.ListIndex = FontIndex
picColor.BackColor = FontColor
End Sub

Private Sub Label2_Click()
'Set variables..
FontIndex = cmbFont.ListIndex
FontColor = picColor.BackColor

'Save settings and unload
Call SaveSettings
Unload Me
End Sub

Private Sub picColor_Click()
On Error Resume Next

CD.DialogTitle = "Choose Color"
CD.ShowColor

picColor.BackColor = CD.Color
End Sub
