VERSION 5.00
Begin VB.Form frmShow 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   3480
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "X"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.Image picture1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblTaz 
      BackStyle       =   0  'Transparent
      Caption         =   "Taz-Mania 3.0"
      Height          =   255
      Left            =   7560
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblFileLocation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   5400
      Width           =   480
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
Timer1.Enabled = False
frmShow.Hide
End Sub

Private Sub cmdPlay_Click()
Timer1.Enabled = True

End Sub

Private Sub cmdStop_Click()
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    frmShow.Left = 0
    frmShow.top = 0
    frmShow.Height = Screen.Height
    frmShow.Width = Screen.Width
          
    Dim midPictureWi As Integer
    Dim midScreenWi As Integer
    Dim midPictureHe As Integer
    Dim midScreenHe As Integer

    midPictureWi = Picture1.Width / 2
    midScreenWi = Screen.Width / 2
    midPictureHe = (Picture1.Height) / 2
    midScreenHe = (Screen.Height) / 2
    


End Sub





Private Sub Timer1_Timer()
    On Error Resume Next
        Timer1.Tag = Timer1.Tag + 1
    If Timer1.Tag >= Int(0.1 * 60) Then
        Timer1.Tag = 0
        cmdRnd_Click
End If
   End Sub



Private Sub cmdNext_Click()
    Timer1.Tag = 0
    If Form1.File1.ListCount Then Form1.File1.ListIndex = (Form1.File1.ListIndex + 1) Mod File1.ListCount
End Sub

Private Sub cmdRnd_Click()
    Form1.File1.ListIndex = Int(Rnd * Form1.File1.ListCount)
End Sub
