VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSlideShowProp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5820
   ClientLeft      =   2640
   ClientTop       =   2925
   ClientWidth     =   8625
   Icon            =   "frmSlideShowProp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8625
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   720
      TabIndex        =   19
      Top             =   5160
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7920
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   33
   End
   Begin VB.Frame Frame2 
      Caption         =   "Play :"
      Height          =   3615
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
      Begin VB.Label Label6 
         Caption         =   $"frmSlideShowProp.frx":08CA
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Press Play."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Choose background color for the back of the screen."
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Set pass time for each picture."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "* The Play a ""Slide Show"" in TAZMANIA's CINIMAS, u need to set some setting on the oter frame, the settin are :"
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Setting :"
      Height          =   3615
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   5535
      Begin VB.TextBox txtSeconds 
         Height          =   285
         Left            =   4320
         TabIndex        =   21
         Top             =   3120
         Width           =   495
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton cmdTextSize 
         Caption         =   "Font - Style and size"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton cmdFontColor 
         Caption         =   "Font Color"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmdBackColor 
         Caption         =   "Back Color"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "How it will look :"
         Height          =   2655
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   2655
         Begin VB.Image Image2 
            Height          =   840
            Left            =   840
            Picture         =   "frmSlideShowProp.frx":0958
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   960
         End
         Begin VB.Label lblText2 
            BackStyle       =   0  'Transparent
            Caption         =   "Taz-Mania 3.0"
            Height          =   255
            Left            =   1080
            TabIndex        =   11
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblText1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "The file location..."
            Height          =   195
            Left            =   360
            TabIndex        =   10
            Top             =   600
            Width           =   1260
         End
         Begin VB.Label lblBackColor 
            BorderStyle     =   1  'Fixed Single
            Height          =   2175
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Label Label8 
         Caption         =   "sec."
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Change picture each"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label lblTextStyle 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblTextSize 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblSample 
         Alignment       =   2  'Center
         Caption         =   "Sample"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8400
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   6600
      Picture         =   "frmSlideShowProp.frx":0D9A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to TAZMANIA's CINIMAS!!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmSlideShowProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmSlideShowProp.Hide
End Sub
Private Sub cmdFontColor_Click()
    CD1.Color = lblText1.ForeColor
    CD1.ShowColor
    lblText1.ForeColor = CD1.Color
    lblText2.ForeColor = CD1.Color
    lblSample.ForeColor = CD1.Color
    frmShow.lblFileLocation.ForeColor = CD1.Color
    frmShow.lblTaz.ForeColor = CD1.Color
End Sub

Private Sub cmdBackColor_Click()
    CD1.Color = lblBackColor.BackColor
    CD1.ShowColor
    lblBackColor.BackColor = CD1.Color
    lblSample.BackColor = CD1.Color
    frmShow.BackColor = CD1.Color
End Sub

Private Sub cmdStart_Click()
    frmShow.Timer1.Interval = txtSeconds * 100
    frmShow.Show
End Sub

Private Sub cmdTextSize_Click()
    CD1.FontSize = lblText1.Font.Size
    CD1.FontName = lblText1.Font.Name
    CD1.ShowFont
    lblTextSize.Caption = CD1.FontSize
    lblTextStyle.Caption = CD1.FontName
    
    frmShow.lblFileLocation.Font.Size = CD1.FontSize
    frmShow.lblFileLocation.Font.Name = CD1.FontName
    
    frmShow.lblTaz.Font.Size = CD1.FontSize
    frmShow.lblTaz.Font.Name = CD1.FontName
End Sub
