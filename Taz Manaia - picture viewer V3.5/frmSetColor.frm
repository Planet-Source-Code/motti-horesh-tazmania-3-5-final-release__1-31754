VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetColor 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting - Colors"
   ClientHeight    =   5010
   ClientLeft      =   3750
   ClientTop       =   3120
   ClientWidth     =   6540
   Icon            =   "frmSetColor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6540
   Begin VB.Frame cmdClose 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4320
         TabIndex        =   24
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Caption         =   "other areas:"
         Height          =   2895
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton cmdViewAreaBg 
            Caption         =   "View Area bg"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdFileFrameBg 
            Caption         =   " file frame bg"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblFileFrameBg 
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblViewAreaBg 
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "After you'll choose color it will auto applay it self!"
            Height          =   615
            Left            =   120
            TabIndex        =   21
            Top             =   2160
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Font Color:"
         Height          =   2895
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton cmdFileAreaTC 
            Caption         =   "File Area "
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton cmdSlideShowTC 
            Caption         =   "SlideShow - frame"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdWallpaperTC 
            Caption         =   "Wallpaper -frame"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblFileAreaTC 
            Alignment       =   2  'Center
            Caption         =   "Sample"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label lblSlideShowTC 
            Alignment       =   2  'Center
            Caption         =   "Sample"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label lblWallpaperTC 
            Alignment       =   2  'Center
            Caption         =   "Sample"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Back Ground color:"
         Height          =   3855
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton cmdFileAreasBg 
            Caption         =   "File Areas"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   3000
            Width           =   1695
         End
         Begin VB.CommandButton cmdSlideShowBg 
            Caption         =   "SlideShow-frame"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CommandButton cmdWallPaperBg 
            Caption         =   "wallPaper-frame"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton cmdBgColor 
            Caption         =   "BackGround"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblFileAreasBg 
            Height          =   255
            Left            =   360
            TabIndex        =   11
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label lblSlideShowBg 
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label lblWallPaperBg 
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblBgColor 
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   720
            Width           =   1215
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "TazMania 3.0"
   End
End
Attribute VB_Name = "frmSetColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBgColor_Click()
    CommonDialog1.Color = lblBgColor.BackColor
    CommonDialog1.ShowColor
    lblBgColor.BackColor = CommonDialog1.Color
    Form1.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdFileAreasBg_Click()
    CommonDialog1.Color = lblFileAreasBg.BackColor
    CommonDialog1.ShowColor
    lblFileAreasBg.BackColor = CommonDialog1.Color
    Form1.File1.BackColor = CommonDialog1.Color
    Form1.Dir1.BackColor = CommonDialog1.Color
    Form1.Drive1.BackColor = CommonDialog1.Color
    lblFileAreaTC.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdFileAreaTC_Click()
    CommonDialog1.Color = lblFileAreaTC.ForeColor
    CommonDialog1.ShowColor
    lblFileAreaTC.ForeColor = CommonDialog1.Color
    Form1.File1.ForeColor = CommonDialog1.Color
    Form1.Dir1.ForeColor = CommonDialog1.Color
    Form1.Drive1.ForeColor = CommonDialog1.Color
End Sub

Private Sub cmdFileFrameBg_Click()
    CommonDialog1.Color = lblFileFrameBg.BackColor
    CommonDialog1.ShowColor
    lblFileFrameBg.BackColor = CommonDialog1.Color
    Form1.Frame1.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdOK_Click()
frmSetColor.Hide

End Sub

Private Sub cmdSlideShowBg_Click()
    CommonDialog1.Color = lblSlideShowBg.BackColor
    CommonDialog1.ShowColor
    lblSlideShowBg.BackColor = CommonDialog1.Color
    Form1.frmSlideShow.BackColor = CommonDialog1.Color
    lblSlideShowTC.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdSlideShowTC_Click()
    CommonDialog1.Color = lblSlideShowTC.ForeColor
    CommonDialog1.ShowColor
    lblSlideShowTC.ForeColor = CommonDialog1.Color
    Form1.frmSlideShow.ForeColor = CommonDialog1.Color
End Sub

Private Sub cmdViewAreaBg_Click()
    CommonDialog1.Color = lblViewAreaBg.BackColor
    CommonDialog1.ShowColor
    lblViewAreaBg.BackColor = CommonDialog1.Color
    Form1.Picture1.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdWallPaperBg_Click()
    CommonDialog1.Color = lblWallPaperBg.BackColor
    CommonDialog1.ShowColor
    lblWallPaperBg.BackColor = CommonDialog1.Color
    Form1.frmWallPaper.BackColor = CommonDialog1.Color
    lblWallpaperTC.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdWallpaperTC_Click()
    CommonDialog1.Color = lblWallpaperTC.ForeColor
    CommonDialog1.ShowColor
    lblWallpaperTC.ForeColor = CommonDialog1.Color
    Form1.frmWallPaper.ForeColor = CommonDialog1.Color
    Form1.Label2.ForeColor = CommonDialog1.Color
End Sub

Private Sub Command1_Click()
frmSetColor.Hide
End Sub

Private Sub Form_Load()
    lblBgColor.BackColor = Form1.BackColor
    lblWallPaperBg.BackColor = Form1.frmWallPaper.BackColor
    lblSlideShowBg.BackColor = Form1.frmSlideShow.BackColor
    lblFileAreasBg.BackColor = Form1.File1.BackColor
    
    lblWallpaperTC.BackColor = Form1.frmWallPaper.BackColor
    lblSlideShowTC.BackColor = Form1.frmSlideShow.BackColor
    lblFileAreaTC.BackColor = Form1.File1.BackColor
    
    lblWallpaperTC.ForeColor = Form1.frmWallPaper.ForeColor
    lblSlideShowTC.ForeColor = Form1.frmSlideShow.ForeColor
    lblFileAreaTC.ForeColor = Form1.File1.ForeColor
    
    lblFileFrameBg.BackColor = Form1.Frame1.BackColor
    lblViewAreaBg.BackColor = Form1.Picture1.BackColor
End Sub

