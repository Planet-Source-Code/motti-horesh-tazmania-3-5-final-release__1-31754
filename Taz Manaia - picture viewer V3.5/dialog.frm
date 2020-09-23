VERSION 5.00
Begin VB.Form dialog 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4995
   ClientLeft      =   4485
   ClientTop       =   3300
   ClientWidth     =   6300
   ClipControls    =   0   'False
   Icon            =   "dialog.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3447.638
   ScaleMode       =   0  'User
   ScaleWidth      =   5916.026
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   3375
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   5655
      Begin VB.Label Label6 
         Caption         =   "Now you can choose the size of TAZMANIA "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   5415
      End
      Begin VB.Label Label4 
         Caption         =   "Now you can choose the colors of the interface.!!!!"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   """Slide Show over the whole screen, like ""black screen"",  and settening for the sldeshow, like, set time, background and each."
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Now, All the options, and there are alot of them, are orgenized in menus,"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Now Taz-Mania Support more image files - *.jpg *.gif and *.bmp."
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Now you can set the picture as a wallpaper for the desktop."
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Now you can see the files that are in the same directory as a slide show (like in Power Point)"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1920
      TabIndex        =   0
      Top             =   4080
      Width           =   1260
   End
   Begin VB.Label Label5 
      Caption         =   "What's new In this Version of  TAZMANIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   0
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3600
      Picture         =   "dialog.frx":08CA
      Top             =   3960
      Width           =   480
   End
End
Attribute VB_Name = "dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
    Form1.Show
End Sub

