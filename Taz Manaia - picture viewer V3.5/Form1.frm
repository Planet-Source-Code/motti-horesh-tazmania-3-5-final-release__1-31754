VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   Caption         =   "TAZ MANIA - Picture Viewer"
   ClientHeight    =   5760
   ClientLeft      =   2655
   ClientTop       =   2295
   ClientWidth     =   9360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9360
   Begin VB.Frame frmSlideShow 
      Caption         =   "Slide Show"
      Height          =   1695
      Left            =   7080
      TabIndex        =   12
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton Command2 
         Caption         =   "stop Slideshow"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "start slideshow"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "What's new?"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Frame frmWallPaper 
      Caption         =   "Wallpaper"
      Height          =   1335
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdWallpaper 
         Caption         =   "Set as wallpaper"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2655
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00808080&
         Height          =   2115
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H00808080&
         Height          =   2625
         Left            =   120
         Pattern         =   "*.GIF;*.jpg;*.bmp"
         TabIndex        =   5
         Top             =   2760
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   2880
      ScaleHeight     =   3315
      ScaleWidth      =   5955
      TabIndex        =   2
      Top             =   1920
      Width           =   6015
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   3
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   5280
      Width           =   6015
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3375
      Left            =   8880
      TabIndex        =   0
      Top             =   1920
      Width           =   255
   End
   Begin VB.PictureBox PicScreen 
      Height          =   1095
      Left            =   600
      ScaleHeight     =   1035
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileWallpaper 
         Caption         =   "Set As wallpaper"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Options"
      Begin VB.Menu mnuOptionsColor 
         Caption         =   "S&et Colors"
      End
   End
   Begin VB.Menu mnuOptionsPropSlide 
      Caption         =   "Slide Show"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "H&elp"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private ScrollArea As CScrollArea





Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
    dialog.Show
End Sub

Private Sub cmdWallpaper_Click()
    On Error GoTo Fehler
    Dim xFile As String
       
       xFile = WinPath & "Taz-Maina.bmp"
    If 1 = 2 Then
     
        PicScreen.Cls
        PicScreen.PaintPicture LoadPicture(Dir1.Path + "\" + File1.FileName), 0, 0, PicScreen.ScaleWidth, PicScreen.ScaleHeight
        'FoxAlphaBlend picScreen.HDC, picScreen.ScaleWidth - picFlomix.ScaleWidth, picScreen.ScaleHeight - picFlomix.ScaleHeight - 30, picFlomix.ScaleWidth, picFlomix.ScaleHeight, picFlomix.HDC, 0, 0, 128, 0, 1
    Else
        Set PicScreen.Picture = LoadPicture(Dir1.Path + "\" + File1.FileName)
    End If
 
    'Picture zu Image
    picPreview.PaintPicture PicScreen.Picture, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight
    picPreview.Refresh
    SavePicture PicScreen.Picture, xFile
       Label2.Caption = ""
   
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, ByVal xFile, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
 Label2.Caption = "The wallpaper is ready!!!"

Fehler:
End Sub

Private Sub Dir1_Change()
On Error GoTo err
    File1.Path = Dir1.Path
    Exit Sub
err:
    MsgBox "This Drive/Device is unavailble"
End Sub

Private Sub Drive1_Change()
On Error GoTo Err1
    Dir1.Path = Drive1.Drive
    Exit Sub
Err1:
    MsgBox "Taz-Mania is unable to read from this device, make sure that there is a disk in this device"
End Sub


Private Sub File1_Click()
On Error GoTo ErrPicLoad
    picPreview.Picture = LoadPicture(Dir1.Path + "\" + File1.FileName)
    ScrollArea.ReSizeArea
    Label2.Caption = ""
    frmShow.Picture1.Picture = LoadPicture(Form1.Dir1.Path + "\" + Form1.File1.FileName)
    frmShow.lblFileLocation.Caption = (Dir1.Path + "\" + File1.FileName)
    Exit Sub
ErrPicLoad:
    MsgBox "Taz-Mania is unable to preview this picture"
    
End Sub

Private Sub Form_Load()
    Set ScrollArea = New CScrollArea
    Set ScrollArea.VBar = VScroll1
    Set ScrollArea.HBar = HScroll1
    Set ScrollArea.InnerPicture = picPreview
    Set ScrollArea.FramePicture = Picture1
    ScrollArea.ReSizeArea
    
End Sub

Private Sub Form_Resize()
On Error GoTo error
 Picture1.Width = Form1.Width - 3400
    Picture1.Height = Form1.Height - 3000
    HScroll1.Width = Picture1.Width
    HScroll1.top = Picture1.top + Picture1.Height
    VScroll1.Height = Picture1.Height
    VScroll1.Left = Picture1.Left + Picture1.Width
    Set ScrollArea = New CScrollArea
    Set ScrollArea.VBar = VScroll1
    Set ScrollArea.HBar = HScroll1
    Set ScrollArea.InnerPicture = picPreview
    Set ScrollArea.FramePicture = Picture1
    ScrollArea.ReSizeArea
    
error:
End Sub



Private Sub Form_Terminate()
End
End Sub

Private Sub mnuFileExit_Click()
    'Exit the program
    
    End
End Sub



Private Sub mnuFileWallpaper_Click()
    On Error GoTo Fehler
    Dim xFile As String

       xFile = WinPath & "Taz-Maina.bmp"
    If 1 = 2 Then
     
        PicScreen.Cls
        PicScreen.PaintPicture LoadPicture(Dir1.Path + "\" + File1.FileName), 0, 0, PicScreen.ScaleWidth, PicScreen.ScaleHeight
        'FoxAlphaBlend picScreen.HDC, picScreen.ScaleWidth - picFlomix.ScaleWidth, picScreen.ScaleHeight - picFlomix.ScaleHeight - 30, picFlomix.ScaleWidth, picFlomix.ScaleHeight, picFlomix.HDC, 0, 0, 128, 0, 1
    Else
        Set PicScreen.Picture = LoadPicture(Dir1.Path + "\" + File1.FileName)
    End If
 
    'Picture zu Image
    picPreview.PaintPicture PicScreen.Picture, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight
    picPreview.Refresh
    SavePicture PicScreen.Picture, xFile
       Label2.Caption = ""
   
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, ByVal xFile, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
 Label2.Caption = "The wallpaper is ready!!!"

Fehler:
End Sub

Private Sub mnuHelpAbout_Click()
    'Shows "About window
    frmAbout.Show
    Form1.Enabled = False
End Sub



Private Sub mnuOptionsColor_Click()
    frmSetColor.Show
End Sub

Private Sub mnuOptionsPropSlide_Click()
    frmSlideShowProp.Show
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
        Timer1.Tag = Timer1.Tag + 1
    If Timer1.Tag >= Int(0.1 * 60) Then
        Timer1.Tag = 0
        File1.ListIndex = Int(Rnd * File1.ListCount)
End If
   End Sub
Private Sub cmdNext_Click()
    Timer1.Tag = 0
    If File1.ListCount Then File1.ListIndex = (File1.ListIndex + 1) Mod File1.ListCount
End Sub
Private Sub cmdRnd_Click()
    File1.ListIndex = Int(Rnd * File1.ListCount)
End Sub
