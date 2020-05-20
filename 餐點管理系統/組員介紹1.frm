VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form8 
   BackColor       =   &H00C0FFFF&
   Caption         =   "介紹"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14775
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   10170
   ScaleWidth      =   14775
   WindowState     =   2  '最大化
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "我的介紹"
      BeginProperty Font 
         Name            =   "華康墨字體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   14415
      Begin VB.CommandButton Command9 
         Caption         =   "結束"
         Height          =   975
         Left            =   13200
         Picture         =   "組員介紹1.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   8520
         Width           =   975
      End
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   735
         Left            =   6120
         TabIndex        =   3
         Top             =   8640
         Visible         =   0   'False
         Width           =   2535
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   5
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   4471
         _cy             =   1296
      End
      Begin VB.Image Image5 
         Height          =   3735
         Left            =   4680
         Picture         =   "組員介紹1.frx":0E52
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   5295
      End
      Begin VB.Image Image4 
         Height          =   1890
         Left            =   840
         Picture         =   "組員介紹1.frx":109EC
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   2760
      End
      Begin VB.Image Image3 
         Height          =   3735
         Left            =   5040
         Picture         =   "組員介紹1.frx":14297
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3495
      End
      Begin VB.Image Image2 
         Height          =   3840
         Left            =   240
         Picture         =   "組員介紹1.frx":19567
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   5355
      End
      Begin VB.Image Image1 
         Height          =   9360
         Left            =   9000
         Picture         =   "組員介紹1.frx":2732D
         Top             =   240
         Width           =   5400
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   $"組員介紹1.frx":356BE
         ForeColor       =   &H00000000&
         Height          =   3495
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   5055
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command9_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  MDIForm1.StatusBar1.Panels(5) = " 亞瑟小子_YEAH "
  MDIForm1.StatusBar1.Panels(4) = "製作者介紹"
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Timer1_Timer()
 'label1
If Label2.Left < Form8.ScaleHeight And A = 0 Then
A = 0
Label2.Left = Label2.Left + 50
Else
A = 1
Label2.Left = Label2.Left - 50
If Label2.Left <= 0 And Label2.Width Then
A = 0
Else
End If
End If
End Sub
Private Sub Form_Activate()

WindowsMediaPlayer1.Controls.play
WindowsMediaPlayer1.URL = "手機鈴聲 - 亞瑟小子_YEAH 1.mp3"
End Sub
