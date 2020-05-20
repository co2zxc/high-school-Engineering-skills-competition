VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C0FFFF&
   Caption         =   "介紹"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   10005
   ScaleWidth      =   15210
   WindowState     =   2  '最大化
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6120
      Top             =   9000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束"
      Height          =   975
      Left            =   14280
      Picture         =   "Form6.frx":16A0
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   8760
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "咖啡的苦，不是我能想像的 ，但是你們可以，這也將是最美好的回憶"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   9600
      Width           =   7815
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '單線固定
      Height          =   1980
      Left            =   5760
      Picture         =   "Form6.frx":24F2
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3015
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '單線固定
      Height          =   4935
      Left            =   8760
      Picture         =   "Form6.frx":38A5
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   6420
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   3660
      Left            =   9840
      Picture         =   "Form6.frx":C124
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5175
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   2700
      Left            =   120
      Picture         =   "Form6.frx":DC8A
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   5295
   End
   Begin VB.Label Label3 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   $"Form6.frx":12870
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   3975
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   8535
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "動機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  MDIForm1.StatusBar1.Panels(4) = "系統動機"
End Sub

Private Sub Timer1_Timer()
 'label1
If Label1.Left < Form6.ScaleHeight And A = 0 Then
A = 0
Label1.Left = Label1.Left + 50
Else
A = 1
Label1.Left = Label1.Left - 50
If Label1.Left <= 0 And Label1.Width Then
A = 0
Else
End If
End If

End Sub
