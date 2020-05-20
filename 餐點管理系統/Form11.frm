VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form11"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15045
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   10035
   ScaleWidth      =   15045
   WindowState     =   2  '最大化
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   8400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束"
      Height          =   975
      Left            =   13920
      Picture         =   "Form11.frx":16A0
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "母親的愛永不會逝去，這是永恆的愛。我們這一家即將推出艾菲爾鐵塔系列商品"
      BeginProperty Font 
         Name            =   "華康布丁體(P)"
         Size            =   12.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   9480
      Width           =   9015
   End
   Begin VB.Image Image4 
      Height          =   7050
      Left            =   9120
      Picture         =   "Form11.frx":24F2
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   5925
   End
   Begin VB.Image Image3 
      Height          =   3495
      Left            =   5400
      Picture         =   "Form11.frx":CB6F
      Top             =   4080
      Width           =   5025
   End
   Begin VB.Image Image2 
      Height          =   4485
      Left            =   3600
      Picture         =   "Form11.frx":1242F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6930
   End
   Begin VB.Image Image1 
      Height          =   6210
      Left            =   240
      Picture         =   "Form11.frx":1BBF7
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   5085
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "新上市產品"
      BeginProperty Font 
         Name            =   "華康墨字體(P)"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim A As Integer

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Timer1_Timer()
 'label3
If Label3.Left < Form11.ScaleHeight And A = 0 Then
A = 0
Label3.Left = Label3.Left + 50
Else
A = 1
Label3.Left = Label3.Left - 50
If Label3.Left <= 0 And Label3.Width Then
A = 0
Else
End If
End If


End Sub
