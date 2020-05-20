VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00C0FFFF&
   Caption         =   "介紹"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   11130
   WindowState     =   2  '最大化
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "豬的介紹"
      BeginProperty Font 
         Name            =   "華康墨字體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13935
      Begin VB.CommandButton Command9 
         BackColor       =   &H008080FF&
         Caption         =   "結束"
         Height          =   975
         Left            =   12960
         Picture         =   "組員介紹2.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   8400
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   4125
         Left            =   7680
         Picture         =   "組員介紹2.frx":0E52
         Top             =   360
         Width           =   6000
      End
      Begin VB.Image Image1 
         Height          =   3975
         Left            =   480
         Picture         =   "組員介紹2.frx":5D22
         Top             =   5040
         Width           =   6000
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  '單線固定
         Caption         =   $"組員介紹2.frx":53794
         Height          =   3255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   4455
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  '單線固定
         Height          =   6765
         Left            =   4200
         Picture         =   "組員介紹2.frx":53BFA
         Top             =   2640
         Width           =   9660
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command9_Click()
Unload Me
End Sub

Private Sub Form_Load()
  MDIForm1.StatusBar1.Panels(4) = "製作者介紹"
End Sub
