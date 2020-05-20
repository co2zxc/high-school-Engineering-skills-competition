VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   9570
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command7 
      Caption         =   "輸入角度"
      Height          =   975
      Left            =   7920
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CSC"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SEC"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "COT"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TAN"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "COS"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   360
      ScaleHeight     =   2835
      ScaleWidth      =   7395
      TabIndex        =   1
      Top             =   240
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIN"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() ' sin
Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
  Picture1.Scale (-10, 10)-(10, -10)
  Picture1.DrawWidth = 1  '線條寬度
  Picture1.Line (-10, 0)-(10, 0)  '畫X軸
  Picture1.Line (0, -10)-(0, 10)  '畫Y軸
  pi = 3.141593
  For x = -2 * pi To 2 * pi Step 0.01  '-2π<=X<=2π
    Y = 3 * Sin(x)
    Picture1.PSet (x, Y)  '畫點
  Next x
End Sub


Private Sub Command2_Click()  ' cos
Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Scale (-10, 10)-(10, -10)
  Picture1.DrawWidth = 1
  Picture1.Line (-10, 0)-(10, 0)
  Picture1.Line (0, -10)-(0, 10)
  pi = 3.141593
  For x = -2 * pi To 2 * pi Step 0.01
    Y = 3 * Cos(x)
    Picture1.PSet (x, Y)
  Next x
End Sub

Private Sub Command3_Click() ' tan
Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Scale (-10, 10)-(10, -10)
  Picture1.DrawWidth = 1
  Picture1.Line (-10, 0)-(10, 0)
  Picture1.Line (0, -10)-(0, 10)
  pi = 3.141593
  For x = -2 * pi To 2 * pi Step 0.01
    Y = 3 * Tan(x)
    Picture1.PSet (x, Y)
  Next x
End Sub

Private Sub Command4_Click() ' cot
Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Scale (-10, 10)-(10, -10)
  Picture1.DrawWidth = 1
  Picture1.Line (-10, 0)-(10, 0)
  Picture1.Line (0, -10)-(0, 10)
  pi = 3.141593
  For x = -2 * pi To 2 * pi Step 0.01
    Y = 3 * 1 / Tan(x)
    Picture1.PSet (x, Y)
  Next x
End Sub

Private Sub Command5_Click() ' sce
Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Scale (-10, 10)-(10, -10)
  Picture1.DrawWidth = 1
  Picture1.Line (-10, 0)-(10, 0)
  Picture1.Line (0, -10)-(0, 10)
  pi = 3.141593
  For x = -2 * pi To 2 * pi Step 0.01
    Y = 3 * 1 / Cos(x)
    Picture1.PSet (x, Y)
  Next x
End Sub

Private Sub Command6_Click() ' csc
Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Scale (-10, 10)-(10, -10)
  Picture1.DrawWidth = 1
  Picture1.Line (-10, 0)-(10, 0)
  Picture1.Line (0, -10)-(0, 10)
  pi = 3.141593
  For x = -2 * pi To 2 * pi Step 0.01
    Y = 3 * 1 / Sin(x)
    Picture1.PSet (x, Y)
  Next x
End Sub



Private Sub Command7_Click()
Dim a As Double
If KeyAscii = 8 Then Exit Sub
Picture1.ForeColor = RGB(255, 0, 0)
pi = 3.141593
a = InputBox("請輸入角度", "角度", 0)
b = a * pi / 180
Picture1.Line (b, 10)-(b, -10)




End Sub
