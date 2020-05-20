VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "訊號調變與解調變"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7500
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "輸出"
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      DrawStyle       =   1  '破折線
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   6915
      TabIndex        =   8
      Top             =   2280
      Width           =   6975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Text            =   "5"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Text            =   "10"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "1"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Text            =   "0.001"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "振福大小:"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "載波頻率之訊號大小:"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "輸入訊號之頻率大小:"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "取樣時間:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
t = Text1.Text
f1 = Text2.Text
f0 = Text3.Text
A = Text4.Text
pi = 3.14159265358979
c1 = 2 * pi * f1
c0 = 2 * pi * f0

Picture1.DrawStyle = 0
Picture1.Scale (-0.5, 10)-(2.5, -10)
Picture1.Cls
Picture1.Line (-0.1, -5)-(2.1, -5)
Picture1.Line (0, -6)-(0, 6)
For i = 0 To 2 Step 0.0001
If A <= 0 Then
y = 0
Else
y = (A / 2 * Cos((c0 + c1) * i) + A / 2 * Cos((c0 - c1) * i)) * 5 / A
End If
Picture1.PSet (i, y)
Next

For i = 0 To 2 Step 0.2
Picture1.CurrentX = i - 0.1
Picture1.CurrentY = -6
Picture1.Print i
Next

Picture1.DrawStyle = 2
Picture1.Line (-0, 0)-(2, 0)

If A <= 0 Then
Else
Picture1.CurrentX = -0.1
Picture1.CurrentY = A * 5 / A
Picture1.Print A
End If

Picture1.CurrentX = -0.1
Picture1.CurrentY = 0
Picture1.Print 0

If A <= 0 Then
Else
Picture1.CurrentX = -0.1
Picture1.CurrentY = -A * 5 / A
Picture1.Print -A
End If

End Sub

