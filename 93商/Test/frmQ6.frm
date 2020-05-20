VERSION 5.00
Begin VB.Form frmQ6 
   AutoRedraw      =   -1  'True
   Caption         =   "求係數"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2438
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Cls"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1358
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   278
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim n%
Do
    n = Val(InputBox("請輸入n值(1<=n<=50，n為正整數)" & vbCrLf & "求(X+1)^n之各項細數:"))
Loop Until Int(n + 0.5) = Fix(n) And Abs(n) = n And n >= 1 And n <= 50
Print "輸入:"; Trim(str(n))
Dim ans$, t
Dim i%, j%
ans = "1,1"
For i = 2 To n
    t = Split(ans, ",")
    ans = "1"
    For j = 0 To UBound(t) - 1
        ans = ans & "," & Trim(str(Val(t(j)) + Val(t(j + 1))))
    Next j
    ans = ans & ",1"
Next i
Print "輸出:"; ans
Print
End Sub

Private Sub cmdCls_Click()
Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub
