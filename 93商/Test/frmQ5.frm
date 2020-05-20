VERSION 5.00
Begin VB.Form frmQ5 
   AutoRedraw      =   -1  'True
   Caption         =   "字串重組"
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
Attribute VB_Name = "frmQ5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim tmp
tmp = InputBox("請輸入字串")
Print "輸入:"; tmp
tmp = Split(tmp)
Dim s1$, s2$
s1 = tmp(0): s2 = tmp(1)
Dim ans$
ans = ans & Left(s1, 3)
Dim i%
Erase tmp
ReDim tmp(7)
For i = 0 To 2
    tmp(i) = Mid(ans, i + 1, 1)
Next i
Dim flag As Boolean
Dim t$, j%, k%, l%
j = 1
For i = 1 To 5
    l = 0
    Do
        flag = True
        t = Mid(s2, j + l, 1)
        For k = 0 To j + 1
            If t = tmp(k) Then flag = False: Exit For
        Next k
        l = l + 1
    Loop Until flag
    tmp(j + 2) = t
    j = j + 1
    ans = ans & t
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
