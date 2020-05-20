VERSION 5.00
Begin VB.Form frmQ5 
   AutoRedraw      =   -1  'True
   Caption         =   "學生成績的排名次"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   5535
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Cls"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   1748
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim score%(51, 2), tmp%(51, 2), n%

Private Sub cmdStart_Click()
Dim i%, j%
Do
    n = Val(InputBox("請輸入學生人數:"))
Loop Until n > 0 And n <= 50 And Int(n + 0.5) = Fix(n) And Abs(n) = n
For i = 1 To n
    score(i, 1) = i
    tmp(i, 1) = i
    score(i, 2) = Val(InputBox("請輸入學生學號為" & i & "的成績:"))
    tmp(i, 2) = score(i, 2)
Next i
Dim t%
For i = 1 To n - 1
    For j = i To n
        If tmp(i, 2) < tmp(j, 2) Then
            t = tmp(i, 2): tmp(i, 2) = tmp(j, 2): tmp(j, 2) = t
            t = tmp(i, 1): tmp(i, 1) = tmp(j, 1): tmp(j, 1) = t
        End If
    Next j
Next i
Dim total% '累積人數
Dim s% '名次
t = tmp(1, 2) '分數
s = 1
For i = 1 To n + 1
    If tmp(i, 2) = tmp(i - 1, 2) Then
        total = total + 1
        's = s + 1
        tmp(i, 0) = s
    Else
        total = total + 1
        tmp(i, 0) = total
        s = total
        't = tmp(i, 2)
    End If
Next i
Print "學號", "程式設計", "名次"
For i = 1 To n
    Print Trim(Str(score(i, 1))), Trim(Str(score(i, 2))), Trim(Str(fncFind(score(i, 1))))
Next i
Print

End Sub

Private Function fncFind%(id%)
Dim i%
For i = 1 To n
    If tmp(i, 1) = id Then fncFind = tmp(i, 0): Exit For
Next i
End Function

Private Sub cmdCls_Click()
Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub
