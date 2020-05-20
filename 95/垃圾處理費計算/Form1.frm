VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "垃圾處理費計算"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim x$
x = InputBox("請輸入水費的度數(半形空一格)電費的度數")
Dim n
n = Split(x, " ")
Dim A&, B&
A = Val(n(0))
B = Val(n(1))
Dim tmp&, answer&
tmp = (A + B) / 2
If A < 50 And B < 100 Then
    answer = tmp * 0.6
ElseIf (A < 50 And B >= 100 And B <= 200) Or (B < 100 And A >= 50 And A <= 100) Then
    answer = tmp * 0.8
ElseIf A > 100 And B > 200 Then
    answer = tmp * 1.4
ElseIf (A > 100 And B >= 100 And B <= 200) Or (B > 200 And A >= 50 And A <= 100) Then
    answer = tmp * 1.2
Else
    answer = tmp
End If
Print "輸入:" & x
Print "所需垃圾處理費用為" & answer * 5 & "元"
End Sub
