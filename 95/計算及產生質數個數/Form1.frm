VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "�p��β��ͽ�ƭӼ�"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '�ù�����
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim n!, i%, t$
ReDim ary%(0)
Do
    n = Val(InputBox("�п�J>=2�������:"))
Loop Until n >= 2 And Fix(n) = n
ReDim ary%(n)
For i = 2 To n
    ary(i) = i
Next i
'��***�R��z�k***��
Dim flag As Boolean '���ƬO�_�����
Dim j%
For i = 2 To n
    flag = False
    For j = 2 To i
        If i Mod j = 0 And j <> i Then flag = flag + 1
    Next j
    If Not flag Then
        For j = i To n
            If j <> i And j Mod i = 0 Then ary(j) = 0
        Next j
    End If
Next i
'��***�R��z�k***��
Dim tmp%
For i = 2 To n - 1
    For j = i To n
        If ary(i) < ary(j) Then tmp = ary(i): ary(i) = ary(j): ary(j) = tmp
    Next j
Next i
Dim answer$, num%
If n >= 5 Then
    For i = 4 To 2 Step -1
        answer = answer & ary(i) & ","
        num = num + 1
    Next i
    answer = Left(answer, Len(answer) - 1)
ElseIf n = 2 Then
    answer = "2"
    num = 1
Else
    answer = "2,3"
    num = 2
End If
Print "��ƭӼƦ�" & num & "�ӡA�̤j��3�ӭȼ�" & answer
End Sub
