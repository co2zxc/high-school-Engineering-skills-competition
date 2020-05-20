VERSION 5.00
Begin VB.Form frmQ3 
   AutoRedraw      =   -1  'True
   Caption         =   "�H�Ѱ����k"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   5535
   StartUpPosition =   2  '�ù�����
End
Attribute VB_Name = "frmQ3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Horse$(1) '0:x,1:y
Dim Table(8, 8) As Boolean 'T:Block,(x,y)
Dim Flag As Boolean
'T:��J���ثe��m�P�@�ǻ�ê��
'F:��J���ʼƦr��A�B�����ʦܷs�y��
Dim Buffer$ '��J���ثe��m�P�@�ǻ�ê��
Dim TheKey$

Private Sub Form_Load()
Flag = True '��J���ثe��m�P�@�ǻ�ê��
Buffer = ""
End Sub

Private Sub Form_Activate()
If Flag = True Then Print "���ثe��m�P�@�ǻ�ê��:";
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
TheKey = Chr(KeyAscii)
If TheKey = "9" Then Print TheKey; "(�������{��)": End
If Flag Then
    If KeyAscii <> 13 Then
        Print TheKey; '��X����
        Buffer = Buffer & TheKey
    Else
        Call subDefault '���ثe��m�P�@�ǻ�ê��]
        Flag = False
        'Buffer = ""
        Print
        Print "��J���ʼƦr��:";
    End If
Else
    If TheKey > "7" Or TheKey < "0" Then Exit Sub
    Print TheKey
    Print "�����ʦܷs��m:";
    If fncOutRange Then
        Print Horse(0); " "; Horse(1);
        Print "(�]�W�X�ѽL�~�ӫO����y��)";
    ElseIf fncBlocked Then
         Print Horse(0); " "; Horse(1);
        Print "(�]���}����ӫO����y��)";
    Else
        Call subMove
        Print Horse(0); " "; Horse(1);
    End If
    Print
    Print "��J���ʼƦr��:";
End If
End Sub

Private Sub subDefault()
Dim tmp
tmp = Split(Buffer)
Horse(0) = tmp(0)
Horse(1) = tmp(1)
Dim i%
For i = 2 To UBound(tmp) Step 2
    Table(Val(tmp(i)), Val(tmp(i + 1))) = True
Next i
End Sub

Private Function fncOutRange() As Boolean
Select Case TheKey
Case "0"
    If (Val(Horse(0)) + 1 > 8) Or (Val(Horse(1)) + 2 > 8) Then fncOutRange = True
Case "1"
    If (Val(Horse(0)) - 1 < 1) Or (Val(Horse(1)) + 2 > 8) Then fncOutRange = True
Case "2"
    If (Val(Horse(0)) - 2 < 1) Or (Val(Horse(1)) + 1 > 8) Then fncOutRange = True
Case "3"
    If (Val(Horse(0)) - 2 < 1) Or (Val(Horse(1)) - 1 < 1) Then fncOutRange = True
Case "4"
    If (Val(Horse(0)) - 1 < 1) Or (Val(Horse(1)) - 2 < 1) Then fncOutRange = True
Case "5"
    If (Val(Horse(0)) + 1 > 8) Or (Val(Horse(1)) - 2 < 1) Then fncOutRange = True
Case "6"
    If (Val(Horse(0)) + 2 > 8) Or (Val(Horse(1)) - 1 < 1) Then fncOutRange = True
Case "7"
    If (Val(Horse(0)) + 2 > 8) Or (Val(Horse(1)) + 1 > 8) Then fncOutRange = True
End Select
End Function

Private Function fncBlocked() As Boolean
Select Case TheKey
Case "0", "1"
    If Table(Horse(0), Horse(1) + 1) Then fncBlocked = True
Case "2", "3"
    If Table(Horse(0) - 1, Horse(1)) Then fncBlocked = True
Case "4", "5"
    If Table(Horse(0), Horse(1) - 1) Then fncBlocked = True
Case "6", "7"
    If Table(Horse(0) + 1, Horse(1)) Then fncBlocked = True
End Select
End Function

Private Sub subMove()
Select Case TheKey
Case "0"
    Horse(0) = Val(Horse(0)) + 1
    Horse(1) = Val(Horse(1)) + 2
Case "1"
    Horse(0) = Val(Horse(0)) - 1
    Horse(1) = Val(Horse(1)) + 2
Case "2"
    Horse(0) = Val(Horse(0)) - 2
    Horse(1) = Val(Horse(1)) + 1
Case "3"
    Horse(0) = Val(Horse(0)) - 2
    Horse(1) = Val(Horse(1)) - 1
Case "4"
    Horse(0) = Val(Horse(0)) - 1
    Horse(1) = Val(Horse(1)) - 2
Case "5"
    Horse(0) = Val(Horse(0)) + 1
    Horse(1) = Val(Horse(1)) - 2
Case "6"
    Horse(0) = Val(Horse(0)) + 2
    Horse(1) = Val(Horse(1)) - 1
Case "7"
    Horse(0) = Val(Horse(0)) + 2
    Horse(1) = Val(Horse(1)) + 1
End Select
End Sub
