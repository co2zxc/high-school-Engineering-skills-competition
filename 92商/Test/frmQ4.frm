VERSION 5.00
Begin VB.Form frmQ4 
   AutoRedraw      =   -1  'True
   Caption         =   "數列"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   6900
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Cls"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ary&()

Private Sub cmdStart_Click()
ReDim ary(1)
Dim n&, Flag As Boolean
Do
    n = Val(InputBox("請輸入大於1的正整數"))
Loop Until Abs(n) = n And Int(n + 0.5) = Fix(n) And n > 0
Dim tmp$
Print Trim(Str(n)),
ary(0) = n
n = n ^ 2
Print Trim(Str(n))
ary(1) = n
tmp = Trim(Str(n))
If Len(tmp) >= 2 Then n = fncMin(n) * 10 + fncMax(n)
Do
    tmp = Trim(Str(n))
    n = n ^ 2
    Print tmp, Trim(Str(n))
    If Len(tmp) >= 2 Then n = fncMin(n) * 10 + fncMax(n)
    If Not fncChk(n) Then
        ReDim Preserve ary(UBound(ary) + 1) 'If ary(UBound(ary)) <> 0 Then
        ary(UBound(ary)) = n
    Else
        Print "*" & Trim(Str(n)), Trim(Str(n ^ 2))
        Flag = True
    End If
Loop Until Flag
Print
End Sub

Private Function fncMax&(num&)
Dim i%, tmp$
For i = 0 To Len(Trim(Str(num))) - 1
    tmp = Trim(Str(num))
    If Val(Mid(tmp, i + 1, 1)) > fncMax Then fncMax = Val(Mid(tmp, i + 1, 1))
Next i
End Function

Private Function fncMin&(num&)
Dim i%, tmp$
fncMin = 99999
For i = 0 To Len(Trim(Str(num))) - 1
    tmp = Trim(Str(num))
    If Val(Mid(tmp, i + 1, 1)) < fncMin Then fncMin = Val(Mid(tmp, i + 1, 1))
Next i
End Function

Private Function fncChk(num&) As Boolean
Dim i%
For i = 0 To UBound(ary)
    If num = ary(i) Then fncChk = True: Exit For
Next i
End Function

Private Sub cmdCls_Click()
Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub
