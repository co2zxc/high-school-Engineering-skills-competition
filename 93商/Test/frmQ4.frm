VERSION 5.00
Begin VB.Form frmQ4 
   AutoRedraw      =   -1  'True
   Caption         =   "�έp"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdCls 
      Caption         =   "Cls"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCount_Click()
Dim n%, data!(), avg#, zero#, middin#, many#
Dim i%, tmp
ReDim data(0)
i = 0
Do
    Do
        tmp = InputBox("��J(0��ܵ���)", "��" & (i + 1) & "�ӼƦr")
        If tmp = 0 Then Exit Do
    Loop Until tmp >= 1 And tmp <= 100
    If tmp = 0 Then Exit Do
    If i <> 0 Then ReDim Preserve data(i + 1)
    data(i) = tmp
    i = i + 1
Loop
n = i
Print "��J(0��ܵ���)�G";
For i = 0 To n - 1
    Print Trim(Str(data(i))) & ",";
Next i
Print "0"

Print "��X�G"
For i = 0 To n - 1
    avg = avg + data(i)
Next i
avg = avg / n
Print "  �����ơG" & avg & IIf(Fix(avg) = avg, ".0", "")

For i = 0 To n - 1
    zero = zero + (data(i) - avg) ^ 2
Next i
zero = zero / (n - 1)
Print "  �ܲ��ơG" & Fix(zero * 100) / 100

Dim j%
For i = 0 To n - 2
    For j = i To n - 1
        If data(i) > data(j) Then tmp = data(i): data(i) = data(j): data(j) = tmp
    Next j
Next i
tmp = n
If n Mod 2 <> 0 Then
    middin = data((n + 1) / 2)
Else
    middin = (data(n / 2 - 1) + data(n / 2 + 1 - 1)) / 2
End If
Print "  ����ơG" & middin

Dim times%, ttimes%
ttimes = 1
For i = 1 To n - 1
    If data(i) = data(i - 1) Then
        ttimes = ttimes + 1
    Else
        If ttimes > times Then
            times = ttimes
            many = data(i - 1)
            ttimes = 1
        End If
    End If
Next i
Print "  ���@�ơG" & many
End Sub

Private Sub cmdCls_Click()
Cls
End Sub
