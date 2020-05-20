VERSION 5.00
Begin VB.Form frmQ2 
   AutoRedraw      =   -1  'True
   Caption         =   "二進位「直式乘法」"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7470
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Cls"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim n1$, n2$

Private Sub cmdStart_Click()
Dim tmp
tmp = InputBox("請輸入值:")
tmp = Split(tmp)
n1 = tmp(0): n2 = tmp(2)
Erase tmp
Print Spc(10 - Len(n1) + 1); n1
Print Spc(10 - Len(n2) + 1); n2
Print String(20, "-")
Dim i%, save$(100)
For i = 1 To Len(n2)
    tmp = Val(Mid(n2, Len(n2) - i + 1, 1))
    'save(i) = Trim(str(fncCval(n1) And tmp))
    If tmp = 0 Then
        save(i) = 0
    Else
        save(i) = n1
    End If
    save(i) = String(Len(n1) - Len(save(i)), "0") & save(i)
    Print Spc(10 - Len(save(i)) - i + 1 + 1); save(i)
Next i
Print String(20, "-")
Dim sum&
For i = 1 To 2 * Len(n2)
    sum = sum + Val(save(i)) * 10 ^ (i - 1)
Next i

For i = 1 To Len(Trim(str(sum)))
    tmp = Val(Mid(sum, Len(Trim(str(sum))) - i + 1, 1))
    If tmp > 1 Then
        sum = sum - 2 * 10 ^ (i - 1)
        sum = sum + 10 ^ i
    End If
Next i
Print Spc(10 - 2 * Len(n2) + 1); Format(sum, String(2 * Len(n2), "0"))
Print
End Sub

Private Function fncCval&(str$)
Dim i%, tmp%
For i = 1 To Len(str)
    tmp = Val(Mid(str, i, 1))
    fncCval = fncCval + tmp * 2 ^ (Len(str) - i)
Next i
End Function

Private Sub cmdCls_Click()
Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub

