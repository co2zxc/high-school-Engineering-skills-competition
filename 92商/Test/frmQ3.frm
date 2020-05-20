VERSION 5.00
Begin VB.Form frmQ3 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6225
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   548
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Cls"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   1808
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enCode$, Code$, strLen%, spaceLen

Private Sub cmdStart_Click()
enCode = "": Code = "": strLen = 0: spaceLen = 0
enCode = UCase(InputBox("請輸入明碼文字(大小寫英文字母"))
strLen = Len(enCode)
Code = Code & "**"
Dim i%, tmp$
For i = 1 To Len(enCode)
    If Mid(enCode, i, 1) = " " Then spaceLen = spaceLen + 1
Next i
strLen = strLen - spaceLen
For i = 1 To Len(enCode)
    tmp = Mid(enCode, i, 1)
    Select Case tmp
    Case " ": Code = Code & "*": spaceLen = spaceLen + 1
    Case Else: Code = Code & fncCode1(tmp)
    End Select
Next i
Code = Code & "#"
Code = Code & fncCode2(strLen)
Code = Code & "***"
Print enCode
Print Code
Print
End Sub

Private Function fncCode1$(str$)
Dim s$
s = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & "ABCDEFGHI"
Dim t%
t = InStr(s, str) + 64 + (strLen Mod 10)
fncCode1 = fncCode2(t)

End Function

Private Function fncCode2$(num%)
Dim s$
s = "bcdefghija"
Dim i%, tmp$, t%
tmp = Trim(str(num))
For i = 1 To Len(tmp)
    t = val(Mid(tmp, i, 1))
    If t = 0 Then t = 10
    fncCode2 = fncCode2 & Mid(s, t, 1)
Next i
End Function

Private Sub cmdCls_Click()
Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub

