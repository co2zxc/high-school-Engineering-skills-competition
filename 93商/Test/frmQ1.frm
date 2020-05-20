VERSION 5.00
Begin VB.Form frmQ1 
   AutoRedraw      =   -1  'True
   Caption         =   "加密法"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   2138
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Cls"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1358
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   578
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim enCode$, Code$
enCode = InputBox("請輸入欲加密的字串")
Print "輸入:"; enCode
Dim i%
For i = 1 To Len(enCode)
    Code = Code & fncC(Asc(Mid(enCode, i, 1)))
Next i
Print "輸出:"; Code
Print
End Sub

Private Function fncC$(ascCode%)
Dim tmp1$, tmp2$
tmp1 = fnc2(ascCode)
tmp2 = Val(Right(tmp1, 3))
tmp2 = Format(Trim(Str(Val(tmp2) Xor Val(111))), "000")
tmp1 = Left(tmp1, Len(tmp1) - 3) & tmp2
fncC = Chr(fnc10(tmp1))
End Function

Private Function fnc2$(ByVal num&)
Do
    fnc2 = Trim(Str(num Mod 2)) & fnc2
    num = num \ 2
Loop Until num = 0
End Function

Private Function fnc10&(ByVal num$)
Dim i%, tmp$
For i = 1 To Len(num)
    tmp = Mid(num, Len(num) - i + 1, 1)
    fnc10 = fnc10 + Val(tmp) * 2 ^ (i - 1)
Next i
End Function

Private Sub cmdCls_Click()
Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub
