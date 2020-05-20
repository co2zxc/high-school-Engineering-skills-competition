VERSION 5.00
Begin VB.Form frmQ2 
   AutoRedraw      =   -1  'True
   Caption         =   "質因數分解"
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
Attribute VB_Name = "frmQ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim n%, ans$
Do
    n = Val(InputBox("請輸入一正整數"))
Loop Until Abs(n) = n And Int(n + 0.5) = Fix(n) And n >= 3 And n <= 1000
Print "輸入:"; Trim(Str(n))
Dim i%, tmp%, j%, t%
tmp = n

t = 2
Do
    't = 0
    Do
        't = 0
        'i = i + 1
        For j = t To n
            If n Mod j = 0 Then Exit Do
        Next j
    Loop
    t = j ' - 1
    If n Mod t = 0 Then
        n = n / t
        ans = ans & Trim(Str(t)) & "*"
    End If
Loop Until n = 1

Print "輸出:"; Trim(Str(tmp)); "="; Left(ans, Len(ans) - 1)
Print
End Sub

Private Sub cmdCls_Click()
Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub
