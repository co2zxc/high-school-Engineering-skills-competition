VERSION 5.00
Begin VB.Form frmQ7 
   AutoRedraw      =   -1  'True
   Caption         =   "十進位轉二進位"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   5865
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Cls"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim n!
Do
    n = Val(InputBox("請輸入十進位之1-100的正整數"))
Loop Until Int(n + 0.5) = Fix(n) And n = Abs(n) And n > 0
Dim ans$, tmp!, m%
tmp = n
Do
    ans = (n Mod 2) & ans
    n = n \ 2
Loop Until n = 0
ans = String(8 - Len(ans), "0") & ans
Print ans
End Sub

Private Sub cmdCls_Click()
Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub
