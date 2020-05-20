VERSION 5.00
Begin VB.Form frmQ3 
   AutoRedraw      =   -1  'True
   Caption         =   "聯立方程式"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   578
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Cls"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1358
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2138
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim tmp
tmp = InputBox("請輸入參數")
Print "輸入:"; tmp
tmp = Split(tmp, ",")
Dim a1%, b1%, c1%, a2%, b2%, c2%
a1 = tmp(0): b1 = tmp(1): c1 = tmp(2)
a2 = tmp(3): b2 = tmp(4): c2 = tmp(5)
Dim delta&, deltaX&, deltaY&
delta = a1 * b2 - b1 * a2
deltaX = -c1 * b2 + b1 * c2
deltaY = -a1 * c2 + a2 * c1
Dim x#, y#
Print "輸出:";
If delta <> 0 Then
    x = deltaX / delta
    y = deltaY / delta
    Print "X="; x; "Y="; y
ElseIf delta = 0 And deltaX = 0 And deltaY = 0 Then
    Print "無限多解"
Else 'delta=0,deltaX<>0,deltaY<>0
    Print "無解"
End If
Print
End Sub

Private Sub cmdCls_Click()
Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub
