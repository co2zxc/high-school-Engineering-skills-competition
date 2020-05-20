VERSION 5.00
Begin VB.Form frmQ6 
   AutoRedraw      =   -1  'True
   Caption         =   "請寫一個程式作兩個矩陣相乘"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5025
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtOut 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   3
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox txtIn 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   2
      Text            =   "frmQ6.frx":0000
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2033
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   833
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim A%(), B%()

Private Sub cmdStart_Click()
Dim tmp, data
data = Split(txtIn, vbCrLf)
tmp = Split(data(1))
ReDim A(2, UBound(tmp) + 1), B(UBound(tmp) + 1, 2)
Dim i%, j%
For i = 1 To 2
    tmp = Split(data(i - 1))
    For j = 0 To UBound(tmp)
        A(i, j + 1) = tmp(j)
    Next j
Next i
For i = 1 To UBound(tmp) + 1
    tmp = Split(data(i - 1 + 3))
    For j = 0 To 1
        B(i, j + 1) = tmp(j)
    Next j
Next i
Dim c#(2, 2), k%
For i = 1 To 2
    For j = 1 To 2
        For k = 1 To UBound(A, 2)
            c(i, j) = c(i, j) + A(i, k) * B(k, j)
        Next k
    Next j
Next i
Dim t$
For i = 1 To 2
    For j = 1 To 2
        t = t & " " & c(i, j)
    Next j
    t = t & vbCrLf
Next i
txtOut = t
End Sub

Private Sub cmdEnd_Click()
End
End Sub
