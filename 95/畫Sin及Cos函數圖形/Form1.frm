VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "畫Sin及Cos函數圖形"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   6870
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清除"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdDrao 
      Caption         =   "畫出"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   4320
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "函數圖形"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   4200
      Width           =   3375
      Begin VB.OptionButton optCos 
         Caption         =   "Cos函數圖形"
         Height          =   180
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optSin 
         Caption         =   "Sin函數圖形"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3975
      Left            =   240
      ScaleHeight     =   3915
      ScaleWidth      =   6315
      TabIndex        =   5
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDrao_Click()
Dim pi$
pi = 3.14

Picture1.Line (0, 0)-(0, Picture1.ScaleHeight), vbRed
Picture1.Line (0, Picture1.ScaleHeight / 2)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 2), vbRed
Dim i#, y#
For i = -360 To 0 Step 0.5
    y = (i * pi / 180)
    If optSin Then
        Picture1.PSet (Abs(i * 2 * 8), Sin(y) * 1000 + Picture1.ScaleHeight / 2)
    Else
        Picture1.PSet (Abs(i * 2 * 8), -Cos(y) * 1000 + Picture1.ScaleHeight / 2)
    End If
Next i
End Sub

Private Sub cmdClear_Click()
Picture1.Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub
