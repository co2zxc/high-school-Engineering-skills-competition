VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "低通濾波器的頻率響應"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Cls
Dim pi#
pi = 4 * Atn(1)
Dim R&, C#, F&
R = Val(InputBox("輸入電阻值R，單位為歐姆:", , "1600"))
C = Val(InputBox("輸入電容值C，單位為法拉:", , "0.000001"))
F = Val(InputBox("輸入頻率值F，單位為赫芝:", , "1000"))
Print "輸入電阻值R，單位為歐姆=" & R
Print "輸入電容值C，單位為法拉=" & C
Print "輸入頻率值F，單位為赫芝=" & F
Dim Amplitude#
Amplitude = 1 / (Sqr(1 + (2 * pi * F * R * C) ^ 2))
Dim Z#, Phase#
Z = 20 * Log(Amplitude) / Log(10)
Phase = -Atn(1 / ((1 / (2 * pi * F * C)) / R)) * 180 / pi
Print "濾波器的大小Z=" & Round(Z, 3) & "dB"
Print "濾波器的相角Phase=" & Round(Phase, 3)
End Sub

