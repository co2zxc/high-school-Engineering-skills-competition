VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "�C�q�o�i�����W�v�T��"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '�ù�����
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
R = Val(InputBox("��J�q����R�A��쬰�کi:", , "1600"))
C = Val(InputBox("��J�q�e��C�A��쬰�k��:", , "0.000001"))
F = Val(InputBox("��J�W�v��F�A��쬰����:", , "1000"))
Print "��J�q����R�A��쬰�کi=" & R
Print "��J�q�e��C�A��쬰�k��=" & C
Print "��J�W�v��F�A��쬰����=" & F
Dim Amplitude#
Amplitude = 1 / (Sqr(1 + (2 * pi * F * R * C) ^ 2))
Dim Z#, Phase#
Z = 20 * Log(Amplitude) / Log(10)
Phase = -Atn(1 / ((1 / (2 * pi * F * C)) / R)) * 180 / pi
Print "�o�i�����j�pZ=" & Round(Z, 3) & "dB"
Print "�o�i�����ۨ�Phase=" & Round(Phase, 3)
End Sub

