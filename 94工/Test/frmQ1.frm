VERSION 5.00
Begin VB.Form frmQ1 
   AutoRedraw      =   -1  'True
   Caption         =   "同音代換加密法(Homophonic Substitution Ciphers)"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5895
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCode 
      Caption         =   "加密"
      Height          =   495
      Left            =   2340
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   3
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox txtEnCode 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   1
      Text            =   "abcabcabcabc"
      Top             =   240
      Width           =   4695
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   2880
      Picture         =   "frmQ1.frx":0000
      Top             =   2280
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   2880
      Picture         =   "frmQ1.frx":0208
      Top             =   720
      Width           =   135
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "密文"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label lblEnCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "明文"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmQ1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim enCode$, s$(10, 15) '(y,1)=A,(y,2)=MaxTimes,(y,3)=CountTimes

Private Sub cmdCode_Click()
txtCode = "accede"
Call subDefault

enCode = UCase(txtEnCode)
Dim i%, j%, tmp$, key$, Times$, MaxTimes$
For i = 1 To Len(enCode)
    tmp = Mid(enCode, i, 1)
    key = Trim(Str(Asc(tmp) - 64))
    Times = s(Val(key), 3)
    MaxTimes = s(Val(key), 2)
    txtCode = txtCode & s(Val(key), Val(Times) + 3) & " "
    Times = Trim(Str(Val(Times) + 1))
    If Val(Times) > Val(MaxTimes) Then
        Times = "1"
    End If
    s(Val(key), 3) = Times
Next i




txtEnCode = ""
End Sub

Private Sub subDefault()
Dim i%
For i = 1 To 10
    s(i, 3) = "1"
Next i

End Sub

Private Sub Form_Load()
Dim i%
For i = 1 To 10
    s(i, 1) = Chr(i + 64)
Next i
'A
s(1, 2) = "8"
s(1, 4) = "09"
s(1, 5) = "12"
s(1, 6) = "33"
s(1, 7) = "47"
s(1, 8) = "53"
s(1, 9) = "67"
s(1, 10) = "78"
s(1, 11) = "92"
'B
s(2, 2) = "2"
s(2, 4) = "48"
s(2, 5) = "81"
'C
s(3, 2) = "3"
s(3, 4) = "13"
s(3, 5) = "41"
s(3, 6) = "62"
'D
s(4, 2) = "4"
s(4, 4) = "01"
s(4, 5) = "03"
s(4, 6) = "45"
s(4, 7) = "79"
'E
s(5, 2) = "12"
s(5, 4) = "14"
s(5, 5) = "16"
s(5, 6) = "24"
s(5, 7) = "44"
s(5, 8) = "46"
s(5, 9) = "55"
s(5, 10) = "57"
s(5, 11) = "64"
s(5, 12) = "74"
s(5, 13) = "82"
s(5, 14) = "87"
s(5, 15) = "98"
'F
s(6, 2) = "2"
s(6, 4) = "10"
s(6, 5) = "31"
'G
s(7, 2) = "2"
s(7, 4) = "06"
s(7, 5) = "25"
'H
s(8, 2) = "6"
s(8, 4) = "23"
s(8, 5) = "39"
s(8, 6) = "50"
s(8, 7) = "56"
s(8, 8) = "65"
s(8, 9) = "68"
'I
s(9, 2) = "6"
s(9, 4) = "32"
s(9, 5) = "70"
s(9, 6) = "73"
s(9, 7) = "83"
s(9, 8) = "88"
s(9, 9) = "93"
'J
s(10, 2) = "1"
s(10, 4) = "15"
End Sub
