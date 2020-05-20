VERSION 5.00
Begin VB.Form frmQ5 
   AutoRedraw      =   -1  'True
   Caption         =   "整數分割"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtShow 
      Height          =   2415
      Left            =   113
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   3
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "作答鈕"
      Height          =   375
      Left            =   3473
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtN 
      Alignment       =   1  '靠右對齊
      Height          =   375
      Left            =   1898
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblN 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "輸入一整數N:"
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
      Left            =   113
      TabIndex        =   0
      Top             =   120
      Width           =   1425
   End
End
Attribute VB_Name = "frmQ5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
txtShow = ""
Dim n%
n = Val(txtN)
Dim i%, j%, t$
t = ""
For i = 1 To n
    For j = 1 To n - i
        t = t & " "
    Next j
    t = t & Trim(Str(n - i + 1))
    If i <> 1 Then
        For j = 1 To i - 1
            t = t & " 1"
        Next j
    End If
    t = t & vbCrLf
Next i
txtShow = t
txtN = ""
End Sub
