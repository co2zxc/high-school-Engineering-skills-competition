VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "急速穿梭"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   3600
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Text            =   "80 1/5 0.9"
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Text            =   "80 1/5 0.8"
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Text            =   "100 1/5 0.9"
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算極速"
      Height          =   375
      Left            =   1193
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "100 1/5 0.8"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   3480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "結果:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   540
   End
   Begin VB.Label lbl 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
lbl = ""
Dim i%, j%, tmp, ok#(2)
Dim a%, b%, flag As Boolean
For i = 0 To 3
    tmp = Split(txt(i), " ")
    ok(0) = Val(tmp(0))
    flag = False
    a = 0
    b = 0
    For j = 1 To Len(tmp(1))
        If Mid(tmp(1), j, 1) = "/" Then flag = True: j = j + 1
        If flag Then
            b = b * 10 + Val(Mid(tmp(1), j, 1))
        Else
            a = a * 10 + Val(Mid(tmp(1), j, 1))
        End If
    Next j
    'If b = 0 Then b = 1
    ok(1) = a / b
    ok(2) = Val(tmp(2))
    If i > 0 Then lbl = lbl & vbCrLf
    lbl = lbl & Trim(Str(Int(ok(0) * ok(1) / ok(2) + 0.5)))
Next i
End Sub
