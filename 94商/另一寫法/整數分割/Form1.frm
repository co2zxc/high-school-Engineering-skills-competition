VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "整數分割"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5595
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox Textbox2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   1440
      Width           =   4815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "作答鈕"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox textbox1 
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
      Left            =   2400
      TabIndex        =   1
      Text            =   "textbox1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
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
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "輸入一整數N:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim err_count%
Dim unable_restart As Boolean

Private Sub cmdOk_Click()
Textbox2.ForeColor = RGB(0, 0, 0)
If unable_restart Then
    If textbox1 = "***" Then
        unable_restart = False
    Else
        Textbox2.Alignment = 0
        Textbox2.ForeColor = RGB("&hFF", 0, 0)
        Textbox2 = "???"
    End If
    Exit Sub
End If

If Not (Val(textbox1) >= 1 And Val(textbox1) <= 10) And Val(textbox1) <> Fix(Val(textbox1)) Then
    lbl = "輸入錯誤"
    err_count = err_count + 1
    If err_count >= 3 Then lbl = "輸入超過3次"
Else
    err_count = 0
    lbl = ""
End If
Textbox2.Alignment = 2
Dim n%
n = Val(textbox1)
Dim i%, j%, tmp$
tmp = ""
For i = n To 1 Step -1
    tmp = tmp & i
    For j = 1 To (-i + n)
      tmp = tmp & 1
    Next j
    tmp = tmp & vbCrLf
Next i
Textbox2 = tmp
unable_restart = True
End Sub
