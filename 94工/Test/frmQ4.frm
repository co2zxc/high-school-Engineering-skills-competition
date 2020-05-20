VERSION 5.00
Begin VB.Form frmQ4 
   AutoRedraw      =   -1  'True
   Caption         =   "急速穿梭"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   3900
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCount 
      Caption         =   "計算"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtShow 
      Height          =   1095
      Left            =   1223
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2700
      Width           =   1455
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1223
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmQ4.frx":0000
      Top             =   300
      Width           =   1455
   End
End
Attribute VB_Name = "frmQ4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCount_Click()
txtShow = ""
Dim data%(), s$, tmp1, tmp2, tmp3, tmp4
tmp1 = Split(txtData, vbCrLf)
Dim i%, j%
For i = 0 To 3
    tmp2 = Split(tmp1(i))
    tmp3 = Split(tmp2(1), "/")
    tmp2(1) = Val(tmp3(0) / tmp3(1))
    s = Trim(str(Int(Val(tmp2(0)) * Val(tmp2(1)) / Val(tmp2(2)) + 0.5)))
    txtShow = txtShow & s & vbCrLf
Next i
txtData = "100 3/5 0.8" & vbCrLf
txtData = txtData & "100 3/4 0.9" & vbCrLf
txtData = txtData & "80 4/7 0.8" & vbCrLf
txtData = txtData & "80 2/5 0.9" & vbCrLf
End Sub
