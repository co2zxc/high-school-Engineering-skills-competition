VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click!"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N$, strmod$, x#

Private Sub cmdClick_Click()
N = InputBox("請輸入想轉換的字串", "請輸入值")
x = (Asc(N)): Call sub0
Print "2進制:"; x
Print "16進制："; Hex(Asc(N))
End Sub

Public Sub sub0()
Do
    strmod = strmod & (x Mod 2)
    x = x \ 2
Loop Until x = 0
x = Val(strmod)
End Sub
