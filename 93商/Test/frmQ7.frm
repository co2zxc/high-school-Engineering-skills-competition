VERSION 5.00
Begin VB.Form frmQ7 
   AutoRedraw      =   -1  'True
   Caption         =   "9*9 ­¼ªkªí"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   308
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Cls"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   1388
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   2468
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim i%, j%, t%
j = 1
For i = 1 To 9
    t = 4 - Abs(5 - i)
    For j = 1 To 9
        If j > (5 - t) And j < (5 + t) Then
            Me.ForeColor = Me.BackColor
        Else
            Me.ForeColor = vbBlack
        End If
        Print Tab((j - 1) * 5 + IIf(Len(fncCstr(i * j)) = 1, 1, 0)); fncCstr(i * j);
    Next j
    Print
Next i
End Sub

Private Function fncCstr$(value&)
fncCstr = Trim(Str(value))
End Function

Private Sub cmdCls_Click()
Cls
End Sub

Private Sub cmdEnd_Click()
End
End Sub
