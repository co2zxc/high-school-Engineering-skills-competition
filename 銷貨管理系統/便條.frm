VERSION 5.00
Begin VB.Form frmDrawLine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "便條"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   DrawWidth       =   4
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6375
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmd2 
      Caption         =   "清除"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "回主畫面"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
End
Attribute VB_Name = "frmDrawLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd2_Click()
Set Picture = Nothing
End Sub

Private Sub cmd1_Click()
 Unload Me
    frmMDIMain.Show
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  CurrentX = X
  CurrentY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Button
   Case Is = vbLeftButton
      Line -(X, Y), vbBlack
   Case Is = vbRightButton
      Line -(X, Y), vbRed
   Case Is = vbMiddleButton
      Line -(X, Y), BackColor
   End Select
End Sub
