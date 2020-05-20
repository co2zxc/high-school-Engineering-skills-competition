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
   StartUpPosition =   3  '╰参箇砞
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   600
      Top             =   2160
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d, e As Byte

Private Sub Form_Load()
 e = 0
End Sub

Private Sub Timer2_Timer()
  If e = 0 Then e = 1 Else e = 0
  a = Val(Mid(Time, 4, 1))  ': 计
  b = Val(Mid(Time, 5, 1))  ': 计
  c = Val(Mid(Time, 7, 1))  'だ牧: 计
  d = Val(Mid(Time, 8, 1))  'だ牧: 计
 
  If OpenUsbDevice(VendorID, ProductID) Then Label1.Caption = "Connect" Else Label1.Caption = "DisConnect"
  OutDataEightByte &H21, a, b, c, d, e, 0, 0
End Sub

