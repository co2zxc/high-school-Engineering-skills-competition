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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3480
      Top             =   2280
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
Dim a(7), i As Byte

Private Sub Form_Load()
  a(0) = &H1
  a(1) = &H2
  a(2) = &H4
  a(3) = &H8
  a(4) = &H10
  a(5) = &H20
  a(6) = &H40
  a(7) = &H80
  i = 0
End Sub

Private Sub Timer1_Timer()
  If OpenUsbDevice(VendorID, ProductID) Then Label1.Caption = "Connect" Else Label1.Caption = "DisConnect"
  OutDataEightByte &H11, a(i), 0, 0, 0, 0, 0, 0
  If i = 7 Then i = 0 Else i = i + 1
End Sub
