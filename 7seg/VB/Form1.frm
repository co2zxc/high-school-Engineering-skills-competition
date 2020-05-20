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
   Begin VB.Timer Timer2 
      Interval        =   100
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
Dim j As Byte

Private Sub Form_Load()
 j = 0
End Sub

Private Sub Timer2_Timer()
  If OpenUsbDevice(VendorID, ProductID) Then Label1.Caption = "Connect" Else Label1.Caption = "DisConnect"
  OutDataEightByte &H21, 0, 1, 2, 3, 0, 0, 0
  If j = 3 Then j = 0 Else j = j + 1
End Sub

