VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "首頁"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.Menu order 
      Caption         =   "訂購管理"
      Begin VB.Menu reorder 
         Caption         =   "訂貨處理"
         Begin VB.Menu materiel 
            Caption         =   "物料課"
         End
         Begin VB.Menu accountant 
            Caption         =   "會計課"
         End
      End
      Begin VB.Menu return 
         Caption         =   "退貨處理"
      End
   End
   Begin VB.Menu save 
      Caption         =   "存貨管理"
      Begin VB.Menu stock 
         Caption         =   "進貨處理"
      End
      Begin VB.Menu take 
         Caption         =   "領貨處理"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub accountant_Click()
    Form3.Show
End Sub

Private Sub materiel_Click()
  Dim X, password As Integer
  X = InputBox("請輸入密碼", "修模零件庫存管理系統") '輸入密碼
  password = Val(X)            '將字串資料轉成數值
    If password = 123 Then
     MsgBox "歡迎您使用訂貨系統"
     Form2.Show
    Else
     MsgBox "密碼錯誤，無法登入"
    End If
  
  End Sub

Private Sub return_Click()
     Form4.Show
End Sub

Private Sub stock_Click()
    Form6.Show
End Sub

Private Sub take_Click()
    Form8.Show
End Sub
