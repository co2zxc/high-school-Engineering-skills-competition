VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   Caption         =   "����"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   WindowState     =   2  '�̤j��
   Begin VB.Menu order 
      Caption         =   "�q�ʺ޲z"
      Begin VB.Menu reorder 
         Caption         =   "�q�f�B�z"
         Begin VB.Menu materiel 
            Caption         =   "���ƽ�"
         End
         Begin VB.Menu accountant 
            Caption         =   "�|�p��"
         End
      End
      Begin VB.Menu return 
         Caption         =   "�h�f�B�z"
      End
   End
   Begin VB.Menu save 
      Caption         =   "�s�f�޲z"
      Begin VB.Menu stock 
         Caption         =   "�i�f�B�z"
      End
      Begin VB.Menu take 
         Caption         =   "��f�B�z"
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
  X = InputBox("�п�J�K�X", "�׼ҹs��w�s�޲z�t��") '��J�K�X
  password = Val(X)            '�N�r�����ন�ƭ�
    If password = 123 Then
     MsgBox "�w��z�ϥέq�f�t��"
     Form2.Show
    Else
     MsgBox "�K�X���~�A�L�k�n�J"
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
