VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "USB�����]�p�P���γ]�p�J�� �{���d�Ҥ@"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5205
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Click_Flag As Integer
Private Sub Form_Click()
 
 If Click_Flag = 0 Then
    Print "USB�����]�p�P���γ]�p�J��:����VB~~~~~~~~"
    Click_Flag = 1
 Else
    Cls
    Click_Flag = 0
 End If
End Sub

