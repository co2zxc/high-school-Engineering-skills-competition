VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form Login 
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "�n�J"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton ���� 
      Caption         =   "����(&C)"
      Height          =   375
      Left            =   2370
      TabIndex        =   5
      Top             =   1605
      Width           =   1215
   End
   Begin VB.TextBox pwd 
      Height          =   285
      IMEMode         =   3  '�Ȥ�
      Left            =   2175
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   930
      Width           =   1440
   End
   Begin VB.CommandButton �T�w 
      Caption         =   "�T�w(&S)"
      Height          =   360
      Left            =   735
      TabIndex        =   1
      Top             =   1605
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo AllEmp 
      Height          =   330
      Left            =   2130
      TabIndex        =   0
      Top             =   330
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "���w�m�W�G"
      Height          =   330
      Left            =   1005
      TabIndex        =   4
      Top             =   390
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "��J�K�X�G"
      Height          =   300
      Left            =   1020
      TabIndex        =   3
      Top             =   990
      Width           =   960
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AllEmp_Change()
pwd.SetFocus
End Sub

Private Sub Form_Load()
AllEmp.ListField = "�m�W"
Set AllEmp.RowSource = getAllemp()
End Sub

Private Sub ����_Click()
If MsgBox("�T�w�������t��?", 4) = 6 Then
    End
End If
End Sub

Private Sub �T�w_Click()
Set NowUser = AllEmp.RowSource
NowUser.MoveFirst
NowUser.Find "�m�W='" & AllEmp.Text & "'"
If NowUser.EOF = True Then
    MsgBox "�������u�A�Э��s�ާ@"
    Exit Sub
End If
c = NowUser("�K�X")
If Len(c) <> Len(pwd.Text) Then
    MsgBox "�K�X��J���~"
    Exit Sub
End If
For i = 1 To Len(c)
    If Asc(Mid(c, i, 1)) <> Asc(Mid(pwd.Text, i, 1)) Then
        MsgBox "�K�X��J���~"
        Exit Sub
    End If
Next
Form1.Show
Form1.NowUserName = AllEmp.Text
If NowUser("���O�W��") = "��J" Then
    Form1.�Ȥ�έp.Enabled = False
    Form1.���~�P��.Enabled = False
    Form1.��b��.Enabled = False
    Form1.�X�f�ΰh�f.Enabled = False
ElseIf NowUser("���O�W��") = "�X�f" Then
    Form1.�Ȥ�έp.Enabled = False
    Form1.���~�P��.Enabled = False
    Form1.�q��B�z.Enabled = False
End If
Unload Me
End Sub
