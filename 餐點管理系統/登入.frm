VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "�n�J"
   ClientHeight    =   4095
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "�n�J.frx":0000
   ScaleHeight     =   2419.461
   ScaleMode       =   0  '�ϥΪ̦ۭq
   ScaleWidth      =   5633.674
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "�n�J.frx":BFC2
      Height          =   360
      Left            =   2880
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "���u�s��"
      BoundColumn     =   "���u�s��"
      Text            =   "��ܭ��u�s��- - -"
      Object.DataMember      =   "���u"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�رd���r��"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�T�w"
      Default         =   -1  'True
      Height          =   390
      Left            =   2160
      TabIndex        =   3
      Top             =   3480
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  '�Ȥ�
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2880
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "�ϥΪ̦W��(&U):"
      BeginProperty Font 
         Name            =   "�رd���r��"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "�K�X(&P):"
      BeginProperty Font 
         Name            =   "�رd���r��(P)"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   2760
      Width           =   2160
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public user As String
Public ID As String



Private Sub cmdOK_Click()
  Do While Not DataEnvironment1.rs���u.EOF
      If DataEnvironment1.rs���u("���u�s��") = DataCombo1 Then
         ID = DataCombo1
         user = DataEnvironment1.rs���u("���u�m�W")
         Password = DataEnvironment1.rs���u("�K�X")
         Exit Do
      End If
      DataEnvironment1.rs���u.MoveNext
    Loop
    
    If txtPassword = Password Then
        Unload Me
        MDIForm1.Show
    Else
        MsgBox "�K�X���~�A�Э��s��J!", , "�n�J"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

