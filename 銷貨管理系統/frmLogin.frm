VERSION 5.00
Begin VB.Form Password 
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "�����q�i�P�s�t��"
   ClientHeight    =   5475
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3234.812
   ScaleMode       =   0  '�ϥΪ̦ۭq
   ScaleWidth      =   7816.724
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command1 
      Caption         =   "�M��"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�T�w"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "���}"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2400
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   465
      IMEMode         =   3  '�Ȥ�
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      Caption         =   "�w��ϥ�"
      Height          =   1935
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   5415
      Begin VB.TextBox txtuser 
         Height          =   495
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label 
         Alignment       =   2  '�m�����
         Caption         =   "�п�J�t�αb��"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  '�m�����
         Caption         =   "�п�J�t�αK�X"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Login_Time As Integer '�O���K�X��J����
Public LoginSucceeded As Boolean
Dim mnuMkCk As String
Dim mnuInfo As String
Private Sub cmdCancel_Click()
   
   End
   
End Sub


Private Sub cmdOK_Click()
    '�ˬd�K�X�����T��
    If txtuser = "qwe" And txtPassword = "123" Then
        '�N�ǻ����\�T���ܩI�s���Ƶ{����
        '�Ƶ{�����{���X�m�󦹳B
        '�]�w�����ܼƬO��²�檺�覡
         frmMDIMain.Show
            Unload Me
    ElseIf txtuser = "asd" And txtPassword = "456" Then
        
         frmMDIMain.Show
         frmMDIMain.mnuMkCk.Enabled = False
         Unload Me
    ElseIf txtuser = "zxc" And txtPassword = "789" Then
        
         frmMDIMain.Show
         frmMDIMain.mnuInfo.Enabled = False
         frmMDIMain.mnuMkCk.Enabled = False
         Unload Me
    Else
    
            Login_Time = Login_Time + 1
            
            MsgBox "�K�X���~,�|�� " & 3 - Login_Time & " �����|", 48, "�n�J�T��"
            
         
         If Login_Time = 3 Then
         
         End
            
        
        End If
      End If
        ' MsgBox "�K�X���~�A�Э��s��J!", , "�n�J"
       ' txtPassword.SetFocus
       ' SendKeys "{Home}+{End}"
         
        

End Sub

Private Sub Command1_Click()
txtuser = ""
txtPassword = ""
End Sub

