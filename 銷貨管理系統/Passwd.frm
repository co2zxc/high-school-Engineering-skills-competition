VERSION 5.00
Begin VB.Form Passwd 
   BackColor       =   &H00FAEEEB&
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "���}���޲z�t�εn�J"
   ClientHeight    =   4320
   ClientLeft      =   2820
   ClientTop       =   3870
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F0D3CC&
      Caption         =   "�����t��"
      Height          =   495
      Left            =   4440
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   7335
      TabIndex        =   6
      Top             =   0
      Width           =   7335
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7200
         Y1              =   1920
         Y2              =   1920
      End
   End
   Begin VB.CommandButton Cmd_Flase 
      BackColor       =   &H00F0D3CC&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_Ture 
      BackColor       =   &H00F0D3CC&
      Caption         =   "�T�w�n�J"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaskColor       =   &H8000000F&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00EFCFC7&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  '�Ȥ�
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00EFCFC7&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   2  '����
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   7200
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "�޲z���K�X�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "�޲z���b���G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "Passwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Login_Time As Integer '�O���K�X��J����
Dim QryRs As ADODB.Recordset

Private Sub Cmd_Flase_Click()
    End
End Sub

Private Sub Cmd_Flase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmd_Ture_Click()
If Text(0).Text = "manager" Then
    frmMDIMain.Show
    Unload Me
Else
        '�K�X�T�{�@�~
        Sqlstr = " select ad_id,ad_pw from ad" _
               & " where ad_id = '" & Trim(Text(0).Text) & "'" _
               & " and   ad_pw = '" & Trim(Text(1).Text) & "'" _
               
        Set QryRs = New ADODB.Recordset
        QryRs.Open Sqlstr, cn, adOpenStatic
        If QryRs.RecordCount = 0 Then
            Login_Time = Login_Time + 1
            
            
            MsgBox "�K�X���~,�|�� " & 3 - Login_Time & " �����|", 48, "�n�J�T��"
            
            Text(0).Text = ""
            Text(1).Text = ""
            Text(0).SetFocus
            
            If Login_Time = 3 Then End
            
        Else
            
            frmMDIMain.Show
            
            Unload Me
        End If
End If

End Sub

Private Sub Cmd_Ture_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Command1_Click()
    msgend = MsgBox("�T�w�n���}���t�ζ�?", vbYesNo, "�T���T�{")
    If msgend = vbYes Then
        End
    Else
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Login_Time = 0
End Sub


Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub '��enter
    If KeyAscii = 8 Then: Exit Sub
   
End Sub

