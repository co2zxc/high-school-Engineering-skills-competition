VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "MDIForm1"
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10245
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  '�t�ιw�]��
   WindowState     =   2  '�̤j��
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   8040
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '������U��
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2548
            MinWidth        =   2548
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5274
            MinWidth        =   5274
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2542
            MinWidth        =   2542
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3529
            MinWidth        =   3529
         EndProperty
      EndProperty
   End
   Begin VB.Menu F 
      Caption         =   "�t�Τ��� "
      Begin VB.Menu F1 
         Caption         =   "�t�ΰʾ� "
      End
      Begin VB.Menu F2 
         Caption         =   "�t�Υت� "
      End
   End
   Begin VB.Menu A 
      Caption         =   "�Ȥ���"
      Begin VB.Menu a1 
         Caption         =   "�޲z�Ȥ� "
      End
      Begin VB.Menu A2 
         Caption         =   "�d�߫Ȥ���"
      End
   End
   Begin VB.Menu B 
      Caption         =   "�ӫ~��ƺ޲z "
   End
   Begin VB.Menu C 
      Caption         =   "�P����"
   End
   Begin VB.Menu H 
      Caption         =   "���ʰT�� "
      Begin VB.Menu H1 
         Caption         =   "�s���~���W���M�S�����u�f "
      End
      Begin VB.Menu H2 
         Caption         =   "�P��e�Q�W "
      End
   End
   Begin VB.Menu G 
      Caption         =   "�խ����"
      Begin VB.Menu G1 
         Caption         =   "�קʯ�"
      End
      Begin VB.Menu G2 
         Caption         =   "��i�T"
      End
   End
   Begin VB.Menu E 
      Caption         =   "���� "
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a1_Click()
 If Not (ActiveForm Is Nothing) Then Unload Me
  Form1.Show
End Sub

Private Sub A2_Click()
 If Not (ActiveForm Is Nothing) Then Unload Me
  Form4.Show
End Sub

Private Sub B_Click()
If frmLogin.user = "" Then
MsgBox "�Ϧ~�v�u�����A���i�ϤJ���t��", 0 + 64, "���~�T��"
Else
 If Not (ActiveForm Is Nothing) Then Unload Me
  Form2.Show
End If


End Sub

Private Sub C_Click()

If Not (ActiveForm Is Nothing) Then Unload Me
  Form3.Show
End Sub

Private Sub D_Click()
'If frmLogin.user = "" Then
'MsgBox "�Ϧ~�v�u�����A���i�ϤJ���t��", 0 + 64, "���~�T��"
'Else
If Not (ActiveForm Is Nothing) Then Unload Me
  Form5.Show
'End If


End Sub

Private Sub E_Click()
End
End Sub

Private Sub F1_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form6.Show
End Sub

Private Sub F2_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form7.Show
End Sub

Private Sub G1_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form8.Show
End Sub

Private Sub G2_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form9.Show
End Sub

Private Sub H1_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form11.Show
End Sub

Private Sub H2_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form10.Show
End Sub

Private Sub MDIForm_Load()

If frmLogin.user = "��i�T" Then
  MDIForm1.StatusBar1.Panels(4) = "¾��G�`��"
  Else
End If

  If frmLogin.user = "�קʯ�" Then
  MDIForm1.StatusBar1.Panels(4) = "¾��G����"
  Else
End If

If frmLogin.user = "²���I" Then
  MDIForm1.StatusBar1.Panels(4) = "¾��G�M��u"
  Else
End If

 StatusBar1.Panels(1) = "�ϥΪ̡G" & frmLogin.user
End Sub

Private Sub Timer1_Timer()
  StatusBar1.Panels(2) = "�{�b�ɶ��G" & Now()
End Sub
