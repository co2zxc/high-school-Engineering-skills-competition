VERSION 5.00
Begin VB.Form frmChoose 
   Caption         =   "��ܪ��"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   4380
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�T�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cboStaff 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ComboBox cboCust 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblChoose 
      AutoSize        =   -1  'True
      Caption         =   "�g��H"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label lblChoose 
      AutoSize        =   -1  'True
      Caption         =   "�Ȥ�W��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   960
   End
   Begin VB.Label lblChoose 
      AutoSize        =   -1  'True
      Caption         =   "�п�ܫȤ�P�g��H"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2700
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�ˬdcboCust�����Ȥ�O�_�����ƪ��ܼ�
Dim pre_Cust_No As String

'��s�W�@���D�ɫ�,���Q�~��h�i���U�����s, �H�^������J�ɪ����A
Private Sub cmdCancel_Click()
    rsSell.CancelUpdate
    Me.Hide
    frmSell.Show
    '���U�����s�����A��1
    Main_cancel = 1
End Sub

'�b���U�T�w�s��,��ܪ�椤���Ȥ��g��H��ƱN�|�g�^��ڤ�,�ç����@���D�ɪ��s�W
Private Sub cmdOK_Click()
    If cboStaff.Text <> "���I��" And cboStaff.Text <> "" And _
        cboCust.Text <> "���I��" And cboCust.Text <> "" Then
        
        '�NcboStaff���ȥH�s�����Φ��g��P�f�檺txtSell(1)��
        rsStaff.MoveFirst
        rsStaff.Find "�m�W='" & cboStaff.Text & "'", , adSearchForward, 1
        If rsStaff.EOF = False Then
            frmSell.txtSell(1).Text = rsStaff.Fields("���u�s��")
        End If
        
        '�NcboCust���ȥH�s�����Φ��g��P�f�檺txtSell(2)��
        rsCust.MoveFirst
        rsCust.Find "���q�W��='" & cboCust.Text & "'", , adSearchForward, 1
        If rsCust.EOF = False Then
            frmSell.txtSell(2).Text = rsCust.Fields("�Ȥ�s��")
        End If
    Else
        MsgBox ("�п�ܸg��H�ΫȤ�!!")
        Exit Sub
    End If
    rsSell.Update
    Me.Hide
    frmSell.Show
    '���U�T�w�s�����A��0
    Main_cancel = 0
End Sub

'�����J,�s�@cboStaff�PcboCust
Public Sub Form_Activate()
    
    '�s�@cboStaff
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from ���u��ƪ�"
    rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
    
    '�b�C�@���I�s��ܪ���,�������NcboStaff �����Ȳ���,�M�᭫�sAddItem�i��
    cboStaff.Clear
    '��cboStaff���M�欰�W�����
    Do Until rsStaff.EOF
        cboStaff.AddItem rsStaff.Fields("�m�W")
        rsStaff.MoveNext
    Loop
    
    '�s�@cboCust
    Set rsCust = New ADODB.Recordset
    sql_rsCust = "select * from �Ȥ���"
    rsCust.Open sql_rsCust, cn, adOpenDynamic, adLockOptimistic
    
    Set rsOrder = New ADODB.Recordset
    sql_rsOrder = "select * from �q��D�� where �O�_�־P=0 order by �Ȥ�s��"
    rsOrder.Open sql_rsOrder, cn, adOpenDynamic, adLockOptimistic
    
    '�b�C�@���I�s��ܪ���,�������NcboCust �����Ȳ���,�M�᭫�sAddItem�i��
    pre_Cust_No = ""
    cboCust.Clear
    '��cboCust���M�欰�W�����,�ñN���ƪ��Ȥ�h��
    Do Until rsOrder.EOF
        If pre_Cust_No <> rsOrder.Fields("�Ȥ�s��") Then
            rsCust.MoveFirst
            rsCust.Find "�Ȥ�s��='" & rsOrder.Fields("�Ȥ�s��") & "'"
            cboCust.AddItem rsCust.Fields("���q�W��")
        End If
        pre_Cust_No = rsOrder.Fields("�Ȥ�s��")
        rsOrder.MoveNext
    Loop
    'combo Box����l��
    cboStaff.Text = "���I��"
    cboCust.Text = "���I��"
End Sub






