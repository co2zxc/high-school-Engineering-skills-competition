VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCust 
   Caption         =   "�Ȥ��ƪ�"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   8115
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   6360
      TabIndex        =   25
      Top             =   3600
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   6720
      TabIndex        =   24
      Top             =   3600
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   720
      TabIndex        =   23
      Top             =   3600
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   1080
      TabIndex        =   22
      Top             =   3600
      Width           =   350
   End
   Begin VB.TextBox txtArea 
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
      Left            =   1560
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     ��"
      Height          =   375
      Index           =   3
      Left            =   6600
      TabIndex        =   19
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�s     �W"
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   18
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�R     ��"
      Height          =   375
      Index           =   4
      Left            =   6600
      TabIndex        =   17
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�x     �s"
      Height          =   375
      Index           =   0
      Left            =   6600
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     ��"
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   15
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�^�D�e��"
      Height          =   375
      Index           =   5
      Left            =   6600
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtCust 
      DataField       =   "�a�Ͻs��"
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
      Index           =   6
      Left            =   1560
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtCust 
      Alignment       =   1  '�a�k���
      DataField       =   "�q��"
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
      Index           =   4
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   3000
   End
   Begin VB.TextBox txtCust 
      DataField       =   "�p���H"
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
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   3000
   End
   Begin VB.TextBox txtCust 
      DataField       =   "���q�t�d�H"
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
      Index           =   2
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   3000
   End
   Begin VB.TextBox txtCust 
      DataField       =   "���q�W��"
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
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   720
      Width           =   3000
   End
   Begin VB.TextBox txtCust 
      Alignment       =   1  '�a�k���
      DataField       =   "�Ȥ�s��"
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
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   3000
   End
   Begin MSDataListLib.DataCombo dcboArea 
      Height          =   390
      Left            =   1560
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtCust 
      DataField       =   "���q�a�}"
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
      Index           =   5
      Left            =   2880
      TabIndex        =   11
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label lblRecord 
      BorderStyle     =   1  '��u�T�w
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
      Left            =   1440
      TabIndex        =   21
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "���q�a�}"
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
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "�q��"
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
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "�p���H"
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
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "���q�t�d�H"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "���q�W��"
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
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "�Ȥ�s��"
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
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ƪ��s��
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '�x�s
            For i = 1 To 5
                If txtCust(i).Text = "" Then
                    MsgBox "��p!��Ƥ���ť�!", , "�����q�q���ѥ��������q"
                    txtCust(i).SetFocus
                    Exit Sub
                End If
            Next i
                
            rsArea.MoveFirst
            rsArea.Find "�a�ϦW��='" & dcboArea.Text & "'", , 1
            If rsArea.EOF Then
                MsgBox "�п�ܥ��T���a�ϦW��!", , "�����q�q���ѥ��������q"
                dcboArea.SetFocus
                Exit Sub
            Else
                txtCust(6).Text = dcboArea.BoundText
            End If
            
            rsCust.Update
            control_status False
            Call cmdMove_Click(3)

        Case 1      '����
            rsCust.CancelUpdate
            If rsCust.RecordCount <> 0 Then
                '�Y�O�s�W���A�U������,�h�n���㵧�ƪ����
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                For Each oText In Me.txtCust
                    Set oText.DataSource = rsCust
                Next
                dcboArea.BoundText = txtCust(6).Text
                control_status False
            Else
                Call cmdEdit_Click(5)
            End If
            Add_flag = 0
            
        Case 2      '�s�W
            Add_flag = 1
            Call Add_No(rsCust, 1, 1)
            control_status True
            txtCust(0).Text = Main_No
            dcboArea.Text = "���I��"
            
        Case 3      '�ק�
            control_status True
            
        Case 4      '�R��
            If MsgBox("�T�w�n�R�������O��??", 32 + vbYesNo, "�����q�q���ѥ��������q") = vbYes Then
                rsCust.Delete
                If rsCust.RecordCount <> 0 Then
                    Call cmdMove_Click(2)
                Else
                    Call cmdEdit_Click(5)
                End If
            End If
            
        Case 5      '�^�D�e��
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            rsCust.Close
            Set rsCust = Nothing
    End Select
End Sub

'��ƪ�����
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsCust)
    lblRecord.Caption = "�Ȥ��ơG" & intRecord & "/" & intTotal
End Sub

'��ƪ��s��
Private Sub Form_Load()
    '1.�}�ҫȤ���
    Set rsCust = New ADODB.Recordset
    sql_rsCust = "select * from �Ȥ���"
    rsCust.Open sql_rsCust, cn, adOpenDynamic, adLockOptimistic
    
    '2.�N��ƫ��w����椤���P������
    For Each oText In Me.txtCust
        Set oText.DataSource = rsCust
    Next
    
    '3.�s�@dcboArea
    Set rsArea = New ADODB.Recordset
    sql_rsArea = "select * from �a�Ϫ�"
    rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
    
    Set dcboArea.DataSource = rsCust
    dcboArea.DataField = "�a�Ͻs��"
    Set dcboArea.RowSource = rsArea
    dcboArea.ListField = "�a�ϦW��"
    dcboArea.BoundColumn = "�a�Ͻs��"
    rsCust.MoveFirst
    
    '4.�ˬd�O�_����
    If rsCust.RecordCount <> 0 Then
        cmdMove_Click (0)
        control_status False
    Else
        '�Y�L���
        Call cmdEdit_Click(2)
        control_status True
    End If
    
    '5.��檺���e�]�w
    frmCust.Height = 4920
    frmCust.Width = 8235
    
End Sub

'��������A
Private Sub control_status(control_flag As Boolean)
    '�T�w���ܪ����A
    txtCust(0).Enabled = False
    txtArea.Enabled = False
    txtCust(6).Visible = False
    
    '�i���������A
    For i = 1 To 5
        txtCust(i).Enabled = control_flag
    Next
    dcboArea.Visible = control_flag
    txtArea.Visible = Not control_flag
    
    '�x�s,�����s
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next
    '�s�W,�ק�,�R��,�^�D�e���s
    For i = 2 To 5
        cmdEdit(i).Visible = Not control_flag
    Next i
    '���ʶs
    For i = 0 To 3
        cmdMove(i).Enabled = Not control_flag
    Next i
    
End Sub

'��txtArea���ȻP��Ƴs��
Private Sub txtCust_Change(Index As Integer)
    Select Case Index
        Case 6
            Set rsArea = New ADODB.Recordset
            sql_rsArea = "select * from �a�Ϫ�"
            rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
            
            rsArea.Find "�a�Ͻs��='" & txtCust(6).Text & "'"
            If rsArea.EOF = False Then
                txtArea.Text = rsArea.Fields("�a�ϦW��")
            End If
    End Select
End Sub
