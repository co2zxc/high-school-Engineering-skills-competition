VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSell 
   Caption         =   "�P�f��"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11340
   WindowState     =   2  '�̤j��
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "�C�L����"
      Height          =   375
      Left            =   1200
      TabIndex        =   26
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "�s���q��"
      Height          =   375
      Left            =   1200
      TabIndex        =   25
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�^�D�e��"
      Height          =   375
      Index           =   3
      Left            =   7800
      TabIndex        =   24
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraSelDetail 
      Caption         =   "�P�f����"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2520
      TabIndex        =   5
      Top             =   3960
      Width           =   6495
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "�x     �s"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   21
         Top             =   2880
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dgSelDetail 
         Height          =   2295
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1.2
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "��     ��"
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   10
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.Frame fraSell 
      Caption         =   "�P�f�D��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   6495
      Begin VB.TextBox txtSell 
         Alignment       =   1  '�a�k���
         DataField       =   "�X�f��"
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
         Index           =   3
         Left            =   1575
         TabIndex        =   22
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "<<"
         Height          =   355
         Index           =   1
         Left            =   960
         TabIndex        =   19
         Top             =   2520
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "|<"
         Height          =   355
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   2520
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">|"
         Height          =   355
         Index           =   3
         Left            =   4920
         TabIndex        =   17
         Top             =   2520
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">>"
         Height          =   355
         Index           =   2
         Left            =   4560
         TabIndex        =   16
         Top             =   2520
         Width           =   350
      End
      Begin VB.ComboBox cboEmployee 
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
         TabIndex        =   15
         Top             =   960
         Width           =   2500
      End
      Begin VB.TextBox txtCust 
         DataField       =   "�Ȥ�W��"
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
         TabIndex        =   14
         Top             =   1440
         Width           =   2500
      End
      Begin VB.TextBox txtStaff 
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
         TabIndex        =   13
         Top             =   960
         Width           =   2500
      End
      Begin VB.TextBox txtSell 
         DataField       =   "�Ȥ�s��"
         Height          =   345
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "��     ��"
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "��     ��"
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�x     �s"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtSell 
         DataField       =   "�g��H�s��"
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
         TabIndex        =   4
         Top             =   960
         Width           =   2500
      End
      Begin VB.TextBox txtSell 
         Alignment       =   1  '�a�k���
         DataField       =   "�P�f��s��"
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
         TabIndex        =   3
         Top             =   480
         Width           =   2500
      End
      Begin VB.Label lblSell 
         AutoSize        =   -1  'True
         Caption         =   "�X�f��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   720
         TabIndex        =   23
         Top             =   2070
         Width           =   720
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
         Left            =   1320
         TabIndex        =   20
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label lblSell 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   930
         TabIndex        =   11
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label lblSell 
         AutoSize        =   -1  'True
         Caption         =   "�g��H"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label lblSell 
         AutoSize        =   -1  'True
         Caption         =   "�P�f�渹"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   510
         TabIndex        =   1
         Top             =   660
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSelDetail_list As ADODB.Recordset     '�P�f���Ӫ��ȦsRecordSet
Dim update_flag As Integer                  '�����x�s�s���X��(�ΨӨM�w�O�_��s�w�s�P�־P���)

'�D�ɪ��s��
Public Sub cmdEdit_Click(index As Integer)
    Select Case index
        Case 0      '�x�s
            rsStaff.MoveFirst
            rsStaff.Find "�m�W='" & cboEmployee.Text & "'", , 1
            If rsStaff.EOF Then
                MsgBox "�п�ܸg��H!", , "�����q�q���ѥ��������q"
                cboEmployee.SetFocus
                Exit Sub
            Else
                If txtSell(3).Text = "" Then
                    MsgBox "�ж�g�X�f��!", , "�����q�q���ѥ��������q"
                    txtSell(3).SetFocus
                    Exit Sub
                Else
                    If IsDate(txtSell(3).Text) = False Then
                        MsgBox "�z��J������榡���~,�ЦA�T�{�@�M!", , "�����q�q���ѥ��������q"
                        txtSell(3).SetFocus
                        Exit Sub
                    Else
                        '�x�s��
                        txtSell(1).Text = rsStaff.Fields("���u�s��")
                        
                        rsSell.Fields(0) = txtSell(0).Text
                        rsSell.Fields(1) = rsOrder.Fields(0)
                        rsSell.Fields(2) = txtSell(1).Text
                        
                        Set rsCust = New ADODB.Recordset
                        rsCust.Open "select * from �Ȥ���", cn, adOpenDynamic, adLockOptimistic
                        rsCust.MoveFirst
                        rsCust.Find "���q�W��='" & txtCust.Text & "'", , 1
                        rsSell.Fields(3) = rsCust.Fields(0)
                        rsSell.Fields(4) = txtCust.Text
                        rsSell.Fields(5) = txtSell(3).Text
                        rsSell.Update
                        
                        control_status False
                        Call cmdMove_Click(3)
                    End If
                End If
            End If
            
            '��s�W�@���P�f�D�ɫ�,�Y�i�J�P�f���Ӫ��s�W
            If Add_flag = 1 Then
                '�P�f���Ӹ�ƨӷ�
                Call SelDetail_Data
            End If
            Add_flag = 0
            
        Case 1      '����
            rsSell.CancelUpdate
            If rsSell.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                
                '�^�Э�Ӫ���Ƥ��e
                For Each oText In Me.txtSell
                    Set oText.DataSource = rsSell
                Next
                control_status False
            Else
                Call cmdEdit_Click(3)
                MsgBox "�ثe�õL����P�f���!", , "�����q�q���ѥ��������q"
            End If
            
            Add_flag = 0
                
        Case 2      '�ק�
            control_status True
            '�U�Ԧ�������w�]�Ȭ���Ӫ����u�m�W
            cboEmployee.Text = txtStaff.Text
            fraSelDetail.Visible = True
            
        Case 3      '�^�D�e��
            '1.�ѭq���,��ܭn�s�W�@�i�P�f��(���s�W���A)
            If cmdSell_flag = 1 Then
                '2.�����F�s�W�{��,�Y���;P�f��F,�G�n��s�w�s�P�־P����
                If update_flag = 1 Then
                    '��s�w�s�P�־P��줧�Ƶ{��
                    Call Data_Check
                    Call frmOrder.dgOrDetail_refresh
                    intRecord = bm_move
                    Call frmOrder.cmdSell_status
                End If
            Else
                '��ѥ\���'�P�f��d��'�\�ඵ�بӶ}�ҾP�f��(���s�����A)
                frmMDIMain.mnuExit.Enabled = True
                '�^��D�e���ɶ�����\�ඵ�ت����A
                Call frmMDIMain.mnuSell_status
            End If
                        
            update_flag = 0
            cmdSell_flag = 0
            Unload Me
            
    End Select
End Sub

'��ƪ�����
Private Sub cmdMove_Click(index As Integer)
    Call rs_Move(index, rsSell)
    lblRecord.Caption = "�P�f��ơG" & intRecord & "/" & intTotal
End Sub

'���Ӫ��s��
Private Sub cmdSubEdit_Click(index As Integer)
    Select Case index
        Case 0      '�����x�s
            
            '�����x�s���X��(�Y������s�w�s��ƻP�q�檺�־P���X��)
            update_flag = 1
            
            '�N�P�f���Ӫ��Ȧs���rsSelDetail_list�g�i��Ʈw���P�f���Ӹ�ƪ�
            rsSelDetail_list.MoveFirst
            Do Until rsSelDetail_list.EOF
                Call Add_SubNo(rsSell_Detail, 2, txtSell(0))
                rsSell_Detail.AddNew
                rsSell_Detail.Fields(0) = txtSell(0).Text
                rsSell_Detail.Fields(1) = Sub_No
                rsSell_Detail.Fields(2) = rsSelDetail_list.Fields("���~�s��")
                rsSell_Detail.Fields(3) = RTrim(rsSelDetail_list.Fields("���~�W��"))
                rsSell_Detail.Fields(4) = rsSelDetail_list.Fields("���")
                rsSell_Detail.Fields(5) = rsSelDetail_list.Fields("�P�f�ƶq")
                rsSell_Detail.Fields(6) = Int(rsSelDetail_list.Fields("���")) * Int(rsSelDetail_list.Fields("�P�f�ƶq"))
                rsSell_Detail.Update
                
                rsSelDetail_list.MoveNext
            Loop
            
            Call SelDetail_refresh
            subcontrol_status True
            '�x�s�P�����s���i��
            For i = 0 To 1
                cmdSubEdit(i).Visible = False
            Next i
            
        Case 1      '���Ө���
            If MsgBox("�P�f���ӥ��������,�_�h�P�f�D�ɱN�Q�R��!�T�w�n�������Ӹ�ƪ��s�W?", 48 + vbYesNo, "�����q�q���ѥ��������q") = vbYes Then
                rsSell.Delete
                If rsSell.RecordCount <> 0 Then
                    Call cmdMove_Click(2)
                    
                    subcontrol_status True
                    '�x�s�P�����s���i��
                    For i = 0 To 1
                        cmdSubEdit(i).Visible = False
                    Next i
                Else
                    Call cmdEdit_Click(3)
                    MsgBox "�ثe�õL����P�f���!", , "�����q�q���ѥ��������q"
                End If
            End If
    End Select
            
End Sub



Private Sub cmdPreview_Click()
RptOrder.Show
End Sub

Private Sub cmdPrinter_Click()
    '�ˬd�O�_���w�˦L���
    If Printers.Count <> 0 Then
        RptOrder.PrintReport True, rptRangeAllPages
    Else
        MsgBox ("�Х��w�˦L���!!")
    End If
End Sub





'�D�ɪ���Ƴs��
Private Sub Form_Load()
    '1.�}��RecordSet
    Set rsSell = New ADODB.Recordset
    sql_rsSell = "select * from �P�f�D�� order by �P�f��s��"
    rsSell.Open sql_rsSell, cn, adOpenDynamic, adLockOptimistic
    
    '2.�N��ƫ��w���P������
    For Each oText In Me.txtSell
        Set oText.DataSource = rsSell
    Next
    Set txtCust.DataSource = rsSell
    
    '3.�s�@cboEmployee���M�涵��
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from ���u��ƪ�"
    rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
    
    cboEmployee.Clear
    Do Until rsStaff.EOF
        cboEmployee.AddItem rsStaff.Fields("�m�W")
        rsStaff.MoveNext
    Loop
    
    '4.�P�_�O�q���̪����O�Ӷ}�ҾP�f��
    '�Y�ѭq��(���;P�f�� �s)�Ӫ�,�h�����i�J�s�W���A
    If cmdSell_flag = 1 Then
        Add_flag = 1
        Call Add_No(rsSell, 4, 1)
        control_status True
        txtSell(0).Text = Main_No
        txtCust.Text = frmOrder.txtCustomer.Text
        cboEmployee.Text = "--�п�ܸg��H--"
        txtSell(3).Text = Date
    Else
        '�ѥ\��� �P�f��d�� ���O�Ӫ�,�h�����e�{�X�P�f���
        Call cmdMove_Click(0)
        control_status False
        For i = 0 To 1
            cmdSubEdit(i).Visible = False
        Next i
    End If
End Sub

'���Ӹ�ƪ��s��
Private Sub SelDetail_refresh()
    '1.�}�ҾP�f����,�åH�P�f��s����P�f�D�����p
    Set rsSell_Detail = New ADODB.Recordset
    sql_rsSell_Detail = "select * from �P�f���� where �P�f��s��='" & txtSell(0).Text & "'"
    rsSell_Detail.Open sql_rsSell_Detail, cn, adOpenDynamic, adLockOptimistic
    
    '2.�N��ƫ��w��DataGrid
    Set dgSelDetail.DataSource = rsSell_Detail
    
    For i = 0 To 2
        dgSelDetail.Columns.Item(i).Visible = False
    Next i
    For i = 4 To 6
        dgSelDetail.Columns.Item(i).Alignment = dbgRight
    Next i
End Sub

'��ƪ��s��
Private Sub txtSell_Change(index As Integer)
    '�P�f���Ӹ�Ƴs��
    Call SelDetail_refresh
    
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from ���u��ƪ�"
    rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
        
    rsStaff.MoveFirst
    rsStaff.Find "���u�s��='" & txtSell(1).Text & "'", , adSearchForward, 1
    If rsStaff.EOF = False Then
        txtStaff.Text = rsStaff.Fields("�m�W")
        cboEmployee.Text = rsStaff.Fields("�m�W")
    End If
End Sub

'�D�ɱ�������A�]�w
Private Sub control_status(control_flag As Boolean)
    For i = 1 To 2
        txtSell(i).Visible = False
    Next i
    txtSell(0).Enabled = False
    txtCust.Enabled = False
    txtStaff.Enabled = False
    
    txtSell(3).Enabled = control_flag
    txtStaff.Visible = Not control_flag
    cboEmployee.Visible = control_flag
    
    '�x�s,���� �s
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next i
    '�ק�,�^�D�e�� �s
    For i = 2 To 3
        cmdEdit(i).Visible = Not control_flag
    Next i
    
    fraSelDetail.Visible = Not control_flag
End Sub

'���ӱ�������A�]�w
Private Sub subcontrol_status(subcontrol_flag As Boolean)
    fraSell.Enabled = subcontrol_flag
    cmdEdit(3).Enabled = subcontrol_flag
End Sub

'�P�f���Ӹ�ƨӷ�
Private Sub SelDetail_Data()
    '1.���s�@�@�ӪŪ�RecordSet
    Set rsSelDetail_list = New ADODB.Recordset
    rsSelDetail_list.Fields.Append "���~�s��", adChar, 3
    rsSelDetail_list.Fields.Append "���~�W��", adChar, 20
    rsSelDetail_list.Fields.Append "���", adInteger
    rsSelDetail_list.Fields.Append "�P�f�ƶq", adInteger
    rsSelDetail_list.Open
    
    '2.�A�N�۹������q����Ӹ�Ʃ��RecordSet��
    Do Until rsOrder_Detail.EOF
        rsSelDetail_list.AddNew
        rsSelDetail_list.Fields(0) = rsOrder_Detail.Fields("���~�s��")
        rsSelDetail_list.Fields(1) = rsOrder_Detail.Fields("���~�W��")
        rsSelDetail_list.Fields(2) = rsOrder_Detail.Fields("���")
        rsSelDetail_list.Fields(3) = rsOrder_Detail.Fields("�q�f�ƶq")
        rsSelDetail_list.Update
        
        rsOrder_Detail.MoveNext
    Loop
    
    '�ñN���Ȯɩʪ�RecordSet-rsSelDetail_list��Ƨe�{��DataGrid-dgSelDetail��
    Set dgSelDetail.DataSource = rsSelDetail_list
    dgSelDetail.Columns.Item(0).Visible = False
    
    subcontrol_status False
End Sub

'��s�P�f�־P��쬰1�P�Ͳ��־P��쬰3
Private Sub Data_Check()
    Set rsProdtStore = New ADODB.Recordset
    rsProdtStore.Open "select * from ���~�w�s�έp��", cn, adOpenKeyset, adLockOptimistic
    
    '1.��X�P�f�椤���q��s��=�ثe�q�f�椤���q��s�����O��
    rsSell.MoveFirst
    rsSell.Find "�q��s��='" & frmOrder.txtOrder(0).Text & "'", , 1
    '2.�����,��ܦ����q��w�g���ͤF�۹������P�f��F,�ҥH�P�f�־P��1,�Ͳ��־P��3
    If rsSell.EOF = False Then
        rsOrder_Detail.Filter = "�q��s��='" & frmOrder.txtOrder(0).Text & "' and �Ͳ��־P=2"
        '�w�Ͳ�����(�i�ǳƾP�f)���q����ӵ���
        intAccount(1) = rsOrder_Detail.RecordCount
        
        '3.�@������s�q����Ӫ��P�f�־P�P�Ͳ��־P���
        Do Until rsOrder_Detail.EOF
            rsOrder_Detail.Fields("�P�f�־P") = 1
            rsOrder_Detail.Fields("�Ͳ��־P") = 3
            rsOrder_Detail.Update
            
            '4.��s���~�w�s�έp����'�P�f�q'����
            '�Y�ҥͲ����ƶq�O�ΨӸɨ��w���s�q(�Ȥᬰ�����q�q���ѥ��������q����), _
             �h���ݧ�s�P�f�q����
            If txtSell(2).Text <> "C000" Then
                rsProdtStore.MoveFirst
                rsProdtStore.Find "���~�s��='" & rsOrder_Detail.Fields("���~�s��") & "'", , 1
                If rsProdtStore.EOF = False Then
                    rsProdtStore.Fields("�P�f�q") = rsProdtStore.Fields("�P�f�q") + rsOrder_Detail.Fields("�q�f�ƶq")
                    rsProdtStore.Fields("�{�s�q") = rsProdtStore.Fields("�W���w�s�q") + rsProdtStore.Fields("�Ͳ��q") - rsProdtStore.Fields("�P�f�q")
                    rsProdtStore.Fields("�Ȧs�q") = rsProdtStore.Fields("�{�s�q")
                    rsProdtStore.Update
                End If
            End If
            rsOrder_Detail.MoveNext
        Loop
        rsOrder_Detail.Filter = adFilterNone
                
        '5.��q����Ӫ����ƻP�i�ǳƾP�f���q����ӵ��Ƥ@��, _
           �h�i��s�q��D�ɪ��P�f�־P����
        If intAccount(0) = intAccount(1) Then
            rsOrder.Fields("�P�f�־P") = 1
            rsOrder.Update
        End If
    End If
End Sub
