VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCheck 
   Caption         =   "���~�L�s��"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     ��"
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   31
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�x     �s"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   22
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�^�D�e��"
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   21
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame fraComputer 
      Caption         =   "�q���L�I��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   5895
      Begin VB.CommandButton cmdMove 
         Caption         =   "<<"
         Height          =   355
         Index           =   1
         Left            =   600
         TabIndex        =   29
         Top             =   4200
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "|<"
         Height          =   355
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   4200
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">|"
         Height          =   355
         Index           =   3
         Left            =   5160
         TabIndex        =   27
         Top             =   4200
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">>"
         Height          =   355
         Index           =   2
         Left            =   4800
         TabIndex        =   26
         Top             =   4200
         Width           =   350
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "��     ��"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "�x     �s"
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "��     ��"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtComputer 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3840
         TabIndex        =   20
         Top             =   1200
         Width           =   1600
      End
      Begin VB.TextBox txtComputer 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   19
         Top             =   1200
         Width           =   1600
      End
      Begin VB.TextBox txtComputer 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgComputer 
         Height          =   2415
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4260
         _Version        =   393216
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
         Left            =   960
         TabIndex        =   30
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label lblComputer 
         AutoSize        =   -1  'True
         Caption         =   "��ڽL�s�q"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   4080
         TabIndex        =   17
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label lblComputer 
         AutoSize        =   -1  'True
         Caption         =   "�����w�s�q"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   16
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label lblComputer 
         AutoSize        =   -1  'True
         Caption         =   "���~�W��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   15
         Top             =   960
         Width           =   840
      End
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
      Index           =   1
      Left            =   1560
      TabIndex        =   12
      Top             =   1200
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
      Index           =   0
      Left            =   1560
      TabIndex        =   11
      Top             =   720
      Width           =   2500
   End
   Begin MSDataListLib.DataCombo dcboStaff 
      Height          =   360
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   720
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     ��"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtCheck 
      Alignment       =   1  '�a�k���
      DataField       =   "�L�I���"
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
      TabIndex        =   6
      Top             =   1680
      Width           =   2500
   End
   Begin VB.TextBox txtCheck 
      DataField       =   "�L�I�H���s��"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   2500
   End
   Begin VB.TextBox txtCheck 
      DataField       =   "�L�I�s��"
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
      Top             =   240
      Width           =   2500
   End
   Begin MSDataListLib.DataCombo dcboStaff 
      Height          =   360
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtCheck 
      DataField       =   "�]�֤H���s��"
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   2500
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "�]�֤H��"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "�L�I���"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "�L�I�H��"
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
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "�L�I�s��"
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
      Width           =   960
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

'���~�L�s���s��
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '�D�ɭק�
            control_status True
            fraComputer.Visible = True
            
        Case 1      '�^�D�e��
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            '�Y��N��ڽL�I���G�O���U��,�h�i�i��U���������Ҫ��إ�
            Set rsComputer = New ADODB.Recordset
            sql_rsComputer = "select * from �q���L�I��"
            rsComputer.Open sql_rsComputer, cn, adOpenDynamic, adLockOptimistic
            If rsComputer.Fields("��ڽL�s�q") <> "" Then
                MsgBox ("�ЫإߤU����������!!")
                frmMDIMain.mnuInitial.Enabled = True
            End If
            
        Case 2      '�D���x�s
            For i = 0 To 1
                rsStaff.MoveFirst
                rsStaff.Find "�m�W='" & dcboStaff(i).Text & "'", , 1
                If rsStaff.EOF Then
                    MsgBox "�п��" & lblCheck(i + 1).Caption & "", , "�����q�q���ѥ��������q"
                    dcboStaff(i).SetFocus
                    Exit Sub
                Else
                    txtCheck(i + 1).Text = dcboStaff(i).BoundText
                    rsCheck.Fields(i + 1) = txtCheck(i + 1).Text
                End If
            Next i
            
            If txtCheck(3).Text = "" Then
                MsgBox "�ж�g�L�I���!", , "�����q�q���ѥ��������q"
                txtCheck(3).SetFocus
                Exit Sub
            Else
                If IsDate(txtCheck(3).Text) = False Then
                    MsgBox "�z��J������榡���~,�ЦA�T�{�@�M!", , "�����q�q���ѥ��������q"
                    txtCheck(3).SetFocus
                    Exit Sub
                Else
                    rsCheck.Update
                    control_status False
                End If
            End If
                    
            '�Y�O�s�W���A,�h�N��ڽL�s�q���ȹw�]�������w�s�q����
            If Add_flag = 1 Then
                '�@����ڽL�s�q���w�]��
                Do Until rsComputer.EOF
                    rsComputer.Fields("��ڽL�s�q") = rsComputer.Fields("�����w�s�q")
                    rsComputer.MoveNext
                Loop
                Call cmdMove_Click(0)
            End If
            
            Add_flag = 0

        Case 3      '����
            '��ƥ��s�W����,���U�����s�h�^��D�e��
            If Add_flag = 1 Then
                Call cmdEdit_Click(1)
            Else
                control_status False
            End If
            Add_flag = 0
    End Select
End Sub

'��ƪ�����
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsComputer)
    lblRecord.Caption = "�L�I���G�G" & intRecord & "/" & intTotal
    
    '�ǥѸ�ƪ�����,���O���i�H��ܩ��r�����
    Call dgComputer_Click
End Sub

'�q���L�I���s��
Private Sub cmdSubEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '���ӭק�
            subcontrol_status True
            
        Case 1      '�����x�s
            rsComputer.Fields("��ڽL�s�q") = txtComputer(2).Text
            '�N�L�I�t���q�g�J��Ʈw
            rsComputer.Fields("�L�I�t���q") = rsComputer.Fields("�����w�s�q") - rsComputer.Fields("��ڽL�s�q")
            rsComputer.Update
            
            subcontrol_status False
            
        Case 2      '���Ө���
            subcontrol_status False
            Call rsComputer_Cnn
    End Select
End Sub

'��DataGrid���O����ܩ��r�����
Private Sub dgComputer_Click()
    For i = 0 To 2
        If rsComputer.Fields(i + 2) <> "" Then
            txtComputer(i).Text = rsComputer.Fields(i + 2)
        Else
            '�Y��ڽL�s�q�٥�����,�h���H0�N��
            txtComputer(i).Text = 0
        End If
    Next
End Sub

'���~�L�s���s��
Private Sub Form_Load()
    '1.�}��RecordSet
    Set rsCheck = New ADODB.Recordset
    sql_rsCheck = "select * from ���~�L�s��"
    rsCheck.Open sql_rsCheck, cn, adOpenDynamic, adLockOptimistic
    
    '2.�N��ƫ��w����椤���P������
    For Each oText In Me.txtCheck
        Set oText.DataSource = rsCheck
    Next
    
    '3.�s�@dcboStaff(�s�׸�ƮɥΤ�)
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from ���u��ƪ�"
    rsStaff.Open sql_rsStaff, cn, adOpenKeyset, adLockReadOnly
    For i = 0 To 1
        Set dcboStaff(i).DataSource = rsCheck
        Set dcboStaff(i).RowSource = rsStaff
        dcboStaff(i).ListField = "�m�W"
        dcboStaff(i).BoundColumn = "���u�s��"
    Next
    dcboStaff(0).DataField = "�L�I�H���s��"
    dcboStaff(1).DataField = "�]�֤H���s��"
    
    '4.�ˬd���~�L�s����ƬO�_����
    '������h�i�J�s�W���A
    If txtCheck(1).Text = "" And txtCheck(2).Text = "" Then
        '�s�W���X��
        Add_flag = 1
        For i = 0 To 1
            dcboStaff(i).Text = "--�п�ܭt�d�H��--"
        Next i
        control_status True
    Else
        '�_�h�����N��Ƨe�{�X��
        control_status False
        Call rsComputer_Cnn
    End If
    
    '5.��檺���e�]�w
    frmCheck.Height = 7755
    frmCheck.Width = 6795
    
End Sub

'�q���L�I���s��
Private Sub rsComputer_Cnn()
    '1.�}�ҹq���L�I��
    Set rsComputer = New ADODB.Recordset
    sql_rsComputer = "select * from �q���L�I�� where �L�I�s��='" & txtCheck(0).Text & "' order by ���~�s��"
    rsComputer.Open sql_rsComputer, cn, adOpenDynamic, adLockOptimistic
          
    '2.�N��ƫ��w��DataGrid,�����ýL�I�s���P�L�I�t���q���
    Set dgComputer.DataSource = rsComputer
    For i = 0 To 1
        dgComputer.Columns.Item(i).Visible = False
    Next i
    dgComputer.Columns.Item(5).Visible = False
    dgComputer.Columns.Item(2).Width = 2000
    For i = 3 To 4
        dgComputer.Columns.Item(i).Alignment = dbgRight
    Next i
    
    '3.�]�w�����J��������A
    subcontrol_status False
    
    '4.���O����ܩ��r�����
    Call cmdMove_Click(0)
End Sub

'��ƪ��s��
Private Sub txtCheck_Change(Index As Integer)
    Select Case Index
        Case 0
            '�����~�L�s��P�q���L�I����ƥi�H�s��
            Call rsComputer_Cnn
            
        Case 1
            '��ܥX�L�I�H���W��
            Call txtStaff_Cnn(1)
            
        Case 2
            '��ܥX�]�֤H���W��
            Call txtStaff_Cnn(2)
    End Select
End Sub

'���~�L�s���P�����󪬺A����
Private Sub control_status(control_flag As Boolean)
    txtCheck(0).Enabled = False
    For i = 0 To 1
        txtStaff(i).Enabled = False
        txtStaff(i).Visible = Not control_flag
        dcboStaff(i).Visible = control_flag
    Next
    txtCheck(3).Enabled = control_flag
    
    '�D�ɭק�,�^�D�e���s
    For i = 0 To 1
        cmdEdit(i).Visible = Not control_flag
    Next i
    
    '�D���x�s�s,�����s
    For i = 2 To 3
        cmdEdit(i).Visible = control_flag
    Next i
    
    '�ج[fraComputer�����A
    fraComputer.Visible = Not control_flag
    fraComputer.Enabled = Not control_flag
End Sub

'�q���L�I���P�����󪬺A����
Private Sub subcontrol_status(subcontrol_flag As Boolean)
    For i = 0 To 1
        txtComputer(i).Enabled = False
    Next i
    txtComputer(2).Enabled = subcontrol_flag
    dgComputer.Enabled = Not subcontrol_flag
    
    '���ӭק�s
    cmdSubEdit(0).Visible = Not subcontrol_flag
    '�����x�s,�����s
    For i = 1 To 2
        cmdSubEdit(i).Visible = subcontrol_flag
    Next i
    
    For i = 0 To 3
        cmdEdit(i).Enabled = Not subcontrol_flag
        cmdMove(i).Enabled = Not subcontrol_flag
    Next i
End Sub

'txtStaff�����
Private Sub txtStaff_Cnn(dcbostaff_index As Integer)
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from ���u��ƪ�"
    rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
    
    rsStaff.MoveFirst
    rsStaff.Find "���u�s��='" & txtCheck(dcbostaff_index).Text & "'", , adSearchForward, 1
    If rsStaff.EOF = False Then
        txtStaff(dcbostaff_index - 1).Text = rsStaff.Fields("�m�W")
    End If
End Sub
