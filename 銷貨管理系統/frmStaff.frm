VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmStaff 
   Caption         =   "���u��ƪ�"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   9825
   Begin VB.CommandButton cmdRelation 
      Caption         =   "�p���H���"
      Height          =   375
      Left            =   8160
      TabIndex        =   36
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   960
      TabIndex        =   32
      Top             =   3720
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   600
      TabIndex        =   31
      Top             =   3720
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   7320
      TabIndex        =   30
      Top             =   3720
      Width           =   315
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   6960
      TabIndex        =   29
      Top             =   3720
      Width           =   350
   End
   Begin VB.Frame fraMarry 
      Caption         =   "�B�ê��p"
      Height          =   615
      Left            =   5040
      TabIndex        =   28
      Top             =   720
      Width           =   2055
      Begin VB.OptionButton optMarry 
         Caption         =   "�w�B"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optMarry 
         Caption         =   "���B"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraSex 
      Caption         =   "�ʧO"
      Height          =   615
      Left            =   1560
      TabIndex        =   25
      Top             =   720
      Width           =   2055
      Begin VB.OptionButton optSex 
         Caption         =   "�k"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optSex 
         Caption         =   "�k"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
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
      TabIndex        =   24
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     ��"
      Height          =   375
      Index           =   3
      Left            =   8160
      TabIndex        =   23
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�s     �W"
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   22
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�R     ��"
      Height          =   375
      Index           =   4
      Left            =   8160
      TabIndex        =   21
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�x     �s"
      Height          =   375
      Index           =   0
      Left            =   8160
      TabIndex        =   20
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     ��"
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   19
      Top             =   2040
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo dcboArea 
      Height          =   390
      Left            =   1560
      TabIndex        =   18
      Top             =   2880
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
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�^�D�e��"
      Height          =   375
      Index           =   5
      Left            =   8160
      TabIndex        =   17
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtStaff 
      DataField       =   "�a�Ͻs��"
      Height          =   375
      Index           =   7
      Left            =   1560
      TabIndex        =   16
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtStaff 
      DataField       =   "��}"
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
      Left            =   2880
      TabIndex        =   15
      Top             =   2880
      Width           =   4215
   End
   Begin VB.TextBox txtStaff 
      DataField       =   "�Ǿ�"
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
      Left            =   5040
      TabIndex        =   14
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
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
      TabIndex        =   13
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
      Alignment       =   1  '�a�k���
      DataField       =   "�X�ͤ��"
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
      Left            =   5040
      TabIndex        =   12
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
      Alignment       =   1  '�a�k���
      DataField       =   "�����Ҧr��"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
      DataField       =   "�m�W"
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
      Left            =   5040
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
      Alignment       =   1  '�a�k���
      DataField       =   "���u�s��"
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
      TabIndex        =   9
      Top             =   240
      Width           =   2055
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
      TabIndex        =   33
      Top             =   3720
      Width           =   5655
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "�B�ê��p"
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
      Index           =   8
      Left            =   3960
      TabIndex        =   8
      Top             =   840
      Width           =   960
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "�Ǿ�"
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
      Index           =   7
      Left            =   4440
      TabIndex        =   7
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "��}"
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
      Index           =   6
      Left            =   960
      TabIndex        =   6
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label lblStaff 
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
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "�X�ͤ��"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "�����Ҧr��"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "�ʧO"
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
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "�m�W"
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
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "���u�s��"
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
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ƪ��s��
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '�x�s
            For i = 1 To 6
                If txtStaff(i).Text = "" Then
                    MsgBox "��p!��Ƥ���ť�!", , "�����q�q���ѥ��������q"
                    txtStaff(i).SetFocus
                    Exit Sub
                Else
                    If IsDate(txtStaff(3).Text) = False Then
                        MsgBox "�z��J������榡���~,�ЦA�T�{�@�M!", , "�����q�q���ѥ��������q"
                        txtStaff(3).SetFocus
                        Exit Sub
                    End If
                End If
            Next i
            
      
                    rsArea.MoveFirst
                    rsArea.Find "�a�ϦW��='" & dcboArea.Text & "'", , 1
                    If rsArea.EOF Then
                        MsgBox "�п�ܥ��T���a�ϦW��!", , "�����q�q���ѥ��������q"
                        dcboArea.SetFocus
                        Exit Sub
                    Else
                        txtStaff(7).Text = dcboArea.BoundText
                        rsStaff.Update
                        control_status False
                    End If
             
            
        Case 1      '����
            rsStaff.CancelUpdate
            If rsStaff.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                For Each oText In Me.txtStaff
                    Set oText.DataSource = rsStaff
                Next
                dcboArea.BoundText = txtStaff(7).Text
                
                control_status False
            Else
                Call cmdEdit_Click(5)
                MsgBox "�ثe�õL������u���!", , "�����q�q���ѥ��������q"
            End If
            Add_flag = 0

        Case 2      '�s�W
            Add_flag = 1
            Call Add_No(rsStaff, 2, 1)
            txtStaff(0).Text = Main_No
            control_status True
            dcboArea.Text = "--�п��--"
            optSex(0).Value = True
            optMarry(0).Value = True
            
        Case 3      '�ק�
            control_status True
            
        Case 4      '�R��
            If MsgBox("�T�w�n�R�������O��??", 32 + vbYesNo, "�����q�q���ѥ��������q") = vbYes Then
                rsStaff.Delete
                If rsStaff.RecordCount <> 0 Then
                    Call cmdMove_Click(2)
                Else
                    Call cmdEdit_Click(5)
                    MsgBox "�ثe�w�L������u���!", , "�����q�q���ѥ��������q"
                End If
            End If
            
        Case 5      '�^�D�e��
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            rsStaff.Close
            Set rsStaff = Nothing
            
    End Select
End Sub

'��ƪ�����
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsStaff)
    lblRecord.Caption = "���u��ơG" & intRecord & "/" & intTotal
    
    '�k��
    If rsStaff.Fields("�ʧO") = 1 Then
        optSex(0).Value = True
    Else
        optSex(1).Value = True
    End If
    
    '�w�B
    If rsStaff.Fields("�B�ê��p") = 1 Then
        optMarry(1).Value = True
    Else
        optMarry(0).Value = True
    End If
End Sub

Private Sub cmdRelation_Click()
    '���}�e���O���ثe������
    bm_move = intRecord
    frmRelation.Show 1
End Sub

'��ƪ��s��
Private Sub Form_Load()
    '1.�}�ҭ��u��ƪ�
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from ���u��ƪ�"
    rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
    
    '2.�N��ƫ��w���P������
    For Each oText In Me.txtStaff
        Set oText.DataSource = rsStaff
    Next
        
    '3.�s�@DataCombo
    '�s�@�a�ϦW��dcboArea
    Set rsArea = New ADODB.Recordset
    sql_rsArea = "select * from �a�Ϫ�"
    rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
    
    Set dcboArea.DataSource = rsStaff
    dcboArea.DataField = "�a�Ͻs��"
    Set dcboArea.RowSource = rsArea
    dcboArea.ListField = "�a�ϦW��"
    dcboArea.BoundColumn = "�a�Ͻs��"
    rsStaff.MoveFirst
    
    '4.�����J�ɪ����A�]�w
    If rsStaff.RecordCount <> 0 Then
        Call cmdMove_Click(0)
        control_status False
    Else
        Call cmdEdit_Click(2)
    End If
    
    '5.��檺���e�]�w
    frmStaff.Height = 5025
    frmStaff.Width = 9945
    
End Sub

'��������A�]�w
Private Sub control_status(control_flag As Boolean)
    '�T�w���ܪ����A
    '���u�s��,�a�ϦW��,�����W��,¾�٬Ҥ��i�ާ@
    txtStaff(0).Enabled = False
    txtArea.Enabled = False
    '�a�Ͻs��,�����s��,¾�ٽs��������
    For i = 7 To 7
        txtStaff(i).Visible = False
    Next
    
    '���A�i����
    For i = 1 To 6
        txtStaff(i).Enabled = control_flag
    Next
    fraSex.Enabled = control_flag
    fraMarry.Enabled = control_flag
        
    '�a�ϦW��,�����W��,¾��
    dcboArea.Visible = control_flag
    txtArea.Visible = Not control_flag
    
    '�x�s,�����s
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next i
    
    '�s�W,�ק�,�R��,�^�D�e���s
    For i = 2 To 5
        cmdEdit(i).Visible = Not control_flag
    Next i
    
    '��Ʋ���
    For i = 0 To 3
        cmdMove(i).Enabled = Not control_flag
    Next i
    '���u�p���H
    cmdRelation.Visible = Not control_flag
    
End Sub

'�N�ȶ��txtArea txtDepart txtDuty��,�ûPrsStaff�s��
Private Sub txtStaff_Change(Index As Integer)
    Select Case Index
        Case 7      '�a�ϦW�����
            Set rsArea = New ADODB.Recordset
            sql_rsArea = "select * from �a�Ϫ�"
            rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
            
            rsArea.Find "�a�Ͻs��='" & txtStaff(7).Text & "'"
            If rsArea.EOF = False Then
                txtArea.Text = rsArea.Fields("�a�ϦW��")
            End If
    End Select
End Sub
