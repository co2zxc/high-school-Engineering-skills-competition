VERSION 5.00
Begin VB.Form frmInfo 
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6120
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   840
      TabIndex        =   13
      Top             =   2520
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   2520
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   5160
      TabIndex        =   11
      Top             =   2520
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   4800
      TabIndex        =   10
      Top             =   2520
      Width           =   350
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�^�D�e��"
      Height          =   375
      Index           =   5
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     ��"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�x     �s"
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�R     ��"
      Height          =   375
      Index           =   4
      Left            =   4440
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�s     �W"
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     ��"
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
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
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
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
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   1815
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
      Left            =   1200
      TabIndex        =   14
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   75
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   75
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsInfo As ADODB.Recordset       '�@�Ϊ�RecordSet����
Dim sql_Area As String              '�N���}�Ҧa�Ϫ��r���ܼ�
Dim sql_Depart As String            '�N���}�ҳ����������r���ܼ�
Dim sql_Duty As String              '�N���}��¾�ٵ��Ū��r���ܼ�
Dim strData_show As String          '��ܩ�lblRecord���r��

'��ƪ��s��
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '�x�s
            If txtFields(1).Text <> "" Then
                rsInfo.Update
            Else
                MsgBox "�ж�g" & lblLabel(1).Caption & "!", , "�����q�q���ѥ��������q"
                txtFields(1).SetFocus
                Exit Sub
            End If
            Call cmdMove_Click(3)
            control_status False

        Case 1      '����
            rsInfo.CancelUpdate
            If rsInfo.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                For Each oText In Me.txtFields
                    Set oText.DataSource = rsInfo
                Next
                
                control_status False
            Else
                Call cmdEdit_Click(5)
                MsgBox "�ثe�õL����" & strData_show & "!", , "�����q�q���ѥ��������q"
            End If

        Case 2      '�s�W
            Add_flag = 1
            
            Select Case Menu_index
                Case 1
                    Call Add_No(rsInfo, 6, 2)
                Case 2
                    Call Add_No(rsInfo, 7, 2)
                Case 3
                    Call Add_No(rsInfo, 8, 2)
            End Select
            
            control_status True
            txtFields(0).Text = Main_No
            
        Case 3      '�ק�
            control_status True
            
        Case 4      '�R��
            If MsgBox("�T�w�n�R�������O��??", 32 + vbYesNo, "�����q�q���ѥ��������q") = vbYes Then
                rsInfo.Delete
                If rsInfo.RecordCount <> 0 Then
                    Call cmdMove_Click(2)
                Else
                    Call cmdEdit_Click(5)
                    MsgBox "�ثe�õL����" & strData_show & "!", , "�����q�q���ѥ��������q"
                End If
            End If
            
        Case 5      '���}
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            '�N�D��椤����ܶ��ش_��
            frmMDIMain.mnuArea.Enabled = True

    End Select
End Sub

'��ƪ�����
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsInfo)
    lblRecord.Caption = strData_show & "�G" & intRecord & "/" & intTotal
End Sub

'��ƪ��s��
Private Sub Form_Load()
    '1.�HMenu_index���ϧO,�Ӷ}�Ҥ��PRecordSet
    Set rsInfo = New ADODB.Recordset
    Select Case Menu_index
        Case 1      '����������
            sql_Depart = "select * from ���������� ORDER BY �����s��"
            rsInfo.Open sql_Depart, cn, adOpenDynamic, adLockOptimistic
            
            txtFields(0).DataField = "�����s��"
            txtFields(1).DataField = "�����W��"
            frmInfo.Caption = "����������"
            lblLabel(0).Caption = "�����s��"
            lblLabel(1).Caption = "�����W��"
            strData_show = "�������"
            
        Case 2      '¾�ٵ��Ū�
            sql_Duty = "select * from ¾�ٵ��Ū� ORDER BY ¾�ٽs��"
            rsInfo.Open sql_Duty, cn, adOpenDynamic, adLockOptimistic
            
            txtFields(0).DataField = "¾�ٽs��"
            txtFields(1).DataField = "¾��"
            frmInfo.Caption = "¾�ٵ��Ū�"
            lblLabel(0).Caption = "¾�ٽs��"
            lblLabel(1).Caption = "¾��"
            strData_show = "¾�ٸ��"
            
        Case 3      '�a�Ϫ�
            sql_Area = "select * from �a�Ϫ� ORDER BY �a�Ͻs��"
            rsInfo.Open sql_Area, cn, adOpenDynamic, adLockOptimistic
            
            txtFields(0).DataField = "�a�Ͻs��"
            txtFields(1).DataField = "�a�ϦW��"
            frmInfo.Caption = "�a�Ϫ�"
            lblLabel(0).Caption = "�a�Ͻs��"
            lblLabel(1).Caption = "�a�ϦW��"
            strData_show = "�a�ϸ��"
            
    End Select
    
    '2.�N��ƫ��w��P������
    For Each oText In Me.txtFields
        Set oText.DataSource = rsInfo
    Next
    
    '3.�ˬd�O�_�����
    If rsInfo.RecordCount <> 0 Then
        '�����J�ɪ����A�]�w
        control_status False
        Call cmdMove_Click(0)
    Else
        Call cmdEdit_Click(2)
    End If
  
    '4.��檺���e�]�w
    frmInfo.Height = 3825
    frmInfo.Width = 6240
End Sub

'��������A�]�w
Private Sub control_status(control_flag As Boolean)
    txtFields(0).Enabled = False
    txtFields(1).Enabled = control_flag
    '�x�s,�����s
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next i
    '�s�W,�ק�,�R��,�^�D�e���s
    For i = 2 To 5
        cmdEdit(i).Visible = Not control_flag
    Next i
    '���ʶs
    For i = 0 To 3
        cmdMove(i).Enabled = Not control_flag
    Next i
    
End Sub

