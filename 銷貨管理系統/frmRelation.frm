VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelation 
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   8115
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   6120
      TabIndex        =   21
      Top             =   3240
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   6480
      TabIndex        =   20
      Top             =   3240
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   600
      TabIndex        =   19
      Top             =   3240
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   960
      TabIndex        =   18
      Top             =   3240
      Width           =   350
   End
   Begin VB.TextBox txtArea 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     �}"
      Height          =   375
      Index           =   5
      Left            =   6240
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     ��"
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�x     �s"
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�R     ��"
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�s     �W"
      Height          =   375
      Index           =   2
      Left            =   6240
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��     ��"
      Height          =   375
      Index           =   3
      Left            =   6240
      TabIndex        =   14
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtRelation 
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
      Index           =   4
      Left            =   1920
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtRelation 
      DataField       =   "�p���H�q��"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   1800
      Width           =   2500
   End
   Begin VB.TextBox txtRelation 
      DataField       =   "���Y"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   2500
   End
   Begin MSDataListLib.DataCombo dcboArea 
      Height          =   360
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
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
   Begin VB.TextBox txtRelation 
      DataField       =   "�p���H��}"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txtRelation 
      DataField       =   "�p���H�m�W"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   2500
   End
   Begin VB.Label lblStaff 
      BorderStyle     =   1  '��u�T�w
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
      Left            =   1920
      TabIndex        =   23
      Top             =   360
      Width           =   2535
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
      TabIndex        =   22
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      Caption         =   "���Y"
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
      Left            =   1320
      TabIndex        =   12
      Top             =   1395
      Width           =   495
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      Caption         =   "�p���H��}"
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
      Left            =   600
      TabIndex        =   11
      Top             =   2355
      Width           =   1215
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      Caption         =   "�p���H�q��"
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
      TabIndex        =   10
      Top             =   1875
      Width           =   1215
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      Caption         =   "�p���H�m�W"
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
      Left            =   600
      TabIndex        =   9
      Top             =   915
      Width           =   1215
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      Caption         =   "���u�m�W"
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
      Left            =   840
      TabIndex        =   8
      Top             =   435
      Width           =   975
   End
End
Attribute VB_Name = "frmRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ƪ��s��
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '�x�s
            For i = 0 To 3
                If txtRelation(i).Text = "" Then
                    MsgBox "�ж�g" & lblRelation(i + 1).Caption & "!", , "�����q�q���ѥ��������q"
                    txtRelation(i).SetFocus
                    Exit Sub
                Else
                    rsArea.MoveFirst
                    rsArea.Find "�a�ϦW��='" & dcboArea.Text & "'", , 1
                    If rsArea.EOF Then
                        MsgBox "�п�ܦa��!", , "�����q�q���ѥ��������q"
                        dcboArea.SetFocus
                        Exit Sub
                    Else
                        rsStaff.MoveFirst
                        rsStaff.Find "�m�W='" & lblStaff.Caption & "'", , 1
                        rsRelation.Fields("���u�s��") = rsStaff.Fields(0)
                        txtRelation(4).Text = dcboArea.BoundText
                        rsRelation.Update
                    End If
                End If
            Next i
            Call cmdMove_Click(3)
            control_status False

        Case 1      '����
            rsRelation.CancelUpdate
            If rsRelation.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                For Each oText In Me.txtRelation
                    Set oText.DataSource = rsRelation
                Next
                dcboArea.BoundText = txtRelation(4).Text
                control_status False
            Else
                Call cmdEdit_Click(5)
                MsgBox "�ثe�S�����u" & frmStaff.txtStaff(1).Text & "�������p���H�����!", , "�����q�q���ѥ��������q"
            End If

        Case 2      '�s�W
            Add_flag = 1
            rsRelation.AddNew
            control_status True
            dcboArea.Text = "�п��"
            
        Case 3      '�ק�
            control_status True
            dcboArea.Text = txtArea.Text
            
        Case 4      '�R��
            If MsgBox("�T�w�n�R�������O��??", 32 + vbYesNo, "�����q�q���ѥ��������q") = vbYes Then
                rsRelation.Delete
                If rsRelation.RecordCount <> 0 Then
                    rsRelation.MoveNext
                Else
                    Call cmdEdit_Click(5)
                    MsgBox "�ثe�S�����u" & frmStaff.txtStaff(1).Text & "�������p���H�����!", , "�����q�q���ѥ��������q"
                End If
            End If
            
        Case 5      '���}
            '�n�^��frmStaff�e,�NfrmStaff�ثe�����Ʃ�^�h
            intRecord = bm_move
            Unload Me
            rsRelation.Close
            Set rsRelation = Nothing
    End Select
End Sub

'��ƪ�����
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsRelation)
    lblRecord.Caption = "�p���H��ơG" & intRecord & "/" & intTotal
    
End Sub

'��ƪ��s��
Private Sub Form_Load()
    '1.�}�ҭ��u�p���H��ƪ�
    Set rsRelation = New ADODB.Recordset
    sql_rsRelation = "select * from ���u�p���H��ƪ� where ���u�s��='" & frmStaff.txtStaff(0).Text & "'"
    rsRelation.Open sql_rsRelation, cn, adOpenDynamic, adLockOptimistic
    
    '2.�N��ƫ��w�ܪ�椤���P������W
    For Each oText In Me.txtRelation
        Set oText.DataSource = rsRelation
    Next
    
    '���u�m�W�ӷ�
    lblStaff.Caption = frmStaff.txtStaff(1).Text
    Me.Caption = "���u " & lblStaff.Caption & " ���p���H�������"
    
    '3.�ˬd�O�_�����
    '�Y�L��ƴN�i�J�s�W���A
    If rsRelation.RecordCount = 0 Then
        Call cmdEdit_Click(2)
    Else
        '�]�w�����J�ɱ�������A
        control_status False
        Call cmdMove_Click(0)
    End If
       
    '4.�s�@DataCombo
    '�s�@dcboArea
    Set rsArea = New ADODB.Recordset
    sql_rsArea = "select * from �a�Ϫ�"
    rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
    
    Set dcboArea.DataSource = rsRelation
    dcboArea.DataField = "�a�Ͻs��"
    Set dcboArea.RowSource = rsArea
    dcboArea.ListField = "�a�ϦW��"
    dcboArea.BoundColumn = "�a�Ͻs��"
        
End Sub

'��������A�]�w
Private Sub control_status(control_flag As Boolean)
    '�T�w���ܪ����A
    txtArea.Enabled = False
    txtRelation(4).Visible = False
    
    '���A�i����
    For i = 0 To 4
        txtRelation(i).Enabled = control_flag
    Next
    dcboArea.Visible = control_flag
    txtArea.Visible = Not control_flag
    
    '�x�s,�����s
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next i
    '�s�W,�ק�,�R��,���}�s
    For i = 2 To 5
        cmdEdit(i).Visible = Not control_flag
    Next i
    
    '���ʶs
    For i = 0 To 3
        cmdMove(i).Enabled = Not control_flag
    Next i
End Sub

'���W����ܤ����txtArea��P��Ƴs��
Private Sub txtRelation_Change(Index As Integer)
    Select Case Index
        Case 4      '�a�ϦW�����
            Set rsArea = New ADODB.Recordset
            sql_rsArea = "select * from �a�Ϫ�"
            rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
            
            rsArea.Find "�a�Ͻs��='" & txtRelation(4).Text & "'"
            If rsArea.EOF = False Then
                txtArea.Text = rsArea.Fields("�a�ϦW��")
                dcboArea.Text = rsArea.Fields("�a�ϦW��")
            End If
    End Select
End Sub
