VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMkCk 
   Caption         =   "�Ͳ��־P��"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8400
   StartUpPosition =   1  '���ݵ�������
   Begin VB.CommandButton cmdReturn 
      Caption         =   "�^�D�e��"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "��     �P"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraDetail 
      Caption         =   "���P�f�����~���"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   600
      TabIndex        =   4
      Top             =   3600
      Width           =   4935
      Begin MSDataGridLib.DataGrid dgDetail 
         Height          =   2175
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3836
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
   End
   Begin VB.ListBox lstNot 
      Height          =   2370
      Left            =   3240
      Style           =   1  '���إ]�t�֨����
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ListBox lstOK 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblInCheck 
      AutoSize        =   -1  'True
      Caption         =   "�����f����ơG"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   1680
   End
   Begin VB.Label lblInCheck 
      AutoSize        =   -1  'True
      Caption         =   "���X�f����ơG"
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
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1680
   End
End
Attribute VB_Name = "frmMkCk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�־P�Ͳ��־P���(�N�Ͳ��q��s�ܮw�s��,�ñN�Ͳ��־P�����O���R��)
Private Sub cmdOK_Click()
    Set rsProdtStore = New ADODB.Recordset
    rsProdtStore.Open "select * from ���~�w�s�έp�� order by ���~�s��", cn, adOpenDynamic, adLockOptimistic
    
    Set rsMake = New ADODB.Recordset
    rsMake.Open "select * from �Ͳ��־P�� order by ���~�s��", cn, adOpenDynamic, adLockOptimistic

    For i = 0 To lstNot.ListCount - 1
        If lstNot.Selected(i) Then
            '1.��X�Ͳ��־P���P�M�椤�Q�Ŀ蠟�ۦP���~�W��
            rsMake.MoveFirst
            rsMake.Find "���~�W��='" & RTrim(lstNot.List(i)) & "'", , 1
            '2.��X�w�s���P�����~�۹��������~�s��
            rsProdtStore.MoveFirst
            rsProdtStore.Find "���~�s��='" & rsMake.Fields(0) & "'", , 1
            If rsProdtStore.EOF = False Then
                '3.��s��Ͳ��q�P�{�s�q,�Ȧs�q
                rsProdtStore.Fields("�Ͳ��q") = rsProdtStore.Fields("�Ͳ��q") + rsMake.Fields("�w�p�Ͳ��q")
                rsProdtStore.Fields("�{�s�q") = rsProdtStore.Fields("�W���w�s�q") + rsProdtStore.Fields("�Ͳ��q") - rsProdtStore.Fields("�P�f�q")
                '���~�T�w�w�Ͳ���,�Ȧs�q�Y����{�s�q
                rsProdtStore.Fields("�Ȧs�q") = rsProdtStore.Fields("�{�s�q")
                rsProdtStore.Update
                '4.��s�q����Ӥ����Ͳ��־P��쬰2,��ܲ��~�w�Ͳ�,�i�i��P�f�F
                rsOrder_Detail.Filter = "���~�s��='" & rsMake.Fields(0) & "' and �Ͳ��־P=1"
                Do Until rsOrder_Detail.EOF
                    rsOrder_Detail.Fields("�Ͳ��־P") = 2
                    
                    rsOrder_Detail.MoveNext
                Loop
                
                '5.�R���Ͳ��־P���۹������O��
                rsMake.Delete
                If rsMake.RecordCount <> 0 Then
                    rsMake.MoveNext
                End If
                rsOrder_Detail.Filter = adFilterNone
            End If
        End If
    Next i
    
    '6.������ܩ�lstOK�PlstNot�������
    Call Form_Load
End Sub

'�^�D�e��
Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmdOK.Enabled = False
    
    Set rsOrder_Detail = New ADODB.Recordset
    rsOrder_Detail.Open "select * from �q����� where �P�f�־P=0 order by ���~�s��", cn, adOpenDynamic, adLockOptimistic
     
    lstOK.Clear
    lstNot.Clear
    pre_prodt = ""
    '�w�����Ͳ�,�B�w�־P�L����ƫh�e�{��lstOK��
    rsOrder_Detail.Filter = "�Ͳ��־P=2"
    Do Until rsOrder_Detail.EOF
        '�Y���~������,�h�u�n�H�@���N��Y�i,�������Ƨe�{
        If pre_prodt <> rsOrder_Detail.Fields("���~�s��") Then
            lstOK.AddItem rsOrder_Detail.Fields("���~�W��")
        End If
        pre_prodt = rsOrder_Detail.Fields("���~�s��")
        rsOrder_Detail.MoveNext
    Loop
    rsOrder_Detail.Filter = adFilterNone
    
    '�u���ͥͲ���,�٦b�Ͳ����q�����
    pre_prodt = ""
    rsOrder_Detail.Filter = "�Ͳ��־P=1"
    Do Until rsOrder_Detail.EOF
        If pre_prodt <> rsOrder_Detail.Fields("���~�s��") Then
            lstNot.AddItem rsOrder_Detail.Fields("���~�W��")
        End If
        pre_prodt = rsOrder_Detail.Fields("���~�s��")
        rsOrder_Detail.MoveNext
    Loop
    rsOrder_Detail.Filter = adFilterNone
       
    rsOrder_Detail.Sort = "�q��s�� asc"
    Set dgDetail.DataSource = rsOrder_Detail
    
    For i = 1 To 2
        dgDetail.Columns.Item(i).Visible = False
    Next i
    
    For i = 7 To 8
        dgDetail.Columns.Item(i).Visible = False
    Next i
        
    For i = 4 To 6
        dgDetail.Columns.Item(i).Visible = False
    Next i
    dgDetail.Columns.Item(5).Visible = True
    dgDetail.Columns.Item(5).Alignment = dbgRight
    dgDetail.Columns.Item(3).Width = 1900
        
End Sub

''�w�Ͳ������'�Q�Ŀ��,�־P�s�~�ϥ�
Private Sub lstNot_Click()
    cmdOK.Enabled = True
End Sub
