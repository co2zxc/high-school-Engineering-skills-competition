VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOrder 
   Caption         =   "�q�f��"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11880
   WindowState     =   2  '�̤j��
   Begin VB.CommandButton cmdSell 
      Caption         =   "���;P�f��"
      Height          =   375
      Left            =   6720
      TabIndex        =   41
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�^�D�e��"
      Height          =   375
      Index           =   5
      Left            =   7920
      TabIndex        =   31
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraOrder 
      Caption         =   "�q��D��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   2160
      TabIndex        =   15
      Top             =   600
      Width           =   6975
      Begin VB.CommandButton cmdMove 
         Caption         =   "<<"
         Height          =   355
         Index           =   1
         Left            =   1560
         TabIndex        =   40
         Top             =   2280
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "|<"
         Height          =   355
         Index           =   0
         Left            =   1200
         TabIndex        =   39
         Top             =   2280
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">|"
         Height          =   355
         Index           =   3
         Left            =   4920
         TabIndex        =   38
         Top             =   2280
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">>"
         Height          =   355
         Index           =   2
         Left            =   4560
         TabIndex        =   37
         Top             =   2280
         Width           =   350
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�s     �W"
         Height          =   375
         Index           =   2
         Left            =   5640
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "��     ��"
         Height          =   375
         Index           =   3
         Left            =   5640
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�R      ��"
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtOrder 
         DataField       =   "�q��s��"
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
         TabIndex        =   19
         Top             =   360
         Width           =   2685
      End
      Begin VB.TextBox txtOrder 
         Alignment       =   1  '�a�k���
         DataField       =   "�q�f��"
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
         Left            =   1920
         TabIndex        =   17
         Top             =   1800
         Width           =   2685
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�x     �s"
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "��     ��"
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtOrder 
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
         Left            =   1920
         TabIndex        =   18
         Top             =   840
         Width           =   2685
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
         Left            =   1920
         TabIndex        =   33
         Top             =   840
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo dcboStaff 
         Height          =   390
         Left            =   1920
         TabIndex        =   16
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.TextBox txtOrder 
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
         Index           =   2
         Left            =   1920
         TabIndex        =   26
         Top             =   1320
         Width           =   2685
      End
      Begin VB.TextBox txtCustomer 
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
         TabIndex        =   34
         Top             =   1320
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo dcboCust 
         Height          =   390
         Left            =   1920
         TabIndex        =   25
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
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
         Left            =   1920
         TabIndex        =   36
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         Caption         =   "�q��s��"
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
         Left            =   990
         TabIndex        =   30
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         Caption         =   "�g��H"
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
         Left            =   1200
         TabIndex        =   29
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ�W��"
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
         Left            =   990
         TabIndex        =   28
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         Caption         =   "�q�f��"
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
         Index           =   3
         Left            =   1200
         TabIndex        =   27
         Top             =   1920
         Width           =   630
      End
   End
   Begin VB.Frame fraOrDetail 
      Caption         =   "�q�����"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2160
      TabIndex        =   0
      Top             =   3600
      Width           =   6975
      Begin VB.TextBox txtProdt 
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
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cboProdt 
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
         ItemData        =   "frmOrder.frx":0000
         Left            =   240
         List            =   "frmOrder.frx":0002
         TabIndex        =   32
         Top             =   720
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgOrDetail 
         Height          =   2055
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         Caption         =   "�R      ��"
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   14
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "��      ��"
         Height          =   375
         Index           =   3
         Left            =   5640
         TabIndex        =   13
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "�s      �W"
         Height          =   375
         Index           =   2
         Left            =   5640
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "��     ��"
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "�x     �s"
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtOrDetail 
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
         Index           =   3
         Left            =   4440
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtOrDetail 
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
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtOrDetail 
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
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtOrDetail 
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
         Left            =   3240
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblOrDetail 
         AutoSize        =   -1  'True
         Caption         =   "�`     ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   4680
         TabIndex        =   9
         Top             =   480
         Width           =   690
      End
      Begin VB.Label lblOrDetail 
         AutoSize        =   -1  'True
         Caption         =   "�q �f �� �q"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   3360
         TabIndex        =   8
         Top             =   480
         Width           =   960
      End
      Begin VB.Label lblOrDetail 
         AutoSize        =   -1  'True
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   480
         Width           =   630
      End
      Begin VB.Label lblOrDetail 
         AutoSize        =   -1  'True
         Caption         =   "�� �~ �W ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   720
         TabIndex        =   6
         Top             =   480
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bm_detail As Variant            '�O�����ӭק�e���O��
Dim Prodt_Name As String            '�O�����ӭק�ecboProdt����
Dim intpre_NumCk(1) As Integer      '�O���ק�e���q�f�ƶq�P�Ͳ��־P���

'�D�ɸ�ƪ��s��
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0     '�D���x�s
            '�D�ɸ�ƥ�������,���Ӹ�Ƥ~���g
            rsStaff.MoveFirst
            rsStaff.Find "�m�W='" & dcboStaff.Text & "'", , 1
            If rsStaff.EOF Then
                MsgBox "�п�ܸg��H!", , "�����q�q���ѥ��������q"
                dcboStaff.SetFocus
                Exit Sub
            Else
                rsCust.MoveFirst
                rsCust.Find "���q�W��='" & dcboCust.Text & "'", , 1
                If rsCust.EOF Then
                    MsgBox "�п�ܫȤ�!", , "�����q�q���ѥ��������q"
                    dcboCust.SetFocus
                    Exit Sub
                Else
                    If txtOrder(3).Text = "" Then
                        MsgBox "�ж�g�q�f��!", , "�����q�q���ѥ��������q"
                        txtOrder(3).SetFocus
                        Exit Sub
                    Else
                        If IsDate(txtOrder(3).Text) = False Then
                            MsgBox "�z��J������榡���~,�ЦA�T�{�@�M!", , "�����q�q���ѥ��������q"
                            txtOrder(3).SetFocus
                            Exit Sub
                        Else
                            txtOrder(1).Text = dcboStaff.BoundText
                            txtOrder(2).Text = dcboCust.BoundText
                            If Add_flag = 1 Then
                                '�s�W�@���O����,�P�f�־P��쪺�w�]�Ȭ�0
                                rsOrder.Fields("�P�f�־P") = 0
                            End If
                            rsOrder.Update
                            Call cmdMove_Click(3)
                            control_status False
                            cmdSell.Visible = True
                        End If
                    End If
                End If
            End If
            
            '�D�ɦb"�s�W"���A�U,�x�s��Y�i�J���ӷs�W���A
            If Add_flag = 1 Then
                Call cmdSubEdit_Click(2)
            End If
            
            Add_flag = 0

        Case 1     '�D�ɨ���
            rsOrder.CancelUpdate
            If rsOrder.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                dcboStaff.BoundText = txtOrder(1).Text
                dcboCust.BoundText = txtOrder(2).Text
                For Each oText In Me.txtOrder
                    Set oText.DataSource = rsOrder
                Next
                control_status False
                cmdSell.Visible = True
            Else
                Call cmdEdit_Click(5)
            End If
            
            Add_flag = 0
            
        Case 2     '�D�ɷs�W
            '�N�D�ɻP���ӽs�׶s��ܥX��,�A�����A������
           ' For i = 0 To 4
           '     cmdEdit(i).Enabled = True
           '     cmdSubEdit(i).Enabled = True
           ' Next i
            
            Add_flag = 1
            Call Add_No(rsOrder, 3, 1)
            control_status True
            cmdSell.Visible = False
            txtOrder(0).Text = Main_No
            txtOrder(3).Text = Date
            dcboStaff.Text = "�п�ܸg��H"
            dcboCust.Text = "--�п�ܫȤ�--"

        Case 3     '�D�ɭק�
            control_status True
            '�ק�ɫȤ���줣�i���,�G��ܥXtxtCustomer
            dcboCust.Visible = False
            txtCustomer.Visible = True
            fraOrDetail.Visible = True
            fraOrDetail.Enabled = False
            
            cmdSell.Visible = False
            
        Case 4     '�D�ɧR��
            If MsgBox("�T�w�n�R�������O��??", 32 + vbYesNo, "�����q�q���ѥ��������q") = vbYes Then
                rsOrder.Delete
                If rsOrder.RecordCount <> 0 Then
                    Call cmdMove_Click(2)
                Else
                    Call cmdEdit_Click(5)
                    MsgBox "�ثe�w�L����q����!", , "�����q�q���ѥ��������q"
                End If
            End If
            
        Case 5     '�^�D�e��
            Set rsOrder_Detail = New ADODB.Recordset
            rsOrder_Detail.Open "select * from �q����� where �Ͳ��־P=0 order by ���~�s��", cn, adOpenDynamic, adLockOptimistic
            
            '�Y�ثe�٦����Ͳ������~,�Y�|�X�{�Ͳ���
            If rsOrder_Detail.RecordCount <> 0 Then
                frmMake.Show 1
            End If
                
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            Call frmMDIMain.mnuSell_status
    End Select
End Sub

'��ƪ�����
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsOrder)
    lblRecord.Caption = "�q���ơG" & intRecord & "/" & intTotal
    
    Call cmdSell_status
End Sub

'���;P�f��
Private Sub cmdSell_Click()
    cmdSell_flag = 1
    bm_move = intRecord
    frmSell.Show
End Sub

'���Ӹ�ƪ��s��
Private Sub cmdSubEdit_Click(Index As Integer)
    Select Case Index
        Case 0    '�����x�s
            '�ˬd��ƬO�_��g����
            rsProdt.MoveFirst
            rsProdt.Find "���~�W��='" & cboProdt.Text & "'", , 1
            If rsProdt.EOF Then
                MsgBox "�п�ܭq�ʪ����~!", , "�����q�q���ѥ��������q"
                cboProdt.SetFocus
                Exit Sub
            Else
                For i = 1 To 2
                    If txtOrDetail(i).Text = "" Then
                        MsgBox "�ж�g" & lblOrDetail(i).Caption & "!", , "�����q�q���ѥ��������q"
                        txtOrDetail(i).SetFocus
                        Exit Sub
                    Else
                        If IsNumeric(txtOrDetail(i).Text) = False Then
                            MsgBox "�ж�g�Ʀr�榡�����!", , "�����q�q���ѥ��������q"
                            txtOrDetail(i).SetFocus
                            Exit Sub
                        End If
                    End If
                Next i
                
                '�`��
                txtOrDetail(3).Text = Int(txtOrDetail(1).Text) * Int(txtOrDetail(2).Text)
                
                '���U�s�W�s�᪺�x�s���A,�h�|�I�sAddNew
                If Edit_flag <> 1 Then
                    rsOrder_Detail.AddNew
                    rsOrder_Detail.Fields(0) = txtOrder(0).Text
                    rsOrder_Detail.Fields(1) = Sub_No
                    '�P�f�־P�P�Ͳ��־P���w�]��0(�Y���־P)
                    For i = 7 To 8
                        rsOrder_Detail.Fields(i) = 0
                    Next i
                End If
                
                '���~�s��
                txtOrDetail(0).Text = rsProdt.Fields(0)
                rsOrder_Detail.Fields("���~�s��") = txtOrDetail(0).Text
                For i = 1 To 3
                   rsOrder_Detail.Fields(i + 3) = txtOrDetail(i).Text
                Next i
                rsOrder_Detail.Fields("���~�W��") = cboProdt.Text
                rsOrder_Detail.Update
            
                subcontrol_status False
                cmdSell.Visible = True
                lblOrDetail(3).Visible = True
                txtOrDetail(3).Visible = True
            End If
                
            Set rsMake = New ADODB.Recordset
            rsMake.Open "select * from �Ͳ��־P�� order by  ���~�s��", cn, adOpenDynamic, adLockOptimistic
            
            '�ק��
            If Edit_flag = 1 Then
                rsOrder_Detail.Bookmark = bm_detail
                
                '1.�Y�q�f�ƶq�Q���F
                If intpre_NumCk(0) <> Int(rsOrder_Detail.Fields("�q�f�ƶq")) Then
                    '2.���ɦp�G�w�g���ͥͲ���F(�Ͳ��־P=1),�ƦܬO�w�������~���Ͳ��F(�Ͳ��־P=2)
                    If intpre_NumCk(1) = 1 Or intpre_NumCk(1) = 2 Then
                        If rsMake.RecordCount <> 0 Then
                            rsMake.MoveFirst
                        End If
                        rsMake.Find "���~�s��='" & rsOrder_Detail.Fields("���~�s��") & "'", , 1
                        
                        '3.�Y�q�f�ƶq�W�h�F,����W�[���ƶq��n�[�J�Ͳ��椤
                        If Sgn(rsOrder_Detail.Fields("�q�f�ƶq") - intpre_NumCk(0)) = 1 Then
                            If rsMake.EOF Then
                                rsMake.AddNew
                                rsMake.Fields(0) = rsOrder_Detail.Fields("���~�s��")
                                rsMake.Fields(1) = rsOrder_Detail.Fields("���~�W��")
                                '�ҭn�Ͳ����q�f�ƶq������h�ק�e���ƶq
                                rsMake.Fields(2) = Int(rsOrder_Detail.Fields("�q�f�ƶq") - intpre_NumCk(0))
                            Else
                                rsMake.Fields(2) = rsMake.Fields(2) + Int(rsOrder_Detail.Fields("�q�f�ƶq") - intpre_NumCk(0))
                            End If
                        Else
                            '4.�Y�q�f�ƶq��֤F,����p�G�Ͳ��椤�������~,�h�i�N��֪����~���� _
                               �p�G���~�w�g�Ͳ��F,���N���H�w�s�q�x�s��ܮw��
                            If rsMake.EOF = False Then
                                rsMake.Fields(2) = rsMake.Fields(2) - Abs(rsOrder_Detail.Fields("�q�f�ƶq") - intpre_NumCk(0))
                            Else
                                Call dgOrDetail_refresh
                                '�M�����Ӫ��X�Э�
                                Edit_flag = 0
                                Exit Sub
            
                            End If
                        End If
                        rsMake.Update
                        
                        '5.�ñN�Ͳ��־P���אּ1,�H�K�b�i�J�Ͳ����q
                        rsOrder_Detail.Fields("�Ͳ��־P") = 1
                        rsOrder_Detail.Update
                    End If
                End If
            End If
                            
            Call dgOrDetail_refresh
            Call cmdSell_status
            '�M�����Ӫ��X�Э�
            Edit_flag = 0
         
        Case 1    '���Ө���
            '���Ӧb�Ĥ@����"�s�W"���A�U(���Ӥ������Ȯ�)
            If rsOrder_Detail.RecordCount = 0 Then
                If MsgBox("�q����Ӥ����������,�_�h�q��D�ɱN�Q�R��!!", 48 + vbYesNo, "�����q�q���ѥ��������q") = vbYes Then
                    rsOrder.Delete
                    '�R����Y�D���٦��O��,�h�i��O����refresh
                    If rsOrder.RecordCount <> 0 Then
                        Call cmdMove_Click(2)
                    Else
                       '�Y�R����q�f�欰�Ū�,�Y�^��D�e��
                        MsgBox "�q�f�椤�w�L�O��!!", , "�����q�q���ѥ��������q"
                        Call cmdEdit_Click(5)
                    End If
                Else
                    '�_�h���X�T�����
                    Exit Sub
                End If
            Else
                Call dgOrDetail_refresh
            End If
            
            lblOrDetail(3).Visible = True
            txtOrDetail(3).Visible = True
            '�^�Ъ����J�ɪ����A
            subcontrol_status False
            cmdSell.Visible = True
    
            Edit_flag = 0
                    
        Case 2    '���ӷs�W
            Call cboProdt_AddItem
            subcontrol_status True
            cmdSell.Visible = False
            Call Add_SubNo(rsOrder_Detail, 1, txtOrder(0))
            For i = 0 To 2
                txtOrDetail(i).Text = ""
            Next i
            cboProdt.Text = "--�п�ܲ��~--"
            '�`����줣�i��
            txtOrDetail(3).Visible = False
            lblOrDetail(3).Visible = False
            
        Case 3    '���ӭק�
            '�ק���Ӫ��X�Э�
            Edit_flag = 1
            
            '1.�O���ثe�O������m
            bm_detail = rsOrder_Detail.Bookmark
            '2.�O���ثe�O�����q�f�ƶq�P�Ͳ���P����
            intpre_NumCk(0) = rsOrder_Detail.Fields("�q�f�ƶq")
            intpre_NumCk(1) = rsOrder_Detail.Fields("�Ͳ��־P")
            
            '3.��cboProdt�e�{��Ӫ����~�W��
            '(1)�O�����cboProdt����
            Prodt_Name = rsOrder_Detail.Fields(3)
            '(2)����cboProdt������
            Call cboProdt_AddItem
            '(3)�e�{���cboProdt�W����
            cboProdt.Text = Prodt_Name
            
            subcontrol_status True
            cmdSell.Visible = False
            
        Case 4    '���ӧR��
            '���@���H�W�����Ӹ�Ʈ�,�h�����R�������O���Y�i
            If MsgBox("�T�w�n�R�������O��??", 32 + vbYesNo, "�����q�q���ѥ��������q") = vbYes Then
                rsOrder_Detail.Delete
                If rsOrder_Detail.RecordCount <> 0 Then
                    rsOrder_Detail.MoveNext
                Else
                    MsgBox "�ѩ󦹵��q��w�L������Ӹ��,�G�D�ɸ�ƱN�Q�R��!", , "�����q�q���ѥ��������q"
                    rsOrder.Delete
                    '�R����Y�D���٦��O��,�h�i��O����refresh
                    If rsOrder.RecordCount <> 0 Then
                        Call cmdMove_Click(2)
                    Else
                        Call cmdEdit_Click(5)
                        '�Y�R����q�f�欰�Ū�,�N�A�i�J�D�ɷs�W�����A
                        MsgBox "�q�f�椤�w�L�O��!!", , "�����w�w���q"
                    End If
                End If
            End If
    End Select
End Sub

'��DataGrid�����O����ܩ�txtOrDetail�W
Private Sub dgOrDetail_Click()
    If rsOrder_Detail.RecordCount <> 0 Then
        For i = 1 To 3
            txtOrDetail(i).Text = rsOrder_Detail.Fields(i + 3)
        Next i
        txtOrDetail(0).Text = rsOrder_Detail.Fields(2)
        txtProdt.Text = rsOrder_Detail.Fields("���~�W��")
    End If
End Sub

'��ƪ��s��
Private Sub txtOrder_Change(Index As Integer)
    Select Case Index
        Case 0      '�q��s��
            Call dgOrDetail_refresh
            
        Case 1      '�g��H�s��
            Set rsStaff = New ADODB.Recordset
            sql_rsStaff = "select * from ���u��ƪ�"
            rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
            
            rsStaff.MoveFirst
            rsStaff.Find "���u�s��='" & txtOrder(1).Text & "'", , adSearchForward, 1
            If rsStaff.EOF = False Then
                txtStaff.Text = rsStaff.Fields("�m�W")
            End If

        Case 2      '�Ȥ�s��
            Set rsCust = New ADODB.Recordset
            sql_rsCust = "select * from �Ȥ���"
            rsCust.Open sql_rsCust, cn, adOpenDynamic, adLockOptimistic
            
            rsCust.Find "�Ȥ�s��='" & txtOrder(2).Text & "'", , adSearchForward, 1
            If rsCust.EOF = False Then
                txtCustomer.Text = rsCust.Fields("���q�W��")
            End If
    End Select
End Sub

'�b�D�ɪ��g��H���@��F�T�ӷP������,�@�Ӭ��x�s��txtOrder(1),
'�@�Ӭ���ܥ�txtStaff,�@�Ӭ��s�׸�Ʈɥ�dcboStaff ;�b�Ȥ����
'�]��F�T�ӷP������ , �x�s�Ϊ�txtOrder(2), ��ܥΪ�txtCustomer
'�P�s�ץΪ�dcboCust

'�D�ɸ�ƪ��s��
Private Sub Form_Load()
    '1.�}�ҭq��D��
    Set rsOrder = New ADODB.Recordset
    sql_rsOrder = "select * from �q��D��"
    rsOrder.Open sql_rsOrder, cn, adOpenDynamic, adLockOptimistic
    
    '2.�N��ƫ��w���D�ɤ����P������
    For Each oText In Me.txtOrder
        Set oText.DataSource = rsOrder
    Next
    
    '3.�s�@DataCombo
    '�s�@dcboStaff(�s�׸�ƮɥΤ�)
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from ���u��ƪ�"
    rsStaff.Open sql_rsStaff, cn, adOpenKeyset, adLockReadOnly
    Set dcboStaff.DataSource = rsOrder
    dcboStaff.DataField = "�g��H�s��"
    Set dcboStaff.RowSource = rsStaff
    dcboStaff.ListField = "�m�W"
    dcboStaff.BoundColumn = "���u�s��"
    
    '�s�@dcboCust(�s�׸�ƮɥΤ�)
    Set rsCust = New ADODB.Recordset
    sql_rsCust = "select * from �Ȥ���"
    rsCust.Open sql_rsCust, cn, adOpenKeyset, adLockReadOnly
    Set dcboCust.DataSource = rsOrder
    dcboCust.DataField = "�Ȥ�s��"
    Set dcboCust.RowSource = rsCust
    dcboCust.ListField = "���q�W��"
    dcboCust.BoundColumn = "�Ȥ�s��"
    
    '4.�ˬd�q�f�椤�O�_�����
    '�Y�q�f�欰�Ū��h�����i�J�D�ɪ��s�W���A
    If rsOrder.RecordCount = 0 Then
        Call cmdEdit_Click(2)
    Else
        '�����J��,��������A�]�w
        control_status False
        subcontrol_status False
        Call cmdMove_Click(0)
    End If
    
End Sub

'���Ӹ�ƪ��s��
Public Sub dgOrDetail_refresh()
    '1.�}�ҭq�����
    Set rsOrder_Detail = New ADODB.Recordset
    sql_rsOrder_Detail = "select * from �q����� where �q��s��='" & txtOrder(0).Text & "'"
    rsOrder_Detail.Open sql_rsOrder_Detail, cn, adOpenDynamic, adLockOptimistic
    
    '�����q�檺���ӵ���
    intAccount(0) = rsOrder_Detail.RecordCount
    
    '2.�N��ƫ��w��DataGrid,�ñN�־P������ð_��
    Set dgOrDetail.DataSource = rsOrder_Detail
    For i = 0 To 2
        dgOrDetail.Columns.Item(i).Visible = False
    Next i
    For i = 7 To 8
        dgOrDetail.Columns.Item(i).Visible = False
    Next i
    
    dgOrDetail.Columns.Item(3).Width = 2000
    For i = 4 To 6
        dgOrDetail.Columns.Item(i).Alignment = dbgRight
    Next i
    
    '3.���Ӹ�Ƹ��J��,��������A�]�w
    subcontrol_status False
        
    '4.��DataGrid���O����ܩ��r�����
    Call dgOrDetail_Click
End Sub

'�D�ɱ�������A�]�w
Private Sub control_status(control_flag As Boolean)
    '�T�w���ܪ����A
    For i = 0 To 4
        cmdEdit(i).Enabled = True
    Next i
    txtOrder(0).Enabled = False
    txtOrder(1).Visible = False
    txtOrder(2).Visible = False
    txtStaff.Enabled = False
    txtCustomer.Enabled = False
    
    '�i���������A
    For i = 1 To 3
        txtOrder(i).Enabled = control_flag
    Next
    dcboStaff.Visible = control_flag
    dcboCust.Visible = control_flag
    txtStaff.Visible = Not control_flag
    txtCustomer.Visible = Not control_flag
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
    
    '�D�ɷs�W��,fraOrDetail�|����;�D�ɭק��,fraOrDetail�|���,�����i�ާ@
    fraOrDetail.Enabled = Not control_flag
    fraOrDetail.Visible = Not control_flag
    
End Sub

'���ӱ�������A�]�w
Private Sub subcontrol_status(subcontrol_flag As Boolean)
    '�T�w���ܪ����A
    For i = 0 To 4
        cmdSubEdit(i).Enabled = True
    Next i
    txtOrDetail(0).Visible = False
    txtOrDetail(3).Enabled = False
    txtProdt.Enabled = False
    
    '�i���������A
    For i = 1 To 2
        txtOrDetail(i).Enabled = subcontrol_flag
    Next
    cboProdt.Visible = subcontrol_flag
    txtProdt.Visible = Not subcontrol_flag
    dgOrDetail.Enabled = Not subcontrol_flag
    '�x�s,�����s
    For i = 0 To 1
        cmdSubEdit(i).Visible = subcontrol_flag
    Next i
    '�s�W,�ק�,�R���s
    For i = 2 To 4
        cmdSubEdit(i).Visible = Not subcontrol_flag
    Next i
    
    '����ӽs�׮�,�D�ɪ�������A
    fraOrder.Enabled = Not subcontrol_flag
    '�^�D�e���s
    cmdEdit(5).Enabled = Not subcontrol_flag
    
End Sub

'�s�@cboProdt
Private Sub cboProdt_AddItem()
    '1.�}�Ҳ��~���
    Set rsProdt = New ADODB.Recordset
    sql_rsProdt = "select * from ���~���"
    rsProdt.Open sql_rsProdt, cn, adOpenKeyset, adLockReadOnly
    
    '2.���M��cboProdt������
    cboProdt.Clear
    '3.�A�N���~�W�٥[�JcboProdt��
    Do Until rsProdt.EOF
        '�q����Ӥ��w�������~,�h���A�X�{
        If rsOrder_Detail.RecordCount <> 0 Then
            rsOrder_Detail.MoveFirst
        End If
        rsOrder_Detail.Find "���~�s��='" & rsProdt.Fields(0) & "'", , 1
        If rsOrder_Detail.EOF Then
            cboProdt.AddItem rsProdt.Fields("���~�W��")
        End If
        rsProdt.MoveNext
    Loop
End Sub

Private Sub cboProdt_Click()
    rsProdt.MoveFirst
    rsProdt.Find "���~�W��='" & cboProdt.Text & "'", , adSearchForward, 1
    If rsProdt.EOF = False Then
        '�N���I��cboProdt�����ȼg�^txtOrDetail(0)��
        txtOrDetail(0).Text = rsProdt.Fields("���~�s��")
        '�����쪺�w�]�Ȭ���ĳ���
        txtOrDetail(1).Text = rsProdt.Fields("��ĳ���")
    End If
End Sub

'��檺�i�ާ@�P�_
Public Sub cmdSell_status()
    cmdSell.Enabled = False
    '�q�榳��Ʈ�
    If rsOrder.EOF = False And rsOrder.BOF = False Then
        '���P�f�e�q�檺�s�׶s�ҥi�ϥ�,���;P�f��s�h�n�����p�ӨM�w�i�_�Q�ϥ�
        If rsOrder.Fields("�P�f�־P") = 0 Then
            rsOrder_Detail.Filter = "�q��s��='" & txtOrder(0).Text & "' and �Ͳ��־P=2"
            intAccount(1) = rsOrder_Detail.RecordCount
            
            '���i�q�檺���~���w�Ͳ�������,�~�i�i�J�P�f���q
            If intAccount(0) = intAccount(1) Then
                cmdSell.Enabled = True
            Else
                cmdSell.Enabled = False
            End If
            '���P�f���e,�q�����٥i�ק�
            For i = 0 To 4
                cmdEdit(i).Enabled = True
                cmdSubEdit(i).Enabled = True
            Next i
            
            rsOrder_Detail.Filter = adFilterNone
            Call dgOrDetail_refresh
        Else
            '�P�f�沣�ͫ�,�q�檺�s�׶s�P���;P�f��s�Ҥ��i�ϥ�,
            cmdSell.Enabled = False
            For i = 0 To 4
                cmdEdit(i).Enabled = False
                cmdSubEdit(i).Enabled = False
            Next i
            '�u���D�ɪ��s�W�s�i�ϥ�
            cmdEdit(2).Enabled = True
        End If
    End If
End Sub

