VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBuy_Detail 
   Caption         =   "�ӫ~�d��(�U�Ԧ�)"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   5895
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "��ܥ���"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�^�D�e��"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3960
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo dcboWhere2 
      Height          =   330
      Left            =   1800
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcboWhere 
      Height          =   360
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.CommandButton cmdWhere 
      Caption         =   "�j�M"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dgBuy_Detail 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
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
   Begin VB.Label Label2 
      Alignment       =   2  '�m�����
      Caption         =   "�~�P"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      Caption         =   "���O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   1200
   End
End
Attribute VB_Name = "frmBuy_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As Connection
Dim rs As Recordset
Dim rs_dcbo2 As Recordset
Dim rs_dcbo As Recordset


Private Sub cmdWhere_Click()
    Dim mySQL As String
        mySQL = "select ���~�s��,���~�W��,�w���s�q,��즨��,��ĳ���,�t�ӽs��,���O�s�� from ���~��� where ���O = '" & dcboWhere.Text & "' and �~�P = '" & dcboWhere2.Text & "'"
    '����'���N��b��Ʈw�O�N���媺�ഫ�ҥH���s�q���γ]�w
    '�s�@dgBuy_Detail��RecordSet����
            
    Set rs = New Recordset
    
        rs.Open mySQL, cn, adOpenStatic, adLockOptimistic
    
    Set dgBuy_Detail.DataSource = rs
End Sub
 
Private Sub Command2_Click()
    Set cn = New Connection
    cn.CursorLocation = adUseClient
    cn.Open "dsn=dbRose"
    Set rs = New Recordset
        rs.Open "select ���~�s��,���~�W��,�w���s�q,��즨��,��ĳ��� " _
            & "from ���~���", cn, adOpenStatic, adLockOptimistic
     Set dgBuy_Detail.DataSource = rs
End Sub

Private Sub Form_Load()
    Set cn = New Connection
    cn.CursorLocation = adUseClient
    cn.Open "dsn=dbRose"
    Set rs = New Recordset
        rs.Open "select ���~�s��,���~�W��,�w���s�q,��즨��,��ĳ��� " _
            & "from ���~���", cn, adOpenStatic, adLockOptimistic
            
    Set dgBuy_Detail.DataSource = rs

    '�s�@dcboWhere��RecordSet����
    Set rs_dcbo = New Recordset
        rs_dcbo.Open "select * from ���O���", cn, adOpenStatic, adLockOptimistic
        
    Set dcboWhere.DataSource = rs_dcbo
        dcboWhere.DataField = "���O"
    Set dcboWhere.RowSource = rs_dcbo
        dcboWhere.ListField = "���O"


    Set rs_dcbo2 = New Recordset
        rs_dcbo2.Open "select * from �t�Ӹ��", cn, adOpenStatic, adLockOptimistic
    
        
    Set dcboWhere2.DataSource = rs_dcbo2
        dcboWhere2.DataField = "�t��"
    Set dcboWhere2.RowSource = rs_dcbo2
        dcboWhere2.ListField = "�t��"
End Sub

Private Sub Command1_Click()
Unload Me
frmMDIMain.Show
End Sub
