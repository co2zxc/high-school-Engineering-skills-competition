VERSION 5.00
Begin VB.Form frmProdtStore 
   Caption         =   "���~�w�s�έp��"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6525
   Begin VB.TextBox txtProdtStore 
      Alignment       =   1  '�a�k���
      DataField       =   "�{�s�q"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtProdtStore 
      Alignment       =   1  '�a�k���
      DataField       =   "�P�f�q"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtProdtStore 
      Alignment       =   1  '�a�k���
      DataField       =   "�Ͳ��q"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtProdtStore 
      Alignment       =   1  '�a�k���
      DataField       =   "�W���w�s�q"
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
      Left            =   2160
      TabIndex        =   14
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtProdtStore 
      Alignment       =   1  '�a�k���
      DataField       =   "�w���s�q"
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
      Left            =   2160
      TabIndex        =   13
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtProdtStore 
      DataField       =   "���~�W��"
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
      Left            =   2160
      TabIndex        =   12
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   5160
      TabIndex        =   7
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   4800
      TabIndex        =   6
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "�^�D�e��"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "�Ͳ��q"
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
      Left            =   1200
      TabIndex        =   11
      Top             =   1920
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
      Left            =   1560
      TabIndex        =   10
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "�{�s�q"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "�P�f�q"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "�W���w�s�q"
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
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "�w���s�q"
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
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   960
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "���~�W��"
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
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmProdtStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�������i���
Private Sub cmdReturn_Click()
    Unload Me
    frmMDIMain.mnuExit.Enabled = True
End Sub

'��ƪ�����
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsProdtStore)
    lblRecord.Caption = "�w�s��ơG" & intRecord & "/" & intTotal
End Sub

'��ƪ��s��
Private Sub Form_Load()
    '1.�}�Ҳ��Ӯw�s�έp��
    Set rsProdtStore = New ADODB.Recordset
    sql_rsProdtStore = "select * from ���~�w�s�έp��"
    rsProdtStore.Open sql_rsProdtStore, cn, adOpenDynamic, adLockOptimistic
    
    '2.�N��ƫ��w�ܪ�椤���P������W
    For Each oText In Me.txtProdtStore
        Set oText.DataSource = rsProdtStore
    Next
    
    '3.�]�w�����J�ɱ�������A
    For i = 0 To 5
        txtProdtStore(i).Enabled = False
    Next i
    
    '4.����lblRecord�������ƭ�
    Call cmdMove_Click(0)
    
    '5.��檺���e�]�w
    frmProdtStore.Height = 4740
    frmProdtStore.Width = 6645
End Sub


