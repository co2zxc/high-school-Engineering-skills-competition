VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   BackColor       =   &H00C0FFFF&
   Caption         =   "�e�Q�W"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form10"
   MDIChild        =   -1  'True
   Picture         =   "�P��Ʀ�.frx":0000
   ScaleHeight     =   10380
   ScaleWidth      =   15240
   WindowState     =   2  '�̤j��
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�P��e�Q�W"
      Height          =   7455
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   14415
      Begin VB.CommandButton Command9 
         Caption         =   "����"
         Height          =   975
         Left            =   13320
         Picture         =   "�P��Ʀ�.frx":22D3
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Top             =   6240
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "�P��Ʀ�.frx":3125
         Height          =   6735
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   11880
         _Version        =   393216
         BackColor       =   16777215
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
         DataMember      =   "�P��e�Q�W"
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
      Begin VB.Label Label1 
         BackStyle       =   0  '�z��
         Caption         =   "�n�����D�A�|�û����d�b�ڭ̤ߤ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   8040
         TabIndex        =   4
         Top             =   6720
         Width           =   4815
      End
      Begin VB.Image Image2 
         Height          =   3375
         Left            =   6720
         Picture         =   "�P��Ʀ�.frx":3144
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   4620
      End
      Begin VB.Image Image1 
         Height          =   2895
         Left            =   10680
         Picture         =   "�P��Ʀ�.frx":86DC
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   3450
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "�P��Ʀ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  MDIForm1.StatusBar1.Panels(2) = "�����ӫ~�Ʀ�"
  
If DataEnvironment1.rs�P��e�Q�W.State <> adStateClosed Then DataEnvironment1.rs�P��e�Q�W.Close
  DataEnvironment1.�P��e�Q�W
  Set rs = DataEnvironment1.rs�P��e�Q�W
  Set DataGrid1.DataSource = DataEnvironment1
  DataGrid1.DataMember = "�P��e�Q�W"
  
End Sub

Private Sub Command9_Click()
 Unload Me
End Sub



