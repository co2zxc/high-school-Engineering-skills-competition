VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "�i�f�B�z"
   ClientHeight    =   3390
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3390
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   WindowState     =   2  '�̤j��
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10080
      Top             =   360
   End
   Begin VB.Frame Frame1 
      Caption         =   "���u���"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   29
      Top             =   240
      Width           =   9135
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   6360
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "���u�s��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   33
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "���u�m�W"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   32
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�i�f��ԲӸ��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   12375
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   1800
         TabIndex        =   36
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   8400
         TabIndex        =   34
         Text            =   "Text13"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Text            =   "Text4"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5880
         TabIndex        =   15
         Text            =   "Text6"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   5880
         TabIndex        =   14
         Text            =   "Text7"
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   6480
         TabIndex        =   13
         Text            =   "Text8"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Text            =   "Text9"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "�x�s��s"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4440
         MaskColor       =   &H80000000&
         Picture         =   "�i�f�B�z.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   11
         Top             =   4560
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "�i�f�B�z.frx":0974
         Left            =   1200
         List            =   "�i�f�B�z.frx":098A
         TabIndex        =   10
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�s�W�q��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         TabIndex        =   9
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�d�߭q��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6720
         TabIndex        =   8
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�ק�q��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7920
         TabIndex        =   7
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "�R���q��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9000
         TabIndex        =   6
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "��^"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10080
         TabIndex        =   5
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text11"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Text            =   "Text12"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   5880
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "�`���B"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   35
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "�i�f��s��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   28
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "�i�f����"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "�t�ӦW��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   26
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "�s��s��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "�ؤo"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "�i�f�ƶq"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "�t�Ӧa�}"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   21
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "�t�ӹq�ܸ��X"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   20
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "�t�Ӷǯu���X"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   19
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label16 
         Caption         =   "�t�ӽs��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   7680
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�з���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "�q����"
         Caption         =   "�q����"
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
         DataField       =   "�i�f��s��"
         Caption         =   "�i�f��s��"
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
      BeginProperty Column02 
         DataField       =   "�s��s��"
         Caption         =   "�s��s��"
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
      BeginProperty Column03 
         DataField       =   "�ؤo"
         Caption         =   "�ؤo"
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
      BeginProperty Column04 
         DataField       =   "���"
         Caption         =   "���"
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
      BeginProperty Column05 
         DataField       =   "�ƶq"
         Caption         =   "�ƶq"
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
      BeginProperty Column06 
         DataField       =   "�`���B"
         Caption         =   "�`���B"
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
      BeginProperty Column07 
         DataField       =   "�w�s�q"
         Caption         =   "�w�s�q"
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
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            Button          =   -1  'True
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   10920
      TabIndex        =   38
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   12480
      TabIndex        =   37
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command6_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  Call Lock_Data
  
  
  Command1.Enabled = False
  Command3.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
End Sub
Sub Lock_Data()   '����r���
 
    Text3.Enabled = False
    Text4.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Text8.Enabled = False
    Text9.Enabled = False
    Text11.Enabled = False
    Text12.Enabled = False
    Text13.Enabled = False

    
   Combo1.Enabled = False
   Combo2.Enabled = False
   Combo3.Enabled = False
  
End Sub

Sub UnLock_Data()  '�Ѷ}��r���
 
    Text3.Enabled = True
    Text4.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    Text8.Enabled = True
    Text9.Enabled = True
    Text11.Enabled = True
    Text12.Enabled = True
    Text13.Enabled = True
 
  Combo1.Enabled = True
  Combo2.Enabled = True
  Combo3.Enabled = True

 
 
End Sub
Private Sub Command1_Click()


  Command1.Enabled = False
  Command3.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  
  Command2.Enabled = True
    
  Call Lock_Data

End Sub
Private Sub Command2_Click()   '�s�W
  
    Call UnLock_Data
    
   
    Text3.Text = ""
    Text4.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
   
   Text6.Enabled = False
   Text7.Enabled = False
   Text8.Enabled = False
   Text9.Enabled = False
   

   Combo1.Clear
   Combo2.Clear
   Combo3.Clear
   
 Command2.Enabled = False
 Command4.Enabled = False
 Command1.Enabled = True
 
End Sub
Private Sub Timer1_Timer()
  
  Label11 = Format(Date$, "yyyy/mm/dd")
  Label12 = Format(Time$, "ampm h:mm:ss")
 
End Sub
