VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "�h�f�B�z"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   WindowState     =   2  '�̤j��
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9840
      Top             =   240
   End
   Begin VB.Frame Frame2 
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
      Left            =   480
      TabIndex        =   30
      Top             =   120
      Width           =   8655
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5160
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   480
         Width           =   1335
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
         Left            =   240
         TabIndex        =   34
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
         Left            =   3600
         TabIndex        =   33
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�h�f��ԲӸ��"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   12375
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   1200
         TabIndex        =   36
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Text            =   "Text4"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5880
         TabIndex        =   16
         Text            =   "Text6"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   5880
         TabIndex        =   15
         Text            =   "Text7"
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   6480
         TabIndex        =   14
         Text            =   "Text8"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   6480
         TabIndex        =   13
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
         Left            =   3120
         MaskColor       =   &H80000000&
         Picture         =   "�h�f�B�z.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   12
         Top             =   4560
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "�h�f�B�z.frx":0974
         Left            =   1800
         List            =   "�h�f�B�z.frx":098A
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�s�W   �h�f��"
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
         Left            =   4200
         TabIndex        =   10
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�d��  �h�f��"
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
         Left            =   5760
         TabIndex        =   9
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�ק�   �h�f��"
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
         Left            =   7320
         TabIndex        =   8
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "�R��  �h�f��"
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
         Left            =   8640
         TabIndex        =   7
         Top             =   4560
         Width           =   1215
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
         TabIndex        =   6
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text11"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Text            =   "Text12"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   8280
         TabIndex        =   3
         Text            =   "Text13"
         Top             =   3960
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   5880
         TabIndex        =   2
         Top             =   720
         Width           =   1335
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
         TabIndex        =   35
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "�h�f��s��"
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
         TabIndex        =   29
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "�h�f����"
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "�h�f�ƶq"
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
         Top             =   3840
         Width           =   1455
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
         Left            =   7200
         TabIndex        =   23
         Top             =   3960
         Width           =   1095
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   7800
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4260
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "�h�f����"
         Caption         =   "�h�f����"
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
         DataField       =   "�h�f��s��"
         Caption         =   "�h�f��s��"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Button          =   -1  'True
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
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
      Left            =   10680
      TabIndex        =   38
      Top             =   240
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
      Left            =   12240
      TabIndex        =   37
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


  Command1.Enabled = False
  Command3.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  
  Command2.Enabled = True
    
  Call Lock_Data

End Sub

Private Sub Command6_Click()  '��^
  Unload Me
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

Private Sub Form_Load()
  Call Lock_Data
  
  
  Command1.Enabled = False
  Command3.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
End Sub
Private Sub Timer1_Timer()
  
  Label11 = Format(Date$, "yyyy/mm/dd")
  Label12 = Format(Time$, "ampm h:mm:ss")
 
End Sub
