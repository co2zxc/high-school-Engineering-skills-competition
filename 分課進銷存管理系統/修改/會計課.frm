VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "�|�p��"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   WindowState     =   2  '�̤j��
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9960
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Text            =   "Text7"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Text            =   "Text8"
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Text            =   "Text9"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Text            =   "Text10"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��^"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      TabIndex        =   2
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�f�֭q��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   1
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�d�߭q��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   0
      Top             =   6840
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   840
      TabIndex        =   9
      Top             =   5160
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�з���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "�Ҩ�s��s��"
         Caption         =   "�Ҩ�s��s��"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "�q�f�ƶq"
         Caption         =   "�q�f�ƶq"
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
         DataField       =   "���B"
         Caption         =   "���B"
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
         EndProperty
         BeginProperty Column04 
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
      Left            =   10800
      TabIndex        =   23
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
      Left            =   12360
      TabIndex        =   22
      Top             =   240
      Width           =   2535
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
      Left            =   720
      TabIndex        =   21
      Top             =   600
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
      Left            =   4920
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "�q��s��"
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
      Left            =   720
      TabIndex        =   19
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "�q����"
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
      Left            =   720
      TabIndex        =   18
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label6 
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
      Left            =   4920
      TabIndex        =   17
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label7 
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
      Left            =   4920
      TabIndex        =   16
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label8 
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
      Left            =   4920
      TabIndex        =   15
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label9 
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
      Left            =   4920
      TabIndex        =   14
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label10 
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
      Left            =   4920
      TabIndex        =   13
      Top             =   3960
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
 Unload Me
End Sub

Private Sub Timer1_Timer()
  
  Label11 = Format(Date$, "yyyy/mm/dd")
  Label12 = Format(Time$, "ampm h:mm:ss")
 
End Sub
