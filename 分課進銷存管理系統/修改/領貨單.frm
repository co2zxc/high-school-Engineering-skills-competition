VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "��f��"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  '�t�ιw�]��
   WindowState     =   2  '�̤j��
   Begin VB.Frame Frame1 
      Caption         =   "���u���"
      Height          =   1095
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   9135
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2640
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   6600
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   360
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
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1695
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
         Height          =   495
         Left            =   4920
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��f��ԲӸ��"
      Height          =   1455
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   9135
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   6720
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "��f��s��"
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
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "��f����"
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
         Left            =   4800
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�C�L"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   1
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��^"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6720
      TabIndex        =   0
      Top             =   5640
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1095
      Left            =   2760
      TabIndex        =   2
      Top             =   3600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1931
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
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
      ColumnCount     =   3
      BeginProperty Column00 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
 Unload Me
End Sub
