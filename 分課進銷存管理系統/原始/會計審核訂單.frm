VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form11 
   Caption         =   "審核訂單"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.Frame Frame2 
      Caption         =   "訂貨單詳細資料"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   480
      TabIndex        =   8
      Top             =   1920
      Width           =   12255
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Text            =   "Text4"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   6600
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   7440
         TabIndex        =   12
         Text            =   "Text10"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   7440
         TabIndex        =   11
         Text            =   "Text9"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   6600
         TabIndex        =   10
         Text            =   "Text8"
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   6600
         TabIndex        =   9
         Text            =   "Text7"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "訂貨單編號"
         BeginProperty Font 
            Name            =   "標楷體"
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
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "訂貨單日期"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "廠商"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "廠商傳真號碼"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   19
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "廠商電話號碼"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   18
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "廠商地址"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "廠商名稱"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "員工資料"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   9135
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   6360
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "員工編號"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "員工姓名"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   2
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "審核"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      TabIndex        =   0
      Top             =   8520
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   1200
      TabIndex        =   1
      Top             =   6120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "零件編號"
         Caption         =   "零件編號"
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
         DataField       =   "尺寸"
         Caption         =   "尺寸"
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
         DataField       =   "單價"
         Caption         =   "單價"
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
         DataField       =   "訂貨數量"
         Caption         =   "訂貨數量"
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
         DataField       =   "金額"
         Caption         =   "金額"
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
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
