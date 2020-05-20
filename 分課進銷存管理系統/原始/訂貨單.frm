VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   Caption         =   "訂貨單"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form10"
   ScaleHeight     =   5280
   ScaleWidth      =   8310
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Text            =   "Text10"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   7200
      TabIndex        =   19
      Text            =   "Text9"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   6720
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   2160
      Width           =   1695
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
      Height          =   975
      Left            =   6720
      TabIndex        =   8
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "列印"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   7
      Top             =   7320
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   840
      TabIndex        =   6
      Top             =   5160
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "模具零件編號"
         Caption         =   "模具零件編號"
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1695
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
      Left            =   4920
      TabIndex        =   17
      Top             =   3960
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
      Left            =   4920
      TabIndex        =   15
      Top             =   3360
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
      Left            =   4920
      TabIndex        =   14
      Top             =   2760
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
      Left            =   4920
      TabIndex        =   13
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "廠商編號"
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
      Left            =   4920
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "訂單日期"
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
      Left            =   720
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "訂單編號"
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
      Left            =   720
      TabIndex        =   4
      Top             =   1440
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
      Left            =   4920
      TabIndex        =   2
      Top             =   600
      Width           =   1455
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
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Form2.Show
End Sub

Private Sub Command2_Click()
  Form2.Show
End Sub

