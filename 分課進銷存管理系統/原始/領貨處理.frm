VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   Caption         =   "領貨處理"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10080
      Top             =   360
   End
   Begin VB.Frame Frame2 
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
      TabIndex        =   18
      Top             =   600
      Width           =   9135
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
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
         Left            =   4680
         TabIndex        =   22
         Top             =   480
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
         Left            =   840
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "領貨單詳細資料"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   480
      TabIndex        =   0
      Top             =   2040
      Width           =   11055
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   5400
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5400
         TabIndex        =   8
         Text            =   "Text6"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "儲存更新"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4200
         Picture         =   "領貨處理.frx":0000
         TabIndex        =   7
         Top             =   3120
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "領貨處理.frx":0974
         Left            =   5400
         List            =   "領貨處理.frx":098A
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "新增領單"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5280
         TabIndex        =   5
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "查詢領單"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6360
         TabIndex        =   4
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "修改領單"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7440
         TabIndex        =   3
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "刪除領單"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8640
         TabIndex        =   2
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "返回"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9720
         TabIndex        =   1
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "領貨單編號"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "領貨單日期"
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
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "零件編號"
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
         Left            =   3840
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "尺寸"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "領貨數量"
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
         Left            =   3960
         TabIndex        =   12
         Top             =   1800
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   480
      TabIndex        =   17
      Top             =   6960
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   2778
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "領貨日期"
         Caption         =   "領貨日期"
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
         DataField       =   "領單編號"
         Caption         =   "領單編號"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "數量"
         Caption         =   "數量"
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
         DataField       =   "庫存量"
         Caption         =   "庫存量"
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
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "標楷體"
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
      TabIndex        =   24
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "標楷體"
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
      TabIndex        =   23
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form8"
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
Sub Lock_Data()   '鎖住文字方塊
 
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    
   Combo1.Enabled = False

  
End Sub

Sub UnLock_Data()  '解開文字方塊
 
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
 
  Combo1.Enabled = True

 
End Sub
Private Sub Command1_Click()


  Command1.Enabled = False
  Command3.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  
  Command2.Enabled = True
    
  Call Lock_Data

End Sub
Private Sub Command2_Click()   '新增
  
    Call UnLock_Data
    
   
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
   

   Combo1.Clear

   
 Command2.Enabled = False
 Command4.Enabled = False
 Command1.Enabled = True
 
End Sub
Private Sub Timer1_Timer()
  
  Label11 = Format(Date$, "yyyy/mm/dd")
  Label12 = Format(Time$, "ampm h:mm:ss")
 
End Sub

