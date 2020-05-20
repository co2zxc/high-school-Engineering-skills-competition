VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBuy_Detail 
   Caption         =   "商品查詢(下拉式)"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   5895
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "顯示全部"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "回主畫面"
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
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdWhere 
      Caption         =   "搜尋"
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
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
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
      Alignment       =   2  '置中對齊
      Caption         =   "品牌"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "類別"
      BeginProperty Font 
         Name            =   "新細明體"
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
        mySQL = "select 產品編號,產品名稱,安全存量,單位成本,建議售價,廠商編號,類別編號 from 產品資料 where 類別 = '" & dcboWhere.Text & "' and 品牌 = '" & dcboWhere2.Text & "'"
    '註解'的意思在資料庫是代表中文的轉換所以全存量不用設定
    '製作dgBuy_Detail的RecordSet物件
            
    Set rs = New Recordset
    
        rs.Open mySQL, cn, adOpenStatic, adLockOptimistic
    
    Set dgBuy_Detail.DataSource = rs
End Sub
 
Private Sub Command2_Click()
    Set cn = New Connection
    cn.CursorLocation = adUseClient
    cn.Open "dsn=dbRose"
    Set rs = New Recordset
        rs.Open "select 產品編號,產品名稱,安全存量,單位成本,建議售價 " _
            & "from 產品資料", cn, adOpenStatic, adLockOptimistic
     Set dgBuy_Detail.DataSource = rs
End Sub

Private Sub Form_Load()
    Set cn = New Connection
    cn.CursorLocation = adUseClient
    cn.Open "dsn=dbRose"
    Set rs = New Recordset
        rs.Open "select 產品編號,產品名稱,安全存量,單位成本,建議售價 " _
            & "from 產品資料", cn, adOpenStatic, adLockOptimistic
            
    Set dgBuy_Detail.DataSource = rs

    '製作dcboWhere的RecordSet物件
    Set rs_dcbo = New Recordset
        rs_dcbo.Open "select * from 類別資料", cn, adOpenStatic, adLockOptimistic
        
    Set dcboWhere.DataSource = rs_dcbo
        dcboWhere.DataField = "類別"
    Set dcboWhere.RowSource = rs_dcbo
        dcboWhere.ListField = "類別"


    Set rs_dcbo2 = New Recordset
        rs_dcbo2.Open "select * from 廠商資料", cn, adOpenStatic, adLockOptimistic
    
        
    Set dcboWhere2.DataSource = rs_dcbo2
        dcboWhere2.DataField = "廠商"
    Set dcboWhere2.RowSource = rs_dcbo2
        dcboWhere2.ListField = "廠商"
End Sub

Private Sub Command1_Click()
Unload Me
frmMDIMain.Show
End Sub
