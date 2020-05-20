VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "物料課"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.Timer Timer3 
      Left            =   10920
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Left            =   10200
      Top             =   960
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   12480
      Top             =   7800
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"物料課.frx":0000
      OLEDBString     =   $"物料課.frx":0094
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "物料課"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "物料課.frx":0128
      Height          =   2535
      Left            =   600
      TabIndex        =   38
      Top             =   7440
      Width           =   11415
      _ExtentX        =   20135
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10200
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "訂單詳細資料"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   12375
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   6
         Left            =   8280
         TabIndex        =   39
         Top             =   3960
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   5
         Left            =   2400
         TabIndex        =   37
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   4
         Left            =   2400
         TabIndex        =   36
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   3
         Left            =   1200
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   34
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   5880
         TabIndex        =   27
         Top             =   720
         Width           =   1335
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
         Left            =   10080
         TabIndex        =   20
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "刪除訂單"
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
         Left            =   9000
         TabIndex        =   19
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "修改訂單"
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
         Left            =   7920
         TabIndex        =   18
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "查詢訂單"
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
         Left            =   6720
         TabIndex        =   17
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "新增訂單"
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
         Left            =   5520
         TabIndex        =   16
         Top             =   4560
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "物料課.frx":013D
         Left            =   2400
         List            =   "物料課.frx":0156
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
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
         Left            =   4320
         MaskColor       =   &H80000000&
         Picture         =   "物料課.frx":0176
         Style           =   1  '圖片外觀
         TabIndex        =   14
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Text            =   "Text9"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Text            =   "Text8"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         Text            =   "Text7"
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Text            =   "Text6"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label16 
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
         Left            =   4440
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label15 
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
         Left            =   4440
         TabIndex        =   25
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label14 
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
         Left            =   4440
         TabIndex        =   24
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label13 
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
         Left            =   4440
         TabIndex        =   23
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "總金額"
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
         Left            =   7200
         TabIndex        =   13
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "訂貨數量"
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
         TabIndex        =   11
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "單價"
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
         TabIndex        =   9
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label7 
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
         Left            =   360
         TabIndex        =   7
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label6 
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
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Left            =   4440
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
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
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
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
      Left            =   600
      TabIndex        =   28
      Top             =   240
      Width           =   8655
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5160
         TabIndex        =   32
         Text            =   "Text2"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Text            =   "Text1"
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
         Left            =   3600
         TabIndex        =   31
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
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   1455
      End
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
      Left            =   12600
      TabIndex        =   22
      Top             =   240
      Width           =   2535
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
      Left            =   11040
      TabIndex        =   21
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As Recordset                    '宣告資料表
Dim db As Database                     '宣告資料庫
Dim RS_User As Recordset
Dim i As Integer
Dim s As Recordset
Dim SQL As String
Dim rsson As Recordset

Private Sub Combo1_Change()
If Combo1 <> "" Then
   Text3(3).Text = Combo1()
   End If
End Sub

Private Sub Command1_Click()

Dim iii As Integer

For iii = 0 To s.Fields.Count - 1
    
    s.Fields(iii) = Text3(iii).Text

Next
s.Update

  
  Command1.Enabled = False
  Command3.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  
  Command2.Enabled = True
    
  Call Lock_Data
End Sub

Private Sub Command2_Click()  '新增
s.AddNew
    Call UnLock_Data
    
   

   Combo1.Clear
   Combo2.Clear
  
 Command2.Enabled = False
 Command4.Enabled = False
 Command5.Enabled = False
 Command1.Enabled = True
 
End Sub


Private Sub Command5_Click()  '刪除
    dbName = App.Path & "\db1.mdb"
     Set db = DBEngine.Workspaces(0).OpenDatabase(dbName)
      SQL = SQL & "delete * "
      SQL = SQL & "from 物料課"
      
X = MsgBox("確定要刪除這筆資料嗎?", vbQuestion + vbYesNo, "刪除資料 ")
   If X = vbYes Then
       
      s.Delete
      s.MoveFirst
    
    End If

End Sub

Private Sub Command6_Click()   '返回
   Unload Me
End Sub



Private Sub Form_Load()
  dbName = App.Path & "\db1.mdb" '''''''''''''''''''''''''''''''''''課室
   Set db = DBEngine.Workspaces(0).OpenDatabase(dbName)
    SQL = "select 物料課.*"
    SQL = SQL & "from  物料課 "
    
    
  Set s = db.OpenRecordset(SQL, dbOpenDynaset)
  Call Lock_Data
  
  Command1.Enabled = False
  Command3.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = False
Timer2.Interval = 3000
End Sub

Private Sub Timer1_Timer()
  
  Label11 = Format(Date$, "yyyy/mm/dd")
  Label12 = Format(Time$, "ampm h:mm:ss")
 
End Sub



Sub Lock_Data()   '鎖住文字方塊
 



    
   Combo1.Enabled = False
   Combo2.Enabled = False
  
End Sub

Sub UnLock_Data()  '解開文字方塊

    
  Combo1.Enabled = True
  Combo2.Enabled = True
 
 
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Timer3.Enabled = True
Timer3.Interval = 3000
Adodc1.Recordset.Requery
Text3(6).Text = Val(Text3(4).Text) * Val(Text3(5).Text)
End Sub

Private Sub Timer3_Timer()

Timer2.Enabled = True
Timer3.Enabled = False
Timer2.Interval = 3000
Adodc1.Recordset.Requery

End Sub
