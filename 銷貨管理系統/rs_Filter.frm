VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuy1_Detail 
   Caption         =   "產品查詢(關鍵字)"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   9915
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "回主畫面"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "搜尋"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtFilter 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "輸入產品關鍵字:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   1755
   End
End
Attribute VB_Name = "frmBuy1_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim cn As Connection
Dim rs As Recordset
Dim WithEvents dbfile As ADODB.Connection
Attribute dbfile.VB_VarHelpID = -1

Dim WithEvents dbrset As ADODB.Recordset
Attribute dbrset.VB_VarHelpID = -1

Dim criteria As String

Private Sub cmdFilter_Click()
   '設定條件式
    
    'criteria = "產品名稱 = like '" & txtFilter.Text & "'+"%""
criteria = "產品名稱 like '*" & txtFilter.Text & "*'"
    
    '檢查文字方塊中是否有輸入資料
    If txtFilter.Text = "" Then
        MsgBox ("請輸入欲篩選的產品編號!!")
        txtFilter.SetFocus
    Else
        '將條件式帶入
        rs.Filter = criteria
        If rs.RecordCount = 0 Then
           MsgBox "抱歉!沒有你所輸入產品編號:" & txtFilter.Text
           
           '回復到表單載入時的狀態
           txtFilter.Text = ""
           txtFilter.SetFocus
           Call Form_Load
        End If
    End If
    ' End If
End Sub

Private Sub Command1_Click()
Unload Me
frmMDIMain.Show
End Sub

Private Sub Form_Load()
    Set cn = New Connection
        cn.CursorLocation = adUseClient
        cn.Open "dbRose"
        
    '表單載入時 , 同時也將資料載入
    Set rs = New Recordset
        rs.Open "select 產品編號,產品名稱,安全存量,單位成本,建議售價 " _
            & "from 產品資料", cn, adOpenStatic, adLockOptimistic
            
    Set dgBuy_Detail.DataSource = rs
End Sub

