VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0FFFF&
   Caption         =   "客戶篩選"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   9495
   ScaleWidth      =   14130
   WindowState     =   2  '最大化
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "查詢條件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   9495
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "取消篩選"
         Height          =   975
         Left            =   8160
         Picture         =   "Form4.frx":16A0
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   4560
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "確定篩選"
         Height          =   975
         Left            =   6840
         Picture         =   "Form4.frx":303F
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   2280
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5250
      Left            =   960
      TabIndex        =   1
      Top             =   3240
      Width           =   12615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form4.frx":49DE
         Height          =   4575
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   8070
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7920
      Top             =   2400
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=0104.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=0104.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "客戶資料"
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
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "客戶查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset


Private Sub Command1_Click()
 Text9 = ""
 Combo1 = ""
  Combo2 = ""
  Text11 = ""
  Text12 = ""
End Sub

Private Sub Command12_Click()
aa
On errow GoTo 11
Select Case Combo1.Text
Case "客戶編號"
Text11.Text = "Where " & Combo1.Text & Combo2.Text & Text9
Case Else
Text11.Text = "Where " & Combo1.Text & Combo2.Text & "'" & Text9 & "%'"
End Select
Text12.Text = "select * from 客戶資料 " & Text11
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\0104.mdb;Persist Security Info=False"
  Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = Text12
Adodc1.Refresh
Set rs = Adodc1.Recordset
Set DataGrid1.DataSource = Adodc1
 If rs.EOF Then
    MsgBox ("無符合查詢條件記錄！")
    End If
Exit Sub
11
MsgBox "打錯了啦！"
aa

End Sub


Private Sub DataGrid1_Click()
  aa
End Sub

Private Sub Form_Load()
  MDIForm1.StatusBar1.Panels(4) = "客戶查詢系統"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\0104.mdb;Persist Security Info=False"
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from 客戶資料"
Adodc1.Refresh
Set rs = Adodc1.Recordset
Set DataGrid1.DataSource = Adodc1

For I = 0 To rs.Fields.Count - 1
Combo1.AddItem rs.Fields(I).Name
Next I

'For I = 0 To 6
'Set Combo1.DataSource = DataEnvironment1
'Combo1.DataMember = "COMMAND1"
'For A = 0 To DataEnvironment1.rsCommand1.Fields.Count - 1
'Combo1.AddItem DataEnvironment1.rs會員編號查詢.Fields(A).Name
'Next A
'Next I

Combo2.AddItem " > "
Combo2.AddItem " < "
Combo2.AddItem " = "
Combo2.AddItem " >= "
Combo2.AddItem " <= "
Combo2.AddItem " <> "
Combo2.AddItem " like "

End Sub
Sub aa()
MDIForm1.StatusBar1.Panels(3).Text = "第" & rs.AbsolutePosition & "筆 / 共" & rs.RecordCount & "筆"
End Sub

