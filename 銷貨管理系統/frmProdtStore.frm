VERSION 5.00
Begin VB.Form frmProdtStore 
   Caption         =   "產品庫存統計表"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6525
   Begin VB.TextBox txtProdtStore 
      Alignment       =   1  '靠右對齊
      DataField       =   "現存量"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2160
      TabIndex        =   17
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtProdtStore 
      Alignment       =   1  '靠右對齊
      DataField       =   "銷貨量"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   16
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtProdtStore 
      Alignment       =   1  '靠右對齊
      DataField       =   "生產量"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   15
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtProdtStore 
      Alignment       =   1  '靠右對齊
      DataField       =   "上期庫存量"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   14
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtProdtStore 
      Alignment       =   1  '靠右對齊
      DataField       =   "安全存量"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   13
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtProdtStore 
      DataField       =   "產品名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   12
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   5160
      TabIndex        =   7
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   4800
      TabIndex        =   6
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "回主畫面"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "生產量"
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
      Index           =   3
      Left            =   1200
      TabIndex        =   11
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label lblRecord 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "現存量"
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
      Index           =   5
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "銷貨量"
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
      Index           =   4
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "上期庫存量"
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
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "安全存量"
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
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   960
   End
   Begin VB.Label lblProdtStore 
      AutoSize        =   -1  'True
      Caption         =   "產品名稱"
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
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmProdtStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'結束此張表單
Private Sub cmdReturn_Click()
    Unload Me
    frmMDIMain.mnuExit.Enabled = True
End Sub

'資料的移動
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsProdtStore)
    lblRecord.Caption = "庫存資料：" & intRecord & "/" & intTotal
End Sub

'資料的連結
Private Sub Form_Load()
    '1.開啟產敏庫存統計表
    Set rsProdtStore = New ADODB.Recordset
    sql_rsProdtStore = "select * from 產品庫存統計表"
    rsProdtStore.Open sql_rsProdtStore, cn, adOpenDynamic, adLockOptimistic
    
    '2.將資料指定至表單中的感知元件上
    For Each oText In Me.txtProdtStore
        Set oText.DataSource = rsProdtStore
    Next
    
    '3.設定表單載入時控制項的狀態
    For i = 0 To 5
        txtProdtStore(i).Enabled = False
    Next i
    
    '4.重整lblRecord中的筆數值
    Call cmdMove_Click(0)
    
    '5.表單的長寬設定
    frmProdtStore.Height = 4740
    frmProdtStore.Width = 6645
End Sub


