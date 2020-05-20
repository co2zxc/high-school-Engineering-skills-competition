VERSION 5.00
Begin VB.Form frmChoose 
   Caption         =   "選擇表單"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   4380
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
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
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cboStaff 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ComboBox cboCust 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblChoose 
      AutoSize        =   -1  'True
      Caption         =   "經手人"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label lblChoose 
      AutoSize        =   -1  'True
      Caption         =   "客戶名稱"
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
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   960
   End
   Begin VB.Label lblChoose 
      AutoSize        =   -1  'True
      Caption         =   "請選擇客戶與經手人"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2700
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'檢查cboCust中的客戶是否有重複的變數
Dim pre_Cust_No As String

'當新增一筆主檔後,不想繼續則可按下取消鈕, 以回到表單載入時的狀態
Private Sub cmdCancel_Click()
    rsSell.CancelUpdate
    Me.Hide
    frmSell.Show
    '按下取消鈕的狀態為1
    Main_cancel = 1
End Sub

'在按下確定鈕後,選擇表單中的客戶於經手人資料將會寫回單據中,並完成一筆主檔的新增
Private Sub cmdOK_Click()
    If cboStaff.Text <> "請點選" And cboStaff.Text <> "" And _
        cboCust.Text <> "請點選" And cboCust.Text <> "" Then
        
        '將cboStaff的值以編號的形式寫到銷貨單的txtSell(1)中
        rsStaff.MoveFirst
        rsStaff.Find "姓名='" & cboStaff.Text & "'", , adSearchForward, 1
        If rsStaff.EOF = False Then
            frmSell.txtSell(1).Text = rsStaff.Fields("員工編號")
        End If
        
        '將cboCust的值以編號的形式寫到銷貨單的txtSell(2)中
        rsCust.MoveFirst
        rsCust.Find "公司名稱='" & cboCust.Text & "'", , adSearchForward, 1
        If rsCust.EOF = False Then
            frmSell.txtSell(2).Text = rsCust.Fields("客戶編號")
        End If
    Else
        MsgBox ("請選擇經手人或客戶!!")
        Exit Sub
    End If
    rsSell.Update
    Me.Hide
    frmSell.Show
    '按下確定鈕的狀態為0
    Main_cancel = 0
End Sub

'表單載入,製作cboStaff與cboCust
Public Sub Form_Activate()
    
    '製作cboStaff
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from 員工資料表"
    rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
    
    '在每一次呼叫選擇表單時,都必須將cboStaff 中的值移除,然後重新AddItem進來
    cboStaff.Clear
    '讓cboStaff的清單為名稱顯示
    Do Until rsStaff.EOF
        cboStaff.AddItem rsStaff.Fields("姓名")
        rsStaff.MoveNext
    Loop
    
    '製作cboCust
    Set rsCust = New ADODB.Recordset
    sql_rsCust = "select * from 客戶資料"
    rsCust.Open sql_rsCust, cn, adOpenDynamic, adLockOptimistic
    
    Set rsOrder = New ADODB.Recordset
    sql_rsOrder = "select * from 訂單主檔 where 是否核銷=0 order by 客戶編號"
    rsOrder.Open sql_rsOrder, cn, adOpenDynamic, adLockOptimistic
    
    '在每一次呼叫選擇表單時,都必須將cboCust 中的值移除,然後重新AddItem進來
    pre_Cust_No = ""
    cboCust.Clear
    '讓cboCust的清單為名稱顯示,並將重複的客戶去除
    Do Until rsOrder.EOF
        If pre_Cust_No <> rsOrder.Fields("客戶編號") Then
            rsCust.MoveFirst
            rsCust.Find "客戶編號='" & rsOrder.Fields("客戶編號") & "'"
            cboCust.AddItem rsCust.Fields("公司名稱")
        End If
        pre_Cust_No = rsOrder.Fields("客戶編號")
        rsOrder.MoveNext
    Loop
    'combo Box的初始值
    cboStaff.Text = "請點選"
    cboCust.Text = "請點選"
End Sub






