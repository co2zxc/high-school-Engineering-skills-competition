VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCust 
   Caption         =   "客戶資料表"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   8115
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   6360
      TabIndex        =   25
      Top             =   3600
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   6720
      TabIndex        =   24
      Top             =   3600
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   720
      TabIndex        =   23
      Top             =   3600
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   1080
      TabIndex        =   22
      Top             =   3600
      Width           =   350
   End
   Begin VB.TextBox txtArea 
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
      Left            =   1560
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "修     改"
      Height          =   375
      Index           =   3
      Left            =   6600
      TabIndex        =   19
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "新     增"
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   18
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "刪     除"
      Height          =   375
      Index           =   4
      Left            =   6600
      TabIndex        =   17
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "儲     存"
      Height          =   375
      Index           =   0
      Left            =   6600
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "取     消"
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   15
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "回主畫面"
      Height          =   375
      Index           =   5
      Left            =   6600
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtCust 
      DataField       =   "地區編號"
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
      Index           =   6
      Left            =   1560
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtCust 
      Alignment       =   1  '靠右對齊
      DataField       =   "電話"
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
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   3000
   End
   Begin VB.TextBox txtCust 
      DataField       =   "聯絡人"
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
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   3000
   End
   Begin VB.TextBox txtCust 
      DataField       =   "公司負責人"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   3000
   End
   Begin VB.TextBox txtCust 
      DataField       =   "公司名稱"
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
      Left            =   1560
      TabIndex        =   7
      Top             =   720
      Width           =   3000
   End
   Begin VB.TextBox txtCust 
      Alignment       =   1  '靠右對齊
      DataField       =   "客戶編號"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   3000
   End
   Begin MSDataListLib.DataCombo dcboArea 
      Height          =   390
      Left            =   1560
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtCust 
      DataField       =   "公司地址"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   2640
      Width           =   3255
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
      Left            =   1440
      TabIndex        =   21
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "公司地址"
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
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "電話"
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
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人"
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
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "公司負責人"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "公司名稱"
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
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號"
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
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'資料的編修
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '儲存
            For i = 1 To 5
                If txtCust(i).Text = "" Then
                    MsgBox "抱歉!資料不能空白!", , "金光益電器股份有限公司"
                    txtCust(i).SetFocus
                    Exit Sub
                End If
            Next i
                
            rsArea.MoveFirst
            rsArea.Find "地區名稱='" & dcboArea.Text & "'", , 1
            If rsArea.EOF Then
                MsgBox "請選擇正確的地區名稱!", , "金光益電器股份有限公司"
                dcboArea.SetFocus
                Exit Sub
            Else
                txtCust(6).Text = dcboArea.BoundText
            End If
            
            rsCust.Update
            control_status False
            Call cmdMove_Click(3)

        Case 1      '取消
            rsCust.CancelUpdate
            If rsCust.RecordCount <> 0 Then
                '若是新增狀態下的取消,則要重整筆數的顯示
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                For Each oText In Me.txtCust
                    Set oText.DataSource = rsCust
                Next
                dcboArea.BoundText = txtCust(6).Text
                control_status False
            Else
                Call cmdEdit_Click(5)
            End If
            Add_flag = 0
            
        Case 2      '新增
            Add_flag = 1
            Call Add_No(rsCust, 1, 1)
            control_status True
            txtCust(0).Text = Main_No
            dcboArea.Text = "請點選"
            
        Case 3      '修改
            control_status True
            
        Case 4      '刪除
            If MsgBox("確定要刪除此筆記錄??", 32 + vbYesNo, "金光益電器股份有限公司") = vbYes Then
                rsCust.Delete
                If rsCust.RecordCount <> 0 Then
                    Call cmdMove_Click(2)
                Else
                    Call cmdEdit_Click(5)
                End If
            End If
            
        Case 5      '回主畫面
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            rsCust.Close
            Set rsCust = Nothing
    End Select
End Sub

'資料的移動
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsCust)
    lblRecord.Caption = "客戶資料：" & intRecord & "/" & intTotal
End Sub

'資料的連結
Private Sub Form_Load()
    '1.開啟客戶資料
    Set rsCust = New ADODB.Recordset
    sql_rsCust = "select * from 客戶資料"
    rsCust.Open sql_rsCust, cn, adOpenDynamic, adLockOptimistic
    
    '2.將資料指定給表單中的感知元件
    For Each oText In Me.txtCust
        Set oText.DataSource = rsCust
    Next
    
    '3.製作dcboArea
    Set rsArea = New ADODB.Recordset
    sql_rsArea = "select * from 地區表"
    rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
    
    Set dcboArea.DataSource = rsCust
    dcboArea.DataField = "地區編號"
    Set dcboArea.RowSource = rsArea
    dcboArea.ListField = "地區名稱"
    dcboArea.BoundColumn = "地區編號"
    rsCust.MoveFirst
    
    '4.檢查是否有值
    If rsCust.RecordCount <> 0 Then
        cmdMove_Click (0)
        control_status False
    Else
        '若無資料
        Call cmdEdit_Click(2)
        control_status True
    End If
    
    '5.表單的長寬設定
    frmCust.Height = 4920
    frmCust.Width = 8235
    
End Sub

'控制項的狀態
Private Sub control_status(control_flag As Boolean)
    '固定不變的狀態
    txtCust(0).Enabled = False
    txtArea.Enabled = False
    txtCust(6).Visible = False
    
    '可切換的狀態
    For i = 1 To 5
        txtCust(i).Enabled = control_flag
    Next
    dcboArea.Visible = control_flag
    txtArea.Visible = Not control_flag
    
    '儲存,取消鈕
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next
    '新增,修改,刪除,回主畫面鈕
    For i = 2 To 5
        cmdEdit(i).Visible = Not control_flag
    Next i
    '移動鈕
    For i = 0 To 3
        cmdMove(i).Enabled = Not control_flag
    Next i
    
End Sub

'讓txtArea的值與資料連動
Private Sub txtCust_Change(Index As Integer)
    Select Case Index
        Case 6
            Set rsArea = New ADODB.Recordset
            sql_rsArea = "select * from 地區表"
            rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
            
            rsArea.Find "地區編號='" & txtCust(6).Text & "'"
            If rsArea.EOF = False Then
                txtArea.Text = rsArea.Fields("地區名稱")
            End If
    End Select
End Sub
