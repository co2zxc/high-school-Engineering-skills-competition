VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSell 
   Caption         =   "銷貨單"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11340
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "列印報表"
      Height          =   375
      Left            =   1200
      TabIndex        =   26
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "瀏覽訂單"
      Height          =   375
      Left            =   1200
      TabIndex        =   25
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "回主畫面"
      Height          =   375
      Index           =   3
      Left            =   7800
      TabIndex        =   24
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraSelDetail 
      Caption         =   "銷貨明細"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2520
      TabIndex        =   5
      Top             =   3960
      Width           =   6495
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "儲     存"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   21
         Top             =   2880
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dgSelDetail 
         Height          =   2295
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1.2
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
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
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "取     消"
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   10
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.Frame fraSell 
      Caption         =   "銷貨主檔"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   6495
      Begin VB.TextBox txtSell 
         Alignment       =   1  '靠右對齊
         DataField       =   "出貨日"
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
         Index           =   3
         Left            =   1575
         TabIndex        =   22
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "<<"
         Height          =   355
         Index           =   1
         Left            =   960
         TabIndex        =   19
         Top             =   2520
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "|<"
         Height          =   355
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   2520
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">|"
         Height          =   355
         Index           =   3
         Left            =   4920
         TabIndex        =   17
         Top             =   2520
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">>"
         Height          =   355
         Index           =   2
         Left            =   4560
         TabIndex        =   16
         Top             =   2520
         Width           =   350
      End
      Begin VB.ComboBox cboEmployee 
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
         TabIndex        =   15
         Top             =   960
         Width           =   2500
      End
      Begin VB.TextBox txtCust 
         DataField       =   "客戶名稱"
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
         TabIndex        =   14
         Top             =   1440
         Width           =   2500
      End
      Begin VB.TextBox txtStaff 
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
         TabIndex        =   13
         Top             =   960
         Width           =   2500
      End
      Begin VB.TextBox txtSell 
         DataField       =   "客戶編號"
         Height          =   345
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "修     改"
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "取     消"
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "儲     存"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtSell 
         DataField       =   "經手人編號"
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
         TabIndex        =   4
         Top             =   960
         Width           =   2500
      End
      Begin VB.TextBox txtSell 
         Alignment       =   1  '靠右對齊
         DataField       =   "銷貨單編號"
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
         TabIndex        =   3
         Top             =   480
         Width           =   2500
      End
      Begin VB.Label lblSell 
         AutoSize        =   -1  'True
         Caption         =   "出貨日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   720
         TabIndex        =   23
         Top             =   2070
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
         Left            =   1320
         TabIndex        =   20
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label lblSell 
         AutoSize        =   -1  'True
         Caption         =   "客戶"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   930
         TabIndex        =   11
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label lblSell 
         AutoSize        =   -1  'True
         Caption         =   "經手人"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label lblSell 
         AutoSize        =   -1  'True
         Caption         =   "銷貨單號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   510
         TabIndex        =   1
         Top             =   660
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSelDetail_list As ADODB.Recordset     '銷貨明細的暫存RecordSet
Dim update_flag As Integer                  '明細儲存鈕的旗標(用來決定是否更新庫存與核銷資料)

'主檔的編修
Public Sub cmdEdit_Click(index As Integer)
    Select Case index
        Case 0      '儲存
            rsStaff.MoveFirst
            rsStaff.Find "姓名='" & cboEmployee.Text & "'", , 1
            If rsStaff.EOF Then
                MsgBox "請選擇經手人!", , "金光益電器股份有限公司"
                cboEmployee.SetFocus
                Exit Sub
            Else
                If txtSell(3).Text = "" Then
                    MsgBox "請填寫出貨日!", , "金光益電器股份有限公司"
                    txtSell(3).SetFocus
                    Exit Sub
                Else
                    If IsDate(txtSell(3).Text) = False Then
                        MsgBox "您輸入的日期格式有誤,請再確認一遍!", , "金光益電器股份有限公司"
                        txtSell(3).SetFocus
                        Exit Sub
                    Else
                        '儲存用
                        txtSell(1).Text = rsStaff.Fields("員工編號")
                        
                        rsSell.Fields(0) = txtSell(0).Text
                        rsSell.Fields(1) = rsOrder.Fields(0)
                        rsSell.Fields(2) = txtSell(1).Text
                        
                        Set rsCust = New ADODB.Recordset
                        rsCust.Open "select * from 客戶資料", cn, adOpenDynamic, adLockOptimistic
                        rsCust.MoveFirst
                        rsCust.Find "公司名稱='" & txtCust.Text & "'", , 1
                        rsSell.Fields(3) = rsCust.Fields(0)
                        rsSell.Fields(4) = txtCust.Text
                        rsSell.Fields(5) = txtSell(3).Text
                        rsSell.Update
                        
                        control_status False
                        Call cmdMove_Click(3)
                    End If
                End If
            End If
            
            '當新增一筆銷貨主檔後,即進入銷貨明細的新增
            If Add_flag = 1 Then
                '銷貨明細資料來源
                Call SelDetail_Data
            End If
            Add_flag = 0
            
        Case 1      '取消
            rsSell.CancelUpdate
            If rsSell.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                
                '回覆原來的資料內容
                For Each oText In Me.txtSell
                    Set oText.DataSource = rsSell
                Next
                control_status False
            Else
                Call cmdEdit_Click(3)
                MsgBox "目前並無任何銷貨資料!", , "金光益電器股份有限公司"
            End If
            
            Add_flag = 0
                
        Case 2      '修改
            control_status True
            '下拉式方塊的預設值為原來的員工姓名
            cboEmployee.Text = txtStaff.Text
            fraSelDetail.Visible = True
            
        Case 3      '回主畫面
            '1.由訂單來,表示要新增一張銷貨單(為新增狀態)
            If cmdSell_flag = 1 Then
                '2.完成了新增程序,即產生銷貨單了,故要更新庫存與核銷欄位值
                If update_flag = 1 Then
                    '更新庫存與核銷欄位之副程式
                    Call Data_Check
                    Call frmOrder.dgOrDetail_refresh
                    intRecord = bm_move
                    Call frmOrder.cmdSell_status
                End If
            Else
                '表由功能表的'銷貨單查詢'功能項目來開啟銷貨單(為瀏覽狀態)
                frmMDIMain.mnuExit.Enabled = True
                '回到主畫面時須重整功能項目的狀態
                Call frmMDIMain.mnuSell_status
            End If
                        
            update_flag = 0
            cmdSell_flag = 0
            Unload Me
            
    End Select
End Sub

'資料的移動
Private Sub cmdMove_Click(index As Integer)
    Call rs_Move(index, rsSell)
    lblRecord.Caption = "銷貨資料：" & intRecord & "/" & intTotal
End Sub

'明細的編修
Private Sub cmdSubEdit_Click(index As Integer)
    Select Case index
        Case 0      '明細儲存
            
            '明細儲存的旗標(即必須更新庫存資料與訂單的核銷之旗標)
            update_flag = 1
            
            '將銷貨明細的暫存資料rsSelDetail_list寫進資料庫的銷貨明細資料表中
            rsSelDetail_list.MoveFirst
            Do Until rsSelDetail_list.EOF
                Call Add_SubNo(rsSell_Detail, 2, txtSell(0))
                rsSell_Detail.AddNew
                rsSell_Detail.Fields(0) = txtSell(0).Text
                rsSell_Detail.Fields(1) = Sub_No
                rsSell_Detail.Fields(2) = rsSelDetail_list.Fields("產品編號")
                rsSell_Detail.Fields(3) = RTrim(rsSelDetail_list.Fields("產品名稱"))
                rsSell_Detail.Fields(4) = rsSelDetail_list.Fields("單價")
                rsSell_Detail.Fields(5) = rsSelDetail_list.Fields("銷貨數量")
                rsSell_Detail.Fields(6) = Int(rsSelDetail_list.Fields("單價")) * Int(rsSelDetail_list.Fields("銷貨數量"))
                rsSell_Detail.Update
                
                rsSelDetail_list.MoveNext
            Loop
            
            Call SelDetail_refresh
            subcontrol_status True
            '儲存與取消鈕不可見
            For i = 0 To 1
                cmdSubEdit(i).Visible = False
            Next i
            
        Case 1      '明細取消
            If MsgBox("銷貨明細必須有資料,否則銷貨主檔將被刪除!確定要取消明細資料的新增?", 48 + vbYesNo, "金光益電器股份有限公司") = vbYes Then
                rsSell.Delete
                If rsSell.RecordCount <> 0 Then
                    Call cmdMove_Click(2)
                    
                    subcontrol_status True
                    '儲存與取消鈕不可見
                    For i = 0 To 1
                        cmdSubEdit(i).Visible = False
                    Next i
                Else
                    Call cmdEdit_Click(3)
                    MsgBox "目前並無任何銷貨資料!", , "金光益電器股份有限公司"
                End If
            End If
    End Select
            
End Sub



Private Sub cmdPreview_Click()
RptOrder.Show
End Sub

Private Sub cmdPrinter_Click()
    '檢查是否有安裝印表機
    If Printers.Count <> 0 Then
        RptOrder.PrintReport True, rptRangeAllPages
    Else
        MsgBox ("請先安裝印表機!!")
    End If
End Sub





'主檔的資料連結
Private Sub Form_Load()
    '1.開啟RecordSet
    Set rsSell = New ADODB.Recordset
    sql_rsSell = "select * from 銷貨主檔 order by 銷貨單編號"
    rsSell.Open sql_rsSell, cn, adOpenDynamic, adLockOptimistic
    
    '2.將資料指定給感知元件
    For Each oText In Me.txtSell
        Set oText.DataSource = rsSell
    Next
    Set txtCust.DataSource = rsSell
    
    '3.製作cboEmployee的清單項目
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from 員工資料表"
    rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
    
    cboEmployee.Clear
    Do Until rsStaff.EOF
        cboEmployee.AddItem rsStaff.Fields("姓名")
        rsStaff.MoveNext
    Loop
    
    '4.判斷是從哪裡的指令來開啟銷貨單
    '若由訂單(產生銷貨單 鈕)來的,則直接進入新增狀態
    If cmdSell_flag = 1 Then
        Add_flag = 1
        Call Add_No(rsSell, 4, 1)
        control_status True
        txtSell(0).Text = Main_No
        txtCust.Text = frmOrder.txtCustomer.Text
        cboEmployee.Text = "--請選擇經手人--"
        txtSell(3).Text = Date
    Else
        '由功能表 銷貨單查詢 指令來的,則直接呈現出銷貨資料
        Call cmdMove_Click(0)
        control_status False
        For i = 0 To 1
            cmdSubEdit(i).Visible = False
        Next i
    End If
End Sub

'明細資料的連結
Private Sub SelDetail_refresh()
    '1.開啟銷貨明細,並以銷貨單編號跟銷貨主檔關聯
    Set rsSell_Detail = New ADODB.Recordset
    sql_rsSell_Detail = "select * from 銷貨明細 where 銷貨單編號='" & txtSell(0).Text & "'"
    rsSell_Detail.Open sql_rsSell_Detail, cn, adOpenDynamic, adLockOptimistic
    
    '2.將資料指定給DataGrid
    Set dgSelDetail.DataSource = rsSell_Detail
    
    For i = 0 To 2
        dgSelDetail.Columns.Item(i).Visible = False
    Next i
    For i = 4 To 6
        dgSelDetail.Columns.Item(i).Alignment = dbgRight
    Next i
End Sub

'資料的連動
Private Sub txtSell_Change(index As Integer)
    '銷貨明細資料連結
    Call SelDetail_refresh
    
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from 員工資料表"
    rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
        
    rsStaff.MoveFirst
    rsStaff.Find "員工編號='" & txtSell(1).Text & "'", , adSearchForward, 1
    If rsStaff.EOF = False Then
        txtStaff.Text = rsStaff.Fields("姓名")
        cboEmployee.Text = rsStaff.Fields("姓名")
    End If
End Sub

'主檔控制項的狀態設定
Private Sub control_status(control_flag As Boolean)
    For i = 1 To 2
        txtSell(i).Visible = False
    Next i
    txtSell(0).Enabled = False
    txtCust.Enabled = False
    txtStaff.Enabled = False
    
    txtSell(3).Enabled = control_flag
    txtStaff.Visible = Not control_flag
    cboEmployee.Visible = control_flag
    
    '儲存,取消 鈕
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next i
    '修改,回主畫面 鈕
    For i = 2 To 3
        cmdEdit(i).Visible = Not control_flag
    Next i
    
    fraSelDetail.Visible = Not control_flag
End Sub

'明細控制項的狀態設定
Private Sub subcontrol_status(subcontrol_flag As Boolean)
    fraSell.Enabled = subcontrol_flag
    cmdEdit(3).Enabled = subcontrol_flag
End Sub

'銷貨明細資料來源
Private Sub SelDetail_Data()
    '1.先製作一個空的RecordSet
    Set rsSelDetail_list = New ADODB.Recordset
    rsSelDetail_list.Fields.Append "產品編號", adChar, 3
    rsSelDetail_list.Fields.Append "產品名稱", adChar, 20
    rsSelDetail_list.Fields.Append "單價", adInteger
    rsSelDetail_list.Fields.Append "銷貨數量", adInteger
    rsSelDetail_list.Open
    
    '2.再將相對應的訂單明細資料放於此RecordSet中
    Do Until rsOrder_Detail.EOF
        rsSelDetail_list.AddNew
        rsSelDetail_list.Fields(0) = rsOrder_Detail.Fields("產品編號")
        rsSelDetail_list.Fields(1) = rsOrder_Detail.Fields("產品名稱")
        rsSelDetail_list.Fields(2) = rsOrder_Detail.Fields("單價")
        rsSelDetail_list.Fields(3) = rsOrder_Detail.Fields("訂貨數量")
        rsSelDetail_list.Update
        
        rsOrder_Detail.MoveNext
    Loop
    
    '並將此暫時性的RecordSet-rsSelDetail_list資料呈現於DataGrid-dgSelDetail中
    Set dgSelDetail.DataSource = rsSelDetail_list
    dgSelDetail.Columns.Item(0).Visible = False
    
    subcontrol_status False
End Sub

'更新銷貨核銷欄位為1與生產核銷欄位為3
Private Sub Data_Check()
    Set rsProdtStore = New ADODB.Recordset
    rsProdtStore.Open "select * from 產品庫存統計表", cn, adOpenKeyset, adLockOptimistic
    
    '1.找出銷貨單中的訂單編號=目前訂貨單中的訂單編號之記錄
    rsSell.MoveFirst
    rsSell.Find "訂單編號='" & frmOrder.txtOrder(0).Text & "'", , 1
    '2.有找到,表示此筆訂單已經產生了相對應的銷貨單了,所以銷貨核銷為1,生產核銷為3
    If rsSell.EOF = False Then
        rsOrder_Detail.Filter = "訂單編號='" & frmOrder.txtOrder(0).Text & "' and 生產核銷=2"
        '已生產完畢(可準備銷貨)的訂單明細筆數
        intAccount(1) = rsOrder_Detail.RecordCount
        
        '3.一筆筆更新訂單明細的銷貨核銷與生產核銷欄位
        Do Until rsOrder_Detail.EOF
            rsOrder_Detail.Fields("銷貨核銷") = 1
            rsOrder_Detail.Fields("生產核銷") = 3
            rsOrder_Detail.Update
            
            '4.更新產品庫存統計表中的'銷貨量'欄位值
            '若所生產的數量是用來補足安全存量(客戶為金光益電器股份有限公司本身), _
             則不需更新銷貨量欄位值
            If txtSell(2).Text <> "C000" Then
                rsProdtStore.MoveFirst
                rsProdtStore.Find "產品編號='" & rsOrder_Detail.Fields("產品編號") & "'", , 1
                If rsProdtStore.EOF = False Then
                    rsProdtStore.Fields("銷貨量") = rsProdtStore.Fields("銷貨量") + rsOrder_Detail.Fields("訂貨數量")
                    rsProdtStore.Fields("現存量") = rsProdtStore.Fields("上期庫存量") + rsProdtStore.Fields("生產量") - rsProdtStore.Fields("銷貨量")
                    rsProdtStore.Fields("暫存量") = rsProdtStore.Fields("現存量")
                    rsProdtStore.Update
                End If
            End If
            rsOrder_Detail.MoveNext
        Loop
        rsOrder_Detail.Filter = adFilterNone
                
        '5.當訂單明細的筆數與可準備銷貨的訂單明細筆數一樣, _
           則可更新訂單主檔的銷貨核銷欄位值
        If intAccount(0) = intAccount(1) Then
            rsOrder.Fields("銷貨核銷") = 1
            rsOrder.Update
        End If
    End If
End Sub
