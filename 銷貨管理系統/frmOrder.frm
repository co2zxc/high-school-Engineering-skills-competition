VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOrder 
   Caption         =   "訂貨單"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11880
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdSell 
      Caption         =   "產生銷貨單"
      Height          =   375
      Left            =   6720
      TabIndex        =   41
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "回主畫面"
      Height          =   375
      Index           =   5
      Left            =   7920
      TabIndex        =   31
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraOrder 
      Caption         =   "訂單主檔"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   2160
      TabIndex        =   15
      Top             =   600
      Width           =   6975
      Begin VB.CommandButton cmdMove 
         Caption         =   "<<"
         Height          =   355
         Index           =   1
         Left            =   1560
         TabIndex        =   40
         Top             =   2280
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "|<"
         Height          =   355
         Index           =   0
         Left            =   1200
         TabIndex        =   39
         Top             =   2280
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">|"
         Height          =   355
         Index           =   3
         Left            =   4920
         TabIndex        =   38
         Top             =   2280
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">>"
         Height          =   355
         Index           =   2
         Left            =   4560
         TabIndex        =   37
         Top             =   2280
         Width           =   350
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "新     增"
         Height          =   375
         Index           =   2
         Left            =   5640
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "修     改"
         Height          =   375
         Index           =   3
         Left            =   5640
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "刪      除"
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtOrder 
         DataField       =   "訂單編號"
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
         Left            =   1920
         TabIndex        =   19
         Top             =   360
         Width           =   2685
      End
      Begin VB.TextBox txtOrder 
         Alignment       =   1  '靠右對齊
         DataField       =   "訂貨日"
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
         Left            =   1920
         TabIndex        =   17
         Top             =   1800
         Width           =   2685
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "儲     存"
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "取     消"
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtOrder 
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
         Left            =   1920
         TabIndex        =   18
         Top             =   840
         Width           =   2685
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
         Left            =   1920
         TabIndex        =   33
         Top             =   840
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo dcboStaff 
         Height          =   390
         Left            =   1920
         TabIndex        =   16
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.TextBox txtOrder 
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
         Index           =   2
         Left            =   1920
         TabIndex        =   26
         Top             =   1320
         Width           =   2685
      End
      Begin VB.TextBox txtCustomer 
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
         Left            =   1920
         TabIndex        =   34
         Top             =   1320
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo dcboCust 
         Height          =   390
         Left            =   1920
         TabIndex        =   25
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
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
         Left            =   1920
         TabIndex        =   36
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         Caption         =   "訂單編號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   30
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         Caption         =   "經手人"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   29
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         Caption         =   "客戶名稱"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   990
         TabIndex        =   28
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         Caption         =   "訂貨日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1200
         TabIndex        =   27
         Top             =   1920
         Width           =   630
      End
   End
   Begin VB.Frame fraOrDetail 
      Caption         =   "訂單明細"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2160
      TabIndex        =   0
      Top             =   3600
      Width           =   6975
      Begin VB.TextBox txtProdt 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cboProdt 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmOrder.frx":0000
         Left            =   240
         List            =   "frmOrder.frx":0002
         TabIndex        =   32
         Top             =   720
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgOrDetail 
         Height          =   2055
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         Caption         =   "刪      除"
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   14
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "修      改"
         Height          =   375
         Index           =   3
         Left            =   5640
         TabIndex        =   13
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "新      增"
         Height          =   375
         Index           =   2
         Left            =   5640
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "取     消"
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "儲     存"
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtOrDetail 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4440
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtOrDetail 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtOrDetail 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtOrDetail 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblOrDetail 
         AutoSize        =   -1  'True
         Caption         =   "總     價"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   4680
         TabIndex        =   9
         Top             =   480
         Width           =   690
      End
      Begin VB.Label lblOrDetail 
         AutoSize        =   -1  'True
         Caption         =   "訂 貨 數 量"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   3360
         TabIndex        =   8
         Top             =   480
         Width           =   960
      End
      Begin VB.Label lblOrDetail 
         AutoSize        =   -1  'True
         Caption         =   "單    價"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   480
         Width           =   630
      End
      Begin VB.Label lblOrDetail 
         AutoSize        =   -1  'True
         Caption         =   "產 品 名 稱"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   720
         TabIndex        =   6
         Top             =   480
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bm_detail As Variant            '記錄明細修改前之記錄
Dim Prodt_Name As String            '記錄明細修改前cboProdt的值
Dim intpre_NumCk(1) As Integer      '記錄修改前的訂貨數量與生產核銷欄位

'主檔資料的編修
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0     '主檔儲存
            '主檔資料必須完整,明細資料才能填寫
            rsStaff.MoveFirst
            rsStaff.Find "姓名='" & dcboStaff.Text & "'", , 1
            If rsStaff.EOF Then
                MsgBox "請選擇經手人!", , "金光益電器股份有限公司"
                dcboStaff.SetFocus
                Exit Sub
            Else
                rsCust.MoveFirst
                rsCust.Find "公司名稱='" & dcboCust.Text & "'", , 1
                If rsCust.EOF Then
                    MsgBox "請選擇客戶!", , "金光益電器股份有限公司"
                    dcboCust.SetFocus
                    Exit Sub
                Else
                    If txtOrder(3).Text = "" Then
                        MsgBox "請填寫訂貨日!", , "金光益電器股份有限公司"
                        txtOrder(3).SetFocus
                        Exit Sub
                    Else
                        If IsDate(txtOrder(3).Text) = False Then
                            MsgBox "您輸入的日期格式有誤,請再確認一遍!", , "金光益電器股份有限公司"
                            txtOrder(3).SetFocus
                            Exit Sub
                        Else
                            txtOrder(1).Text = dcboStaff.BoundText
                            txtOrder(2).Text = dcboCust.BoundText
                            If Add_flag = 1 Then
                                '新增一筆記錄時,銷貨核銷欄位的預設值為0
                                rsOrder.Fields("銷貨核銷") = 0
                            End If
                            rsOrder.Update
                            Call cmdMove_Click(3)
                            control_status False
                            cmdSell.Visible = True
                        End If
                    End If
                End If
            End If
            
            '主檔在"新增"狀態下,儲存後即進入明細新增狀態
            If Add_flag = 1 Then
                Call cmdSubEdit_Click(2)
            End If
            
            Add_flag = 0

        Case 1     '主檔取消
            rsOrder.CancelUpdate
            If rsOrder.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                dcboStaff.BoundText = txtOrder(1).Text
                dcboCust.BoundText = txtOrder(2).Text
                For Each oText In Me.txtOrder
                    Set oText.DataSource = rsOrder
                Next
                control_status False
                cmdSell.Visible = True
            Else
                Call cmdEdit_Click(5)
            End If
            
            Add_flag = 0
            
        Case 2     '主檔新增
            '將主檔與明細編修鈕顯示出來,再做狀態的切換
           ' For i = 0 To 4
           '     cmdEdit(i).Enabled = True
           '     cmdSubEdit(i).Enabled = True
           ' Next i
            
            Add_flag = 1
            Call Add_No(rsOrder, 3, 1)
            control_status True
            cmdSell.Visible = False
            txtOrder(0).Text = Main_No
            txtOrder(3).Text = Date
            dcboStaff.Text = "請選擇經手人"
            dcboCust.Text = "--請選擇客戶--"

        Case 3     '主檔修改
            control_status True
            '修改時客戶欄位不可更改,故顯示出txtCustomer
            dcboCust.Visible = False
            txtCustomer.Visible = True
            fraOrDetail.Visible = True
            fraOrDetail.Enabled = False
            
            cmdSell.Visible = False
            
        Case 4     '主檔刪除
            If MsgBox("確定要刪除此筆記錄??", 32 + vbYesNo, "金光益電器股份有限公司") = vbYes Then
                rsOrder.Delete
                If rsOrder.RecordCount <> 0 Then
                    Call cmdMove_Click(2)
                Else
                    Call cmdEdit_Click(5)
                    MsgBox "目前已無任何訂單資料!", , "金光益電器股份有限公司"
                End If
            End If
            
        Case 5     '回主畫面
            Set rsOrder_Detail = New ADODB.Recordset
            rsOrder_Detail.Open "select * from 訂單明細 where 生產核銷=0 order by 產品編號", cn, adOpenDynamic, adLockOptimistic
            
            '若目前還有未生產之產品,即會出現生產單
            If rsOrder_Detail.RecordCount <> 0 Then
                frmMake.Show 1
            End If
                
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            Call frmMDIMain.mnuSell_status
    End Select
End Sub

'資料的移動
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsOrder)
    lblRecord.Caption = "訂單資料：" & intRecord & "/" & intTotal
    
    Call cmdSell_status
End Sub

'產生銷貨單
Private Sub cmdSell_Click()
    cmdSell_flag = 1
    bm_move = intRecord
    frmSell.Show
End Sub

'明細資料的編修
Private Sub cmdSubEdit_Click(Index As Integer)
    Select Case Index
        Case 0    '明細儲存
            '檢查資料是否填寫完整
            rsProdt.MoveFirst
            rsProdt.Find "產品名稱='" & cboProdt.Text & "'", , 1
            If rsProdt.EOF Then
                MsgBox "請選擇訂購的產品!", , "金光益電器股份有限公司"
                cboProdt.SetFocus
                Exit Sub
            Else
                For i = 1 To 2
                    If txtOrDetail(i).Text = "" Then
                        MsgBox "請填寫" & lblOrDetail(i).Caption & "!", , "金光益電器股份有限公司"
                        txtOrDetail(i).SetFocus
                        Exit Sub
                    Else
                        If IsNumeric(txtOrDetail(i).Text) = False Then
                            MsgBox "請填寫數字格式之資料!", , "金光益電器股份有限公司"
                            txtOrDetail(i).SetFocus
                            Exit Sub
                        End If
                    End If
                Next i
                
                '總價
                txtOrDetail(3).Text = Int(txtOrDetail(1).Text) * Int(txtOrDetail(2).Text)
                
                '按下新增鈕後的儲存狀態,則會呼叫AddNew
                If Edit_flag <> 1 Then
                    rsOrder_Detail.AddNew
                    rsOrder_Detail.Fields(0) = txtOrder(0).Text
                    rsOrder_Detail.Fields(1) = Sub_No
                    '銷貨核銷與生產核銷欄位預設為0(即未核銷)
                    For i = 7 To 8
                        rsOrder_Detail.Fields(i) = 0
                    Next i
                End If
                
                '產品編號
                txtOrDetail(0).Text = rsProdt.Fields(0)
                rsOrder_Detail.Fields("產品編號") = txtOrDetail(0).Text
                For i = 1 To 3
                   rsOrder_Detail.Fields(i + 3) = txtOrDetail(i).Text
                Next i
                rsOrder_Detail.Fields("產品名稱") = cboProdt.Text
                rsOrder_Detail.Update
            
                subcontrol_status False
                cmdSell.Visible = True
                lblOrDetail(3).Visible = True
                txtOrDetail(3).Visible = True
            End If
                
            Set rsMake = New ADODB.Recordset
            rsMake.Open "select * from 生產核銷表 order by  產品編號", cn, adOpenDynamic, adLockOptimistic
            
            '修改時
            If Edit_flag = 1 Then
                rsOrder_Detail.Bookmark = bm_detail
                
                '1.若訂貨數量被更改了
                If intpre_NumCk(0) <> Int(rsOrder_Detail.Fields("訂貨數量")) Then
                    '2.此時如果已經產生生產單了(生產核銷=1),甚至是已完成產品的生產了(生產核銷=2)
                    If intpre_NumCk(1) = 1 Or intpre_NumCk(1) = 2 Then
                        If rsMake.RecordCount <> 0 Then
                            rsMake.MoveFirst
                        End If
                        rsMake.Find "產品編號='" & rsOrder_Detail.Fields("產品編號") & "'", , 1
                        
                        '3.若訂貨數量增多了,那麼增加的數量亦要加入生產單中
                        If Sgn(rsOrder_Detail.Fields("訂貨數量") - intpre_NumCk(0)) = 1 Then
                            If rsMake.EOF Then
                                rsMake.AddNew
                                rsMake.Fields(0) = rsOrder_Detail.Fields("產品編號")
                                rsMake.Fields(1) = rsOrder_Detail.Fields("產品名稱")
                                '所要生產的訂貨數量必須減去修改前的數量
                                rsMake.Fields(2) = Int(rsOrder_Detail.Fields("訂貨數量") - intpre_NumCk(0))
                            Else
                                rsMake.Fields(2) = rsMake.Fields(2) + Int(rsOrder_Detail.Fields("訂貨數量") - intpre_NumCk(0))
                            End If
                        Else
                            '4.若訂貨數量減少了,那麼如果生產單中有此產品,則可將減少的產品扣除 _
                               如果產品已經生產了,那就先以庫存量儲存於倉庫中
                            If rsMake.EOF = False Then
                                rsMake.Fields(2) = rsMake.Fields(2) - Abs(rsOrder_Detail.Fields("訂貨數量") - intpre_NumCk(0))
                            Else
                                Call dgOrDetail_refresh
                                '清除明細的旗標值
                                Edit_flag = 0
                                Exit Sub
            
                            End If
                        End If
                        rsMake.Update
                        
                        '5.並將生產核銷欄位改為1,以便在進入生產階段
                        rsOrder_Detail.Fields("生產核銷") = 1
                        rsOrder_Detail.Update
                    End If
                End If
            End If
                            
            Call dgOrDetail_refresh
            Call cmdSell_status
            '清除明細的旗標值
            Edit_flag = 0
         
        Case 1    '明細取消
            '明細在第一筆的"新增"狀態下(明細中未有值時)
            If rsOrder_Detail.RecordCount = 0 Then
                If MsgBox("訂單明細中必須有資料,否則訂單主檔將被刪除!!", 48 + vbYesNo, "金光益電器股份有限公司") = vbYes Then
                    rsOrder.Delete
                    '刪除後若主檔還有記錄,則進行記錄的refresh
                    If rsOrder.RecordCount <> 0 Then
                        Call cmdMove_Click(2)
                    Else
                       '若刪除後訂貨單為空的,即回到主畫面
                        MsgBox "訂貨單中已無記錄!!", , "金光益電器股份有限公司"
                        Call cmdEdit_Click(5)
                    End If
                Else
                    '否則跳出訊息方塊
                    Exit Sub
                End If
            Else
                Call dgOrDetail_refresh
            End If
            
            lblOrDetail(3).Visible = True
            txtOrDetail(3).Visible = True
            '回覆表單載入時的狀態
            subcontrol_status False
            cmdSell.Visible = True
    
            Edit_flag = 0
                    
        Case 2    '明細新增
            Call cboProdt_AddItem
            subcontrol_status True
            cmdSell.Visible = False
            Call Add_SubNo(rsOrder_Detail, 1, txtOrder(0))
            For i = 0 To 2
                txtOrDetail(i).Text = ""
            Next i
            cboProdt.Text = "--請選擇產品--"
            '總價欄位不可見
            txtOrDetail(3).Visible = False
            lblOrDetail(3).Visible = False
            
        Case 3    '明細修改
            '修改明細的旗標值
            Edit_flag = 1
            
            '1.記錄目前記錄的位置
            bm_detail = rsOrder_Detail.Bookmark
            '2.記錄目前記錄的訂貨數量與生產欄銷欄位值
            intpre_NumCk(0) = rsOrder_Detail.Fields("訂貨數量")
            intpre_NumCk(1) = rsOrder_Detail.Fields("生產核銷")
            
            '3.讓cboProdt呈現原來的產品名稱
            '(1)記錄原來cboProdt的值
            Prodt_Name = rsOrder_Detail.Fields(3)
            '(2)重整cboProdt中的值
            Call cboProdt_AddItem
            '(3)呈現原來cboProdt上的值
            cboProdt.Text = Prodt_Name
            
            subcontrol_status True
            cmdSell.Visible = False
            
        Case 4    '明細刪除
            '有一筆以上的明細資料時,則直接刪除此筆記錄即可
            If MsgBox("確定要刪除此筆記錄??", 32 + vbYesNo, "金光益電器股份有限公司") = vbYes Then
                rsOrder_Detail.Delete
                If rsOrder_Detail.RecordCount <> 0 Then
                    rsOrder_Detail.MoveNext
                Else
                    MsgBox "由於此筆訂單已無任何明細資料,故主檔資料將被刪除!", , "金光益電器股份有限公司"
                    rsOrder.Delete
                    '刪除後若主檔還有記錄,則進行記錄的refresh
                    If rsOrder.RecordCount <> 0 Then
                        Call cmdMove_Click(2)
                    Else
                        Call cmdEdit_Click(5)
                        '若刪除後訂貨單為空的,就再進入主檔新增的狀態
                        MsgBox "訂貨單中已無記錄!!", , "玫瑰針針公司"
                    End If
                End If
            End If
    End Select
End Sub

'讓DataGrid中的記錄顯示於txtOrDetail上
Private Sub dgOrDetail_Click()
    If rsOrder_Detail.RecordCount <> 0 Then
        For i = 1 To 3
            txtOrDetail(i).Text = rsOrder_Detail.Fields(i + 3)
        Next i
        txtOrDetail(0).Text = rsOrder_Detail.Fields(2)
        txtProdt.Text = rsOrder_Detail.Fields("產品名稱")
    End If
End Sub

'資料的連動
Private Sub txtOrder_Change(Index As Integer)
    Select Case Index
        Case 0      '訂單編號
            Call dgOrDetail_refresh
            
        Case 1      '經手人編號
            Set rsStaff = New ADODB.Recordset
            sql_rsStaff = "select * from 員工資料表"
            rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
            
            rsStaff.MoveFirst
            rsStaff.Find "員工編號='" & txtOrder(1).Text & "'", , adSearchForward, 1
            If rsStaff.EOF = False Then
                txtStaff.Text = rsStaff.Fields("姓名")
            End If

        Case 2      '客戶編號
            Set rsCust = New ADODB.Recordset
            sql_rsCust = "select * from 客戶資料"
            rsCust.Open sql_rsCust, cn, adOpenDynamic, adLockOptimistic
            
            rsCust.Find "客戶編號='" & txtOrder(2).Text & "'", , adSearchForward, 1
            If rsCust.EOF = False Then
                txtCustomer.Text = rsCust.Fields("公司名稱")
            End If
    End Select
End Sub

'在主檔的經手人欄位共放了三個感知元件,一個為儲存用txtOrder(1),
'一個為顯示用txtStaff,一個為編修資料時用dcboStaff ;在客戶欄位
'也放了三個感知元件 , 儲存用的txtOrder(2), 顯示用的txtCustomer
'與編修用的dcboCust

'主檔資料的連結
Private Sub Form_Load()
    '1.開啟訂單主檔
    Set rsOrder = New ADODB.Recordset
    sql_rsOrder = "select * from 訂單主檔"
    rsOrder.Open sql_rsOrder, cn, adOpenDynamic, adLockOptimistic
    
    '2.將資料指定給主檔內的感知元件
    For Each oText In Me.txtOrder
        Set oText.DataSource = rsOrder
    Next
    
    '3.製作DataCombo
    '製作dcboStaff(編修資料時用之)
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from 員工資料表"
    rsStaff.Open sql_rsStaff, cn, adOpenKeyset, adLockReadOnly
    Set dcboStaff.DataSource = rsOrder
    dcboStaff.DataField = "經手人編號"
    Set dcboStaff.RowSource = rsStaff
    dcboStaff.ListField = "姓名"
    dcboStaff.BoundColumn = "員工編號"
    
    '製作dcboCust(編修資料時用之)
    Set rsCust = New ADODB.Recordset
    sql_rsCust = "select * from 客戶資料"
    rsCust.Open sql_rsCust, cn, adOpenKeyset, adLockReadOnly
    Set dcboCust.DataSource = rsOrder
    dcboCust.DataField = "客戶編號"
    Set dcboCust.RowSource = rsCust
    dcboCust.ListField = "公司名稱"
    dcboCust.BoundColumn = "客戶編號"
    
    '4.檢查訂貨單中是否有資料
    '若訂貨單為空的則直接進入主檔的新增狀態
    If rsOrder.RecordCount = 0 Then
        Call cmdEdit_Click(2)
    Else
        '表單載入時,控制項的狀態設定
        control_status False
        subcontrol_status False
        Call cmdMove_Click(0)
    End If
    
End Sub

'明細資料的連結
Public Sub dgOrDetail_refresh()
    '1.開啟訂單明細
    Set rsOrder_Detail = New ADODB.Recordset
    sql_rsOrder_Detail = "select * from 訂單明細 where 訂單編號='" & txtOrder(0).Text & "'"
    rsOrder_Detail.Open sql_rsOrder_Detail, cn, adOpenDynamic, adLockOptimistic
    
    '此筆訂單的明細筆數
    intAccount(0) = rsOrder_Detail.RecordCount
    
    '2.將資料指定給DataGrid,並將核銷欄位隱藏起來
    Set dgOrDetail.DataSource = rsOrder_Detail
    For i = 0 To 2
        dgOrDetail.Columns.Item(i).Visible = False
    Next i
    For i = 7 To 8
        dgOrDetail.Columns.Item(i).Visible = False
    Next i
    
    dgOrDetail.Columns.Item(3).Width = 2000
    For i = 4 To 6
        dgOrDetail.Columns.Item(i).Alignment = dbgRight
    Next i
    
    '3.明細資料載入時,控制項的狀態設定
    subcontrol_status False
        
    '4.讓DataGrid中記錄顯示於文字方塊中
    Call dgOrDetail_Click
End Sub

'主檔控制項的狀態設定
Private Sub control_status(control_flag As Boolean)
    '固定不變的狀態
    For i = 0 To 4
        cmdEdit(i).Enabled = True
    Next i
    txtOrder(0).Enabled = False
    txtOrder(1).Visible = False
    txtOrder(2).Visible = False
    txtStaff.Enabled = False
    txtCustomer.Enabled = False
    
    '可切換的狀態
    For i = 1 To 3
        txtOrder(i).Enabled = control_flag
    Next
    dcboStaff.Visible = control_flag
    dcboCust.Visible = control_flag
    txtStaff.Visible = Not control_flag
    txtCustomer.Visible = Not control_flag
    '儲存,取消鈕
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next i
    '新增,修改,刪除,回主畫面鈕
    For i = 2 To 5
        cmdEdit(i).Visible = Not control_flag
    Next i
    '移動鈕
    For i = 0 To 3
        cmdMove(i).Enabled = Not control_flag
    Next i
    
    '主檔新增時,fraOrDetail會隱藏;主檔修改時,fraOrDetail會顯示,但不可操作
    fraOrDetail.Enabled = Not control_flag
    fraOrDetail.Visible = Not control_flag
    
End Sub

'明細控制項的狀態設定
Private Sub subcontrol_status(subcontrol_flag As Boolean)
    '固定不變的狀態
    For i = 0 To 4
        cmdSubEdit(i).Enabled = True
    Next i
    txtOrDetail(0).Visible = False
    txtOrDetail(3).Enabled = False
    txtProdt.Enabled = False
    
    '可切換的狀態
    For i = 1 To 2
        txtOrDetail(i).Enabled = subcontrol_flag
    Next
    cboProdt.Visible = subcontrol_flag
    txtProdt.Visible = Not subcontrol_flag
    dgOrDetail.Enabled = Not subcontrol_flag
    '儲存,取消鈕
    For i = 0 To 1
        cmdSubEdit(i).Visible = subcontrol_flag
    Next i
    '新增,修改,刪除鈕
    For i = 2 To 4
        cmdSubEdit(i).Visible = Not subcontrol_flag
    Next i
    
    '當明細編修時,主檔的控制項狀態
    fraOrder.Enabled = Not subcontrol_flag
    '回主畫面鈕
    cmdEdit(5).Enabled = Not subcontrol_flag
    
End Sub

'製作cboProdt
Private Sub cboProdt_AddItem()
    '1.開啟產品資料
    Set rsProdt = New ADODB.Recordset
    sql_rsProdt = "select * from 產品資料"
    rsProdt.Open sql_rsProdt, cn, adOpenKeyset, adLockReadOnly
    
    '2.先清除cboProdt中的值
    cboProdt.Clear
    '3.再將產品名稱加入cboProdt中
    Do Until rsProdt.EOF
        '訂單明細中已有的產品,則不再出現
        If rsOrder_Detail.RecordCount <> 0 Then
            rsOrder_Detail.MoveFirst
        End If
        rsOrder_Detail.Find "產品編號='" & rsProdt.Fields(0) & "'", , 1
        If rsOrder_Detail.EOF Then
            cboProdt.AddItem rsProdt.Fields("產品名稱")
        End If
        rsProdt.MoveNext
    Loop
End Sub

Private Sub cboProdt_Click()
    rsProdt.MoveFirst
    rsProdt.Find "產品名稱='" & cboProdt.Text & "'", , adSearchForward, 1
    If rsProdt.EOF = False Then
        '將所點選cboProdt中的值寫回txtOrDetail(0)中
        txtOrDetail(0).Text = rsProdt.Fields("產品編號")
        '單價欄位的預設值為建議售價
        txtOrDetail(1).Text = rsProdt.Fields("建議售價")
    End If
End Sub

'表單的可操作與否
Public Sub cmdSell_status()
    cmdSell.Enabled = False
    '訂單有資料時
    If rsOrder.EOF = False And rsOrder.BOF = False Then
        '未銷貨前訂單的編修鈕皆可使用,產生銷貨單鈕則要視情況來決定可否被使用
        If rsOrder.Fields("銷貨核銷") = 0 Then
            rsOrder_Detail.Filter = "訂單編號='" & txtOrder(0).Text & "' and 生產核銷=2"
            intAccount(1) = rsOrder_Detail.RecordCount
            
            '此張訂單的產品都已生產完畢時,才可進入銷貨階段
            If intAccount(0) = intAccount(1) Then
                cmdSell.Enabled = True
            Else
                cmdSell.Enabled = False
            End If
            '未銷貨之前,訂單資料還可修改
            For i = 0 To 4
                cmdEdit(i).Enabled = True
                cmdSubEdit(i).Enabled = True
            Next i
            
            rsOrder_Detail.Filter = adFilterNone
            Call dgOrDetail_refresh
        Else
            '銷貨單產生後,訂單的編修鈕與產生銷貨單鈕皆不可使用,
            cmdSell.Enabled = False
            For i = 0 To 4
                cmdEdit(i).Enabled = False
                cmdSubEdit(i).Enabled = False
            Next i
            '只有主檔的新增鈕可使用
            cmdEdit(2).Enabled = True
        End If
    End If
End Sub

