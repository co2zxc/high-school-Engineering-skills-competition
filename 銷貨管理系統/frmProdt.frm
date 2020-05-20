VERSION 5.00
Begin VB.Form frmProdt 
   Caption         =   "產品資料表"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   8385
   Begin VB.TextBox txtProdt 
      Alignment       =   1  '靠右對齊
      DataField       =   "類別"
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
      Index           =   8
      Left            =   1680
      TabIndex        =   28
      Top             =   4200
      Width           =   2500
   End
   Begin VB.TextBox txtProdt 
      Alignment       =   1  '靠右對齊
      DataField       =   "品牌"
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
      Index           =   7
      Left            =   1680
      TabIndex        =   26
      Top             =   3240
      Width           =   2500
   End
   Begin VB.TextBox txtProdt 
      Alignment       =   1  '靠右對齊
      DataField       =   "類別編號"
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
      Left            =   1680
      TabIndex        =   25
      Top             =   3720
      Width           =   2500
   End
   Begin VB.TextBox txtProdt 
      Alignment       =   1  '靠右對齊
      DataField       =   "廠商編號"
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
      Left            =   1680
      TabIndex        =   24
      Top             =   2760
      Width           =   2500
   End
   Begin VB.TextBox txtProdt 
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
      Index           =   1
      Left            =   1680
      TabIndex        =   20
      Top             =   840
      Width           =   2500
   End
   Begin VB.TextBox txtProdt 
      Alignment       =   1  '靠右對齊
      DataField       =   "建議售價"
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
      Left            =   1680
      TabIndex        =   18
      Top             =   2280
      Width           =   2500
   End
   Begin VB.TextBox txtProdt 
      Alignment       =   1  '靠右對齊
      DataField       =   "單位成本"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   2500
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   4440
      TabIndex        =   15
      Top             =   4800
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   4800
      TabIndex        =   14
      Top             =   4800
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Top             =   4800
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   840
      TabIndex        =   12
      Top             =   4800
      Width           =   350
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "回主畫面"
      Height          =   375
      Index           =   5
      Left            =   4680
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "取     消"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "儲     存"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "刪     除"
      Height          =   375
      Index           =   4
      Left            =   4680
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "新     增"
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "修     改"
      Height          =   375
      Index           =   3
      Left            =   4680
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtProdt 
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
      Index           =   2
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   2500
   End
   Begin VB.TextBox txtProdt 
      Alignment       =   1  '靠右對齊
      DataField       =   "產品編號"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   360
      Width           =   2500
   End
   Begin VB.Label lblProdt 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "類別"
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
      Index           =   8
      Left            =   600
      TabIndex        =   27
      Top             =   4320
      Width           =   960
   End
   Begin VB.Label lblProdt 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "廠商"
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
      Index           =   7
      Left            =   840
      TabIndex        =   23
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label lblProdt 
      AutoSize        =   -1  'True
      Caption         =   "類別編號"
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
      Index           =   6
      Left            =   600
      TabIndex        =   22
      Top             =   3840
      Width           =   960
   End
   Begin VB.Label lblProdt 
      AutoSize        =   -1  'True
      Caption         =   "廠商編號"
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
      Left            =   600
      TabIndex        =   21
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label lblProdt 
      AutoSize        =   -1  'True
      Caption         =   "建議售價"
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
      Left            =   600
      TabIndex        =   19
      Top             =   2355
      Width           =   960
   End
   Begin VB.Label lblProdt 
      AutoSize        =   -1  'True
      Caption         =   "單位成本"
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
      Left            =   600
      TabIndex        =   17
      Top             =   1875
      Width           =   960
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
      Left            =   1200
      TabIndex        =   16
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label lblProdt 
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
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   1395
      Width           =   960
   End
   Begin VB.Label lblProdt 
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
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   915
      Width           =   975
   End
   Begin VB.Label lblProdt 
      AutoSize        =   -1  'True
      Caption         =   "產品編號"
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
      Left            =   600
      TabIndex        =   4
      Top             =   435
      Width           =   975
   End
End
Attribute VB_Name = "frmProdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'資料的編修
Private Sub cmdEdit_Click(index As Integer)
    Select Case index
        Case 0      '儲存
            If txtProdt(1).Text = "" Then
                MsgBox "請填寫產品名稱!", , "金光益電器股份有限公司"
                txtProdt(1).SetFocus
                Exit Sub
            Else
                For i = 2 To 6
                    If txtProdt(i).Text = "" Then
                        MsgBox "" & lblProdt(i).Caption & "欄位不能空白!", , "金光益電器股份有限公司"
                        txtProdt(i).SetFocus
                        Exit Sub
                    Else
                        If IsNumeric(txtProdt(i).Text) = False Then
                            MsgBox "" & lblProdt(i).Caption & "需填數字格式之資料!", , "金光益電器股份有限公司"
                            txtProdt(i).SetFocus
                            Exit Sub
                        End If
                    End If
                Next i
            End If
            rsProdt.Update
            control_status False

        Case 1      '取消
            rsProdt.CancelUpdate
            If rsProdt.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                For Each oText In Me.txtProdt
                    Set oText.DataSource = rsProdt
                Next
                control_status False
            Else
                Call cmdEdit_Click(5)
                MsgBox "目前並無任何產品資料!", , "金光益電器股份有限公司"
            End If

        Case 2      '新增
            Add_flag = 1
            Call Add_No(rsProdt, 5, 2)
            control_status True
            txtProdt(0).Text = Main_No
            For i = 2 To 8
                txtProdt(i).Text = 0
            Next i
            
        Case 3      '修改
            control_status True
            
        Case 4      '刪除
            If MsgBox("確定要刪除此筆記錄??", 32 + vbYesNo, "金光益電器股份有限公司") = vbYes Then
                rsProdt.Delete
                If rsProdt.RecordCount <> 0 Then
                    rsProdt.MoveNext
                Else
                    Call cmdEdit_Click(5)
                    MsgBox "目前已無任何產品資料!", , "金光益電器股份有限公司"
                End If
            End If
            
        Case 5      '回主畫面
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            rsProdt.Close
            Set rsProdt = Nothing
    End Select
End Sub

'資料的移動
Private Sub cmdMove_Click(index As Integer)
    Call rs_Move(index, rsProdt)
    lblRecord.Caption = "產品資料：" & intRecord & "/" & intTotal
End Sub

'資料的連結
Private Sub Form_Load()
    '1.開啟產品資料
    Set rsProdt = New ADODB.Recordset
    sql_rsProdt = "select * from 產品資料"
    rsProdt.Open sql_rsProdt, cn, adOpenDynamic, adLockOptimistic
    
    '2.將資料指定給表單中的感知元件
    For Each oText In Me.txtProdt
        Set oText.DataSource = rsProdt
    Next
    
    '3.檢查是否有資料
    If rsProdt.RecordCount <> 0 Then
        Call cmdMove_Click(0)
        '表單載入時的狀態設定
        control_status False
    Else
        '進入新增狀態
        Call cmdEdit_Click(2)
    End If
    
    '4.表單的長寬設定
    frmProdt.Height = 4290
    frmProdt.Width = 6495
End Sub

'控制項狀態設定
Private Sub control_status(control_flag As Boolean)
    txtProdt(0).Enabled = False
    For i = 1 To 8
        txtProdt(i).Enabled = control_flag
    Next i
    
    '儲存,取消鈕
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next i
    '新增,修改,刪除,華回主畫面鈕
    For i = 2 To 5
        cmdEdit(i).Visible = Not control_flag
    Next i
    '移動鈕
    For i = 0 To 3
        cmdMove(i).Enabled = Not control_flag
    Next i
End Sub

