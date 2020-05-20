VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmStaff 
   Caption         =   "員工資料表"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   9825
   Begin VB.CommandButton cmdRelation 
      Caption         =   "聯絡人資料"
      Height          =   375
      Left            =   8160
      TabIndex        =   36
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   960
      TabIndex        =   32
      Top             =   3720
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   600
      TabIndex        =   31
      Top             =   3720
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   7320
      TabIndex        =   30
      Top             =   3720
      Width           =   315
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   6960
      TabIndex        =   29
      Top             =   3720
      Width           =   350
   End
   Begin VB.Frame fraMarry 
      Caption         =   "婚姻狀況"
      Height          =   615
      Left            =   5040
      TabIndex        =   28
      Top             =   720
      Width           =   2055
      Begin VB.OptionButton optMarry 
         Caption         =   "已婚"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optMarry 
         Caption         =   "未婚"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraSex 
      Caption         =   "性別"
      Height          =   615
      Left            =   1560
      TabIndex        =   25
      Top             =   720
      Width           =   2055
      Begin VB.OptionButton optSex 
         Caption         =   "女"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optSex 
         Caption         =   "男"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
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
      TabIndex        =   24
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "修     改"
      Height          =   375
      Index           =   3
      Left            =   8160
      TabIndex        =   23
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "新     增"
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   22
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "刪     除"
      Height          =   375
      Index           =   4
      Left            =   8160
      TabIndex        =   21
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "儲     存"
      Height          =   375
      Index           =   0
      Left            =   8160
      TabIndex        =   20
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "取     消"
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   19
      Top             =   2040
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo dcboArea 
      Height          =   390
      Left            =   1560
      TabIndex        =   18
      Top             =   2880
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
   Begin VB.CommandButton cmdEdit 
      Caption         =   "回主畫面"
      Height          =   375
      Index           =   5
      Left            =   8160
      TabIndex        =   17
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtStaff 
      DataField       =   "地區編號"
      Height          =   375
      Index           =   7
      Left            =   1560
      TabIndex        =   16
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtStaff 
      DataField       =   "住址"
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
      Left            =   2880
      TabIndex        =   15
      Top             =   2880
      Width           =   4215
   End
   Begin VB.TextBox txtStaff 
      DataField       =   "學歷"
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
      Left            =   5040
      TabIndex        =   14
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
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
      TabIndex        =   13
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
      Alignment       =   1  '靠右對齊
      DataField       =   "出生日期"
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
      Left            =   5040
      TabIndex        =   12
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
      Alignment       =   1  '靠右對齊
      DataField       =   "身分證字號"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
      DataField       =   "姓名"
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
      Left            =   5040
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
      Alignment       =   1  '靠右對齊
      DataField       =   "員工編號"
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
      TabIndex        =   9
      Top             =   240
      Width           =   2055
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
      TabIndex        =   33
      Top             =   3720
      Width           =   5655
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "婚姻狀況"
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
      Left            =   3960
      TabIndex        =   8
      Top             =   840
      Width           =   960
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "學歷"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "住址"
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
      Left            =   960
      TabIndex        =   6
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label lblStaff 
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
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "出生日期"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "身分證字號"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "性別"
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
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
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
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      Caption         =   "員工編號"
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
      Width           =   960
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'資料的編修
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '儲存
            For i = 1 To 6
                If txtStaff(i).Text = "" Then
                    MsgBox "抱歉!資料不能空白!", , "金光益電器股份有限公司"
                    txtStaff(i).SetFocus
                    Exit Sub
                Else
                    If IsDate(txtStaff(3).Text) = False Then
                        MsgBox "您輸入的日期格式有誤,請再確認一遍!", , "金光益電器股份有限公司"
                        txtStaff(3).SetFocus
                        Exit Sub
                    End If
                End If
            Next i
            
      
                    rsArea.MoveFirst
                    rsArea.Find "地區名稱='" & dcboArea.Text & "'", , 1
                    If rsArea.EOF Then
                        MsgBox "請選擇正確的地區名稱!", , "金光益電器股份有限公司"
                        dcboArea.SetFocus
                        Exit Sub
                    Else
                        txtStaff(7).Text = dcboArea.BoundText
                        rsStaff.Update
                        control_status False
                    End If
             
            
        Case 1      '取消
            rsStaff.CancelUpdate
            If rsStaff.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                For Each oText In Me.txtStaff
                    Set oText.DataSource = rsStaff
                Next
                dcboArea.BoundText = txtStaff(7).Text
                
                control_status False
            Else
                Call cmdEdit_Click(5)
                MsgBox "目前並無任何員工資料!", , "金光益電器股份有限公司"
            End If
            Add_flag = 0

        Case 2      '新增
            Add_flag = 1
            Call Add_No(rsStaff, 2, 1)
            txtStaff(0).Text = Main_No
            control_status True
            dcboArea.Text = "--請選擇--"
            optSex(0).Value = True
            optMarry(0).Value = True
            
        Case 3      '修改
            control_status True
            
        Case 4      '刪除
            If MsgBox("確定要刪除此筆記錄??", 32 + vbYesNo, "金光益電器股份有限公司") = vbYes Then
                rsStaff.Delete
                If rsStaff.RecordCount <> 0 Then
                    Call cmdMove_Click(2)
                Else
                    Call cmdEdit_Click(5)
                    MsgBox "目前已無任何員工資料!", , "金光益電器股份有限公司"
                End If
            End If
            
        Case 5      '回主畫面
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            rsStaff.Close
            Set rsStaff = Nothing
            
    End Select
End Sub

'資料的移動
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsStaff)
    lblRecord.Caption = "員工資料：" & intRecord & "/" & intTotal
    
    '男生
    If rsStaff.Fields("性別") = 1 Then
        optSex(0).Value = True
    Else
        optSex(1).Value = True
    End If
    
    '已婚
    If rsStaff.Fields("婚姻狀況") = 1 Then
        optMarry(1).Value = True
    Else
        optMarry(0).Value = True
    End If
End Sub

Private Sub cmdRelation_Click()
    '離開前先記錄目前的筆數
    bm_move = intRecord
    frmRelation.Show 1
End Sub

'資料的連結
Private Sub Form_Load()
    '1.開啟員工資料表
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from 員工資料表"
    rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
    
    '2.將資料指定給感知元件
    For Each oText In Me.txtStaff
        Set oText.DataSource = rsStaff
    Next
        
    '3.製作DataCombo
    '製作地區名稱dcboArea
    Set rsArea = New ADODB.Recordset
    sql_rsArea = "select * from 地區表"
    rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
    
    Set dcboArea.DataSource = rsStaff
    dcboArea.DataField = "地區編號"
    Set dcboArea.RowSource = rsArea
    dcboArea.ListField = "地區名稱"
    dcboArea.BoundColumn = "地區編號"
    rsStaff.MoveFirst
    
    '4.表單載入時的狀態設定
    If rsStaff.RecordCount <> 0 Then
        Call cmdMove_Click(0)
        control_status False
    Else
        Call cmdEdit_Click(2)
    End If
    
    '5.表單的長寬設定
    frmStaff.Height = 5025
    frmStaff.Width = 9945
    
End Sub

'控制項的狀態設定
Private Sub control_status(control_flag As Boolean)
    '固定不變的狀態
    '員工編號,地區名稱,部門名稱,職稱皆不可操作
    txtStaff(0).Enabled = False
    txtArea.Enabled = False
    '地區編號,部門編號,職稱編號皆隱藏
    For i = 7 To 7
        txtStaff(i).Visible = False
    Next
    
    '狀態可切換
    For i = 1 To 6
        txtStaff(i).Enabled = control_flag
    Next
    fraSex.Enabled = control_flag
    fraMarry.Enabled = control_flag
        
    '地區名稱,部門名稱,職稱
    dcboArea.Visible = control_flag
    txtArea.Visible = Not control_flag
    
    '儲存,取消鈕
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next i
    
    '新增,修改,刪除,回主畫面鈕
    For i = 2 To 5
        cmdEdit(i).Visible = Not control_flag
    Next i
    
    '資料移動
    For i = 0 To 3
        cmdMove(i).Enabled = Not control_flag
    Next i
    '員工聯絡人
    cmdRelation.Visible = Not control_flag
    
End Sub

'將值填於txtArea txtDepart txtDuty中,並與rsStaff連動
Private Sub txtStaff_Change(Index As Integer)
    Select Case Index
        Case 7      '地區名稱欄位
            Set rsArea = New ADODB.Recordset
            sql_rsArea = "select * from 地區表"
            rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
            
            rsArea.Find "地區編號='" & txtStaff(7).Text & "'"
            If rsArea.EOF = False Then
                txtArea.Text = rsArea.Fields("地區名稱")
            End If
    End Select
End Sub
