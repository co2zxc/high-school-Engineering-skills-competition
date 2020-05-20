VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelation 
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   8115
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   6120
      TabIndex        =   21
      Top             =   3240
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   6480
      TabIndex        =   20
      Top             =   3240
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   600
      TabIndex        =   19
      Top             =   3240
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   960
      TabIndex        =   18
      Top             =   3240
      Width           =   350
   End
   Begin VB.TextBox txtArea 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "離     開"
      Height          =   375
      Index           =   5
      Left            =   6240
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "取     消"
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "儲     存"
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "刪     除"
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "新     增"
      Height          =   375
      Index           =   2
      Left            =   6240
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "修     改"
      Height          =   375
      Index           =   3
      Left            =   6240
      TabIndex        =   14
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtRelation 
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
      Index           =   4
      Left            =   1920
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtRelation 
      DataField       =   "聯絡人電話"
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
      TabIndex        =   2
      Top             =   1800
      Width           =   2500
   End
   Begin VB.TextBox txtRelation 
      DataField       =   "關係"
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
      TabIndex        =   1
      Top             =   1320
      Width           =   2500
   End
   Begin MSDataListLib.DataCombo dcboArea 
      Height          =   360
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtRelation 
      DataField       =   "聯絡人住址"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txtRelation 
      DataField       =   "聯絡人姓名"
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
      TabIndex        =   0
      Top             =   840
      Width           =   2500
   End
   Begin VB.Label lblStaff 
      BorderStyle     =   1  '單線固定
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
      TabIndex        =   23
      Top             =   360
      Width           =   2535
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
      TabIndex        =   22
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      Caption         =   "關係"
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
      Left            =   1320
      TabIndex        =   12
      Top             =   1395
      Width           =   495
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人住址"
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
      TabIndex        =   11
      Top             =   2355
      Width           =   1215
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人電話"
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
      TabIndex        =   10
      Top             =   1875
      Width           =   1215
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人姓名"
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
      TabIndex        =   9
      Top             =   915
      Width           =   1215
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      Caption         =   "員工姓名"
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
      Left            =   840
      TabIndex        =   8
      Top             =   435
      Width           =   975
   End
End
Attribute VB_Name = "frmRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'資料的編修
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '儲存
            For i = 0 To 3
                If txtRelation(i).Text = "" Then
                    MsgBox "請填寫" & lblRelation(i + 1).Caption & "!", , "金光益電器股份有限公司"
                    txtRelation(i).SetFocus
                    Exit Sub
                Else
                    rsArea.MoveFirst
                    rsArea.Find "地區名稱='" & dcboArea.Text & "'", , 1
                    If rsArea.EOF Then
                        MsgBox "請選擇地區!", , "金光益電器股份有限公司"
                        dcboArea.SetFocus
                        Exit Sub
                    Else
                        rsStaff.MoveFirst
                        rsStaff.Find "姓名='" & lblStaff.Caption & "'", , 1
                        rsRelation.Fields("員工編號") = rsStaff.Fields(0)
                        txtRelation(4).Text = dcboArea.BoundText
                        rsRelation.Update
                    End If
                End If
            Next i
            Call cmdMove_Click(3)
            control_status False

        Case 1      '取消
            rsRelation.CancelUpdate
            If rsRelation.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                For Each oText In Me.txtRelation
                    Set oText.DataSource = rsRelation
                Next
                dcboArea.BoundText = txtRelation(4).Text
                control_status False
            Else
                Call cmdEdit_Click(5)
                MsgBox "目前沒有員工" & frmStaff.txtStaff(1).Text & "的任何聯絡人之資料!", , "金光益電器股份有限公司"
            End If

        Case 2      '新增
            Add_flag = 1
            rsRelation.AddNew
            control_status True
            dcboArea.Text = "請選擇"
            
        Case 3      '修改
            control_status True
            dcboArea.Text = txtArea.Text
            
        Case 4      '刪除
            If MsgBox("確定要刪除此筆記錄??", 32 + vbYesNo, "金光益電器股份有限公司") = vbYes Then
                rsRelation.Delete
                If rsRelation.RecordCount <> 0 Then
                    rsRelation.MoveNext
                Else
                    Call cmdEdit_Click(5)
                    MsgBox "目前沒有員工" & frmStaff.txtStaff(1).Text & "的任何聯絡人之資料!", , "金光益電器股份有限公司"
                End If
            End If
            
        Case 5      '離開
            '要回到frmStaff前,將frmStaff目前的筆數放回去
            intRecord = bm_move
            Unload Me
            rsRelation.Close
            Set rsRelation = Nothing
    End Select
End Sub

'資料的移動
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsRelation)
    lblRecord.Caption = "聯絡人資料：" & intRecord & "/" & intTotal
    
End Sub

'資料的連結
Private Sub Form_Load()
    '1.開啟員工聯絡人資料表
    Set rsRelation = New ADODB.Recordset
    sql_rsRelation = "select * from 員工聯絡人資料表 where 員工編號='" & frmStaff.txtStaff(0).Text & "'"
    rsRelation.Open sql_rsRelation, cn, adOpenDynamic, adLockOptimistic
    
    '2.將資料指定至表單中的感知元件上
    For Each oText In Me.txtRelation
        Set oText.DataSource = rsRelation
    Next
    
    '員工姓名來源
    lblStaff.Caption = frmStaff.txtStaff(1).Text
    Me.Caption = "員工 " & lblStaff.Caption & " 的聯絡人相關資料"
    
    '3.檢查是否有資料
    '若無資料就進入新增狀態
    If rsRelation.RecordCount = 0 Then
        Call cmdEdit_Click(2)
    Else
        '設定表單載入時控制項的狀態
        control_status False
        Call cmdMove_Click(0)
    End If
       
    '4.製作DataCombo
    '製作dcboArea
    Set rsArea = New ADODB.Recordset
    sql_rsArea = "select * from 地區表"
    rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
    
    Set dcboArea.DataSource = rsRelation
    dcboArea.DataField = "地區編號"
    Set dcboArea.RowSource = rsArea
    dcboArea.ListField = "地區名稱"
    dcboArea.BoundColumn = "地區編號"
        
End Sub

'控制項的狀態設定
Private Sub control_status(control_flag As Boolean)
    '固定不變的狀態
    txtArea.Enabled = False
    txtRelation(4).Visible = False
    
    '狀態可切換
    For i = 0 To 4
        txtRelation(i).Enabled = control_flag
    Next
    dcboArea.Visible = control_flag
    txtArea.Visible = Not control_flag
    
    '儲存,取消鈕
    For i = 0 To 1
        cmdEdit(i).Visible = control_flag
    Next i
    '新增,修改,刪除,離開鈕
    For i = 2 To 5
        cmdEdit(i).Visible = Not control_flag
    Next i
    
    '移動鈕
    For i = 0 To 3
        cmdMove(i).Enabled = Not control_flag
    Next i
End Sub

'讓名稱顯示之欄位txtArea能與資料連動
Private Sub txtRelation_Change(Index As Integer)
    Select Case Index
        Case 4      '地區名稱欄位
            Set rsArea = New ADODB.Recordset
            sql_rsArea = "select * from 地區表"
            rsArea.Open sql_rsArea, cn, adOpenKeyset, adLockReadOnly
            
            rsArea.Find "地區編號='" & txtRelation(4).Text & "'"
            If rsArea.EOF = False Then
                txtArea.Text = rsArea.Fields("地區名稱")
                dcboArea.Text = rsArea.Fields("地區名稱")
            End If
    End Select
End Sub
