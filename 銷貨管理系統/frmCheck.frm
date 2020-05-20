VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCheck 
   Caption         =   "產品盤存表"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdEdit 
      Caption         =   "取     消"
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   31
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "儲     存"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   22
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "回主畫面"
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   21
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame fraComputer 
      Caption         =   "電腦盤點表"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   5895
      Begin VB.CommandButton cmdMove 
         Caption         =   "<<"
         Height          =   355
         Index           =   1
         Left            =   600
         TabIndex        =   29
         Top             =   4200
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "|<"
         Height          =   355
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   4200
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">|"
         Height          =   355
         Index           =   3
         Left            =   5160
         TabIndex        =   27
         Top             =   4200
         Width           =   350
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">>"
         Height          =   355
         Index           =   2
         Left            =   4800
         TabIndex        =   26
         Top             =   4200
         Width           =   350
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "取     消"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "儲     存"
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubEdit 
         Caption         =   "修     改"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtComputer 
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
         Left            =   3840
         TabIndex        =   20
         Top             =   1200
         Width           =   1600
      End
      Begin VB.TextBox txtComputer 
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
         Left            =   2160
         TabIndex        =   19
         Top             =   1200
         Width           =   1600
      End
      Begin VB.TextBox txtComputer 
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
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgComputer 
         Height          =   2415
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4260
         _Version        =   393216
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
         Left            =   960
         TabIndex        =   30
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label lblComputer 
         AutoSize        =   -1  'True
         Caption         =   "實際盤存量"
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
         Left            =   4080
         TabIndex        =   17
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label lblComputer 
         AutoSize        =   -1  'True
         Caption         =   "應有庫存量"
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
         Left            =   2400
         TabIndex        =   16
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label lblComputer 
         AutoSize        =   -1  'True
         Caption         =   "產品名稱"
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
         Left            =   840
         TabIndex        =   15
         Top             =   960
         Width           =   840
      End
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
      Index           =   1
      Left            =   1560
      TabIndex        =   12
      Top             =   1200
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
      Index           =   0
      Left            =   1560
      TabIndex        =   11
      Top             =   720
      Width           =   2500
   End
   Begin MSDataListLib.DataCombo dcboStaff 
      Height          =   360
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   720
      Width           =   2505
      _ExtentX        =   4419
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
   Begin VB.CommandButton cmdEdit 
      Caption         =   "修     改"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtCheck 
      Alignment       =   1  '靠右對齊
      DataField       =   "盤點日期"
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
      TabIndex        =   6
      Top             =   1680
      Width           =   2500
   End
   Begin VB.TextBox txtCheck 
      DataField       =   "盤點人員編號"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   2500
   End
   Begin VB.TextBox txtCheck 
      DataField       =   "盤點編號"
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
      Top             =   240
      Width           =   2500
   End
   Begin MSDataListLib.DataCombo dcboStaff 
      Height          =   360
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Width           =   2505
      _ExtentX        =   4419
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
   Begin VB.TextBox txtCheck 
      DataField       =   "稽核人員編號"
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   2500
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "稽核人員"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "盤點日期"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "盤點人員"
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
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "盤點編號"
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
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

'產品盤存表的編修
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '主檔修改
            control_status True
            fraComputer.Visible = True
            
        Case 1      '回主畫面
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            '若亦將實際盤點結果記錄下來,則可進行下期期初環境的建立
            Set rsComputer = New ADODB.Recordset
            sql_rsComputer = "select * from 電腦盤點表"
            rsComputer.Open sql_rsComputer, cn, adOpenDynamic, adLockOptimistic
            If rsComputer.Fields("實際盤存量") <> "" Then
                MsgBox ("請建立下期期初環境!!")
                frmMDIMain.mnuInitial.Enabled = True
            End If
            
        Case 2      '主檔儲存
            For i = 0 To 1
                rsStaff.MoveFirst
                rsStaff.Find "姓名='" & dcboStaff(i).Text & "'", , 1
                If rsStaff.EOF Then
                    MsgBox "請選擇" & lblCheck(i + 1).Caption & "", , "金光益電器股份有限公司"
                    dcboStaff(i).SetFocus
                    Exit Sub
                Else
                    txtCheck(i + 1).Text = dcboStaff(i).BoundText
                    rsCheck.Fields(i + 1) = txtCheck(i + 1).Text
                End If
            Next i
            
            If txtCheck(3).Text = "" Then
                MsgBox "請填寫盤點日期!", , "金光益電器股份有限公司"
                txtCheck(3).SetFocus
                Exit Sub
            Else
                If IsDate(txtCheck(3).Text) = False Then
                    MsgBox "您輸入的日期格式有誤,請再確認一遍!", , "金光益電器股份有限公司"
                    txtCheck(3).SetFocus
                    Exit Sub
                Else
                    rsCheck.Update
                    control_status False
                End If
            End If
                    
            '若是新增狀態,則將實際盤存量的值預設為應有庫存量的值
            If Add_flag = 1 Then
                '作為實際盤存量的預設值
                Do Until rsComputer.EOF
                    rsComputer.Fields("實際盤存量") = rsComputer.Fields("應有庫存量")
                    rsComputer.MoveNext
                Loop
                Call cmdMove_Click(0)
            End If
            
            Add_flag = 0

        Case 3      '取消
            '資料未新增完成,按下取消鈕則回到主畫面
            If Add_flag = 1 Then
                Call cmdEdit_Click(1)
            Else
                control_status False
            End If
            Add_flag = 0
    End Select
End Sub

'資料的移動
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsComputer)
    lblRecord.Caption = "盤點結果：" & intRecord & "/" & intTotal
    
    '藉由資料的移動,讓記錄可以顯示於文字方塊中
    Call dgComputer_Click
End Sub

'電腦盤點表的編修
Private Sub cmdSubEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '明細修改
            subcontrol_status True
            
        Case 1      '明細儲存
            rsComputer.Fields("實際盤存量") = txtComputer(2).Text
            '將盤點差異量寫入資料庫
            rsComputer.Fields("盤點差異量") = rsComputer.Fields("應有庫存量") - rsComputer.Fields("實際盤存量")
            rsComputer.Update
            
            subcontrol_status False
            
        Case 2      '明細取消
            subcontrol_status False
            Call rsComputer_Cnn
    End Select
End Sub

'讓DataGrid的記錄顯示於文字方塊中
Private Sub dgComputer_Click()
    For i = 0 To 2
        If rsComputer.Fields(i + 2) <> "" Then
            txtComputer(i).Text = rsComputer.Fields(i + 2)
        Else
            '若實際盤存量還未有值,則先以0代替
            txtComputer(i).Text = 0
        End If
    Next
End Sub

'產品盤存表的連結
Private Sub Form_Load()
    '1.開啟RecordSet
    Set rsCheck = New ADODB.Recordset
    sql_rsCheck = "select * from 產品盤存表"
    rsCheck.Open sql_rsCheck, cn, adOpenDynamic, adLockOptimistic
    
    '2.將資料指定給表單中的感知元件
    For Each oText In Me.txtCheck
        Set oText.DataSource = rsCheck
    Next
    
    '3.製作dcboStaff(編修資料時用之)
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from 員工資料表"
    rsStaff.Open sql_rsStaff, cn, adOpenKeyset, adLockReadOnly
    For i = 0 To 1
        Set dcboStaff(i).DataSource = rsCheck
        Set dcboStaff(i).RowSource = rsStaff
        dcboStaff(i).ListField = "姓名"
        dcboStaff(i).BoundColumn = "員工編號"
    Next
    dcboStaff(0).DataField = "盤點人員編號"
    dcboStaff(1).DataField = "稽核人員編號"
    
    '4.檢查產品盤存表的資料是否完整
    '不完整則進入新增狀態
    If txtCheck(1).Text = "" And txtCheck(2).Text = "" Then
        '新增的旗標
        Add_flag = 1
        For i = 0 To 1
            dcboStaff(i).Text = "--請選擇負責人員--"
        Next i
        control_status True
    Else
        '否則直接將資料呈現出來
        control_status False
        Call rsComputer_Cnn
    End If
    
    '5.表單的長寬設定
    frmCheck.Height = 7755
    frmCheck.Width = 6795
    
End Sub

'電腦盤點表的連結
Private Sub rsComputer_Cnn()
    '1.開啟電腦盤點表
    Set rsComputer = New ADODB.Recordset
    sql_rsComputer = "select * from 電腦盤點表 where 盤點編號='" & txtCheck(0).Text & "' order by 產品編號"
    rsComputer.Open sql_rsComputer, cn, adOpenDynamic, adLockOptimistic
          
    '2.將資料指定給DataGrid,並隱藏盤點編號與盤點差異量欄位
    Set dgComputer.DataSource = rsComputer
    For i = 0 To 1
        dgComputer.Columns.Item(i).Visible = False
    Next i
    dgComputer.Columns.Item(5).Visible = False
    dgComputer.Columns.Item(2).Width = 2000
    For i = 3 To 4
        dgComputer.Columns.Item(i).Alignment = dbgRight
    Next i
    
    '3.設定表單載入控制項的狀態
    subcontrol_status False
    
    '4.讓記錄顯示於文字方塊中
    Call cmdMove_Click(0)
End Sub

'資料的連動
Private Sub txtCheck_Change(Index As Integer)
    Select Case Index
        Case 0
            '讓產品盤存表與電腦盤點表的資料可以連動
            Call rsComputer_Cnn
            
        Case 1
            '顯示出盤點人員名稱
            Call txtStaff_Cnn(1)
            
        Case 2
            '顯示出稽核人員名稱
            Call txtStaff_Cnn(2)
    End Select
End Sub

'產品盤存表的感知元件狀態切換
Private Sub control_status(control_flag As Boolean)
    txtCheck(0).Enabled = False
    For i = 0 To 1
        txtStaff(i).Enabled = False
        txtStaff(i).Visible = Not control_flag
        dcboStaff(i).Visible = control_flag
    Next
    txtCheck(3).Enabled = control_flag
    
    '主檔修改,回主畫面鈕
    For i = 0 To 1
        cmdEdit(i).Visible = Not control_flag
    Next i
    
    '主檔儲存鈕,取消鈕
    For i = 2 To 3
        cmdEdit(i).Visible = control_flag
    Next i
    
    '框架fraComputer的狀態
    fraComputer.Visible = Not control_flag
    fraComputer.Enabled = Not control_flag
End Sub

'電腦盤點表的感知元件狀態切換
Private Sub subcontrol_status(subcontrol_flag As Boolean)
    For i = 0 To 1
        txtComputer(i).Enabled = False
    Next i
    txtComputer(2).Enabled = subcontrol_flag
    dgComputer.Enabled = Not subcontrol_flag
    
    '明細修改鈕
    cmdSubEdit(0).Visible = Not subcontrol_flag
    '明細儲存,取消鈕
    For i = 1 To 2
        cmdSubEdit(i).Visible = subcontrol_flag
    Next i
    
    For i = 0 To 3
        cmdEdit(i).Enabled = Not subcontrol_flag
        cmdMove(i).Enabled = Not subcontrol_flag
    Next i
End Sub

'txtStaff的顯示
Private Sub txtStaff_Cnn(dcbostaff_index As Integer)
    Set rsStaff = New ADODB.Recordset
    sql_rsStaff = "select * from 員工資料表"
    rsStaff.Open sql_rsStaff, cn, adOpenDynamic, adLockOptimistic
    
    rsStaff.MoveFirst
    rsStaff.Find "員工編號='" & txtCheck(dcbostaff_index).Text & "'", , adSearchForward, 1
    If rsStaff.EOF = False Then
        txtStaff(dcbostaff_index - 1).Text = rsStaff.Fields("姓名")
    End If
End Sub
