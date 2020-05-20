VERSION 5.00
Begin VB.Form frmInfo 
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6120
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   355
      Index           =   1
      Left            =   840
      TabIndex        =   13
      Top             =   2520
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   355
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   2520
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   355
      Index           =   3
      Left            =   5160
      TabIndex        =   11
      Top             =   2520
      Width           =   350
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">>"
      Height          =   355
      Index           =   2
      Left            =   4800
      TabIndex        =   10
      Top             =   2520
      Width           =   350
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "回主畫面"
      Height          =   375
      Index           =   5
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "取     消"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "儲     存"
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "刪     除"
      Height          =   375
      Index           =   4
      Left            =   4440
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "新     增"
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "修     改"
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
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
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
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
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   1815
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
      TabIndex        =   14
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   75
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   75
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsInfo As ADODB.Recordset       '共用的RecordSet物件
Dim sql_Area As String              '代替開啟地區表的字串變數
Dim sql_Depart As String            '代替開啟部門分類表的字串變數
Dim sql_Duty As String              '代替開啟職稱等級表的字串變數
Dim strData_show As String          '顯示於lblRecord的字樣

'資料的編修
Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
        Case 0      '儲存
            If txtFields(1).Text <> "" Then
                rsInfo.Update
            Else
                MsgBox "請填寫" & lblLabel(1).Caption & "!", , "金光益電器股份有限公司"
                txtFields(1).SetFocus
                Exit Sub
            End If
            Call cmdMove_Click(3)
            control_status False

        Case 1      '取消
            rsInfo.CancelUpdate
            If rsInfo.RecordCount <> 0 Then
                If Add_flag = 1 Then
                    Call cmdMove_Click(3)
                End If
                For Each oText In Me.txtFields
                    Set oText.DataSource = rsInfo
                Next
                
                control_status False
            Else
                Call cmdEdit_Click(5)
                MsgBox "目前並無任何" & strData_show & "!", , "金光益電器股份有限公司"
            End If

        Case 2      '新增
            Add_flag = 1
            
            Select Case Menu_index
                Case 1
                    Call Add_No(rsInfo, 6, 2)
                Case 2
                    Call Add_No(rsInfo, 7, 2)
                Case 3
                    Call Add_No(rsInfo, 8, 2)
            End Select
            
            control_status True
            txtFields(0).Text = Main_No
            
        Case 3      '修改
            control_status True
            
        Case 4      '刪除
            If MsgBox("確定要刪除此筆記錄??", 32 + vbYesNo, "金光益電器股份有限公司") = vbYes Then
                rsInfo.Delete
                If rsInfo.RecordCount <> 0 Then
                    Call cmdMove_Click(2)
                Else
                    Call cmdEdit_Click(5)
                    MsgBox "目前並無任何" & strData_show & "!", , "金光益電器股份有限公司"
                End If
            End If
            
        Case 5      '離開
            Unload Me
            frmMDIMain.mnuExit.Enabled = True
            '將主表單中的選擇項目復原
            frmMDIMain.mnuArea.Enabled = True

    End Select
End Sub

'資料的移動
Private Sub cmdMove_Click(Index As Integer)
    Call rs_Move(Index, rsInfo)
    lblRecord.Caption = strData_show & "：" & intRecord & "/" & intTotal
End Sub

'資料的連結
Private Sub Form_Load()
    '1.以Menu_index做區別,來開啟不同RecordSet
    Set rsInfo = New ADODB.Recordset
    Select Case Menu_index
        Case 1      '部門分類表
            sql_Depart = "select * from 部門分類表 ORDER BY 部門編號"
            rsInfo.Open sql_Depart, cn, adOpenDynamic, adLockOptimistic
            
            txtFields(0).DataField = "部門編號"
            txtFields(1).DataField = "部門名稱"
            frmInfo.Caption = "部門分類表"
            lblLabel(0).Caption = "部門編號"
            lblLabel(1).Caption = "部門名稱"
            strData_show = "部門資料"
            
        Case 2      '職稱等級表
            sql_Duty = "select * from 職稱等級表 ORDER BY 職稱編號"
            rsInfo.Open sql_Duty, cn, adOpenDynamic, adLockOptimistic
            
            txtFields(0).DataField = "職稱編號"
            txtFields(1).DataField = "職稱"
            frmInfo.Caption = "職稱等級表"
            lblLabel(0).Caption = "職稱編號"
            lblLabel(1).Caption = "職稱"
            strData_show = "職稱資料"
            
        Case 3      '地區表
            sql_Area = "select * from 地區表 ORDER BY 地區編號"
            rsInfo.Open sql_Area, cn, adOpenDynamic, adLockOptimistic
            
            txtFields(0).DataField = "地區編號"
            txtFields(1).DataField = "地區名稱"
            frmInfo.Caption = "地區表"
            lblLabel(0).Caption = "地區編號"
            lblLabel(1).Caption = "地區名稱"
            strData_show = "地區資料"
            
    End Select
    
    '2.將資料指定到感知元件中
    For Each oText In Me.txtFields
        Set oText.DataSource = rsInfo
    Next
    
    '3.檢查是否有資料
    If rsInfo.RecordCount <> 0 Then
        '表單載入時的狀態設定
        control_status False
        Call cmdMove_Click(0)
    Else
        Call cmdEdit_Click(2)
    End If
  
    '4.表單的長寬設定
    frmInfo.Height = 3825
    frmInfo.Width = 6240
End Sub

'控制項的狀態設定
Private Sub control_status(control_flag As Boolean)
    txtFields(0).Enabled = False
    txtFields(1).Enabled = control_flag
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
    
End Sub

