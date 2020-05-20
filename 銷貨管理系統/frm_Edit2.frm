VERSION 5.00
Begin VB.Form frmOrd_Detail2 
   Caption         =   "品牌新增"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5610
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdLast 
      Caption         =   "最末筆"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   1800
      Width           =   800
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一筆"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   840
      Width           =   800
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一筆"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   1320
      Width           =   800
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "第一筆"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   360
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "回主畫面"
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Top             =   1920
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "儲存"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   1920
      Width           =   915
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "刪除"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   915
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "修改"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   915
   End
   Begin VB.TextBox txtDetail 
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
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   2500
   End
   Begin VB.TextBox txtDetail 
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
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   2500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "廠商"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "廠商編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frmOrd_Detail2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As Connection
Dim rs As Recordset
Dim i As Integer

Private Sub Form_Load()
    Set cn = New Connection
    cn.Open "dsn=dbRose"
    Set rs = New Recordset
    rs.Open "select 廠商編號,廠商 " _
    & "from 廠商資料", cn, adOpenStatic, adLockOptimistic
    
    Dim oText As TextBox
    For Each oText In Me.txtDetail
        Set oText.DataSource = rs
    Next
        txtDetail(0).DataField = "廠商編號"
        txtDetail(1).DataField = "廠商"
        
        '表單載入時,文字方塊為不可修改
        For i = 0 To 1
            txtDetail(i).Enabled = False
        Next
        
End Sub

Private Sub cmdAdd_Click()
    rs.AddNew
    '文字方塊變為可修改
    For i = 0 To 1
        txtDetail(i).Enabled = True
    Next
        txtDetail(0).SetFocus
End Sub

Private Sub cmdCancel_Click()
    rs.CancelUpdate
    '文字方塊變為不可修改
    For i = 0 To 1
        txtDetail(i).Enabled = False
    Next
End Sub

Private Sub cmdDel_Click()
    If MsgBox("確定要刪除此筆資料嗎？", vbYesNo + 48, "金光益電器股份有限公司") = vbYes Then
        rs.Delete
        '呼叫MoveNext方法來更新資料
        rs.MoveNext
        If rs.EOF Then rs.MoveLast
    End If
End Sub

Private Sub cmdEdit_Click()
    '文字方塊變為可修改
    For i = 0 To 1
        txtDetail(i).Enabled = True
    Next
End Sub

Private Sub cmdExit_Click()

     Unload Me
    frmMDIMain.Show
End Sub

Private Sub cmdFirst_Click()
    rs.MoveFirst
End Sub

Private Sub cmdLast_Click()
    rs.MoveLast
End Sub

Private Sub cmdNext_Click()
    rs.MoveNext
    If rs.EOF Then
        rs.MoveLast
        MsgBox ("這已是最後一筆資料!!")
    End If
End Sub

Private Sub cmdPrevious_Click()
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveFirst
        MsgBox ("這已是第一筆資料!!")
        
    End If
End Sub

Private Sub cmdUpdate_Click()
    rs.Update
    '文字方塊變為不可修改
    For i = 0 To 1
        txtDetail(i).Enabled = False
    Next
End Sub

