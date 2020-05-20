VERSION 5.00
Begin VB.Form frmMake 
   Caption         =   "出貨表"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6915
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdOK 
      Caption         =   "確     定"
      Height          =   375
      Left            =   5400
      TabIndex        =   33
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtMake 
      Alignment       =   1  '靠右對齊
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
      Index           =   9
      Left            =   4440
      TabIndex        =   32
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox txtMake 
      Alignment       =   1  '靠右對齊
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
      Left            =   4440
      TabIndex        =   31
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox txtMake 
      Alignment       =   1  '靠右對齊
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
      Left            =   4440
      TabIndex        =   30
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtMake 
      Alignment       =   1  '靠右對齊
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
      Left            =   4440
      TabIndex        =   29
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox txtMake 
      Alignment       =   1  '靠右對齊
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
      Left            =   4440
      TabIndex        =   28
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtMake 
      Alignment       =   1  '靠右對齊
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
      Left            =   4440
      TabIndex        =   27
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtMake 
      Alignment       =   1  '靠右對齊
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
      Left            =   4440
      TabIndex        =   26
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtMake 
      Alignment       =   1  '靠右對齊
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
      Left            =   4440
      TabIndex        =   25
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtMake 
      Alignment       =   1  '靠右對齊
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
      Left            =   4440
      TabIndex        =   24
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtMake 
      Alignment       =   1  '靠右對齊
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
      Left            =   4440
      TabIndex        =   23
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblNum 
      Alignment       =   1  '靠右對齊
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
      Index           =   9
      Left            =   2640
      TabIndex        =   22
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label lblNum 
      Alignment       =   1  '靠右對齊
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
      Index           =   8
      Left            =   2640
      TabIndex        =   21
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblNum 
      Alignment       =   1  '靠右對齊
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
      Index           =   7
      Left            =   2640
      TabIndex        =   20
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblNum 
      Alignment       =   1  '靠右對齊
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
      Index           =   6
      Left            =   2640
      TabIndex        =   19
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblNum 
      Alignment       =   1  '靠右對齊
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
      Index           =   5
      Left            =   2640
      TabIndex        =   18
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblNum 
      Alignment       =   1  '靠右對齊
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
      Index           =   4
      Left            =   2640
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblNum 
      Alignment       =   1  '靠右對齊
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
      Index           =   3
      Left            =   2640
      TabIndex        =   16
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblNum 
      Alignment       =   1  '靠右對齊
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
      Index           =   2
      Left            =   2640
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblNum 
      Alignment       =   1  '靠右對齊
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
      Index           =   1
      Left            =   2640
      TabIndex        =   14
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblNum 
      Alignment       =   1  '靠右對齊
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
      Index           =   0
      Left            =   2640
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblProdt 
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
      Index           =   9
      Left            =   600
      TabIndex        =   12
      Top             =   5400
      Width           =   1605
   End
   Begin VB.Label lblProdt 
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
      Index           =   8
      Left            =   600
      TabIndex        =   11
      Top             =   4920
      Width           =   1605
   End
   Begin VB.Label lblProdt 
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
      Index           =   7
      Left            =   600
      TabIndex        =   10
      Top             =   4440
      Width           =   1605
   End
   Begin VB.Label lblProdt 
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
      Index           =   6
      Left            =   600
      TabIndex        =   9
      Top             =   3960
      Width           =   1605
   End
   Begin VB.Label lblProdt 
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
      Index           =   5
      Left            =   600
      TabIndex        =   8
      Top             =   3480
      Width           =   1605
   End
   Begin VB.Label lblProdt 
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
      Index           =   4
      Left            =   600
      TabIndex        =   7
      Top             =   3000
      Width           =   1605
   End
   Begin VB.Label lblProdt 
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
      Index           =   3
      Left            =   600
      TabIndex        =   6
      Top             =   2520
      Width           =   1605
   End
   Begin VB.Label lblProdt 
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
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   1605
   End
   Begin VB.Label lblProdt 
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
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   1605
   End
   Begin VB.Label lblProdt 
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
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   1605
   End
   Begin VB.Label lblMake 
      AutoSize        =   -1  'True
      Caption         =   "預計出貨量"
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
      Left            =   4800
      TabIndex        =   2
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label lblMake 
      AutoSize        =   -1  'True
      Caption         =   "現存量"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblMake 
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
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   960
   End
End
Attribute VB_Name = "frmMake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'經過整理的訂單明細資料之暫存檔
Dim rsOr_detail_list As ADODB.Recordset

Private Sub cmdOK_Click()
    Set rsMake = New ADODB.Recordset
    rsMake.Open "select * from 生產核銷表 order by 產品編號", cn, adOpenDynamic, adLockOptimistic

    '1.將必須生產的產品數量記錄在生產核銷表中
    For i = 0 To rsOr_detail_list.RecordCount - 1
        If rsMake.RecordCount <> 0 Then
            rsMake.MoveFirst
        End If
        '2.生產核銷表內已有此項產品,即進行數量的加總即可
        rsMake.Find "產品名稱='" & lblProdt(i).Caption & "'", , 1
        If rsMake.EOF Then
            rsProdtStore.MoveFirst
            rsProdtStore.Find "產品名稱='" & lblProdt(i).Caption & "'", , 1
            rsMake.AddNew
            rsMake.Fields(0) = rsProdtStore.Fields(0)
            rsMake.Fields(1) = lblProdt(i).Caption
            rsMake.Fields(2) = Int(txtMake(i).Text)
        Else
            rsMake.Fields(2) = rsMake.Fields(2) + Int(txtMake(i).Text)
        End If
        rsMake.Update
               
        '3.按下確定鈕表示確定了要生產的數量,故訂單明細的生產核銷欄位即要更新為1
        rsOrder_Detail.Filter = "產品名稱='" & lblProdt(i).Caption & "' and 生產核銷=0"
        Do Until rsOrder_Detail.EOF
            rsOrder_Detail.Fields("生產核銷") = 1
            rsOrder_Detail.MoveNext
        Loop
    Next i
    
    '4.將最新的暫存量更新至庫存表中(暫存量=原來的暫存量-訂貨數量)
    rsProdtStore.MoveFirst
    Do Until rsProdtStore.EOF
        rsOr_detail_list.MoveFirst
        rsOr_detail_list.Find "產品編號='" & rsProdtStore.Fields(0) & "'", , 1
        If rsOr_detail_list.EOF = False Then
            rsProdtStore.Fields("暫存量") = rsProdtStore.Fields("暫存量") - rsOr_detail_list.Fields("數量")
        End If
        rsProdtStore.MoveNext
    Loop
    
    Unload Me
End Sub

Private Sub Form_Load()
    '1.先將表單中的控制項隱藏,再依據所需求顯示出所需的控制項
    For i = 0 To 9
        lblProdt(i).Visible = False
        lblNum(i).Visible = False
        txtMake(i).Visible = False
    Next i
        
    Set rsOrder_Detail = New ADODB.Recordset
    rsOrder_Detail.Open "select * from 訂單明細", cn, adOpenDynamic, adLockOptimistic
        
    '2.製作一個暫存訂單明細之產品數量的recordset(且還未通知生產的訂單明細)
    Set rsOr_detail_list = New ADODB.Recordset
    rsOr_detail_list.Fields.Append "產品編號", adChar, 3
    rsOr_detail_list.Fields.Append "產品名稱", adChar, 20
    rsOr_detail_list.Fields.Append "數量", adInteger
    rsOr_detail_list.Open
    
    '3.重整訂單明細中所訂的產品之數量(使其成為一筆筆不重複之資料)
    pre_prodt = ""
    rsOrder_Detail.Sort = "產品編號 asc"
    rsOrder_Detail.Filter = "生產核銷=0"
    Do Until rsOrder_Detail.EOF
        If pre_prodt <> rsOrder_Detail.Fields("產品編號") Then
            rsOr_detail_list.AddNew
            rsOr_detail_list.Fields(0) = rsOrder_Detail.Fields("產品編號")
            rsOr_detail_list.Fields(1) = Trim(rsOrder_Detail.Fields("產品名稱"))
            rsOr_detail_list.Fields(2) = rsOrder_Detail.Fields("訂貨數量")
            rsOr_detail_list.Update
        Else
            '重複則只需加總訂貨數量
            rsOr_detail_list.Fields(2) = rsOr_detail_list.Fields(2) + rsOrder_Detail.Fields("訂貨數量")
        End If
        pre_prodt = rsOrder_Detail.Fields("產品編號")
        
        rsOrder_Detail.MoveNext
    Loop
    rsOrder_Detail.Filter = adFilterNone

    Set rsProdtStore = New ADODB.Recordset
    sql_rsProdtStore = "select * from 產品庫存統計表"
    rsProdtStore.Open sql_rsProdtStore, cn, adOpenDynamic, adLockOptimistic
               
    '4.將此暫存檔遞增排序,並只顯示出與暫存檔同筆數的控制項
    rsOr_detail_list.Sort = "產品編號 asc"
    For i = 0 To rsOr_detail_list.RecordCount - 1
        lblProdt(i).Visible = True
        lblNum(i).Visible = True
        txtMake(i).Visible = True
        
        '將產品名稱放於lblProdt中
        lblProdt(i).Caption = RTrim(rsOr_detail_list.Fields("產品名稱"))
        rsProdtStore.MoveFirst
        rsProdtStore.Find "產品名稱='" & lblProdt(i).Caption & "'", , 1
        If rsProdtStore.Fields(0) = rsOr_detail_list.Fields(0) Then
            '呈現產品現存量
            lblNum(i).Caption = rsProdtStore.Fields("現存量")
         End If
         txtMake(i).Text = rsOr_detail_list.Fields("數量")
         txtMake(i).Enabled = False
        rsOr_detail_list.MoveNext
    Next i

End Sub


