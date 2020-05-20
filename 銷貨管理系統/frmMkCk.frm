VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMkCk 
   Caption         =   "生產核銷表"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8400
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdReturn 
      Caption         =   "回主畫面"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "核     銷"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraDetail 
      Caption         =   "未銷貨的產品資料"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   600
      TabIndex        =   4
      Top             =   3600
      Width           =   4935
      Begin MSDataGridLib.DataGrid dgDetail 
         Height          =   2175
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3836
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
   End
   Begin VB.ListBox lstNot 
      Height          =   2370
      Left            =   3240
      Style           =   1  '項目包含核取方塊
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ListBox lstOK 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblInCheck 
      AutoSize        =   -1  'True
      Caption         =   "未取貨之資料："
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
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   1680
   End
   Begin VB.Label lblInCheck 
      AutoSize        =   -1  'True
      Caption         =   "未出貨之資料："
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
      TabIndex        =   0
      Top             =   840
      Width           =   1680
   End
End
Attribute VB_Name = "frmMkCk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'核銷生產核銷欄位(將生產量更新至庫存表中,並將生產核銷表內的記錄刪除)
Private Sub cmdOK_Click()
    Set rsProdtStore = New ADODB.Recordset
    rsProdtStore.Open "select * from 產品庫存統計表 order by 產品編號", cn, adOpenDynamic, adLockOptimistic
    
    Set rsMake = New ADODB.Recordset
    rsMake.Open "select * from 生產核銷表 order by 產品編號", cn, adOpenDynamic, adLockOptimistic

    For i = 0 To lstNot.ListCount - 1
        If lstNot.Selected(i) Then
            '1.找出生產核銷表中與清單中被勾選之相同產品名稱
            rsMake.MoveFirst
            rsMake.Find "產品名稱='" & RTrim(lstNot.List(i)) & "'", , 1
            '2.找出庫存表中與此產品相對應之產品編號
            rsProdtStore.MoveFirst
            rsProdtStore.Find "產品編號='" & rsMake.Fields(0) & "'", , 1
            If rsProdtStore.EOF = False Then
                '3.更新其生產量與現存量,暫存量
                rsProdtStore.Fields("生產量") = rsProdtStore.Fields("生產量") + rsMake.Fields("預計生產量")
                rsProdtStore.Fields("現存量") = rsProdtStore.Fields("上期庫存量") + rsProdtStore.Fields("生產量") - rsProdtStore.Fields("銷貨量")
                '產品確定已生產後,暫存量即等於現存量
                rsProdtStore.Fields("暫存量") = rsProdtStore.Fields("現存量")
                rsProdtStore.Update
                '4.更新訂單明細中的生產核銷欄位為2,表示產品已生產,可進行銷貨了
                rsOrder_Detail.Filter = "產品編號='" & rsMake.Fields(0) & "' and 生產核銷=1"
                Do Until rsOrder_Detail.EOF
                    rsOrder_Detail.Fields("生產核銷") = 2
                    
                    rsOrder_Detail.MoveNext
                Loop
                
                '5.刪除生產核銷表中相對應之記錄
                rsMake.Delete
                If rsMake.RecordCount <> 0 Then
                    rsMake.MoveNext
                End If
                rsOrder_Detail.Filter = adFilterNone
            End If
        End If
    Next i
    
    '6.重整顯示於lstOK與lstNot中的資料
    Call Form_Load
End Sub

'回主畫面
Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmdOK.Enabled = False
    
    Set rsOrder_Detail = New ADODB.Recordset
    rsOrder_Detail.Open "select * from 訂單明細 where 銷貨核銷=0 order by 產品編號", cn, adOpenDynamic, adLockOptimistic
     
    lstOK.Clear
    lstNot.Clear
    pre_prodt = ""
    '已完成生產,且已核銷過的資料則呈現於lstOK中
    rsOrder_Detail.Filter = "生產核銷=2"
    Do Until rsOrder_Detail.EOF
        '若產品有重複,則只要以一項代表即可,不須重複呈現
        If pre_prodt <> rsOrder_Detail.Fields("產品編號") Then
            lstOK.AddItem rsOrder_Detail.Fields("產品名稱")
        End If
        pre_prodt = rsOrder_Detail.Fields("產品編號")
        rsOrder_Detail.MoveNext
    Loop
    rsOrder_Detail.Filter = adFilterNone
    
    '只產生生產單,還在生產階段的資料
    pre_prodt = ""
    rsOrder_Detail.Filter = "生產核銷=1"
    Do Until rsOrder_Detail.EOF
        If pre_prodt <> rsOrder_Detail.Fields("產品編號") Then
            lstNot.AddItem rsOrder_Detail.Fields("產品名稱")
        End If
        pre_prodt = rsOrder_Detail.Fields("產品編號")
        rsOrder_Detail.MoveNext
    Loop
    rsOrder_Detail.Filter = adFilterNone
       
    rsOrder_Detail.Sort = "訂單編號 asc"
    Set dgDetail.DataSource = rsOrder_Detail
    
    For i = 1 To 2
        dgDetail.Columns.Item(i).Visible = False
    Next i
    
    For i = 7 To 8
        dgDetail.Columns.Item(i).Visible = False
    Next i
        
    For i = 4 To 6
        dgDetail.Columns.Item(i).Visible = False
    Next i
    dgDetail.Columns.Item(5).Visible = True
    dgDetail.Columns.Item(5).Alignment = dbgRight
    dgDetail.Columns.Item(3).Width = 1900
        
End Sub

''已生產之資料'被勾選後,核銷鈕才使用
Private Sub lstNot_Click()
    cmdOK.Enabled = True
End Sub
