VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00C0FFFF&
   Caption         =   "商品紀錄"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "商品紀錄.frx":0000
   ScaleHeight     =   10245
   ScaleWidth      =   15240
   WindowState     =   2  '最大化
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "篩選照片"
      Height          =   1095
      Left            =   12600
      Picture         =   "商品紀錄.frx":22D3
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "刪除照片"
      Height          =   1095
      Left            =   13560
      Picture         =   "商品紀錄.frx":33EF
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   4080
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "商品表"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   8040
      TabIndex        =   2
      Top             =   240
      Width           =   6855
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "增加照片"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   4440
         TabIndex        =   20
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "結束"
         Height          =   975
         Left            =   5520
         Picture         =   "商品紀錄.frx":3A65
         Style           =   1  '圖片外觀
         TabIndex        =   19
         Top             =   5880
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3720
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "最未筆"
         Height          =   975
         Left            =   3840
         Picture         =   "商品紀錄.frx":48B7
         Style           =   1  '圖片外觀
         TabIndex        =   18
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "下一筆"
         Height          =   975
         Left            =   2640
         Picture         =   "商品紀錄.frx":987E
         Style           =   1  '圖片外觀
         TabIndex        =   17
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "上一筆"
         Height          =   975
         Left            =   1440
         Picture         =   "商品紀錄.frx":A3D8
         Style           =   1  '圖片外觀
         TabIndex        =   16
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "第一筆"
         Height          =   975
         Left            =   240
         Picture         =   "商品紀錄.frx":F4DA
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   5880
         Width           =   975
      End
      Begin VB.TextBox txt 
         DataField       =   "商品編號"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Index           =   0
         Left            =   1965
         TabIndex        =   8
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txt 
         DataField       =   "商品名稱"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Index           =   1
         Left            =   1965
         TabIndex        =   7
         Top             =   855
         Width           =   3375
      End
      Begin VB.TextBox txt 
         DataField       =   "大分類"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Index           =   2
         Left            =   1965
         TabIndex        =   6
         Top             =   1245
         Width           =   3375
      End
      Begin VB.TextBox txt 
         DataField       =   "小分類"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Index           =   3
         Left            =   1965
         TabIndex        =   5
         Top             =   1620
         Width           =   3375
      End
      Begin VB.TextBox txt 
         DataField       =   "商品價格"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Index           =   4
         Left            =   1965
         TabIndex        =   4
         Top             =   1995
         Width           =   1320
      End
      Begin VB.TextBox txt 
         DataField       =   "商品照片"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Index           =   5
         Left            =   1920
         TabIndex        =   3
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   720
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "商品編號："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   525
         Width           =   975
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "商品名稱："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "大分類："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1290
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "小分類："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1665
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "商品價格："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "商品照片："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2430
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "商品紀錄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   7215
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "商品紀錄.frx":112B6
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   11245
         _Version        =   393216
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "Command2"
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
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "商品紀錄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Private Sub Command1_Click()
rs.MoveFirst
aa
End Sub

Private Sub Command13_Click()
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)
If CommonDialog1.FileName = "" Then
MsgBox "請找出一張照片來"
Else
SavePicture Image1.Picture, App.Path & "\" & txt(0) & ".jpg"
txt(5) = txt(0) & ".jpg"
End If
rs.MoveLast
rs.MovePrevious

End Sub

Private Sub Command14_Click()
Image1.Picture = LoadPicture("")
txt(5).Text = ""
End Sub

Private Sub Command2_Click()
rs.MovePrevious
If rs.BOF = True Then
   MsgBox "已經到第一筆了！"
   rs.MoveFirst
Else
aa
End If
End Sub

Private Sub Command3_Click()
rs.MoveNext
If rs.EOF = True Then
   MsgBox "已經到最末筆了！"
   rs.MoveLast
Else
aa
End If
End Sub

Private Sub Command4_Click()
rs.MoveLast
aa
End Sub

Private Sub Command7_Click()


End Sub

Private Sub Command5_Click()
Image1.Picture = LoadPicture("")
txt(5).Text = ""
End Sub

Private Sub Command6_Click()
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)
If CommonDialog1.FileName = "" Then
MsgBox "請找出一張照片來"
Else
SavePicture Image1.Picture, App.Path & "\" & txt(0) & ".jpg"
txt(5) = txt(0) & ".jpg"
End If
rs.MoveLast
rs.MovePrevious

End Sub

Private Sub Command8_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  MDIForm1.StatusBar1.Panels(4) = "商品管理系統"
Set rs = DataEnvironment1.rsCommand2

aa
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rs.Close
  
End Sub

Private Sub txt_Change(Index As Integer)
If txt(5) <> "" Then Image1.Picture = LoadPicture(App.Path & "\" & txt(5)) Else Image1.Picture = LoadPicture("")
End Sub

Private Sub aa()
   MDIForm1.StatusBar1.Panels(3).Text = "第" & rs.AbsolutePosition & "筆 / 共" & rs.RecordCount & "筆"
End Sub



