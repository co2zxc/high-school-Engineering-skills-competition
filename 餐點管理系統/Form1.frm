VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "會員紀錄"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10230
   ScaleWidth      =   15240
   WindowState     =   2  '最大化
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "編輯"
      Height          =   975
      Left            =   10560
      Picture         =   "Form1.frx":2007
      Style           =   1  '圖片外觀
      TabIndex        =   33
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "結束"
      Height          =   975
      Left            =   15360
      Picture         =   "Form1.frx":39A6
      Style           =   1  '圖片外觀
      TabIndex        =   31
      Top             =   7680
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5520
      Top             =   8400
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "新增"
      Height          =   975
      Left            =   11520
      Picture         =   "Form1.frx":47F8
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "刪除"
      Height          =   975
      Left            =   12480
      Picture         =   "Form1.frx":6BA6
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "儲存"
      Enabled         =   0   'False
      Height          =   975
      Left            =   13560
      Picture         =   "Form1.frx":8B39
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   7560
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "記錄內容"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   6960
      TabIndex        =   7
      Top             =   240
      Width           =   7455
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form1.frx":CA3C
         Height          =   6015
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   10610
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
         DataMember      =   "Command1"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "第一筆"
      Height          =   975
      Left            =   6480
      Picture         =   "Form1.frx":CA5B
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "上一筆"
      Height          =   975
      Left            =   7440
      Picture         =   "Form1.frx":E837
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "下一筆"
      Height          =   975
      Left            =   8400
      Picture         =   "Form1.frx":13939
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "最未筆"
      Height          =   975
      Left            =   9360
      Picture         =   "Form1.frx":14493
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   7560
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "會員資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6690
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   5895
      Begin VB.TextBox Text1 
         DataField       =   "照片"
         DataSource      =   "DataEnvironment1"
         Height          =   270
         Left            =   2040
         TabIndex        =   32
         Top             =   3480
         Width           =   3375
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "女"
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "男"
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txt客戶地址 
         DataField       =   "客戶地址"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txt客戶電話 
         DataField       =   "客戶電話"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   2040
         TabIndex        =   25
         Top             =   2715
         Width           =   3375
      End
      Begin VB.TextBox txt入員日期 
         DataField       =   "入員日期"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   2040
         TabIndex        =   23
         Top             =   2340
         Width           =   1320
      End
      Begin VB.TextBox txt客戶生日 
         DataField       =   "客戶生日"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   1920
         Width           =   1320
      End
      Begin VB.CheckBox chkVIP 
         BackColor       =   &H00C0FFFF&
         DataField       =   "VIP"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         Top             =   1575
         Width           =   330
      End
      Begin VB.TextBox txt客戶性別 
         DataField       =   "客戶性別"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   4320
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt客戶姓名 
         DataField       =   "客戶姓名"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txt客戶編號 
         DataField       =   "客戶編號"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   435
         Width           =   3375
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3600
         Top             =   4080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "篩選照片"
         Height          =   1095
         Left            =   3120
         Picture         =   "Form1.frx":1945A
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "刪除照片"
         Height          =   1095
         Left            =   4200
         Picture         =   "Form1.frx":1A576
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "照片："
         DataSource      =   "DataEnvironment1"
         Height          =   180
         Index           =   8
         Left            =   195
         TabIndex        =   28
         Top             =   3525
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶地址："
         DataSource      =   "DataEnvironment1"
         Height          =   180
         Index           =   7
         Left            =   195
         TabIndex        =   26
         Top             =   3135
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶電話："
         DataSource      =   "DataEnvironment1"
         Height          =   180
         Index           =   6
         Left            =   195
         TabIndex        =   24
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "入員日期："
         DataSource      =   "DataEnvironment1"
         Height          =   180
         Index           =   5
         Left            =   195
         TabIndex        =   22
         Top             =   2385
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶生日："
         DataSource      =   "DataEnvironment1"
         Height          =   180
         Index           =   4
         Left            =   195
         TabIndex        =   20
         Top             =   1995
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "VIP："
         DataSource      =   "DataEnvironment1"
         Height          =   180
         Index           =   3
         Left            =   195
         TabIndex        =   18
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶性別："
         DataSource      =   "DataEnvironment1"
         Height          =   180
         Index           =   2
         Left            =   195
         TabIndex        =   16
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶姓名："
         DataSource      =   "DataEnvironment1"
         Height          =   180
         Index           =   1
         Left            =   195
         TabIndex        =   14
         Top             =   855
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶編號："
         DataSource      =   "DataEnvironment1"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   12
         Top             =   480
         Width           =   900
      End
      Begin VB.Image Image1 
         Height          =   2295
         Left            =   480
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   2055
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   735
      Left            =   11280
      TabIndex        =   35
      Top             =   8880
      Visible         =   0   'False
      Width           =   3015
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5318
      _cy             =   1296
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "會員紀錄"
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
      Height          =   495
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
rs.MoveFirst
aa
End Sub

'Private Sub Command10_Click()
'rs.Update
'rs.MoveLast
'rs.MovePrevious
'aa
'Frame2.Enabled = True
'Frame1.Enabled = False
'Command1.Enabled = True
'Command2.Enabled = True
'Command3.Enabled = True
'Command4.Enabled = True
'Command5.Enabled = True
'Command6.Enabled = True
'Command7.Enabled = True
'End Sub


Private Sub Command11_Click()
DataGrid1.Enabled = True
End Sub

Private Sub Command12_Click()
Frame1.Enabled = True
Frame2.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = True
Command9.Enabled = False
DataGrid1.Enabled = False
End Sub

Private Sub Command13_Click()
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)
If CommonDialog1.FileName = "" Then
MsgBox "請找出一張照片來"
Else
SavePicture Image1.Picture, App.Path & "\" & txt客戶編號 & ".jpg"
Text1 = txt客戶編號 & ".jpg"
End If
rs.MoveLast
rs.MovePrevious
aa

'CommonDialog1.DialogTitle = "匯入圖片"
  'CommonDialog1.Filter = "(相片圖)*.jpg|*.jpg"
  'CommonDialog1.ShowOpen
  'Set Image1.Picture = LoadPicture(CommonDialog1.FileName)
  'SavePicture Image1.Picture, App.Path & "\image\" & txt客戶編號 & ".jpg"
   'Text1 = txt客戶編號 & ".jpg"
End Sub

Private Sub Command14_Click()
Image1.Picture = LoadPicture("")
Text1.Text = ""
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

Private Sub aa()
 MDIForm1.StatusBar1.Panels(3).Text = "第" & rs.AbsolutePosition & "筆 / 共" & rs.RecordCount & "筆"
End Sub

Private Sub Command5_Click()
DataEnvironment1.rsCommand1.AddNew
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = True
Command12.Enabled = False
DataGrid1.Enabled = False
aa
rs.AddNew
Frame1.Enabled = True
Frame2.Enabled = False
End Sub

Private Sub Command6_Click()
 C = MsgBox("是否刪除？", vbYesNo)
If C = vbYes Then
rs.Delete
DoEvents
End If
aa
End Sub

Private Sub Command7_Click()
 Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
 Command9.Enabled = True
 rs.Update
Frame2.Enabled = True
Frame1.Enabled = False
aa
End Sub

Private Sub Command9_Click()
 Unload Me
End Sub

Private Sub DataGrid1_Click()
 aa
End Sub

Private Sub Form_Load()
  MDIForm1.StatusBar1.Panels(5) = "精選好歌：MY LOVE"
  MDIForm1.StatusBar1.Panels(4) = "客戶管理系統"
Frame1.Enabled = False
Set rs = DataEnvironment1.rsCommand1
End Sub

Private Sub Option1_Click()
 txt客戶性別 = "男"
End Sub


Private Sub Option2_Click()
txt客戶性別 = "女"
End Sub



Private Sub Form_Unload(Cancel As Integer)
'  rs.Close
End Sub

Private Sub txt客戶編號_Change()
If txt客戶性別 = "男" Then Option1.Value = True Else Option1.Value = False
If txt客戶性別 = "女" Then Option2.Value = True Else Option2.Value = False
If Text1 <> "" Then Image1.Picture = LoadPicture(App.Path & "\" & Text1) Else Image1.Picture = LoadPicture("")
End Sub

Private Sub Text1_Change()
If Text1 <> "" Then Image1.Picture = LoadPicture(App.Path & "\" & Text1) Else Image1.Picture = LoadPicture("")
End Sub

Private Sub Form_Activate()
WindowsMediaPlayer1.Controls.play
WindowsMediaPlayer1.URL = "(鈴聲)Westlife - my love.mp3"
End Sub
