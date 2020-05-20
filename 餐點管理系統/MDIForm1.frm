VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "MDIForm1"
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10245
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   8040
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2548
            MinWidth        =   2548
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5274
            MinWidth        =   5274
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2542
            MinWidth        =   2542
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3529
            MinWidth        =   3529
         EndProperty
      EndProperty
   End
   Begin VB.Menu F 
      Caption         =   "系統介紹 "
      Begin VB.Menu F1 
         Caption         =   "系統動機 "
      End
      Begin VB.Menu F2 
         Caption         =   "系統目的 "
      End
   End
   Begin VB.Menu A 
      Caption         =   "客戶資料"
      Begin VB.Menu a1 
         Caption         =   "管理客戶 "
      End
      Begin VB.Menu A2 
         Caption         =   "查詢客戶資料"
      End
   End
   Begin VB.Menu B 
      Caption         =   "商品資料管理 "
   End
   Begin VB.Menu C 
      Caption         =   "銷售資料"
   End
   Begin VB.Menu H 
      Caption         =   "活動訊息 "
      Begin VB.Menu H1 
         Caption         =   "新產品的上市和特價的優惠 "
      End
      Begin VB.Menu H2 
         Caption         =   "銷售前十名 "
      End
   End
   Begin VB.Menu G 
      Caption         =   "組員資料"
      Begin VB.Menu G1 
         Caption         =   "尤廷紜"
      End
      Begin VB.Menu G2 
         Caption         =   "伍虹蓉"
      End
   End
   Begin VB.Menu E 
      Caption         =   "結束 "
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a1_Click()
 If Not (ActiveForm Is Nothing) Then Unload Me
  Form1.Show
End Sub

Private Sub A2_Click()
 If Not (ActiveForm Is Nothing) Then Unload Me
  Form4.Show
End Sub

Private Sub B_Click()
If frmLogin.user = "" Then
MsgBox "使年權線不足，不可使入本系統", 0 + 64, "錯誤訊息"
Else
 If Not (ActiveForm Is Nothing) Then Unload Me
  Form2.Show
End If


End Sub

Private Sub C_Click()

If Not (ActiveForm Is Nothing) Then Unload Me
  Form3.Show
End Sub

Private Sub D_Click()
'If frmLogin.user = "" Then
'MsgBox "使年權線不足，不可使入本系統", 0 + 64, "錯誤訊息"
'Else
If Not (ActiveForm Is Nothing) Then Unload Me
  Form5.Show
'End If


End Sub

Private Sub E_Click()
End
End Sub

Private Sub F1_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form6.Show
End Sub

Private Sub F2_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form7.Show
End Sub

Private Sub G1_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form8.Show
End Sub

Private Sub G2_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form9.Show
End Sub

Private Sub H1_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form11.Show
End Sub

Private Sub H2_Click()
If Not (ActiveForm Is Nothing) Then Unload Me
  Form10.Show
End Sub

Private Sub MDIForm_Load()

If frmLogin.user = "伍虹蓉" Then
  MDIForm1.StatusBar1.Panels(4) = "職位：總裁"
  Else
End If

  If frmLogin.user = "尤廷紜" Then
  MDIForm1.StatusBar1.Panels(4) = "職位：店長"
  Else
End If

If frmLogin.user = "簡明富" Then
  MDIForm1.StatusBar1.Panels(4) = "職位：清潔工"
  Else
End If

 StatusBar1.Panels(1) = "使用者：" & frmLogin.user
End Sub

Private Sub Timer1_Timer()
  StatusBar1.Panels(2) = "現在時間：" & Now()
End Sub
