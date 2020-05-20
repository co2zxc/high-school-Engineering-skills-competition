VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "登入"
   ClientHeight    =   4095
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "登入.frx":0000
   ScaleHeight     =   2419.461
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   5633.674
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "登入.frx":BFC2
      Height          =   360
      Left            =   2880
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "員工編號"
      BoundColumn     =   "員工編號"
      Text            =   "選擇員工編號- - -"
      Object.DataMember      =   "員工"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "華康墨字體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   390
      Left            =   2160
      TabIndex        =   3
      Top             =   3480
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  '暫止
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2880
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "使用者名稱(&U):"
      BeginProperty Font 
         Name            =   "華康墨字體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "密碼(&P):"
      BeginProperty Font 
         Name            =   "華康墨字體(P)"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   2760
      Width           =   2160
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public user As String
Public ID As String



Private Sub cmdOK_Click()
  Do While Not DataEnvironment1.rs員工.EOF
      If DataEnvironment1.rs員工("員工編號") = DataCombo1 Then
         ID = DataCombo1
         user = DataEnvironment1.rs員工("員工姓名")
         Password = DataEnvironment1.rs員工("密碼")
         Exit Do
      End If
      DataEnvironment1.rs員工.MoveNext
    Loop
    
    If txtPassword = Password Then
        Unload Me
        MDIForm1.Show
    Else
        MsgBox "密碼錯誤，請重新輸入!", , "登入"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

