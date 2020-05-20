VERSION 5.00
Begin VB.Form Password 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "金光益進銷存系統"
   ClientHeight    =   5475
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3234.812
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   7816.724
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command1 
      Caption         =   "清除"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "離開"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2400
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   465
      IMEMode         =   3  '暫止
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      Caption         =   "歡迎使用"
      Height          =   1935
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   5415
      Begin VB.TextBox txtuser 
         Height          =   495
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label 
         Alignment       =   2  '置中對齊
         Caption         =   "請輸入系統帳號"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  '置中對齊
         Caption         =   "請輸入系統密碼"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Login_Time As Integer '記錄密碼輸入次數
Public LoginSucceeded As Boolean
Dim mnuMkCk As String
Dim mnuInfo As String
Private Sub cmdCancel_Click()
   
   End
   
End Sub


Private Sub cmdOK_Click()
    '檢查密碼的正確性
    If txtuser = "qwe" And txtPassword = "123" Then
        '將傳遞成功訊息至呼叫本副程式之
        '副程式的程式碼置於此處
        '設定全域變數是最簡單的方式
         frmMDIMain.Show
            Unload Me
    ElseIf txtuser = "asd" And txtPassword = "456" Then
        
         frmMDIMain.Show
         frmMDIMain.mnuMkCk.Enabled = False
         Unload Me
    ElseIf txtuser = "zxc" And txtPassword = "789" Then
        
         frmMDIMain.Show
         frmMDIMain.mnuInfo.Enabled = False
         frmMDIMain.mnuMkCk.Enabled = False
         Unload Me
    Else
    
            Login_Time = Login_Time + 1
            
            MsgBox "密碼錯誤,尚有 " & 3 - Login_Time & " 次機會", 48, "登入訊息"
            
         
         If Login_Time = 3 Then
         
         End
            
        
        End If
      End If
        ' MsgBox "密碼錯誤，請重新輸入!", , "登入"
       ' txtPassword.SetFocus
       ' SendKeys "{Home}+{End}"
         
        

End Sub

Private Sub Command1_Click()
txtuser = ""
txtPassword = ""
End Sub

