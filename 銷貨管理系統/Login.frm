VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form Login 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "登入"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton 取消 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   2370
      TabIndex        =   5
      Top             =   1605
      Width           =   1215
   End
   Begin VB.TextBox pwd 
      Height          =   285
      IMEMode         =   3  '暫止
      Left            =   2175
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   930
      Width           =   1440
   End
   Begin VB.CommandButton 確定 
      Caption         =   "確定(&S)"
      Height          =   360
      Left            =   735
      TabIndex        =   1
      Top             =   1605
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo AllEmp 
      Height          =   330
      Left            =   2130
      TabIndex        =   0
      Top             =   330
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "指定姓名："
      Height          =   330
      Left            =   1005
      TabIndex        =   4
      Top             =   390
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "輸入密碼："
      Height          =   300
      Left            =   1020
      TabIndex        =   3
      Top             =   990
      Width           =   960
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AllEmp_Change()
pwd.SetFocus
End Sub

Private Sub Form_Load()
AllEmp.ListField = "姓名"
Set AllEmp.RowSource = getAllemp()
End Sub

Private Sub 取消_Click()
If MsgBox("確定結束本系統?", 4) = 6 Then
    End
End If
End Sub

Private Sub 確定_Click()
Set NowUser = AllEmp.RowSource
NowUser.MoveFirst
NowUser.Find "姓名='" & AllEmp.Text & "'"
If NowUser.EOF = True Then
    MsgBox "未找到員工，請重新操作"
    Exit Sub
End If
c = NowUser("密碼")
If Len(c) <> Len(pwd.Text) Then
    MsgBox "密碼輸入錯誤"
    Exit Sub
End If
For i = 1 To Len(c)
    If Asc(Mid(c, i, 1)) <> Asc(Mid(pwd.Text, i, 1)) Then
        MsgBox "密碼輸入錯誤"
        Exit Sub
    End If
Next
Form1.Show
Form1.NowUserName = AllEmp.Text
If NowUser("類別名稱") = "輸入" Then
    Form1.客戶統計.Enabled = False
    Form1.產品銷售.Enabled = False
    Form1.對帳單.Enabled = False
    Form1.出貨及退貨.Enabled = False
ElseIf NowUser("類別名稱") = "出貨" Then
    Form1.客戶統計.Enabled = False
    Form1.產品銷售.Enabled = False
    Form1.訂單處理.Enabled = False
End If
Unload Me
End Sub
