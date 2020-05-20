VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7080
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

    Private Sub Command1_Click()

    Dim strPathName As String

    strPathName = ""

    strPathName = InputBox("請輸入需要刪除的文件夾名稱︰", "刪除文件夾")

    If strPathName = "" Then Exit Sub

    On Error GoTo ErrorHandle

    SetAttr strPathName, vbNormal '此行主要是為了檢查文件夾名稱的有效性

    RecurseTree strPathName

    Label1.Caption = "文件夾" & strPathName & "已經刪除！"

    Exit Sub

ErrorHandle:

    MsgBox "無效的文件夾名稱:" & strPathName

    End Sub

    Sub RecurseTree(CurrPath As String)

    Dim sFileName As String

    Dim newPath As String

    Dim sPath As String

    Static oldPath As String

    sPath = CurrPath & "\"

    sFileName = Dir(sPath, 31) '31的含義︰31=vbNormal+vbReadOnly+vbHidden+vbSystem+vbVolume+vbDirectory

    Do While sFileName <> ""

    If sFileName <> "." And sFileName <> ".." Then

    If GetAttr(sPath & sFileName) And vbDirectory Then '如果是目錄和文件夾

    newPath = sPath & sFileName

    RecurseTree newPath

    sFileName = Dir(sPath, 31)

    Else

    SetAttr sPath & sFileName, vbNormal

    Kill (sPath & sFileName)

    Label1.Caption = sPath & sFileName '顯示刪除過程

    sFileName = Dir

    End If

    Else

    sFileName = Dir

    End If

    DoEvents

    Loop

    SetAttr CurrPath, vbNormal

    RmDir CurrPath

    Label1.Caption = CurrPath

    End Sub
