VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmJukebox 
   BorderStyle     =   1  '單線固定
   Caption         =   "點唱機"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4440
   StartUpPosition =   1  '所屬視窗中央
   Begin MSComDlg.CommonDialog cdlgFile 
      Left            =   240
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "請選擇音樂檔案"
      Filter          =   "MP3檔案|*.mp3|Wave音效檔|*.wav|MIDI數位音效檔|*.mid|全部檔案|*.*"
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "上一首(&P)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一首(&N)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清除(&C)"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "刪除(&D)"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "加入(&A)"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "停止(&S)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "播放(&P)"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox lstFile 
      Height          =   2040
      Left            =   1080
      MultiSelect     =   2  '進階多重選取
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin MCI.MMControl mmc1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "frmJukebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Dim aryFileList, i As Integer
On Error GoTo Error_Handler
    Dim Flags As Long
    Flags = &H800& Or &H4& Or &H200& Or &H1000&
    cdlgFile.Flags = Flags
    cdlgFile.ShowOpen
    aryFileList = Split(cdlgFile.FileName)
    If UBound(aryFileList) = 0 Then
      lstFile.AddItem aryFileList(0)
    Else
      If Right(aryFileList(0), 1) <> "\" Then aryFileList(0) = aryFileList(0) & "\"
      For i = 1 To UBound(aryFileList)
          lstFile.AddItem aryFileList(0) & aryFileList(i)
      Next
    End If
Error_Handler:
End Sub

Private Sub cmdClear_Click()
    lstFile.Clear
End Sub

Private Sub cmdDelete_Click()
   Dim i As Integer
   If lstFile.SelCount > 0 Then
    For i = lstFile.ListCount - 1 To 0 Step -1
      If lstFile.Selected(i) Then lstFile.RemoveItem i
    Next
   End If
End Sub

Private Sub cmdNext_Click()
   If (lstFile.ListCount - 1) > lstFile.ListIndex Then
      lstFile.ListIndex = lstFile.ListIndex + 1
      mmc1.Command = "stop"
      mmc1.Command = "close"
      mmc1.FileName = lstFile.List(lstFile.ListIndex)
      mmc1.Command = "open"
      mmc1.Command = "play"
    End If
End Sub

Private Sub cmdPlay_Click()
   lstFile.ListIndex = 0
   cmdPlay.Enabled = False
   cmdAdd.Enabled = False
   cmdDelete.Enabled = False
   cmdClear.Enabled = False
   cmdStop.Enabled = True
   cmdNext.Enabled = True
   cmdPrev.Enabled = True
   mmc1.FileName = lstFile.List(lstFile.ListIndex)
   mmc1.Command = "close"
   mmc1.Command = "open"
   mmc1.Command = "play"
End Sub

Private Sub cmdPrev_Click()
   If lstFile.ListIndex > 0 Then
      lstFile.ListIndex = lstFile.ListIndex - 1
      mmc1.Command = "stop"
      mmc1.Command = "close"
      mmc1.FileName = lstFile.List(lstFile.ListIndex)
      mmc1.Command = "open"
      mmc1.Command = "play"
    End If
End Sub

Private Sub cmdQuit_Click()
   Unload Me
frmMDIMain.Show
End Sub

Private Sub cmdStop_Click()
  mmc1.Command = "stop"
  mmc1.Command = "close"
  cmdPlay.Enabled = True
  cmdAdd.Enabled = True
  cmdDelete.Enabled = True
  cmdClear.Enabled = True
  cmdStop.Enabled = False
  cmdNext.Enabled = False
  cmdPrev.Enabled = False
End Sub

Private Sub mmc1_StatusUpdate()
   If mmc1.Mode = mciModeStop Then
    If lstFile.ListCount - 1 > lstFile.ListIndex Then
        lstFile.ListIndex = lstFile.ListIndex + 1
        mmc1.Command = "close"
        mmc1.FileName = lstFile.List(lstFile.ListIndex)
        mmc1.Command = "open"
        mmc1.Command = "play"
    Else
        mmc1.Command = "close"
        cmdPlay.Enabled = True
        cmdAdd.Enabled = True
    End If
   End If
End Sub
