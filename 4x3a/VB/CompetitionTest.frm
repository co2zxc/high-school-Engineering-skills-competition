VERSION 5.00
Begin VB.Form USB 
   Caption         =   "CompetitionTest"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7530
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   120
      Top             =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "tt hh:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   26.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   " DEVICE OFF "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5040
      TabIndex        =   0
      Top             =   2520
      Width           =   2085
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReadKeyData(8) As Byte

Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
     
    If (Check) Then
        Label1.Caption = " DEVICE ON "                  ' DEVICE ON
    Else
        Label1.Caption = " DEVICE OFF "                 ' DEVICE OFF
    End If
    
    CloseUsbDevice                                      ' 關閉HID裝置
    isConnect = Check
End Function

Private Sub Timer1_Timer()
    Dim timeStringArray() As String
    Dim Alarm As String
    Dim Check As Boolean

    isConnect
    timeStringArray = Split(Time$, ":", 3)                                                  ' 將時間分成時分秒個別儲存
    
    Label2.Caption = "                                  " & _
                     "Current Time : " & Time$
    
    Check = OpenUsbDevice(VendorID, ProductID)                                              ' 確認USB I/O介面卡實驗板是否連接
    If (Check) Then
        OutDataEightByte &H20, Int(timeStringArray(0) / 10), timeStringArray(0) Mod 10, _
                         Int(timeStringArray(1) / 10), timeStringArray(1) Mod 10, _
                         Int(timeStringArray(2) / 10), timeStringArray(2) Mod 10, 0      ' 將時分秒分成十位數個位數傳送至 USB Device
        ReadData ReadKeyData
        CloseUsbDevice                                                                      ' 關閉HID裝置
    End If
    
' 透過 USB 由 VB 端將資料取出
    Text1.Text = (ReadKeyData(1) - 48) & (ReadKeyData(2) - 48) & ":" & (ReadKeyData(3) - 48) & (ReadKeyData(4) - 48) _
                 & ":" & (ReadKeyData(5) - 48) & (ReadKeyData(6) - 48)
End Sub
