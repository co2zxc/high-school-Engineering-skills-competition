VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "電腦硬體裝修乙級 第一站 第八題"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   8355
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Green LED"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   1200
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   2640
      Width           =   2200
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   1
      Left            =   4680
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   3840
      Width           =   2200
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Red LED"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   4680
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   2640
      Width           =   2200
   End
   Begin VB.Label StatusDisplay 
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
      Left            =   1200
      TabIndex        =   3
      Top             =   4080
      Width           =   2085
   End
   Begin VB.Label Label1 
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
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   6495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TxRedLED As Byte
Public TxCtrl As Byte
Public Temp As Byte

Private Sub Command1_Click(Index As Integer)
    a = Index
    Temp = &H21
    c = 0
End Sub

Private Sub Command2_Click(Index As Integer)
    a = Index
    Temp = &H22
    c = 0
End Sub

Private Sub Form_Load()
    a = 99
    
    ' NO.8
    b(0) = &H81
    b(1) = &H42
    b(2) = &H24
    b(3) = &H18
    d = 3
    
    c = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim Result As Boolean
    Result = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Result) Then
        TxRedLED = &H0                                          ' USB 資料
        TxCtrl = &H20                                           ' 前導碼
        OutDataEightByte TxRedLED, TxCtrl, Temp, 0, 0, 0, 0, 0  ' 傳輸USB資料
        CloseUsbDevice                                          ' 關閉HID裝置
    End If
End Sub


Private Sub Timer1_Timer()
    Dim Result As Boolean
    Dim k As Byte
    Static PreConnect As Boolean
    If c > 20 Then c = 20
    Label1.Caption = "                                  " & _
                     "Current Time : " & Time$

    Result = OpenUsbDevice(VendorID, ProductID)                     ' 確認USB devie是否連線
    
    If (Result) Then
        StatusDisplay.Caption = " DEVICE ON "                       ' DEVICE ON
        If ((PreConnect = False) And (Result = True)) Then
            TxRedLED = &H0                                          ' USB 資料
            OutDataEightByte TxRedLED, TxCtrl, Temp, 0, 0, 0, 0, 0  ' 傳輸USB資料
        End If
        
        If a = 0 Then
            If c > d Then
                TxRedLED = &H0                                      ' USB 資料
                a = 99
                c = 0
            Else
                TxRedLED = b(c)                                     ' USB 資料
            End If
        End If
        
        If a = 1 Then
            TxRedLED = &H0                                          ' USB 資料
            TxCtrl = &H20                                           ' 前導碼
            OutDataEightByte TxRedLED, TxCtrl, Temp, 0, 0, 0, 0, 0  ' 傳輸USB資料
            a = 99
            CloseUsbDevice                                          ' 關閉HID裝置
            PreConnect = Result
            End
        End If
        
        TxCtrl = &H20                                               ' 前導碼
        OutDataEightByte TxRedLED, TxCtrl, Temp, 0, 0, 0, 0, 0      ' 傳輸USB資料
        CloseUsbDevice                                              ' 關閉HID裝置
    Else
        StatusDisplay.Caption = " DEVICE OFF "                      ' DEVICE OFF
        If a = 1 Then
            a = 99
            PreConnect = Result
            End
        End If
    End If
    PreConnect = Result
    c = c + 1

End Sub

