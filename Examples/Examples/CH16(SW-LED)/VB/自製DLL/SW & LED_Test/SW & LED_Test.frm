VERSION 5.00
Begin VB.Form USB 
   Caption         =   "SW & LED_Test"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   6360
   StartUpPosition =   3  '系統預設值
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   23
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   12
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   11
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   6
      Left            =   840
      TabIndex        =   10
      Top             =   2280
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5760
      Top             =   1800
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   3
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   1
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   0
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label9 
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
      Left            =   4080
      TabIndex        =   24
      Top             =   2280
      Width           =   2085
   End
   Begin VB.Label Label20 
      Caption         =   "D26"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "D25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "D24"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "D23"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   "D22"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "D21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "D20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "D19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Input Test :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Output Test :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   5
      Left            =   240
      Shape           =   3  '圓形
      Top             =   840
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   4
      Left            =   1080
      Shape           =   3  '圓形
      Top             =   840
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   3
      Left            =   1920
      Shape           =   3  '圓形
      Top             =   840
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   2
      Left            =   2760
      Shape           =   3  '圓形
      Top             =   840
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   1
      Left            =   3600
      Shape           =   3  '圓形
      Top             =   840
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   0
      Left            =   4440
      Shape           =   3  '圓形
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "SD2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   375
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "SD1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   8
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "SW1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2055
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "SW2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2895
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "SW3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3735
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "SW4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4575
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   全華圖書: USB介面設計與應用設計入門 Visual Basic 範例程式-CH16(許永和編著)
'   程式功能: 利用PB或PD驅動8顆紅色LED，以及透過PC0~PC5讀取指撥開關與按鍵值

Const VendorID = &H1234                                     ' 設定販售商碼
Const ProductID = &H2468                                    ' 設定產品碼
Dim usbHID As New HID

Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = usbHID.OpenHIDDevice(VendorID, ProductID)       ' 確認USB I/O介面卡實驗板是否連接
    
    If (Check) Then
        Label9.Caption = " DEVICE ON "                      ' DEVICE ON
    Else
        Label9.Caption = " DEVICE OFF "                     ' DEVICE OFF
    End If
    
    usbHID.CloseHIDDevice                                   ' 關閉HID裝置
    isConnect = Check
End Function

Private Sub USBAction()
    Dim Check As Boolean
    Dim ReadKeyData() As Byte
    Dim tt As Byte
    Dim Send(8) As Byte
    Dim i As Byte
    Check = usbHID.OpenHIDDevice(VendorID, ProductID)       ' 確認USB I/O介面卡實驗板是否連接
    If (Check) Then                                         ' 判斷 HID Device 是否存在
        ReadKeyData() = usbHID.HidGetReport                 ' 讀取HID資料
            
        For tt = 0 To 5
            If ((ReadKeyData(1) And 2 ^ tt) = 2 ^ tt) Then
                Shape2(tt).FillColor = &HFF&                ' LED 亮
            Else
                Shape2(tt).FillColor = &H80&                ' LED 滅
            End If
        Next tt
        For i = 0 To 7
            If Check1(i).Value = Checked Then               ' 取得 CheckBox 狀態
                Send(1) = Send(1) + 2 ^ i
            End If
        Next i
        Send(0) = &HFF - Send(1)                            ' 反相
        
        usbHID.HIDSetReport Send                            ' 傳輸HID資料
        usbHID.CloseHIDDevice                               ' 關閉HID裝置
    End If
End Sub

Private Sub Timer1_Timer()
    isConnect
    USBAction
End Sub




