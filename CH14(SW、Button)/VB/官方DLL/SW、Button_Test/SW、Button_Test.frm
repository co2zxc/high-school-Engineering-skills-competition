VERSION 5.00
Begin VB.Form USB 
   Caption         =   "SW、Button_Test"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   6750
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6240
      Top             =   1440
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
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   2085
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
      Left            =   6120
      TabIndex        =   7
      Top             =   240
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
      Left            =   5280
      TabIndex        =   6
      Top             =   240
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
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   495
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
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   495
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
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   375
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
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "SD3"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "SD4"
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
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   0
      Left            =   5985
      Shape           =   3  '圓形
      Top             =   600
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   1
      Left            =   5145
      Shape           =   3  '圓形
      Top             =   600
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   2
      Left            =   4305
      Shape           =   3  '圓形
      Top             =   600
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   3
      Left            =   3465
      Shape           =   3  '圓形
      Top             =   600
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   4
      Left            =   2625
      Shape           =   3  '圓形
      Top             =   600
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   5
      Left            =   1785
      Shape           =   3  '圓形
      Top             =   600
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   6
      Left            =   945
      Shape           =   3  '圓形
      Top             =   600
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   7
      Left            =   105
      Shape           =   3  '圓形
      Top             =   600
      Width           =   645
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)  ' 確認USB devie是否連線
     
    If (Check) Then
        Label9.Caption = " DEVICE ON "          ' DEVICE ON
    Else
        Label9.Caption = " DEVICE OFF "         ' DEVICE OFF
    End If
    
    CloseUsbDevice                              ' 關閉HID裝置
    isConnect = Check
End Function

Private Sub USBAction()
    Dim Check As Boolean
    Dim ReadKeyData(8) As Byte
    Dim tt As Byte
    Check = OpenUsbDevice(VendorID, ProductID)  ' 確認USB devie是否連線
    If (Check) Then
        ReadData ReadKeyData                    ' 讀取HID資料
        For tt = 0 To 7
            If ((ReadKeyData(1) And 2 ^ tt) = 2 ^ tt) Then
                Shape2(tt).FillColor = &HFF&    ' LED 亮
            Else
                Shape2(tt).FillColor = &H80&    ' LED 滅
            End If
        Next tt
        CloseUsbDevice                          ' 關閉HID裝置
    End If
End Sub

Private Sub Timer1_Timer()
    isConnect
    USBAction
End Sub



