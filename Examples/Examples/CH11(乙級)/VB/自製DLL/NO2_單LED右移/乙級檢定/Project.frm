VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "電腦硬體裝修乙級 第一站 第二題"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   8970
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1320
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   3720
      Width           =   2200
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3360
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   4920
      Width           =   2200
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
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
      Index           =   2
      Left            =   5400
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   3720
      Width           =   2200
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   8
      Left            =   4680
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   1
      Left            =   8040
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   2
      Left            =   7560
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   3
      Left            =   7080
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   4
      Left            =   6600
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   5
      Left            =   6120
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   6
      Left            =   5640
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   7
      Left            =   5160
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   1
      Left            =   4200
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   2
      Left            =   3720
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   3
      Left            =   3240
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   4
      Left            =   2760
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   5
      Left            =   2280
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   6
      Left            =   1800
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   7
      Left            =   1320
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   8
      Left            =   840
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   375
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
      Height          =   1815
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   7575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   全華圖書: USB介面設計與應用設計入門 Visual Basic 範例程式-CH11(許永和編著)
'   程式功能: 利用PB或PD驅動8顆LED控制(乙級檢定)

Dim a As Integer
Dim b(99) As Byte
Dim c As Integer
Dim d As Integer
Dim Send(8) As Byte
Const VendorID = &H1234
Const ProductID = &H2468
Dim usbHID As New HID

Private Sub Command1_Click(Index As Integer)
    a = Index
    Send(1) = &H21
    c = 0
End Sub

Private Sub Command2_Click(Index As Integer)
    a = Index
    Send(1) = &H22
    c = 0
End Sub

Private Sub Form_Load()
    a = 99
    
    ' NO.2
    b(0) = &H80
    b(1) = &H40
    b(2) = &H20
    b(3) = &H10
    b(4) = &H8
    b(5) = &H4
    b(6) = &H2
    b(7) = &H1
    d = 7
    
    c = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim Result As Boolean
    Result = usbHID.OpenHIDDevice(VendorID, ProductID)          ' 確認USB I/O介面卡實驗板是否連接
    If (Result) Then
        Send(0) = &H0                                           ' USB 資料
        usbHID.HIDSetReport Send                                ' 傳輸USB資料
        usbHID.CloseHIDDevice                                   ' 關閉HID裝置
    End If
End Sub


Private Sub Timer1_Timer()
    Dim Result As Boolean
    Dim k As Byte
    Dim i As Integer
    Static PreConnect As Boolean
    If c > 20 Then c = 20
    Label1.Caption = "                                            " & _
                     "Current Time : " & Time$

    Result = usbHID.OpenHIDDevice(VendorID, ProductID)              ' 確認USB I/O介面卡實驗板是否連接
    
    If (Result) Then
        
        If ((PreConnect = False) And (Result = True)) Then
            Send(0) = &H0                                           ' USB 資料
            Send(1) = &H21                                          ' 前導碼
            
            For i = 1 To 8                                             ' DEVICE ON
            Shape1(i).FillStyle = 0
            Shape2(i).FillStyle = 0
            Next i
            
            usbHID.HIDSetReport Send                                ' 傳輸USB資料
        End If
        
        If a = 0 Then
            If c > d Then
                Send(0) = &H0                                       ' USB 資料
                a = 99
                Shape2(9 - c).FillColor = &H8000&
                c = 0
            Else
                If c > 0 Then
                Shape2(9 - c).FillColor = &H8000&
                End If
                Send(0) = b(c)                                      ' USB 資料
                Shape2(8 - c).FillColor = &HFF00&
            End If
            
        End If
        
        If a = 1 Then
            Send(0) = &H0                                           ' USB 資料
            usbHID.HIDSetReport Send                                ' 傳輸USB資料
            a = 99
            usbHID.CloseHIDDevice                                   ' 關閉HID裝置
            PreConnect = Result
            End
        End If
        
        If a = 2 Then
            If c > d Then
                Send(0) = &H0                                       ' USB 資料
                a = 99
                Shape1(c).FillColor = &H80&
                c = 0
            Else
                If c > 0 Then
                Shape1(c).FillColor = &H80&
                End If
                Send(0) = 2 ^ c                                    ' USB 資料
                Shape1(c + 1).FillColor = &HFF&
            End If
            
        End If
        
        usbHID.HIDSetReport Send                                    ' 傳輸USB資料
        usbHID.CloseHIDDevice                                       ' 關閉HID裝置
    Else
        For i = 1 To 8
        Shape1(i).FillStyle = 1
        Shape2(i).FillStyle = 1                                     ' DEVICE OFF
        Next i

        If a = 1 Then
            a = 99
            PreConnect = Result
            End
        End If
    End If
    PreConnect = Result
    c = c + 1
End Sub


