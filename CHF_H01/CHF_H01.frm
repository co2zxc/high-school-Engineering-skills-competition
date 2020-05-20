VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '單線固定
   Caption         =   "102學年度  工業類科學生技藝競賽  電腦修護職種  第二站  崗位號碼：98"
   ClientHeight    =   6450
   ClientLeft      =   180
   ClientTop       =   255
   ClientWidth     =   8400
   FillColor       =   &H000000FF&
   FillStyle       =   0  '實心
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   12
      Charset         =   136
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   8400
   StartUpPosition =   2  '螢幕中央
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
      Index           =   0
      Left            =   5248
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   3960
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3473
      TabIndex        =   1
      Text            =   "98"
      Top             =   5213
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
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
      Left            =   5248
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   5040
      Width           =   2200
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   2880
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
      Left            =   953
      TabIndex        =   5
      Top             =   600
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label StatusDisplay 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "No Device"
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
      Height          =   375
      Left            =   2100
      TabIndex        =   3
      Top             =   4140
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "Station Number"
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
      Height          =   375
      Left            =   953
      TabIndex        =   2
      Top             =   5228
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   7
      Left            =   938
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   6
      Left            =   1778
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   5
      Left            =   2618
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   4
      Left            =   3458
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   3
      Left            =   4298
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   2
      Left            =   5138
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   1
      Left            =   5978
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   645
      Index           =   0
      Left            =   6818
      Shape           =   3  '圓形
      Top             =   2760
      Width           =   645
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
Public TxHour As Byte
Public TxMinute As Byte
Public TxSecond As Byte
Public TxStNumber As Byte
'

Private Sub Command1_Click(Index As Integer)
    a = Index
    c = 0
End Sub
'

Private Sub Form_Load()
    a = 99
    b(0) = &H1
    b(1) = &H3
    b(2) = &H7
    b(3) = &HF
    b(4) = &H1F
    b(5) = &H3F
    b(6) = &H7F
    b(7) = &HFF
    b(8) = &HFE
    b(9) = &HFC
    b(10) = &HF8
    b(11) = &HF0
    b(12) = &HE0
    b(13) = &HC0
    b(14) = &H80
    b(15) = &H0
    c = 0
    TxStNumber = Val(Text1.Text)
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
    Dim Result As Boolean
    Result = OpenUsbDevice(VendorID, ProductID)
    If (Result) Then
        TxRedLED = &H0
        TxCtrl = &H20
        TxStNumber = 0
        OutDataEightByte TxRedLED, TxCtrl, 0, 0, 0, TxStNumber, 0, 0
        ShowRedLED (&H0)
        CloseUsbDevice
    End If
End Sub
'

Private Sub Text1_Change()
    If (Val(Text1.Text) > 99) Then
        Text1.Text = 99
    End If
    TxStNumber = Val(Text1.Text)
End Sub
'

Private Sub Timer1_Timer()
    Dim Result As Boolean
    Dim k As Byte
    Static PreConnect As Boolean
    If c > 20 Then c = 20
    Label1.Caption = "                                  " & _
                     "Current Time : " & Time$

    Result = OpenUsbDevice(VendorID, ProductID)
    
    If (Result) Then
        ShowNormalLEDs
        If ((PreConnect = False) And (Result = True)) Then
            TxRedLED = &H0
            OutRedLED (TxRedLED)
            ShowRedLED (&H0)
        End If
        
        If a = 0 Then
            If c > 15 Then
                TxRedLED = &H0
                ShowRedLED (&H0)
                a = 99
                c = 0
            Else
                TxRedLED = b(c)
                ShowRedLED (b(c))
            End If
        End If
        
        If a = 1 Then
            TxRedLED = &H0
            TxCtrl = &H20
            TxStNumber = 0
            OutDataEightByte TxRedLED, TxCtrl, 0, 0, 0, TxStNumber, 0, 0
            ShowRedLED (&H0)
            a = 99
            CloseUsbDevice
            PreConnect = Result
            End
        End If
        
        TxHour = Hour(Now)
        TxMinute = Minute(Now)
        TxSecond = Second(Now)
        TxCtrl = &H20
        OutDataEightByte TxRedLED, TxCtrl, TxHour, TxMinute, TxSecond, TxStNumber, 0, 0
        
        ReadData ReadKeyData
        CloseUsbDevice

        Select Case (ReadKeyData(1) And &H3)

            Case 0
                StatusDisplay.Caption = " Moving Logo"

            Case 1
                StatusDisplay.Caption = " Display Station Number "

            Case 2
                 StatusDisplay.Caption = " Display Second "

            Case 3
                 StatusDisplay.Caption = " Display Current Time "

        End Select
    Else
        ShowInitLEDs
         StatusDisplay.Caption = " Device Off "
        If a = 1 Then
            ShowRedLED (&H0)
            a = 99
            PreConnect = Result
            End
        End If
    End If
    PreConnect = Result
    c = c + 1

End Sub
'
