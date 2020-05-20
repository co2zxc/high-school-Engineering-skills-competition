VERSION 5.00
Begin VB.Form USB 
   Caption         =   "CompetitionTest"
   ClientHeight    =   2985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   7365
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   720
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000000FF&
      Caption         =   "Red LED"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2520
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   1725
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   0
      Top             =   1440
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
      Left            =   4920
      TabIndex        =   5
      Top             =   2400
      Width           =   2085
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "Key off "
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
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   1185
   End
   Begin VB.Label Label7 
      Caption         =   "Input Test :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Output Test :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2055
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
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   6495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   全華圖書: USB介面設計與應用設計入門 Visual Basic 範例程式-CH12(許永和編著)
'   程式功能: 利用PB或PD驅動8顆LED，以及透過PC0~PC5、PD1與PD3驅動4顆8x8點矩陣，和使用PD0讀取按鍵值

Const VendorID = &H1234
Const ProductID = &H2468
Dim USBHID As New HID
Dim Data(4, 8) As Byte
Dim numberData(11, 5) As Byte
Dim timeI As Byte
Dim rl As Byte
Dim A As Byte
Dim B As Byte
Dim C As Byte

Private Sub numberDataInit()
'number 0
numberData(0, 0) = &H0
numberData(0, 1) = &H7F
numberData(0, 2) = &H41
numberData(0, 3) = &H7F
numberData(0, 4) = &H0
'number 1
numberData(1, 0) = &H0
numberData(1, 1) = &H42
numberData(1, 2) = &H7F
numberData(1, 3) = &H40
numberData(1, 4) = &H0
'number 2
numberData(2, 0) = &H0
numberData(2, 1) = &H79
numberData(2, 2) = &H49
numberData(2, 3) = &H4F
numberData(2, 4) = &H0
'number 3
numberData(3, 0) = &H0
numberData(3, 1) = &H49
numberData(3, 2) = &H49
numberData(3, 3) = &H7F
numberData(3, 4) = &H0
'number 4
numberData(4, 0) = &H0
numberData(4, 1) = &HF
numberData(4, 2) = &H8
numberData(4, 3) = &H7F
numberData(4, 4) = &H0
'number 5
numberData(5, 0) = &H0
numberData(5, 1) = &H4F
numberData(5, 2) = &H49
numberData(5, 3) = &H79
numberData(5, 4) = &H0
'number 6
numberData(6, 0) = &H0
numberData(6, 1) = &H7F
numberData(6, 2) = &H49
numberData(6, 3) = &H79
numberData(6, 4) = &H0
'number 7
numberData(7, 0) = &H0
numberData(7, 1) = &H7
numberData(7, 2) = &H1
numberData(7, 3) = &H7F
numberData(7, 4) = &H0
'number 8
numberData(8, 0) = &H0
numberData(8, 1) = &H7F
numberData(8, 2) = &H49
numberData(8, 3) = &H7F
numberData(8, 4) = &H0
'number 9
numberData(9, 0) = &H0
numberData(9, 1) = &HF
numberData(9, 2) = &H9
numberData(9, 3) = &H7F
numberData(9, 4) = &H0
':
numberData(10, 0) = &H0
numberData(10, 1) = &H66
numberData(10, 2) = &H66
numberData(10, 3) = &H0
numberData(10, 4) = &H0
rl = 0
timeI = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Send(8) As Byte
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)          ' 確認USB I/O介面卡實驗板是否連接
    If (Check) Then
        For i = 0 To 7
            Send(0) = &H20                                     ' 前導碼
            Send(1) = &H0                                      ' 8*8第一組
            Send(2) = i                                        ' 8*8第n行
            USBHID.HIDSetReport Send                           ' 傳輸HID資
            
            Send(1) = &H1                                      ' 8*8第二組
            USBHID.HIDSetReport Send                           ' 傳輸HID資
            
            Send(1) = &H2                                      ' 8*8第三組
            USBHID.HIDSetReport Send                           ' 傳輸HID資
            
            Send(1) = &H3                                      ' 8*8第四組
            USBHID.HIDSetReport Send                           ' 傳輸HID資
        Next i
    End If
    USBHID.CloseHIDDevice                                      ' 關閉HID裝置
End Sub

Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)   ' 確認USB I/O介面卡實驗板是否連接
     
    If (Check) Then
        Label1.Caption = " DEVICE ON "                  ' DEVICE ON
    Else
        Label1.Caption = " DEVICE OFF "                 ' DEVICE OFF
    End If
    
    USBHID.CloseHIDDevice                               ' 關閉HID裝置
    isConnect = Check
End Function

Private Sub Command9_Click()
    A = 1
    B = 0
    C = 0
End Sub

Private Sub Form_load()
    numberDataInit
End Sub



Private Sub Timer1_Timer()
    Dim i As Byte
    Dim ReadKeyData() As Byte
    Dim Send(8) As Byte
    Dim timeStringArray() As String
    Dim timeArray(9) As Byte
    timeStringArray = Split(Time$, ":", 3)
    
    Label2.Caption = "                                  " & _
                     "Current Time : " & Time$
    
    For i = 0 To 2
        timeArray(3 * i) = hTOd(timeStringArray(i)) \ 16
        timeArray(3 * i + 1) = hTOd(timeStringArray(i)) Mod 16
        timeArray(3 * i + 2) = &HA
    Next i
    
    isConnect
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)
    If (A) Then
        C = C + 1
        If (C Mod 10 = 0) Then
            Send(0) = &H21                                     ' 前導碼
            Send(1) = 2 ^ B                                    ' LED資料
            USBHID.HIDSetReport Send                           ' 傳輸HID資
            B = B + 1
            If (B = 8) Then
                A = 0
            End If
        End If
    Else
        C = C + 1
        If (C Mod 10 = 0) Then
            Send(0) = &H21                                     ' 前導碼
            Send(1) = &H0                                      ' LED資料
            USBHID.HIDSetReport Send                           ' 傳輸HID資
            C = 0
        End If
    End If
    
    'left
    'If (rl <= 0) Then
    '   rl = 33
    'End If
    
    'rl = rl - 1
    
    'For i = 0 To 32
        'left
        Send(0) = &H20                                     ' 前導碼
        Send(1) = ((timeI + rl) \ 8) Mod 4                 ' 8*8第n組
        Send(2) = (timeI + rl) Mod 8                       ' 8*8第n行
        Send(3) = numberData(timeArray(timeI \ 4), (timeI Mod 4) + 1) ' 8*8資料
        USBHID.HIDSetReport Send
        
        'right
        'Send(0) = &H20                                     ' 前導碼
        'Send(1) = ((timeI + rl) \ 8) Mod 4                 ' 8*8第n組
        'Send(2) = (timeI + rl) Mod 8                       ' 8*8第n行
        'Send(3) = numberData(timeArray(timeI \ 4), (timeI Mod 4)) ' 8*8資料
        'USBHID.HIDSetReport Send
    'Next i
    'right
    'rl = rl + 1
    'If (rl >= 32) Then
    '    rl = 0
    'End If
    
    ReadKeyData() = USBHID.HidGetReport                        ' 讀取HID資料
    USBHID.CloseHIDDevice                                      ' 關閉HID裝置
    
    Select Case (ReadKeyData(1) And &H1)
        Case 0
            Label6.Caption = " KEY ON "
        Case 1
            Label6.Caption = " KEY OFF "
    End Select
End Sub

Public Function hTOd(ByVal HEXA As String) As Long
    Dim strTmp As String
    Dim i As Integer
    Dim Numslen As Integer
    Dim Tmpstr As Long
    Tmpstr = 0
    Numslen = Len(HEXA)
    For i = 1 To Numslen
        strTmp = Mid(HEXA, i, 1)
        If IsNumeric(strTmp) Then
           strTmp = strTmp * 16 * (16 ^ (Numslen - i - 1))
        Else
           If Asc(UCase(strTmp)) < 65 Or Asc(UCase(strTmp)) > 70 Then
              hTOd = -1
           End If
           strTmp = (Asc(UCase(strTmp)) - 55) * (16 ^ (Numslen - i))
        End If
        Tmpstr = Tmpstr + CLng(strTmp)
    Next
    hTOd = Tmpstr
End Function

Private Sub Timer2_Timer()

    timeI = timeI + 1
    If (timeI >= 33) Then
        'left
        If (rl <= 0) Then
           rl = 33
        End If
    
         rl = rl - 1
         'right
        'rl = rl + 1
        'If (rl >= 32) Then
        '    rl = 0
        'End If
        timeI = 0
    End If
    
End Sub


