VERSION 5.00
Begin VB.Form USB 
   Caption         =   "全國高級中等學校102學年度工業類科學生技藝競賽"
   ClientHeight    =   4755
   ClientLeft      =   5730
   ClientTop       =   4800
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6975
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   360
      Top             =   4320
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Text            =   "00"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   4320
   End
   Begin VB.CommandButton Command9 
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
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   8
      Left            =   120
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   7
      Left            =   960
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   6
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   4
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   6000
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Station Number:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
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
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   3855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
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
Dim numberData(11, 7) As Byte
Dim clock(32), pic(64) As Byte
Dim timeI, ttt, keyin, logo As Byte
Dim a, B, C, D, x As Byte ' LED
Dim ap As Boolean ' 分號


Private Sub numberDataInit()
'number 0
pic(0) = &H3F
pic(1) = &HA1
pic(2) = &HE1
pic(3) = &HE1
pic(4) = &HE1
pic(5) = &HE1
pic(6) = &HA1
pic(7) = &H3F
pic(8) = &H0
pic(9) = &H6
pic(10) = &H5
pic(11) = &H9
pic(12) = &H91
pic(13) = &HE1
pic(14) = &HE1
pic(15) = &H91
pic(16) = &H9
pic(17) = &H5
pic(18) = &H6
pic(19) = &H0
pic(20) = &HC0
pic(21) = &HC0
pic(22) = &HF8
pic(23) = &H88
pic(24) = &H8F
pic(25) = &H81
pic(26) = &H8F
pic(27) = &H88
pic(28) = &HF8
pic(29) = &HC0
pic(30) = &HC0
pic(31) = &H0
pic(32) = &H0
pic(33) = &H0
pic(34) = &H0
pic(35) = &H0
pic(36) = &H0
pic(37) = &H0
pic(38) = &H0
pic(39) = &H0
pic(40) = &H0
pic(41) = &H0
pic(42) = &H0
pic(43) = &H0
pic(44) = &H0
pic(45) = &H0
pic(46) = &H0
pic(47) = &H0
pic(48) = &H0
pic(49) = &H0
pic(50) = &H0
pic(51) = &H0
pic(52) = &H0
pic(53) = &H0
pic(54) = &H0
pic(55) = &H0
pic(56) = &H0
pic(57) = &H0
pic(58) = &H0
pic(59) = &H0
pic(60) = &H0
pic(61) = &H0
pic(62) = &H0
pic(63) = &H0
numberData(0, 0) = &H0
numberData(0, 1) = &H7E
numberData(0, 2) = &H81
numberData(0, 3) = &H81
numberData(0, 4) = &H81
numberData(0, 5) = &H7E
numberData(0, 6) = &H0
'number 1
numberData(1, 0) = &H0
numberData(1, 1) = &H0
numberData(1, 2) = &H82
numberData(1, 3) = &HFF
numberData(1, 4) = &H80
numberData(1, 5) = &H0
numberData(1, 6) = &H0
'number 2
numberData(2, 0) = &H0
numberData(2, 1) = &HC2
numberData(2, 2) = &HA1
numberData(2, 3) = &H91
numberData(2, 4) = &H89
numberData(2, 5) = &H86
numberData(2, 6) = &H0
'number 3
numberData(3, 0) = &H0
numberData(3, 1) = &H41
numberData(3, 2) = &H81
numberData(3, 3) = &H89
numberData(3, 4) = &H95
numberData(3, 5) = &H63
numberData(3, 6) = &H0
'number 4
numberData(4, 0) = &H0
numberData(4, 1) = &H38
numberData(4, 2) = &H24
numberData(4, 3) = &H22
numberData(4, 4) = &HFF
numberData(4, 5) = &H20
numberData(4, 6) = &H0
'number 5
numberData(5, 0) = &H0
numberData(5, 1) = &H4F
numberData(5, 2) = &H89
numberData(5, 3) = &H89
numberData(5, 4) = &H89
numberData(5, 5) = &H71
numberData(5, 6) = &H0
'number 6
numberData(6, 0) = &H0
numberData(6, 1) = &H7C
numberData(6, 2) = &H92
numberData(6, 3) = &H91
numberData(6, 4) = &H91
numberData(6, 5) = &H60
numberData(6, 6) = &H0
'number 7
numberData(7, 0) = &H0
numberData(7, 1) = &H1
numberData(7, 2) = &H1
numberData(7, 3) = &HF9
numberData(7, 4) = &H5
numberData(7, 5) = &H3
numberData(7, 6) = &H0
'number 8
numberData(8, 0) = &H0
numberData(8, 1) = &H66
numberData(8, 2) = &H99
numberData(8, 3) = &H91
numberData(8, 4) = &H99
numberData(8, 5) = &H66
numberData(8, 6) = &H0
'number 9
numberData(9, 0) = &H0
numberData(9, 1) = &H6
numberData(9, 2) = &H89
numberData(9, 3) = &H89
numberData(9, 4) = &H49
numberData(9, 5) = &H3E
numberData(9, 6) = &H0
':
numberData(10, 0) = &H0
numberData(10, 1) = &H66
numberData(10, 2) = &H66
numberData(10, 3) = &H0
timeI = 0
ap = 0
keyin = 0
x = 0
logo = 0
End Sub

Private Sub Command1_Click()
Dim Send(8) As Byte
Check = USBHID.OpenHIDDevice(VendorID, ProductID)
    Send(0) = &H23
     USBHID.HIDSetReport Send
    USBHID.CloseHIDDevice '關閉HID裝置
    
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Send(8) As Byte
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)          ' 確認USB I/O介面卡實驗板是否連接
    Send(0) = &H23
    USBHID.HIDSetReport Send
    USBHID.CloseHIDDevice                                      ' 關閉HID裝置
    
End Sub

Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)   ' 確認USB I/O介面卡實驗板是否連接
     
    If (Check) Then
    Else
        Label4.Caption = " DEVICE OFF "                 ' DEVICE OFF
        For i = 0 To 8
            Shape1(i).FillColor = &H8000000F
        Next i
    End If
    
    USBHID.CloseHIDDevice                               ' 關閉HID裝置
    isConnect = Check
End Function

Private Sub Command9_Click()
    a = 1
    B = 0
    C = 0
    D = 0
End Sub

Private Sub Form_load()
    numberDataInit
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
        timeArray(3 * i) = hTOd(timeStringArray(i)) \ 16      '0->1  1->8
        timeArray(3 * i + 1) = hTOd(timeStringArray(i)) Mod 16
        timeArray(3 * i + 2) = &HA ' :
    Next i
    
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)
    If (Check) Then
        Send(0) = &H22                                               ' 前導碼
        USBHID.HIDSetReport Send                                      ' 傳輸HID資
    End If
    
    
'--------------------------------------------------------------------------LED
    If (a = 1) Then
        C = C + 1
        If (C Mod 10 = 0) Then
            Send(0) = &H21                                     ' 前導碼
            Send(1) = B                                    ' LED資料
            If (Shape1(D).FillColor = &H0&) Then
                Shape1(D).FillColor = &HFF&
            End If
            USBHID.HIDSetReport Send                           ' 傳輸HID資
            If (B > 254) Then
                a = 2
                B = 1
                D = 1
            Else
                B = B * 2 + 1
                D = D + 1
            End If
        End If
    ElseIf (a = 2) Then
        C = C + 1
        If (C Mod 10 = 0) Then
            Send(0) = &H21                                     ' 前導碼
            Send(1) = 255 - B                                  ' LED資料
            If (Shape1(D).FillColor = &HFF&) Then
                Shape1(D).FillColor = &H0&
            End If
            USBHID.HIDSetReport Send                           ' 傳輸HID資
            If (B > 254) Then
                a = 0
            Else
                B = B * 2 + 1
                D = D + 1
            End If
        End If
    Else
        C = C + 1
        For i = 0 To 8
            Shape1(i).FillColor = &H0
        Next i
        If (C Mod 10 = 0) Then
            Send(0) = &H21                                     ' 前導碼
            Send(1) = &H0                                      ' LED資料
            USBHID.HIDSetReport Send                           ' 傳輸HID資
            C = 0
        End If
    End If



'--------------------------------------------------------------------------按鈕點矩陣
    If (keyin = 0) Then
        For i = 0 To 17
            clock(i) = &H0
        Next i
               
        For i = 18 To 24
            clock(i) = numberData(Val(Text1.Text) \ 10, i - 18)
        Next i
        
        For i = 25 To 31
            clock(i) = numberData(Val(Text1.Text) Mod 10, i - 25)
        Next i
             
    ElseIf (keyin = 1) Then
        
        For i = 0 To 13
            clock(i) = &H0
        Next i
        
        For i = 14 To 17
            clock(i) = numberData(timeArray(2), i - 14)
        Next i
               
        For i = 18 To 24
            clock(i) = numberData(timeArray(6), i - 18)
        Next i
        
        For i = 25 To 31
            clock(i) = numberData(timeArray(7), i - 25)
        Next i
            
    ElseIf (keyin = 2) Then
        
        For i = 0 To 6
            clock(i) = numberData(timeArray(0), i)
        Next i
        
        For i = 7 To 13
            clock(i) = numberData(timeArray(1), i - 7)
        Next i
        
        If (ap = True) Then
            For i = 14 To 17
                clock(i) = numberData(timeArray(2), i - 14)
            Next i
        ElseIf (ap = False) Then
            For i = 14 To 17
                clock(i) = &H0
            Next i
        End If
        
        For i = 18 To 24
            clock(i) = numberData(timeArray(3), i - 18)
        Next i
        
        For i = 25 To 31
            clock(i) = numberData(timeArray(4), i - 25)
        Next i
    Else
        For i = 0 To 31
            clock(i) = pic(i + logo)
        Next i
        
    End If

    Send(0) = &H20                                     ' 前導碼
    Send(1) = ((timeI) \ 8) Mod 4                 ' 8*8第n組
    Send(2) = (timeI) Mod 8                        ' 8*8第n行
    Send(3) = clock(timeI)
    USBHID.HIDSetReport Send
    
    ReadKeyData() = USBHID.HidGetReport                 ' 讀取HID資料
    USBHID.CloseHIDDevice                               ' 關閉HID裝置
            
    
    
    Select Case (ReadKeyData(1) And &H3)  'b0
        Case 0
            Label4.Caption = "Moving Logo"
            keyin = 3       ' Logo
        Case 1
            Label4.Caption = "Display Station Number"
            keyin = 0       ' 崗位
        Case 2
            Label4.Caption = "Display Second"
            keyin = 1       ' 秒鐘
        Case 3
            Label4.Caption = "Display Current Time"
            keyin = 2       ' 沒按按鈕 顯示時間
            
    End Select
    
    timeI = timeI + 1
    If (timeI >= 32) Then
        timeI = 0
    End If
    
    ttt = ttt + 1
    If (ttt > 25) Then
        ttt = 0
        ap = Not ap
    End If
        
    isConnect
End Sub


Private Sub Timer2_Timer()
    If (x = 0) Then
        logo = logo + 1
        If (logo > 31) Then
            x = 1
        End If
    End If
    If (x = 1) Then
        logo = logo - 1
        If (logo < 1) Then
            x = 0
        End If
    End If
End Sub
