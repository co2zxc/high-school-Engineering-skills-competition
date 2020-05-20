VERSION 5.00
Begin VB.Form USB 
   Caption         =   "102學年度 工業類科學生技藝競賽 電腦修護職種 第二站 崗位號碼: 49"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   10455
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Text            =   "49"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Red LED"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   120
      Top             =   840
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00404040&
      FillStyle       =   0  '實心
      Height          =   615
      Index           =   0
      Left            =   8160
      Shape           =   3  '圓形
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   615
      Index           =   1
      Left            =   7080
      Shape           =   3  '圓形
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   615
      Index           =   2
      Left            =   6120
      Shape           =   3  '圓形
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   615
      Index           =   3
      Left            =   5040
      Shape           =   3  '圓形
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   615
      Index           =   4
      Left            =   4080
      Shape           =   3  '圓形
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   615
      Index           =   5
      Left            =   3000
      Shape           =   3  '圓形
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   615
      Index           =   6
      Left            =   2040
      Shape           =   3  '圓形
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   615
      Index           =   7
      Left            =   1080
      Shape           =   3  '圓形
      Top             =   2520
      Width           =   735
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
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Caption         =   "Station Number"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   " Display Current Time"
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
      TabIndex        =   0
      Top             =   4080
      Width           =   3045
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VB_first_Connect As Byte
Dim bit_Led_count As Integer
Dim bit_led As Byte
Dim Maxtrix_Led_count As Integer
Dim one_second As Byte
Dim LedData(15) As Byte
Dim t1 As String


Private Sub Computer_Led(ByVal bit)
    For i = 0 To 7
        Select Case (bit Mod 2)
            Case 0
                 Shape1(i).FillColor = &H0   'Black
            Case 1
             Shape1(i).FillColor = &HFF  'Red
        End Select
        bit = bit \ 2
    Next i
End Sub

Private Sub Led_Clear()
    'Mxtiix_Led
    OutDataEightByte &H28, 11, 11, 11, 11, 0, 0, 0
    '8bit_Led
     OutDataEightByte &H21, 0, 0, 0, 0, 0, 0, 0
End Sub


Private Sub Command1_Click()
    bit_led = 0
End Sub

Private Sub Command2_Click()
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        Call Led_Clear
        OutDataEightByte &H27, 2, 0, 0, 0, 0, 0, 0
        End
    End If
    CloseUsbDevice
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        Call Led_Clear
        OutDataEightByte &H27, 2, 0, 0, 0, 0, 0, 0
    End If
    CloseUsbDevice
End Sub

Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
        If VB_first_Connect = 0 Then
            OutDataEightByte &H27, 1, 0, 0, 0, 0, 0, 0      'VB_OPEN
            Call Led_Clear
            VB_first_Connect = 1
        End If
    Else
       Label1.Caption = " Device Off"
        VB_first_Connect = 0
    End If
    CloseUsbDevice                                      ' 關閉HID裝置
End Function

Private Sub Form_load()
LedData(0) = &H1
LedData(1) = &H3
LedData(2) = &H7
LedData(3) = &HF
LedData(4) = &H1F
LedData(5) = &H3F
LedData(6) = &H7F
LedData(7) = &HFF
LedData(8) = &HFE
LedData(9) = &HFC
LedData(10) = &HF8
LedData(11) = &HF0
LedData(12) = &HE0
LedData(13) = &HC0
LedData(14) = &H80
LedData(15) = 0
bit_led = UBound(LedData) + 1
one_second = 4
End Sub

Private Sub Text1_Change()
 
    Maxtrix_Led_count = 0
        
        If Len(Text1.Text) > 2 Then
            Text1.Text = Mid(Text1.Text, 3, 1) & Mid(Text1.Text, 1, 1)
        ElseIf Len(Text1.Text) < 2 Then
            Text1.Text = "0" & Text1.Text
        End If
        For i = 1 To Len(Text1.Text)
            If Mid(Text1.Text, i, 1) < "0" Or Mid(Text1.Text, i, 1) > "9" Then
                Text1.Text = t1
            ElseIf i = 2 Then
                t1 = Text1.Text
                Maxtrix_Led_count = one_second - 1
            End If
        Next i

End Sub

Private Sub Timer1_Timer()
    Dim i As Byte
    Dim timeStringArray() As String
    Dim timeArray(5) As Byte
    Dim ReadKeyData(8) As Byte
    Dim B1_B0 As Byte
    Dim t1 As Byte '崗位
    Dim t2 As Byte
        
    Label2.Caption = vbNewLine & "Current Time : " & Time$
    Call Timer_string(timeStringArray, timeArray)
    Call isConnect
    Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
        ReadData ReadKeyData  'RXD
        B1_B0 = ReadKeyData(2) * 10 + ReadKeyData(1)
        Select Case (B1_B0)
            Case 10 'OFF ON
                Label1.Caption = " Display Second"
                
                Maxtrix_Led_count = Maxtrix_Led_count + 1
                If Maxtrix_Led_count = one_second / 2 Then
                    OutDataEightByte &H28, 11, 11, timeArray(4), timeArray(5), 1, 0, 0   '11是空白  第六碼 冒號
                ElseIf Maxtrix_Led_count = one_second Then
                    OutDataEightByte &H28, 11, 11, timeArray(4), timeArray(5), 0, 0, 0  '11是空白  第六碼 冒號
                End If
                Maxtrix_Led_count = Maxtrix_Led_count Mod one_second
            Case 1  'ON OFF
                Label1.Caption = " Display Station Number"
                
                Maxtrix_Led_count = Maxtrix_Led_count + 1
                If Maxtrix_Led_count = one_second Then
                    t1 = Mid(Text1.Text, 1, 1)
                    t2 = Mid(Text1.Text, 2, 1)
                    OutDataEightByte &H28, 11, 11, t1, t2, 0, 0, 0   '11是空白  第六碼 冒號
                End If
                Maxtrix_Led_count = Maxtrix_Led_count Mod one_second
            Case 0  'ON ON
                 Label1.Caption = " Display Moving Logo"
            Case 11 'OFF OFF
                 Label1.Caption = " Display Current Time"
                 
                 Maxtrix_Led_count = Maxtrix_Led_count + 1
                 If Maxtrix_Led_count = one_second / 2 Then
                     OutDataEightByte &H28, timeArray(0), timeArray(1), timeArray(2), timeArray(3), 1, 0, 0 '11是空白  第六碼 冒號
                 ElseIf Maxtrix_Led_count = one_second Then
                     OutDataEightByte &H28, timeArray(0), timeArray(1), timeArray(2), timeArray(3), 0, 0, 0 '11是空白  第六碼 冒號
                 End If
                Maxtrix_Led_count = Maxtrix_Led_count Mod one_second
        End Select

        '8bit_led..........
        bit_Led_count = bit_Led_count + 1
        If bit_Led_count = 15 Then
            If (bit_led <= UBound(LedData)) Then
                OutDataEightByte &H21, LedData(bit_led), 0, 0, 0, 0, 0, 0
                Call Computer_Led(LedData(bit_led))
                bit_led = bit_led + 1
            End If
        End If
        bit_Led_count = bit_Led_count Mod 15
        '..........8bit_led
               
    End If
    CloseUsbDevice  ' 關閉HID裝置
End Sub

Private Sub Timer_string(ByRef timeStringArray() As String, ByRef timeArray() As Byte)

    timeStringArray = Split(Time$, ":", 3) 'HH MM SS
    
    timeArray(0) = Mid(timeStringArray(0), 1, 1) 'HH
    timeArray(1) = Mid(timeStringArray(0), 2, 1)
    
    timeArray(2) = Mid(timeStringArray(1), 1, 1) 'MM
    timeArray(3) = Mid(timeStringArray(1), 2, 1)
    
    timeArray(4) = Mid(timeStringArray(2), 1, 1) 'SS
    timeArray(5) = Mid(timeStringArray(2), 2, 1)
End Sub

