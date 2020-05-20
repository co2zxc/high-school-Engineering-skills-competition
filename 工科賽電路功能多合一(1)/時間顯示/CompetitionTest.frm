VERSION 5.00
Begin VB.Form USB 
   Caption         =   "CompetitionTest"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   12285
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   120
      Top             =   1920
   End
   Begin VB.Timer Timer2 
      Interval        =   1150
      Left            =   120
      Top             =   1320
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
      Left            =   2640
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   1725
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   720
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
      TabIndex        =   5
      Top             =   720
      Width           =   6495
      WordWrap        =   -1  'True
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
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
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   " KEY OFF "
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
      Left            =   2220
      TabIndex        =   1
      Top             =   2640
      Width           =   1545
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
      Top             =   2640
      Width           =   2085
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Data(4, 8) As Byte
Dim numberData(11, 5) As Byte
Dim A As Byte
Dim B As Byte
Dim C As Byte
Dim d1, d2, d3, d4, d5, d6
Dim LedData(20) As Byte
Dim ReadKeyData(8) As Byte

Private Sub Form_Unload(Cancel As Integer)
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            OutDataEightByte &H20, 0, i, 0, 0, 0, 0, 0
            OutDataEightByte &H20, 1, i, 0, 0, 0, 0, 0
            OutDataEightByte &H20, 2, i, 0, 0, 0, 0, 0
            OutDataEightByte &H20, 3, i, 0, 0, 0, 0, 0
        Next i
    End If
    OutDataEightByte &H21, 0, 0, 0, 0, 0, 0, 0
    CloseUsbDevice
End Sub

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

Private Sub Command9_Click()
    A = 1
    B = 0
    C = 0
End Sub

Private Sub Form_load()
'number 0
numberData(0, 0) = &H7F
numberData(0, 1) = &H41
numberData(0, 2) = &H7F
'number 1
numberData(1, 0) = &H42
numberData(1, 1) = &H7F
numberData(1, 2) = &H40
'number 2
numberData(2, 0) = &H79
numberData(2, 1) = &H49
numberData(2, 2) = &H4F
'number 3
numberData(3, 0) = &H49
numberData(3, 1) = &H49
numberData(3, 2) = &H7F
'number 4
numberData(4, 0) = &HF
numberData(4, 1) = &H8
numberData(4, 2) = &H7F
'number 5
numberData(5, 0) = &H4F
numberData(5, 1) = &H49
numberData(5, 2) = &H79
'number 6
numberData(6, 0) = &H7F
numberData(6, 1) = &H49
numberData(6, 2) = &H79
'number 7
numberData(7, 0) = &H7
numberData(7, 1) = &H1
numberData(7, 2) = &H7F
'number 8
numberData(8, 0) = &H7F
numberData(8, 1) = &H49
numberData(8, 2) = &H7F
'number 9
numberData(9, 0) = &HF
numberData(9, 1) = &H9
numberData(9, 2) = &H7F
':
numberData(10, 0) = &H66


LedData(0) = &H80
LedData(1) = &H40
LedData(2) = &H20
LedData(3) = &H10
LedData(4) = &H8
LedData(5) = &H4
LedData(6) = &H2
LedData(7) = &H1
LedData(8) = &H2
LedData(9) = &H4
LedData(10) = &H8
LedData(11) = &H10
LedData(12) = &H20
LedData(13) = &H40
LedData(14) = &H80
LedData(15) = &HFF
LedData(16) = &H0
LedData(17) = &HFF
LedData(18) = &H0
LedData(19) = &HFF
LedData(20) = &H0



Check = OpenUsbDevice(VendorID, ProductID)
    Dim timeStringArray() As String
    timeStringArray = Split(Time$, ":")
   
  
    d1 = Mid(timeStringArray(0), 1, 1)
    d2 = Mid(timeStringArray(0), 2, 1)
    d3 = Mid(timeStringArray(1), 1, 1)
    d4 = Mid(timeStringArray(1), 2, 1)
    d5 = Mid(timeStringArray(2), 1, 1)
    d6 = Mid(timeStringArray(2), 2, 1)
    
    Label2.Caption = vbNewLine & "Current Time : " & d1 & d2 & ":" & d3 & d4 & ":" & d5 & d6

    If (Check) Then
    
            For i = 0 To 2
                OutDataEightByte &H20, 0, i + 2, numberData(d1, i), 0, 0, 0, 0
                OutDataEightByte &H20, 0 + (i + 6) \ 8, (i + 6) Mod 8, numberData(d2, i), 0, 0, 0, 0
                OutDataEightByte &H20, 1, i + 4, numberData(d3, i), 0, 0, 0, 0
                OutDataEightByte &H20, 2, i, numberData(d4, i), 0, 0, 0, 0
                OutDataEightByte &H20, 2 + (i + 6) \ 8, (i + 6) Mod 8, numberData(d5, i), 0, 0, 0, 0
                OutDataEightByte &H20, 3, i + 2, numberData(d6, i), 0, 0, 0, 0
            Next i
            
            OutDataEightByte &H20, 1, 2, numberData(10, 0), 0, 0, 0, 0
            OutDataEightByte &H20, 2, 4, numberData(10, 0), 0, 0, 0, 0
    End If

    
CloseUsbDevice

End Sub

Private Sub Timer1_Timer()
    isConnect
    Check = OpenUsbDevice(VendorID, ProductID)
    ReadData LedData  'RXD
    CloseUsbDevice
    Select Case LedData(1)
        Case 0
           If Label1.Caption = " DEVICE ON " Then Label6.Caption = " KEY ON "
        Case 1
            Label6.Caption = " KEY OFF "
    End Select
End Sub

Private Sub Timer2_Timer()

    d6 = d6 + 1
    
    d5 = d5 + (d6 \ 10)
    
    d4 = d4 + (d5 \ 6)
    
    d3 = d3 + (d4 \ 10)
    
    d2 = d2 + (d3 \ 6)
    
    d1 = d1 + (d2 \ 10)
        
    d6 = d6 Mod 10
    d5 = d5 Mod 6
    d4 = d4 Mod 10
    d3 = d3 Mod 6
    d2 = d2 Mod 10
    d1 = d1 Mod 24
    
    Label2.Caption = vbNewLine & "Current Time : " & d1 & d2 & ":" & d3 & d4 & ":" & d5 & d6
        
        
    Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
    
            For i = 0 To 2
                OutDataEightByte &H20, 0, i + 2, numberData(d1, i), 0, 0, 0, 0
                OutDataEightByte &H20, 0 + (i + 6) \ 8, (i + 6) Mod 8, numberData(d2, i), 0, 0, 0, 0
                OutDataEightByte &H20, 1, i + 4, numberData(d3, i), 0, 0, 0, 0
                OutDataEightByte &H20, 2, i, numberData(d4, i), 0, 0, 0, 0
                OutDataEightByte &H20, 2 + (i + 6) \ 8, (i + 6) Mod 8, numberData(d5, i), 0, 0, 0, 0
                OutDataEightByte &H20, 3, i + 2, numberData(d6, i), 0, 0, 0, 0
            Next i
                        
            OutDataEightByte &H20, 1, 2, numberData(10, 0), 0, 0, 0, 0
            OutDataEightByte &H20, 2, 4, numberData(10, 0), 0, 0, 0, 0
            
            
            If (A) Then
                 OutDataEightByte &H21, LedData(B), 0, 0, 0, 0, 0, 0
                 B = B + 1
                 If B = 21 Then A = 0
             End If
            
    End If
    CloseUsbDevice
    
End Sub

Private Sub Timer3_Timer()
    isConnect
    Check = OpenUsbDevice(VendorID, ProductID)
    ReadData ReadKeyData  'RXD
    CloseUsbDevice
    Select Case ReadKeyData(1)
        Case 0
           If Label1.Caption = " DEVICE ON " Then Label6.Caption = " KEY ON "
        Case 1
            Label6.Caption = " KEY OFF "
    End Select
End Sub
