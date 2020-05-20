VERSION 5.00
Begin VB.Form USB 
   Caption         =   "CompetitionTest"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   11505
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   0
      Top             =   2520
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   0
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   240
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000FF00&
      Caption         =   "送出跑馬燈數字"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "案鍵盤 上下左右 鍵 改變跑馬燈方向"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      TabIndex        =   5
      Top             =   240
      Width           =   2775
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
      Left            =   840
      TabIndex        =   2
      Top             =   1800
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
      Left            =   3360
      TabIndex        =   1
      Top             =   1800
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
      Left            =   5280
      TabIndex        =   0
      Top             =   1800
      Width           =   2085
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Dim b As Double
Dim c As Double
Dim LedData(20) As Byte
Dim ReadKeyData(8) As Byte
Dim number(10, 8) As Byte
Dim t1 As String
Dim horse() As Byte
Dim ch  As Byte
Private Declare Function GetKeyState Lib "user32" (ByVal key As Long) As Integer

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


Private Sub Command14_Click()
If Text1.Text <> "" Then
    Dim Check As Boolean
    ckeck = OpenUsbDevice(VendorID, ProductID)
    Dim i, j, L, m     As Integer: L = Len(Text1.Text)
    For i = 0 To 3
        For j = 0 To 7
            OutDataEightByte &H20, i, j, &H0, 0, 0, 0, 0
        Next j
    Next i
    ReDim horse(L * 8 - 1 + 31)
    For i = 0 To L * 8 - 1 Step 8
        m = Mid(Text1.Text, i \ 8 + 1, 1)
        For j = 0 To 7
            horse(i + j) = number(m, j)
        Next j
    Next i
    m = Mid(Text1.Text, 1, 1)
    b = 2
    ch = &H22
    CloseUsbDevice
    End If
Timer1.Enabled = True
a = 0
End Sub

Private Sub Form_Load()

number(0, 2) = &H1F
number(0, 3) = &H11
number(0, 4) = &H11
number(0, 5) = &H1F

number(1, 2) = &H0
number(1, 3) = &H12
number(1, 4) = &H1F
number(1, 5) = &H10

number(2, 2) = &H1D
number(2, 3) = &H15
number(2, 4) = &H15
number(2, 5) = &H17

number(3, 2) = &H15
number(3, 3) = &H15
number(3, 4) = &H15
number(3, 5) = &H1F

number(4, 2) = &H7
number(4, 3) = &H4
number(4, 4) = &H4
number(4, 5) = &H1F

number(5, 2) = &H17
number(5, 3) = &H15
number(5, 4) = &H15
number(5, 5) = &H1D

number(6, 2) = &H1F
number(6, 3) = &H14
number(6, 4) = &H14
number(6, 5) = &H1C

number(7, 2) = &H7
number(7, 3) = &H1
number(7, 4) = &H1
number(7, 5) = &H1F

number(8, 2) = &H1F
number(8, 3) = &H15
number(8, 4) = &H15
number(8, 5) = &H1F

number(9, 2) = &H17
number(9, 3) = &H15
number(9, 4) = &H15
number(9, 5) = &H1F


LedData(0) = &H1
LedData(1) = &H2
LedData(2) = &H4
LedData(3) = &H8
LedData(4) = &H10
LedData(5) = &H20
LedData(6) = &H40
LedData(7) = &H80
LedData(8) = &H40
LedData(9) = &H20
LedData(10) = &H10
LedData(11) = &H8
LedData(12) = &H4
LedData(13) = &H2
LedData(14) = &H1
LedData(15) = &HFF
LedData(16) = &H0
LedData(17) = &HFF
LedData(18) = &H0
LedData(19) = &HFF
LedData(20) = &H0

ReDim horse(0)
horse(0) = &H0

End Sub

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

Private Sub Text1_Change()
If Text1.Text <> "" Then
    If Len(Text1.Text) > 4 Then
        Text1.Text = Mid(Text1.Text, 1, 4)
    End If
    For i = 1 To Len(Text1.Text)
        If Mid(Text1.Text, i, 1) < "0" Or Mid(Text1.Text, i, 1) > "9" Then
            Text1.Text = t1
        ElseIf i = 4 Then
        t1 = Text1.Text
     End If
    Next i
 End If
End Sub

Private Sub Timer1_Timer()
    isConnect
    '8bit LED
    Check = OpenUsbDevice(VendorID, ProductID)
    OutDataEightByte &H21, LedData(a), 0, 0, 0, 0, 0, 0
    a = a + 1
    a = a Mod (UBound(LedData) + 1)
    CloseUsbDevice
End Sub

Private Sub Timer2_Timer()
Check = OpenUsbDevice(VendorID, ProductID)
OutDataEightByte ch, horse(b), 0, 0, 0, 0, 0, 0
b = b + 1
b = b Mod (UBound(horse) + 1)
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


Private Sub Timer4_Timer()
    Dim Check As Boolean
    Dim i, j, L, m     As Integer: L = Len(Text1.Text)

  If GetKeyState(37) < 0 Then     'LEFT
        ch = &H22
        
        If Text1.Text <> "" Then

    ckeck = OpenUsbDevice(VendorID, ProductID)

    For i = 0 To 3
        For j = 0 To 7
            OutDataEightByte &H20, i, j, &H0, 0, 0, 0, 0
        Next j
    Next i
    ReDim horse(L * 8 - 1 + 31)
    For i = 0 To L * 8 - 1 Step 8
        m = Mid(Text1.Text, i \ 8 + 1, 1)
        For j = 0 To 7
            horse(i + j) = number(m, j)
        Next j
    Next i
    m = Mid(Text1.Text, 1, 1)
    b = 2
    CloseUsbDevice
    a = 0
    Timer1.Enabled = True
        End If
 End If
 
 
 
 
 
 
    
    If GetKeyState(39) < 0 Then   'Right
    
    ch = &H23
    
    If Text1.Text <> "" Then
    ckeck = OpenUsbDevice(VendorID, ProductID)
    For i = 0 To 3
        For j = 0 To 7
            OutDataEightByte &H20, i, j, &H0, 0, 0, 0, 0
        Next j
    Next i
    ReDim horse(L * 8 - 1 + 31)
    For i = 0 To L * 8 - 1 Step 8
        m = Mid(Text1.Text, L - (i \ 8), 1)
        For j = 7 To 0 Step -1
            horse(i + (8 - j)) = number(m, j)
        Next j
    Next i
    b = 2
    OutDataEightByte &H21, 0, 0, 0, 0, 0, 0, 0
    a = 0
    Timer1.Enabled = True
    CloseUsbDevice
    End If
    
    End If
        
     
    If GetKeyState(38) < 0 Then    'UP
    
    ch = &H25
    
        ckeck = OpenUsbDevice(VendorID, ProductID)
      
        For i = 0 To L - 1
           m = Val(Mid(Text1.Text, i + 1, 1))
            For j = 0 To 7
               OutDataEightByte &H20, i, j, number(m, j), 0, 0, 0, 0
            Next j
        Next i
        For i = i To 3
            For j = 0 To 7
               OutDataEightByte &H20, i, j, &H0, 0, 0, 0, 0
            Next j
        Next i
      CloseUsbDevice
    
    End If
    
    
    If GetKeyState(40) < 0 Then    'Down
    
    ch = &H26

        ckeck = OpenUsbDevice(VendorID, ProductID)

        For i = 0 To L - 1
         m = Val(Mid(Text1.Text, i + 1, 1))
            For j = 0 To 7
               OutDataEightByte &H20, i, j, number(m, j), 0, 0, 0, 0
            Next j
        Next i
        For i = i To 3
            For j = 0 To 7
               OutDataEightByte &H20, i, j, &H0, 0, 0, 0, 0
            Next j
        Next i
      CloseUsbDevice
    
    End If
End Sub
