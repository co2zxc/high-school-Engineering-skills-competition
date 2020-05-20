VERSION 5.00
Begin VB.Form USB 
   Caption         =   "CompetitionTest"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   10080
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
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
      Left            =   6480
      TabIndex        =   1
      Top             =   240
      Width           =   2085
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
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Data(4, 8) As Byte
Dim a As Byte
Dim b As Integer
Dim d As Byte
Dim ba As Byte
Dim bi(3) As Byte
Dim LedData(20) As Byte
Dim ReadKeyData(8) As Byte
Dim number(10, 8) As Byte
Dim t1 As String
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
    a = 1
    b = 0
    c = 0
End Sub

Private Sub Command14_Click()
        Dim Check As Boolean
        ckeck = OpenUsbDevice(VendorID, ProductID)
        Dim L, i, j   As Integer: L = Len(Text1.Text)
        For i = 0 To L - 1
            Dim m As Integer: m = Val(Mid(Text1.Text, i + 1, 1))
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
      For i = 0 To 63
        Command1(i).BackColor = &HC0C0FF
        Command2(i).BackColor = &HC0C0FF
        Command3(i).BackColor = &HC0C0FF
        Command4(i).BackColor = &HC0C0FF
      Next i
      d = 1
End Sub

Private Sub Form_Load()

number(0, 2) = &H1F
number(0, 3) = &H11
number(0, 4) = &H11
number(0, 5) = &H1F

number(1, 5) = &H1F

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
        OutDataEightByte &H21, 0, 0, 0, 0, 0, 0, 0
    End If
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
    Check = OpenUsbDevice(VendorID, ProductID)
    OutDataEightByte &H21, b, 0, 0, 0, 0, 0, 0
    ReadData ReadKeyData  'RXD
    Select Case ReadKeyData(1)
        Case 0      'ON
            ba = 1
        Case 1   'OFF
            If ba = 1 Then
                ba = 0
                b = b - 1
                If b = -1 Then b = 255
            End If
    End Select
    CloseUsbDevice
End Sub
