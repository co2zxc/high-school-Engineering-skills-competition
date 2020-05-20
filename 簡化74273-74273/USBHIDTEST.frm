VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "USB HID 簡易控制模組 2011-09-28"
   ClientHeight    =   5250
   ClientLeft      =   3480
   ClientTop       =   2565
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   4560
   End
   Begin VB.Label Label8 
      Caption         =   "*Press<Ctrl+Q> --- Return to Windows."
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label7 
      Caption         =   "***Please press F1,F2,F3,F4 or Ctrl+Q key ----->***"
      Height          =   975
      Left            =   1560
      TabIndex        =   6
      Top             =   3480
      Width           =   7095
   End
   Begin VB.Label Label6 
      Caption         =   "* Press<F4> --- Display as Fig. 6."
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   7575
   End
   Begin VB.Label Label5 
      Caption         =   "* Press<F3> --- Display as Fig. 5."
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   6855
   End
   Begin VB.Label Label4 
      Caption         =   "* Press<F2> --- Display as Fig. 4."
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Label Label3 
      Caption         =   "* Press<F1> --- Display as Fig. 3."
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "The program is running ... Press F1,F2,F3,F4 or Ctrl+Q Key to select function:"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   8055
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'USBHID.DLL
'有四個api
'opendevice : 開始與HID裝置通訊
'Writedevice : 傳送資料到HID裝置
'Readdevice : 從HID裝置接收資料
'closedevice : 結束與HID裝置通訊

Public hid As New USB_HID

Const MyVendorID = &H7777       'USB介面卡VID
Const MyProductID = &H1234      'USB介面卡PID
Dim send(7) As Byte
Dim LoopF1, LoopF2, LoopF3, LoopF4 As Boolean
Dim Address, Data, NumberF1, NumberF2, NumberF3, NumberF4, A As Integer
Dim Number(16), DFF(8), LED(13), DFFF2(13), LEDF2(13), LEDF3(2), LEDF4(4) As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF1 Then
LoopF1 = False
LoopF2 = True
LoopF3 = True
LoopF4 = True
hiddevice = hid.opendevice(MyVendorID, MyProductID) ' 開啟 HID 裝
    
Do

    DoEvents
    
 
        '七段輸出
        send(1) = &H2
        send(0) = DFF(NumberF1)
        hid.Writedevice send()
                        
        send(1) = &H0
        hid.Writedevice send()
                
        send(0) = &H40
        send(1) = &H1
        hid.Writedevice send()
    '-----------------------------------
    'LED部分
 
    
    send(1) = &H0
    hid.Writedevice send()
            
    send(0) = LED(NumberF1)
    send(1) = &H3
    hid.Writedevice send()
    
    delay (1)
    
    send(0) = &H0
    hid.Writedevice send()
    
NumberF1 = NumberF1 + 1
If NumberF1 = 8 Then NumberF1 = 0
    
Loop Until LoopF1

End If






If KeyCode = vbKeyF2 Then
LoopF1 = True
LoopF2 = False
LoopF3 = True
LoopF4 = True
hiddevice = hid.opendevice(MyVendorID, MyProductID) ' 開啟 HID 裝
    
Do
    DoEvents
    
        send(1) = &H2
        send(0) = DFFF2(A)
        hid.Writedevice send()
                        
        send(1) = &H0
        hid.Writedevice send()
                
        send(0) = Number(NumberF2)
        send(1) = &H1
        hid.Writedevice send()
    '-----------------------------------
    'LED部分
 
    
    send(1) = &H0
    hid.Writedevice send()
            
    send(0) = LEDF2(A)
    send(1) = &H3
    hid.Writedevice send()
    
    delay (1)
    
    send(0) = &H0
    hid.Writedevice send()
    
A = A + 1
If A = 14 Then
A = 0
NumberF2 = NumberF2 + 1
If NumberF2 = 10 Then
NumberF2 = 0
End If
End If
    
Loop Until LoopF2
End If






If KeyCode = vbKeyF3 Then
LoopF1 = True
LoopF2 = True
LoopF3 = False
LoopF4 = True
hiddevice = hid.opendevice(MyVendorID, MyProductID) ' 開啟 HID 裝置
Do

        DoEvents
        '-------------------------------
        '七段控制參數
        Select Case NumberF3
            Case 0
                Address = DFF(0)
                Data = Number(1)
            Case 1
                Address = DFF(1)
                Data = Number(0)
            Case 2
                Address = DFF(2)
                Data = Number(0)
            Case 3
                Address = DFF(3)
                Data = &H40
            Case 4
                Address = DFF(4)
                Data = Number(1)
            Case 5
                Address = DFF(5)
                Data = Number(1)
            Case 6
                Address = DFF(6)
                Data = Number(3)
            Case 7
                Address = DFF(7)
                Data = Number(0)
        End Select
        '-----------------------------------
        '七段輸出
        send(1) = &H2
        send(0) = Address
        hid.Writedevice send()
                        
        send(1) = &H0
        hid.Writedevice send()
                
        send(0) = Data
        send(1) = &H1
        hid.Writedevice send()
    '-----------------------------------
    'LED部分
    
    send(1) = &H0
    hid.Writedevice send()
            
    send(0) = LEDF3(NumberF3 Mod 2)
    send(1) = &H3
    hid.Writedevice send()
            
    delay (0.5)
            
    send(0) = &H0
    hid.Writedevice send()

NumberF3 = NumberF3 + 1
If NumberF3 = 8 Then
NumberF3 = 0
End If

Loop Until LoopF3
End If









If KeyCode = vbKeyF4 Then

LoopF1 = True
LoopF2 = True
LoopF3 = True
LoopF4 = False
hiddevice = hid.opendevice(MyVendorID, MyProductID) ' 開啟 HID 裝置
Do

        DoEvents
        '-------------------------------
        '七段控制參數
        Select Case NumberF4
            Case 0
                Address = DFF(0)
                Data = Number(Hour(Time) \ 10)
            Case 1
                Address = DFF(1)
                Data = Number(Hour(Time) Mod 10)
            Case 2
                Address = DFF(2)
                Data = &H40
            Case 3
                Address = DFF(3)
                Data = Number(Minute(Time) \ 10)
            Case 4
                Address = DFF(4)
                Data = Number(Minute(Time) Mod 10)
            Case 5
                Address = DFF(5)
                Data = &H40
            Case 6
                Address = DFF(6)
                Data = Number(Second(Time) \ 10)
            Case 7
                Address = DFF(7)
                Data = Number(Second(Time) Mod 10)
        End Select
        '-----------------------------------
        '七段輸出
        send(1) = &H2
        send(0) = Address
        hid.Writedevice send()
                        
        send(1) = &H0
        hid.Writedevice send()
                
        send(0) = Data
        send(1) = &H1
        hid.Writedevice send()
    '-----------------------------------
    'LED部分
    
    send(1) = &H0
    hid.Writedevice send()
            
    send(0) = LEDF4(NumberF4 Mod 4)
    send(1) = &H3
    hid.Writedevice send()
            
    delay (0.5)
            
    send(0) = &H0
    hid.Writedevice send()

NumberF4 = NumberF4 + 1
If NumberF4 = 8 Then
NumberF4 = 0

End If

Loop Until LoopF4


End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 17 Then
LoopEnd = True                 ' 停止清空後結束程式
send(0) = &H0
send(1) = &H0
hid.Writedevice send()
hid.closedevice
End
End If
End Sub

Private Sub Form_Load()
Number(0) = &H3F    '七段數字換算
Number(1) = &H6
Number(2) = &H5B
Number(3) = &H4F
Number(4) = &H66
Number(5) = &H6D
Number(6) = &H7D
Number(7) = &H27
Number(8) = &H7F
Number(9) = &H67
Number(10) = &H77
Number(11) = &H7C
Number(12) = &H58
Number(13) = &H5E
Number(14) = &H79
Number(15) = &H71

DFF(4) = &HF7       '七段address
DFF(5) = &HFB
DFF(6) = &HFD
DFF(7) = &HFE
DFF(0) = &H7F
DFF(1) = &HBF
DFF(2) = &HDF
DFF(3) = &HEF

LED(0) = &H80       'LED亮點位置
LED(1) = &H40
LED(2) = &H20
LED(3) = &H10
LED(4) = &H8
LED(5) = &H4
LED(6) = &H2
LED(7) = &H1

DFFF2(4) = &HF7       '七段address
DFFF2(5) = &HFB
DFFF2(6) = &HFD
DFFF2(7) = &HFE
DFFF2(0) = &H7F
DFFF2(1) = &HBF
DFFF2(2) = &HDF
DFFF2(3) = &HEF
DFFF2(8) = &HFD
DFFF2(9) = &HFB
DFFF2(10) = &HF7
DFFF2(13) = &HBF
DFFF2(12) = &HDF
DFFF2(11) = &HEF

LEDF2(0) = &H80       'LED亮點位置
LEDF2(1) = &H40
LEDF2(2) = &H20
LEDF2(3) = &H10
LEDF2(4) = &H8
LEDF2(5) = &H4
LEDF2(6) = &H2
LEDF2(7) = &H1
LEDF2(0) = &H80       'LED亮點位置
LEDF2(13) = &H40
LEDF2(12) = &H20
LEDF2(11) = &H10
LEDF2(10) = &H8
LEDF2(9) = &H4
LEDF2(8) = &H2

LEDF3(0) = &HAA
LEDF3(1) = &H55

LEDF4(0) = &H88
LEDF4(1) = &H44
LEDF4(2) = &H22
LEDF4(3) = &H11


End Sub



Private Sub Timer1_Timer()
Label1.Caption = "Current time --- " & Hour(Time) & " : " & Minute(Time) & " : " & Second(Time)
End Sub
 Private Sub delay(tm)
 t = Timer
 Do
   my = DoEvents()
 Loop Until Timer - t > tm
 End Sub
