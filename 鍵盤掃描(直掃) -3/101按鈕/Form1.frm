VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7935
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "clean"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "按鈕3"
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "按鈕2"
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "按鈕1"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d1, d2, d3, d4, d5, d6
Dim Data(4, 8) As Byte
Dim number(10, 8) As Byte
Dim ft As String
Dim d
Dim LED As Integer
Dim a As Integer
Private Sub Command1_Click(Index As Integer)
    Dim i As Byte
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            OutDataEightByte &H20, 0, i, number(1, i), 0, 0, 0, 0
            OutDataEightByte &H20, 1, i, number(1, i), 0, 0, 0, 0
            OutDataEightByte &H20, 2, i, number(2, i), 0, 0, 0, 0
            OutDataEightByte &H20, 3, i, number(8, i), 0, 0, 0, 0
        Next i
    End If
    CloseUsbDevice
End Sub

Private Sub Command2_Click()
    Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
        For i = 0 To 7
            OutDataEightByte &H20, 0, i, number(d3, i), 0, 0, 0, 0
            OutDataEightByte &H20, 1, i, number(d4, i), 0, 0, 0, 0
            OutDataEightByte &H20, 2, i, number(d5, i), 0, 0, 0, 0
            OutDataEightByte &H20, 3, i, number(d6, i), 0, 0, 0, 0
        Next i
        d = 1
    End If
    CloseUsbDevice
End Sub

Private Sub Command3_Click()
    Dim i As Byte
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
      For i = 0 To 7
        OutDataEightByte &H20, 0, i, 0, 0, 0, 0, 0
        OutDataEightByte &H20, 1, i, 0, 0, 0, 0, 0
        OutDataEightByte &H20, 2, i, 0, 0, 0, 0, 0
        OutDataEightByte &H20, 3, i, 0, 0, 0, 0, 0
    Next i
        OutDataEightByte &H21, 25, 0, 0, 0, 0, 0, 0
        LED = 255
    End If
    a = 1
    CloseUsbDevice
End Sub

Private Sub Command4_Click()
    Dim i As Byte
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            Data(0, i) = 0
            OutDataEightByte &H20, 0, i, Data(0, i), 0, 0, 0, 0
            OutDataEightByte &H20, 1, i, Data(0, i), 0, 0, 0, 0
            OutDataEightByte &H20, 2, i, Data(0, i), 0, 0, 0, 0
            OutDataEightByte &H20, 3, i, Data(0, i), 0, 0, 0, 0
        Next i
    End If
    OutDataEightByte &H21, 0, 0, 0, 0, 0, 0, 0
    CloseUsbDevice
    d = 0
    a = 0
End Sub

Private Sub Form_Activate()

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

number(10, 3) = &H4F
number(10, 4) = &H60
number(10, 5) = &H76
number(10, 6) = &H60
number(10, 7) = &H4F



Dim timeStringArray() As String
 timeStringArray = Split(Time$, ":")
  
    d1 = Mid(timeStringArray(0), 1, 1)
    d2 = Mid(timeStringArray(0), 2, 1)
    d3 = Mid(timeStringArray(1), 1, 1)
    d4 = Mid(timeStringArray(1), 2, 1)
    d5 = Mid(timeStringArray(2), 1, 1)
    d6 = Mid(timeStringArray(2), 2, 1)
    
Print vbTab & vbTab & "Current time ---" & d1 & d2 & ":" & d3 & d4 & ":" & d5 & d6 & vbNewLine




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

Private Sub Timer1_Timer()
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

Cls
Print vbTab & vbTab & "Current time ---" & d1 & d2 & ":" & d3 & d4 & ":" & d5 & d6 & vbNewLine
Print ft

delay (0.1)
'd
If d = 1 Then
 Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
            For i = 0 To 7
            OutDataEightByte &H20, 0, i, number(d3, i), 0, 0, 0, 0
            OutDataEightByte &H20, 1, i, number(d4, i), 0, 0, 0, 0
            OutDataEightByte &H20, 2, i, number(d5, i), 0, 0, 0, 0
            OutDataEightByte &H20, 3, i, number(d6, i), 0, 0, 0, 0
            Next i
    End If

 CloseUsbDevice
End If
End Sub

Private Sub delay(tm)
 t = Timer
 Do
   my = DoEvents()
 Loop Until Timer - t > tm
 End Sub


Private Sub Timer2_Timer()
'f1
If a = 1 Then
  Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
       LED = (LED + 255) Mod (255 * 2)
       OutDataEightByte &H21, 25 And LED, 0, 0, 0, 0, 0, 0
    End If
    CloseUsbDevice
 End If
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
