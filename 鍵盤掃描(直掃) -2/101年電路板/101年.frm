VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6585
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   3360
   End
   Begin VB.Timer Timer7 
      Interval        =   80
      Left            =   1800
      Top             =   3360
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   3360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ft As String
Private Declare Function GetKeyState Lib "user32" (ByVal key As Long) As Integer
Dim a, b, c, d, e As Integer
Dim LED As Integer
Dim number(10, 8) As Byte
Dim d1, d2, d3, d4, d5, d6
Dim es  As Integer

Private Sub Form_Activate()
ft = "The program is running ... Press A,B,C,D,E or Ctrl+Alt+Q key to selece function:" & vbNewLine
ft = ft & "* Press<Q>---Display as Fig. 3." & vbNewLine
ft = ft & "* Press<W>---Display as Fig. 4." & vbNewLine
ft = ft & "* Press<E>---Display as Fig. 5." & vbNewLine
ft = ft & "* Press<R>---Display as Fig. 6." & vbNewLine
ft = ft & "* Press<T>---Display the Square Wave" & vbNewLine & vbNewLine
ft = ft & vbTab & vbTab & "* Press<Ctrl+Alt+Q>---Return to Windows." & vbNewLine


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
Print ft


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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then

If KeyCode = vbKeyQ < 0 Then
   a = 1
   b = 0
   c = 0
   d = 0
   e = 0
   Timer3.Enabled = True
   Timer7.Enabled = False
   
    For i = 0 To 7
        OutDataEightByte &H20, 0, i, 0, 0, 0, 0, 0
        OutDataEightByte &H20, 1, i, 0, 0, 0, 0, 0
        OutDataEightByte &H20, 2, i, 0, 0, 0, 0, 0
        OutDataEightByte &H20, 3, i, 0, 0, 0, 0, 0
    Next i
    
    LED = 255
    
    OutDataEightByte &H21, 41, 0, 0, 0, 0, 0, 0
    
    OutDataEightByte &H25, es, 0, 0, 0, 0, 0, 0
    
 ElseIf KeyCode = vbKeyW < 0 Then
   a = 0
   b = 1
   c = 0
   d = 0
   e = 0
   Timer3.Enabled = False
   Timer7.Enabled = False
   
            ' f2
                     
            For i = 0 To 7
            'OutDataEightByte &H20, 0, i, number(10, i), 0, 0, 0, 0   '特殊自型
            OutDataEightByte &H20, 1, i, number(2, i), 0, 0, 0, 0
            OutDataEightByte &H20, 2, i, number(5, i), 0, 0, 0, 0
           ' OutDataEightByte &H20, 3, i, number(10, i), 0, 0, 0, 0
            Next i
              
            OutDataEightByte &H21, 0, 0, 0, 0, 0, 0, 0
            
            OutDataEightByte &H25, es, 0, 0, 0, 0, 0, 0
        
    
  ElseIf KeyCode = vbKeyE < 0 Then
     a = 0
     b = 0
     c = 1
     d = 0
     e = 0
     Timer3.Enabled = False
     Timer7.Enabled = False
     
            'f3
            
            For i = 0 To 7
            OutDataEightByte &H20, 0, i, number(1, i), 0, 0, 0, 0
            OutDataEightByte &H20, 1, i, number(1, i), 0, 0, 0, 0
            OutDataEightByte &H20, 2, i, number(2, i), 0, 0, 0, 0
            OutDataEightByte &H20, 3, i, number(8, i), 0, 0, 0, 0
            Next i
              
            OutDataEightByte &H21, 0, 0, 0, 0, 0, 0, 0
            LED = 255
            
            OutDataEightByte &H25, es, 0, 0, 0, 0, 0, 0
    
  
   ElseIf KeyCode = vbKeyR < 0 Then
      a = 0
      b = 0
      c = 0
      d = 1
      e = 0
      Timer3.Enabled = False
      Timer7.Enabled = False
      
            For i = 0 To 7
            OutDataEightByte &H20, 0, i, number(d3, i), 0, 0, 0, 0
            OutDataEightByte &H20, 1, i, number(d4, i), 0, 0, 0, 0
            OutDataEightByte &H20, 2, i, number(d5, i), 0, 0, 0, 0
            OutDataEightByte &H20, 3, i, number(d6, i), 0, 0, 0, 0
            Next i
      
              OutDataEightByte &H21, 0, 0, 0, 0, 0, 0, 0
              
              OutDataEightByte &H25, 0, 0, 0, 0, 0, 0, 0
   
    ElseIf KeyCode = vbKeyT < 0 Then
       a = 0
       b = 0
       c = 0
       d = 0
       e = 1
       Timer3.Enabled = False
       Timer7.Enabled = True
    
       For i = 0 To 7
            OutDataEightByte &H20, 0, i, 0, 0, 0, 0, 0
            OutDataEightByte &H20, 1, i, 0, 0, 0, 0, 0
            OutDataEightByte &H20, 2, i, 0, 0, 0, 0, 0
            OutDataEightByte &H20, 3, i, 0, 0, 0, 0, 0
        Next i
        OutDataEightByte &H21, 0, 0, 0, 0, 0, 0, 0
    
     ElseIf GetKeyState(27) < 0 Then   'Esc
        a = 0
        b = 0
        c = 0
        d = 0
        e = 0
        Timer3.Enabled = True
        Timer7.Enabled = False
       
        
        For i = 0 To 7
          OutDataEightByte &H20, 0, i, 0, 0, 0, 0, 0
          OutDataEightByte &H20, 1, i, 0, 0, 0, 0, 0
          OutDataEightByte &H20, 2, i, 0, 0, 0, 0, 0
          OutDataEightByte &H20, 3, i, 0, 0, 0, 0, 0
         Next i
    
         OutDataEightByte &H21, 0, 0, 0, 0, 0, 0, 0
         
         OutDataEightByte &H25, 0, 0, 0, 0, 0, 0, 0
                  
End If

     End If
     CloseUsbDevice
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

'f4
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

Private Sub Timer2_Timer()
  Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then

If GetKeyState(162) < 0 Then   'Ctrl
        If GetKeyState(164) < 0 Then     'Alt
         If GetKeyState(vbKeyQ) < 0 Then
           End
         End If
        End If
End If

     End If
     CloseUsbDevice
     
End Sub


Private Sub Timer3_Timer()
'f1
If a = 1 Then
  Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
       LED = (LED + 255) Mod (255 * 2)
       OutDataEightByte &H21, 41 And LED, 0, 0, 0, 0, 0, 0
    End If
    CloseUsbDevice
 End If
End Sub

Private Sub Timer7_Timer()
'f5
If e = 1 Then
  Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
    es = (es + 8) Mod 16
        OutDataEightByte &H25, es, 0, 0, 0, 0, 0, 0
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
