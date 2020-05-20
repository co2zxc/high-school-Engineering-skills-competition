VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "103學年度 "
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13350
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   20.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   13350
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox t3 
      Alignment       =   2  '置中對齊
      Height          =   525
      Left            =   7920
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Left            =   600
      Top             =   1560
   End
   Begin VB.Timer Timer4 
      Interval        =   5
      Left            =   600
      Top             =   1200
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   950
      Left            =   600
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.TextBox t2 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Text            =   "057"
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton b1 
      BackColor       =   &H8000000B&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox t1 
      Alignment       =   2  '置中對齊
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   48
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   9975
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   960
      TabIndex        =   6
      Top             =   5400
      Width           =   11895
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   4920
      Width           =   7455
   End
   Begin VB.Shape R 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   7
      Left            =   3360
      Shape           =   3  '圓形
      Top             =   4320
      Width           =   495
   End
   Begin VB.Shape R 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   6
      Left            =   4005
      Shape           =   3  '圓形
      Top             =   4320
      Width           =   495
   End
   Begin VB.Shape R 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   5
      Left            =   4665
      Shape           =   3  '圓形
      Top             =   4320
      Width           =   495
   End
   Begin VB.Shape R 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   4
      Left            =   5310
      Shape           =   3  '圓形
      Top             =   4320
      Width           =   495
   End
   Begin VB.Shape R 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   3
      Left            =   5970
      Shape           =   3  '圓形
      Top             =   4320
      Width           =   495
   End
   Begin VB.Shape R 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   2
      Left            =   6615
      Shape           =   3  '圓形
      Top             =   4320
      Width           =   495
   End
   Begin VB.Shape R 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   1
      Left            =   7260
      Shape           =   3  '圓形
      Top             =   4320
      Width           =   495
   End
   Begin VB.Shape R 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   0
      Left            =   7920
      Shape           =   3  '圓形
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "Station Number"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ptbg, xx, yy As Integer 'ptbg防彈跳
Dim kbd, btn, vbtn  '鍵盤    按鈕    VB按鈕
Private Declare Function GetKeyState Lib "user32" (ByVal key As Long) As Integer
Private Sub b1_Click()
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
        
        OutDataEightByte &H6, 0, 0, 0, 0, 0, 0, 0        'VB_OPEN
        End
    
    End If
    CloseUsbDevice
End
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    t2.Enabled = True
    t3.Enabled = True
    b1.Enabled = True
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Dim Check As Boolean
Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
If (Check) Then
    t3.Text = Chr(KeyAscii)
    OutDataEightByte &HD, (KeyAscii), 0, 0, 0, 0, 0, 0        'VB_OPEN
End If
CloseUsbDevice
End Sub
Private Sub Form_Load()
 Label4 = "The program is running...Press A,B,C,D,F1 or Ctrl+Alt+Q key to select function:" & vbNewLine & "* Press <A> ---Display as Fig. 3." & vbNewLine & "* Press <B> ---Display as Fig. 4." & vbNewLine & "* Press <C> ---Display as Fig. 5." & vbNewLine & "* Press <D> ---Display as Fig. 6." & vbNewLine & "* Press <F1>---Display the Square Wave." & vbNewLine & "* Press <Ctrlt+Alt+Q> ---Return to Windows." & vbNewLine & vbNewLine & "                                    *** Press A,B,C,D,E or Ctrl+Alt+Q -----> ***"
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Dim Check As Boolean
    
    
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
        
        OutDataEightByte &H6, 0, 0, 0, 0, 0, 0, 0        'VB_OPEN
        End
    
    End If
    CloseUsbDevice
End
End Sub

Private Sub Timer1_Timer()
    Dim Check As Boolean
     
     Label3.Caption = "Current Time---" & Time$
    
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
        If (Check) Then

                Select Case kbd
                    Case 0
                        t1.Text = " "
                        OutDataEightByte &O5, 0, 0, 0, 0, 0, 0, 0
                        
                        For i = 0 To 7
                            R(i).FillColor = RGB(128, 0, 0)
                        Next i

                    Case 65
                        If xx = 0 Then
                        R(0).FillColor = RGB(255, 0, 0)
                        R(1).FillColor = RGB(255, 0, 0)
                        R(2).FillColor = RGB(255, 0, 0)
                        R(4).FillColor = RGB(255, 0, 0)
                        R(6).FillColor = RGB(255, 0, 0)
                        xx = xx + 1
                        Else
                        For i = 0 To 7
                            R(i).FillColor = RGB(128, 0, 0)
                        Next i
                        xx = 0
                        End If
                    Case 66
                        OutDataEightByte &H9, t2.Text, 0, 0, 0, 0, 0, 0
                    Case 67
                        t1.Enabled = False
                        t2.Enabled = False
                        t3.Enabled = False
                        b1.Enabled = False
                    Case 68
                        OutDataEightByte &HF, (Hour(Now) \ 10), (Hour(Now) Mod 10), (Minute(Now) \ 10), (Minute(Now) Mod 10), (Second(Now) \ 10), (Second(Now) Mod 10), 0
                    Case 256
                        OutDataEightByte &H5, 0, 0, 0, 0, 0, 0, 0
                        End
                    Case 112
                        OutDataEightByte &HC, 0, 0, 0, 0, 0, 0, 0
                End Select
        Else
            For i = 0 To 7
                R(i).FillColor = RGB(255, 255, 255)
            Next i
            Label2.Caption = " Device Off"
           
        End If
CloseUsbDevice
   
End Sub


Private Sub Timer2_Timer()
    Dim Check As Boolean
    If (Len(t2.Text) >= 1) Then
        If t2.Text > 101 Then t2.Text = 101
    End If
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
        If (Check) Then
       
    
        
    
        End If
    CloseUsbDevice
End Sub

Private Sub Timer3_Timer()

    Dim Check As Boolean
    If (Len(t2.Text) >= 1) Then
        If t2.Text > 101 Then t2.Text = 101
    End If
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
         ReadData ReadKeyData
        
        Select Case ReadKeyData(1)
            Case 0
                btn = 0
            Case 1
                btn = 1
            Case 2
                btn = 2
            Case 3
                btn = 3
            Case 4
                btn = 4
            Case 5
                btn = 5
            Case 6
                btn = 6
            Case 7
                btn = 7
            Case 8
                btn = 8
            Case 9
                btn = 9
            Case 10
                btn = 10
            Case 11
                btn = 11
            Case 12
                btn = 12

        End Select

        
        
    Else
        Timer1.Enabled = True
        Timer2.Enabled = False
        Label2.Caption = " Device Off"
    End If
CloseUsbDevice

End Sub
Private Sub Timer4_Timer()
Dim Check, a As Boolean
a = True
For i = 1 To Len(t2.Text)
    If Not IsNumeric(Mid(t2.Text, i, 1)) Then
        a = False
        t2.Text = Mid(t2.Text, 1, i - 1) & Mid(t2.Text, i + 1)
    End If
Next
If a = True And (Len(t2.Text) >= 1) Then
    If t2.Text > 101 Then
        t2.Text = 101
    End If
ElseIf Len(t2.Text) = 0 Then
    t2.Text = 0
End If
Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
        If kbd = 0 Then
            For i = 0 To 255
                If GetKeyState(i) < 0 Then kbd = i
            Next i
        End If
        If GetKeyState(162) < 0 And GetKeyState(164) < 0 And GetKeyState(81) < 0 Then
            kbd = 256
        ElseIf GetKeyState(27) < 0 Then
            kbd = 0
        End If
    Else
        Timer1.Enabled = True
        Timer2.Enabled = False
        Label2.Caption = " Device Off"
    End If
CloseUsbDevice
End Sub
