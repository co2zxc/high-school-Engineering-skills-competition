VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   30
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   13320
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   480
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   600
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   500
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
      Left            =   3240
      TabIndex        =   2
      Text            =   "6"
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
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
      Index           =   0
      Left            =   7440
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox t1 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   48
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   9975
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
      Left            =   1560
      TabIndex        =   4
      Top             =   3120
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
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim xy As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal key As Long) As Integer


Private Sub Command1_Click(Index As Integer)
 Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
        
        OutDataEightByte 0, &H6, 0, 0, 0, 0, 0, 0        'VB_OPEN
        End
    
    End If
    CloseUsbDevice
End
End Sub


Private Sub Form_Unload(Cancel As Integer)
 Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
        
        OutDataEightByte 0, &H6, 0, 0, 0, 0, 0, 0        'VB_OPEN
        End
    
    End If
    CloseUsbDevice
End
End Sub



Private Sub t1_Change()
    t1.Text = "Current Time:" & Time$
End Sub

Private Sub Timer1_Timer()
Dim Check As Boolean

t1.Text = "Current Time:" & Time$
If t2.Text > 99 Then t2.Text = 99
Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
                ReadData ReadKeyData
                If (ReadKeyData(1) = &H0) Then
                    
                    OutDataEightByte 0, &H4, (Hour(Now) \ 10), (Hour(Now) Mod 10), (Minute(Now) \ 10), (Minute(Now) Mod 10), (Second(Now) \ 10), (Second(Now) Mod 10)
                    Label2.Caption = " DisplayCurrent Time"
                 
                    ElseIf (ReadKeyData(1) = &H1) Then
                   OutDataEightByte 0, &H5, (Hour(Now) \ 10), (Hour(Now) Mod 10), (Minute(Now) \ 10), (Minute(Now) Mod 10), (Second(Now) \ 10), (Second(Now) Mod 10)
                   Label2.Caption = " Display Second"
                    ElseIf (ReadKeyData(1) = &H3) Then
                    Timer1.Enabled = False
                    Timer2.Enabled = True
                   ElseIf (ReadKeyData(1) = &H2) Then
                   OutDataEightByte 0, &H7, (t2.Text), 0, 0, 0, 0, 0
                   Label2.Caption = " Display Second"
                   End If

    Else
       
       Label2.Caption = " Device Off"
       
    End If
    CloseUsbDevice
   
End Sub


Private Sub Timer2_Timer()
Dim Check As Boolean
t1.Text = "Current Time:" & Time$
Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
    
    Select Case xy
    
    Case 1
    OutDataEightByte 0, &H9, (t2.Text), 0, 0, 0, 0, 0
    
    Case 2
    OutDataEightByte 0, &HA, 0, 0, 0, 0, 0, 0
    
    Case 3
    OutDataEightByte 0, &HB, (Hour(Now) \ 10), (Hour(Now) Mod 10), (Minute(Now) \ 10), (Minute(Now) Mod 10), (Second(Now) \ 10), (Second(Now) Mod 10)
    Case 4
    OutDataEightByte 0, &HC, 0, 0, 0, 0, 0, 0
     End
    Case 5
    OutDataEightByte 0, &HD, 0, 0, 0, 0, 0, 0
   
    End Select
    Else
    Timer1.Enabled = True
     Timer2.Enabled = False
     
    End If
    CloseUsbDevice
End Sub

Private Sub Timer3_Timer()

    Dim Check As Boolean
    t1.Text = "Current Time:" & Time$
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
    If GetKeyState(vbKeyF1) < 0 Then
        Timer1.Enabled = False
        Timer2.Enabled = True
        xy = 1
     ElseIf GetKeyState(27) < 0 Then
        Timer1.Enabled = True
        Timer2.Enabled = False
        xy = 0
    ElseIf GetKeyState(vbKeyF2) < 0 Then
        Timer1.Enabled = False
        Timer2.Enabled = True
        xy = 2
    ElseIf GetKeyState(vbKeyF3) < 0 Then
        Timer1.Enabled = False
        Timer2.Enabled = True
        xy = 3
   ElseIf GetKeyState(vbKeyF4) < 0 Then
        Timer1.Enabled = False
        Timer2.Enabled = True
        xy = 5
        
   ElseIf GetKeyState(164) < 0 And GetKeyState(162) < 0 And GetKeyState(vbKeyQ) < 0 Then
        Timer1.Enabled = False
        Timer2.Enabled = True
        xy = 4

End If
    CloseUsbDevice
End If


End Sub
