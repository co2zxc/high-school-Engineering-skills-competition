VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "103學年度 工業類科學生技藝競賽 電腦修護職種 第二站 崗位號碼 : 072"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   21.75
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
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   1200
   End
   Begin VB.TextBox t4 
      Alignment       =   2  '置中對齊
      Height          =   555
      Left            =   2880
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox t3 
      Alignment       =   2  '置中對齊
      Height          =   915
      Left            =   1560
      TabIndex        =   4
      Top             =   3600
      Width           =   9975
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   600
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   500
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
      Left            =   3240
      TabIndex        =   2
      Text            =   "72"
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
      Left            =   11280
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   5640
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
      Caption         =   "Key"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
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
Dim m0, m1, m2 As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal key As Long) As Integer


Private Sub Command1_Click(Index As Integer)
 Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
        
        OutDataEightByte 0, &H6, Hour(Now), Minute(Now), Second(Now), 0, 0, 0        'VB_OPEN
        End
    
    End If
    CloseUsbDevice
End
End Sub




Private Sub Form_Unload(Cancel As Integer)
 Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
        
        OutDataEightByte 0, &H6, Hour(Now), Minute(Now), Second(Now), 0, 0, 0        'VB_OPEN
        End
    
    End If
    CloseUsbDevice
End
End Sub






Private Sub Timer1_Timer()
Dim Check As Boolean

t1.Text = "Current Time : " & Time$
Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    If (Check) Then
                ReadData ReadKeyData
                If (ReadKeyData(1) = &H0) Then
                    OutDataEightByte (t2.Text), &H4, (Hour(Now) \ 10), (Hour(Now) Mod 10), (Minute(Now) \ 10), (Minute(Now) Mod 10), (Second(Now) \ 10), (Second(Now) Mod 10)
                    
                ElseIf (ReadKeyData(1) = &H1 Or ReadKeyData(1) = &H4 Or ReadKeyData(1) = &H7 Or ReadKeyData(1) = &HA) Then
                    OutDataEightByte 0, &H8, 0, 0, 0, 0, 0, 0
                    t3.Text = "Output 250 Hz Signal"
                ElseIf (ReadKeyData(1) = &H2 Or ReadKeyData(1) = &H5 Or ReadKeyData(1) = &H8 Or ReadKeyData(1) = &HB) Then
                    OutDataEightByte 0, &H9, 0, 0, 0, 0, 0, 0
                    t3.Text = "Output 500 Hz Signal"
                ElseIf (ReadKeyData(1) = &H3 Or ReadKeyData(1) = &H6 Or ReadKeyData(1) = &H9 Or ReadKeyData(1) = &HC) Then
                    OutDataEightByte 0, &HA, 0, 0, 0, 0, 0, 0
                    t3.Text = "Output 1000 Hz Signal"
                End If

    Else
       
       t3.Text = " Device Off"
       Timer4.Enabled = True
       
    End If
    CloseUsbDevice
   
End Sub


Private Sub Timer3_Timer()
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
        ReadData ReadKeyData
        If ReadKeyData(1) = &H1 Then
        t4.Text = "1"
        OutDataEightByte 0, &H5, 1, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H2 Then
        t4.Text = "2"
        OutDataEightByte 0, &H5, 2, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H3 Then
        t4.Text = "3"
        OutDataEightByte 0, &H5, 3, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H4 Then
        t4.Text = "4"
        OutDataEightByte 0, &H5, 4, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H5 Then
        t4.Text = "5"
        OutDataEightByte 0, &H5, 5, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H6 Then
        t4.Text = "6"
        OutDataEightByte 0, &H5, 6, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H7 Then
        t4.Text = "7"
        OutDataEightByte 0, &H5, 7, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H8 Then
        t4.Text = "8"
        OutDataEightByte 0, &H5, 8, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H9 Then
        t4.Text = "9"
        OutDataEightByte 0, &H5, 9, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &HA Then
        t4.Text = "*"
        OutDataEightByte 0, &H5, 10, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &HB Then
        t4.Text = "0"
        OutDataEightByte 0, &H5, 0, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &HC Then
        t4.Text = "#"
        OutDataEightByte 0, &H5, 11, 0, 0, 0, 0, 0
        End If
        
    End If
End Sub

Private Sub Timer4_Timer()
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
    t3.Text = "Output 250 Hz Signal"
    Timer4.Enabled = False
    End If
End Sub
