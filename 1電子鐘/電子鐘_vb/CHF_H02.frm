VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "103學年度  工業類科學生技藝競賽  電腦修護職種   第二站  崗位號碼:98"
   ClientHeight    =   5235
   ClientLeft      =   195
   ClientTop       =   420
   ClientWidth     =   8730
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "細明體"
      Size            =   14.25
      Charset         =   136
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   -1  'True
   EndProperty
   ForeColor       =   &H000000FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleMode       =   0  'User
   ScaleWidth      =   8730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "送出"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   2880
   End
   Begin VB.TextBox Text_Station 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   4
      Text            =   "98"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   360
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   2200
   End
   Begin VB.Label Label2 
      Caption         =   "Station Number"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   953
      TabIndex        =   2
      Top             =   600
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label StatusDisplay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   960
      TabIndex        =   1
      Top             =   3720
      Width           =   3525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TxRedLED As Byte
Public TxCtrl As Byte
Public count1 As Integer
Dim temp(4) As Byte
Public count2 As Integer
Public loop_on As Boolean



Private Sub Command2_Click()
   Dim Result As Boolean
   Result = OpenUsbDevice(VendorID, ProductID)
   

   If (Len(Text_Station.Text) = 2) Then
      If (Val(Text_Station.Text) < 100) And (Val(Text_Station.Text) > 0) Then
              
         If (Result) Then
            TxCtrl = &H4
            OutDataEightByte TxRedLED, TxCtrl, Val(Text_Station.Text), 0, 0, 0, 0, 0
            CloseUsbDevice
         End If
         
      End If
   End If
End Sub

'

Private Sub Form_Load()
    a = 99
    b(0) = &H0
    b(1) = &H1
    b(2) = &H3
    b(3) = &H7
    b(4) = &HF
    b(5) = &H1F
    b(6) = &H3F
    b(7) = &H7F
    b(8) = &HFF
    b(9) = &HFE
    b(10) = &HFC
    b(11) = &HF8
    b(12) = &HF0
    b(13) = &HE0
    b(14) = &HC0
    b(15) = &H80
    c = 0
    
    count1 = 0
    temp(0) = Asc(" ")
    temp(1) = Asc(" ")
    temp(2) = Asc(" ")
    temp(3) = Asc(" ")
    
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
    Dim Result As Boolean
    Result = OpenUsbDevice(VendorID, ProductID)
    If (Result) Then
        TxCtrl = &H3
        OutDataEightByte TxRedLED, TxCtrl, 0, 0, 0, 0, 0, 0
        CloseUsbDevice
    End If
End Sub
'



Private Sub Timer1_Timer()                             ''讀取INPUT 與LED動作輸出
    Dim Result As Boolean
    Dim k As Byte
    Static PreConnect As Boolean
    If c > 20 Then c = 20
    Label1.Caption = "                                  " & _
                     "Current Time : " & Time$

    Result = OpenUsbDevice(VendorID, ProductID)
     
    If (Result) Then
       StatusDisplay.Caption = "Display Current Time"
        
        
'       ReadData ReadKeyData
       CloseUsbDevice

    Else
        StatusDisplay.Caption = " DEVICE OFF "
        If a = 1 Then
           PreConnect = Result
           End
        End If
    End If
    PreConnect = Result
    c = c + 1

End Sub

Private Sub Timer2_Timer()                   '
    Dim Result As Boolean
    
    Result = OpenUsbDevice(VendorID, ProductID)
    If (Result) Then
    
       TxCtrl = &H1
       OutDataEightByte TxRedLED, TxCtrl, (Hour(Now) \ 10), (Hour(Now) Mod 10), (Minute(Now) \ 10), (Minute(Now) Mod 10), (Second(Now) \ 10), (Second(Now) Mod 10)
        
'        If (ReadKeyData(1) = 0) Then
'           TxCtrl = &H1
'           OutDataEightByte TxRedLED, TxCtrl, (Hour(Now) \ 10), (Hour(Now) Mod 10), (Minute(Now) \ 10), (Minute(Now) Mod 10), (Second(Now) \ 10), (Second(Now) Mod 10)
'        End If
'        If (ReadKeyData(1) = 1) Then
'           TxCtrl = &H2
'           OutDataEightByte TxRedLED, TxCtrl, (Hour(Now) \ 10), (Hour(Now) Mod 10), (Minute(Now) \ 10), (Minute(Now) Mod 10), (Second(Now) \ 10), (Second(Now) Mod 10)
'        End If
        
        CloseUsbDevice
        
    Else
        StatusDisplay.Caption = " DEVICE OFF "
    End If

End Sub


