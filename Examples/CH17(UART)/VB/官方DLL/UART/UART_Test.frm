VERSION 5.00
Begin VB.Form USB 
   Caption         =   "UART_Test"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   4935
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   2295
      Left            =   480
      TabIndex        =   3
      Top             =   4200
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   120
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "UART_Test.frx":0000
      Left            =   1320
      List            =   "UART_Test.frx":0019
      TabIndex        =   0
      Text            =   "9600"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Transmit :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Receive :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "Baud : "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
     
    If (Check) Then
        Label3.Caption = " DEVICE ON "                  ' DEVICE ON
    Else
        Label3.Caption = " DEVICE OFF "                 ' DEVICE OFF
    End If
    
    CloseUsbDevice                                      ' 關閉HID裝置
    isConnect = Check
End Function

Private Sub USBAction()
    Dim Check As Boolean
    Dim Data As Byte
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then

        Select Case Combo1.Text
        Case "1200": Data = 0
        Case "2400": Data = 1
        Case "4800": Data = 2
        Case "9600": Data = 3
        Case "19200": Data = 4
        Case "38400": Data = 5
        Case "57600": Data = 6
        End Select
        For i = 0 To Len(Text1.Text) - 1
            OutDataEightByte &H20, (Asc(Mid(Text1.Text, i + 1, 1))), Data, 0, 0, 0, 0, 0
        Next i
        CloseUsbDevice                                          ' 關閉HID裝置
    End If
End Sub

Private Sub Command1_Click()
    Text1.Text = ""
End Sub

Private Sub Command2_Click()
    USBAction
End Sub

Private Sub Command3_Click()
    Text2.Text = ""
End Sub
Private Sub Timer1_Timer()
    isConnect
    Dim ReadKeyData(8) As Byte
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
    
    If (Check) Then
        ReadData ReadKeyData                            ' 讀取HID資料
        Text2.Text = Text2.Text & Chr(ReadKeyData(1))
    End If
    CloseUsbDevice                                      ' 關閉HID裝置
End Sub

