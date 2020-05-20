VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "102學年度  工業類科學生技藝競賽  電腦修護職種   第二站  崗位號碼:98"
   ClientHeight    =   2955
   ClientLeft      =   195
   ClientTop       =   420
   ClientWidth     =   8010
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
   ScaleHeight     =   2955
   ScaleMode       =   0  'User
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "更改密碼"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
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
      Left            =   3360
      TabIndex        =   3
      Text            =   "98"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "讀取"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   0
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   885
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
      Height          =   615
      Index           =   1
      Left            =   5248
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   2200
   End
   Begin VB.Label Label3 
      Caption         =   "密碼"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
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
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
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


Private Sub Command1_Click(Index As Integer)
    Dim Result As Boolean

    Result = OpenUsbDevice(VendorID, ProductID)
    If (Result) Then
        
       ReadData ReadKeyData
       CloseUsbDevice
       
       Text1.Text = Chr(ReadKeyData(1)) & Chr(ReadKeyData(2)) & Chr(ReadKeyData(3)) & Chr(ReadKeyData(4))

    Else
       Text1.Text = " "
       

    End If

End Sub

Private Sub Command2_Click()
    Dim Result As Boolean

    
    Result = OpenUsbDevice(VendorID, ProductID)
    If (Result) Then
        
'       strLen = Mid(Text1.Text, 1, 1)
       TxCtrl = &H5
       OutDataEightByte 0, TxCtrl, Mid(Text1.Text, 1, 1), Mid(Text1.Text, 2, 1), Mid(Text1.Text, 3, 1), Mid(Text1.Text, 4, 1), 0, 0
       CloseUsbDevice
       
    Else


    End If
End Sub

'

Private Sub Form_Load()
    
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
    Dim Result As Boolean
    Result = OpenUsbDevice(VendorID, ProductID)
    If (Result) Then

'        OutDataEightByte TxRedLED, TxCtrl, 0, 0, 0, 0, 0, 0
        CloseUsbDevice
    End If
End Sub
'

Private Sub Label1_Click()

End Sub

Private Sub Text_Station_Change()
   Dim Result As Boolean
   Result = OpenUsbDevice(VendorID, ProductID)
   

   If (Len(Text_Station.Text) = 2) Then
      If (Val(Text_Station.Text) < 100) And (Val(Text_Station.Text) > 0) Then
              
         If (Result) Then
            TxCtrl = &H4
            OutDataEightByte 0, TxCtrl, Val(Text_Station.Text), 0, 0, 0, 0, 0
            CloseUsbDevice
         End If
         
      End If
   End If
End Sub


