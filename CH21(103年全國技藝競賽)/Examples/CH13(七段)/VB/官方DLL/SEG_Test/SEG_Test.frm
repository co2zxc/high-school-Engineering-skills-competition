VERSION 5.00
Begin VB.Form USB 
   Caption         =   "SEG_Test"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   6015
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   2520
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '�m�����
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Text            =   "0"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '�m�����
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Text            =   "0"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '�m�����
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '�m�����
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "�п�J0~15"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   4080
      TabIndex        =   13
      Top             =   1920
      Width           =   1725
   End
   Begin VB.Label Label9 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '��u�T�w
      Caption         =   " DEVICE OFF "
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2085
   End
   Begin VB.Label Label3 
      Caption         =   "U6 : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "U7 : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "U5 : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "U4 : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ���عϮ�: USB�����]�p�P���γ]�p�J�� Visual Basic �d�ҵ{��-CH13(�\�éM�s��)
'   �{���\��: �Q��PB��PD�X��4���C�q��ܾ�

'�ܼƫŧi
Dim U4Data As Byte
Dim U5Data As Byte
Dim U6Data As Byte
Dim U7Data As Byte

Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)  ' �T�{USB I/O�����d����O�O�_�s��
     
    If (Check) Then
        Label9.Caption = " DEVICE ON "          ' DEVICE ON
    Else
        Label9.Caption = " DEVICE OFF "         ' DEVICE OFF
    End If
    
    CloseUsbDevice                              ' ����HID�˸m
    isConnect = Check
End Function

Private Sub USBAction()
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)                          ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        OutDataEightByte U4Data, U5Data, U6Data, U7Data, 0, 0, 0, 0     ' �ǿ�HID���
        CloseUsbDevice                                                  ' ����HID�˸m
    End If
End Sub

Private Sub Command1_Click()
    U4Data = Text1.Text                                                 ' U4 Data
    USBAction
End Sub

Private Sub Command2_Click()
    U5Data = Text2.Text                                                 ' U5 Data
    USBAction
End Sub

Private Sub Command3_Click()
    U6Data = Text3.Text                                                 ' U6 Data
    USBAction
End Sub

Private Sub Command4_Click()
    U7Data = Text4.Text                                                 ' U7 Data
    USBAction
End Sub

Private Sub Timer1_Timer()
    isConnect
End Sub

