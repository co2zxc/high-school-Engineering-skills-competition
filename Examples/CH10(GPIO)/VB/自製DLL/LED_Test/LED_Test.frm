VERSION 5.00
Begin VB.Form USB 
   Caption         =   "LED_Test"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4200
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3480
      Top             =   1320
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   6
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "Output Test :"
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
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "D19"
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
      TabIndex        =   16
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "D20"
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
      Left            =   3120
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "D21"
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
      Left            =   2640
      TabIndex        =   14
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "D22"
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
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "D23"
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
      Left            =   1680
      TabIndex        =   12
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "D24"
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
      Left            =   1200
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "D25"
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
      TabIndex        =   10
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label20 
      Caption         =   "D26"
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
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   375
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
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   2085
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ���عϮ�: USB�����]�p�P���γ]�p�J�� Visual Basic �d�ҵ{��-CH10(�\�éM�s��)
'   �{���\��: �Q��PB��PD�X��8������LED

Const VendorID = &H1234                                     ' �]�w�c��ӽX
Const ProductID = &H2468                                    ' �]�w���~�X
Dim usbHID As New HID
Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = usbHID.OpenHIDDevice(VendorID, ProductID)       ' �T�{USB I/O�����d����O�O�_�s��
    
    If (Check) Then
        Label9.Caption = " DEVICE ON "                      ' DEVICE ON
    Else
        Label9.Caption = " DEVICE OFF "                     ' DEVICE OFF
    End If
    
    usbHID.CloseHIDDevice                                   ' ����HID�˸m
    isConnect = Check
End Function

Private Sub USBAction()
    Dim Check As Boolean
    Dim Send(8) As Byte
    Dim i As Byte
    Check = usbHID.OpenHIDDevice(VendorID, ProductID)       ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then                                         ' �P�_ HID Device �O�_�s�b
        For i = 0 To 7
            If Check1(i).Value = Checked Then               ' ���o CheckBox ���A
                Send(0) = Send(0) + 2 ^ i
            End If
        Next i
        Send(0) = &HFF - Send(0)                            ' �Ϭ�
        usbHID.HIDSetReport Send                            ' �ǿ�HID���
        
        usbHID.CloseHIDDevice                               ' ����HID�˸m
    End If
End Sub

Private Sub Timer1_Timer()
    isConnect
    USBAction
End Sub
