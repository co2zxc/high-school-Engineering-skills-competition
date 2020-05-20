VERSION 5.00
Begin VB.Form USB 
   Caption         =   "LCD_Test"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   3045
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   2040
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
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "5678"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "1234"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
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
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   2085
   End
   Begin VB.Label Label2 
      Caption         =   "2. "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1. "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   135
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ���عϮ�: USB�����]�p�P���γ]�p�J�� Visual Basic �d�ҵ{��-CH14(�\�éM�s��)
'   �{���\��: �Q��PB��PD����LCD

Const VendorID = &H1234                                        ' �]�w�c��ӽX
Const ProductID = &H2468                                       ' �]�w���~�X
Dim usbHID As New HID

Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = usbHID.OpenHIDDevice(VendorID, ProductID)          ' �T�{USB I/O�����d����O�O�_�s��
    
    If (Check) Then
        Label3.Caption = " DEVICE ON "                         ' DEVICE ON
    Else
        Label3.Caption = " DEVICE OFF "                        ' DEVICE OFF
    End If
    
    usbHID.CloseHIDDevice                                      ' ����HID�˸m
    isConnect = Check
End Function

Private Sub USBAction()
    Dim Check As Boolean
    Dim Send(8) As Byte
    Check = usbHID.OpenHIDDevice(VendorID, ProductID)          ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        Send(0) = &H21                                         ' �P�_LCD���\��
        Send(1) = 0                                            ' �M��LCM
        usbHID.HIDSetReport Send                               ' �ǿ�HID���
        For i = 0 To Len(Text1.Text) - 1
            Send(0) = &H20                                     ' �e�ɽX
            Send(1) = &H80 + i                                 ' LCD ��m
            Send(2) = (Asc(Mid(Text1.Text, i + 1, 1)))         ' �g�J��� �� LCD �Ҳ�
            usbHID.HIDSetReport Send                           ' �ǿ�HID���
        Next i
        For i = 0 To Len(Text2.Text) - 1
            Send(0) = &H20                                     ' �P�_LCD���\��
            Send(1) = &HC0 + i                                 ' LCD ��m
            Send(2) = (Asc(Mid(Text2.Text, i + 1, 1)))         ' �g�J��� �� LCD �Ҳ�
            usbHID.HIDSetReport Send                           ' �ǿ�HID���
        Next i
        usbHID.CloseHIDDevice                                  ' ����HID�˸m
    End If
End Sub

Private Sub Command1_Click()
    USBAction
End Sub

Private Sub Timer1_Timer()
    isConnect
End Sub


