VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "�q���w��˭פA�� �Ĥ@�� �ĤK�D"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   8355
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Red LED"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   2
      Left            =   5160
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   3840
      Width           =   2200
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   1
      Left            =   3120
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   1
      Top             =   5040
      Width           =   2200
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   1680
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Green LED"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   1080
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   0
      Top             =   3840
      Width           =   2200
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
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
         Name            =   "�s�ө���"
         Size            =   26.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   7575
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   8
      Left            =   600
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   7
      Left            =   1080
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   6
      Left            =   1560
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   5
      Left            =   2040
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   4
      Left            =   2520
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   3
      Left            =   3000
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   2
      Left            =   3480
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   1
      Left            =   3960
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   7
      Left            =   4920
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   6
      Left            =   5400
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   5
      Left            =   5880
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   4
      Left            =   6360
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   3
      Left            =   6840
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   2
      Left            =   7320
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   1
      Left            =   7800
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   8
      Left            =   4440
      Shape           =   3  '���
      Top             =   2880
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ���عϮ�: USB�����]�p�P���γ]�p�J�� Visual Basic �d�ҵ{��-CH11(�\�éM�s��)
'   �{���\��: �Q��PB��PD�X��8��LED����(�A���˩w)

Option Explicit

Public TxRedLED As Byte
Public TxCtrl As Byte

Private Sub Command1_Click(Index As Integer)
    a = Index
    TxCtrl = &H21
    c = 0
End Sub

Private Sub Command2_Click(Index As Integer)
    a = Index
    TxCtrl = &H22
    c = 0
End Sub

Private Sub Form_Load()
    a = 99
    
    ' NO.8
    b(0) = &H81
    b(1) = &H42
    b(2) = &H24
    b(3) = &H18
    d = 3
    
    c = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim Result As Boolean
    Result = OpenUsbDevice(VendorID, ProductID)                 ' �T�{USB I/O�����d����O�O�_�s��
    If (Result) Then
        TxRedLED = &H0                                          ' USB ���
        OutDataEightByte TxRedLED, TxCtrl, 0, 0, 0, 0, 0, 0     ' �ǿ�USB���
        CloseUsbDevice                                          ' ����HID�˸m
    End If
End Sub


Private Sub Timer1_Timer()
    Dim Result As Boolean
    Dim k As Byte
    Dim i As Integer
    Static PreConnect As Boolean
    If c > 20 Then c = 20
    Label1.Caption = "                                            " & _
                     "Current Time : " & Time$

    Result = OpenUsbDevice(VendorID, ProductID)                     ' �T�{USB I/O�����d����O�O�_�s��
    
    If (Result) Then
         
       If ((PreConnect = False) And (Result = True)) Then
            TxRedLED = &H0                                          ' USB ���
            TxCtrl = &H21
            OutDataEightByte TxRedLED, TxCtrl, 0, 0, 0, 0, 0, 0   ' �ǿ�USB���
            
            For i = 1 To 8                                             ' DEVICE ON
            Shape1(i).FillStyle = 0
            Shape2(i).FillStyle = 0
            Next i
            
        End If
        
        If a = 0 Then
            If c > d Then
                TxRedLED = &H0                                       ' USB ���
                a = 99
                Shape2(c).FillColor = &H8000&
                Shape2(9 - c).FillColor = &H8000&
                c = 0
            Else
                If c > 0 Then
                Shape2(c).FillColor = &H8000&
                Shape2(9 - c).FillColor = &H8000&
                End If
                TxRedLED = b(c)                                      ' USB ���
                Shape2(c + 1).FillColor = &HFF00&
                Shape2(8 - c).FillColor = &HFF00&
            End If
        End If
        
        If a = 1 Then
            TxRedLED = &H0                                          ' USB ���
            OutDataEightByte TxRedLED, TxCtrl, 0, 0, 0, 0, 0, 0     ' �ǿ�USB���
            a = 99
            CloseUsbDevice                                          ' ����HID�˸m
            PreConnect = Result
            End
        End If
        
        If a = 2 Then
            If c > 7 Then
                TxRedLED = &H0                                       ' USB ���
                a = 99
                Shape1(c).FillColor = &H80&
                c = 0
            Else
                If c > 0 Then
                Shape1(c).FillColor = &H80&
                End If
                TxRedLED = 2 ^ c                                    ' USB ���
                Shape1(c + 1).FillColor = &HFF&
            End If
            
        End If
        
        OutDataEightByte TxRedLED, TxCtrl, 0, 0, 0, 0, 0, 0         ' �ǿ�USB���
        CloseUsbDevice                                              ' ����HID�˸m
    Else
        For i = 1 To 8                                              ' DEVICE OFF
        Shape1(i).FillStyle = 1
        Shape2(i).FillStyle = 1
        Next i
        
        If a = 1 Then
            a = 99
            PreConnect = Result
            End
        End If
    End If
    PreConnect = Result
    c = c + 1

End Sub

