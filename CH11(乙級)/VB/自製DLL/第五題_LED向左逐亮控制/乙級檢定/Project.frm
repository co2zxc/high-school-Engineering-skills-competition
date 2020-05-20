VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "�q���w��˭פA�� �Ĥ@�� �Ĥ��D"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   8355
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
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
      Left            =   1200
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   4
      Top             =   2640
      Width           =   2200
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
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
      Left            =   4680
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   1
      Top             =   3840
      Width           =   2200
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
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
      Index           =   0
      Left            =   4680
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   0
      Top             =   2640
      Width           =   2200
   End
   Begin VB.Label StatusDisplay 
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
      Left            =   1200
      TabIndex        =   3
      Top             =   4080
      Width           =   2085
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
      Height          =   1575
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   6495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b(99) As Byte
Dim c As Integer
Dim d As Integer
Dim Send(8) As Byte
Const VendorID = &H1234
Const ProductID = &H2468
Dim usbHID As New HID

Private Sub Command1_Click(Index As Integer)
    a = Index
    Send(2) = &H21
    c = 0
End Sub

Private Sub Command2_Click(Index As Integer)
    a = Index
    Send(2) = &H22
    c = 0
End Sub

Private Sub Form_Load()
    a = 99
    
    ' NO.5
    b(0) = &H1
    b(1) = &H3
    b(2) = &H7
    b(3) = &HF
    b(4) = &H1F
    b(5) = &H3F
    b(6) = &H7F
    b(7) = &HFF
    d = 7
    
    c = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim Result As Boolean
    Result = usbHID.OpenHIDDevice(VendorID, ProductID)          ' �T�{USB devie�O�_�s�u
    If (Result) Then
        Send(0) = &H0                                           ' USB ���
        Send(1) = &H20                                          ' �e�ɽX
        usbHID.HIDSetReport Send                                ' �ǿ�USB���
        usbHID.CloseHIDDevice                                   ' ����HID�˸m
    End If
End Sub


Private Sub Timer1_Timer()
    Dim Result As Boolean
    Dim k As Byte
    Static PreConnect As Boolean
    If c > 20 Then c = 20
    Label1.Caption = "                                  " & _
                     "Current Time : " & Time$

    Result = usbHID.OpenHIDDevice(VendorID, ProductID)              ' �T�{USB devie�O�_�s�u
    
    If (Result) Then
        StatusDisplay.Caption = " DEVICE ON "                       ' DEVICE ON
        If ((PreConnect = False) And (Result = True)) Then
            Send(0) = &H0                                           ' USB ���
            usbHID.HIDSetReport Send                                ' �ǿ�USB���
        End If
        
        If a = 0 Then
            If c > d Then
                Send(0) = &H0                                       ' USB ���
                a = 99
                c = 0
            Else
                Send(0) = b(c)                                      ' USB ���
            End If
        End If
        
        If a = 1 Then
            Send(0) = &H0                                           ' USB ���
            Send(1) = &H20                                          ' �e�ɽX
            usbHID.HIDSetReport Send                                ' �ǿ�USB���
            a = 99
            usbHID.CloseHIDDevice                                   ' ����HID�˸m
            PreConnect = Result
            End
        End If
        
        Send(1) = &H20                                              ' �e�ɽX
        usbHID.HIDSetReport Send                                    ' �ǿ�USB���
        usbHID.CloseHIDDevice                                       ' ����HID�˸m
    Else
        StatusDisplay.Caption = " DEVICE OFF "                      ' DEVICE OFF
        If a = 1 Then
            a = 99
            PreConnect = Result
            End
        End If
    End If
    PreConnect = Result
    c = c + 1
End Sub
