VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "�q���w��˭פA�� �Ĥ@�� �ĤQ�D"
   ClientHeight    =   6285
   ClientLeft      =   3345
   ClientTop       =   2835
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   10830
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   360
      Top             =   2400
   End
   Begin VB.TextBox T1 
      Height          =   615
      Left            =   8160
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   5160
      Width           =   2295
   End
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
      Left            =   5520
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   3960
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
      Left            =   3480
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   1
      Top             =   5160
      Width           =   2200
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   1800
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
      Left            =   1440
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   0
      Top             =   3960
      Width           =   2200
   End
   Begin VB.Label Label3 
      Caption         =   "KEY"
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "DEVICE OFF"
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
      Left            =   930
      TabIndex        =   4
      Top             =   5520
      Width           =   1905
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
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   7575
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   8
      Left            =   960
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   7
      Left            =   1440
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   6
      Left            =   1920
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   5
      Left            =   2400
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   4
      Left            =   2880
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   3
      Left            =   3360
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   2
      Left            =   3840
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      FillColor       =   &H00008000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   1
      Left            =   4320
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   7
      Left            =   5280
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   6
      Left            =   5760
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   5
      Left            =   6240
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   4
      Left            =   6720
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   3
      Left            =   7200
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   2
      Left            =   7680
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   1
      Left            =   8160
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   8
      Left            =   4800
      Shape           =   3  '���
      Top             =   3000
      Width           =   375
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
    
  b(0) = &H80
    b(1) = &H40
    b(2) = &H20
    b(3) = &H10
    b(4) = &H8
    b(5) = &H4
    b(6) = &H2
    b(7) = &H1
    b(8) = &H2
    b(9) = &H4
    b(10) = &H8
    b(11) = &H10
    b(12) = &H20
    b(13) = &H40
    b(14) = &H80
    d = 14

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
    Label1.Caption = "                                                    " & _
                     "Current Time : " & Time$

    Result = OpenUsbDevice(VendorID, ProductID)                     ' �T�{USB I/O�����d�O�_�s��
    
    If (Result) Then
        
       If ((PreConnect = False) And (Result = True)) Then
            TxRedLED = &H0                                          ' USB ���
            TxCtrl = &H21
            OutDataEightByte TxRedLED, TxCtrl, 0, 0, 0, 0, 0, 0   ' �ǿ�USB���
            
            For i = 1 To 8                                             ' DEVICE ON
            Shape1(i).FillStyle = 0
            Shape2(i).FillStyle = 0
           Label2.Caption = "DEVICE ON"
            Next i
            
        End If
        
         If a = 0 Then
            If c > d Then
                TxRedLED = &H0                                       ' USB ���
                a = 99
                Shape2(8).FillColor = &H8000&
                c = 0
            Else
                If c > 0 And c < 8 Then
                Shape2(9 - c).FillColor = &H8000&
                Else
                  If c >= 8 Then
                    Shape2(c - 7).FillColor = &H8000&
                    End If
                End If
                TxRedLED = b(c)                                      ' USB ���
                If c < 8 Then
                Shape2(8 - c).FillColor = &HFF00&
                Else
                Shape2(c - 6).FillColor = &HFF00&
                End If
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
     
        OutDataEightByte &H5, &H5, &H5, &H5, &H5, &H5, &H5, &H5      ' �ǿ�USB���
         ReadData ReadKeyData
        
       Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
        ReadData ReadKeyData
        If ReadKeyData(1) = &H1 Then
        T1.Text = "1"
        OutDataEightByte &H5, &H5, &H5, &H5, &H5, &H5, &H5, &H5
        ElseIf ReadKeyData(1) = &H2 Then
        T1.Text = "2"
        OutDataEightByte 0, &H5, 2, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H3 Then
        T1.Text = "3"
        OutDataEightByte 0, &H5, 3, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H4 Then
        T1.Text = "4"
        OutDataEightByte 0, &H5, 4, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H5 Then
        T1.Text = "5"
        OutDataEightByte 0, &H5, 5, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H6 Then
        T1.Text = "6"
        OutDataEightByte 0, &H5, 6, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H7 Then
        T1.Text = "7"
        OutDataEightByte 0, &H5, 7, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H8 Then
        T1.Text = "8"
        OutDataEightByte 0, &H5, 8, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H9 Then
        T1.Text = "9"
        OutDataEightByte 0, &H5, 9, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &HA Then
        T1.Text = "*"
        OutDataEightByte 0, &H5, 10, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &HB Then
        T1.Text = "0"
        OutDataEightByte 0, &H5, 0, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &HC Then
        T1.Text = "#"
        OutDataEightByte 0, &H5, 11, 0, 0, 0, 0, 0
        End If
        
    End If
        
        
        
        
        
        CloseUsbDevice                                              ' ����HID�˸m
        
        
    Else
        For i = 1 To 8                                              ' DEVICE OFF
        Shape1(i).FillStyle = 1
        Shape2(i).FillStyle = 1
        Label2.Caption = "DEVICE OFF"
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

Private Sub Timer2_Timer()
  Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)
    If (Check) Then
        ReadData ReadKeyData
        If ReadKeyData(1) = &H1 Then
        T1.Text = "1"
        OutDataEightByte &H5, 1, 0, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H2 Then
        T1.Text = "2"
       OutDataEightByte &H5, 2, 0, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H3 Then
        T1.Text = "3"
        OutDataEightByte &H5, 3, 0, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H4 Then
        T1.Text = "4"
        OutDataEightByte 0, &H5, 4, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H5 Then
        T1.Text = "5"
        OutDataEightByte 0, &H5, 5, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H6 Then
        T1.Text = "6"
        OutDataEightByte 0, &H5, 6, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H7 Then
        T1.Text = "7"
        OutDataEightByte 0, &H5, 7, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H8 Then
        T1.Text = "8"
        OutDataEightByte 0, &H5, 8, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &H9 Then
        T1.Text = "9"
        OutDataEightByte 0, &H5, 9, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &HA Then
        T1.Text = "*"
        OutDataEightByte 0, &H5, 10, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &HB Then
        T1.Text = "0"
        OutDataEightByte 0, &H5, 0, 0, 0, 0, 0, 0
        ElseIf ReadKeyData(1) = &HC Then
        T1.Text = "#"
        OutDataEightByte 0, &H5, 11, 0, 0, 0, 0, 0
        End If
            End If
End Sub
