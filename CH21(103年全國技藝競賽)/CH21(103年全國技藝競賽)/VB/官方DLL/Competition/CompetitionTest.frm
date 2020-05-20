VERSION 5.00
Begin VB.Form USB 
   Caption         =   "CompetitionTest"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8775
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.OptionButton Option2 
      Caption         =   "Low"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "High"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   840
   End
   Begin VB.Label Label6 
      Alignment       =   2  '�m�����
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  '�m�����
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  '�m�����
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   720
      TabIndex        =   2
      Top             =   2520
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
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
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
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
      Left            =   5160
      TabIndex        =   0
      Top             =   4200
      Width           =   2085
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReadKeyData(8) As Byte

Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)          ' �T�{USB devie�O�_�s�u
     
    If (Check) Then
        Label1.Caption = " DEVICE ON "                  ' DEVICE ON
    Else
        Label1.Caption = " DEVICE OFF "                 ' DEVICE OFF
    End If
    
    CloseUsbDevice                                      ' ����HID�˸m
    isConnect = Check
End Function

Private Sub Timer1_Timer()
    Dim timeStringArray() As String
    Dim ReadKeyData(8) As Byte
    Dim OutData As Byte
    Dim Check As Boolean
    
    isConnect
    If (Option1.Value = True) Then
        OutData = 1
    Else
        OutData = 0
    End If
    
    timeStringArray = Split(Time$, ":", 3)                                                  ' �N�ɶ������ɤ���ӧO�x�s
    
    Label2.Caption = "                                  " & _
                     "Current Time : " & Time$
    
    
    Check = OpenUsbDevice(VendorID, ProductID)                                              ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        OutDataEightByte Int(timeStringArray(0)), Int(timeStringArray(1)), _
                         Int(timeStringArray(2)), OutData, 0, 0, 0, 0                       ' �N�ɤ�������Q��ƭӦ�ƶǰe�� USB Device
                         
        ReadData ReadKeyData                                                                ' Ū��HID���
        
        If (ReadKeyData(1)) Then
            Label4.Caption = " Time UP "
        Else
            Label4.Caption = "   "
        End If
        
        If (ReadKeyData(8)) Then
            Label5.Caption = " Alarm Mode "
        Else
            Label5.Caption = " Clock Mode  "
        End If
        
        
        CloseUsbDevice                                                                      ' ����HID�˸m
    End If
    
    Label3.Caption = "                                   " & _
                     "Alarm Time : " _
                     & Str(ReadKeyData(2) * 10 + ReadKeyData(3)) & ":" _
                     & Str(ReadKeyData(4) * 10 + ReadKeyData(5)) & ":" _
                     & Str(ReadKeyData(6) * 10 + ReadKeyData(7))
End Sub
