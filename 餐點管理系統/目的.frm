VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00C0FFFF&
   Caption         =   "����"
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   Picture         =   "�ت�.frx":0000
   ScaleHeight     =   10875
   ScaleWidth      =   15240
   WindowState     =   2  '�̤j��
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   975
      Left            =   13680
      Picture         =   "�ت�.frx":22D3
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   8400
      Width           =   855
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   2370
      Left            =   2280
      Picture         =   "�ت�.frx":3125
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   3105
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '��u�T�w
      Height          =   2850
      Left            =   8640
      Picture         =   "�ت�.frx":5300
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4425
   End
   Begin VB.Label Label3 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   $"�ت�.frx":9611
      BeginProperty Font 
         Name            =   "�رd���B��"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   3255
      Left            =   2280
      TabIndex        =   1
      Top             =   4080
      Width           =   7455
   End
   Begin VB.Image Image4 
      Height          =   4725
      Left            =   1320
      Picture         =   "�ت�.frx":97B5
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   13200
   End
   Begin VB.Label Label2 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "�ت�"
      BeginProperty Font 
         Name            =   "�رd���r��(P)"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  MDIForm1.StatusBar1.Panels(4) = "�t�Υت�"
End Sub

