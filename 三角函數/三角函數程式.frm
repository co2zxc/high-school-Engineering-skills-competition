VERSION 5.00
Begin VB.Form �T����� 
   BackColor       =   &H0000FFFF&
   Caption         =   "�T�����--�@�̡G������"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form2"
   ScaleHeight     =   5160
   ScaleWidth      =   8505
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "�e��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   3975
      Left            =   960
      ScaleHeight     =   3915
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      Caption         =   "Csc�c="
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Index           =   5
      Left            =   6240
      TabIndex        =   9
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      Caption         =   "Sec�c="
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Index           =   4
      Left            =   6240
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      Caption         =   "Cot�c="
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Index           =   3
      Left            =   6240
      TabIndex        =   7
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      Caption         =   "Tan�c="
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Index           =   2
      Left            =   6240
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      Caption         =   "Cos�c="
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Index           =   1
      Left            =   6240
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      Caption         =   "Sin�c="
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Index           =   0
      Left            =   6240
      TabIndex        =   4
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "���ףc="
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
End
Attribute VB_Name = "�T�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, cx, cy

Private Sub Command1_Click()
Picture1.Cls
Picture1.Line (cx, 0)-(cx, a), QBColor(12)
Picture1.Line (0, cy)-(b, cy), QBColor(12)

Picture1.DrawWidth = 2
degree = Val(Text1)
       Liney = cy - 2000 * Sin(degree / 180 * 3.14159)
       Linex = cx + 2000 * Cos(degree / 180 * 3.14159)
       Picture1.Line (cx, cy)-(Linex, Liney), RGB(0, 0, 255)
       Picture1.Line (Linex, cy)-(Linex, Liney), RGB(0, 0, 255)

aSin = Round(Sin(degree / 180 * 3.14159), 3)
aCos = Round(Cos(degree / 180 * 3.14159), 3)

Label2(0) = "Sin�c= " + Trim(Str(Round(Sin(degree / 180 * 3.14159), 3))) 'trim�h���r�ꪺ�e�B��m�ť�
Label2(1) = "Cos�c= " + Trim(Str(Round(Cos(degree / 180 * 3.14159), 3))) 'round�p���I���
If aCos = 0 Then
   Label2(2) = "Tan�c= " + "�L���j"
   Label2(4) = "Sec�c= " + "�L���j"
   
Else
   Label2(2) = "tan�c= " + Trim(Str(Round(Tan(degree / 180 * 3.14159), 3)))
   Label2(4) = "Sec�c= " + Trim(Str(Round(1 / Cos(degree / 180 * 3.14159), 3)))

End If

If aSin = 0 Then
   Label2(3) = "Cot�c= " + "�L���j"
   Label2(5) = "Csc�c= " + "�L���j"
Else
   Label2(3) = "Cot�c= " + Trim(Str(Round(1 / Tan(degree / 180 * 3.14159), 3)))
   Label2(5) = "Csc�c= " + Trim(Str(Round(1 / Sin(degree / 180 * 3.14159), 3)))
End If
End Sub

Private Sub Form_Resize()
a = Picture1.Height
b = Picture1.Width
cy = a / 2
cx = b / 2
Picture1.Line (cx, 0)-(cx, a), QBColor(12)
Picture1.Line (0, cy)-(b, cy), QBColor(12)

End Sub
   
   
   
