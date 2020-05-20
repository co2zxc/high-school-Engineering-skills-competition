VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "解一元二次方程式"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   5250
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtC 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtB 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtA 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton comCalculate 
      Caption         =   "求解"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   120
      X2              =   5160
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "x="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   18
      Top             =   3600
      Width           =   285
   End
   Begin VB.Label lblFunction 
      BackColor       =   &H8000000A&
      Caption         =   "Function Here"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label lblC 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Input C="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblB 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Input B="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Input A="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   990
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   720
      X2              =   2760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line4 
      X1              =   1440
      X2              =   1440
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   1560
      X2              =   2760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   1560
      X2              =   1560
      Y1              =   960
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   1560
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   180
      Left            =   1920
      TabIndex        =   13
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "B -4AC"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   12
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "2A"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "－"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   10
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "-B＋"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   9
      Top             =   720
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "x＝"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   180
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ax＋Bx＋C＝0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub comCalculate_Click()
txtX(0) = ""
txtX(1) = ""
Dim A%, B%, C%
A = Val(txtA)
B = Val(txtB)
C = Val(txtC)
lblFunction = "(" & Trim(Str(A)) & ")x2+(" & Trim(Str(B)) & ")x+(" & Trim(Str(C)) & ")=0"
Dim tmp#
tmp = B ^ 2 - 4 * A * C '333*333=>溢位,333^2=>OK
If A = 0 And B = 0 And C <> 0 Then '無解
    txtX(0) = "無解"
ElseIf A = 0 And B = 0 And C = 0 Then '無限多組解
    txtX(0) = "無限多組解"
ElseIf A = 0 And B <> 0 Then '-C/B,一解
    txtX(0) = Trim(Str(Round(-C / B, 2)))
    txtX(1) = "只有一解"
ElseIf A <> 0 And B <> 0 And tmp = 0 Then '-b/2a,同根
    txtX(0) = Trim(Str(B / (2 * A)))
    txtX(1) = "同根"
ElseIf A <> 0 And B <> 0 And tmp > 0 Then '公式法,實數解
    txtX(0) = Trim(Str((Round((-B + Sqr(tmp)) / (2 * A), 2))))
    txtX(0) = Trim(Str((Round((-B - Sqr(tmp)) / (2 * A), 2))))
ElseIf A <> 0 And B <> 0 And tmp < 0 Then '公式法,虛數解
    txtX(0) = Trim(Str((Round((-B + Sqr(-tmp)) / (2 * A), 2)))) & "i"
    txtX(0) = Trim(Str((Round((-B - Sqr(-tmp)) / (2 * A), 2)))) & "i"
End If
End Sub
