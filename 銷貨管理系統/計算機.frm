VERSION 5.00
Begin VB.Form number 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "計算機"
   ClientHeight    =   4620
   ClientLeft      =   8835
   ClientTop       =   2295
   ClientWidth     =   4260
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   4620
   ScaleWidth      =   4260
   Begin VB.CommandButton cmdl 
      Caption         =   "/"
      Height          =   495
      Left            =   2520
      TabIndex        =   17
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdx 
      Caption         =   "*"
      Height          =   495
      Left            =   2520
      TabIndex        =   16
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "C"
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
      Left            =   360
      TabIndex        =   15
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdEnd 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "回主畫面"
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdAns 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "="
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
      Left            =   1800
      TabIndex        =   12
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdSub 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "-"
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
      Left            =   2520
      TabIndex        =   14
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "+"
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
      Left            =   2520
      TabIndex        =   13
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdDot 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "."
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
      Left            =   1080
      TabIndex        =   11
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "9"
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
      Index           =   9
      Left            =   1800
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "8"
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
      Index           =   8
      Left            =   1080
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "7"
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
      Index           =   7
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "6"
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
      Index           =   6
      Left            =   1800
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "5"
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
      Index           =   5
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "4"
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
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "3"
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
      Index           =   3
      Left            =   1800
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "2"
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
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "1"
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
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "0"
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
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblBoard 
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "number"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim opt As Integer
Dim total As Double, temp_val As Double
Dim dot As Single, zero As Single
Const Nul = 0, Add = 1, Subb = 2, mul = 3, div = 4, Ans = 5

Private Sub cmdAdd_Click()
    Select Case opt
        Case Nul:   total = temp_val
        Case Subb:  total = total - temp_val
        Case mul:   total = total * temp_val
        Case div:   total = total / temp_val
        Case Add:   total = total + temp_val
    End Select
    lblBoard = Str(total)
    temp_val = 0
    opt = Add
End Sub

Private Sub cmdAns_Click()
    Select Case opt
        Case Nul: total = temp_val
        Case Add: total = total + temp_val
        Case Subb: total = total - temp_val
        Case mul: total = total * temp_val
        Case div: total = total / temp_val
        Case Ans: total = temp_val
    End Select
    lblBoard = Str(total)
    temp_val = 0
    opt = Ans
End Sub

Private Sub cmdCancel_Click()
    opt = Nul
    total = 0
    temp_val = 0
    lblBoard.Caption = Str(temp_val)
End Sub

Private Sub cmdDot_Click()
    dot = True
End Sub

Private Sub cmdEnd_Click()
 Unload Me
    frmMDIMain.Show
End Sub

Private Sub cmdl_Click()
    Select Case opt
        Case Nul:   total = temp_val
        Case Add:   total = total + temp_val
        Case Subb:  total = total - temp_val
        Case mul:   total = total * temp_val
        Case div:   total = total / temp_val
    End Select
    lblBoard = Str$(total)
    temp_val = 0
    opt = div
End Sub

Private Sub cmdNum_Click(index As Integer)
    If dot = True Then
        temp_val = Val(Str(temp_val) + "." + Str(index))
        If index = 0 Then zero = True
        dot = False
    Else
        If zero = True Then
            temp_val = Val(Str(temp_val) + ".0" + Str(index))
            zero = False
        Else
            temp_val = Val(Str(temp_val) + Str(index))
        End If
    End If
    lblBoard.Caption = Str(temp_val)
End Sub

Private Sub cmdSub_Click()
    Select Case opt
        Case Nul:   total = temp_val
        Case Add:   total = total + temp_val
        Case mul:   total = total * temp_val
        Case div:   total = total / temp_val
        Case Subb:  total = total - temp_val
    End Select
  lblBoard = Str(total)
  temp_val = 0
  opt = Subb
End Sub

Private Sub cmdx_Click()
    Select Case opt
        Case Nul:   total = temp_val
        Case Add:   total = total + temp_val
        Case Subb:  total = total - temp_val
        Case div:   total = total / temp_val
        Case mul:   total = total * temp_val
    End Select
    lblBoard = Str(total)
    temp_val = 0
    opt = mul
End Sub
