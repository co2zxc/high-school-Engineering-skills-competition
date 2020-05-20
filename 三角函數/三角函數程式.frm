VERSION 5.00
Begin VB.Form 三角函數 
   BackColor       =   &H0000FFFF&
   Caption         =   "三角函數--作者：陳錫齡"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form2"
   ScaleHeight     =   5160
   ScaleWidth      =   8505
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "畫圖"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      BackStyle       =   0  '透明
      Caption         =   "Cscθ="
      BeginProperty Font 
         Name            =   "標楷體"
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
      BackStyle       =   0  '透明
      Caption         =   "Secθ="
      BeginProperty Font 
         Name            =   "標楷體"
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
      BackStyle       =   0  '透明
      Caption         =   "Cotθ="
      BeginProperty Font 
         Name            =   "標楷體"
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
      BackStyle       =   0  '透明
      Caption         =   "Tanθ="
      BeginProperty Font 
         Name            =   "標楷體"
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
      BackStyle       =   0  '透明
      Caption         =   "Cosθ="
      BeginProperty Font 
         Name            =   "標楷體"
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
      BackStyle       =   0  '透明
      Caption         =   "Sinθ="
      BeginProperty Font 
         Name            =   "標楷體"
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
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "角度θ="
      BeginProperty Font 
         Name            =   "標楷體"
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
Attribute VB_Name = "三角函數"
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

Label2(0) = "Sinθ= " + Trim(Str(Round(Sin(degree / 180 * 3.14159), 3))) 'trim去除字串的前、後置空白
Label2(1) = "Cosθ= " + Trim(Str(Round(Cos(degree / 180 * 3.14159), 3))) 'round小數點位數
If aCos = 0 Then
   Label2(2) = "Tanθ= " + "無限大"
   Label2(4) = "Secθ= " + "無限大"
   
Else
   Label2(2) = "tanθ= " + Trim(Str(Round(Tan(degree / 180 * 3.14159), 3)))
   Label2(4) = "Secθ= " + Trim(Str(Round(1 / Cos(degree / 180 * 3.14159), 3)))

End If

If aSin = 0 Then
   Label2(3) = "Cotθ= " + "無限大"
   Label2(5) = "Cscθ= " + "無限大"
Else
   Label2(3) = "Cotθ= " + Trim(Str(Round(1 / Tan(degree / 180 * 3.14159), 3)))
   Label2(5) = "Cscθ= " + Trim(Str(Round(1 / Sin(degree / 180 * 3.14159), 3)))
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
   
   
   
