VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "A ^ B Mod C = ?"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2813
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtIn 
      Alignment       =   1  '靠右對齊
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
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Text            =   "0"
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox txtIn 
      Alignment       =   1  '靠右對齊
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
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Text            =   "0"
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtIn 
      Alignment       =   1  '靠右對齊
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
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   653
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblShow 
      AutoSize        =   -1  'True
      Caption         =   "A ^ B Mod C = ?"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1545
      TabIndex        =   8
      Top             =   2025
      Width           =   1590
   End
   Begin VB.Label lblIn 
      AutoSize        =   -1  'True
      Caption         =   "C="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   270
   End
   Begin VB.Label lblIn 
      AutoSize        =   -1  'True
      Caption         =   "B="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   270
   End
   Begin VB.Label lblIn 
      AutoSize        =   -1  'True
      Caption         =   "A="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim a&, b&, c&, s
Dim i#
a = 4937
b = 4732
c = 63
s = 1
For i = 1 To b
    s = s * a
    s = s Mod c
Next i
Print s
End Sub

Private Sub cmdCount_Click()
Dim a@, b@, c 'a ^ b Mod c = ?
Dim s@
Dim i@
a = CCur(txtIn(0))
b = CCur(txtIn(1))
c = CCur(txtIn(2))

s = 1
For i = 1 To b
    s = s * a
    s = s Mod c
Next i
lblShow = CStr(s)
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
txtIn_GotFocus (0)
End Sub

Private Sub lblShow_Change()
With lblShow
    .Left = Me.ScaleWidth / 2 - .Width \ 2
End With
End Sub

Private Sub txtIn_GotFocus(Index As Integer)
With txtIn(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub
