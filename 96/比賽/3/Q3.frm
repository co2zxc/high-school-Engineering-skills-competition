VERSION 5.00
Begin VB.Form Q3 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   9240
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   8
      Left            =   7680
      TabIndex        =   18
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   7
      Left            =   6480
      TabIndex        =   17
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   6
      Left            =   5280
      TabIndex        =   16
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   5
      Left            =   7680
      TabIndex        =   15
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   4
      Left            =   6480
      TabIndex        =   14
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   3
      Left            =   5280
      TabIndex        =   13
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   2
      Left            =   7680
      TabIndex        =   12
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   1
      Left            =   6480
      TabIndex        =   11
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   10
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   8
      Left            =   3240
      TabIndex        =   9
      Text            =   "255"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   7
      Left            =   2040
      TabIndex        =   8
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   6
      Left            =   840
      TabIndex        =   7
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   5
      Left            =   3240
      TabIndex        =   6
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Text            =   "255"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Text            =   "255"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "B"
      Height          =   495
      Index           =   8
      Left            =   240
      TabIndex        =   27
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "G"
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   26
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "R"
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "I"
      Height          =   495
      Index           =   5
      Left            =   7800
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "S"
      Height          =   495
      Index           =   4
      Left            =   6600
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "H"
      Height          =   495
      Index           =   3
      Left            =   5280
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "B"
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "G"
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "R"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Q3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim r#, g#, b#
r=
End Sub

Private Sub Label1_Click(Index As Integer)

End Sub
