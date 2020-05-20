VERSION 5.00
Begin VB.Form frmQ2 
   Caption         =   "去除背景得到主體影像"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5325
   StartUpPosition =   2  '螢幕中央
   Begin VB.PictureBox picDifferent 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   3735
      ScaleHeight     =   85
      ScaleMode       =   3  '像素
      ScaleWidth      =   93
      TabIndex        =   9
      Top             =   1774
      Width           =   1455
   End
   Begin VB.PictureBox picGrey 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Index           =   1
      Left            =   1935
      ScaleHeight     =   85
      ScaleMode       =   3  '像素
      ScaleWidth      =   93
      TabIndex        =   8
      Top             =   2783
      Width           =   1455
   End
   Begin VB.PictureBox picLoad 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Index           =   1
      Left            =   135
      ScaleHeight     =   85
      ScaleMode       =   3  '像素
      ScaleWidth      =   101
      TabIndex        =   7
      Top             =   2783
      Width           =   1575
   End
   Begin VB.PictureBox picGrey 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Index           =   0
      Left            =   1935
      ScaleHeight     =   85
      ScaleMode       =   3  '像素
      ScaleWidth      =   93
      TabIndex        =   6
      Top             =   623
      Width           =   1455
   End
   Begin VB.PictureBox picLoad 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Index           =   0
      Left            =   120
      ScaleHeight     =   85
      ScaleMode       =   3  '像素
      ScaleWidth      =   101
      TabIndex        =   5
      Top             =   623
      Width           =   1575
   End
   Begin VB.CommandButton cmdDifferent 
      Caption         =   "兩影像相減"
      Height          =   375
      Left            =   3855
      TabIndex        =   4
      Top             =   1151
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrey 
      Caption         =   "影像2灰階"
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   2183
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrey 
      Caption         =   "影像1灰階"
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   143
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "載入影像2"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2183
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "載入影像1"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   143
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DirPath$

Private Sub cmdDifferent_Click()
Dim i%, j%, R#, G#, B#, tmp#
For i = 0 To picLoad(0).ScaleHeight - 1
    For j = 0 To picLoad(0).ScaleWidth - 1
        picDifferent.PSet (i, j), Not Abs(picGrey(0).Point(i, j) - picGrey(1).Point(i, j))
    Next j
Next i
End Sub

Private Sub cmdGrey_Click(Index As Integer)

Dim i%, j%, R#, G#, B#, tmp#
For i = 0 To picLoad(Index).ScaleHeight - 1
    For j = 0 To picLoad(Index).ScaleWidth - 1
        B = picLoad(Index).Point(i, j) Mod 256
        G = (picLoad(Index).Point(i, j) \ 256) Mod 256
        R = (picLoad(Index).Point(i, j) \ 256 ^ 2) Mod 256
        tmp = (R + G + B) / 3
        picGrey(Index).PSet (i, j), RGB(tmp, tmp, tmp)
    Next j
Next i
End Sub

Private Sub cmdLoad_Click(Index As Integer)
picLoad(Index).Picture = LoadPicture(DirPath & "pic" & IIf(Index = 0, "1", "2") & ".bmp")
End Sub

Private Sub Form_Load()
DirPath = App.Path & IIf(Right(App.Path, 1) = "/", "", "/")
End Sub
