VERSION 5.00
Begin VB.Form KeyValuef 
   Caption         =   " 按鍵值"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "變數 KeyCode 值"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "變數 Shift 值"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "按鍵值："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "KeyValuef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  KeyValuef.KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Label2.Caption = Shift
  Label3.Caption = KeyCode
  
  Select Case Shift
    Case 0
      Label6.Caption = "單一按鍵"
      Label7 = Chr(KeyCode)
    Case 1
      Label6.Caption = "Shift"
      Label7 = Chr(KeyCode)
    Case 2
      Label6.Caption = "Ctrl"
      Label7 = Chr(KeyCode)
    Case 4
      Label6.Caption = "Alt"
      Label7 = Chr(KeyCode)
    Case 3
      Label6.Caption = "Shift + Ctrl"
      Label7 = Chr(KeyCode)
    Case 5
      Label6.Caption = "Shift + Alt"
      Label7 = Chr(KeyCode)
    Case 6
      Label6.Caption = "Ctrl + Alt"
      Label7 = Chr(KeyCode)
    Case 7
      Label6.Caption = "Shift + Ctrl + Alt"
      Label7 = Chr(KeyCode)
  End Select
End Sub

