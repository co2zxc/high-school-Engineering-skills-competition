VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "隨便按顯示對應的ASCII碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32" (ByVal key As Long) As Integer

Private Sub Timer1_Timer()

For i = 0 To 255
    If GetKeyState(i) < 0 Then Label1.Caption = i
Next i

End Sub

