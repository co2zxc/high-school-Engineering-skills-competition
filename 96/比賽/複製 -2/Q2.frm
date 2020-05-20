VERSION 5.00
Begin VB.Form Q2 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4605
   StartUpPosition =   2  '螢幕中央
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lbl2 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lbl 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Messages:"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Circular Queue"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Q2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Front%, Rear% 'text's index
Const w% = 480 'lbl寬
Dim Key%
Dim sHide%
Dim sss%
Dim XD%
Dim Label2() As Object

Private Sub Command1_Click()
Dim num%
Dim i%, flag As Boolean
flag = False
For i = 0 To UBound(Label2) - sHide
    If Label2(i) = "" Then flag = True
Next i
If Not flag Then
    Call subChg
    'Front = Front + 1
    'Label2(Key) = ""
    'Front = Front + 1
End If

num = fncGet
If Not flag Then
    Label2(Key) = num
Else
    Label2(Key) = num
End If
Key = Key + 1
If Key > UBound(Label2) - sHide Then Key = Key - UBound(Label2) - sHide - 1
lbl = "Added " & CStr(num)
Dim f%
If UBound(Label2) = 5 Then
    For i = 0 To 5
        If Label2(i) <> "" Then f = f + 1
    Next i
    If f = 5 Then sss = 0
End If
XD = XD + 1
End Sub

Private Sub Command2_Click()
If UBound(Label2) = 5 Then sss = UBound(Label2) - 1
Dim i%, t$, num$
num = Label2(Front)
For i = 0 To UBound(Label2)
    t = t & Trim(Label2(i))
Next i
If t = "" Then lbl = "Queue is empty": Exit Sub

For i = Front + 1 To UBound(Label2) - sHide
    'If i = Key Then Exit For
    Label2(i - 1) = Label2(i)
    Label2(i) = ""
Next i
If UBound(Label2) > 5 Then Unload Label2(UBound(Label2)) Else Label2(UBound(Label2) - sHide).Visible = False: sHide = sHide + 1
Label2(Front) = ""
Front = Front - 1
If Front <= 0 Then Front = UBound(Label2) - sHide
Rear = Front + 2
If Rear >= UBound(Label2) - sHide Then Rear = Rear - UBound(Label2) - 1 - sHide
Key = Key - 1
If Key = 0 Then Key = UBound(Label2) - sHide
lbl = "Removed " & CStr(num)
XD = XD - 1
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Dim i%

End Sub

Private Sub Form_Load()
Dim i%
For i = 1 To 100
    Load lbl2(i)
Next i
ReDim Label2(5)
For i = 0 To 5
    Set Label2(i) = lbl2(i)
    Label2(i).Left = i * w
    Label2(i).Visible = True
Next i
Randomize Timer
Dim tmp%
tmp = Int((UBound(Label2) + 1) * Rnd)
Label2(tmp) = fncGet
Label2(IIf(tmp = UBound(Label2), 0, tmp + 1)) = fncGet
Front = IIf(tmp <> 0, tmp - 1, 5)
Rear = IIf(tmp = UBound(Label2), 0, tmp + 1)
Key = IIf(tmp = UBound(Label2) + 1, 0, tmp + 2)
If Key > UBound(Label2) Then Key = Key - UBound(Label2) - 1
XD = 5
End Sub

Private Function fncGet%()
Randomize Timer
fncGet = Int(999 * Rnd) + 1
End Function

Private Sub subChg()
'插入Front後
Dim i%
If UBound(Label2) >= 5 Then
    Load Label2(UBound(Label2) + 1)
    Label2(UBound(Label2)).Left = Label2(UBound(Label2) - 1).Left + w
    Label2(UBound(Label2)) = ""
    Label2(UBound(Label2)).Visible = True
Else
    For i = 1 To UBound(Label2)
        If Label2(i) = "" Then
            Label2(UBound(Label2) - sHide - 1).Visible = True
            sHide = sHide - 1
        End If
    Next i

End If

For i = UBound(Label2) - sHide To Front + 1 Step -1
    If i = Key Then Exit For
    Label2(i) = Label2(i - 1)
    Label2(i - 1) = ""
Next i
Front = Front + 1
If Front > UBound(Label2) - sHide Then Front = 0
Rear = Front + 2
If Rear >= UBound(Label2) - sHide Then Rear = Rear - UBound(Label2) - 1 - sHide
End Sub

Private Sub Timer1_Timer()
Dim i%, f%
For i = 0 To 5
    If Label2(i) <> "" Then f = f + 1
Next i
'If f = 0 And UBound(Label2) = 5 Then
If XD > 5 And f = 6 Then
    'If sss = 0 Then Exit Sub
    'If UBound(Label2) = 5 Then Exit Sub
    For i = 0 To UBound(Label2)
        If Label2(i) <> "" Then Label2(i).Visible = True Else Label2(i).Visible = False
        DoEvents
    Next i
End If
End Sub

