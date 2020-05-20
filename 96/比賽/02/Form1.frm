VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   8325
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6480
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Messages:"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblShow 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   3120
      Width           =   2175
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
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Circular Queue"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Front%, Rear% 'text's index
Const w% = 480 'lbl寬
Dim lbls() As Object
Dim Key%

Private Sub cmdAdd_Click()

Dim num%
Dim i%, flag As Boolean
flag = False
For i = 0 To UBound(lbls)
    If lbls(i) = "" Then flag = True
Next i
num = GetRnd
If Not flag Then
    Key = UBound(lbls)
    Call subChg
    'For i = 0 To UBound(lbls)
    '    If lbls(i) = "" Then Key = i: lbls(i) = num
    'Next i
End If
Dim t%
Dim tmp%
For i = 0 To UBound(lbls)
    If lbls(i) = "" Then t = t + 1: tmp = i
Next i
If t = 1 Then Key = tmp
lbls(Key) = num
If t <> 1 Then Key = Key + 1
If Key > UBound(lbls) Then Key = Key - UBound(lbls) - 1
lblShow = "Added " & CStr(num)
Dim f%
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Function GetRnd%()
Randomize Timer
GetRnd = Int(999 * Rnd) + 1
End Function

Private Sub cmdRemove_Click()
If UBound(lbls) = 0 Then lblShow = "Queue is empty": Exit Sub
If Key > UBound(lbls) Then Key = Key - UBound(lbls) - 1
Dim i%, t$, num$
num = lbls(Front)
For i = 0 To UBound(lbls)
    t = t & Trim(lbls(i))
Next i
If t = "" Then lblShow = "Queue is empty": Exit Sub

For i = Front + 1 To UBound(lbls)
    lbls(i - 1) = lbls(i)
    lbls(i) = ""
Next i
Unload lbls(UBound(lbls))
ReDim Preserve lbls(UBound(lbls) - 1)

'lbls(Front) = ""
Front = Front - 1
If Front <= 0 Then Front = UBound(lbls)
Rear = Front + 2
If Rear >= UBound(lbls) Then Rear = Rear - Front - 1
Key = Key - 1
If Key < 0 Then Key = UBound(lbls)
lblShow = "Removed " & CStr(num)
End Sub

Private Sub Form_Load()
Dim i%
For i = 1 To 100
    Load lbl2(i)
Next i
ReDim lbls(5)
For i = 0 To 5
    Set lbls(i) = lbl2(i)
    lbls(i).Left = i * w
    lbls(i).Visible = True
Next i
Randomize Timer
Dim tmp%
tmp = Int((UBound(lbls) + 1) * Rnd)
lbls(tmp) = GetRnd
lbls(IIf(tmp = UBound(lbls), 0, tmp + 1)) = GetRnd
Front = IIf(tmp <> 0, tmp - 1, 5)
'If Front > UBound(lbls) Then Front = 0
Rear = Front + 2
If Rear >= UBound(lbls) Then Rear = Rear - UBound(lbls) - 1
Key = IIf(tmp = UBound(lbls) + 1, 0, tmp + 2)
If Key > UBound(lbls) Then Key = Key - UBound(lbls) - 1
End Sub

Private Sub subChg()
On Error Resume Next
'插入Front後
Dim i%
ReDim Preserve lbls(UBound(lbls) + 1)
Set lbls(UBound(lbls)) = lbl2(UBound(lbls))
Load lbls(UBound(lbls))
lbls(UBound(lbls)).Left = lbls(UBound(lbls) - 1).Left + w
lbls(UBound(lbls)) = ""
lbls(UBound(lbls)).Visible = True

For i = UBound(lbls) To Front + 1 + IIf(Front <> 1, 1, 0) Step -1
    'If i = Key Then Exit For
    lbls(i) = lbls(i - 1)
    lbls(i - 1) = ""
Next i
Front = Front + 1
If Front > UBound(lbls) Then Front = 0
Rear = Front + 2
If Rear >= UBound(lbls) Then Rear = Rear - UBound(lbls) - 1
'Key = Key + 1
End Sub

Private Sub lbls_Change(Index As Integer)
Debug.Print "XD"
End Sub
