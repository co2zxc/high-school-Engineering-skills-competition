VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "計算及產生漢明碼"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   5160
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txt2 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   4
      Top             =   1770
      Width           =   2415
   End
   Begin VB.TextBox txt1 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "產生漢明碼"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lbl2 
      Caption         =   "含有漢明碼的訊息"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   2280
   End
   Begin VB.Label lbl1 
      Caption         =   "欲傳遞的訊息"
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
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Dim tt: tt = txt1
txt2 = ""
Dim n%
n = Len(txt1)
If txt1 = "" Then Exit Sub
If n > 11 Then txt2 = "欲傳遞訊息的長度不超過11位元": Exit Sub
Dim i%
For i = 1 To n
    If Mid(txt1, i, 1) <> "0" And Mid(txt1, i, 1) <> "1" Then txt2 = "欲傳遞訊息的值應是0或1": Exit Sub
Next i
Dim bits%, num%
num = 0
Do
    bits = 2 ^ num
    num = num + 1
Loop Until bits * 2 > n
ReDim ary%(n + num)
Dim flag%, tmp%, keys%
flag = 0
tmp = 0
For i = 1 To UBound(ary)
    If i Mod 2 ^ flag = 0 Then
        ary(i) = -1
        flag = flag + 1
    End If
    If ary(i) = -1 Then
        tmp = tmp + 1
    Else
        ary(i) = Mid(txt1, n + 1 - i + tmp, 1)
    End If
    If ary(i) = 1 Then
        keys = keys Xor i
    End If
Next i
tmp = keys

Dim code$
For i = 1 To 4
    code = Trim(Str(tmp Mod 2)) & code
    tmp = tmp \ 2
Next i
tmp = 0
For i = 0 To num - 1
    ary(2 ^ i) = Mid(code, 4 - i, 1)
Next i
For i = 1 To UBound(ary)
    txt2 = ary(i) & txt2
Next i
End Sub
