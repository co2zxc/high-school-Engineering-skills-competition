VERSION 5.00
Begin VB.Form Q5 
   AutoRedraw      =   -1  'True
   Caption         =   "A^B MOD C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6495
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   6
      Text            =   "68269"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Text            =   "8315"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Text            =   "42418"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A ^ B MOD C"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "C"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "B"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "A"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Q5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim a#, b#, c#, tmp#, sb$, t$
a = CDbl(Text1(0))
b = CDbl(Text1(1))
c = CDbl(Text1(2))
sb = fnc2(b)
GoTo XX
Dim s#, i#
s = 1
For i = Len(sb) - 1 To 0 Step -1
    
    If s < 5000 Then
        s = s * s Mod c
    Else
        s = fnc(s, s, b, a, c)
    End If
    t = Mid(sb, Len(sb) - i, 1)
    'If t = "1" Then s = a * s Mod c
    If t = "1" Then s = fnc(s, a, b, a, c)
Next i
XX:
s = 1
Text2 = fnc(s, s, b, a, c)
End Sub

Private Function fnc(ByRef s#, ByVal st#, ByVal sb$, ByVal a#, ByVal c#) As Double
Dim i#, t$
s = 1
sb = fnc2(sb)
For i = Len(sb) - 1 To 0 Step -1
    If s < 55000 Then
        s = s * st Mod c
    Else
        s = fnc(s, s, CStr(s), a, c)
    End If
    t = Mid(sb, Len(sb) - i, 1)
    If t = "1" Then s = fnc(s, a, CStr(s), a, c) 's = fnc(s, a, CStr(s), a, c)
    DoEvents
Next i
fnc = s
End Function

Private Function fnc33(ByRef s#, ByVal sb$, ByVal a#, ByVal c#) As Double
Dim i#, t$
s = 1
sb = fnc2(sb)
For i = Len(sb) - 1 To 0 Step -1
    If s < 55000 Then
        s = s * s Mod c
    Else
        s = fnc33(s, CStr(s), a, c)
    End If
    t = Mid(sb, Len(sb) - i, 1)
    If t = "1" Then s = a * s Mod c
Next i
fnc = s
End Function

Private Function fnc2(ByVal x#) As String
Dim i#, r#, t#
Do
    t = x Mod 2
    fnc2 = CStr(t) & fnc2
    x = x \ 2
Loop Until x = 0

End Function
