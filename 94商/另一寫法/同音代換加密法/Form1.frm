VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "同音代換加密法(Homophonic Substitution Ciphers)"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   6000
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdOk 
      Caption         =   "加密"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2393
      TabIndex        =   2
      Top             =   1148
      Width           =   1215
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      Caption         =   "密文"
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
      TabIndex        =   4
      Top             =   2280
      Width           =   570
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "明文"
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
      TabIndex        =   0
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Dim A%, B%, C%, D%, E%, F%, G%, H%, I%, J%
Dim sA$(8), sB$(2), sC$(3), sD$(4), sE$(12), sF$(2), sG$(2), sH$(6), sI$(6), sJ$(1)
sA(1) = "09": sA(2) = "12": sA(3) = "33": sA(4) = "47": sA(5) = "53": sA(6) = "67": sA(7) = "76": sA(8) = "92"
sB(1) = "48": sB(2) = "81"
sC(1) = "13": sC(2) = "41": sC(3) = "62"
sD(1) = "01": sD(2) = "03": sD(3) = "45": sD(4) = "79"
sE(1) = "14": sE(2) = "16": sE(3) = "24": sE(4) = "44": sE(5) = "46": sE(6) = "55": sE(7) = "57": sE(8) = "64": sE(9) = "74": sE(10) = "82": sE(11) = "87": sE(12) = "98"
sF(1) = "10": sF(2) = "31"
sG(1) = "06": sG(2) = "25"
sH(1) = "23": sH(2) = "39": sH(3) = "50": sH(4) = "56": sH(5) = "65": sH(6) = "68"
sI(1) = "32": sI(2) = "70": sI(3) = "73": sI(4) = "83": sI(5) = "88": sI(6) = "93"
sJ(1) = "15"
Dim k%, answer$, tmp$
For k = 1 To Len(txt1)
    Select Case Mid(txt1, k, 1)
        Case "a", "A": A = A + 1: A = A Mod 8: tmp = sA(IIf(A = 0, 8, A))
        Case "b", "B": B = B + 1: B = B Mod 2: tmp = sB(IIf(B = 0, 2, B))
        Case "c", "C": C = C + 1: C = C Mod 3: tmp = sC(IIf(C = 0, 3, C))
        Case "d", "D": D = D + 1: D = D Mod 4: tmp = sD(IIf(D = 0, 4, D))
        Case "e", "E": E = E + 1: E = E Mod 12: tmp = sE(IIf(E = 0, 12, E))
        Case "f", "F": F = F + 1: F = F Mod 2: tmp = sF(IIf(F = 0, 2, F))
        Case "g", "G": G = G + 1: G = G Mod 2: tmp = sG(IIf(G = 0, 2, G))
        Case "h", "H": H = H + 1: H = H Mod 6: tmp = sH(IIf(H = 0, 6, H))
        Case "i", "I": I = I + 1: I = I Mod 6: tmp = sI(IIf(I = 0, 6, I))
        Case "j", "J": tmp = sJ(1)
    End Select
    answer = answer & " " & tmp
Next k
txt2 = Trim(answer)
End Sub
