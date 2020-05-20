VERSION 5.00
Begin VB.Form frmQ5 
   AutoRedraw      =   -1  'True
   Caption         =   "³Ì¤pÂ÷®t"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   818
      Width           =   1215
   End
   Begin VB.CommandButton cndEnd 
      Caption         =   "End"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1898
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim DirPath$, FilePath$, FileNo!
DirPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
FilePath = DirPath & "test5.txt"
FileNo = FreeFile
Dim num%, data%()
Open FilePath For Input As #FileNo
Input #FileNo, num
ReDim data(num - 1)
Dim i%
While Not EOF(FileNo)
    Input #FileNo, data(i)
    i = i + 1
Wend
Close #FileNo
Dim Avg!
Avg = 1
For i = 0 To UBound(data)
    Avg = Avg * data(i)
Next i
Avg = Avg ^ (1 / num)
Dim tmp%, j%
For i = 0 To UBound(data) - 1
    For j = i + 1 To UBound(data)
        If data(i) > data(j) Then
            tmp = data(i)
            data(i) = data(j)
            data(j) = tmp
        End If
    Next j
Next i
Dim tmp1!, tmp2%
tmp1 = 99
For i = 0 To UBound(data) - 1
    If Abs(data(i) - Avg) < tmp1 Then
        tmp1 = Abs(data(i) - Avg)
        tmp2 = i
    End If
Next i

Dim s1$, s31!, t!
t = 0
For i = 0 To tmp2
    s1 = s1 & Trim(Str(data(i))) & " "
    t = t + data(i)
Next i
s1 = Left(s1, Len(s1) - 1)
t = t / (tmp2 + 1)
For i = 0 To tmp2
    s31 = s31 + Abs(t - data(i))
Next i

Dim s2$, s32!
t = 0
For i = (tmp2 + 1) To UBound(data)
    s2 = s2 & Trim(Str(data(i))) & " "
    t = t + data(i)
Next i
s2 = Left(s2, Len(s2) - 1)
t = t / (UBound(data) - tmp2)
For i = (tmp2 + 1) To UBound(data)
    s32 = s32 + Abs(t - data(i))
Next i

Dim s3!
s3 = s31 + s32

FileNo = FreeFile
FilePath = DirPath & "result5.txt"
Open FilePath For Output As #FileNo
Print #FileNo, s1
Print #FileNo, s2
Print #FileNo, Trim(Str(s3))
Close #FileNo
MsgBox "Ok!", vbOKOnly
End Sub

Private Sub cndEnd_Click()
End
End Sub
