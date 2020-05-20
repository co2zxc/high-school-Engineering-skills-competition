VERSION 5.00
Begin VB.Form frmQ3 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
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
Attribute VB_Name = "frmQ3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim DirPath$, FilePath$, FileNo%
DirPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
FilePath = DirPath & "test3.txt"
FileNo = FreeFile
Dim num%, data%()
Open FilePath For Input As #FileNo
Input #FileNo, num
ReDim data(num - 1, 1) '0:分子,1:分母
Dim i%
While Not EOF(FileNo)
    Input #FileNo, data(i, 0), data(i, 1)
    i = i + 1
Wend
Close #FileNo
Dim j%, k%, min%
Dim a%, b%, c%, t% '(已通分的)a:整數,b:分子,c:分母
For i = 0 To UBound(data, 1) - 1
    b = data(i, 0) * data(i + 1, 1) + data(i + 1, 0) * data(i, 1)
    c = data(i, 1) * data(i + 1, 1)
    min = IIf(b < c, b, c)
    For k = 2 To min
        If b Mod k = 0 And c Mod k = 0 Then
            b = b / k
            c = c / k
        End If
    Next k
    If b \ c <> 0 Then a = a + b \ c: t = b \ c
    If b > c Then b = b - t * c
    data(i + 1, 0) = b
    data(i + 1, 1) = c
Next i
min = IIf(b < c, b, c)
For k = 2 To min
    If b Mod k = 0 And c Mod k = 0 Then
        b = b / k
        c = c / k
    End If
Next k
If b \ c <> 0 Then a = a + b \ c: t = b \ c
If b > c Then b = b - t * c
FileNo = FreeFile
FilePath = DirPath & "result3.txt"
Open FilePath For Output As #FileNo
Print #FileNo, Trim(Str(a)), Trim(Str(b)), Trim(Str(c))
Close #FileNo
MsgBox "Ok!", vbOKOnly
End Sub

Private Sub cndEnd_Click()
End
End Sub
