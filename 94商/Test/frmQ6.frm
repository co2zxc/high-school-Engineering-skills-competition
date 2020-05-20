VERSION 5.00
Begin VB.Form frmQ6 
   AutoRedraw      =   -1  'True
   Caption         =   "回歸方程式"
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
Attribute VB_Name = "frmQ6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim DirPath$, FilePath$, FileNo%
DirPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
FilePath = DirPath & "test6.txt"
FileNo = FreeFile
Dim num%, data%()
Open FilePath For Input As #FileNo
Input #FileNo, num
ReDim data(num - 1, 1) '0:X,1:Y
Dim i%
While Not EOF(FileNo)
    Input #FileNo, data(i, 0), data(i, 1)
    i = i + 1
Wend
Close #FileNo
Dim avg_x!, avg_y!
For i = 0 To UBound(data, 1)
    avg_x = avg_x + data(i, 0) / num
    avg_y = avg_y + data(i, 1) / num
Next i
Dim a!, b!, b1!, b2!, b3!, b4!, b5!
'Dim j%
For i = 0 To UBound(data, 1)
    b1 = b1 + data(i, 0) * data(i, 1)
Next i
b1 = b1 * num
For i = 0 To UBound(data, 1)
    b2 = b2 + data(i, 0)
Next i
For i = 0 To UBound(data, 1)
    b3 = b3 + data(i, 1)
Next i
For i = 0 To UBound(data, 1)
    b4 = b4 + data(i, 0) ^ 2
Next i
b4 = b4 * num
For i = 0 To UBound(data, 1)
    b5 = b5 + data(i, 0)
Next i
b5 = b5 ^ 2
b = (b1 - b2 * b3) / (b4 - b5)
a = avg_y - b * avg_x
b = Int(b * 10 ^ 3 + 0.5) / 10 ^ 3
a = Int(a * 10 ^ 3 + 0.5) / 10 ^ 3
FileNo = FreeFile
FilePath = DirPath & "result6.txt"
Open FilePath For Output As #FileNo
Print #FileNo, "Y=" & IIf(Left(Trim(Str(a)), 1) = ".", "0", "") & Trim(Str(a)) & "+" & IIf(Left(Trim(Str(b)), 1) = ".", "0", "") & Trim(Str(b))
Close #FileNo
MsgBox "Ok!", vbOKOnly
End Sub

Private Sub cndEnd_Click()
End
End Sub
