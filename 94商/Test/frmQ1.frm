VERSION 5.00
Begin VB.Form frmQ1 
   AutoRedraw      =   -1  'True
   Caption         =   "堆數分配"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cndEnd 
      Caption         =   "End"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1898
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   818
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim DirPath$, FilePath$, FileNo%
DirPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
FilePath = DirPath & "test1.txt"
FileNo = FreeFile
Dim data%(3, 1), i%
Open FilePath For Input As #FileNo
While Not EOF(FileNo)
    Input #FileNo, data(i, 1)
    i = i + 1
Wend
Close #FileNo
data(0, 0) = 1000
data(1, 0) = 500
data(2, 0) = 200
data(3, 0) = 100
Dim tmp%, j%
For i = 0 To 2
    For j = i + 1 To 3
        If data(i, 1) > data(j, 1) Then
            tmp = data(i, 0): data(i, 0) = data(j, 0): data(j, 0) = tmp
            tmp = data(i, 1): data(i, 1) = data(j, 1): data(j, 1) = tmp
        End If
    Next j
Next i
Dim min%
min = 100
For i = 0 To 3
    If data(i, 1) < min Then min = data(i, 1)
Next i
Dim Flag As Boolean, Adder%
Adder = 1
For i = 2 To min
    Flag = True
    For j = 0 To 3
        If data(j, 1) Mod i <> 0 Then Flag = False
    Next j
    If Flag Then
        Adder = Adder * i
        For j = 0 To 3
            data(j, 1) = data(j, 1) / i
        Next j
    End If
Next i
For i = 0 To 2
    For j = i + 1 To 3
        If data(i, 0) < data(j, 0) Then
            tmp = data(i, 0): data(i, 0) = data(j, 0): data(j, 0) = tmp
            tmp = data(i, 1): data(i, 1) = data(j, 1): data(j, 1) = tmp
        End If
    Next j
Next i
FileNo = FreeFile
FilePath = DirPath & "result1.txt"
Open FilePath For Output As #FileNo
Print #FileNo, Trim(Str(Adder))
For i = 0 To 3
    Print #FileNo, Trim(Str(data(i, 0))), Trim(Str(data(i, 1)))
Next i
Close #FileNo
MsgBox "Ok!", vbOKOnly
End Sub

Private Sub cndEnd_Click()
End
End Sub
