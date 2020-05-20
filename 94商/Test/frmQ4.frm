VERSION 5.00
Begin VB.Form frmQ4 
   AutoRedraw      =   -1  'True
   Caption         =   "字串數字算式及總和"
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
      Left            =   3240
      TabIndex        =   1
      Top             =   1898
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   818
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim DirPath$, FilePath$, FileNo%
DirPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
FilePath = DirPath & "test4.txt"
FileNo = FreeFile
Dim tmp, strCount%
Open FilePath For Input As #FileNo
While Not EOF(FileNo)
    Input #FileNo, tmp
Wend
Close #FileNo
tmp = Split(tmp)
strCount = UBound(tmp) + 1
Dim data$(9, 1), i%
For i = 0 To strCount - 1
    data(i, 1) = tmp(i)
Next i
Erase tmp
Dim j%, sum%
For i = 0 To strCount - 1
    For j = 1 To Len(data(i, 1))
        tmp = Mid(data(i, 1), j, 1)
        If tmp >= "0" And tmp <= "9" Then
            tmp = Right(data(i, 1), Len(data(i, 1)) - j + 1)
            data(i, 0) = Trim(Str(Val(tmp)))
            Exit For
        End If
    Next j
Next i
For i = 0 To strCount - 1
    sum = sum + Val(data(i, 0))
Next i
FileNo = FreeFile
FilePath = DirPath & "result4.txt"
Open FilePath For Output As #FileNo
tmp = ""
For i = 0 To strCount - 1
    If data(i, 0) <> "" Then
        tmp = tmp & data(i, 0) & " + "
    End If
Next i
tmp = Left(tmp, Len(tmp) - 2)
tmp = tmp & "= " & Trim(Str(sum))
Print #FileNo, tmp
Close #FileNo
MsgBox "Ok!", vbOKOnly
End Sub

Private Sub cndEnd_Click()
End
End Sub
