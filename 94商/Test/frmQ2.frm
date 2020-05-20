VERSION 5.00
Begin VB.Form frmQ2 
   AutoRedraw      =   -1  'True
   Caption         =   "支票兌現"
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
Attribute VB_Name = "frmQ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim DirPath$, FilePath$, FileNo%
DirPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
FilePath = DirPath & "test2.txt"
FileNo = FreeFile
Dim num%
Open FilePath For Input As #FileNo
While Not EOF(FileNo)
    Input #FileNo, num
Wend
Close #FileNo
Dim data%(4, 1)
data(0, 0) = 2000
data(1, 0) = 1000
data(2, 0) = 500
data(3, 0) = 200
data(4, 0) = 100
Dim i%, tmp%, sum%
For i = 0 To 4
    If num \ data(i, 0) <> 0 Then
        data(i, 1) = data(i, 1) + num \ data(i, 0)
        num = num - data(i, 0) * data(i, 1)
        sum = sum + data(i, 1)
    End If
Next i
FileNo = FreeFile
FilePath = DirPath & "result2.txt"
Open FilePath For Output As #FileNo
Print #FileNo, Trim(Str(sum))
For i = 0 To 4
    Print #FileNo, Trim(Str(data(i, 0))), Trim(Str(data(i, 1)))
Next i
Close #FileNo
MsgBox "Ok!", vbOKOnly
End Sub

Private Sub cndEnd_Click()
End
End Sub
