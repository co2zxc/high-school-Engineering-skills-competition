VERSION 5.00
Begin VB.Form frmQ7 
   AutoRedraw      =   -1  'True
   Caption         =   "後序運算式"
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
Attribute VB_Name = "frmQ7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'堆資料時，不會堆到底層
Dim data$()

Private Sub cmdStart_Click()
Erase data
Dim DirPath$, FilePath$, FileNo!
DirPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
FilePath = DirPath & "test7.txt"
FileNo = FreeFile
Dim tmp$
Open FilePath For Input As #FileNo
While Not EOF(FileNo)
    Input #FileNo, tmp
Wend
Close #FileNo
Dim i%, s$, a!, b!
ReDim data(0)
For i = 0 To Len(tmp) - 1
    s = Mid(tmp, i + 1, 1)
    If s >= "0" And s <= "9" Then
        If i <> 0 Then ReDim Preserve data(UBound(data) + 1)
        data(UBound(data)) = s
    Else
        b = Val(fncPop)
        a = Val(fncPop)
        Select Case s
        Case "+": a = a + b
        Case "-": a = a - b
        Case "*": a = a * b
        Case "/": a = a / b
        End Select
        fncPush (a)
    End If
Next i
FileNo = FreeFile
FilePath = DirPath & "result7.txt"
Open FilePath For Output As #FileNo
Print #FileNo, data(0)
Close #FileNo
MsgBox "Ok!", vbOKOnly
End Sub

Private Function fncPop$()
Dim i%
i = UBound(data)
Do
    fncPop = data(i)
    i = i - 1
Loop Until fncPop <> ""
data(i + 1) = ""
End Function

Private Function fncPush(result!) As Boolean
Dim i%, tmp$
i = UBound(data)
Do
    tmp = data(i)
    If tmp <> "" Then Exit Do
    If i = 0 Then Exit Do
    i = i - 1
Loop ' Until tmp <> ""
data(i + IIf(tmp <> "", 1, 0)) = Trim(Str(result))
End Function

Private Sub cndEnd_Click()
End
End Sub
