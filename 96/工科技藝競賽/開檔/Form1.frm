VERSION 5.00
Begin VB.Form Form1 
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DirPath$

Private Sub Command1_Click()
Dim FileName$, strData$, arydata
FileName = "test.txt"
Print IIf(fncInputFile(FileName, strData, arydata), "Succeed!", "Error!")
End Sub

Private Function fncInputFile(ByVal FileName$, ByRef strData$, ByRef arydata As Variant) As Boolean
Dim FilePath$
FilePath = DirPath & FileName
If Dir(FilePath) = "" Then fncInputFile = False: Exit Function '檔案不存在
fncInputFile = True
Dim i%, FileNo%
strData = ""
FileNo = FreeFile
Open FilePath For Input As #FileNo
i = 0
ReDim arydata(0)
While Not EOF(FileNo)
    ReDim Preserve arydata(i)
    Line Input #FileNo, arydata(i)
    strData = strData & IIf(arydata(i) <> "", vbCrLf, "") & arydata(i)
    i = i + 1
Wend
Close #FileNo
End Function

Private Sub Form_Load()
DirPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
End Sub
