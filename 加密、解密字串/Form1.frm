VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7080
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

    Private Sub Command1_Click()

    Dim strPathName As String

    strPathName = ""

    strPathName = InputBox("�п�J�ݭn�R������󧨦W�١J", "�R�����")

    If strPathName = "" Then Exit Sub

    On Error GoTo ErrorHandle

    SetAttr strPathName, vbNormal '����D�n�O���F�ˬd��󧨦W�٪����ĩ�

    RecurseTree strPathName

    Label1.Caption = "���" & strPathName & "�w�g�R���I"

    Exit Sub

ErrorHandle:

    MsgBox "�L�Ī���󧨦W��:" & strPathName

    End Sub

    Sub RecurseTree(CurrPath As String)

    Dim sFileName As String

    Dim newPath As String

    Dim sPath As String

    Static oldPath As String

    sPath = CurrPath & "\"

    sFileName = Dir(sPath, 31) '31���t�q�J31=vbNormal+vbReadOnly+vbHidden+vbSystem+vbVolume+vbDirectory

    Do While sFileName <> ""

    If sFileName <> "." And sFileName <> ".." Then

    If GetAttr(sPath & sFileName) And vbDirectory Then '�p�G�O�ؿ��M���

    newPath = sPath & sFileName

    RecurseTree newPath

    sFileName = Dir(sPath, 31)

    Else

    SetAttr sPath & sFileName, vbNormal

    Kill (sPath & sFileName)

    Label1.Caption = sPath & sFileName '��ܧR���L�{

    sFileName = Dir

    End If

    Else

    sFileName = Dir

    End If

    DoEvents

    Loop

    SetAttr CurrPath, vbNormal

    RmDir CurrPath

    Label1.Caption = CurrPath

    End Sub
