VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���WAV�ɪ��n���i��"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   310
   ScaleMode       =   3  '����
   ScaleWidth      =   636
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   50
      Left            =   1680
      Max             =   4000
      SmallChange     =   10
      TabIndex        =   4
      Top             =   1200
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�}���ɮר���ܪi��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.PictureBox Picture1 
      DrawWidth       =   3
      ForeColor       =   &H0080FF80&
      Height          =   2535
      Left            =   1680
      ScaleHeight     =   2475
      ScaleWidth      =   6075
      TabIndex        =   7
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Label Label4 
      Caption         =   "�n������(��)�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "�ثe��m  (��)�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "�վ���ܪ���m"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "���}�Ҫ�WAV�n����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim wavdata As String
Dim wav1 As String
Option Explicit
'PCM �������Y
Private Type PCMFORM '�H�U��44byte�����Y
    wRiffFormatTag As String * 4 '1~4�s�񪺬ORIFF�r��
    wfdataSize As Long '5~8�s�񪺬O��ư϶��j�p
' NOTE : ��ư϶��j�p=(�ɮפj�p-8)
    wFormatTag As String * 4 '9~12�s�񪺬OWAVE�r��
    wFormatName  As String * 4 '13~16�s�񪺬O�l�϶��ѧO�W��
    wCsize As Long '17~20�s�񪺬O�l�϶��j�p
    wWavefmt As Integer '21~22�s�񪺬O�n�ɮ榡,0x0001��PCM�榡
    wChannels As Integer '23~24�s�񪺬O�n�D��
    wSamplesPerSec As Long '25~28�s�񪺨C����˼�
    wBytePerSec As Long '29~32�s�񪺬O�C���ƶq
' NOTE : �C���ƶq=(�n�D��*�줸��*�C����˼�/8)
    wBytePerSample As Integer '33~34�s�񪺬O�l�϶��줸��
' NOTE : �l�϶��줸��=(�줸��/8)
    wBitsPerSample As Integer '35~36�s�񪺬O���˦�դ���
    wData As String * 4 '�s�񪺬Odata�r��
    wDataSize As Long '����n�ɤj�p
' NOTE : �o�ӭȬ��ɮפj�p��h���Y(44BYTE)�᪺��
End Type

Private Type bdData '�Ψ��x�s8�줸���n�D�����
    bData1 As Byte '���n�D
    bData2 As Byte '�k�n�D
End Type

'�ǤJPicture������ɦW
Public Sub PaintWaveForm(Pic As PictureBox, sFileName As String)
Dim bData() As Byte '�ΨӦs��8bit���n�D���ɸ��
Dim dbData() As bdData '�Ψ��x�s8�줸���n�D�����
Dim pHandle As PCMFORM '���Y
Dim i As Long, lDataSize As Long
Dim t1 As String
Dim a1 As Long

Open sFileName For Binary As #1 '�}�Ҩӷ���z
Get #1, 1, pHandle 'Ū�X44Byte�����Y
If pHandle.wBitsPerSample = 8 Then '�o�O8bit����
    If pHandle.wChannels = 1 Then '���n�D
    t1 = Left(Val(pHandle.wDataSize / 11025), 7) '���o�p���I�᤭�X
        Text3.Text = Val(t1) '���o�n���ɶ�����
        ReDim bData(1 To pHandle.wDataSize) '�]�wBuffer�j�p�HŪ�J�ɮ�
        'pHandle.wDataSize �����Y�i���O��ڥ��ɤj�p
        Get #1, 45, bData '�N�ɮ�Ū��}�C
        lDataSize = UBound(bData) '���o��Ƶ���
        Pic.Scale (0, -256)-(Pic.ScaleWidth, 0) '�]�wø�ϰϰ�y��
        '8�줸���ɭȬO�ɩ�0~256
        a1 = HScroll1.Value / HScroll1.Max * pHandle.wDataSize
    
            For i = 2 To lDataSize
                Pic.Line (10 * i - 10 * a1, -bData(i - 1))-(10 * i - 10 * a1, -bData(i))
            Next
        End If
Else
MsgBox "��J���ɮצW�٤��ORIFF WAVEfmt�BPCM�榡�B8�줸�γ��n�D", , "��J���ɮצW��" & Text1.Text
End If
    Exit Sub
Close
End Sub




Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "�п�J���T���|", , "���|���~"
Else
Picture1.Cls
wav1 = Text1.Text
PaintWaveForm Picture1, wav1 'ø�X�i��
Close
Call sndPlaySound(wav1, 1) '����wav
Text2.Text = "0"
End If
End Sub

Private Sub HScroll1_Change()
Dim data1 As String
Picture1.Cls
wav1 = Text1.Text
wavdata = Left(HScroll1.Value / HScroll1.Max * Text3.Text, 7) '���o�p���I�᤭�X
Text2.Text = wavdata '��ܥثe�ɶ�
PaintWaveForm Picture1, wav1
Close
End Sub


