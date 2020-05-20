VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "顯示WAV檔的聲音波形"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   310
   ScaleMode       =   3  '像素
   ScaleWidth      =   636
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
      Caption         =   "開啟檔案並顯示波形"
      BeginProperty Font 
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
      Caption         =   "聲音長度(秒)："
      BeginProperty Font 
         Name            =   "新細明體"
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
      Caption         =   "目前位置  (秒)："
      BeginProperty Font 
         Name            =   "新細明體"
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
      Caption         =   "調整顯示的位置"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Caption         =   "欲開啟的WAV聲音檔"
      BeginProperty Font 
         Name            =   "新細明體"
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
'PCM 音檔檔頭
Private Type PCMFORM '以下為44byte的檔頭
    wRiffFormatTag As String * 4 '1~4存放的是RIFF字串
    wfdataSize As Long '5~8存放的是資料區塊大小
' NOTE : 資料區塊大小=(檔案大小-8)
    wFormatTag As String * 4 '9~12存放的是WAVE字串
    wFormatName  As String * 4 '13~16存放的是子區塊識別名稱
    wCsize As Long '17~20存放的是子區塊大小
    wWavefmt As Integer '21~22存放的是聲檔格式,0x0001表PCM格式
    wChannels As Integer '23~24存放的是聲道數
    wSamplesPerSec As Long '25~28存放的每秒取樣數
    wBytePerSec As Long '29~32存放的是每秒資料量
' NOTE : 每秒資料量=(聲道數*位元數*每秒取樣數/8)
    wBytePerSample As Integer '33~34存放的是子區塊位元組
' NOTE : 子區塊位元組=(位元數/8)
    wBitsPerSample As Integer '35~36存放的是取樣位組元數
    wData As String * 4 '存放的是data字串
    wDataSize As Long '實際聲檔大小
' NOTE : 這個值為檔案大小減去檔頭(44BYTE)後的值
End Type

Private Type bdData '用來儲存8位元雙聲道的資料
    bData1 As Byte '左聲道
    bData2 As Byte '右聲道
End Type

'傳入Picture物件及檔名
Public Sub PaintWaveForm(Pic As PictureBox, sFileName As String)
Dim bData() As Byte '用來存放8bit單聲道音檔資料
Dim dbData() As bdData '用來儲存8位元單聲道的資料
Dim pHandle As PCMFORM '檔頭
Dim i As Long, lDataSize As Long
Dim t1 As String
Dim a1 As Long

Open sFileName For Binary As #1 '開啟來源檔z
Get #1, 1, pHandle '讀出44Byte的檔頭
If pHandle.wBitsPerSample = 8 Then '這是8bit音檔
    If pHandle.wChannels = 1 Then '單聲道
    t1 = Left(Val(pHandle.wDataSize / 11025), 7) '取得小數點後五碼
        Text3.Text = Val(t1) '取得聲音時間長度
        ReDim bData(1 To pHandle.wDataSize) '設定Buffer大小以讀入檔案
        'pHandle.wDataSize 由檔頭可知是實際生檔大小
        Get #1, 45, bData '將檔案讀到陣列
        lDataSize = UBound(bData) '取得資料筆數
        Pic.Scale (0, -256)-(Pic.ScaleWidth, 0) '設定繪圖區域座標
        '8位元音檔值是界於0~256
        a1 = HScroll1.Value / HScroll1.Max * pHandle.wDataSize
    
            For i = 2 To lDataSize
                Pic.Line (10 * i - 10 * a1, -bData(i - 1))-(10 * i - 10 * a1, -bData(i))
            Next
        End If
Else
MsgBox "輸入的檔案名稱不是RIFF WAVEfmt、PCM格式、8位元及單聲道", , "輸入的檔案名稱" & Text1.Text
End If
    Exit Sub
Close
End Sub




Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "請輸入正確路徑", , "路徑錯誤"
Else
Picture1.Cls
wav1 = Text1.Text
PaintWaveForm Picture1, wav1 '繪出波形
Close
Call sndPlaySound(wav1, 1) '撥放wav
Text2.Text = "0"
End If
End Sub

Private Sub HScroll1_Change()
Dim data1 As String
Picture1.Cls
wav1 = Text1.Text
wavdata = Left(HScroll1.Value / HScroll1.Max * Text3.Text, 7) '取得小數點後五碼
Text2.Text = wavdata '顯示目前時間
PaintWaveForm Picture1, wav1
Close
End Sub


