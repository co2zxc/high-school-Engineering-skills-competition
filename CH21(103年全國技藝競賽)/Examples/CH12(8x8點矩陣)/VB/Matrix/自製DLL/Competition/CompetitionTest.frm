VERSION 5.00
Begin VB.Form USB 
   Caption         =   "CompetitionTest"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   13680
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   0
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   260
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   1
      Left            =   480
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   259
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   2
      Left            =   480
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   258
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   3
      Left            =   480
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   257
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   4
      Left            =   480
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   256
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   5
      Left            =   480
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   255
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   6
      Left            =   480
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   254
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   7
      Left            =   480
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   253
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   8
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   252
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   9
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   251
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   10
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   250
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   11
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   249
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   12
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   248
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   13
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   247
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   14
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   246
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   15
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   245
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   16
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   244
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   17
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   243
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   18
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   242
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   19
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   241
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   20
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   240
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   21
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   239
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   22
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   238
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   23
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   237
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   24
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   236
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   25
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   235
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   26
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   234
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   27
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   233
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   28
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   232
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   29
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   231
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   30
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   230
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   31
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   229
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   32
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   228
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   33
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   227
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   34
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   226
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   35
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   225
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   36
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   224
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   37
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   223
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   38
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   222
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   39
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   221
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   40
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   220
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   41
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   219
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   42
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   218
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   43
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   217
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   44
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   216
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   45
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   215
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   46
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   214
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   47
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   213
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   48
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   212
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   49
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   211
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   50
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   210
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   51
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   209
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   52
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   208
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   53
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   207
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   54
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   206
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   55
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   205
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   56
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   204
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   57
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   203
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   58
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   202
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   59
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   201
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   60
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   200
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   61
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   199
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   62
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   198
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   63
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   197
      Top             =   3600
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   0
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   196
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   1
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   195
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   2
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   194
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   3
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   193
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   4
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   192
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   5
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   191
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   6
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   190
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   7
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   189
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   8
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   188
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   9
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   187
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   10
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   186
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   11
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   185
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   12
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   184
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   13
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   183
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   14
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   182
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   15
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   181
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   16
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   180
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   17
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   179
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   18
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   178
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   19
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   177
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   20
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   176
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   21
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   175
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   22
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   174
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   23
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   173
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   24
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   172
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   25
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   171
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   26
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   170
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   27
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   169
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   28
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   168
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   29
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   167
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   30
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   166
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   31
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   165
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   32
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   164
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   33
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   163
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   34
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   162
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   35
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   161
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   36
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   160
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   37
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   159
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   38
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   158
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   39
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   157
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   40
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   156
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   41
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   155
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   42
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   154
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   43
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   153
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   44
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   152
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   45
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   151
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   46
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   150
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   47
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   149
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   48
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   148
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   49
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   147
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   50
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   146
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   51
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   145
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   52
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   144
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   53
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   143
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   54
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   142
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   55
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   141
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   56
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   140
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   57
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   139
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   58
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   138
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   59
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   137
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   60
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   136
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   61
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   135
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   62
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   134
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   63
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   133
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   0
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   132
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   1
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   131
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   2
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   130
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   3
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   129
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   4
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   128
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   5
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   127
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   6
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   126
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   7
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   125
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   8
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   124
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   9
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   123
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   10
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   122
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   11
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   121
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   12
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   120
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   13
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   119
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   14
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   118
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   15
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   117
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   16
      Left            =   7920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   116
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   17
      Left            =   7920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   115
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   18
      Left            =   7920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   114
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   19
      Left            =   7920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   113
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   20
      Left            =   7920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   112
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   21
      Left            =   7920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   111
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   22
      Left            =   7920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   110
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   23
      Left            =   7920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   109
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   24
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   108
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   25
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   107
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   26
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   106
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   27
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   105
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   28
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   104
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   29
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   103
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   30
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   102
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   31
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   101
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   32
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   100
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   33
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   99
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   34
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   98
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   35
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   97
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   36
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   96
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   37
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   95
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   38
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   94
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   39
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   93
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   40
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   92
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   41
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   91
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   42
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   90
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   43
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   89
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   44
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   88
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   45
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   87
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   46
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   86
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   47
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   85
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   48
      Left            =   9360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   84
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   49
      Left            =   9360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   83
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   50
      Left            =   9360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   82
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   51
      Left            =   9360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   81
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   52
      Left            =   9360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   80
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   53
      Left            =   9360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   79
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   54
      Left            =   9360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   78
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   55
      Left            =   9360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   77
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   56
      Left            =   9720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   76
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   57
      Left            =   9720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   75
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   58
      Left            =   9720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   74
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   59
      Left            =   9720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   73
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   60
      Left            =   9720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   72
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   61
      Left            =   9720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   71
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   62
      Left            =   9720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   70
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   63
      Left            =   9720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   69
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   0
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   68
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   1
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   67
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   2
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   66
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   3
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   65
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   4
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   64
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   5
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   63
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   6
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   62
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   7
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   61
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   8
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   60
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   9
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   59
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   10
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   58
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   11
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   57
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   12
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   56
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   13
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   55
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   14
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   54
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   15
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   53
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   16
      Left            =   11280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   52
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   17
      Left            =   11280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   51
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   18
      Left            =   11280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   50
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   19
      Left            =   11280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   49
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   20
      Left            =   11280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   48
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   21
      Left            =   11280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   47
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   22
      Left            =   11280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   46
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   23
      Left            =   11280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   45
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   24
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   44
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   25
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   43
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   26
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   42
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   27
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   41
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   28
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   40
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   29
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   39
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   30
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   38
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   31
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   37
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   32
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   36
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   33
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   35
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   34
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   34
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   35
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   33
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   36
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   32
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   37
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   31
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   38
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   30
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   39
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   29
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   40
      Left            =   12360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   28
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   41
      Left            =   12360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   27
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   42
      Left            =   12360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   26
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   43
      Left            =   12360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   25
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   44
      Left            =   12360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   24
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   45
      Left            =   12360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   23
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   46
      Left            =   12360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   22
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   47
      Left            =   12360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   21
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   48
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   20
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   49
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   19
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   50
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   18
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   51
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   17
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   52
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   16
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   53
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   15
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   54
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   14
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   55
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   13
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   56
      Left            =   13080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   12
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   57
      Left            =   13080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   11
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   58
      Left            =   13080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   59
      Left            =   13080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   9
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   60
      Left            =   13080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   8
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   61
      Left            =   13080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   7
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   62
      Left            =   13080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   6
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   63
      Left            =   13080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   5
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000000FF&
      Caption         =   "Red LED"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3840
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   0
      Top             =   120
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '��u�T�w
      Caption         =   " DEVICE OFF "
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   11400
      TabIndex        =   268
      Top             =   4920
      Width           =   2085
   End
   Begin VB.Label Label4 
      Caption         =   "Matrix1 :"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   267
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Matrix2 :"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   266
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Matrix3 :"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   265
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Matrix4 :"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   264
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "Key off "
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   3840
      TabIndex        =   263
      Top             =   4920
      Width           =   1185
   End
   Begin VB.Label Label11 
      Caption         =   "Output Test :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   262
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Input Test :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   261
      Top             =   4920
      Width           =   2055
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ���عϮ�: USB�����]�p�P���γ]�p�J�� Visual Basic �d�ҵ{��-CH12(�\�éM�s��)
'   �{���\��: �Q��PB��PD�X��8��LED�A�H�γz�LPC0~PC5�BPD1�PPD3�X��4��8x8�I�x�}�A�M�ϥ�PD0Ū�������

Const VendorID = &H1234                                        ' �]�w�c��ӽX
Const ProductID = &H2468                                       ' �]�w���~�X
Dim USBHID As New HID
Dim Data(4, 8) As Byte
Dim A As Byte
Dim B As Byte
Dim C As Byte


Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)   ' �T�{USB I/O�����d����O�O�_�s��
     
    If (Check) Then
        Label1.Caption = " DEVICE ON "                  ' DEVICE ON
    Else
        Label1.Caption = " DEVICE OFF "                 ' DEVICE OFF
    End If
    
    USBHID.CloseHIDDevice                               ' ����HID�˸m
    isConnect = Check
End Function

Private Sub Command1_Click(Index As Integer)
    Dim Check As Boolean
    Dim Send(8) As Byte
    If (Command1(Index).BackColor = &HC0C0FF) Then
        Command1(Index).BackColor = &HFF
        Data(0, Index \ 8) = Data(0, Index \ 8) And Not (2 ^ (Index Mod 8)) Or (2 ^ (Index Mod 8))
    Else
        Command1(Index).BackColor = &HC0C0FF
        Data(0, Index \ 8) = Data(0, Index \ 8) And Not (2 ^ (Index Mod 8))
    End If
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)      ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        Send(0) = &H20                                     ' �e�ɽX
        Send(1) = &H0                                      ' 8*8�Ĥ@��
        Send(2) = Index \ 8                                ' 8*8��n��
        Send(3) = Data(0, Index \ 8)                       ' 8*8���
        USBHID.HIDSetReport Send                           ' �ǿ�HID��
    End If
    USBHID.CloseHIDDevice                                  ' ����HID�˸m
End Sub

Private Sub Command2_Click(Index As Integer)
    Dim Check As Boolean
    Dim Send(8) As Byte
    If (Command2(Index).BackColor = &HC0C0FF) Then
        Command2(Index).BackColor = &HFF
        Data(1, Index \ 8) = Data(1, Index \ 8) And Not (2 ^ (Index Mod 8)) Or (2 ^ (Index Mod 8))
    Else
        Command2(Index).BackColor = &HC0C0FF
        Data(1, Index \ 8) = Data(1, Index \ 8) And Not (2 ^ (Index Mod 8))
    End If
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)      ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        Send(0) = &H20                                     ' �e�ɽX
        Send(1) = &H1                                      ' 8*8�ĤG��
        Send(2) = Index \ 8                                ' 8*8��n��
        Send(3) = Data(1, Index \ 8)                       ' 8*8���
        USBHID.HIDSetReport Send                           ' �ǿ�HID��
    End If
    USBHID.CloseHIDDevice                                  ' ����HID�˸m
End Sub

Private Sub Command3_Click(Index As Integer)
    Dim Check As Boolean
    Dim Send(8) As Byte
    If (Command3(Index).BackColor = &HC0C0FF) Then
        Command3(Index).BackColor = &HFF
        Data(2, Index \ 8) = Data(2, Index \ 8) And Not (2 ^ (Index Mod 8)) Or (2 ^ (Index Mod 8))
    Else
        Command3(Index).BackColor = &HC0C0FF
        Data(2, Index \ 8) = Data(2, Index \ 8) And Not (2 ^ (Index Mod 8))
    End If
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)      ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        Send(0) = &H20                                     ' �e�ɽX
        Send(1) = &H2                                      ' 8*8�ĤT��
        Send(2) = Index \ 8                                ' 8*8��n��
        Send(3) = Data(2, Index \ 8)                       ' 8*8���
        USBHID.HIDSetReport Send                           ' �ǿ�HID��
    End If
    USBHID.CloseHIDDevice                                  ' ����HID�˸m
End Sub

Private Sub Command4_Click(Index As Integer)
    Dim Check As Boolean
    Dim Send(8) As Byte
    If (Command4(Index).BackColor = &HC0C0FF) Then
        Command4(Index).BackColor = &HFF
        Data(3, Index \ 8) = Data(3, Index \ 8) And Not (2 ^ (Index Mod 8)) Or (2 ^ (Index Mod 8))
    Else
        Command4(Index).BackColor = &HC0C0FF
        Data(3, Index \ 8) = Data(3, Index \ 8) And Not (2 ^ (Index Mod 8))
    End If
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)      ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        Send(0) = &H20                                     ' �e�ɽX
        Send(1) = &H3                                      ' 8*8�ĥ|��
        Send(2) = Index \ 8                                ' 8*8��n��
        Send(3) = Data(3, Index \ 8)                       ' 8*8���
        USBHID.HIDSetReport Send                           ' �ǿ�HID��
    End If
    USBHID.CloseHIDDevice                                  ' ����HID�˸m
End Sub

Private Sub Command5_Click()
    Dim i As Byte
    Dim Send(8) As Byte
    For i = 0 To 63
        Command1(i).BackColor = &HC0C0FF
    Next i
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)          ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        For i = 0 To 7
            Data(0, i) = 0
            Send(0) = &H20                                     ' �e�ɽX
            Send(1) = &H0                                      ' 8*8�Ĥ@��
            Send(2) = i                                        ' 8*8��n��
            Send(3) = Data(0, i)                               ' 8*8���
            USBHID.HIDSetReport Send                           ' �ǿ�HID��
        Next i
    End If
    USBHID.CloseHIDDevice                                      ' ����HID�˸m
End Sub

Private Sub Command6_Click()
    Dim i As Byte
    Dim Send(8) As Byte
    For i = 0 To 63
        Command2(i).BackColor = &HC0C0FF
    Next i
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)          ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        For i = 0 To 7
            Data(1, i) = 0
            Send(0) = &H20                                     ' �e�ɽX
            Send(1) = &H1                                      ' 8*8�ĤG��
            Send(2) = i                                        ' 8*8��n��
            Send(3) = Data(1, i)                               ' 8*8���
            USBHID.HIDSetReport Send                           ' �ǿ�HID��
        Next i
    End If
    USBHID.CloseHIDDevice                                      ' ����HID�˸m
End Sub

Private Sub Command7_Click()
    Dim i As Byte
    Dim Send(8) As Byte
    For i = 0 To 63
        Command3(i).BackColor = &HC0C0FF
    Next i
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)          ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        For i = 0 To 7
            Data(2, i) = 0
            Send(0) = &H20                                     ' �e�ɽX
            Send(1) = &H2                                      ' 8*8�ĤT��
            Send(2) = i                                        ' 8*8��n��
            Send(3) = Data(2, i)                               ' 8*8���
            USBHID.HIDSetReport Send                           ' �ǿ�HID��
        Next i
    End If
    USBHID.CloseHIDDevice                                      ' ����HID�˸m
End Sub

Private Sub Command8_Click()
    Dim i As Byte
    Dim Send(8) As Byte
    For i = 0 To 63
        Command4(i).BackColor = &HC0C0FF
    Next i
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)          ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        For i = 0 To 7
            Data(3, i) = 0
            Send(0) = &H20                                     ' �e�ɽX
            Send(1) = &H3                                      ' 8*8�ĥ|��
            Send(2) = i                                        ' 8*8��n��
            Send(3) = Data(3, i)                               ' 8*8���
            USBHID.HIDSetReport Send                           ' �ǿ�HID��
        Next i
    End If
    USBHID.CloseHIDDevice                                      ' ����HID�˸m
End Sub

Private Sub Command9_Click()
    A = 1
    B = 0
    C = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Send(8) As Byte
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)          ' �T�{USB I/O�����d����O�O�_�s��
    If (Check) Then
        For i = 0 To 7
            Send(0) = &H20                                     ' �e�ɽX
            Send(1) = &H0                                      ' 8*8�Ĥ@��
            Send(2) = i                                        ' 8*8��n��
            USBHID.HIDSetReport Send                           ' �ǿ�HID��
            
            Send(1) = &H1                                      ' 8*8�ĤG��
            USBHID.HIDSetReport Send                           ' �ǿ�HID��
            
            Send(1) = &H2                                      ' 8*8�ĤT��
            USBHID.HIDSetReport Send                           ' �ǿ�HID��
            
            Send(1) = &H3                                      ' 8*8�ĥ|��
            USBHID.HIDSetReport Send                           ' �ǿ�HID��
        Next i
    End If
    USBHID.CloseHIDDevice                                      ' ����HID�˸m
End Sub

Private Sub Timer1_Timer()
    Dim Send(8) As Byte
    Dim ReadKeyData() As Byte
    isConnect
    Check = USBHID.OpenHIDDevice(VendorID, ProductID)
    If (A) Then
        C = C + 1
        If (C Mod 10 = 0) Then
            Send(0) = &H21                                     ' �e�ɽX
            Send(1) = 2 ^ B                                    ' LED���
            USBHID.HIDSetReport Send                           ' �ǿ�HID��
            B = B + 1
            If (B = 8) Then
                A = 0
            End If
        End If
    Else
        C = C + 1
        If (C Mod 10 = 0) Then
            Send(0) = &H21                                     ' �e�ɽX
            Send(1) = &H0                                      ' LED���
            USBHID.HIDSetReport Send                           ' �ǿ�HID��
            C = 0
        End If
    End If
    
    ReadKeyData() = USBHID.HidGetReport                        ' Ū��HID���
    USBHID.CloseHIDDevice                                      ' ����HID�˸m

    Select Case (ReadKeyData(1) And &H1)
        Case 0
            Label6.Caption = " KEY ON "
        Case 1
            Label6.Caption = " KEY OFF "
    End Select
End Sub

