VERSION 5.00
Begin VB.Form USB 
   Caption         =   "CompetitionTest"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   14655
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text6 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   318
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Delete"
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
      Style           =   1  '圖片外觀
      TabIndex        =   317
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Claear"
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
      Style           =   1  '圖片外觀
      TabIndex        =   316
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADD"
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
      Style           =   1  '圖片外觀
      TabIndex        =   315
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   7
      Left            =   600
      Style           =   1  '圖片外觀
      TabIndex        =   314
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   6
      Left            =   1457
      Style           =   1  '圖片外觀
      TabIndex        =   313
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   5
      Left            =   2314
      Style           =   1  '圖片外觀
      TabIndex        =   312
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   4
      Left            =   3171
      Style           =   1  '圖片外觀
      TabIndex        =   311
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   3
      Left            =   4028
      Style           =   1  '圖片外觀
      TabIndex        =   310
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   4885
      Style           =   1  '圖片外觀
      TabIndex        =   309
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   5742
      Style           =   1  '圖片外觀
      TabIndex        =   308
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   6600
      Style           =   1  '圖片外觀
      TabIndex        =   307
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   7
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   306
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   6
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   305
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   5
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   304
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   4
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   303
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   3
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   302
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   2
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   301
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   1
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   300
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   0
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   299
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   7
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   298
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   6
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   297
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   5
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   296
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   4
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   295
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   3
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   294
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   2
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   293
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   1
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   292
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   0
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   291
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   7
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   290
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   6
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   289
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   5
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   288
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   4
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   287
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   3
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   286
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   2
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   285
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   1
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   284
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   0
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   283
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   7
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   282
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   6
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   281
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   5
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   280
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   4
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   279
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   3
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   278
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   2
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   277
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   276
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      Height          =   495
      Index           =   0
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   275
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   274
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000FF00&
      Caption         =   "送出數字"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      Style           =   1  '圖片外觀
      TabIndex        =   273
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command13 
      Caption         =   "全亮"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   272
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "全亮"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   271
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "全亮"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   270
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "全亮"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   269
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000000FF&
      Caption         =   "Red LED"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3960
      Style           =   1  '圖片外觀
      TabIndex        =   265
      Top             =   240
      Width           =   1725
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   260
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   259
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   258
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   257
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   63
      Left            =   13320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   256
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   62
      Left            =   13320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   255
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   61
      Left            =   13320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   254
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   60
      Left            =   13320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   253
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   59
      Left            =   13320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   252
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   58
      Left            =   13320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   251
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   57
      Left            =   13320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   250
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   56
      Left            =   13320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   249
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   55
      Left            =   12960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   248
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   54
      Left            =   12960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   247
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   53
      Left            =   12960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   246
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   52
      Left            =   12960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   245
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   51
      Left            =   12960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   244
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   50
      Left            =   12960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   243
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   49
      Left            =   12960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   242
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   48
      Left            =   12960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   241
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   47
      Left            =   12600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   240
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   46
      Left            =   12600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   239
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   45
      Left            =   12600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   238
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   44
      Left            =   12600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   237
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   43
      Left            =   12600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   236
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   42
      Left            =   12600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   235
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   41
      Left            =   12600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   234
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   40
      Left            =   12600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   233
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   39
      Left            =   12240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   232
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   38
      Left            =   12240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   231
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   37
      Left            =   12240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   230
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   36
      Left            =   12240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   229
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   35
      Left            =   12240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   228
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   34
      Left            =   12240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   227
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   33
      Left            =   12240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   226
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   32
      Left            =   12240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   225
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   31
      Left            =   11880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   224
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   30
      Left            =   11880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   223
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   29
      Left            =   11880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   222
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   28
      Left            =   11880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   221
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   27
      Left            =   11880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   220
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   26
      Left            =   11880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   219
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   25
      Left            =   11880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   218
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   24
      Left            =   11880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   217
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   23
      Left            =   11520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   216
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   22
      Left            =   11520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   215
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   21
      Left            =   11520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   214
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   20
      Left            =   11520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   213
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   19
      Left            =   11520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   212
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   18
      Left            =   11520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   211
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   17
      Left            =   11520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   210
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   16
      Left            =   11520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   209
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   15
      Left            =   11160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   208
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   14
      Left            =   11160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   207
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   13
      Left            =   11160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   206
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   12
      Left            =   11160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   205
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   11
      Left            =   11160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   204
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   10
      Left            =   11160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   203
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   9
      Left            =   11160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   202
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   8
      Left            =   11160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   201
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   7
      Left            =   10800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   200
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   6
      Left            =   10800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   199
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   5
      Left            =   10800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   198
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   4
      Left            =   10800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   197
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   3
      Left            =   10800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   196
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   2
      Left            =   10800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   195
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   1
      Left            =   10800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   194
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   0
      Left            =   10800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   193
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   63
      Left            =   9960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   192
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   62
      Left            =   9960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   191
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   61
      Left            =   9960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   190
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   60
      Left            =   9960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   189
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   59
      Left            =   9960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   188
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   58
      Left            =   9960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   187
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   57
      Left            =   9960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   186
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   56
      Left            =   9960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   185
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   55
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   184
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   54
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   183
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   53
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   182
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   52
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   181
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   51
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   180
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   50
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   179
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   49
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   178
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   48
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   177
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   47
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   176
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   46
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   175
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   45
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   174
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   44
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   173
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   43
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   172
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   42
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   171
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   41
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   170
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   40
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   169
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   39
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   168
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   38
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   167
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   37
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   166
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   36
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   165
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   35
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   164
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   34
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   163
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   33
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   162
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   32
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   161
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   31
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   160
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   30
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   159
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   29
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   158
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   28
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   157
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   27
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   156
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   26
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   155
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   25
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   154
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   24
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   153
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   23
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   152
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   22
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   151
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   21
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   150
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   20
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   149
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   19
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   148
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   18
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   147
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   17
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   146
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   16
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   145
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   15
      Left            =   7800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   144
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   14
      Left            =   7800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   143
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   13
      Left            =   7800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   142
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   12
      Left            =   7800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   141
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   11
      Left            =   7800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   140
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   10
      Left            =   7800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   139
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   9
      Left            =   7800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   138
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   8
      Left            =   7800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   137
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   7
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   136
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   6
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   135
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   5
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   134
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   4
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   133
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   3
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   132
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   2
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   131
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   1
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   130
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   0
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   129
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   63
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   128
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   62
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   127
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   61
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   126
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   60
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   125
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   59
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   124
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   58
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   123
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   57
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   122
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   56
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   121
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   55
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   120
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   54
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   119
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   53
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   118
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   52
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   117
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   51
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   116
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   50
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   115
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   49
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   114
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   48
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   113
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   47
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   112
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   46
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   111
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   45
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   110
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   44
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   109
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   43
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   108
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   42
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   107
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   41
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   106
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   40
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   105
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   39
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   104
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   38
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   103
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   37
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   102
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   36
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   101
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   35
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   100
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   34
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   99
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   33
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   98
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   32
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   97
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   31
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   96
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   30
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   95
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   29
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   94
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   28
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   93
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   27
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   92
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   26
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   91
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   25
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   90
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   24
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   89
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   23
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   88
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   22
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   87
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   21
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   86
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   20
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   85
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   19
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   84
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   18
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   83
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   17
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   82
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   16
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   81
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   15
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   80
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   14
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   79
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   13
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   78
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   12
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   77
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   11
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   76
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   10
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   75
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   9
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   74
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   8
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   73
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   7
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   72
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   6
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   71
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   5
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   70
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   4
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   69
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   3
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   68
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   2
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   67
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   1
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   66
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   0
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   65
      Top             =   1200
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   63
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   64
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   62
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   63
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   61
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   62
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   60
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   61
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   59
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   60
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   58
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   59
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   57
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   58
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   56
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   57
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   55
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   56
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   54
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   55
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   53
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   54
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   52
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   53
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   51
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   52
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   50
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   51
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   49
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   50
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   48
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   49
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   47
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   48
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   46
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   47
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   45
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   46
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   44
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   45
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   43
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   44
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   42
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   43
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   41
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   42
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   40
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   41
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   39
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   40
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   38
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   39
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   37
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   38
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   36
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   37
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   35
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   36
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   34
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   35
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   33
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   34
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   32
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   33
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   31
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   32
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   30
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   31
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   29
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   30
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   28
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   29
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   27
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   28
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   26
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   25
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   24
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   23
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   22
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   21
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   20
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   19
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   18
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   17
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   16
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   15
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   14
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   13
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   12
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   11
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   10
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   9
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   8
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   7
      Left            =   720
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   6
      Left            =   720
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   5
      Left            =   720
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   4
      Left            =   720
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   3
      Left            =   720
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   2
      Left            =   720
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   1
      Left            =   720
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   0
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   1200
      Width           =   375
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
      Left            =   720
      TabIndex        =   268
      Top             =   5880
      Width           =   2055
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
      Left            =   720
      TabIndex        =   267
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   " KEY OFF "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   3900
      TabIndex        =   266
      Top             =   5880
      Width           =   1545
   End
   Begin VB.Label Label5 
      Caption         =   "Matrix4 :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10800
      TabIndex        =   264
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Matrix3 :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   263
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Matrix2 :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   262
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Matrix1 :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   261
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   " DEVICE OFF "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   11640
      TabIndex        =   8
      Top             =   5880
      Width           =   2085
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Data(4, 8) As Byte
Dim a As Byte
Dim b As Byte
Dim d As Byte
Dim LedData(20) As Byte
Dim ReadKeyData(8) As Byte
Dim number(10, 8) As Byte
Dim t1 As String
Dim bit_number  As Byte
Dim bit_count As Integer
Private Function isConnect() As Boolean
    Dim Check As Boolean
    Check = OpenUsbDevice(VendorID, ProductID)          ' 確認USB devie是否連線
     
    If (Check) Then
        Label1.Caption = " DEVICE ON "                  ' DEVICE ON
    Else
        Label1.Caption = " DEVICE OFF "                 ' DEVICE OFF
    End If
    
    CloseUsbDevice                                      ' 關閉HID裝置
    isConnect = Check
End Function

Private Sub Command1_Click(Index As Integer)
    Dim Check As Boolean
    Dim i As Integer
    
    If d = 1 Then
        For i = 0 To 63
        Command1(i).BackColor = &HC0C0FF
        Command2(i).BackColor = &HC0C0FF
        Command3(i).BackColor = &HC0C0FF
        Command4(i).BackColor = &HC0C0FF
        Next i
        d = 0
        Check = OpenUsbDevice(VendorID, ProductID)
        If (Check) Then
        For i = 0 To 7
            Data(0, i) = 0
            OutDataEightByte &H20, 0, i, Data(0, i), 0, 0, 0, 0
            Data(1, i) = 0
            OutDataEightByte &H20, 1, i, Data(1, i), 0, 0, 0, 0
            Data(2, i) = 0
            OutDataEightByte &H20, 2, i, Data(2, i), 0, 0, 0, 0
            Data(3, i) = 0
            OutDataEightByte &H20, 3, i, Data(3, i), 0, 0, 0, 0
        Next i
        End If
        CloseUsbDevice
    End If
    Dim x: x = Index \ 8
    Dim y: y = Index Mod 8
    Dim D1: D1 = 2 ^ y
    Dim D2: D2 = Not D1
    If (Command1(Index).BackColor = &HC0C0FF) Then
        Command1(Index).BackColor = &HFF
        Data(0, x) = Data(0, x) Or D1
        'byte_number ++++++++
        Text1(x).Text = Data(0, x)
        Text1(x).Text = Hex(Text1(x).Text)
        '++++++++  byte_number
    Else
        Command1(Index).BackColor = &HC0C0FF
        Data(0, x) = Data(0, x) And D2
        'byte_number --------
        Text1(x).Text = Data(0, x)
        Text1(x).Text = Hex(Text1(x).Text)
        '--------  byte_number
    End If
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        OutDataEightByte &H20, 0, x, Data(0, x), 0, 0, 0, 0
    End If
    CloseUsbDevice
       
End Sub

Private Sub Command15_Click(Index As Integer)
    If Command15(Index).BackColor = &H0 Then
        Command15(Index).BackColor = &HFF
         bit_number = bit_number + 2 ^ Index
    Else
        Command15(Index).BackColor = &H0
        bit_number = bit_number - 2 ^ Index
    End If
End Sub

Private Sub Command16_Click()
    Text6.Text = Text6.Text & bit_count & ">>>   " & Hex(bit_number) & vbNewLine
    bit_count = bit_count + 1
End Sub

Private Sub Command17_Click()
    Text6.Text = ""
    bit_count = 0
End Sub

Private Sub Command18_Click()
    If Text6.Text <> "" Then
        bit_count = bit_count - 1
        Text6.Text = Mid(Text6.Text, 1, InStr(Text6.Text, bit_count & ">>>   ") - 1)
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
    Dim Check As Boolean
    Dim i As Integer
    If d = 1 Then
        For i = 0 To 63
        Command1(i).BackColor = &HC0C0FF
        Command2(i).BackColor = &HC0C0FF
        Command3(i).BackColor = &HC0C0FF
        Command4(i).BackColor = &HC0C0FF
        Next i
        d = 0
        Check = OpenUsbDevice(VendorID, ProductID)
        If (Check) Then
        For i = 0 To 7
            Data(0, i) = 0
            OutDataEightByte &H20, 0, i, Data(0, i), 0, 0, 0, 0
            Data(1, i) = 0
            OutDataEightByte &H20, 1, i, Data(1, i), 0, 0, 0, 0
            Data(2, i) = 0
            OutDataEightByte &H20, 2, i, Data(2, i), 0, 0, 0, 0
            Data(3, i) = 0
            OutDataEightByte &H20, 3, i, Data(3, i), 0, 0, 0, 0
        Next i
        End If
        CloseUsbDevice
    End If
    Dim x: x = Index \ 8
    Dim y: y = Index Mod 8
    Dim D1: D1 = 2 ^ y
    Dim D2: D2 = Not D1
    If (Command2(Index).BackColor = &HC0C0FF) Then
        Command2(Index).BackColor = &HFF
        Data(1, x) = Data(1, x) Or D1
        'byte_number ++++++++
        Text2(x).Text = Data(1, x)
        Text2(x).Text = Hex(Text2(x).Text)
        '++++++++  byte_number
    Else
        Command2(Index).BackColor = &HC0C0FF
        Data(1, x) = Data(1, x) And D2
        'byte_number --------
        Text2(x).Text = Data(1, x)
        Text2(x).Text = Hex(Text2(x).Text)
        '--------  byte_number
    End If
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        OutDataEightByte &H20, 1, x, Data(1, x), 0, 0, 0, 0
    End If
    CloseUsbDevice
End Sub

Private Sub Command3_Click(Index As Integer)
    Dim Check As Boolean
    Dim i As Integer
    If d = 1 Then
        For i = 0 To 63
        Command1(i).BackColor = &HC0C0FF
        Command2(i).BackColor = &HC0C0FF
        Command3(i).BackColor = &HC0C0FF
        Command4(i).BackColor = &HC0C0FF
        Next i
        d = 0
        Check = OpenUsbDevice(VendorID, ProductID)
        If (Check) Then
        For i = 0 To 7
            Data(0, i) = 0
            OutDataEightByte &H20, 0, i, Data(0, i), 0, 0, 0, 0
            Data(1, i) = 0
            OutDataEightByte &H20, 1, i, Data(1, i), 0, 0, 0, 0
            Data(2, i) = 0
            OutDataEightByte &H20, 2, i, Data(2, i), 0, 0, 0, 0
            Data(3, i) = 0
            OutDataEightByte &H20, 3, i, Data(3, i), 0, 0, 0, 0
        Next i
        End If
        CloseUsbDevice
    End If
    Dim x: x = Index \ 8
    Dim y: y = Index Mod 8
    Dim D1: D1 = 2 ^ y
    Dim D2: D2 = Not D1
    If (Command3(Index).BackColor = &HC0C0FF) Then
        Command3(Index).BackColor = &HFF
        Data(2, x) = Data(2, x) Or D1
        'byte_number ++++++++
        Text3(x).Text = Data(2, x)
        Text3(x).Text = Hex(Text3(x).Text)
        '++++++++  byte_number
    Else
        Command3(Index).BackColor = &HC0C0FF
        Data(2, x) = Data(2, x) And D2
        'byte_number --------
        Text3(x).Text = Data(2, x)
        Text3(x).Text = Hex(Text3(x).Text)
        '--------  byte_number
    End If
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        OutDataEightByte &H20, 2, x, Data(2, x), 0, 0, 0, 0
    End If
    CloseUsbDevice
End Sub

Private Sub Command4_Click(Index As Integer)
    Dim Check As Boolean
    Dim i As Integer
    If d = 1 Then
        For i = 0 To 63
        Command1(i).BackColor = &HC0C0FF
        Command2(i).BackColor = &HC0C0FF
        Command3(i).BackColor = &HC0C0FF
        Command4(i).BackColor = &HC0C0FF
        Next i
        d = 0
        Check = OpenUsbDevice(VendorID, ProductID)
        If (Check) Then
        For i = 0 To 7
            Data(0, i) = 0
            OutDataEightByte &H20, 0, i, Data(0, i), 0, 0, 0, 0
            Data(1, i) = 0
            OutDataEightByte &H20, 1, i, Data(1, i), 0, 0, 0, 0
            Data(2, i) = 0
            OutDataEightByte &H20, 2, i, Data(2, i), 0, 0, 0, 0
            Data(3, i) = 0
            OutDataEightByte &H20, 3, i, Data(3, i), 0, 0, 0, 0
        Next i
        End If
        CloseUsbDevice
    End If
    Dim x: x = Index \ 8
    Dim y: y = Index Mod 8
    Dim D1: D1 = 2 ^ y
    Dim D2: D2 = Not D1
    If (Command4(Index).BackColor = &HC0C0FF) Then
        Command4(Index).BackColor = &HFF
        Data(3, x) = Data(3, x) Or D1
        'byte_number ++++++++
        Text4(x).Text = Data(3, x)
        Text4(x).Text = Hex(Text4(x).Text)
        '++++++++  byte_number
    Else
        Command4(Index).BackColor = &HC0C0FF
        Data(3, x) = Data(3, x) And D2
        'byte_number ---------
        Text4(x).Text = Data(3, x)
        Text4(x).Text = Hex(Text4(x).Text)
        '--------  byte_number
    End If
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        OutDataEightByte &H20, 3, x, Data(3, x), 0, 0, 0, 0
    End If
    CloseUsbDevice
End Sub

Private Sub Command5_Click()
    Dim i As Byte
    For i = 0 To 63
        Command1(i).BackColor = &HC0C0FF
    Next i
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            Data(0, i) = 0
            OutDataEightByte &H20, 0, i, Data(0, i), 0, 0, 0, 0
        Next i
    End If
    CloseUsbDevice
    'byte_number 0000000
    For i = 0 To 7
        Text1(i).Text = 0
    Next i
    '0000000 byte_number
End Sub

Private Sub Command6_Click()
    Dim i As Byte
    For i = 0 To 63
        Command2(i).BackColor = &HC0C0FF
    Next i
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            Data(1, i) = 0
            OutDataEightByte &H20, 1, i, Data(1, i), 0, 0, 0, 0
        Next i
    End If
    CloseUsbDevice
    'byte_number 0000000
    For i = 0 To 7
        Text2(i).Text = 0
    Next i
    '0000000 byte_number
End Sub

Private Sub Command7_Click()
    Dim i As Byte
    For i = 0 To 63
        Command3(i).BackColor = &HC0C0FF
    Next i
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            Data(2, i) = 0
            OutDataEightByte &H20, 2, i, Data(2, i), 0, 0, 0, 0
        Next i
    End If
    CloseUsbDevice
    'byte_number 0000000
    For i = 0 To 7
        Text3(i).Text = 0
    Next i
    '0000000 byte_number
End Sub

Private Sub Command8_Click()
    Dim i As Byte
    For i = 0 To 63
        Command4(i).BackColor = &HC0C0FF
    Next i
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            Data(3, i) = 0
            OutDataEightByte &H20, 3, i, Data(3, i), 0, 0, 0, 0
        Next i
    End If
    CloseUsbDevice
    'byte_number 0000000
    For i = 0 To 7
        Text4(i).Text = 0
    Next i
    '0000000 byte_number
End Sub

Private Sub Command9_Click()
    a = 1
    b = 0
    c = 0
End Sub

Private Sub Command10_Click()
    Dim i As Byte
    For i = 0 To 63
        Command1(i).BackColor = &HFF
    Next i
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            Data(0, i) = 255
            OutDataEightByte &H20, 0, i, Data(0, i), 0, 0, 0, 0
        Next i
    End If
    CloseUsbDevice
    'byte_number FFFFFFF
    For i = 0 To 7
        Text1(i).Text = "FF"
    Next i
    'FFFFFFF byte_number
End Sub

Private Sub Command11_Click()
    Dim i As Byte
    For i = 0 To 63
        Command2(i).BackColor = &HFF
    Next i
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            Data(1, i) = 255
            OutDataEightByte &H20, 1, i, Data(1, i), 0, 0, 0, 0
        Next i
    End If
    CloseUsbDevice
    'byte_number FFFFFFF
    For i = 0 To 7
        Text2(i).Text = "FF"
    Next i
    'FFFFFFF byte_number
End Sub

Private Sub Command12_Click()
    Dim i As Byte
    For i = 0 To 63
        Command3(i).BackColor = &HFF
    Next i
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            Data(2, i) = 255
            OutDataEightByte &H20, 2, i, Data(2, i), 0, 0, 0, 0
        Next i
    End If
    CloseUsbDevice
    'byte_number FFFFFFF
    For i = 0 To 7
        Text3(i).Text = "FF"
    Next i
    'FFFFFFF byte_number
End Sub

Private Sub Command13_Click()
    Dim i As Byte
    For i = 0 To 63
        Command4(i).BackColor = &HFF
    Next i
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            Data(3, i) = 255
            OutDataEightByte &H20, 3, i, Data(3, i), 0, 0, 0, 0
        Next i
    End If
    CloseUsbDevice
    'byte_number FFFFFFF
    For i = 0 To 7
        Text4(i).Text = "FF"
    Next i
    'FFFFFFF byte_number
End Sub

Private Sub Command14_Click()
        Dim Check As Boolean
        ckeck = OpenUsbDevice(VendorID, ProductID)
        Dim L, i, j   As Integer: L = Len(Text5.Text)
        For i = 0 To L - 1
            Dim m As Integer: m = Val(Mid(Text5.Text, i + 1, 1))
            For j = 0 To 7
               OutDataEightByte &H20, i, j, number(m, j), 0, 0, 0, 0
            Next j
        Next i
        For i = i To 3
            For j = 0 To 7
               OutDataEightByte &H20, i, j, &H0, 0, 0, 0, 0
            Next j
        Next i
      CloseUsbDevice
      For i = 0 To 63
        Command1(i).BackColor = &HC0C0FF
        Command2(i).BackColor = &HC0C0FF
        Command3(i).BackColor = &HC0C0FF
        Command4(i).BackColor = &HC0C0FF
      Next i
      d = 1
     'byte_number 0000000
      For i = 0 To 7
        Text1(i).Text = 0
        Text2(i).Text = 0
        Text3(i).Text = 0
        Text4(i).Text = 0
      Next i
      '0000000 byte_number
      
End Sub

Private Sub Form_Load()

number(0, 2) = &H1F
number(0, 3) = &H11
number(0, 4) = &H11
number(0, 5) = &H1F

number(1, 5) = &H1F

number(2, 2) = &H1D
number(2, 3) = &H15
number(2, 4) = &H15
number(2, 5) = &H17

number(3, 2) = &H15
number(3, 3) = &H15
number(3, 4) = &H15
number(3, 5) = &H1F

number(4, 2) = &H7
number(4, 3) = &H4
number(4, 4) = &H4
number(4, 5) = &H1F

number(5, 2) = &H17
number(5, 3) = &H15
number(5, 4) = &H15
number(5, 5) = &H1D

number(6, 2) = &H1F
number(6, 3) = &H14
number(6, 4) = &H14
number(6, 5) = &H1C

number(7, 2) = &H7
number(7, 3) = &H1
number(7, 4) = &H1
number(7, 5) = &H1F

number(8, 2) = &H1F
number(8, 3) = &H15
number(8, 4) = &H15
number(8, 5) = &H1F

number(9, 2) = &H17
number(9, 3) = &H15
number(9, 4) = &H15
number(9, 5) = &H1F

LedData(0) = &H80
LedData(1) = &H40
LedData(2) = &H20
LedData(3) = &H10
LedData(4) = &H8
LedData(5) = &H4
LedData(6) = &H2
LedData(7) = &H1
LedData(8) = &H2
LedData(9) = &H4
LedData(10) = &H8
LedData(11) = &H10
LedData(12) = &H20
LedData(13) = &H40
LedData(14) = &H80
LedData(15) = &HFF
LedData(16) = &H0
LedData(17) = &HFF
LedData(18) = &H0
LedData(19) = &HFF
LedData(20) = &H0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Check = OpenUsbDevice(VendorID, ProductID)                 ' 確認USB devie是否連線
    If (Check) Then
        For i = 0 To 7
            OutDataEightByte &H20, 0, i, 0, 0, 0, 0, 0
            OutDataEightByte &H20, 1, i, 0, 0, 0, 0, 0
            OutDataEightByte &H20, 2, i, 0, 0, 0, 0, 0
            OutDataEightByte &H20, 3, i, 0, 0, 0, 0, 0
        Next i
    End If
    CloseUsbDevice
End Sub

Private Sub Text5_Change()
If Text5.Text <> "" Then
    If Len(Text5.Text) > 4 Then
        Text5.Text = Mid(Text5.Text, 1, 4)
    End If
    For i = 1 To Len(Text5.Text)
        If Mid(Text5.Text, i, 1) < "0" Or Mid(Text5.Text, i, 1) > "9" Then
            Text5.Text = t1
        ElseIf i = 4 Then
        t1 = Text5.Text
     End If
    Next i
 End If
End Sub

Private Sub Timer1_Timer()
    isConnect
    Check = OpenUsbDevice(VendorID, ProductID)
    If (a) Then
        OutDataEightByte &H21, LedData(b), 0, 0, 0, 0, 0, 0
        b = b + 1
        If b = 21 Then a = 0
    End If
    ReadData ReadKeyData  'RXD
    CloseUsbDevice
    Select Case ReadKeyData(1)
        Case 0
           If Label1.Caption = " DEVICE ON " Then Label6.Caption = " KEY ON "
        Case 1
            Label6.Caption = " KEY OFF "
    End Select
End Sub
