VERSION 5.00
Begin VB.Form Q7 
   AutoRedraw      =   -1  'True
   Caption         =   "西式餐點-點餐服務系統"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   9825
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame1 
      Caption         =   "湯品"
      Height          =   1815
      Index           =   4
      Left            =   3360
      TabIndex        =   154
      Top             =   2280
      Width           =   3255
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   24
         Left            =   1320
         TabIndex        =   169
         Text            =   "100"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   24
         Left            =   2280
         TabIndex        =   168
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   24
         Left            =   2880
         Max             =   20
         TabIndex        =   167
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   25
         Left            =   2280
         TabIndex        =   166
         Text            =   "0"
         Top             =   735
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   25
         Left            =   2880
         Max             =   20
         TabIndex        =   165
         Top             =   735
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   26
         Left            =   2280
         TabIndex        =   164
         Text            =   "0"
         Top             =   975
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   26
         Left            =   2880
         Max             =   20
         TabIndex        =   163
         Top             =   975
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   27
         Left            =   2280
         TabIndex        =   162
         Text            =   "0"
         Top             =   1230
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   27
         Left            =   2880
         Max             =   20
         TabIndex        =   161
         Top             =   1230
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   28
         Left            =   2280
         TabIndex        =   160
         Text            =   "0"
         Top             =   1455
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   28
         Left            =   2880
         Max             =   20
         TabIndex        =   159
         Top             =   1455
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   25
         Left            =   1320
         TabIndex        =   158
         Text            =   "100"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   26
         Left            =   1320
         TabIndex        =   157
         Text            =   "100"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   27
         Left            =   1320
         TabIndex        =   156
         Text            =   "100"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   28
         Left            =   1320
         TabIndex        =   155
         Text            =   "100"
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "品名"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   177
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "單價"
         Height          =   255
         Index           =   13
         Left            =   1440
         TabIndex        =   176
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "數量"
         Height          =   255
         Index           =   14
         Left            =   2280
         TabIndex        =   175
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "雞蓉巧達湯"
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   174
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "海鮮嫩魚湯"
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   173
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "烤洋蔥湯"
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   172
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "南瓜湯"
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   171
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "脆皮濃湯"
         Height          =   255
         Index           =   28
         Left            =   120
         TabIndex        =   170
         Top             =   1455
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "熱飲"
      Height          =   1815
      Index           =   7
      Left            =   120
      TabIndex        =   130
      Top             =   4200
      Width           =   3975
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   42
         Left            =   1560
         TabIndex        =   145
         Text            =   "70"
         Top             =   465
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   42
         Left            =   3000
         TabIndex        =   144
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   42
         Left            =   3600
         Max             =   20
         TabIndex        =   143
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   43
         Left            =   1560
         TabIndex        =   142
         Text            =   "70"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   43
         Left            =   3000
         TabIndex        =   141
         Text            =   "0"
         Top             =   735
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   43
         Left            =   3600
         Max             =   20
         TabIndex        =   140
         Top             =   735
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   44
         Left            =   1560
         TabIndex        =   139
         Text            =   "90"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   44
         Left            =   3000
         TabIndex        =   138
         Text            =   "0"
         Top             =   975
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   44
         Left            =   3600
         Max             =   20
         TabIndex        =   137
         Top             =   975
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   45
         Left            =   1560
         TabIndex        =   136
         Text            =   "70"
         Top             =   1215
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   45
         Left            =   3000
         TabIndex        =   135
         Text            =   "0"
         Top             =   1230
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   45
         Left            =   3600
         Max             =   20
         TabIndex        =   134
         Top             =   1230
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   46
         Left            =   1560
         TabIndex        =   133
         Text            =   "100"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   46
         Left            =   3000
         TabIndex        =   132
         Text            =   "0"
         Top             =   1455
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   46
         Left            =   3600
         Max             =   20
         TabIndex        =   131
         Top             =   1455
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "品名"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   153
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "單價"
         Height          =   255
         Index           =   22
         Left            =   1680
         TabIndex        =   152
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "數量"
         Height          =   255
         Index           =   23
         Left            =   3000
         TabIndex        =   151
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "奶泡熱奶茶"
         Height          =   255
         Index           =   42
         Left            =   120
         TabIndex        =   150
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "熱咖啡"
         Height          =   255
         Index           =   43
         Left            =   120
         TabIndex        =   149
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "蜂蜜柚子茶"
         Height          =   255
         Index           =   44
         Left            =   120
         TabIndex        =   148
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "拿鐵熱咖啡"
         Height          =   255
         Index           =   45
         Left            =   120
         TabIndex        =   147
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "熱金桔檸檬梅子汁"
         Height          =   255
         Index           =   46
         Left            =   120
         TabIndex        =   146
         Top             =   1455
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "沙拉"
      Height          =   2175
      Index           =   1
      Left            =   5760
      TabIndex        =   98
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   11
         Left            =   1320
         TabIndex        =   125
         Text            =   "60"
         Top             =   1695
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   10
         Left            =   1320
         TabIndex        =   124
         Text            =   "60"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   9
         Left            =   1320
         TabIndex        =   123
         Text            =   "60"
         Top             =   1215
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   8
         Left            =   1320
         TabIndex        =   122
         Text            =   "60"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   7
         Left            =   1320
         TabIndex        =   121
         Text            =   "60"
         Top             =   720
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   11
         Left            =   3600
         Max             =   20
         TabIndex        =   111
         Top             =   1710
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   11
         Left            =   3000
         TabIndex        =   110
         Text            =   "0"
         Top             =   1710
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   10
         Left            =   3600
         Max             =   20
         TabIndex        =   109
         Top             =   1455
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   10
         Left            =   3000
         TabIndex        =   108
         Text            =   "0"
         Top             =   1455
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   9
         Left            =   3600
         Max             =   20
         TabIndex        =   107
         Top             =   1230
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   9
         Left            =   3000
         TabIndex        =   106
         Text            =   "0"
         Top             =   1230
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   8
         Left            =   3600
         Max             =   20
         TabIndex        =   105
         Top             =   975
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   8
         Left            =   3000
         TabIndex        =   104
         Text            =   "0"
         Top             =   975
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   7
         Left            =   3600
         Max             =   20
         TabIndex        =   103
         Top             =   735
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   7
         Left            =   3000
         TabIndex        =   102
         Text            =   "0"
         Top             =   735
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   6
         Left            =   3600
         Max             =   20
         TabIndex        =   101
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   6
         Left            =   3000
         TabIndex        =   100
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   6
         Left            =   1320
         TabIndex        =   99
         Text            =   "60"
         Top             =   465
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "義大利醬"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   120
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "千島醬"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   119
         Top             =   1455
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "優格水果沙拉"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   118
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "和風醬"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   117
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "凱薩醬"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   116
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "生菜沙拉"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   115
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "數量"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   114
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "單價"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   113
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "品名"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   112
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "主餐"
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   70
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   1320
         TabIndex        =   88
         Text            =   "250"
         Top             =   465
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   0
         Left            =   3000
         TabIndex        =   87
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   0
         Left            =   3600
         Max             =   20
         TabIndex        =   86
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   1320
         TabIndex        =   85
         Text            =   "380"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   1
         Left            =   3000
         TabIndex        =   84
         Text            =   "0"
         Top             =   735
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   1
         Left            =   3600
         Max             =   20
         TabIndex        =   83
         Top             =   735
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   2
         Left            =   1320
         TabIndex        =   82
         Text            =   "430"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   2
         Left            =   3000
         TabIndex        =   81
         Text            =   "0"
         Top             =   975
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   2
         Left            =   3600
         Max             =   20
         TabIndex        =   80
         Top             =   975
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   3
         Left            =   1320
         TabIndex        =   79
         Text            =   "450"
         Top             =   1215
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   3
         Left            =   3000
         TabIndex        =   78
         Text            =   "0"
         Top             =   1230
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   3
         Left            =   3600
         Max             =   20
         TabIndex        =   77
         Top             =   1230
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   4
         Left            =   1320
         TabIndex        =   76
         Text            =   "300"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   4
         Left            =   3000
         TabIndex        =   75
         Text            =   "0"
         Top             =   1455
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   4
         Left            =   3600
         Max             =   20
         TabIndex        =   74
         Top             =   1455
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   5
         Left            =   1320
         TabIndex        =   73
         Text            =   "570"
         Top             =   1695
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   5
         Left            =   3000
         TabIndex        =   72
         Text            =   "0"
         Top             =   1710
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   5
         Left            =   3600
         Max             =   20
         TabIndex        =   71
         Top             =   1710
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "品名"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "單價"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   96
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "數量"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   95
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "香酥脆皮雞排"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   94
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "特選沙朗牛排"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   93
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "特選菲力牛排"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   92
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "什錦海鮮"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   91
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "法式藍帶豬排"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   90
         Top             =   1455
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "海陸大餐"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   89
         Top             =   1710
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "冷飲"
      Height          =   1815
      Index           =   8
      Left            =   5760
      TabIndex        =   46
      Top             =   4200
      Width           =   3975
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   52
         Left            =   3600
         Max             =   20
         TabIndex        =   61
         Top             =   1455
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   52
         Left            =   3000
         TabIndex        =   60
         Text            =   "0"
         Top             =   1455
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   52
         Left            =   1320
         TabIndex        =   59
         Text            =   "70"
         Top             =   1440
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   51
         Left            =   3600
         Max             =   20
         TabIndex        =   58
         Top             =   1230
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   51
         Left            =   3000
         TabIndex        =   57
         Text            =   "0"
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   51
         Left            =   1320
         TabIndex        =   56
         Text            =   "90"
         Top             =   1215
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   50
         Left            =   3600
         Max             =   20
         TabIndex        =   55
         Top             =   975
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   50
         Left            =   3000
         TabIndex        =   54
         Text            =   "0"
         Top             =   975
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   50
         Left            =   1320
         TabIndex        =   53
         Text            =   "100"
         Top             =   960
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   49
         Left            =   3600
         Max             =   20
         TabIndex        =   52
         Top             =   735
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   49
         Left            =   3000
         TabIndex        =   51
         Text            =   "0"
         Top             =   735
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   49
         Left            =   1320
         TabIndex        =   50
         Text            =   "70"
         Top             =   720
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   48
         Left            =   3600
         Max             =   20
         TabIndex        =   49
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   48
         Left            =   3000
         TabIndex        =   48
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   48
         Left            =   1320
         TabIndex        =   47
         Text            =   "50"
         Top             =   465
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "冰拿鐵"
         Height          =   255
         Index           =   52
         Left            =   120
         TabIndex        =   69
         Top             =   1455
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "芒果汁"
         Height          =   255
         Index           =   51
         Left            =   120
         TabIndex        =   68
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "冰金桔檸檬"
         Height          =   255
         Index           =   50
         Left            =   120
         TabIndex        =   67
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "冰咖啡"
         Height          =   255
         Index           =   49
         Left            =   120
         TabIndex        =   66
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "可口可樂"
         Height          =   255
         Index           =   48
         Left            =   120
         TabIndex        =   65
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "數量"
         Height          =   255
         Index           =   26
         Left            =   3000
         TabIndex        =   64
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "單價"
         Height          =   255
         Index           =   25
         Left            =   1440
         TabIndex        =   63
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "品名"
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "甜點"
      Height          =   1815
      Index           =   5
      Left            =   6600
      TabIndex        =   22
      Top             =   2280
      Width           =   3135
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   34
         Left            =   2760
         Max             =   20
         TabIndex        =   37
         Top             =   1455
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   34
         Left            =   2160
         TabIndex        =   36
         Text            =   "0"
         Top             =   1455
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   34
         Left            =   1320
         TabIndex        =   35
         Text            =   "50"
         Top             =   1440
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   33
         Left            =   2760
         Max             =   20
         TabIndex        =   34
         Top             =   1230
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   33
         Left            =   2160
         TabIndex        =   33
         Text            =   "0"
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   33
         Left            =   1320
         TabIndex        =   32
         Text            =   "50"
         Top             =   1215
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   32
         Left            =   2760
         Max             =   20
         TabIndex        =   31
         Top             =   975
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   32
         Left            =   2160
         TabIndex        =   30
         Text            =   "0"
         Top             =   975
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   32
         Left            =   1320
         TabIndex        =   29
         Text            =   "40"
         Top             =   960
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   31
         Left            =   2760
         Max             =   20
         TabIndex        =   28
         Top             =   735
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   31
         Left            =   2160
         TabIndex        =   27
         Text            =   "0"
         Top             =   735
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   31
         Left            =   1320
         TabIndex        =   26
         Text            =   "50"
         Top             =   720
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   30
         Left            =   2760
         Max             =   20
         TabIndex        =   25
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   30
         Left            =   2160
         TabIndex        =   24
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   30
         Left            =   1320
         TabIndex        =   23
         Text            =   "30"
         Top             =   465
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "柳橙水果凍"
         Height          =   255
         Index           =   34
         Left            =   120
         TabIndex        =   45
         Top             =   1455
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "提拉米蘇"
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   44
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "義式布丁"
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   43
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "焦糖蛋糕"
         Height          =   255
         Index           =   31
         Left            =   120
         TabIndex        =   42
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "雞蛋布丁"
         Height          =   255
         Index           =   30
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "數量"
         Height          =   255
         Index           =   17
         Left            =   2160
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "單價"
         Height          =   255
         Index           =   16
         Left            =   1440
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "品名"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "前菜"
      Height          =   1815
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   3255
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   16
         Left            =   1320
         TabIndex        =   129
         Text            =   "80"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   15
         Left            =   1320
         TabIndex        =   128
         Text            =   "80"
         Top             =   1215
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   14
         Left            =   1320
         TabIndex        =   127
         Text            =   "80"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   13
         Left            =   1320
         TabIndex        =   126
         Text            =   "80"
         Top             =   720
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   16
         Left            =   2880
         Max             =   20
         TabIndex        =   13
         Top             =   1455
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   16
         Left            =   2280
         TabIndex        =   12
         Text            =   "0"
         Top             =   1455
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   15
         Left            =   2880
         Max             =   20
         TabIndex        =   11
         Top             =   1230
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   15
         Left            =   2280
         TabIndex        =   10
         Text            =   "0"
         Top             =   1230
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   14
         Left            =   2880
         Max             =   20
         TabIndex        =   9
         Top             =   975
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   14
         Left            =   2280
         TabIndex        =   8
         Text            =   "0"
         Top             =   975
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   13
         Left            =   2880
         Max             =   20
         TabIndex        =   7
         Top             =   735
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   13
         Left            =   2280
         TabIndex        =   6
         Text            =   "0"
         Top             =   735
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   12
         Left            =   2880
         Max             =   20
         TabIndex        =   5
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   12
         Left            =   2280
         TabIndex        =   4
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   12
         Left            =   1320
         TabIndex        =   3
         Text            =   "80"
         Top             =   465
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "黑菌鵝肝醬"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   21
         Top             =   1455
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "香蒜烤田螺"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   20
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "煙燻鮭魚"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   19
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "泰式嫩菲力"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   18
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "洋蔥鱈魚肝"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "數量"
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "單價"
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "品名"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "總金額"
      Height          =   495
      Left            =   3900
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "數量清除"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label lbltotal 
      Caption         =   "等待客人點餐中"
      Height          =   375
      Left            =   7680
      TabIndex        =   178
      Top             =   6120
      Width           =   1935
   End
End
Attribute VB_Name = "Q7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error Resume Next
Dim i%
For i = 0 To Text2.UBound
    Text2(i) = 0
Next i
lbltotal = "等待客人點餐中"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim total#, i%
total = 0
For i = 0 To Text2.UBound
    If CInt(Text2(i)) <> 0 Then
        total = total + CInt(Text1(i)) * CInt(Text2(i))
    End If
Next i
total = Int(total + total * 0.05 + 0.5)
lbltotal = CStr(total)
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i%
For i = 0 To VScroll1.UBound
    VScroll1(i).Value = 20
Next i
End Sub

Private Sub Text2_Change(Index As Integer)
If Text2(Index) = "" Then Text2(Index) = "0"
If CInt(Text2(Index)) > 20 Then Text2(Index) = "20"
If CInt(Text2(Index)) < 0 Then Text2(Index) = "0"
End Sub

Private Sub VScroll1_Change(Index As Integer)
Text2(Index) = CStr(20 - VScroll1(Index).Value)
End Sub
