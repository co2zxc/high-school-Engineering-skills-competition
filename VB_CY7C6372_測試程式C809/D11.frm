VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form USB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CY7C63723 I/O Test  Date : 2012-06-28"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12240
   Icon            =   "D11.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "INPUT TEST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   9120
      TabIndex        =   119
      Top             =   120
      Width           =   3015
      Begin VB.CheckBox Check2 
         Height          =   375
         Index           =   7
         Left            =   480
         TabIndex        =   130
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   129
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   375
         Index           =   5
         Left            =   960
         TabIndex        =   128
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   127
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   126
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   125
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   124
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   123
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1560
         TabIndex        =   122
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Read"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   120
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Bit : 8   7   6   5   4   3   2   1"
         Height          =   255
         Left            =   240
         TabIndex        =   121
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   9720
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LCD"
      Height          =   615
      Left            =   13080
      TabIndex        =   118
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      Caption         =   "LED & INPUT Test:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6120
      TabIndex        =   106
      Top             =   120
      Width           =   2895
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   115
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   114
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   5
         Left            =   960
         TabIndex        =   113
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   112
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   111
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   110
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   109
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   108
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   107
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label30 
         Caption         =   "Bit : 8   7   6   5   4   3   2   1"
         Height          =   255
         Left            =   240
         TabIndex        =   116
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "LCD 顯示器 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3120
      TabIndex        =   100
      Top             =   120
      Width           =   2895
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   360
         TabIndex        =   103
         Text            =   "5678"
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   360
         TabIndex        =   102
         Text            =   "1234"
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   101
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label29 
         Caption         =   "2. "
         Height          =   375
         Left            =   120
         TabIndex        =   105
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label28 
         Caption         =   "1. "
         Height          =   375
         Left            =   120
         TabIndex        =   104
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   99
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "七段顯示器 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   94
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command9 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   117
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1560
         TabIndex        =   96
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   95
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "U6 : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   98
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "U5 : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   97
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "USB Packet Data - IN/OUT Test :"
      Height          =   4095
      Left            =   6360
      TabIndex        =   51
      Top             =   5160
      Width           =   8535
      Begin VB.CommandButton Command4 
         Caption         =   "Connect"
         Height          =   495
         Left            =   5640
         TabIndex        =   93
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   3600
         TabIndex        =   91
         Text            =   "1234"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   840
         TabIndex        =   89
         Text            =   "7777"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Transfer"
         Height          =   615
         Left            =   120
         TabIndex        =   88
         Top             =   3240
         Width           =   3855
      End
      Begin VB.TextBox DataOUT 
         Height          =   495
         Index           =   7
         Left            =   7800
         TabIndex        =   87
         Text            =   "00"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox DataOUT 
         Height          =   495
         Index           =   6
         Left            =   7080
         TabIndex        =   86
         Text            =   "00"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox DataOUT 
         Height          =   495
         Index           =   5
         Left            =   6360
         TabIndex        =   85
         Text            =   "00"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox DataOUT 
         Height          =   495
         Index           =   4
         Left            =   5640
         TabIndex        =   84
         Text            =   "00"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox DataOUT 
         Height          =   495
         Index           =   3
         Left            =   4560
         TabIndex        =   83
         Text            =   "00"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox DataOUT 
         Height          =   495
         Index           =   2
         Left            =   3840
         TabIndex        =   82
         Text            =   "00"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox DataOUT 
         Height          =   495
         Index           =   1
         Left            =   3120
         TabIndex        =   81
         Text            =   "00"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox DataOUT 
         Height          =   495
         Index           =   0
         Left            =   2400
         TabIndex        =   80
         Text            =   "00"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox DataIN 
         Height          =   495
         Index           =   7
         Left            =   7800
         TabIndex        =   59
         Text            =   "00"
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox DataIN 
         Height          =   495
         Index           =   6
         Left            =   7080
         TabIndex        =   58
         Text            =   "00"
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox DataIN 
         Height          =   495
         Index           =   5
         Left            =   6360
         TabIndex        =   57
         Text            =   "00"
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox DataIN 
         Height          =   495
         Index           =   4
         Left            =   5640
         TabIndex        =   56
         Text            =   "00"
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox DataIN 
         Height          =   495
         Index           =   3
         Left            =   4560
         TabIndex        =   55
         Text            =   "00"
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox DataIN 
         Height          =   495
         Index           =   2
         Left            =   3840
         TabIndex        =   54
         Text            =   "00"
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox DataIN 
         Height          =   495
         Index           =   1
         Left            =   3120
         TabIndex        =   53
         Text            =   "00"
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox DataIN 
         Height          =   495
         Index           =   0
         Left            =   2400
         TabIndex        =   52
         Text            =   "00"
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "PID :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   92
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "VID :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   90
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   " __"
         Height          =   255
         Left            =   5280
         TabIndex        =   79
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "07"
         Height          =   255
         Left            =   7920
         TabIndex        =   78
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "06"
         Height          =   255
         Left            =   7200
         TabIndex        =   77
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "05"
         Height          =   255
         Left            =   6480
         TabIndex        =   76
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "04"
         Height          =   255
         Left            =   5760
         TabIndex        =   75
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "03"
         Height          =   255
         Left            =   4680
         TabIndex        =   74
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "02"
         Height          =   255
         Left            =   3960
         TabIndex        =   73
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "01"
         Height          =   255
         Left            =   3240
         TabIndex        =   72
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "IN(Device to PC) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   2640
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "00"
         Height          =   255
         Left            =   2520
         TabIndex        =   70
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   " __"
         Height          =   255
         Left            =   5280
         TabIndex        =   69
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "07"
         Height          =   255
         Left            =   7920
         TabIndex        =   68
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "06"
         Height          =   255
         Left            =   7200
         TabIndex        =   67
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "05"
         Height          =   255
         Left            =   6480
         TabIndex        =   66
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "04"
         Height          =   255
         Left            =   5760
         TabIndex        =   65
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "03"
         Height          =   255
         Left            =   4680
         TabIndex        =   64
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "02"
         Height          =   255
         Left            =   3960
         TabIndex        =   63
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "01"
         Height          =   255
         Left            =   3240
         TabIndex        =   62
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "00"
         Height          =   255
         Left            =   2520
         TabIndex        =   61
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "OUT(PC to Device) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   1800
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   2
      Left            =   240
      TabIndex        =   32
      Top             =   5520
      Width           =   5895
      Begin VB.Frame Frame2 
         Caption         =   "P0=IN    P2=OUT"
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   45
         Top             =   2520
         Width           =   5655
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   20.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   48
            Text            =   "00"
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "傳 送"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   4080
            TabIndex        =   47
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   20.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   720
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   46
            Text            =   "00"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "P0"
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   50
            Top             =   480
            Width           =   195
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "P2"
            Height          =   195
            Index           =   4
            Left            =   2160
            TabIndex        =   49
            Top             =   480
            Width           =   195
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "P0=OUT    P2=OUT"
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   5655
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   20.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   42
            Text            =   "00"
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "傳 送"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   4080
            TabIndex        =   41
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   20.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   720
            MaxLength       =   2
            TabIndex        =   40
            Text            =   "00"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "P0"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   44
            Top             =   480
            Width           =   195
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "P2"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   43
            Top             =   480
            Width           =   195
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "P0=IN    P2=IN"
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   5655
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   20.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   36
            Text            =   "00"
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "傳 送"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   4080
            TabIndex        =   35
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   20.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   720
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   34
            Text            =   "00"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "P2"
            Height          =   195
            Index           =   1
            Left            =   2160
            TabIndex        =   38
            Top             =   480
            Width           =   195
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "P0"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   37
            Top             =   480
            Width           =   195
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   3
      Left            =   240
      TabIndex        =   28
      Top             =   5520
      Width           =   5895
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "PDIUSBD11 GenIO"
         Height          =   195
         Index           =   5
         Left            =   2160
         TabIndex        =   31
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Version 3.0"
         Height          =   195
         Index           =   4
         Left            =   2400
         TabIndex        =   30
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "有任何問題歡迎到 http://www.usblab.idv.tw 發表問題"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   29
         Top             =   2880
         Width           =   4425
      End
   End
   Begin VB.Timer DelayTimer 
      Enabled         =   0   'False
      Left            =   840
      Top             =   9720
   End
   Begin VB.Timer ButtonTimer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   9600
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   5520
      Width           =   5895
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "按鍵輸入狀態"
         Height          =   195
         Index           =   2
         Left            =   2190
         TabIndex        =   26
         Top             =   960
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   3
         Left            =   4200
         Shape           =   2  'Oval
         Top             =   2160
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   2
         Left            =   3120
         Shape           =   2  'Oval
         Top             =   2160
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   1
         Left            =   2040
         Shape           =   2  'Oval
         Top             =   2160
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   0
         Left            =   960
         Shape           =   2  'Oval
         Top             =   2160
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   5520
      Width           =   5895
      Begin VB.CommandButton Command1 
         Caption         =   "傳送"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   480
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "00"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "00"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "00"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "00"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "00"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "00"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "00"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "00"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   11
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   10
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   9
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   8
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   7
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   4
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "已接收到 PDIUSBD11 的十六進制數值"
         Height          =   195
         Index           =   1
         Left            =   1290
         TabIndex        =   23
         Top             =   2400
         Width           =   3315
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   3
         Left            =   4920
         TabIndex        =   22
         Top             =   2640
         Width           =   105
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   2
         Left            =   4920
         TabIndex        =   21
         Top             =   960
         Width           =   105
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "7"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   20
         Top             =   2640
         Width           =   105
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "7"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   19
         Top             =   960
         Width           =   105
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "輸入要傳送到 PDIUSBD11 的十六進制數值"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   18
         Top             =   720
         Width           =   3735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2220
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   21537
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4095
      Left            =   120
      TabIndex        =   25
      Top             =   5160
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7223
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "一般 IO 控制"
            Key             =   ""
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "測試按鍵功能"
            Key             =   ""
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "測試 USB 傳輸"
            Key             =   ""
            Object.Tag             =   "3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "關於 D11"
            Key             =   ""
            Object.Tag             =   "4"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim D11 As New HID
Dim DealyFlag As Boolean
Dim SegNumber(0 To 9) As Byte
'Const VID = &H1234
'Const PID = &H5678
Const VID = &H7777
Const PID = &H1234
Dim Timeout As Boolean



Private Sub ButtonTimer1_Timer()
    Dim Send(0 To 7) As Byte
    Dim Rece() As String
    Dim temp As Integer
    
    CheckDevice
    If D11.OpenHIDDevice(VID, PID) = True Then
        Send(0) = 1
        D11.WriteHIDData Send()
        delay (50)
        Rece() = D11.ReadHIDData
        D11.CloseHIDDevice
        
        For temp = 0 To 3
            If (Rece(temp) = "FF") Then
                Shape1(temp).BorderColor = RGB(255, 0, 0)
                Shape1(temp).FillColor = RGB(255, 0, 0)
            Else
                Shape1(temp).BorderColor = RGB(0, 0, 0)
                Shape1(temp).FillColor = RGB(0, 0, 0)
            End If
        Next temp
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim Send(0 To 7) As Byte
    Dim Rece() As String
    Dim temp As Integer
    
    Select Case Index
        Case 1  'USB
            CheckDevice
            If D11.OpenHIDDevice(VID, PID) = True Then
                For temp = 4 To 7
                    Send(temp) = "&h" & Text1(temp - 4).Text
                Next temp
                Send(0) = 3
                D11.WriteHIDData Send()
                
                For temp = 4 To 7
                    Send(temp) = "&h" & Text1(temp).Text
                Next temp
                Send(0) = 4
                D11.WriteHIDData Send()
                delay (50)
                Rece() = D11.ReadHIDData
                D11.CloseHIDDevice
                
                For temp = 0 To 7
                    Text2(temp).Text = Rece(temp)
                Next temp
            End If
    End Select
End Sub

Private Sub Command10_Click()   ' READ
    Dim Send(0 To 7) As Byte
    Dim recevice() As String
    Dim i As Byte
    CheckDevice                                         ' 判斷 HID Device 是否存在
    If D11.OpenHIDDevice(VID, PID) = True Then
        For i = 0 To 7
            If Check2(i).Value = Checked Then           ' 取得 CheckBox 狀態
                Send(0) = Send(0) + 2 ^ i
            End If
        Next i
        Send(0) = &HFF - Send(0)                        ' 反相
        D11.WriteHIDData Send()                         ' 傳送至 HID Device
        
        
        Timeout = False
        Timer1.Enabled = True
        Timer1.Interval = 50
        Do
             DoEvents
        Loop Until Timeout = True


        recevice() = D11.ReadHIDData
        
        Do While (recevice(0) = "00")
        
            D11.WriteHIDData Send()                         ' 傳送至 HID Device
            recevice() = D11.ReadHIDData
        Loop
        
        
        'If (recevice(0) <> "00") Then
            Text10.Text = recevice(0)
        'End If
        
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
    Dim Send(0 To 7) As Byte
    Dim Rece() As String
    
    Select Case Index
        Case 0  'P0=IN  P2=IN
            CheckDevice
            If D11.OpenHIDDevice(VID, PID) = True Then
                Send(0) = 5
                D11.WriteHIDData Send()
                delay (50)
                Rece() = D11.ReadHIDData
                D11.CloseHIDDevice
                
                If Rece(0) = "01" Then
                    Text3(0).Text = CStr(Rece(1))
                    Text3(1).Text = CStr(Rece(2))
                End If
            End If
        Case 1  'P0=OUT P2=OUT
            CheckDevice
            If D11.OpenHIDDevice(VID, PID) = True Then
                Send(0) = 6
                
                If Text3(2).Text = "" Then
                Else
                    Send(1) = "&h" & Trim(Text3(2).Text)
                End If
                
                If Text3(3).Text = "" Then
                Else
                    Send(2) = "&h" & Trim(Text3(3).Text)
                End If
                
                D11.WriteHIDData Send()
                delay (50)
                Rece() = D11.ReadHIDData
                D11.CloseHIDDevice
            End If
        Case 2  'P0=IN  P2=OUT
            CheckDevice
            If D11.OpenHIDDevice(VID, PID) = True Then
                Send(0) = 7
                
                If Text3(5).Text = "" Then
                Else
                    Send(1) = "&h" & Trim(Text3(5).Text)
                End If
                
                D11.WriteHIDData Send()
                delay (50)
                Rece() = D11.ReadHIDData
                D11.CloseHIDDevice
                
                If Rece(0) = "02" Then
                    Text3(4).Text = CStr(Rece(1))
                End If
            End If
    End Select
End Sub

Private Sub Command3_Click()
    Dim Send(0 To 7) As Byte
    Dim Rece() As String
    Dim temp As Integer
    
    CheckDevice
    If D11.OpenHIDDevice(VID, PID) = True Then
        For temp = 0 To 7
            Send(temp) = "&h" & DataOUT(temp).Text
        Next temp
'        Send(0) = 1
        D11.WriteHIDData Send()
        
        'delay (50)
        
        'Rece() = D11.ReadHIDData
        'D11.CloseHIDDevice
        
   '     For temp = 0 To 7
   '         DataIN(temp).Text = Rece(temp)
   '     Next temp
    End If
End Sub
Private Sub Command4_Click()
CheckDevice
End Sub

'=============================================
'LCD 顯示器:
'=============================================
Private Sub Command7_Click()
    Dim Send(0 To 7) As Byte
    Dim i As Byte
    CheckDevice
    If D11.OpenHIDDevice(VID, PID) = True Then
        ' 07 06 05 04 03 02 01 00
        ' D7 D6 D5 D4 xx xx EN RS
        '------------------------------
        ' LCD Initial - 設定為 4bit mode
        '------------------------------
        delay (500)
        WriteIns (&H33)                             ' 送出兩次 0x03 至LCM模組的 D4~D7
        delay (500)
        WriteIns (&H32)                             ' 先送出 0x03後, 再送出 0x02 至LCM模組的 D4~D7
        WriteIns (&H1)                              ' 清除顯示資料
        WriteIns (&HE)                              ' 顯示功能ON, 游標顯示
        
        delay (500)
        WriteIns (&H80)                             ' 設定LCM啟始位置-第一列
        For i = 0 To Len(Text8.Text) - 1
            WriteData (Asc(Mid(Text8.Text, i + 1, 1)))  ' 寫入資料 至 LCM 模組
        Next i
        
        delay (50)
        
        WriteIns (&HC0)                             ' 設定LCM啟始位置-第二列
        For i = 0 To Len(Text9.Text) - 1
            WriteData (Asc(Mid(Text9.Text, i + 1, 1)))  ' 寫入資料 至 LCM 模組
        Next i
    End If
End Sub
'---------------------------------------------
' LCD Pin 定義 :
'---------------------------------------------
' Bit : 07 06 05 04 03 02 01 00
'       D7 D6 D5 D4 xx xx EN RS
'=============================================
' LCD 顯示器 : 指令
'=============================================
Public Sub WriteIns(data As Byte)
    Dim tmp As Byte

    tmp = &HF0 And ((data And &HF0))
    SendPacket (0)
    SendPacket (tmp + 2)                                ' RS = 0, EN = 1
    SendPacket (tmp)                                    ' RS = 0, EN = 0
    
    tmp = &HF0 And ((data And &HF) * 16)
    SendPacket (0)
    SendPacket (tmp + 2)                                ' RS = 0, EN = 1
    SendPacket (tmp)                                    ' RS = 0, EN = 0
End Sub
'=============================================
' LCD 顯示器 : 資料
'=============================================
Public Sub WriteData(data As Byte)
    Dim tmp As Byte

    tmp = &HF0 And ((data And &HF0))
    SendPacket (1)
    SendPacket (tmp + 3)                                ' RS = 1, EN = 1
    SendPacket (tmp + 1)                                ' RS = 1, EN = 0
    
    tmp = &HF0 And ((data And &HF) * 16)
    SendPacket (1)
    SendPacket (tmp + 3)                                ' RS = 1, EN = 1
    SendPacket (tmp + 1)                                ' RS = 1, EN = 0
End Sub

Private Sub SendPacket(data As Byte)
    Dim Send(0 To 7) As Byte
    Send(0) = data
    D11.WriteHIDData Send()                             ' 傳送至 HID Device
    delay (20)
End Sub

'=============================================
' 七段顯示器-U5 :
'=============================================
Private Sub Command6_Click()
    Dim Send(0 To 7) As Byte
    Dim i As Byte
    CheckDevice                                         ' 判斷 HID Device 是否存在
    If D11.OpenHIDDevice(VID, PID) = True Then
        Send(0) = SegNumber(Val(Text6.Text))            ' 取得顯示的數字代碼
        Send(1) = 1                                     ' 選擇第一顆 七段顯示器
        D11.WriteHIDData Send()                         ' 傳送至 HID Device
    End If
End Sub

'=============================================
' 七段顯示器-U6 :
'=============================================
Private Sub Command9_Click()
    Dim Send(0 To 7) As Byte
    Dim i As Byte
    CheckDevice                                         ' 判斷 HID Device 是否存在
    If D11.OpenHIDDevice(VID, PID) = True Then
        
        Send(0) = SegNumber(Val(Text7.Text))            ' 取得顯示的數字代碼
        Send(1) = 2                                     ' 選擇第二顆 七段顯示器
        D11.WriteHIDData Send()                         ' 傳送至 HID Device
    End If
End Sub


'=============================================
' LED TEST
'=============================================
Private Sub Command8_Click()
    Dim Send(0 To 7) As Byte
    Dim i As Byte
    CheckDevice                                         ' 判斷 HID Device 是否存在
    If D11.OpenHIDDevice(VID, PID) = True Then
        For i = 0 To 7
            If Check1(i).Value = Checked Then           ' 取得 CheckBox 狀態
                Send(0) = Send(0) + 2 ^ i
            End If
        Next i
        Send(0) = &HFF - Send(0)                        ' 反相
        D11.WriteHIDData Send()                         ' 傳送至 HID Device
    End If
End Sub

Private Sub DelayTimer_Timer()
    DelayTimer.Enabled = False
    DelayTimer.Interval = 0
    DealyFlag = True
End Sub

Private Sub Form_Load()
    Dim temp As Integer

    For temp = 0 To 3
        Frame1(temp).Visible = False
    Next temp
    Frame1(2).Visible = True
    
    CheckDevice                                         ' 判斷 HID Device 是否存在
    
    SegNumber(0) = &HFC
    SegNumber(1) = &H60
    SegNumber(2) = &HDA
    SegNumber(3) = &HF2
    SegNumber(4) = &H66
    SegNumber(5) = &HB6
    SegNumber(6) = &H3E
    SegNumber(7) = &HE4
    SegNumber(8) = &HFE
    SegNumber(9) = &HE6
    
End Sub

Private Sub TabStrip1_Click()
    Dim temp As Integer
    
    For temp = 0 To 3
        Frame1(temp).Visible = False
    Next temp
    ButtonTimer1.Enabled = False
    
    Select Case TabStrip1.SelectedItem
        Case "測試 USB 傳輸"
            Frame1(0).Visible = True
            
        Case "測試按鍵功能"
            Frame1(1).Visible = True
            ButtonTimer1.Enabled = True
            ButtonTimer1.Interval = 100

        Case "一般 IO 控制"
            Frame1(2).Visible = True
            
        Case "關於 D11"
            Frame1(3).Visible = True
            
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = CheckKey(KeyAscii)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = CheckKey(KeyAscii)
End Sub

Public Function CheckKey(iValue As Integer) As Integer
    If (Chr$(iValue) Like "[!A-Fa-f0-9]") Then
        CheckKey = 0
    Else
        If (Chr$(iValue) Like "[a-f]") Then
            CheckKey = iValue - 32
        Else
            CheckKey = iValue
        End If
    End If
End Function

Public Function CheckDevice() As Boolean
    Dim ChkStatus As Boolean
    
    ChkStatus = D11.OpenHIDDevice(VID, PID)
    If ChkStatus = True Then
        StatusBar1.Panels(1).Text = "已找到 HID 裝置"
    Else
        StatusBar1.Panels(1).Text = "請插入 HID 裝置"
    End If
    D11.CloseHIDDevice
    
    CheckDevice = ChkStatus
End Function

Public Sub delay(iValue As Integer)
    DealyFlag = False
    DelayTimer.Enabled = True
    DelayTimer.Interval = iValue
    
    Do
        DoEvents
    Loop Until DealyFlag = True
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 2
            KeyAscii = CheckKey(KeyAscii)
        Case 3
            KeyAscii = CheckKey(KeyAscii)
        Case 4
            KeyAscii = CheckKey(KeyAscii)
    End Select
End Sub

Private Sub Timer1_Timer()
Timeout = True
Timer1.Enabled = False
End Sub
