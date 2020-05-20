VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form3 
   BackColor       =   &H00C0FFFF&
   Caption         =   "銷售紀錄"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  '最大化
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3120
      TabIndex        =   827
      Text            =   "Text2"
      Top             =   240
      Width           =   2895
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Form3.frx":2007
      Height          =   435
      Left            =   4560
      TabIndex        =   825
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   393216
      ListField       =   "客戶編號"
      BoundColumn     =   "客戶編號"
      Text            =   "請選擇您的客戶編號- - - "
      Object.DataMember      =   "客戶"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame9 
      Caption         =   "總金額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1455
      Left            =   11280
      TabIndex        =   7
      Top             =   7920
      Width           =   2655
      Begin VB.Label Label4 
         Alignment       =   2  '置中對齊
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   27.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   855
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   13361
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   12648447
      ForeColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "蛋糕"
      TabPicture(0)   =   "Form3.frx":2026
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "飲料"
      TabPicture(1)   =   "Form3.frx":2042
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(2)=   "Frame8"
      Tab(1).Control(3)=   "Frame10"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "麵包"
      TabPicture(2)   =   "Form3.frx":205E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame15"
      Tab(2).Control(1)=   "Frame14"
      Tab(2).Control(2)=   "Frame13"
      Tab(2).Control(3)=   "Frame12"
      Tab(2).Control(4)=   "Frame11"
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame15 
         Caption         =   "丹麥類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   -70320
         TabIndex        =   760
         Top             =   4440
         Width           =   3975
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   114
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   819
            Top             =   3240
            Value           =   1
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   114
            Left            =   840
            TabIndex        =   818
            Top             =   3240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   114
            Left            =   1800
            TabIndex        =   817
            Text            =   "1"
            Top             =   3240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   113
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   812
            Top             =   2880
            Value           =   1
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   113
            Left            =   840
            TabIndex        =   811
            Top             =   2880
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   113
            Left            =   1800
            TabIndex        =   810
            Text            =   "1"
            Top             =   2880
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   112
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   805
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   112
            Left            =   840
            TabIndex        =   804
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   112
            Left            =   1800
            TabIndex        =   803
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   111
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   798
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   111
            Left            =   840
            TabIndex        =   797
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   111
            Left            =   1800
            TabIndex        =   796
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   110
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   791
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   110
            Left            =   840
            TabIndex        =   790
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   110
            Left            =   1800
            TabIndex        =   789
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   109
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   784
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   109
            Left            =   840
            TabIndex        =   783
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   109
            Left            =   1800
            TabIndex        =   782
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   108
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   777
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   108
            Left            =   840
            TabIndex        =   776
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   108
            Left            =   1800
            TabIndex        =   775
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   107
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   770
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   107
            Left            =   840
            TabIndex        =   769
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   107
            Left            =   1800
            TabIndex        =   768
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   106
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   763
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   106
            Left            =   840
            TabIndex        =   762
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   106
            Left            =   1800
            TabIndex        =   761
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   114
            Left            =   480
            TabIndex        =   823
            Top             =   3240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Index           =   114
            Left            =   2520
            TabIndex        =   822
            Top             =   3240
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   115
            Left            =   1200
            TabIndex        =   821
            Top             =   3360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   114
            Left            =   120
            TabIndex        =   820
            Top             =   3240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   113
            Left            =   480
            TabIndex        =   816
            Top             =   2880
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Index           =   113
            Left            =   2520
            TabIndex        =   815
            Top             =   2880
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   114
            Left            =   1200
            TabIndex        =   814
            Top             =   3000
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   113
            Left            =   120
            TabIndex        =   813
            Top             =   2880
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   112
            Left            =   480
            TabIndex        =   809
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   112
            Left            =   2520
            TabIndex        =   808
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   113
            Left            =   1200
            TabIndex        =   807
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   112
            Left            =   120
            TabIndex        =   806
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   111
            Left            =   480
            TabIndex        =   802
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   111
            Left            =   2520
            TabIndex        =   801
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   112
            Left            =   1200
            TabIndex        =   800
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   111
            Left            =   120
            TabIndex        =   799
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   110
            Left            =   480
            TabIndex        =   795
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   110
            Left            =   2520
            TabIndex        =   794
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   111
            Left            =   1200
            TabIndex        =   793
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   110
            Left            =   120
            TabIndex        =   792
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   109
            Left            =   480
            TabIndex        =   788
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   109
            Left            =   2520
            TabIndex        =   787
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   110
            Left            =   1200
            TabIndex        =   786
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   109
            Left            =   120
            TabIndex        =   785
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   108
            Left            =   480
            TabIndex        =   781
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   108
            Left            =   2520
            TabIndex        =   780
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   109
            Left            =   1200
            TabIndex        =   779
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   108
            Left            =   120
            TabIndex        =   778
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   107
            Left            =   480
            TabIndex        =   774
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   107
            Left            =   2520
            TabIndex        =   773
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   108
            Left            =   1200
            TabIndex        =   772
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   107
            Left            =   120
            TabIndex        =   771
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   106
            Left            =   480
            TabIndex        =   767
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   106
            Left            =   2520
            TabIndex        =   766
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   107
            Left            =   1200
            TabIndex        =   765
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   106
            Left            =   120
            TabIndex        =   764
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "本土化類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -70440
         TabIndex        =   689
         Top             =   360
         Width           =   4215
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   105
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   755
            Top             =   3480
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   105
            Left            =   840
            TabIndex        =   754
            Top             =   3480
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   105
            Left            =   1800
            TabIndex        =   753
            Text            =   "1"
            Top             =   3480
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   104
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   748
            Top             =   3120
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   104
            Left            =   840
            TabIndex        =   747
            Top             =   3120
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   104
            Left            =   1800
            TabIndex        =   746
            Text            =   "1"
            Top             =   3120
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   103
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   741
            Top             =   2760
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   103
            Left            =   840
            TabIndex        =   740
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   103
            Left            =   1800
            TabIndex        =   739
            Text            =   "1"
            Top             =   2760
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   102
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   734
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   102
            Left            =   840
            TabIndex        =   733
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   102
            Left            =   1800
            TabIndex        =   732
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   101
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   727
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   101
            Left            =   840
            TabIndex        =   726
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   101
            Left            =   1800
            TabIndex        =   725
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   100
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   720
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   100
            Left            =   840
            TabIndex        =   719
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   100
            Left            =   1800
            TabIndex        =   718
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   99
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   713
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   99
            Left            =   840
            TabIndex        =   712
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   99
            Left            =   1800
            TabIndex        =   711
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   98
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   706
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   98
            Left            =   840
            TabIndex        =   705
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   98
            Left            =   1800
            TabIndex        =   704
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   97
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   699
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   97
            Left            =   840
            TabIndex        =   698
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   97
            Left            =   1800
            TabIndex        =   697
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   96
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   692
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   96
            Left            =   840
            TabIndex        =   691
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   96
            Left            =   1800
            TabIndex        =   690
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   105
            Left            =   480
            TabIndex        =   759
            Top             =   3480
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   105
            Left            =   2520
            TabIndex        =   758
            Top             =   3480
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   106
            Left            =   1200
            TabIndex        =   757
            Top             =   3600
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   105
            Left            =   120
            TabIndex        =   756
            Top             =   3480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   104
            Left            =   480
            TabIndex        =   752
            Top             =   3120
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   104
            Left            =   2520
            TabIndex        =   751
            Top             =   3120
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   105
            Left            =   1200
            TabIndex        =   750
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   104
            Left            =   120
            TabIndex        =   749
            Top             =   3120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   103
            Left            =   480
            TabIndex        =   745
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   103
            Left            =   2520
            TabIndex        =   744
            Top             =   2760
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   104
            Left            =   1200
            TabIndex        =   743
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   103
            Left            =   120
            TabIndex        =   742
            Top             =   2760
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   102
            Left            =   480
            TabIndex        =   738
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   102
            Left            =   2520
            TabIndex        =   737
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   103
            Left            =   1200
            TabIndex        =   736
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   102
            Left            =   120
            TabIndex        =   735
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   101
            Left            =   480
            TabIndex        =   731
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   101
            Left            =   2520
            TabIndex        =   730
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   102
            Left            =   1200
            TabIndex        =   729
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   101
            Left            =   120
            TabIndex        =   728
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   100
            Left            =   480
            TabIndex        =   724
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   100
            Left            =   2520
            TabIndex        =   723
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   101
            Left            =   1200
            TabIndex        =   722
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   100
            Left            =   120
            TabIndex        =   721
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   99
            Left            =   480
            TabIndex        =   717
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   99
            Left            =   2520
            TabIndex        =   716
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   100
            Left            =   1200
            TabIndex        =   715
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   99
            Left            =   120
            TabIndex        =   714
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   98
            Left            =   480
            TabIndex        =   710
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   98
            Left            =   2520
            TabIndex        =   709
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   99
            Left            =   1200
            TabIndex        =   708
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   98
            Left            =   120
            TabIndex        =   707
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   97
            Left            =   480
            TabIndex        =   703
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   97
            Left            =   2520
            TabIndex        =   702
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   98
            Left            =   1200
            TabIndex        =   701
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   97
            Left            =   120
            TabIndex        =   700
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   96
            Left            =   480
            TabIndex        =   696
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   96
            Left            =   2520
            TabIndex        =   695
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   97
            Left            =   1200
            TabIndex        =   694
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   96
            Left            =   120
            TabIndex        =   693
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "日本類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   -65880
         TabIndex        =   604
         Top             =   480
         Width           =   4215
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   95
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   684
            Top             =   4200
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   95
            Left            =   840
            TabIndex        =   683
            Top             =   4200
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   95
            Left            =   1800
            TabIndex        =   682
            Text            =   "1"
            Top             =   4200
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   94
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   677
            Top             =   3840
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   94
            Left            =   840
            TabIndex        =   676
            Top             =   3840
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   94
            Left            =   1800
            TabIndex        =   675
            Text            =   "1"
            Top             =   3840
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   93
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   670
            Top             =   3480
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   93
            Left            =   840
            TabIndex        =   669
            Top             =   3480
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   93
            Left            =   1800
            TabIndex        =   668
            Text            =   "1"
            Top             =   3480
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   92
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   663
            Top             =   3120
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   92
            Left            =   840
            TabIndex        =   662
            Top             =   3120
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   92
            Left            =   1800
            TabIndex        =   661
            Text            =   "1"
            Top             =   3120
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   91
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   656
            Top             =   2760
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   91
            Left            =   840
            TabIndex        =   655
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   91
            Left            =   1800
            TabIndex        =   654
            Text            =   "1"
            Top             =   2760
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   90
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   649
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   90
            Left            =   840
            TabIndex        =   648
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   90
            Left            =   1800
            TabIndex        =   647
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   89
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   642
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   89
            Left            =   840
            TabIndex        =   641
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   89
            Left            =   1800
            TabIndex        =   640
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   88
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   635
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   88
            Left            =   840
            TabIndex        =   634
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   88
            Left            =   1800
            TabIndex        =   633
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   87
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   628
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   87
            Left            =   840
            TabIndex        =   627
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   87
            Left            =   1800
            TabIndex        =   626
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   86
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   621
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   86
            Left            =   840
            TabIndex        =   620
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   86
            Left            =   1800
            TabIndex        =   619
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   85
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   614
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   85
            Left            =   840
            TabIndex        =   613
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   85
            Left            =   1800
            TabIndex        =   612
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   84
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   607
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   84
            Left            =   840
            TabIndex        =   606
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   84
            Left            =   1800
            TabIndex        =   605
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   95
            Left            =   480
            TabIndex        =   688
            Top             =   4200
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   95
            Left            =   2520
            TabIndex        =   687
            Top             =   4200
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   96
            Left            =   1200
            TabIndex        =   686
            Top             =   4320
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   95
            Left            =   120
            TabIndex        =   685
            Top             =   4200
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   94
            Left            =   480
            TabIndex        =   681
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   94
            Left            =   2520
            TabIndex        =   680
            Top             =   3840
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   95
            Left            =   1200
            TabIndex        =   679
            Top             =   3960
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   94
            Left            =   120
            TabIndex        =   678
            Top             =   3840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   93
            Left            =   480
            TabIndex        =   674
            Top             =   3480
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   93
            Left            =   2520
            TabIndex        =   673
            Top             =   3480
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   94
            Left            =   1200
            TabIndex        =   672
            Top             =   3600
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   93
            Left            =   120
            TabIndex        =   671
            Top             =   3480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   92
            Left            =   480
            TabIndex        =   667
            Top             =   3120
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   92
            Left            =   2520
            TabIndex        =   666
            Top             =   3120
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   93
            Left            =   1200
            TabIndex        =   665
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   92
            Left            =   120
            TabIndex        =   664
            Top             =   3120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   91
            Left            =   480
            TabIndex        =   660
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   91
            Left            =   2520
            TabIndex        =   659
            Top             =   2760
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   92
            Left            =   1200
            TabIndex        =   658
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   91
            Left            =   120
            TabIndex        =   657
            Top             =   2760
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   90
            Left            =   480
            TabIndex        =   653
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   90
            Left            =   2520
            TabIndex        =   652
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   91
            Left            =   1200
            TabIndex        =   651
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   90
            Left            =   120
            TabIndex        =   650
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   89
            Left            =   480
            TabIndex        =   646
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   89
            Left            =   2520
            TabIndex        =   645
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   90
            Left            =   1200
            TabIndex        =   644
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   89
            Left            =   120
            TabIndex        =   643
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   88
            Left            =   480
            TabIndex        =   639
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   88
            Left            =   2520
            TabIndex        =   638
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   89
            Left            =   1200
            TabIndex        =   637
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   88
            Left            =   120
            TabIndex        =   636
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   87
            Left            =   480
            TabIndex        =   632
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   87
            Left            =   2520
            TabIndex        =   631
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   88
            Left            =   1200
            TabIndex        =   630
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   87
            Left            =   120
            TabIndex        =   629
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   86
            Left            =   480
            TabIndex        =   625
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   86
            Left            =   2520
            TabIndex        =   624
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   87
            Left            =   1200
            TabIndex        =   623
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   86
            Left            =   120
            TabIndex        =   622
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   85
            Left            =   480
            TabIndex        =   618
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   85
            Left            =   2520
            TabIndex        =   617
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   86
            Left            =   1200
            TabIndex        =   616
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   85
            Left            =   120
            TabIndex        =   615
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   84
            Left            =   480
            TabIndex        =   611
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   84
            Left            =   2520
            TabIndex        =   610
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   85
            Left            =   1200
            TabIndex        =   609
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   84
            Left            =   120
            TabIndex        =   608
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "歐式類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -74760
         TabIndex        =   554
         Top             =   3960
         Width           =   3975
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   83
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   599
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   83
            Left            =   840
            TabIndex        =   598
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   83
            Left            =   1800
            TabIndex        =   597
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   82
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   592
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   82
            Left            =   840
            TabIndex        =   591
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   82
            Left            =   1800
            TabIndex        =   590
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   81
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   585
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   81
            Left            =   840
            TabIndex        =   584
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   81
            Left            =   1800
            TabIndex        =   583
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   80
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   578
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   80
            Left            =   840
            TabIndex        =   577
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   80
            Left            =   1800
            TabIndex        =   576
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   79
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   571
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   79
            Left            =   840
            TabIndex        =   570
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   79
            Left            =   1800
            TabIndex        =   569
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   78
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   564
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   78
            Left            =   840
            TabIndex        =   563
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   78
            Left            =   1800
            TabIndex        =   562
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   77
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   557
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   77
            Left            =   840
            TabIndex        =   556
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   77
            Left            =   1800
            TabIndex        =   555
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   83
            Left            =   480
            TabIndex        =   603
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   83
            Left            =   2520
            TabIndex        =   602
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   84
            Left            =   1200
            TabIndex        =   601
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   83
            Left            =   120
            TabIndex        =   600
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   82
            Left            =   480
            TabIndex        =   596
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   82
            Left            =   2520
            TabIndex        =   595
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   83
            Left            =   1200
            TabIndex        =   594
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   82
            Left            =   120
            TabIndex        =   593
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   81
            Left            =   480
            TabIndex        =   589
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   81
            Left            =   2520
            TabIndex        =   588
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   82
            Left            =   1200
            TabIndex        =   587
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   81
            Left            =   120
            TabIndex        =   586
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   80
            Left            =   480
            TabIndex        =   582
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   80
            Left            =   2520
            TabIndex        =   581
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   81
            Left            =   1200
            TabIndex        =   580
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   80
            Left            =   120
            TabIndex        =   579
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   79
            Left            =   480
            TabIndex        =   575
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   79
            Left            =   2520
            TabIndex        =   574
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   80
            Left            =   1200
            TabIndex        =   573
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   79
            Left            =   120
            TabIndex        =   572
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   78
            Left            =   480
            TabIndex        =   568
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   78
            Left            =   2520
            TabIndex        =   567
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   79
            Left            =   1200
            TabIndex        =   566
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   78
            Left            =   120
            TabIndex        =   565
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   77
            Left            =   480
            TabIndex        =   561
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   77
            Left            =   2520
            TabIndex        =   560
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   78
            Left            =   1200
            TabIndex        =   559
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   77
            Left            =   120
            TabIndex        =   558
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "吐司類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   -74760
         TabIndex        =   497
         Top             =   480
         Width           =   4095
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   76
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   549
            Top             =   2760
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   76
            Left            =   840
            TabIndex        =   548
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   76
            Left            =   1800
            TabIndex        =   547
            Text            =   "1"
            Top             =   2760
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   75
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   542
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   75
            Left            =   840
            TabIndex        =   541
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   75
            Left            =   1800
            TabIndex        =   540
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   74
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   535
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   74
            Left            =   840
            TabIndex        =   534
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   74
            Left            =   1800
            TabIndex        =   533
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   73
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   528
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   73
            Left            =   840
            TabIndex        =   527
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   73
            Left            =   1800
            TabIndex        =   526
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   72
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   521
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   72
            Left            =   840
            TabIndex        =   520
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   72
            Left            =   1800
            TabIndex        =   519
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   71
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   514
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   71
            Left            =   840
            TabIndex        =   513
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   71
            Left            =   1800
            TabIndex        =   512
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   70
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   507
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   70
            Left            =   840
            TabIndex        =   506
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   70
            Left            =   1800
            TabIndex        =   505
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   69
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   500
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   69
            Left            =   840
            TabIndex        =   499
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   69
            Left            =   1800
            TabIndex        =   498
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   76
            Left            =   480
            TabIndex        =   553
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   76
            Left            =   2520
            TabIndex        =   552
            Top             =   2760
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   77
            Left            =   1200
            TabIndex        =   551
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   76
            Left            =   120
            TabIndex        =   550
            Top             =   2760
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   75
            Left            =   480
            TabIndex        =   546
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   75
            Left            =   2520
            TabIndex        =   545
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   76
            Left            =   1200
            TabIndex        =   544
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   75
            Left            =   120
            TabIndex        =   543
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   74
            Left            =   480
            TabIndex        =   539
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   74
            Left            =   2520
            TabIndex        =   538
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   75
            Left            =   1200
            TabIndex        =   537
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   74
            Left            =   120
            TabIndex        =   536
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   73
            Left            =   480
            TabIndex        =   532
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   73
            Left            =   2520
            TabIndex        =   531
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   74
            Left            =   1200
            TabIndex        =   530
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   73
            Left            =   120
            TabIndex        =   529
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   72
            Left            =   480
            TabIndex        =   525
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   72
            Left            =   2520
            TabIndex        =   524
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   73
            Left            =   1200
            TabIndex        =   523
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   72
            Left            =   120
            TabIndex        =   522
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   71
            Left            =   480
            TabIndex        =   518
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   71
            Left            =   2520
            TabIndex        =   517
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   72
            Left            =   1200
            TabIndex        =   516
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   71
            Left            =   120
            TabIndex        =   515
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   70
            Left            =   480
            TabIndex        =   511
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   70
            Left            =   2520
            TabIndex        =   510
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   71
            Left            =   1200
            TabIndex        =   509
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   70
            Left            =   120
            TabIndex        =   508
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   69
            Left            =   480
            TabIndex        =   504
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   69
            Left            =   2520
            TabIndex        =   503
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   70
            Left            =   1200
            TabIndex        =   502
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   69
            Left            =   120
            TabIndex        =   501
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "奶茶"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74520
         TabIndex        =   461
         Top             =   4440
         Width           =   3615
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   68
            Left            =   1800
            TabIndex        =   492
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   68
            Left            =   840
            TabIndex        =   491
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   68
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   490
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   67
            Left            =   1800
            TabIndex        =   485
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   67
            Left            =   840
            TabIndex        =   484
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   67
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   483
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   66
            Left            =   1800
            TabIndex        =   478
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   66
            Left            =   840
            TabIndex        =   477
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   66
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   476
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   65
            Left            =   1800
            TabIndex        =   471
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   65
            Left            =   840
            TabIndex        =   470
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   65
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   469
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   64
            Left            =   1800
            TabIndex        =   464
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   64
            Left            =   840
            TabIndex        =   463
            Top             =   240
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   64
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   462
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   68
            Left            =   120
            TabIndex        =   496
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   69
            Left            =   1200
            TabIndex        =   495
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   68
            Left            =   2520
            TabIndex        =   494
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   68
            Left            =   480
            TabIndex        =   493
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   67
            Left            =   120
            TabIndex        =   489
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   68
            Left            =   1200
            TabIndex        =   488
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   67
            Left            =   2520
            TabIndex        =   487
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   67
            Left            =   480
            TabIndex        =   486
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   66
            Left            =   120
            TabIndex        =   482
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   67
            Left            =   1200
            TabIndex        =   481
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   66
            Left            =   2520
            TabIndex        =   480
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   66
            Left            =   480
            TabIndex        =   479
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   65
            Left            =   120
            TabIndex        =   475
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   66
            Left            =   1200
            TabIndex        =   474
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   65
            Left            =   2520
            TabIndex        =   473
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   65
            Left            =   480
            TabIndex        =   472
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   64
            Left            =   120
            TabIndex        =   468
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   65
            Left            =   1200
            TabIndex        =   467
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   64
            Left            =   2520
            TabIndex        =   466
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   64
            Left            =   480
            TabIndex        =   465
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "冰沙"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -65760
         TabIndex        =   418
         Top             =   720
         Width           =   3975
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   63
            Left            =   1800
            TabIndex        =   456
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   63
            Left            =   840
            TabIndex        =   455
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   63
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   454
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   62
            Left            =   1800
            TabIndex        =   449
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   62
            Left            =   840
            TabIndex        =   448
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   62
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   447
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   61
            Left            =   1800
            TabIndex        =   442
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   61
            Left            =   840
            TabIndex        =   441
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   61
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   440
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   60
            Left            =   1800
            TabIndex        =   435
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   60
            Left            =   840
            TabIndex        =   434
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   60
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   433
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   59
            Left            =   1800
            TabIndex        =   428
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   59
            Left            =   840
            TabIndex        =   427
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   59
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   426
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   58
            Left            =   1800
            TabIndex        =   421
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   58
            Left            =   840
            TabIndex        =   420
            Top             =   240
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   58
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   419
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   63
            Left            =   120
            TabIndex        =   460
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   64
            Left            =   1200
            TabIndex        =   459
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   63
            Left            =   2520
            TabIndex        =   458
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   63
            Left            =   480
            TabIndex        =   457
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   62
            Left            =   120
            TabIndex        =   453
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   63
            Left            =   1200
            TabIndex        =   452
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   62
            Left            =   2520
            TabIndex        =   451
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   62
            Left            =   480
            TabIndex        =   450
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   61
            Left            =   120
            TabIndex        =   446
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   62
            Left            =   1200
            TabIndex        =   445
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   61
            Left            =   2520
            TabIndex        =   444
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   61
            Left            =   480
            TabIndex        =   443
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   60
            Left            =   120
            TabIndex        =   439
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   61
            Left            =   1200
            TabIndex        =   438
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   60
            Left            =   2520
            TabIndex        =   437
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   60
            Left            =   480
            TabIndex        =   436
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   59
            Left            =   120
            TabIndex        =   432
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   60
            Left            =   1200
            TabIndex        =   431
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   59
            Left            =   2520
            TabIndex        =   430
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   59
            Left            =   480
            TabIndex        =   429
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   58
            Left            =   120
            TabIndex        =   425
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   59
            Left            =   1200
            TabIndex        =   424
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   58
            Left            =   2520
            TabIndex        =   423
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   58
            Left            =   480
            TabIndex        =   422
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "茶"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   -70200
         TabIndex        =   340
         Top             =   600
         Width           =   3975
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   57
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   413
            Top             =   3840
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   57
            Left            =   840
            TabIndex        =   412
            Top             =   3840
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   57
            Left            =   1800
            TabIndex        =   411
            Text            =   "1"
            Top             =   3840
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   56
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   406
            Top             =   3480
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   56
            Left            =   840
            TabIndex        =   405
            Top             =   3480
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   56
            Left            =   1800
            TabIndex        =   404
            Text            =   "1"
            Top             =   3480
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   55
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   399
            Top             =   3120
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   55
            Left            =   840
            TabIndex        =   398
            Top             =   3120
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   55
            Left            =   1800
            TabIndex        =   397
            Text            =   "1"
            Top             =   3120
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   54
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   392
            Top             =   2760
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   54
            Left            =   840
            TabIndex        =   391
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   54
            Left            =   1800
            TabIndex        =   390
            Text            =   "1"
            Top             =   2760
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   53
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   385
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   53
            Left            =   840
            TabIndex        =   384
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   53
            Left            =   1800
            TabIndex        =   383
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   52
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   378
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   52
            Left            =   840
            TabIndex        =   377
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   52
            Left            =   1800
            TabIndex        =   376
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   51
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   371
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   51
            Left            =   840
            TabIndex        =   370
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   51
            Left            =   1800
            TabIndex        =   369
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   50
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   364
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   50
            Left            =   840
            TabIndex        =   363
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   50
            Left            =   1800
            TabIndex        =   362
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   49
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   357
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   49
            Left            =   840
            TabIndex        =   356
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   49
            Left            =   1800
            TabIndex        =   355
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   48
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   350
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   48
            Left            =   840
            TabIndex        =   349
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   48
            Left            =   1800
            TabIndex        =   348
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   47
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   343
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   47
            Left            =   840
            TabIndex        =   342
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   47
            Left            =   1800
            TabIndex        =   341
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   57
            Left            =   480
            TabIndex        =   417
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   57
            Left            =   2520
            TabIndex        =   416
            Top             =   3840
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   58
            Left            =   1200
            TabIndex        =   415
            Top             =   3960
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   57
            Left            =   120
            TabIndex        =   414
            Top             =   3840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   56
            Left            =   480
            TabIndex        =   410
            Top             =   3480
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   56
            Left            =   2520
            TabIndex        =   409
            Top             =   3480
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   57
            Left            =   1200
            TabIndex        =   408
            Top             =   3600
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   56
            Left            =   120
            TabIndex        =   407
            Top             =   3480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   55
            Left            =   480
            TabIndex        =   403
            Top             =   3120
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   55
            Left            =   2520
            TabIndex        =   402
            Top             =   3120
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   56
            Left            =   1200
            TabIndex        =   401
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   55
            Left            =   120
            TabIndex        =   400
            Top             =   3120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   54
            Left            =   480
            TabIndex        =   396
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   54
            Left            =   2520
            TabIndex        =   395
            Top             =   2760
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   55
            Left            =   1200
            TabIndex        =   394
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   54
            Left            =   120
            TabIndex        =   393
            Top             =   2760
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   53
            Left            =   480
            TabIndex        =   389
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   53
            Left            =   2520
            TabIndex        =   388
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   54
            Left            =   1200
            TabIndex        =   387
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   386
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   52
            Left            =   480
            TabIndex        =   382
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   52
            Left            =   2520
            TabIndex        =   381
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   53
            Left            =   1200
            TabIndex        =   380
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   52
            Left            =   120
            TabIndex        =   379
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   51
            Left            =   480
            TabIndex        =   375
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   51
            Left            =   2520
            TabIndex        =   374
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   52
            Left            =   1200
            TabIndex        =   373
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   51
            Left            =   120
            TabIndex        =   372
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   50
            Left            =   480
            TabIndex        =   368
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   50
            Left            =   2520
            TabIndex        =   367
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   51
            Left            =   1200
            TabIndex        =   366
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   50
            Left            =   120
            TabIndex        =   365
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   49
            Left            =   480
            TabIndex        =   361
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   49
            Left            =   2520
            TabIndex        =   360
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   50
            Left            =   1200
            TabIndex        =   359
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   49
            Left            =   120
            TabIndex        =   358
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   48
            Left            =   480
            TabIndex        =   354
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   48
            Left            =   2520
            TabIndex        =   353
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   49
            Left            =   1200
            TabIndex        =   352
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   48
            Left            =   120
            TabIndex        =   351
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   47
            Left            =   480
            TabIndex        =   347
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   47
            Left            =   2520
            TabIndex        =   346
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   48
            Left            =   1200
            TabIndex        =   345
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   47
            Left            =   120
            TabIndex        =   344
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "咖啡"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -74520
         TabIndex        =   276
         Top             =   600
         Width           =   3975
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   46
            Left            =   1800
            TabIndex        =   335
            Text            =   "1"
            Top             =   3120
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   46
            Left            =   840
            TabIndex        =   334
            Top             =   3120
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   46
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   333
            Top             =   3120
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   45
            Left            =   1800
            TabIndex        =   328
            Text            =   "1"
            Top             =   2760
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   45
            Left            =   840
            TabIndex        =   327
            Top             =   2760
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   45
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   326
            Top             =   2760
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   44
            Left            =   1800
            TabIndex        =   321
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   44
            Left            =   840
            TabIndex        =   320
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   44
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   319
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   43
            Left            =   1800
            TabIndex        =   314
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   43
            Left            =   840
            TabIndex        =   313
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   43
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   312
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   42
            Left            =   1800
            TabIndex        =   307
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   42
            Left            =   840
            TabIndex        =   306
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   42
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   305
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   41
            Left            =   1800
            TabIndex        =   300
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   41
            Left            =   840
            TabIndex        =   299
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   41
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   298
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   40
            Left            =   1800
            TabIndex        =   293
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   40
            Left            =   840
            TabIndex        =   292
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   40
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   291
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   39
            Left            =   1800
            TabIndex        =   286
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   39
            Left            =   840
            TabIndex        =   285
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   39
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   284
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   38
            Left            =   1800
            TabIndex        =   279
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   38
            Left            =   840
            TabIndex        =   278
            Top             =   240
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   38
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   277
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   46
            Left            =   120
            TabIndex        =   339
            Top             =   3120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   47
            Left            =   1200
            TabIndex        =   338
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   46
            Left            =   2520
            TabIndex        =   337
            Top             =   3120
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   46
            Left            =   480
            TabIndex        =   336
            Top             =   3120
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   45
            Left            =   120
            TabIndex        =   332
            Top             =   2760
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   46
            Left            =   1200
            TabIndex        =   331
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   45
            Left            =   2520
            TabIndex        =   330
            Top             =   2760
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   45
            Left            =   480
            TabIndex        =   329
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   44
            Left            =   120
            TabIndex        =   325
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   45
            Left            =   1200
            TabIndex        =   324
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   44
            Left            =   2520
            TabIndex        =   323
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   44
            Left            =   480
            TabIndex        =   322
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   43
            Left            =   120
            TabIndex        =   318
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   44
            Left            =   1200
            TabIndex        =   317
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   43
            Left            =   2520
            TabIndex        =   316
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   43
            Left            =   480
            TabIndex        =   315
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   311
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   43
            Left            =   1200
            TabIndex        =   310
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   42
            Left            =   2520
            TabIndex        =   309
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   42
            Left            =   480
            TabIndex        =   308
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   41
            Left            =   120
            TabIndex        =   304
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   42
            Left            =   1200
            TabIndex        =   303
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   41
            Left            =   2520
            TabIndex        =   302
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   41
            Left            =   480
            TabIndex        =   301
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   297
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   41
            Left            =   1200
            TabIndex        =   296
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   40
            Left            =   2520
            TabIndex        =   295
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   40
            Left            =   480
            TabIndex        =   294
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   39
            Left            =   120
            TabIndex        =   290
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   40
            Left            =   1200
            TabIndex        =   289
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   39
            Left            =   2520
            TabIndex        =   288
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   39
            Left            =   480
            TabIndex        =   287
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   283
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   39
            Left            =   1200
            TabIndex        =   282
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   38
            Left            =   2520
            TabIndex        =   281
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   38
            Left            =   480
            TabIndex        =   280
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "其他類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   8880
         TabIndex        =   212
         Top             =   720
         Width           =   4095
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   37
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   271
            Top             =   3120
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   37
            Left            =   840
            TabIndex        =   270
            Top             =   3120
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   37
            Left            =   1800
            TabIndex        =   269
            Text            =   "1"
            Top             =   3120
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   36
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   264
            Top             =   2760
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   36
            Left            =   840
            TabIndex        =   263
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   36
            Left            =   1800
            TabIndex        =   262
            Text            =   "1"
            Top             =   2760
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   35
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   257
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   35
            Left            =   840
            TabIndex        =   256
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   35
            Left            =   1800
            TabIndex        =   255
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   34
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   250
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   34
            Left            =   840
            TabIndex        =   249
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   34
            Left            =   1800
            TabIndex        =   248
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   33
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   243
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   33
            Left            =   840
            TabIndex        =   242
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   33
            Left            =   1800
            TabIndex        =   241
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   32
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   236
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   32
            Left            =   840
            TabIndex        =   235
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   32
            Left            =   1800
            TabIndex        =   234
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   31
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   229
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   31
            Left            =   840
            TabIndex        =   228
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   31
            Left            =   1800
            TabIndex        =   227
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   30
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   222
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   30
            Left            =   840
            TabIndex        =   221
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   30
            Left            =   1800
            TabIndex        =   220
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   29
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   215
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   29
            Left            =   840
            TabIndex        =   214
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   29
            Left            =   1800
            TabIndex        =   213
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   37
            Left            =   480
            TabIndex        =   275
            Top             =   3120
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   37
            Left            =   2520
            TabIndex        =   274
            Top             =   3120
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   38
            Left            =   1200
            TabIndex        =   273
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   272
            Top             =   3120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   36
            Left            =   480
            TabIndex        =   268
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   36
            Left            =   2520
            TabIndex        =   267
            Top             =   2760
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   37
            Left            =   1200
            TabIndex        =   266
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   36
            Left            =   120
            TabIndex        =   265
            Top             =   2760
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   35
            Left            =   480
            TabIndex        =   261
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   35
            Left            =   2520
            TabIndex        =   260
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   36
            Left            =   1200
            TabIndex        =   259
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   258
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   34
            Left            =   480
            TabIndex        =   254
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   34
            Left            =   2520
            TabIndex        =   253
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   35
            Left            =   1200
            TabIndex        =   252
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   251
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   33
            Left            =   480
            TabIndex        =   247
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   33
            Left            =   2520
            TabIndex        =   246
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   33
            Left            =   1200
            TabIndex        =   245
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   244
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   32
            Left            =   480
            TabIndex        =   240
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   32
            Left            =   2520
            TabIndex        =   239
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   32
            Left            =   1200
            TabIndex        =   238
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   237
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   31
            Left            =   480
            TabIndex        =   233
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   31
            Left            =   2520
            TabIndex        =   232
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   31
            Left            =   1200
            TabIndex        =   231
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   230
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   30
            Left            =   480
            TabIndex        =   226
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   30
            Left            =   2520
            TabIndex        =   225
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   30
            Left            =   1200
            TabIndex        =   224
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   223
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   29
            Left            =   480
            TabIndex        =   219
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   29
            Left            =   2520
            TabIndex        =   218
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   29
            Left            =   1200
            TabIndex        =   217
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   216
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "起士類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   4560
         TabIndex        =   6
         Top             =   720
         Width           =   4215
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   12
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   95
            Top             =   2760
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   12
            Left            =   840
            TabIndex        =   94
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   12
            Left            =   1800
            TabIndex        =   93
            Text            =   "1"
            Top             =   2760
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   11
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   88
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   11
            Left            =   840
            TabIndex        =   87
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   11
            Left            =   1800
            TabIndex        =   86
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   10
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   81
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   10
            Left            =   840
            TabIndex        =   80
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   10
            Left            =   1800
            TabIndex        =   79
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   9
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   74
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   9
            Left            =   840
            TabIndex        =   73
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   9
            Left            =   1800
            TabIndex        =   72
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   8
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   67
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   8
            Left            =   840
            TabIndex        =   66
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   8
            Left            =   1800
            TabIndex        =   65
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   7
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   60
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   7
            Left            =   840
            TabIndex        =   59
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   7
            Left            =   1800
            TabIndex        =   58
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   0
            Left            =   840
            TabIndex        =   14
            Top             =   240
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   0
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   13
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   0
            Left            =   1800
            TabIndex        =   12
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   1
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   11
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   1
            Left            =   840
            TabIndex        =   10
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   1
            Left            =   1800
            TabIndex        =   9
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   12
            Left            =   480
            TabIndex        =   99
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   12
            Left            =   2520
            TabIndex        =   98
            Top             =   2760
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   12
            Left            =   1200
            TabIndex        =   97
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   96
            Top             =   2760
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   11
            Left            =   480
            TabIndex        =   92
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   11
            Left            =   2520
            TabIndex        =   91
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   11
            Left            =   1200
            TabIndex        =   90
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   89
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   10
            Left            =   480
            TabIndex        =   85
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   10
            Left            =   2520
            TabIndex        =   84
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   10
            Left            =   1200
            TabIndex        =   83
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   82
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   9
            Left            =   480
            TabIndex        =   78
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   9
            Left            =   2520
            TabIndex        =   77
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   9
            Left            =   1200
            TabIndex        =   76
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   75
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   8
            Left            =   480
            TabIndex        =   71
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   8
            Left            =   2520
            TabIndex        =   70
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   8
            Left            =   1200
            TabIndex        =   69
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   68
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   64
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   7
            Left            =   2520
            TabIndex        =   63
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   7
            Left            =   1200
            TabIndex        =   62
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   61
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   22
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   21
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   0
            Left            =   2520
            TabIndex        =   20
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   19
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   1
            Left            =   2520
            TabIndex        =   18
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   17
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "蛋糕類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   4560
         TabIndex        =   5
         Top             =   4080
         Width           =   3975
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   17
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   130
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   17
            Left            =   840
            TabIndex        =   129
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   17
            Left            =   1800
            TabIndex        =   128
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   16
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   123
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   16
            Left            =   840
            TabIndex        =   122
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   16
            Left            =   1800
            TabIndex        =   121
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   15
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   116
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   15
            Left            =   840
            TabIndex        =   115
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   15
            Left            =   1800
            TabIndex        =   114
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   14
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   109
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   14
            Left            =   840
            TabIndex        =   108
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   14
            Left            =   1800
            TabIndex        =   107
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   13
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   102
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   13
            Left            =   840
            TabIndex        =   101
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   13
            Left            =   1800
            TabIndex        =   100
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   3
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   32
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   3
            Left            =   840
            TabIndex        =   31
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   3
            Left            =   1800
            TabIndex        =   30
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   2
            Left            =   1800
            TabIndex        =   25
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   2
            Left            =   840
            TabIndex        =   24
            Top             =   240
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   2
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   23
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   17
            Left            =   480
            TabIndex        =   134
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   17
            Left            =   2520
            TabIndex        =   133
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   17
            Left            =   1200
            TabIndex        =   132
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   131
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   16
            Left            =   480
            TabIndex        =   127
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   16
            Left            =   2520
            TabIndex        =   126
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   16
            Left            =   1200
            TabIndex        =   125
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   124
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   15
            Left            =   480
            TabIndex        =   120
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   15
            Left            =   2520
            TabIndex        =   119
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   15
            Left            =   1200
            TabIndex        =   118
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   117
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   14
            Left            =   480
            TabIndex        =   113
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   14
            Left            =   2520
            TabIndex        =   112
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   14
            Left            =   1200
            TabIndex        =   111
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   110
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   13
            Left            =   480
            TabIndex        =   106
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   13
            Left            =   2520
            TabIndex        =   105
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   13
            Left            =   1200
            TabIndex        =   104
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   103
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   36
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   3
            Left            =   2520
            TabIndex        =   35
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   34
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   28
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   2
            Left            =   2520
            TabIndex        =   27
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label5 
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   26
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "慕斯類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   3975
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   23
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   172
            Top             =   3120
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   23
            Left            =   840
            TabIndex        =   171
            Top             =   3120
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   23
            Left            =   1800
            TabIndex        =   170
            Text            =   "1"
            Top             =   3120
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   22
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   165
            Top             =   2760
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   22
            Left            =   840
            TabIndex        =   164
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   22
            Left            =   1800
            TabIndex        =   163
            Text            =   "1"
            Top             =   2760
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   21
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   158
            Top             =   2400
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   21
            Left            =   840
            TabIndex        =   157
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   21
            Left            =   1800
            TabIndex        =   156
            Text            =   "1"
            Top             =   2400
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   20
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   151
            Top             =   2040
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   20
            Left            =   840
            TabIndex        =   150
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   20
            Left            =   1800
            TabIndex        =   149
            Text            =   "1"
            Top             =   2040
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   19
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   144
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   19
            Left            =   840
            TabIndex        =   143
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   19
            Left            =   1800
            TabIndex        =   142
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   18
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   137
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   18
            Left            =   840
            TabIndex        =   136
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   18
            Left            =   1800
            TabIndex        =   135
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   6
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   53
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   6
            Left            =   840
            TabIndex        =   52
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   6
            Left            =   1800
            TabIndex        =   51
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   5
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   46
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   5
            Left            =   840
            TabIndex        =   45
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   5
            Left            =   1800
            TabIndex        =   44
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   4
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   39
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   4
            Left            =   840
            TabIndex        =   38
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   4
            Left            =   1800
            TabIndex        =   37
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   23
            Left            =   480
            TabIndex        =   176
            Top             =   3120
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   23
            Left            =   2520
            TabIndex        =   175
            Top             =   3120
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   23
            Left            =   1200
            TabIndex        =   174
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   173
            Top             =   3120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   22
            Left            =   480
            TabIndex        =   169
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   22
            Left            =   2520
            TabIndex        =   168
            Top             =   2760
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   22
            Left            =   1200
            TabIndex        =   167
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   166
            Top             =   2760
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   21
            Left            =   480
            TabIndex        =   162
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   21
            Left            =   2520
            TabIndex        =   161
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   21
            Left            =   1200
            TabIndex        =   160
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   159
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   20
            Left            =   480
            TabIndex        =   155
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   20
            Left            =   2520
            TabIndex        =   154
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   20
            Left            =   1200
            TabIndex        =   153
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   152
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   19
            Left            =   480
            TabIndex        =   148
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   19
            Left            =   2520
            TabIndex        =   147
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   19
            Left            =   1200
            TabIndex        =   146
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   145
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   18
            Left            =   480
            TabIndex        =   141
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   18
            Left            =   2520
            TabIndex        =   140
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   18
            Left            =   1200
            TabIndex        =   139
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   138
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   57
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   6
            Left            =   2520
            TabIndex        =   56
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   55
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   54
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   50
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   5
            Left            =   2520
            TabIndex        =   49
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   5
            Left            =   1200
            TabIndex        =   48
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   43
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   4
            Left            =   2520
            TabIndex        =   42
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   41
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "巧克力類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   360
         TabIndex        =   3
         Top             =   4320
         Width           =   3855
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   28
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   207
            Top             =   1680
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   28
            Left            =   840
            TabIndex        =   206
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   28
            Left            =   1800
            TabIndex        =   205
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   27
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   200
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   27
            Left            =   840
            TabIndex        =   199
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   27
            Left            =   1800
            TabIndex        =   198
            Text            =   "1"
            Top             =   1320
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   26
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   193
            Top             =   960
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   26
            Left            =   840
            TabIndex        =   192
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   26
            Left            =   1800
            TabIndex        =   191
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   25
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   186
            Top             =   600
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   25
            Left            =   840
            TabIndex        =   185
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   25
            Left            =   1800
            TabIndex        =   184
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   255
            Index           =   24
            Left            =   2160
            Max             =   99
            Min             =   1
            TabIndex        =   179
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Index           =   24
            Left            =   840
            TabIndex        =   178
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   24
            Left            =   1800
            TabIndex        =   177
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   28
            Left            =   480
            TabIndex        =   211
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   28
            Left            =   2520
            TabIndex        =   210
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   28
            Left            =   1200
            TabIndex        =   209
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   208
            Top             =   1680
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   27
            Left            =   480
            TabIndex        =   204
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   27
            Left            =   2520
            TabIndex        =   203
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   27
            Left            =   1200
            TabIndex        =   202
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   201
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   26
            Left            =   480
            TabIndex        =   197
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   26
            Left            =   2520
            TabIndex        =   196
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   26
            Left            =   1200
            TabIndex        =   195
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   194
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   25
            Left            =   480
            TabIndex        =   190
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   25
            Left            =   2520
            TabIndex        =   189
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   25
            Left            =   1200
            TabIndex        =   188
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   187
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '透明
            Height          =   255
            Index           =   24
            Left            =   480
            TabIndex        =   183
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   24
            Left            =   2520
            TabIndex        =   182
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "數量"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   24
            Left            =   1200
            TabIndex        =   181
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   180
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結帳"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   2385
      Left            =   10440
      Picture         =   "Form3.frx":207A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4020
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   735
      Left            =   10560
      TabIndex        =   826
      Top             =   9840
      Visible         =   0   'False
      Width           =   3135
      URL             =   ""
      rate            =   1
      balance         =   -9
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   7
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   60
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5530
      _cy             =   1296
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "銷售紀錄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   0
      TabIndex        =   824
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "NO."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   26.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim SUM As Integer

Private Sub Check1_Click(Index As Integer)
  SUM = 0
  For I = 0 To 114
    N = VScroll1(I).Value
    P = Val(Label5(I))
    If Check1(I).Value Then SUM = SUM + N * P
  Next I
  Label4 = SUM
End Sub

Private Sub Command1_Click()
    If frmLogin.user = "" Then
MsgBox "你忘了登入本系統，請重新登入", 0 + 64, "錯誤訊息"
For A = 0 To 114
Check1(A).Value = False
Label4 = ""
Next A
Else

 DataEnvironment1.銷售
 For I = 0 To 114
    If Check1(I).Value Then
    DataEnvironment1.rs銷售.AddNew
    DataEnvironment1.rs銷售("客戶編號") = DataCombo1
    DataEnvironment1.rs銷售("銷售編號") = Text2
    DataEnvironment1.rs銷售("銷售日期") = Now()
    DataEnvironment1.rs銷售("商品編號") = Label6(I)
    DataEnvironment1.rs銷售("數量") = VScroll1(I).Value
    DataEnvironment1.rs銷售("員工編號") = frmLogin.ID
    DataEnvironment1.rs銷售.Update
    End If
    Next I
    DataEnvironment1.rs銷售.Close
If DataEnvironment1.rs銷售詳細.State <> adStateClosed Then DataEnvironment1.rs銷售詳細.Close
     C = MsgBox("列印單據？", vbYesNo)
If C = vbYes Then
   DataEnvironment1.銷售詳細 Text2
   Set DataReport1.DataSource = DataEnvironment1
  DataReport1.DataMember = "銷售詳細"
   DataReport1.Show
     End If
 Call Form_Load
    MsgBox "歡迎下次在來"
End If
End Sub

Private Sub DataCombo1_Change()
If DataCombo1.Visible = True Then
Command1.Enabled = True
Else
MsgBox "請選擇你的編號", 0 + 18, "注意"
End If
End Sub


Private Sub Form_Load()


 MDIForm1.StatusBar1.Panels(5) = "精選好歌：MY LOVE"
  MDIForm1.StatusBar1.Panels(4) = "銷售商品系統"

 For I = 0 To 114
   Check1(I).Value = False
   VScroll1(I).Value = 1
 Next I
 
 I = 0
 DataEnvironment1.商品
 Do While Not DataEnvironment1.rs商品.EOF
  Set rs = DataEnvironment1.rs銷售詳細
      Label5(I) = DataEnvironment1.rs商品("商品價格")
      Label2(I) = DataEnvironment1.rs商品("商品名稱")
      Label6(I) = DataEnvironment1.rs商品("商品編號")
      DataEnvironment1.rs商品.MoveNext
      I = I + 1
      If I = 114 Then Exit Do
 Loop
 DataEnvironment1.rs商品.Close
 Text2 = Format(Now(), "YYMMDDHHMMSS")

End Sub


Private Sub Label4_Click()
  SUM = 0
  For I = 0 To 114
    N = VScroll1(I).Value
    P = Val(Label5(I))
    If Check1(I).Value Then SUM = SUM + N * P
  Next I
  Label4 = SUM
End Sub

Private Sub Text1_Change(Index As Integer)
  Call Check1_Click(Index)
End Sub

Private Sub VScroll1_Change(Index As Integer)
  Text1(Index) = VScroll1(Index).Value
End Sub

Private Sub Form_Activate()
WindowsMediaPlayer1.Controls.play
WindowsMediaPlayer1.URL = "(鈴聲)Westlife - my love.mp3"
End Sub

