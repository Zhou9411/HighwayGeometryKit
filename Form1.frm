VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "高速公路坐标高程计算软件4.3"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   8295
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "详细说明"
      Height          =   420
      Left            =   1440
      TabIndex        =   235
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Command29 
      Caption         =   "计算器"
      Height          =   420
      Left            =   4920
      TabIndex        =   236
      Top             =   7560
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   238
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "线元法计算"
      TabPicture(0)   =   "Form1.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(13)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(15)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "缓和曲线"
      TabPicture(1)   =   "Form1.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(12)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1(14)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "直线"
      TabPicture(2)   =   "Form1.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(17)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "方位角"
      TabPicture(3)   =   "Form1.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(10)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame1(11)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame1(7)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "竖曲线"
      TabPicture(4)   =   "Form1.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1(19)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame1(18)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame1(6)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "交点法计算平曲线"
      TabPicture(5)   =   "Form1.frx":0D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1(4)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Command39"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame1(5)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "宽度渐变"
      TabPicture(6)   =   "Form1.frx":0D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1(20)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame1(3)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Frame1(1)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "说明"
      TabPicture(7)   =   "Form1.frx":0D8E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Text63"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Command71"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).ControlCount=   2
      Begin VB.CommandButton Command71 
         Caption         =   "赞助该软件"
         Height          =   420
         Left            =   -72240
         TabIndex        =   414
         Top             =   6840
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Caption         =   "已知平曲线要素(起算点可以选择ZH或HZ点，从大半径向小半径R方向推算)"
         Height          =   2175
         Index           =   14
         Left            =   -74880
         TabIndex        =   399
         Top             =   360
         Width           =   7815
         Begin VB.TextBox Text3 
            Height          =   300
            Left            =   6030
            TabIndex        =   55
            ToolTipText     =   "必须为度.分秒的格式"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text4 
            Height          =   300
            Left            =   4530
            TabIndex        =   54
            ToolTipText     =   "D"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text5 
            Height          =   300
            Left            =   3045
            TabIndex        =   53
            ToolTipText     =   "C"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text6 
            Height          =   300
            Left            =   1560
            TabIndex        =   52
            ToolTipText     =   "J"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text7 
            Height          =   300
            Left            =   1605
            TabIndex        =   57
            ToolTipText     =   "注意A不能取零"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text8 
            Height          =   300
            Left            =   120
            TabIndex        =   56
            ToolTipText     =   "R不能取零,左偏取负,右偏取正"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command13 
            Caption         =   "保存到文件"
            Height          =   375
            Left            =   6120
            TabIndex        =   47
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton Command14 
            Caption         =   "从文件载入"
            Height          =   375
            Left            =   4560
            TabIndex        =   46
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Text21 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   3480
            TabIndex        =   50
            ToolTipText     =   "缓和曲线参数文件"
            Top             =   240
            Width           =   4095
         End
         Begin VB.OptionButton Option7 
            Caption         =   "起算点在缓1上"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option8 
            Caption         =   "起算点在缓2上"
            Height          =   255
            Left            =   1800
            TabIndex        =   49
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox Text107 
            Height          =   300
            Left            =   120
            TabIndex        =   51
            ToolTipText     =   "J"
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "起算点切线方位角F"
            Height          =   255
            Index           =   29
            Left            =   6000
            TabIndex        =   406
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "起算点Y坐标D"
            Height          =   255
            Index           =   28
            Left            =   4560
            TabIndex        =   405
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "起算点X坐标C"
            Height          =   255
            Index           =   27
            Left            =   3240
            TabIndex        =   404
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "起算点里程J"
            Height          =   255
            Index           =   26
            Left            =   1680
            TabIndex        =   403
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "回旋参数A"
            Height          =   255
            Index           =   31
            Left            =   1800
            TabIndex        =   402
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "圆曲线半径R"
            Height          =   255
            Index           =   30
            Left            =   240
            TabIndex        =   401
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "HY点里程KHY"
            Height          =   255
            Index           =   134
            Left            =   240
            TabIndex        =   400
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "横断面图生成(土石方戴帽子)"
         Height          =   1095
         Index           =   1
         Left            =   -74880
         TabIndex        =   396
         Top             =   6120
         Width           =   7815
         Begin VB.TextBox Text113 
            Height          =   270
            Left            =   3600
            TabIndex        =   228
            Text            =   "20"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command65 
            Caption         =   "宋体"
            Height          =   375
            Left            =   3360
            TabIndex        =   232
            Top             =   660
            Width           =   735
         End
         Begin VB.CommandButton Command61 
            Caption         =   "载入参数并生成横断面CAD脚本(仿宋)"
            Height          =   375
            Left            =   120
            TabIndex        =   231
            Top             =   660
            Width           =   3255
         End
         Begin VB.CommandButton Command62 
            Caption         =   "编辑设计线"
            Height          =   375
            Left            =   120
            TabIndex        =   226
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command63 
            Caption         =   "编辑地面线"
            Height          =   375
            Left            =   1320
            TabIndex        =   227
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text111 
            Height          =   270
            Left            =   6600
            TabIndex        =   230
            Text            =   "K000+000"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text112 
            Height          =   270
            Left            =   5160
            TabIndex        =   229
            Text            =   "0.4"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command64 
            Caption         =   "载入参数并生成Excel计算表"
            Height          =   375
            Left            =   5040
            TabIndex        =   233
            Top             =   660
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "间距："
            Height          =   195
            Index           =   135
            Left            =   3000
            TabIndex        =   409
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "里程格式："
            Height          =   195
            Index           =   132
            Left            =   5760
            TabIndex        =   398
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "文字高度："
            Height          =   195
            Index           =   133
            Left            =   4320
            TabIndex        =   397
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "特征点计算"
         Height          =   3735
         Index           =   4
         Left            =   -74880
         TabIndex        =   376
         Top             =   3000
         Width           =   7815
         Begin VB.CommandButton Command66 
            Caption         =   "生成交点的CAD脚本"
            Height          =   375
            Left            =   5640
            TabIndex        =   168
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox Text80 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   172
            TabStop         =   0   'False
            ToolTipText     =   "度.分秒的格式"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text81 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   171
            TabStop         =   0   'False
            ToolTipText     =   "Y"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text82 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   170
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command37 
            Caption         =   "导出坐标到txt文本"
            Height          =   375
            Left            =   3720
            TabIndex        =   167
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox Combo5 
            Height          =   300
            Left            =   1200
            TabIndex        =   166
            ToolTipText     =   "必须从第二个交点开始才能计算"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text69 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   176
            TabStop         =   0   'False
            ToolTipText     =   "度.分秒的格式"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Text74 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   175
            TabStop         =   0   'False
            ToolTipText     =   "Y"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Text75 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   174
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Text76 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   180
            TabStop         =   0   'False
            ToolTipText     =   "度.分秒的格式"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text77 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   179
            TabStop         =   0   'False
            ToolTipText     =   "Y"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text78 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   178
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text79 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   184
            TabStop         =   0   'False
            ToolTipText     =   "度.分秒的格式"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox Text83 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   183
            TabStop         =   0   'False
            ToolTipText     =   "Y"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox Text84 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   182
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox Text85 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   181
            TabStop         =   0   'False
            ToolTipText     =   "度.分秒的格式"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox Text86 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   177
            TabStop         =   0   'False
            ToolTipText     =   "度.分秒的格式"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text87 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   173
            TabStop         =   0   'False
            ToolTipText     =   "度.分秒的格式"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Text88 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   169
            TabStop         =   0   'False
            ToolTipText     =   "度.分秒的格式"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text89 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   0
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   186
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   2520
            Width           =   920
         End
         Begin VB.TextBox Text90 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   0
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   188
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   2520
            Width           =   920
         End
         Begin VB.TextBox Text91 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   0
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   185
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   2520
            Width           =   920
         End
         Begin VB.TextBox Text89 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   1
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   187
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   2520
            Width           =   920
         End
         Begin VB.TextBox Text91 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   1
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   189
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   2880
            Width           =   920
         End
         Begin VB.TextBox Text90 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   1
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   192
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   2880
            Width           =   920
         End
         Begin VB.TextBox Text89 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   2
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   190
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   2880
            Width           =   920
         End
         Begin VB.TextBox Text89 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   3
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   191
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   2880
            Width           =   920
         End
         Begin VB.TextBox Text91 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   2
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   193
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   3240
            Width           =   920
         End
         Begin VB.TextBox Text91 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   3
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   194
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   3240
            Width           =   920
         End
         Begin VB.Label Label1 
            Caption         =   "交点序号             对应的特征点：  "
            Height          =   375
            Index           =   109
            Left            =   360
            TabIndex        =   395
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "切线方位角Q"
            Height          =   255
            Index           =   114
            Left            =   5520
            TabIndex        =   394
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Y坐标"
            Height          =   255
            Index           =   115
            Left            =   4200
            TabIndex        =   393
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "X坐标"
            Height          =   255
            Index           =   116
            Left            =   2640
            TabIndex        =   392
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "ZH点："
            Height          =   255
            Left            =   240
            TabIndex        =   391
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "HY点："
            Height          =   255
            Left            =   240
            TabIndex        =   390
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "YH点："
            Height          =   255
            Left            =   240
            TabIndex        =   389
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "HZ点："
            Height          =   255
            Left            =   240
            TabIndex        =   388
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "里程K"
            Height          =   255
            Index           =   9
            Left            =   1200
            TabIndex        =   387
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "回旋参数A="
            Height          =   255
            Index           =   100
            Left            =   1800
            TabIndex        =   386
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "圆曲线长O="
            Height          =   255
            Index           =   104
            Left            =   5760
            TabIndex        =   385
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "圆半径R="
            Height          =   255
            Index           =   105
            Left            =   120
            TabIndex        =   384
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "回旋参数B="
            Height          =   255
            Index           =   124
            Left            =   3720
            TabIndex        =   383
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "缓1Ls1="
            Height          =   255
            Index           =   125
            Left            =   120
            TabIndex        =   382
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "切线长T2="
            Height          =   255
            Index           =   127
            Left            =   5760
            TabIndex        =   381
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "缓2Ls2="
            Height          =   255
            Index           =   128
            Left            =   2040
            TabIndex        =   380
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "切线长T1="
            Height          =   255
            Index           =   129
            Left            =   3840
            TabIndex        =   379
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "L总="
            Height          =   255
            Index           =   130
            Left            =   360
            TabIndex        =   378
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "偏角α="
            Height          =   255
            Index           =   131
            Left            =   2040
            TabIndex        =   377
            Top             =   3240
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "宽度参数"
         Height          =   3855
         Index           =   3
         Left            =   -74880
         TabIndex        =   365
         Top             =   360
         Width           =   7815
         Begin VB.CommandButton Command70 
            Caption         =   "9860"
            Height          =   375
            Left            =   3360
            TabIndex        =   199
            Top             =   240
            Width           =   615
         End
         Begin VB.Frame Frame1 
            Caption         =   "左幅宽度参数"
            Height          =   1455
            Index           =   21
            Left            =   120
            TabIndex        =   371
            Top             =   720
            Width           =   7575
            Begin VB.OptionButton Option3 
               Caption         =   "三次抛物线渐变"
               Height          =   255
               Left            =   4800
               TabIndex        =   205
               Top             =   720
               Width           =   1575
            End
            Begin VB.OptionButton Option4 
               Caption         =   "线性渐变"
               Height          =   255
               Left            =   4800
               TabIndex        =   206
               Top             =   1080
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.TextBox Text95 
               Height          =   300
               Left            =   3216
               TabIndex        =   204
               ToolTipText     =   "I向路基外流水取正值,向中桩流水取负值"
               Top             =   960
               Width           =   1335
            End
            Begin VB.CommandButton Command49 
               Caption         =   "输入"
               Height          =   375
               Left            =   6480
               TabIndex        =   207
               ToolTipText     =   "注意各段的参数修改后一定要点输入按钮并注意存盘"
               Top             =   960
               Width           =   975
            End
            Begin VB.ComboBox Combo7 
               Height          =   300
               Left            =   240
               TabIndex        =   202
               ToolTipText     =   "第几个分段"
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox Text96 
               Height          =   300
               Left            =   1728
               TabIndex        =   203
               ToolTipText     =   "取前后两个竖曲线之间的直线上的任意一点里程"
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox Text97 
               Height          =   300
               Left            =   1800
               TabIndex        =   201
               ToolTipText     =   "分段总数"
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "左幅宽度"
               Height          =   255
               Index           =   107
               Left            =   3480
               TabIndex        =   375
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "分段序号"
               Height          =   255
               Index           =   108
               Left            =   480
               TabIndex        =   374
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "分段终点Ki"
               Height          =   255
               Index           =   110
               Left            =   1920
               TabIndex        =   373
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "左幅宽度分段数目："
               Height          =   255
               Index           =   111
               Left            =   120
               TabIndex        =   372
               Top             =   300
               Width           =   1695
            End
         End
         Begin VB.CommandButton Command55 
            Caption         =   "5800"
            Height          =   375
            Left            =   2640
            TabIndex        =   198
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton Command50 
            Caption         =   "保存到文件"
            Height          =   375
            Left            =   1320
            TabIndex        =   197
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command51 
            Caption         =   "载入文件"
            Height          =   375
            Left            =   120
            TabIndex        =   196
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text99 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4440
            TabIndex        =   200
            ToolTipText     =   "横坡参数文件"
            Top             =   240
            Width           =   3255
         End
         Begin VB.Frame Frame1 
            Caption         =   "右幅宽度参数"
            Height          =   1455
            Index           =   2
            Left            =   120
            TabIndex        =   366
            Top             =   2280
            Width           =   7575
            Begin VB.TextBox Text94 
               Height          =   300
               Left            =   3240
               TabIndex        =   211
               ToolTipText     =   "I向路基外流水取正值,向中桩流水取负值"
               Top             =   960
               Width           =   1335
            End
            Begin VB.OptionButton Option5 
               Caption         =   "三次抛物线渐变"
               Height          =   255
               Left            =   4800
               TabIndex        =   212
               Top             =   720
               Width           =   1575
            End
            Begin VB.OptionButton Option6 
               Caption         =   "线性渐变"
               Height          =   255
               Left            =   4800
               TabIndex        =   213
               Top             =   1080
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.CommandButton Command58 
               Caption         =   "输入"
               Height          =   375
               Left            =   6480
               TabIndex        =   214
               ToolTipText     =   "注意各段的参数修改后一定要点输入按钮并注意存盘"
               Top             =   960
               Width           =   975
            End
            Begin VB.ComboBox Combo8 
               Height          =   300
               Left            =   240
               TabIndex        =   209
               ToolTipText     =   "第几个分段"
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox Text101 
               Height          =   300
               Left            =   1725
               TabIndex        =   210
               ToolTipText     =   "取前后两个竖曲线之间的直线上的任意一点里程"
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox Text102 
               Height          =   300
               Left            =   1800
               TabIndex        =   208
               ToolTipText     =   "分段总数"
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "右幅宽度"
               Height          =   255
               Index           =   106
               Left            =   3510
               TabIndex        =   370
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "分段序号"
               Height          =   255
               Index           =   113
               Left            =   480
               TabIndex        =   369
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "分段终点Ki"
               Height          =   255
               Index           =   117
               Left            =   1920
               TabIndex        =   368
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "右幅宽度分段数目："
               Height          =   255
               Index           =   123
               Left            =   120
               TabIndex        =   367
               Top             =   300
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "所求点K参数及结果"
         Height          =   1680
         Index           =   20
         Left            =   -74880
         TabIndex        =   357
         Top             =   4320
         Width           =   7815
         Begin VB.TextBox Text98 
            Height          =   270
            Left            =   3840
            TabIndex        =   221
            Text            =   "3"
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton Command52 
            Caption         =   "编辑里程文件"
            Height          =   375
            Left            =   120
            TabIndex        =   223
            ToolTipText     =   "编辑现有的zap数据文件"
            Top             =   1200
            Width           =   1275
         End
         Begin VB.TextBox Text100 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4680
            TabIndex        =   225
            ToolTipText     =   "横坡参数文件"
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CommandButton Command53 
            Caption         =   "从数据文件输入里程并导出到Excel中"
            Height          =   375
            Left            =   1470
            TabIndex        =   224
            Top             =   1200
            Width           =   3180
         End
         Begin VB.TextBox Text103 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   220
            TabStop         =   0   'False
            ToolTipText     =   "边桩对应里程KG的设计高程"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text104 
            Height          =   300
            Left            =   960
            TabIndex        =   215
            ToolTipText     =   "Kstart"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text105 
            Height          =   300
            Left            =   2520
            TabIndex        =   216
            ToolTipText     =   "Kend"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text106 
            Height          =   300
            Left            =   4440
            TabIndex        =   217
            Text            =   "10"
            ToolTipText     =   "间距不能取零"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command54 
            Caption         =   "计算起点"
            Height          =   375
            Left            =   6240
            TabIndex        =   218
            ToolTipText     =   "只计算Kstart"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command57 
            Caption         =   "导出宽度到Excel文件"
            Height          =   375
            Left            =   5640
            TabIndex        =   222
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text110 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   219
            ToolTipText     =   "边桩对应的里程KG"
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "位小数"
            Height          =   255
            Index           =   137
            Left            =   4320
            TabIndex        =   411
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "右幅宽度"
            Height          =   255
            Index           =   118
            Left            =   1920
            TabIndex        =   364
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "里程K="
            Height          =   195
            Index           =   119
            Left            =   360
            TabIndex        =   363
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "～"
            Height          =   255
            Index           =   120
            Left            =   2280
            TabIndex        =   362
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "间隔："
            Height          =   255
            Index           =   121
            Left            =   3960
            TabIndex        =   361
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "m"
            Height          =   255
            Index           =   122
            Left            =   5200
            TabIndex        =   360
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "左幅宽度"
            Height          =   255
            Index           =   126
            Left            =   240
            TabIndex        =   359
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "保留"
            Height          =   255
            Index           =   112
            Left            =   3360
            TabIndex        =   358
            Top             =   840
            Width           =   495
         End
      End
      Begin VB.TextBox Text63 
         BackColor       =   &H8000000A&
         Height          =   6255
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   356
         Text            =   "Form1.frx":0DAA
         Top             =   480
         Width           =   7815
      End
      Begin VB.Frame Frame1 
         Caption         =   "从数据文件输入里程偏距偏角来计算"
         Height          =   1335
         Index           =   15
         Left            =   120
         TabIndex        =   355
         Top             =   5880
         Width           =   7815
         Begin VB.CommandButton Command45 
            Caption         =   "编辑现有的wsc数据文件"
            Height          =   375
            Left            =   360
            TabIndex        =   43
            Top             =   840
            Width           =   3195
         End
         Begin VB.TextBox Text93 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   3960
            TabIndex        =   42
            ToolTipText     =   "平曲线参数文件"
            Top             =   240
            Width           =   3495
         End
         Begin VB.CommandButton Command44 
            Caption         =   "从wsc/excel文件载入里程偏距偏角"
            Height          =   375
            Left            =   360
            TabIndex        =   41
            Top             =   240
            Width           =   3200
         End
         Begin VB.CommandButton Command42 
            Caption         =   "导出到Excel"
            Height          =   375
            Left            =   5760
            TabIndex        =   45
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton Command41 
            Caption         =   "生成CAD脚本"
            Height          =   375
            Left            =   3840
            TabIndex        =   44
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.CommandButton Command39 
         Caption         =   "生成线元法平曲线参数（仅供参考，ini格式）"
         Height          =   375
         Left            =   -72960
         TabIndex        =   195
         Top             =   6840
         Width           =   4335
      End
      Begin VB.Frame Frame1 
         Caption         =   "已知平曲线要素"
         Height          =   2535
         Index           =   5
         Left            =   -74880
         TabIndex        =   344
         Top             =   360
         Width           =   7815
         Begin VB.TextBox Text109 
            Height          =   300
            Left            =   240
            TabIndex        =   158
            ToolTipText     =   "Ls=缓1=缓2，没有缓和曲线可取0"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox Text108 
            Height          =   300
            Left            =   1875
            TabIndex        =   159
            ToolTipText     =   "T"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton Command60 
            Caption         =   "推算交点里程"
            Height          =   375
            Left            =   6480
            TabIndex        =   163
            Top             =   1440
            Width           =   1240
         End
         Begin VB.CommandButton Command59 
            Caption         =   "检验R正负"
            Height          =   375
            Left            =   5430
            TabIndex        =   162
            Top             =   1440
            Width           =   1035
         End
         Begin VB.TextBox Text73 
            Height          =   300
            Left            =   1845
            TabIndex        =   151
            ToolTipText     =   "第一个交点可取夹直线上任意一点"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text72 
            Height          =   300
            Left            =   3330
            TabIndex        =   152
            ToolTipText     =   "X"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text71 
            Height          =   300
            Left            =   4830
            TabIndex        =   153
            ToolTipText     =   "Y"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text70 
            Height          =   300
            Left            =   1875
            TabIndex        =   156
            ToolTipText     =   "T"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command34 
            Caption         =   "输入"
            Height          =   375
            Left            =   4800
            TabIndex        =   161
            ToolTipText     =   "注意各段的参数修改后一定要点输入按钮并注意存盘"
            Top             =   1440
            Width           =   615
         End
         Begin VB.ComboBox Combo6 
            Height          =   300
            Left            =   240
            TabIndex        =   150
            ToolTipText     =   "第一个交点可取夹直线上任意一点，T,R,Ls,L不用输入"
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox Text68 
            Height          =   300
            Left            =   3360
            TabIndex        =   157
            ToolTipText     =   "L=缓1+圆+缓2"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox Text67 
            Height          =   300
            Left            =   240
            TabIndex        =   155
            ToolTipText     =   "Ls=缓1=缓2，没有缓和曲线可取0"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox Text66 
            Height          =   300
            Left            =   6240
            TabIndex        =   154
            ToolTipText     =   "R不能取零,左偏取负,右偏取正"
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton Command33 
            Caption         =   "保存到文件"
            Height          =   375
            Left            =   6360
            TabIndex        =   165
            Top             =   1920
            Width           =   1360
         End
         Begin VB.CommandButton Command32 
            Caption         =   "从文件载入"
            Height          =   375
            Left            =   4800
            TabIndex        =   164
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox Text65 
            Height          =   300
            Left            =   1200
            TabIndex        =   147
            Text            =   "3"
            ToolTipText     =   "交点总数"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text64 
            Height          =   300
            Left            =   3120
            TabIndex        =   148
            Text            =   "10"
            ToolTipText     =   "第一个交点序号，必须是自然数"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4080
            TabIndex        =   149
            ToolTipText     =   "平曲线参数文件"
            Top             =   240
            Width           =   3615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "手工输入AB"
            Height          =   255
            Left            =   3360
            TabIndex        =   160
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "回旋参数A"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   408
            Top             =   1800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label LabelT 
            Caption         =   "回旋参数B"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   407
            Top             =   1800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "交点里程"
            Height          =   255
            Index           =   103
            Left            =   2160
            TabIndex        =   354
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "交点X坐标"
            Height          =   255
            Index           =   102
            Left            =   3600
            TabIndex        =   353
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "交点Y坐标"
            Height          =   255
            Index           =   101
            Left            =   5040
            TabIndex        =   352
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label LabelT 
            Caption         =   "缓2长Ls2"
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   351
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "交点序号"
            Height          =   375
            Index           =   12
            Left            =   480
            TabIndex        =   350
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "曲线长度L"
            Height          =   255
            Left            =   3600
            TabIndex        =   349
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "缓1长Ls1"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   348
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "圆曲线半径R"
            Height          =   255
            Left            =   6360
            TabIndex        =   347
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "交点数目："
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   346
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "起始交点号："
            Height          =   255
            Index           =   10
            Left            =   1920
            TabIndex        =   345
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "方位角"
         Height          =   2775
         Index           =   7
         Left            =   -74640
         TabIndex        =   335
         Top             =   360
         Width           =   7335
         Begin VB.TextBox Text22 
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   89
            ToolTipText     =   "格式为度.分秒，如150.03125=150°03′12.5″"
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox Text23 
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   88
            ToolTipText     =   "保留了三位小数"
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Frame Frame1 
            Caption         =   "P2点"
            Height          =   1335
            Index           =   9
            Left            =   3840
            TabIndex        =   339
            Top             =   240
            Width           =   2295
            Begin VB.TextBox TextY2 
               Height          =   375
               Left            =   600
               TabIndex        =   87
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox TextX2 
               Height          =   375
               Left            =   600
               TabIndex        =   86
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Y2："
               Height          =   255
               Index           =   54
               Left            =   120
               TabIndex        =   341
               Top             =   840
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "X2："
               Height          =   255
               Index           =   53
               Left            =   120
               TabIndex        =   340
               Top             =   360
               Width           =   375
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "P1点"
            Height          =   1335
            Index           =   8
            Left            =   840
            TabIndex        =   336
            Top             =   240
            Width           =   2295
            Begin VB.TextBox TextY1 
               Height          =   375
               Left            =   600
               TabIndex        =   85
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox TextX1 
               Height          =   375
               Left            =   600
               TabIndex        =   84
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Y1："
               Height          =   255
               Index           =   52
               Left            =   120
               TabIndex        =   338
               Top             =   840
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "X1："
               Height          =   255
               Index           =   51
               Left            =   120
               TabIndex        =   337
               Top             =   360
               Width           =   375
            End
         End
         Begin VB.Label Label1 
            Caption         =   "P1->P2的方位角为："
            Height          =   255
            Index           =   56
            Left            =   1080
            TabIndex        =   343
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "P1->P2的距离为："
            Height          =   255
            Index           =   55
            Left            =   1200
            TabIndex        =   342
            Top             =   1800
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "直线计算"
         Height          =   3015
         Index           =   17
         Left            =   -74760
         TabIndex        =   328
         Top             =   480
         Width           =   7575
         Begin VB.CommandButton Command31 
            Caption         =   "复制方位角中P1->P2的方位角"
            Height          =   375
            Left            =   3480
            TabIndex        =   81
            Top             =   1800
            Width           =   2655
         End
         Begin VB.CommandButton Command30 
            Caption         =   "复制方位角中P2点"
            Height          =   375
            Left            =   5400
            TabIndex        =   76
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton Command10 
            Caption         =   "复制方位角中P1点"
            Height          =   375
            Left            =   3480
            TabIndex        =   75
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox Text62 
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   3240
            TabIndex        =   83
            Text            =   "Text62"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox Text61 
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   1320
            TabIndex        =   82
            Text            =   "Text61"
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox Text60 
            Height          =   375
            Left            =   1680
            TabIndex        =   80
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox Text59 
            Height          =   375
            Left            =   1680
            TabIndex        =   79
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Text58 
            Height          =   375
            Left            =   1680
            TabIndex        =   78
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text57 
            Height          =   375
            Left            =   1680
            TabIndex        =   77
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Y2="
            Height          =   255
            Index           =   50
            Left            =   2880
            TabIndex        =   334
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "P2坐标：X2="
            Height          =   255
            Index           =   49
            Left            =   240
            TabIndex        =   333
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "P1到P2的方位角="
            Height          =   375
            Index           =   48
            Left            =   240
            TabIndex        =   332
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "P1到P2的距离="
            Height          =   255
            Index           =   47
            Left            =   240
            TabIndex        =   331
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Y1="
            Height          =   255
            Index           =   46
            Left            =   1080
            TabIndex        =   330
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "P1坐标：X1="
            Height          =   255
            Index           =   45
            Left            =   360
            TabIndex        =   329
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "角度转换"
         Height          =   1335
         Index           =   11
         Left            =   -74640
         TabIndex        =   321
         Top             =   5280
         Width           =   7335
         Begin VB.TextBox Text56 
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   3360
            TabIndex        =   98
            Text            =   "0"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Text55 
            Height          =   375
            Left            =   480
            TabIndex        =   97
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Text54 
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   3360
            TabIndex        =   96
            Text            =   "0"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox Text53 
            Height          =   375
            Left            =   480
            TabIndex        =   95
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "（度.分秒）"
            Height          =   255
            Index           =   68
            Left            =   5160
            TabIndex        =   327
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "（度）     ="
            Height          =   255
            Index           =   67
            Left            =   2160
            TabIndex        =   326
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "角1"
            Height          =   255
            Index           =   69
            Left            =   120
            TabIndex        =   325
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "（度）"
            Height          =   255
            Index           =   71
            Left            =   5160
            TabIndex        =   324
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "（度.分秒）="
            Height          =   255
            Index           =   70
            Left            =   2160
            TabIndex        =   323
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "角1"
            Height          =   255
            Index           =   66
            Left            =   120
            TabIndex        =   322
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "角度运算"
         Height          =   1935
         Index           =   10
         Left            =   -74640
         TabIndex        =   311
         Top             =   3240
         Width           =   7335
         Begin VB.TextBox Text52 
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   480
            TabIndex        =   93
            Text            =   "Text52"
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox Text51 
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   4560
            TabIndex        =   94
            Text            =   "Text51"
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox Text50 
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   4560
            TabIndex        =   92
            Text            =   "Text50"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text49 
            Height          =   375
            Left            =   2520
            TabIndex        =   91
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text48 
            Height          =   375
            Left            =   480
            TabIndex        =   90
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "=角3"
            Height          =   255
            Index           =   63
            Left            =   4080
            TabIndex        =   320
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "（弧度）"
            Height          =   255
            Index           =   62
            Left            =   2160
            TabIndex        =   319
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "（度）"
            Height          =   255
            Index           =   64
            Left            =   6120
            TabIndex        =   318
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "=角3"
            Height          =   255
            Index           =   61
            Left            =   120
            TabIndex        =   317
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "（度.分秒）"
            Height          =   255
            Index           =   60
            Left            =   6120
            TabIndex        =   316
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "说明：角度1、2格式均为度.分秒格式。13°12′15.5″应输入13.12155。"
            Height          =   375
            Index           =   65
            Left            =   240
            TabIndex        =   315
            Top             =   1440
            Width           =   6015
         End
         Begin VB.Label Label1 
            Caption         =   "=角3"
            Height          =   255
            Index           =   59
            Left            =   4080
            TabIndex        =   314
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "+角2"
            Height          =   255
            Index           =   58
            Left            =   2160
            TabIndex        =   313
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "角1"
            Height          =   255
            Index           =   57
            Left            =   120
            TabIndex        =   312
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "所求点K参数及结果"
         Height          =   3000
         Index           =   19
         Left            =   -74880
         TabIndex        =   297
         Top             =   4250
         Width           =   7815
         Begin VB.TextBox Text115 
            Height          =   300
            Left            =   5280
            TabIndex        =   132
            Text            =   "3"
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton Command67 
            Caption         =   "编辑zas"
            Height          =   375
            Left            =   920
            TabIndex        =   144
            ToolTipText     =   "正交多偏距"
            Top             =   2520
            Width           =   795
         End
         Begin VB.CommandButton Command46 
            Caption         =   "编辑zap"
            Height          =   375
            Left            =   70
            TabIndex        =   143
            ToolTipText     =   "斜交单偏距"
            Top             =   2520
            Width           =   795
         End
         Begin VB.TextBox Text92 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   5520
            TabIndex        =   146
            ToolTipText     =   "横坡参数文件"
            Top             =   2520
            Width           =   2175
         End
         Begin VB.CommandButton Command43 
            Caption         =   "从文件输入里程偏距偏角并导出到Excel中"
            Height          =   375
            Left            =   1800
            TabIndex        =   145
            Top             =   2520
            Width           =   3660
         End
         Begin VB.TextBox Text35 
            Height          =   300
            Left            =   1920
            TabIndex        =   130
            Text            =   "1"
            ToolTipText     =   "T为边桩到中桩的距离,左幅取负,右幅取正"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox TEXTI 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   137
            ToolTipText     =   "I向路基外流水取正值,向中桩流水取负值"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text36 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   136
            TabStop         =   0   'False
            ToolTipText     =   "边桩对应里程KG的设计高程"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text37 
            Height          =   300
            Left            =   960
            TabIndex        =   126
            ToolTipText     =   "Kstart"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text38 
            Height          =   300
            Left            =   2520
            TabIndex        =   127
            ToolTipText     =   "Kend"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text39 
            Height          =   300
            Left            =   4440
            TabIndex        =   128
            Text            =   "20"
            ToolTipText     =   "间距不能取零"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text40 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   138
            TabStop         =   0   'False
            ToolTipText     =   "边桩高程G"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command23 
            Caption         =   "计算起点"
            Height          =   375
            Left            =   6240
            TabIndex        =   134
            ToolTipText     =   "只计算Kstart"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command24 
            Caption         =   "导出标高到txt文本"
            Height          =   375
            Left            =   3240
            TabIndex        =   141
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton Command25 
            Caption         =   "生成中桩高程的CAD脚本文件"
            Height          =   375
            Left            =   240
            TabIndex        =   140
            Top             =   2040
            Width           =   2775
         End
         Begin VB.CommandButton Command28 
            Caption         =   "导出标高到Excel文件"
            Height          =   375
            Left            =   5520
            TabIndex        =   142
            Top             =   2040
            Width           =   2055
         End
         Begin VB.TextBox Text43 
            Height          =   300
            Left            =   240
            TabIndex        =   129
            Text            =   "1"
            ToolTipText     =   "设计高程到中桩的距离"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text44 
            Height          =   300
            Left            =   3480
            TabIndex        =   131
            Text            =   "90"
            ToolTipText     =   "右角，度.分秒的格式"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text45 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   135
            ToolTipText     =   "边桩对应的里程KG"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text46 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   139
            TabStop         =   0   'False
            ToolTipText     =   "边桩高程G2"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text47 
            Height          =   300
            Left            =   6360
            TabIndex        =   133
            Text            =   "0.18"
            ToolTipText     =   "修改后边桩高程G2实时改变"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "小数位"
            Height          =   255
            Index           =   136
            Left            =   5280
            TabIndex        =   410
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "偏距T"
            Height          =   255
            Index           =   92
            Left            =   2280
            TabIndex        =   310
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "横坡I"
            Height          =   255
            Index           =   97
            Left            =   3840
            TabIndex        =   309
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "设计高程H"
            Height          =   255
            Index           =   96
            Left            =   2160
            TabIndex        =   308
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "里程K="
            Height          =   375
            Index           =   87
            Left            =   360
            TabIndex        =   307
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "～"
            Height          =   255
            Index           =   88
            Left            =   2280
            TabIndex        =   306
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "间隔："
            Height          =   255
            Index           =   89
            Left            =   3960
            TabIndex        =   305
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "m"
            Height          =   255
            Index           =   90
            Left            =   5200
            TabIndex        =   304
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "边桩高程G"
            Height          =   255
            Index           =   98
            Left            =   5160
            TabIndex        =   303
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "设计高程到中桩距离TS"
            Height          =   255
            Index           =   91
            Left            =   120
            TabIndex        =   302
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "偏角P"
            Height          =   255
            Index           =   93
            Left            =   3840
            TabIndex        =   301
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "边桩对应里程KG"
            Height          =   255
            Index           =   95
            Left            =   240
            TabIndex        =   300
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "边桩高程G2"
            Height          =   255
            Index           =   99
            Left            =   6600
            TabIndex        =   299
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "边桩向下"
            Height          =   255
            Index           =   94
            Left            =   6600
            TabIndex        =   298
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "横坡参数(左右横坡带百分号)"
         Height          =   1815
         Index           =   18
         Left            =   -74880
         TabIndex        =   290
         Top             =   2370
         Width           =   7815
         Begin VB.OptionButton Option1 
            Caption         =   "线性渐变"
            Height          =   255
            Left            =   240
            TabIndex        =   122
            Top             =   1440
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "三次抛物线渐变"
            Height          =   255
            Left            =   1680
            TabIndex        =   123
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox Text30 
            Height          =   300
            Left            =   4680
            TabIndex        =   121
            ToolTipText     =   "I向路基外流水取正值,向中桩流水取负值"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text31 
            Height          =   300
            Left            =   3216
            TabIndex        =   120
            ToolTipText     =   "I向路基外流水取正值,向中桩流水取负值"
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command20 
            Caption         =   "输入"
            Height          =   375
            Left            =   6240
            TabIndex        =   124
            ToolTipText     =   "注意各段的参数修改后一定要点输入按钮并注意存盘"
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox Combo4 
            Height          =   300
            Left            =   240
            TabIndex        =   118
            ToolTipText     =   "第几个分段"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text32 
            Height          =   300
            Left            =   1728
            TabIndex        =   119
            ToolTipText     =   "取前后两个竖曲线之间的直线上的任意一点里程"
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command21 
            Caption         =   "保存到文件"
            Height          =   375
            Left            =   6240
            TabIndex        =   125
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command22 
            Caption         =   "从文件载入"
            Height          =   375
            Left            =   4680
            TabIndex        =   114
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Text33 
            Height          =   300
            Left            =   1200
            TabIndex        =   115
            ToolTipText     =   "分段总数"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text34 
            Height          =   300
            Left            =   3216
            TabIndex        =   116
            ToolTipText     =   "Kqd"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text42 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4680
            TabIndex        =   117
            ToolTipText     =   "横坡参数文件"
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "右幅横坡"
            Height          =   255
            Index           =   86
            Left            =   4950
            TabIndex        =   296
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "左幅横坡"
            Height          =   255
            Index           =   85
            Left            =   3480
            TabIndex        =   295
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "分段序号"
            Height          =   375
            Index           =   83
            Left            =   480
            TabIndex        =   294
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "变坡点桩号Ki"
            Height          =   255
            Index           =   84
            Left            =   1920
            TabIndex        =   293
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "曲线段数目："
            Height          =   255
            Index           =   81
            Left            =   120
            TabIndex        =   292
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "起点里程："
            Height          =   255
            Index           =   82
            Left            =   2280
            TabIndex        =   291
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "已知竖曲线要素(坡度EF不带百分号)"
         Height          =   1935
         Index           =   6
         Left            =   -74880
         TabIndex        =   281
         Top             =   360
         Width           =   7815
         Begin VB.CommandButton Command69 
            Caption         =   "9860"
            Height          =   375
            Left            =   6960
            TabIndex        =   113
            ToolTipText     =   "生成卡西欧5800计算器主程序"
            Top             =   1320
            Width           =   735
         End
         Begin VB.CommandButton Command36 
            Caption         =   "检验"
            Height          =   375
            Left            =   3060
            TabIndex        =   109
            Top             =   1320
            Width           =   615
         End
         Begin VB.CommandButton Command48 
            Caption         =   "5800"
            Height          =   375
            Left            =   6180
            TabIndex        =   112
            ToolTipText     =   "生成卡西欧5800计算器主程序"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox Text24 
            Height          =   300
            Left            =   1200
            TabIndex        =   99
            ToolTipText     =   "分段总数必须大于3"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            Caption         =   "载入文件"
            Height          =   375
            Left            =   3780
            TabIndex        =   110
            Top             =   1320
            Width           =   1035
         End
         Begin VB.CommandButton Command12 
            Caption         =   "保存到文件"
            Height          =   375
            Left            =   4920
            TabIndex        =   111
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Text25 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   1728
            TabIndex        =   102
            ToolTipText     =   "取前后两个竖曲线之间的直线上的任意一点里程"
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox Combo3 
            Height          =   300
            Left            =   240
            TabIndex        =   101
            ToolTipText     =   "第几个分段"
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton Command19 
            Caption         =   "输入"
            Height          =   375
            Left            =   2400
            TabIndex        =   108
            ToolTipText     =   "注意各段的参数修改后一定要点输入按钮并注意存盘"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox Text26 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   1320
            TabIndex        =   107
            ToolTipText     =   "F不带百分号,上坡取正,下坡取负"
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox TEXTE 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   240
            TabIndex        =   106
            ToolTipText     =   "E不带百分号,上坡取正,下坡取负"
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox Text27 
            Height          =   300
            Left            =   6195
            TabIndex        =   105
            ToolTipText     =   "R取绝对值,正值"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text28 
            Height          =   300
            Left            =   4704
            TabIndex        =   104
            ToolTipText     =   "Y取竖曲线顶点的高程"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text29 
            Height          =   300
            Left            =   3216
            TabIndex        =   103
            ToolTipText     =   "J取竖曲线顶点里程"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text41 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4680
            TabIndex        =   100
            ToolTipText     =   "竖曲线参数文件"
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "变坡点数目："
            Height          =   255
            Index           =   72
            Left            =   120
            TabIndex        =   289
            ToolTipText     =   "变坡点数目必须包含起点和终点"
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "分段终点Ki"
            Height          =   255
            Index           =   75
            Left            =   1920
            TabIndex        =   288
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "分段序号"
            Height          =   375
            Index           =   74
            Left            =   480
            TabIndex        =   287
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "第二坡度F(%)"
            Height          =   255
            Index           =   80
            Left            =   1260
            TabIndex        =   286
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "第一坡度E(%)"
            Height          =   255
            Index           =   79
            Left            =   120
            TabIndex        =   285
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "竖曲线半径R"
            Height          =   255
            Index           =   78
            Left            =   6360
            TabIndex        =   284
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "变坡点高程Y"
            Height          =   255
            Index           =   77
            Left            =   4800
            TabIndex        =   283
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "变坡点里程J"
            Height          =   255
            Index           =   76
            Left            =   3360
            TabIndex        =   282
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "所求点K参数及计算结果"
         Height          =   2535
         Index           =   12
         Left            =   -74880
         TabIndex        =   267
         Top             =   2640
         Width           =   7815
         Begin VB.CommandButton Command27 
            Caption         =   "导出坐标到Excel文件"
            Height          =   375
            Left            =   5400
            TabIndex        =   74
            Top             =   2000
            Width           =   2055
         End
         Begin VB.TextBox Text20 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   6195
            Locked          =   -1  'True
            TabIndex        =   71
            TabStop         =   0   'False
            ToolTipText     =   "Y"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Text19 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   4704
            Locked          =   -1  'True
            TabIndex        =   70
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command18 
            Caption         =   "导出坐标到txt文件"
            Height          =   375
            Left            =   3240
            TabIndex        =   73
            Top             =   2000
            Width           =   1935
         End
         Begin VB.CommandButton Command17 
            Caption         =   "计算"
            Height          =   375
            Left            =   6360
            TabIndex        =   66
            ToolTipText     =   "只计算Kstart"
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton Command16 
            Caption         =   "生成中桩坐标的CAD脚本文件"
            Height          =   375
            Left            =   240
            TabIndex        =   72
            Top             =   2000
            Width           =   2775
         End
         Begin VB.TextBox Text18 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Text17 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   1728
            Locked          =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            ToolTipText     =   "Y"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Text16 
            Height          =   300
            Left            =   3216
            TabIndex        =   69
            TabStop         =   0   'False
            ToolTipText     =   "度.分秒的格式,若改变则以新方位角为基准重新计算(G,H)"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Text15 
            Height          =   300
            Left            =   4704
            TabIndex        =   65
            ToolTipText     =   "Z"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text14 
            Height          =   300
            Left            =   4440
            TabIndex        =   60
            Text            =   "10"
            ToolTipText     =   "间距不能取零"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text13 
            Height          =   300
            Left            =   2520
            TabIndex        =   59
            ToolTipText     =   "Kend"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text12 
            Height          =   300
            Left            =   960
            TabIndex        =   58
            ToolTipText     =   "Kstart"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text11 
            Height          =   300
            Left            =   3216
            TabIndex        =   64
            ToolTipText     =   "S"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text10 
            Height          =   300
            Left            =   1728
            TabIndex        =   63
            ToolTipText     =   "P"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text9 
            Height          =   300
            Left            =   240
            TabIndex        =   62
            ToolTipText     =   "T"
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command15 
            Caption         =   "查看可计算的范围"
            Height          =   375
            Left            =   5520
            TabIndex        =   61
            ToolTipText     =   "必须先输入A和R值"
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "坐标H"
            Height          =   255
            Index           =   44
            Left            =   6600
            TabIndex        =   280
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "坐标G"
            Height          =   255
            Index           =   43
            Left            =   5040
            TabIndex        =   279
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "中桩X坐标"
            Height          =   255
            Index           =   40
            Left            =   480
            TabIndex        =   278
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "中桩Y坐标"
            Height          =   255
            Index           =   41
            Left            =   1920
            TabIndex        =   277
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "切线方位角Q"
            Height          =   255
            Index           =   42
            Left            =   3360
            TabIndex        =   276
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "第二偏角Z"
            Height          =   255
            Index           =   39
            Left            =   4920
            TabIndex        =   275
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "m"
            Height          =   255
            Index           =   35
            Left            =   5200
            TabIndex        =   274
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "间隔："
            Height          =   375
            Index           =   34
            Left            =   3960
            TabIndex        =   273
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "～"
            Height          =   255
            Index           =   33
            Left            =   2280
            TabIndex        =   272
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "里程K="
            Height          =   375
            Index           =   32
            Left            =   360
            TabIndex        =   271
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "第二偏距S"
            Height          =   255
            Index           =   38
            Left            =   3480
            TabIndex        =   270
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "第一偏角P"
            Height          =   255
            Index           =   37
            Left            =   2040
            TabIndex        =   269
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "第一偏距T"
            Height          =   255
            Index           =   36
            Left            =   480
            TabIndex        =   268
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "所求点K参数及计算结果"
         Height          =   2535
         Index           =   13
         Left            =   120
         TabIndex        =   253
         Top             =   3240
         Width           =   7815
         Begin VB.TextBox Text116 
            Height          =   375
            Left            =   2400
            TabIndex        =   37
            Text            =   "7"
            ToolTipText     =   "1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引1-249"
            Top             =   2000
            Width           =   495
         End
         Begin VB.CommandButton Command68 
            Caption         =   "反算"
            Height          =   375
            Left            =   7160
            TabIndex        =   30
            Top             =   840
            Width           =   550
         End
         Begin VB.TextBox Text114 
            Height          =   375
            Left            =   3840
            TabIndex        =   38
            Text            =   "3"
            ToolTipText     =   "保留的小数位"
            Top             =   2000
            Width           =   375
         End
         Begin VB.CommandButton Command40 
            Caption         =   "生成中桩CAD脚本"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   2000
            Width           =   1695
         End
         Begin VB.CommandButton Command26 
            Caption         =   "导出坐标到Excel"
            Height          =   375
            Left            =   6120
            TabIndex        =   40
            Top             =   2000
            Width           =   1575
         End
         Begin VB.TextBox TEXTT 
            Height          =   300
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "T"
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox TEXTP 
            Height          =   300
            Left            =   1605
            TabIndex        =   26
            ToolTipText     =   "P度.分秒的格式"
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox TEXTS 
            Height          =   300
            Left            =   3090
            TabIndex        =   27
            ToolTipText     =   "S"
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox TEXTKstart 
            Height          =   300
            Left            =   960
            TabIndex        =   21
            ToolTipText     =   "Kstart"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TEXTKend 
            Height          =   300
            Left            =   2520
            TabIndex        =   22
            ToolTipText     =   "Kend"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TEXToffset 
            Height          =   300
            Left            =   4440
            TabIndex        =   23
            Text            =   "20"
            ToolTipText     =   "间距不能取零"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TEXTZ 
            Height          =   300
            Left            =   4590
            TabIndex        =   28
            ToolTipText     =   "Z度.分秒的格式"
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox TEXTQ 
            Height          =   300
            Left            =   3090
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "度.分秒的格式,若改变则以新方位角为基准重新计算(G,H)"
            Top             =   1500
            Width           =   1335
         End
         Begin VB.TextBox TEXTY 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   1605
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Y"
            Top             =   1500
            Width           =   1335
         End
         Begin VB.TextBox TEXTX 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   1500
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   "计算起点"
            Height          =   375
            Left            =   6120
            TabIndex        =   29
            ToolTipText     =   "只计算Kstart"
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox TEXTG 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   4590
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "X"
            Top             =   1500
            Width           =   1335
         End
         Begin VB.TextBox TEXTH 
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   6075
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Y"
            Top             =   1500
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            Caption         =   "导出坐标到txt文本"
            Height          =   375
            Left            =   4320
            TabIndex        =   39
            Top             =   2000
            Width           =   1695
         End
         Begin VB.CommandButton Command11 
            Caption         =   "添加到全站仪的CSV文件"
            Height          =   375
            Left            =   5400
            TabIndex        =   24
            ToolTipText     =   "适用于PENTAX R-322N型的全站仪,该格式可以直接通过连线输入全站仪,若间距为零则只计算Kstart,可选择同一个文件进行添加"
            Top             =   200
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "颜色："
            Height          =   255
            Index           =   139
            Left            =   1920
            TabIndex        =   413
            Top             =   2055
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "小数位："
            Height          =   255
            Index           =   138
            Left            =   3120
            TabIndex        =   412
            Top             =   2060
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "第一偏距T"
            Height          =   255
            Index           =   17
            Left            =   360
            TabIndex        =   266
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "第一偏角P"
            Height          =   255
            Index           =   18
            Left            =   1800
            TabIndex        =   265
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "第二偏距S"
            Height          =   255
            Index           =   19
            Left            =   3360
            TabIndex        =   264
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "里程K="
            Height          =   375
            Index           =   13
            Left            =   360
            TabIndex        =   263
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "～"
            Height          =   255
            Index           =   14
            Left            =   2280
            TabIndex        =   262
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "间隔："
            Height          =   255
            Index           =   15
            Left            =   3960
            TabIndex        =   261
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "m"
            Height          =   255
            Index           =   16
            Left            =   5200
            TabIndex        =   260
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "第二偏角Z"
            Height          =   255
            Index           =   20
            Left            =   4800
            TabIndex        =   259
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "切线方位角Q"
            Height          =   255
            Index           =   23
            Left            =   3240
            TabIndex        =   258
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "中桩Y坐标"
            Height          =   255
            Index           =   22
            Left            =   1920
            TabIndex        =   257
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "中桩X坐标"
            Height          =   255
            Index           =   21
            Left            =   360
            TabIndex        =   256
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "坐标G"
            Height          =   255
            Index           =   24
            Left            =   4920
            TabIndex        =   255
            Top             =   1260
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "坐标H"
            Height          =   255
            Index           =   25
            Left            =   6480
            TabIndex        =   254
            Top             =   1260
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "已知平曲线要素"
         Height          =   2775
         Index           =   0
         Left            =   120
         TabIndex        =   239
         Top             =   360
         Width           =   7815
         Begin VB.CommandButton Command38 
            Caption         =   "9860"
            Height          =   375
            Left            =   7080
            TabIndex        =   20
            ToolTipText     =   "生成卡西欧5800计算器主程序"
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton Command56 
            Caption         =   "检验Ki"
            Height          =   375
            Left            =   6940
            TabIndex        =   16
            ToolTipText     =   "检验分段终点是否有误"
            Top             =   2220
            Width           =   720
         End
         Begin VB.CommandButton Command47 
            Caption         =   "5800"
            Height          =   375
            Left            =   6420
            TabIndex        =   19
            ToolTipText     =   "生成卡西欧5800计算器主程序"
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton Command35 
            Caption         =   "从圆曲线推算HY点"
            Height          =   375
            Left            =   2460
            TabIndex        =   17
            ToolTipText     =   "把特征点计算页面中ZH点的XYQ输入"
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4560
            TabIndex        =   3
            ToolTipText     =   "平曲线参数文件"
            Top             =   240
            Width           =   3135
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Left            =   900
            TabIndex        =   5
            ToolTipText     =   "各分段的类型"
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox TextQD 
            Height          =   300
            Left            =   3120
            TabIndex        =   2
            ToolTipText     =   "KQD"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TextPartNumber 
            Height          =   300
            Left            =   1200
            TabIndex        =   1
            ToolTipText     =   "分段总数"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "载入文件"
            Height          =   375
            Left            =   4200
            TabIndex        =   0
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "保存到文件"
            Height          =   375
            Left            =   5220
            TabIndex        =   18
            Top             =   840
            Width           =   1155
         End
         Begin VB.TextBox TEXTR 
            Height          =   300
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "R不能取零,左偏取负,右偏取正"
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox TEXTA 
            Height          =   300
            Left            =   1605
            TabIndex        =   12
            ToolTipText     =   "注意A不能取零"
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox TEXTB 
            Height          =   300
            Left            =   3090
            TabIndex        =   13
            ToolTipText     =   "注意B不能取零"
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox TEXTO 
            Height          =   300
            Left            =   4590
            TabIndex        =   14
            ToolTipText     =   "O"
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "第几个分段"
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton Command9 
            Caption         =   "输入"
            Height          =   375
            Left            =   6240
            TabIndex        =   15
            ToolTipText     =   "注意各段的参数修改后一定要点输入按钮并注意存盘"
            Top             =   2220
            Width           =   615
         End
         Begin VB.TextBox TEXTF 
            Height          =   300
            Left            =   6075
            TabIndex        =   10
            ToolTipText     =   "必须为度.分秒的格式"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox TEXTD 
            Height          =   300
            Left            =   4590
            TabIndex        =   9
            ToolTipText     =   "D"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox TEXTC 
            Height          =   300
            Left            =   3090
            TabIndex        =   8
            ToolTipText     =   "C"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox TEXTJ 
            Height          =   300
            Left            =   1605
            TabIndex        =   7
            ToolTipText     =   "J"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox TEXTKI 
            Height          =   300
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Ki"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "类型"
            Height          =   375
            Index           =   3
            Left            =   1260
            TabIndex        =   252
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "起点里程："
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   251
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "曲线段数目："
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   250
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label LabelR 
            Caption         =   "圆曲线半径R"
            Height          =   255
            Left            =   240
            TabIndex        =   249
            Top             =   2040
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label LabelA 
            Caption         =   "第一回旋参数A"
            Height          =   255
            Left            =   1680
            TabIndex        =   248
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label LabelB 
            Caption         =   "第二回旋参数B"
            Height          =   255
            Left            =   3150
            TabIndex        =   247
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label LabelO 
            Caption         =   "圆曲线长O"
            Height          =   255
            Left            =   4560
            TabIndex        =   246
            Top             =   2040
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "分段序号"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   245
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "起算点切线方位角F"
            Height          =   255
            Index           =   8
            Left            =   5925
            TabIndex        =   244
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "起算点Y坐标D"
            Height          =   255
            Index           =   7
            Left            =   4560
            TabIndex        =   243
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "起算点X坐标C"
            Height          =   255
            Index           =   6
            Left            =   3120
            TabIndex        =   242
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "起算点里程J"
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   241
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "分段终点Ki"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   240
            Top             =   1320
            Width           =   1095
         End
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退  出"
      Height          =   420
      Left            =   6720
      TabIndex        =   237
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "更新记录"
      Height          =   420
      Left            =   240
      TabIndex        =   234
      Top             =   7560
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
'《高速公路测量伴侣》双心软件系列------坐标高程计算程序
Dim factor3(1000, 9) As Variant                                                 '在通用部分中申明平曲线参数数组,通过交点法计算得到
Dim factor2(-5 To 1000, 9) As Variant                                           '在通用部分中申明交点参数数组：Kjd，X，Y，R，Ls，T，L
Dim factor4(5000, 5) As Variant                                                 '全局数组，分别存储平曲线：点号,K,T,P,S,Z
Dim shu(1000, 5) As Variant                                                     '在通用部分中申明参数数组竖曲线参数
Dim hengpo(1000, 3) As Variant                                                  '存储横坡参数的数组KI,左，右，渐变类型
Dim kuandu(1000, 2) As Variant                                                  '存储左幅宽度参数的数组Ki，左，渐变类型
Dim kuandu2(1000, 2) As Variant                                                 '存储右幅宽度参数的数组Ki，右，渐变类型
Dim Partnumber, number, Saveflag As Integer                                     '曲线要素存盘标志：SaveFlag=1已存盘，Save=0未存盘
Dim K, J, C, D, F, R, A, B, O, x, y, Q, T, P, S, Z, U, V, G, H, dL As Variant   '默认全部为变体
Dim Kqd, Kstart, Kend, offset As Variant                                        '起点里程
Dim partnumber1 As Integer, partnumber2 As Integer
Dim E, i As Variant
Dim QB As Variant                                                               '桥梁涵洞平行布置基准切线方位角
Dim ndianhao As Variant                                                         '把文件中所求点数目保存在全局变量中，导出到cad或excel按钮中要用到
Dim ndianhao2 As Variant                                                        '竖曲线所求点总数
Dim ndh As Variant                                                              '把文件中所求点数目保存在全局变量中，导出到cad或excel按钮中要用到
Dim ndh2 As Variant                                                             '竖曲线所求点序号
Dim anniu As Variant                                                            '如果按下了生成cad脚本按钮就把anniu全局变量赋值为1，在ceyi子程序里通过判断该变量给TPSZ赋值。
Dim anniu2 As Variant
Dim banben As Variant                                                           '交点法版本0=新版本，1=老版本
'定义一下常量
Const pai = 3.14159265358979
Const xlCenter = -4108
Const xlNone = -4142
Const xlAutomatic = -4105
Const xlDiagonalDown = 5
Const xlDiagonalUp = 6
Const xlEdgeLeft = 7
Const xlEdgeTop = 8
Const xlEdgeBottom = 9
Const xlEdgeRight = 10
Const xlInsideVertical = 11
Const xlInsideHorizontal = 12
Const xlContinuous = 1
Const xlMedium = -4138
Const xlThick = 4
Const xlthin = 2
Const xlPrintNoComments = -4142
Const xlPortrait = 1
Const xlPaperA4 = 9
Const xlDownThenOver = 1
Const xlPrintErrorsDisplayed = 0
Const xlCalculationManual = -4135
Const xlFillDefault = 0
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Check1_Click()
    If Check1.Value = 0 Then                                                    '如果没有被选中就是完全缓和曲线
        Label7(1).Visible = False
        LabelT(1).Visible = False
        Text109.Visible = False
        Text108.Visible = False
    Else                                                                        '如果选中了就是不完全缓和曲线
        Label7(1).Visible = True
        LabelT(1).Visible = True
        Text109.Visible = True
        Text108.Visible = True
    End If
End Sub

Private Sub Command61_Click()
    '载入参数并生成横断面CAD脚本
    On Error GoTo Label1
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "从文件载入设计线参数..."
        .DefaultExt = "sjx"
        .Filter = "设计线参数(*.sjx)|*.sjx|地面线参数(*.dmx)|*.dmx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text109 = .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 17
    Line Input #17, shuoming1
    Line Input #17, shuoming2
    Line Input #17, shuoming3
    
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "从文件载入地面线参数..."
        .DefaultExt = "dmx"
        .Filter = "地面线参数(*.dmx)|*.dmx|设计线参数(*.sjx)|*.sjx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text109 = Text109 & "  " & .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 18
    Line Input #18, shuoming1
    Line Input #18, shuoming2
    Line Input #18, shuoming3
    
    
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "生成横断面的CAD脚本..."
        .DefaultExt = "scr"
        .Filter = "横断面CAD脚本(*.scr)|*.scr"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
    End With
    Open CommonDialog1.FileName For Output As 19
    
    '设置坐标输入的优先级为键盘输入优先
    Print #19, "osnapcoord" & vbCrLf & "1"
    
    '加载线型center
    Print #19, "-linetype" & vbCrLf & "L" & vbCrLf & "center" & vbCrLf & "s" & vbCrLf
    '新建图层-中线，线型center长短划
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "中线" & vbCrLf & "L" & vbCrLf & "center" & vbCrLf & "中线" & vbCrLf
    '新建图层-设计线'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "设计线" & vbCrLf & "c" & vbCrLf & "1" & vbCrLf & "设计线" & vbCrLf
    '新建图层-设计线标注'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "设计线标注" & vbCrLf & "c" & vbCrLf & "3" & vbCrLf & "设计线标注" & vbCrLf
    '新建图层-坡脚线标注'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "坡脚线标注" & vbCrLf & "c" & vbCrLf & "3" & vbCrLf & "坡脚线标注" & vbCrLf
    '新建图层-地面线'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "地面线" & vbCrLf & "c" & vbCrLf & "7" & vbCrLf & "地面线" & vbCrLf
    '新建图层-地面线标高'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "地面线标高" & vbCrLf & "c" & vbCrLf & "7" & vbCrLf & "地面线标高" & vbCrLf
    '新建图层-地面线平距'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "地面线平距" & vbCrLf & "c" & vbCrLf & "11" & vbCrLf & "地面线平距" & vbCrLf
    '新建图层-中线标注'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "中线标注" & vbCrLf & "c" & vbCrLf & "7" & vbCrLf & "中线标注" & vbCrLf
    Height1 = Val(Text112)                                                      '设置文字高度,默认0.4
    length1 = Height1 * 9.5                                                     '设置引线长度
    length2 = Height1 * 1.1                                                     '设置引线偏移距离
    
    '设置文字样式
    '标注设计线距离高程的文字样式：
    Print #19, "-style"                                                         '建立文字样式
    Print #19, "标高字体"                                                       '文字样式名称
    Print #19, "simfang.ttf"                                                    '字体名称
    Write #19, Height1                                                          '文字高度
    Print #19, "1"                                                              '宽度比例
    Print #19, "0"                                                              '倾斜角度
    Print #19, "n"                                                              '反向显示文字=否
    Print #19, "n"                                                              '颠倒显示文字=否
    
    '标注里程的文字样式：
    Print #19, "-style"                                                         '建立文字样式
    Print #19, "里程字体"                                                       '文字样式名称
    Print #19, "simfang.ttf"                                                    '字体名称
    Write #19, Height1 * 1.5                                                    '文字高度
    Print #19, "1"                                                              '宽度比例
    Print #19, "0"                                                              '倾斜角度
    Print #19, "n"                                                              '反向显示文字=否
    Print #19, "n"                                                              '颠倒显示文字=否
    
    Dim str1 As String
    Dim str2 As String
    Dim sjx() As String
    Dim dmx() As String
    dy = 0
    Line Input #17, str2                                                        '首先读入设计线的第一个里程文件准备比较
    sjx = Split(str2, ",")
    Do While EOF(18) = False                                                    '读地面线文件到结束且不为空值
        Line Input #18, str1                                                    '读地面线
        dmx = Split(str1, ",")
        If UBound(dmx) = -1 Then Exit Do                                        '如果ubound(dmx)=-1,直接引用dmx(0)会造成数组下标越界，必须退出循环
        
        Do While Not EOF(17) And UBound(sjx) <> -1                              '读设计线到结束且不为空值
            If Val(sjx(0)) < Val(dmx(0)) Then                                   '如果设计线里程小就继续读设计线
                Line Input #17, str2
                sjx = Split(str2, ",")
            Else
                Exit Do                                                         '如果设计里程大就退出循环
            End If
        Loop
        
        If UBound(sjx) >= 0 And UBound(dmx) >= 0 Then
            If CDec(Val(sjx(0))) = CDec(Val(dmx(0))) Then                       '如果在设计线中找到对应的桩号就计算
                '画中线：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "中线" & vbCrLf   '设置当前图层为中线
                Print #19, "line" & vbCrLf & 0 & "," & Val(sjx(1)) + Height1 * 3 + dy
                Print #19, 0 & "," & Val(dmx(2)) - Height1 * 2 + dy & vbCrLf
                
                '画设计线：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "设计线" & vbCrLf '设置当前图层为设计线层
                Print #19, "line"
                '求左幅坡脚线
                For I1 = 2 To UBound(sjx) - 3 Step 2
                    If UBound(sjx) >= 4 Then
                        If UBound(sjx) Mod 2 = 1 Then                           '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                            Xa = Val(sjx(I1))
                            Ya = Val(sjx(1)) + Val(sjx(I1 + 1))
                            Xb = Val(sjx(I1 + 2))
                            Yb = Val(sjx(1)) + Val(sjx(I1 + 3))
                        Else                                                    '如果是里程1，平距1，高程1，平距2，高程2格式
                            Xa = Val(sjx(I1 - 1))
                            Ya = Val(sjx(I1))
                            Xb = Val(sjx(I1 + 1))
                            Yb = Val(sjx(I1 + 2))
                        End If
                        Xc = Val(dmx(1))
                        Yc = Val(dmx(2))
                        If UBound(dmx) >= 4 Then
                            Xd = Val(dmx(3))
                            Yd = Val(dmx(4))
                        Else
                            Xd = Xc - 13
                            Yd = Yc
                        End If
                        Call 计算交点(Xa, Ya, Xb, Yb, Xc, Yc, Xd, Yd, tt1, ss1, Xjd1, Yjd1) '计算左幅交点
                        If tt1 >= 0 And tt1 < 1 Then                            '要确保设计线足够长才有交点
                            Print #19, Xjd1 & "," & Yjd1 + dy
                            Exit For
                        End If
                    End If
                Next
                '画中间的特征点：
                For I1 = 5 To UBound(sjx) - 1 Step 2
                    If UBound(sjx) Mod 2 = 1 Then                               '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                        Print #19, Val(sjx(I1 - 1)) & "," & Val(sjx(1)) + Val(sjx(I1)) + dy
                    Else                                                        '如果是里程1，平距1，高程1，平距2，高程2格式
                        Print #19, Val(sjx(I1 - 2)) & "," & Val(sjx(I1 - 1)) + dy
                    End If
                Next
                '求右幅坡脚线
                For I1 = UBound(sjx) - 1 To 4 Step -2
                    If UBound(sjx) >= 5 Then
                        If UBound(sjx) Mod 2 = 1 Then                           '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                            Xa = Val(sjx(I1 - 2))
                            Ya = Val(sjx(1)) + Val(sjx(I1 - 1))
                            Xb = Val(sjx(I1))
                            Yb = Val(sjx(1)) + Val(sjx(I1 + 1))
                        Else                                                    '如果是里程1，平距1，高程1，平距2，高程2格式
                            Xa = Val(sjx(I1 - 2))
                            Ya = Val(sjx(I1 - 1))
                            Xb = Val(sjx(I1))
                            Yb = Val(sjx(I1 + 1))
                        End If
                        i2 = UBound(dmx)
                        If UBound(dmx) >= 4 Then
                            Xc = Val(dmx(i2 - 3))
                            Yc = Val(dmx(i2 - 2))
                        Else
                            Xc = Val(dmx(1)) + 13
                            Yc = Val(dmx(2))
                        End If
                        Xd = Val(dmx(i2 - 1))
                        Yd = Val(dmx(i2))
                        Call 计算交点(Xa, Ya, Xb, Yb, Xc, Yc, Xd, Yd, tt1, ss1, xjd2, Yjd2) '计算右幅交点
                        If tt1 >= 0 And tt1 < 1 Then                            '要确保设计线足够长才有交点
                            Print #19, xjd2 & "," & Yjd2 + dy
                            Exit For
                        End If
                    End If
                Next
                Print #19, ""                                                   '结束画设计线命令
                
                '标注设计线距离高程的文字样式：
                Print #19, "textstyle"                                          '建立文字样式
                Print #19, "标高字体"                                           '文字样式名称
                
                
                '标注坡脚线：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "坡脚线标注" & vbCrLf '设置当前图层为坡脚线标注
                Print #19, "line" & vbCrLf & Xjd1 & "," & Yjd1 + dy & vbCrLf _
                & "@" & length1 & "<90" & vbCrLf                                '画左幅坡脚引线
                Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                Xjd1 - Height1 / 5 & "," & Yjd1 + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                Print #19, "距离:" & Round(Xjd1, 3)                             '标注左幅坡脚距离
                Print #19, "text" & vbCrLf & "j" & vbCrLf & "tl" & vbCrLf & _
                Xjd1 + Height1 / 2 & "," & Yjd1 + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                Print #19, "高程:" & Round(Yjd1, 3)                             '标注左幅坡脚高程
                
                
                Print #19, "line" & vbCrLf & xjd2 & "," & Yjd2 + dy & vbCrLf _
                & "@" & length1 & "<90" & vbCrLf                                '画右幅坡脚引线
                Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                xjd2 - Height1 / 5 & "," & Yjd2 + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                Print #19, "距离:" & Round(xjd2, 3)                             '标注右幅坡脚距离
                Print #19, "text" & vbCrLf & "j" & vbCrLf & "tl" & vbCrLf & _
                xjd2 + Height1 / 2 & "," & Yjd2 + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                Print #19, "高程:" & Round(Yjd2, 3)                             '标注右幅坡脚高程
                
                '标注设计特征点：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "设计线标注" & vbCrLf '设置当前图层为设计线标注
                For I1 = 5 To UBound(sjx) - 1 Step 2
                    If UBound(sjx) Mod 2 = 1 Then                               '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                        Print #19, "line" & vbCrLf & Val(sjx(I1 - 1)) & "," & Val(sjx(1)) + Val(sjx(I1)) + dy & vbCrLf _
                        & "@" & length1 & "<90" & vbCrLf                        '画引线
                        '标注距离：
                        Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                        Val(sjx(I1 - 1)) - Height1 / 5 & "," & Val(sjx(1)) + Val(sjx(I1)) + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                        Print #19, "距离:" & Val(sjx(I1 - 1))
                        '标注高程：
                        Print #19, "text" & vbCrLf & "j" & vbCrLf & "tl" & vbCrLf & _
                        Val(sjx(I1 - 1)) + Height1 / 2 & "," & Val(sjx(1)) + Val(sjx(I1)) + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                        Print #19, "高程:" & Val(sjx(1)) + Val(sjx(I1))
                    Else                                                        '如果是里程1，平距1，高程1，平距2，高程2格式
                        Print #19, "line" & vbCrLf & Val(sjx(I1 - 2)) & "," & Val(sjx(I1 - 1)) + dy & vbCrLf _
                        & "@" & length1 & "<90" & vbCrLf                        '画引线
                        '标注距离：
                        Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                        Val(sjx(I1 - 2)) - Height1 / 5 & "," & Val(sjx(I1 - 1)) + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                        Print #19, "距离:" & Val(sjx(I1 - 2))
                        '标注高程：
                        Print #19, "text" & vbCrLf & "j" & vbCrLf & "tl" & vbCrLf & _
                        Val(sjx(I1 - 2)) + Height1 / 2 & "," & Val(sjx(I1 - 1)) + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                        Print #19, "高程:" & Val(sjx(I1 - 1))
                    End If
                Next
                
                '画地面线：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "地面线" & vbCrLf '设置当前图层为地面线层
                Print #19, "line"
                Print #19, Xjd1 & "," & Yjd1 + dy
                For I1 = 2 To UBound(dmx) Step 2
                    Print #19, Val(dmx(I1 - 1)) & "," & Val(dmx(I1)) + dy
                Next
                Print #19, xjd2 & "," & Yjd2 + dy & vbCrLf
                '画引线：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "地面线标高" & vbCrLf '设置当前图层为地面线标高
                For I1 = 2 To UBound(dmx) Step 2
                    Print #19, "line" & vbCrLf & Val(dmx(I1 - 1)) + Height1 * 0.5773 & "," & Val(dmx(I1)) + Height1 + dy & vbCrLf & _
                    Val(dmx(I1 - 1)) & "," & Val(dmx(I1)) + dy & vbCrLf & _
                    Val(dmx(I1 - 1)) - Height1 * 0.5773 & "," & Val(dmx(I1)) + Height1 + dy & vbCrLf & _
                    Val(dmx(I1 - 1)) + Height1 * 2 & "," & Val(dmx(I1)) + Height1 + dy & vbCrLf
                    '标注高程：
                    Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                    Val(dmx(I1 - 1)) - Height1 * 1.4 & "," & Val(dmx(I1)) + Height1 * 1.1 + dy & vbCrLf & "0" '文字起点位置和标注方向
                    Print #19, Val(dmx(I1))
                Next
                '标注平距：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "地面线平距" & vbCrLf '设置当前图层为地面线平距
                For I1 = 2 To UBound(dmx) Step 2
                    Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                    Val(dmx(I1 - 1)) - Height1 * 1.4 & "," & Val(dmx(I1)) + Height1 * 2.4 + dy & vbCrLf & "0" '文字起点位置和标注方向
                    Print #19, Round(Val(dmx(I1 - 1)), 3)
                Next
                
                '求地面中桩高程
                If UBound(dmx) = 2 Then
                    hd = Round(Val(dmx(2)), 3)
                ElseIf UBound(dmx) > 2 Then
                    For I1 = 2 To UBound(dmx) - 2 Step 2                        '求中桩高程
                        If dmx(I1 - 1) = 0 Or dmx(I1 - 1) < 0 And dmx(I1 + 1) > 0 Then
                            hd = Val(dmx(I1)) + Abs(Val(dmx(I1 - 1))) * (Val(dmx(I1 + 2)) - Val(dmx(I1))) / (Abs(Val(dmx(I1 - 1))) + Val(dmx(I1 + 1)))
                            hd = Round(hd, 3)
                            Exit For
                        End If
                    Next
                End If
                
                '计算填方段或挖方段面积
                smax = UBound(sjx)
                dmax = UBound(dmx)
                Dim aa(200)                                                     '先把多边形顶点xy存放到aa中，有I1个顶点
                aa(0) = Xjd1
                aa(1) = Yjd1
                For I1 = 4 To UBound(sjx) - 2 Step 2
                    If UBound(sjx) Mod 2 = 1 Then                               '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                        aa(I1 - 2) = Val(sjx(I1))
                        aa(I1 - 1) = Val(sjx(1)) + Val(sjx(I1 + 1))
                    Else                                                        '如果是里程1，平距1，高程1，平距2，高程2格式
                        aa(I1 - 2) = Val(sjx(I1 - 1))
                        aa(I1 - 1) = Val(sjx(I1))
                    End If
                Next
                aa(I1 - 2) = xjd2
                aa(I1 - 1) = Yjd2
                For i2 = UBound(dmx) - 1 To 1 Step -2
                    If Val(dmx(i2)) > Xjd1 And Val(dmx(i2)) < xjd2 Then
                        aa(I1) = Val(dmx(i2))
                        aa(I1 + 1) = Val(dmx(i2 + 1))
                        I1 = I1 + 2
                    End If
                Next
                aa(I1) = Xjd1
                aa(I1 + 1) = Yjd1
                '已知坐标求多边形面积：不能计算半填半挖段。
                ss = 0
                For i2 = 0 To I1 - 2 Step 2
                    ss = ss + aa(i2) * aa(i2 + 3) - aa(i2 + 1) * aa(i2 + 2)
                Next
                ss = Round(Abs(ss) / 2, 3)
                
                '标注里程、设计标高、断面面积、填挖高
                '标注里程的文字样式：
                Print #19, "textstyle"                                          '建立文字样式
                Print #19, "里程字体"                                           '文字样式名称
                
                '计算设计高程
                If UBound(sjx) Mod 2 = 1 Then                                   '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                    Hs = Val(sjx(1))                                            '设计高程
                Else                                                            '如果是里程1,平距1，高程1，平距2，高程2格式
                    For i4 = 1 To UBound(sjx) - 2
                        If Val(sjx(i4)) <= 0 And Val(sjx(i4 + 2)) > 0 Then
                            Hs = Val(sjx(i4 + 1))                               '设计高程
                            Exit For
                        End If
                    Next
                End If
                
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "中线标注" & vbCrLf '设置当前图层为中线标注
                Print #19, "text" & vbCrLf & "j" & vbCrLf & "mc" & vbCrLf & _
0 & "," & Val(dmx(2)) - Height1 * 4 + dy & vbCrLf & "0"                         '文字起点位置和标注方向
                Print #19, Format(Val(sjx(0)), "\" & Text111) & "  Hs=" & Hs & "  Hd=" & hd & _
                "  Dh=" & Round(Hs - hd, 3) & "  AA=" & ss
                
                '向上偏移dy：
                If Val(Text113) = 0 Then
                    dy = dy + Val(sjx(1)) - Val(dmx(2)) + length1 * 2
                Else
                    dy = dy + Val(Text113)
                End If
                
            End If
        End If
    Loop
    
    '设置当前图层为0层
    Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "0" & vbCrLf                  '设置当前图层为0层
    Print #19, "zoom" & vbCrLf & "e"
    
    Close (17)
    Close (18)
    Close (19)
    msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
    If MsgBox(msg, vbYesNo) = vbYes Then
        Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
    End If
Label1: Close
End Sub

Private Sub Command65_Click()
    '载入参数并生成横断面CAD脚本
    On Error GoTo Label1
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "从文件载入设计线参数..."
        .DefaultExt = "sjx"
        .Filter = "设计线参数(*.sjx)|*.sjx|地面线参数(*.dmx)|*.dmx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text109 = .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 17
    Line Input #17, shuoming1
    Line Input #17, shuoming2
    Line Input #17, shuoming3
    
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "从文件载入地面线参数..."
        .DefaultExt = "dmx"
        .Filter = "地面线参数(*.dmx)|*.dmx|设计线参数(*.sjx)|*.sjx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text109 = Text109 & "  " & .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 18
    Line Input #18, shuoming1
    Line Input #18, shuoming2
    Line Input #18, shuoming3
    
    
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "生成横断面的CAD脚本..."
        .DefaultExt = "scr"
        .Filter = "横断面CAD脚本(*.scr)|*.scr"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
    End With
    Open CommonDialog1.FileName For Output As 19
    
    '设置坐标输入的优先级为键盘输入优先
    Print #19, "osnapcoord" & vbCrLf & "1"
    
    '加载线型center
    Print #19, "-linetype" & vbCrLf & "L" & vbCrLf & "center" & vbCrLf & "s" & vbCrLf
    '新建图层-中线，线型center长短划
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "中线" & vbCrLf & "L" & vbCrLf & "center" & vbCrLf & "中线" & vbCrLf
    '新建图层-设计线'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "设计线" & vbCrLf & "c" & vbCrLf & "1" & vbCrLf & "设计线" & vbCrLf
    '新建图层-设计线标注'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "设计线标注" & vbCrLf & "c" & vbCrLf & "3" & vbCrLf & "设计线标注" & vbCrLf
    '新建图层-坡脚线标注'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "坡脚线标注" & vbCrLf & "c" & vbCrLf & "3" & vbCrLf & "坡脚线标注" & vbCrLf
    '新建图层-地面线'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "地面线" & vbCrLf & "c" & vbCrLf & "7" & vbCrLf & "地面线" & vbCrLf
    '新建图层-地面线标高'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "地面线标高" & vbCrLf & "c" & vbCrLf & "7" & vbCrLf & "地面线标高" & vbCrLf
    '新建图层-地面线平距'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "地面线平距" & vbCrLf & "c" & vbCrLf & "11" & vbCrLf & "地面线平距" & vbCrLf
    '新建图层-中线标注'1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #19, "-layer" & vbCrLf & "m" & vbCrLf & "中线标注" & vbCrLf & "c" & vbCrLf & "7" & vbCrLf & "中线标注" & vbCrLf
    Height1 = Val(Text112)                                                      '设置文字高度,默认0.4
    length1 = Height1 * 9.5                                                     '设置引线长度
    length2 = Height1 * 1.1                                                     '设置引线偏移距离
    
    '设置文字样式
    '标注设计线距离高程的文字样式：
    Print #19, "-style"                                                         '建立文字样式
    Print #19, "标高字体"                                                       '文字样式名称
    Print #19, "宋体"                                                           '字体名称
    Write #19, Height1                                                          '文字高度
    Print #19, "1"                                                              '宽度比例
    Print #19, "0"                                                              '倾斜角度
    Print #19, "n"                                                              '反向显示文字=否
    Print #19, "n"                                                              '颠倒显示文字=否
    
    '标注里程的文字样式：
    Print #19, "-style"                                                         '建立文字样式
    Print #19, "里程字体"                                                       '文字样式名称
    Print #19, "宋体"                                                           '字体名称
    Write #19, Height1 * 1.5                                                    '文字高度
    Print #19, "1"                                                              '宽度比例
    Print #19, "0"                                                              '倾斜角度
    Print #19, "n"                                                              '反向显示文字=否
    Print #19, "n"                                                              '颠倒显示文字=否
    
    Dim str1 As String
    Dim str2 As String
    Dim sjx() As String
    Dim dmx() As String
    dy = 0
    Line Input #17, str2                                                        '首先读入设计线的第一个里程文件准备比较
    sjx = Split(str2, ",")
    Do While EOF(18) = False                                                    '读地面线文件到结束且不为空值
        Line Input #18, str1                                                    '读地面线
        dmx = Split(str1, ",")
        If UBound(dmx) = -1 Then Exit Do                                        '如果ubound(dmx)=-1,直接引用dmx(0)会造成数组下标越界，必须退出循环
        
        Do While Not EOF(17) And UBound(sjx) <> -1                              '读设计线到结束且不为空值
            If Val(sjx(0)) < Val(dmx(0)) Then                                   '如果设计线里程小就继续读设计线
                Line Input #17, str2
                sjx = Split(str2, ",")
            Else
                Exit Do                                                         '如果设计里程大就退出循环
            End If
        Loop
        
        If UBound(sjx) >= 0 And UBound(dmx) >= 0 Then
            If CDec(Val(sjx(0))) = CDec(Val(dmx(0))) Then                       '如果在设计线中找到对应的桩号就计算
                '画中线：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "中线" & vbCrLf   '设置当前图层为中线
                Print #19, "line" & vbCrLf & 0 & "," & Val(sjx(1)) + Height1 * 3 + dy
                Print #19, 0 & "," & Val(dmx(2)) - Height1 * 2 + dy & vbCrLf
                
                '画设计线：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "设计线" & vbCrLf '设置当前图层为设计线层
                Print #19, "line"
                '求左幅坡脚线
                For I1 = 2 To UBound(sjx) - 3 Step 2
                    If UBound(sjx) >= 4 Then
                        If UBound(sjx) Mod 2 = 1 Then                           '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                            Xa = Val(sjx(I1))
                            Ya = Val(sjx(1)) + Val(sjx(I1 + 1))
                            Xb = Val(sjx(I1 + 2))
                            Yb = Val(sjx(1)) + Val(sjx(I1 + 3))
                        Else                                                    '如果是里程1，平距1，高程1，平距2，高程2格式
                            Xa = Val(sjx(I1 - 1))
                            Ya = Val(sjx(I1))
                            Xb = Val(sjx(I1 + 1))
                            Yb = Val(sjx(I1 + 2))
                        End If
                        Xc = Val(dmx(1))
                        Yc = Val(dmx(2))
                        If UBound(dmx) >= 4 Then
                            Xd = Val(dmx(3))
                            Yd = Val(dmx(4))
                        Else
                            Xd = Xc - 13
                            Yd = Yc
                        End If
                        Call 计算交点(Xa, Ya, Xb, Yb, Xc, Yc, Xd, Yd, tt1, ss1, Xjd1, Yjd1) '计算左幅交点
                        If tt1 >= 0 And tt1 < 1 Then                            '要确保设计线足够长才有交点
                            Print #19, Xjd1 & "," & Yjd1 + dy
                            Exit For
                        End If
                    End If
                Next
                '画中间的特征点：
                For I1 = 5 To UBound(sjx) - 1 Step 2
                    If UBound(sjx) Mod 2 = 1 Then                               '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                        Print #19, Val(sjx(I1 - 1)) & "," & Val(sjx(1)) + Val(sjx(I1)) + dy
                    Else                                                        '如果是里程1，平距1，高程1，平距2，高程2格式
                        Print #19, Val(sjx(I1 - 2)) & "," & Val(sjx(I1 - 1)) + dy
                    End If
                Next
                '求右幅坡脚线
                For I1 = UBound(sjx) - 1 To 4 Step -2
                    If UBound(sjx) >= 5 Then
                        If UBound(sjx) Mod 2 = 1 Then                           '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                            Xa = Val(sjx(I1 - 2))
                            Ya = Val(sjx(1)) + Val(sjx(I1 - 1))
                            Xb = Val(sjx(I1))
                            Yb = Val(sjx(1)) + Val(sjx(I1 + 1))
                        Else                                                    '如果是里程1，平距1，高程1，平距2，高程2格式
                            Xa = Val(sjx(I1 - 2))
                            Ya = Val(sjx(I1 - 1))
                            Xb = Val(sjx(I1))
                            Yb = Val(sjx(I1 + 1))
                        End If
                        i2 = UBound(dmx)
                        If UBound(dmx) >= 4 Then
                            Xc = Val(dmx(i2 - 3))
                            Yc = Val(dmx(i2 - 2))
                        Else
                            Xc = Val(dmx(1)) + 13
                            Yc = Val(dmx(2))
                        End If
                        Xd = Val(dmx(i2 - 1))
                        Yd = Val(dmx(i2))
                        Call 计算交点(Xa, Ya, Xb, Yb, Xc, Yc, Xd, Yd, tt1, ss1, xjd2, Yjd2) '计算右幅交点
                        If tt1 >= 0 And tt1 < 1 Then                            '要确保设计线足够长才有交点
                            Print #19, xjd2 & "," & Yjd2 + dy
                            Exit For
                        End If
                    End If
                Next
                Print #19, ""                                                   '结束画设计线命令
                
                '标注设计线距离高程的文字样式：
                Print #19, "textstyle"                                          '建立文字样式
                Print #19, "标高字体"                                           '文字样式名称
                
                
                '标注坡脚线：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "坡脚线标注" & vbCrLf '设置当前图层为坡脚线标注
                Print #19, "line" & vbCrLf & Xjd1 & "," & Yjd1 + dy & vbCrLf _
                & "@" & length1 & "<90" & vbCrLf                                '画左幅坡脚引线
                Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                Xjd1 - Height1 / 5 & "," & Yjd1 + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                Print #19, "距离:" & Round(Xjd1, 3)                             '标注左幅坡脚距离
                Print #19, "text" & vbCrLf & "j" & vbCrLf & "tl" & vbCrLf & _
                Xjd1 + Height1 / 2 & "," & Yjd1 + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                Print #19, "高程:" & Round(Yjd1, 3)                             '标注左幅坡脚高程
                
                
                Print #19, "line" & vbCrLf & xjd2 & "," & Yjd2 + dy & vbCrLf _
                & "@" & length1 & "<90" & vbCrLf                                '画右幅坡脚引线
                Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                xjd2 - Height1 / 5 & "," & Yjd2 + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                Print #19, "距离:" & Round(xjd2, 3)                             '标注右幅坡脚距离
                Print #19, "text" & vbCrLf & "j" & vbCrLf & "tl" & vbCrLf & _
                xjd2 + Height1 / 2 & "," & Yjd2 + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                Print #19, "高程:" & Round(Yjd2, 3)                             '标注右幅坡脚高程
                
                '标注设计特征点：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "设计线标注" & vbCrLf '设置当前图层为设计线标注
                For I1 = 5 To UBound(sjx) - 1 Step 2
                    If UBound(sjx) Mod 2 = 1 Then                               '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                        Print #19, "line" & vbCrLf & Val(sjx(I1 - 1)) & "," & Val(sjx(1)) + Val(sjx(I1)) + dy & vbCrLf _
                        & "@" & length1 & "<90" & vbCrLf                        '画引线
                        '标注距离：
                        Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                        Val(sjx(I1 - 1)) - Height1 / 5 & "," & Val(sjx(1)) + Val(sjx(I1)) + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                        Print #19, "距离:" & Val(sjx(I1 - 1))
                        '标注高程：
                        Print #19, "text" & vbCrLf & "j" & vbCrLf & "tl" & vbCrLf & _
                        Val(sjx(I1 - 1)) + Height1 / 2 & "," & Val(sjx(1)) + Val(sjx(I1)) + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                        Print #19, "高程:" & Val(sjx(1)) + Val(sjx(I1))
                    Else                                                        '如果是里程1，平距1，高程1，平距2，高程2格式
                        Print #19, "line" & vbCrLf & Val(sjx(I1 - 2)) & "," & Val(sjx(I1 - 1)) + dy & vbCrLf _
                        & "@" & length1 & "<90" & vbCrLf                        '画引线
                        '标注距离：
                        Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                        Val(sjx(I1 - 2)) - Height1 / 5 & "," & Val(sjx(I1 - 1)) + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                        Print #19, "距离:" & Val(sjx(I1 - 2))
                        '标注高程：
                        Print #19, "text" & vbCrLf & "j" & vbCrLf & "tl" & vbCrLf & _
                        Val(sjx(I1 - 2)) + Height1 / 2 & "," & Val(sjx(I1 - 1)) + length2 * 2 + dy & vbCrLf & "90" '文字起点位置和标注方向
                        Print #19, "高程:" & Val(sjx(I1 - 1))
                    End If
                Next
                
                '画地面线：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "地面线" & vbCrLf '设置当前图层为地面线层
                Print #19, "line"
                Print #19, Xjd1 & "," & Yjd1 + dy
                For I1 = 2 To UBound(dmx) Step 2
                    Print #19, Val(dmx(I1 - 1)) & "," & Val(dmx(I1)) + dy
                Next
                Print #19, xjd2 & "," & Yjd2 + dy & vbCrLf
                '画引线：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "地面线标高" & vbCrLf '设置当前图层为地面线标高
                For I1 = 2 To UBound(dmx) Step 2
                    Print #19, "line" & vbCrLf & Val(dmx(I1 - 1)) + Height1 * 0.5773 & "," & Val(dmx(I1)) + Height1 + dy & vbCrLf & _
                    Val(dmx(I1 - 1)) & "," & Val(dmx(I1)) + dy & vbCrLf & _
                    Val(dmx(I1 - 1)) - Height1 * 0.5773 & "," & Val(dmx(I1)) + Height1 + dy & vbCrLf & _
                    Val(dmx(I1 - 1)) + Height1 * 2 & "," & Val(dmx(I1)) + Height1 + dy & vbCrLf
                    '标注高程：
                    Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                    Val(dmx(I1 - 1)) - Height1 * 1.4 & "," & Val(dmx(I1)) + Height1 * 1.1 + dy & vbCrLf & "0" '文字起点位置和标注方向
                    Print #19, Val(dmx(I1))
                Next
                '标注平距：
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "地面线平距" & vbCrLf '设置当前图层为地面线平距
                For I1 = 2 To UBound(dmx) Step 2
                    Print #19, "text" & vbCrLf & "j" & vbCrLf & "bl" & vbCrLf & _
                    Val(dmx(I1 - 1)) - Height1 * 1.4 & "," & Val(dmx(I1)) + Height1 * 2.4 + dy & vbCrLf & "0" '文字起点位置和标注方向
                    Print #19, Round(Val(dmx(I1 - 1)), 3)
                Next
                
                '求地面中桩高程
                If UBound(dmx) = 2 Then
                    hd = Round(Val(dmx(2)), 3)
                ElseIf UBound(dmx) > 2 Then
                    For I1 = 2 To UBound(dmx) - 2 Step 2                        '求中桩高程
                        If dmx(I1 - 1) = 0 Or dmx(I1 - 1) < 0 And dmx(I1 + 1) > 0 Then
                            hd = Val(dmx(I1)) + Abs(Val(dmx(I1 - 1))) * (Val(dmx(I1 + 2)) - Val(dmx(I1))) / (Abs(Val(dmx(I1 - 1))) + Val(dmx(I1 + 1)))
                            hd = Round(hd, 3)
                            Exit For
                        End If
                    Next
                End If
                
                '计算填方段或挖方段面积
                smax = UBound(sjx)
                dmax = UBound(dmx)
                Dim aa(200)                                                     '先把多边形顶点xy存放到aa中，有I1个顶点
                aa(0) = Xjd1
                aa(1) = Yjd1
                For I1 = 4 To UBound(sjx) - 2 Step 2
                    If UBound(sjx) Mod 2 = 1 Then                               '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                        aa(I1 - 2) = Val(sjx(I1))
                        aa(I1 - 1) = Val(sjx(1)) + Val(sjx(I1 + 1))
                    Else                                                        '如果是里程1，平距1，高程1，平距2，高程2格式
                        aa(I1 - 2) = Val(sjx(I1 - 1))
                        aa(I1 - 1) = Val(sjx(I1))
                    End If
                Next
                aa(I1 - 2) = xjd2
                aa(I1 - 1) = Yjd2
                For i2 = UBound(dmx) - 1 To 1 Step -2
                    If Val(dmx(i2)) > Xjd1 And Val(dmx(i2)) < xjd2 Then
                        aa(I1) = Val(dmx(i2))
                        aa(I1 + 1) = Val(dmx(i2 + 1))
                        I1 = I1 + 2
                    End If
                Next
                aa(I1) = Xjd1
                aa(I1 + 1) = Yjd1
                '已知坐标求多边形面积：不能计算半填半挖段。
                ss = 0
                For i2 = 0 To I1 - 2 Step 2
                    ss = ss + aa(i2) * aa(i2 + 3) - aa(i2 + 1) * aa(i2 + 2)
                Next
                ss = Round(Abs(ss) / 2, 3)
                
                '标注里程、设计标高、断面面积、填挖高
                '标注里程的文字样式：
                Print #19, "textstyle"                                          '建立文字样式
                Print #19, "里程字体"                                           '文字样式名称
                
                '计算设计高程
                If UBound(sjx) Mod 2 = 1 Then                                   '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                    Hs = Val(sjx(1))                                            '设计高程
                Else                                                            '如果是里程1,平距1，高程1，平距2，高程2格式
                    For i4 = 1 To UBound(sjx) - 2
                        If Val(sjx(i4)) <= 0 And Val(sjx(i4 + 2)) > 0 Then
                            Hs = Val(sjx(i4 + 1))                               '设计高程
                            Exit For
                        End If
                    Next
                End If
                
                Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "中线标注" & vbCrLf '设置当前图层为中线标注
                Print #19, "text" & vbCrLf & "j" & vbCrLf & "mc" & vbCrLf & _
0 & "," & Val(dmx(2)) - Height1 * 4 + dy & vbCrLf & "0"                         '文字起点位置和标注方向
                Print #19, Format(Val(sjx(0)), "\" & Text111) & "  Hs=" & Hs & "  Hd=" & hd & _
                "  Dh=" & Round(Hs - hd, 3) & "  AA=" & ss
                
                '向上偏移dy：
                If Val(Text113) = 0 Then
                    dy = dy + Val(sjx(1)) - Val(dmx(2)) + length1 * 2
                Else
                    dy = dy + Val(Text113)
                End If
                
            End If
        End If
    Loop
    
    '设置当前图层为0层
    Print #19, "-layer" & vbCrLf & "s" & vbCrLf & "0" & vbCrLf                  '设置当前图层为0层
    Print #19, "zoom" & vbCrLf & "e"
    
    Close (17)
    Close (18)
    Close (19)
    msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
    If MsgBox(msg, vbYesNo) = vbYes Then
        Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
    End If
Label1: Close
End Sub

Private Sub Command66_Click()
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "生成交点坐标的AUTOCAD脚本文件..."
        .DefaultExt = "scr"
        .Filter = "AUTOCAD脚本文件(*.scr)|*.scr"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
    End With
    Open CommonDialog1.FileName For Output As 20
    Print #20, "setvar"
    Print #20, "osnapcoord"
    Print #20, "1"                                                              '设为1也可以，2=键盘输入优先，脚本例外
    Print #20, "UCS"                                                            '在世界坐标系下画图
    Print #20, "W"
    
    '新建图层交点
    Print #20, "-layer"                                                         '首先新建图层
    Print #20, "m"                                                              'm=生成并设置为当前层
    Print #20, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "交点" '新建图层名称为交点
    Print #20, "c"                                                              '修改图层颜色
    Print #20, "130"                                                            '1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #20, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "交点" '输入需要修改颜色的图层里程标注
    Print #20, ""                                                               '结束图层设置命令
    
    Print #20, "line"
    For I1 = Val(Text64) To Val(Text64) + Val(Text65) - 1
        Print #20, factor2(I1, 2) & "," & factor2(I1, 1)                        '在世界坐标系下画图XY必须互换，Z=0
    Next
    Print #20, ""                                                               '结束line命令
    
    '新建图层交点标注
    Print #20, "-layer"                                                         '首先新建图层
    Print #20, "m"                                                              'm=生成并设置为当前层
    Print #20, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "交点标注" '新建图层名称为交点
    Print #20, "c"                                                              '修改图层颜色
    Print #20, "130"                                                            '1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
    Print #20, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "交点标注" '输入需要修改颜色的图层里程标注
    Print #20, ""                                                               '结束图层设置命令
    
    '设置文字样式
    Height1 = Val(TEXToffset) / 10                                              '设置文字高度
    length1 = Val(TEXToffset) / 10                                              '设置引线长度
    length2 = Val(TEXToffset) / 25                                              '设置引线偏移距离
    Print #20, "-style"                                                         '建立文字样式
    Print #20, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "交点标注" '文字样式名称
    Print #20, "simfang.ttf"                                                    '字体名称
    Write #20, Height1                                                          '文字高度
    Print #20, "1"                                                              '宽度比例
    Print #20, "0"                                                              '倾斜角度
    Print #20, "n"                                                              '反向显示文字=否
    Print #20, "n"                                                              '颠倒显示文字=否
    
    '标注文字
    For I1 = Val(Text64) To Val(Text64) + Val(Text65) - 1
        Print #20, "text"                                                       '输入单行文字text
        Print #20, "j"                                                          '修改文字格式
        Print #20, "mc"                                                         '居中对齐
        If factor2(I1, 3) < 0 Then
            Print #20, factor2(I1, 2) & "," & factor2(I1, 1) + 1.5 * length1    '文字起点位置
        Else
            Print #20, factor2(I1, 2) & "," & factor2(I1, 1) - 1.5 * length1    '文字起点位置
        End If
        Print #20, "0"                                                          '文字标注方向
        Print #20, "JD" & I1 & ",K=" & factor2(I1, 0) & ",X=" & factor2(I1, 1) & ",Y=" & factor2(I1, 2)
        '交点号，里程，XY
        '注意：单行文字命令自动结束的！不需要用回车来结束
    Next
    
    '设置当前图层为0层
    Print #20, "-layer"                                                         '设置图层
    Print #20, "s"                                                              '设置当前图层
    Print #20, "0"                                                              '设置当前图层为0层
    Print #20, ""                                                               '结束-layer命令
    Print #20, "zoom"
    Print #20, "e"
    Close #20
    msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
    If MsgBox(msg, vbYesNo) = vbYes Then
        Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
    End If
End Sub

Private Sub Command67_Click()
    Dim file1 As String
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "编辑现有的zas文件..."
        .DefaultExt = "zas"
        .Filter = "输入的曲线参数(*.zas)|*.zas"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        file1 = .FileName
    End With
    If Dir(file1) = "" Then
        '如果文件不存在就先把三行说明行写入文件
        Open CommonDialog1.FileName For Output As 1
        Print #1, "说明：前3行是说明行不要删除，第4行之后是数据行。平距左负右正。"
        Print #1, "第4行开始输入里程1,平距1,平距2,平距3,……,用英文逗号隔开"
        Print #1, "下一行输入输入里程1,平距1,平距2,平距3,……,回车后输入下一行。"
        Close #1
    End If
    '如果文件存在就直接打开：
    Shell "notepad.exe " & """" & file1 & """", vbMaximizedFocus
End Sub

Private Sub Command68_Click()
    If Text2 = "" Then
        MsgBox "请先载入平曲线参数！", vbInformation + vbOKOnly, "已知坐标反算里程偏距"
    Else
        Form2.Show
        Form2.Text1.SetFocus
    End If
End Sub

Private Sub Command69_Click()
    '生成卡西欧9860计算器主程序
    On Error Resume Next
    '查找FA-124程序编辑窗口句柄
    Dim Sx As String
    Dim ckbt As String, cklm As String
    Sx = GetOpenWindowNames("CASIO FA-124 - ")
    '    MsgBox Sx
    Dim hwnd As Long
    Dim hWnd2 As Long, hwnd3 As Long, hwnd4 As Long
    hwnd = FindWindow(vbNullString, Sx)                                         '通过完整标题名找句柄
    hWnd2 = FindWindowEx(hwnd, 0, "MDIClient", vbNullString)
    If hWnd2 <> 0 Then
        EnumChildWindows hWnd2, AddressOf EnumChildProc, ByVal 0&
        '        MsgBox Module1.cklm
        hwnd3 = FindWindowEx(hWnd2, 0, Module1.cklm, vbNullString)              '模块级变量cklm必须带上父变量model1.cklm
    End If
    hwnd4 = FindWindowEx(hwnd3, 0, "Edit", vbNullString)
    '    MsgBox hwnd4
    
    If hwnd4 = 0 Then
        MsgBox "请先打开CASIO FA-124传输软件，调出程序编辑窗口，并把光标定位到要输入的位置！", vbOKOnly + vbExclamation, "导出CASIO fx-9750G/9850G/9860G程序"
        Exit Sub
    End If
    Me.WindowState = 1                                                          '主程序最小化
    
    If Dir(App.Path & "\closeIME.exe") <> "" Then Shell App.Path & "\closeIME.exe", vbHide
    
    '把casio fa-124程序编辑框带到前台并关闭输入法
    '    SetWindowPos hwnd, 0, 0, 0, 0, 0, 3     '设置置顶'Private Const HWND_TOP = 0
    '    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE     '取消置顶
    '以下代码激活父窗体：
    SetForegroundWindow hwnd
    '以下代码激活子窗体：
    'Private Const WM_MOUSEACTIVATE = &H21
    SendMessage hwnd4, WM_MOUSEACTIVATE, 0, 0
    Sleep 2000                                                                  '等待焦点切换到FA-124程序编辑框
    
    '    '以下代码把前台进程的输入法关闭：
    '    '-------------------------------------------------------------------------------------------------------
    '    Dim hwnd As Long                    '前面已经定义过了，不要定义
    On Error Resume Next
    Dim hCurThread As Long
    Dim sCaption As String
    Dim sBuffer As String
    Dim hCurKBDLayout As Long
    hwnd = GetForegroundWindow
    sCaption = Space(255)
    GetWindowText hwnd, sCaption, 255
    If InStr(sCaption, Chr(0)) Then
        sCaption = Left(sCaption, InStr(sCaption, Chr(0)) - 1)
    End If
    hCurThread = GetWindowThreadProcessId(hwnd, ByVal 0&)
    hCurKBDLayout = GetKeyboardLayout(hCurThread)
    If ImmIsIME(hCurKBDLayout) = 1 Then
        sBuffer = Space(255)
        RetCount = ImmGetDescription(ByVal hCurKBDLayout, sBuffer, 255)
        sBuffer = Left(sBuffer, RetCount)
        '        sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        sBuffer = "English(American)"
    End If
    Debug.Print sBuffer
    If sBuffer <> "English(American)" And sBuffer <> "" Then SendKeys "^ ", False
    Sleep 1000                                                                  '延迟1秒确保ctrl+空格生效
    '-------------------------------------------------------------------------------------------------------
    
    '修正大小写状态
    If GetKeyState(vbKeyCapital) = 1 Then                                       '如果CAPSLOCK被按下
        MySendKey vbKeyCapital                                                  '就发送{capslock}键
    End If
    
    '开始写程序了：
    '变量K回显：
    '        SendKeys "Deg:""K=""", False
    MySendKeytext "Deg:""K="""
    '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '        SendKeys "Str 1:1", False
    MySendKeytext "Str 1:1"
    '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '        SendKeys "%+k{UP 5}{ENTER 2}", False 'θ
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn, vbKeyReturn
    '        SendKeys "K", False
    MySendKeytext "K"
    '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '        SendKeys "%+k{UP 6}{ENTER}", False 'r
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
    '        SendKeys ":Prog ""A1"":", False
    MySendKeytext ":Prog ""A1"":"
    '        SendKeys "%+k{UP 6}{ENTER}", False 'r
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
    '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '        SendKeys "K{ENTER}", False
    MySendKeytext "K"
    MySendKey vbKeyReturn
    '变量T回显：
    '    SendKeys "Deg:""T=""", False
    MySendKeytext "Deg:""T="""
    '    SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '    SendKeys "Str 1:2", False
    MySendKeytext "Str 1:2"
    '    SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '    SendKeys "%+k{UP 5}{ENTER 2}", False 'θ
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn, vbKeyReturn
    '    SendKeys "T", False
    MySendKeytext "T"
    '            SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '    SendKeys "%+k{UP 6}{ENTER}", False  'r
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
    '        SendKeys ":Prog ""A1"":", False
    MySendKeytext ":Prog ""A1"":"
    '        SendKeys "%+k{UP 6}{ENTER}", False 'r
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
    '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '   SendKeys "T{ENTER}", False
    MySendKeytext "T"
    MySendKey vbKeyReturn
    
    '变量P回显：
    '    SendKeys "Deg:""P=""", False
    MySendKeytext "Deg:""P="""
    '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '    SendKeys "Str 1:3", False
    MySendKeytext "Str 1:3"
    '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '        SendKeys "%+k{UP 5}{ENTER 2}", False 'θ
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn, vbKeyReturn
    '        SendKeys "P", False
    MySendKeytext "P"
    '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '        SendKeys "%+k{UP 6}{ENTER}", False 'r
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
    '        SendKeys ":Prog ""A1"":", False
    MySendKeytext ":Prog ""A1"":"
    '        SendKeys "%+k{UP 6}{ENTER}", False 'r
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
    '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '    SendKeys "P{ENTER}", False
    MySendKeytext "P"
    MySendKey vbKeyReturn
    
    '    SendKeys "K{+}T", False
    MySendKeytext "K+T"
    '    SendKeys "%+k{DOWN 2}{ENTER}", False '×
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
    '    SendKeys "cos {(}P{)}", False
    MySendKeytext "cos (P)"
    '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '    SendKeys "V{ENTER}", False
    MySendKeytext "V"
    MySendKey vbKeyReturn
    
    
    For n1 = 2 To Val(Text24) - 1
        'SendKeys "If V", False
        MySendKeytext "If V"
        'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
        SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        'SendKeys shu(n1, 0), False
        MySendKeytext shu(n1, 0)
        'SendKeys ":Then ", False
        MySendKeytext ":Then "
        'SendKeys shu(n1, 1), False
        MySendKeytext shu(n1, 1)
        '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        'SendKeys "J:", False
        MySendKeytext "J:"
        'SendKeys shu(n1, 2), False
        MySendKeytext shu(n1, 2)
        '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        'SendKeys "Y:", False
        MySendKeytext "Y:"
        'SendKeys shu(n1, 3), False
        MySendKeytext shu(n1, 3)
        '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        'SendKeys "R:", False
        MySendKeytext "R:"
        'SendKeys Format(shu(n1, 4) * 100, "0.#####"), False
        MySendKeytext Format(shu(n1, 4) * 100, "0.#####")
        'SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        'SendKeys "E:", False
        MySendKeytext "E:"
        'SendKeys Format(shu(n1, 5) * 100, "0.#####"), False
        MySendKeytext Format(shu(n1, 5) * 100, "0.#####")
        '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '                SendKeys "F:", False
        MySendKeytext "F:"
        'SendKeys "Goto 1:IfEnd{ENTER}", False
        MySendKeytext "Goto 1:IfEnd"
        MySendKey vbKeyReturn
    Next
    'SendKeys "Lbl 1:Prog ""GC""", False '最后一个sendkeys等待发送完毕
    MySendKeytext "Lbl 1:Prog ""GC"""
    '    SendKeys "%+p{DOWN 4}{ENTER}", False '◢
    SendKeyALTplus vbKeyP, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    '边桩高程计算，考虑超高
    hengpo(0, 0) = hengpo(1, 0)
    hengpo(0, 1) = hengpo(1, 1)
    hengpo(0, 2) = hengpo(1, 2)
    'If K2 <> K1 Then I = I1 + (KG - K1) * (I2 - I1) / (K2 - K1)
    'If K2 = K1 Then I = I1
    'If K4 <> K3 Then I = I3 + (KG - K3) * (I4 - I3) / (K4 - K3)
    'If K4 = K3 Then I = I3
    'ts = Val(Text43)
    'G = H - I * (Abs(T * Sin(P1)) - Abs(ts))
    
    For n1 = 1 To Val(Text33)
        K1 = CDec(hengpo(n1 - 1, 0))
        K2 = CDec(hengpo(n1, 0))
        K3 = CDec(hengpo(n1 - 1, 0))
        K4 = CDec(hengpo(n1, 0))
        I1 = CDec(hengpo(n1 - 1, 1))
        i2 = CDec(hengpo(n1, 1))
        I3 = CDec(hengpo(n1 - 1, 2))
        i4 = CDec(hengpo(n1, 2))
        If hengpo(n1, 3) = 0 Then                                               '如果是线性超高
            If i2 <> I1 And I3 <> i4 And K2 <> K1 And K3 <> K4 Then
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys hengpo(n1, 0), False
                MySendKeytext hengpo(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys I1, False
                MySendKeytext I1
                'SendKeys "{+}{(}V-", False
                MySendKeytext "+(V-"
                'SendKeys K3, False
                MySendKeytext K3
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (i2 - I1), False
                MySendKeytext (i2 - I1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (K2 - K1), False
                MySendKeytext (K2 - K1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                
                '                SendKeys I3, False
                MySendKeytext I3
                '                SendKeys "{+}{(}V-", False
                MySendKeytext "+(V-"
                '                SendKeys K3, False
                MySendKeytext K3
                '                SendKeys "{)}", False
                MySendKeytext ")"
                '                SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "{(}", False
                MySendKeytext "("
                '                SendKeys (i4 - I3), False
                MySendKeytext (i4 - I3)
                '                SendKeys "{)}", False
                MySendKeytext ")"
                '                SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "{(}", False
                MySendKeytext "("
                '                SendKeys (K4 - K3), False
                MySendKeytext (K4 - K3)
                '                SendKeys "{)}", False
                MySendKeytext ")"
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "B:", False
                MySendKeytext "B:"
                '                SendKeys "Goto 2:IfEnd{ENTER}", False
                MySendKeytext "Goto 2:IfEnd"
                MySendKey vbKeyReturn
                
                '                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                '                I1 & "+(V-" & K1 & ")*(" & (i2 - I1) & ")/(" & (K2 - K1) & ")→A:" & _
                '                I3 & "+(V-" & K3 & ")*(" & (i4 - I3) & ")/(" & (K4 - K3) & ")→B:" & _
                '                "Goto 2:IfEnd↙"
            ElseIf i2 <> I1 And I3 = i4 And K2 <> K1 Then
                'SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys hengpo(n1, 0), False
                MySendKeytext hengpo(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys I1, False
                MySendKeytext I1
                'SendKeys "{+}{(}V-", False
                MySendKeytext "+(V-"
                'SendKeys K1, False
                MySendKeytext K1
                'SendKeys "{)}", False
                MySendKeytext ")"
                '                SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (i2 - I1), False
                MySendKeytext (i2 - I1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                '                SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (K2 - K1), False
                MySendKeytext (K2 - K1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                
                '                SendKeys I3, False
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 2:IfEnd{ENTER}", False
                MySendKeytext "Goto 2:IfEnd"
                MySendKey vbKeyReturn
                '                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                '                I1 & "+(V-" & K1 & ")*(" & (i2 - I1) & ")/(" & (K2 - K1) & ")→A:" & _
                '                I3 & "→B:" & _
                '                "Goto 2:IfEnd↙"
            ElseIf i2 = I1 And I3 <> i4 And K3 <> K4 Then
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys hengpo(n1, 0), False
                MySendKeytext hengpo(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys I1, False
                MySendKeytext I1
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                
                'SendKeys I3, False
                MySendKeytext I3
                'SendKeys "{+}{(}V-", False
                MySendKeytext "+(V-"
                'SendKeys K3, False
                SendKeys K3
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (i4 - I3), False
                MySendKeytext (i4 - I3)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (K4 - K3), False
                MySendKeytext (K4 - K3)
                '                SendKeys "{)}", False
                MySendKeytext ")"
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 2:IfEnd{ENTER}", False
                MySendKeytext "Goto 2:IfEnd"
                MySendKey vbKeyReturn
                
                '                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                '                I1 & "→A:" & _
                '                I3 & "+(V-" & K3 & ")*(" & (i4 - I3) & ")/(" & (K4 - K3) & ")→B:" & _
                '                "Goto 2:IfEnd↙"
            Else
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys hengpo(n1, 0), False
                MySendKeytext hengpo(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys I1, False
                MySendKeytext I1
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                
                'SendKeys I3, False
                MySendKeytext I3
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 2:IfEnd{ENTER}", False
                MySendKeytext "Goto 2:IfEnd"
                MySendKey vbKeyReturn
                '                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                '                I1 & "→A:" & _
                '                I3 & "→B:" & _
                '                "Goto 2:IfEnd↙"
            End If
        ElseIf hengpo(n1, 3) = 1 Then                                           '如果是三次抛物线超高
            If i2 <> I1 And I3 <> i4 And K2 <> K1 And K3 <> K4 Then
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys hengpo(n1, 0), False
                MySendKeytext hengpo(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys "{(}K-", False
                MySendKeytext "(K-"
                'SendKeys K1, False
                MySendKeytext K1
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (K2 - K1), False
                MySendKeytext (K2 - K1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "I:", False
                MySendKeytext "I:"
                
                'SendKeys I1, False
                MySendKeytext I1
                'SendKeys "{+}{(}3", False
                MySendKeytext "+(3"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{I}{^}2-2", False
                MySendKeytext "I^2-2"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{I}{^}3{)}", False
                MySendKeytext "I^3)"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (i2 - I1), False
                MySendKeytext (i2 - I1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                
                'SendKeys I3, False
                MySendKeytext I3
                '                SendKeys "{+}{(}3", False
                MySendKeytext "+(3"
                '                SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "{I}{^}2-2", False
                MySendKeytext "I^2-2"
                '                SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "{I}{^}3{)}", False
                MySendKeytext "I^3)"
                '                SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "{(}", False
                MySendKeytext "("
                '                SendKeys (i4 - I3), False
                MySendKeytext (i4 - I3)
                '                SendKeys "{)}", False
                MySendKeytext "{)}"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 2:IfEnd{ENTER}", False
                MySendKeytext "Goto 2:IfEnd"
                MySendKey vbKeyReturn
                '                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                '                "(K-" & K1 & ")/(" & (K2 - K1) & ")→I:" & _
                '                I1 & "+(3×I^2-2×I^3)×(" & (i2 - I1) & ")→A:" & _
                '                I3 & "+(3×I^2-2×I^3)×(" & (i4 - I3) & ")→B:" & _
                '                "Goto 2:IfEnd↙"
            ElseIf i2 <> I1 And I3 = i4 And K2 <> K1 Then
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys hengpo(n1, 0), False
                MySendKeytext hengpo(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys "{(}K-", False
                MySendKeytext "(K-"
                'SendKeys K1, False
                MySendKeytext K1
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (K2 - K1), False
                MySendKeytext (K2 - K1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "I:", False
                MySendKeytext "I:"
                
                'SendKeys I1, False
                MySendKeytext I1
                'SendKeys "{+}{(}3", False
                MySendKeytext "+(3"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{I}{^}2-2", False
                MySendKeytext "I^2-2"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{I}{^}3{)}", False
                MySendKeytext "I^3)"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (i2 - I1), False
                MySendKeytext (i2 - I1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                
                'SendKeys I3, False
                MySendKeytext I3
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 2:IfEnd{ENTER}", False
                MySendKeytext "Goto 2:IfEnd"
                MySendKey vbKeyReturn
                '                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                '                "(K-" & K1 & ")/(" & (K2 - K1) & ")→I:" & _
                '                I1 & "+(3×I^2-2×I^3)×(" & (i2 - I1) & ")→A:" & _
                '                I3 & "→B:" & _
                '                "Goto 2:IfEnd↙"
            ElseIf i2 = I1 And I3 <> i4 And K3 <> K4 Then
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys hengpo(n1, 0), False
                MySendKeytext hengpo(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys "{(}K-", False
                MySendKeytext "(K-"
                'SendKeys K1, False
                MySendKeytext K1
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (K2 - K1), False
                MySendKeytext (K2 - K1)
                '                SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "I:", False
                MySendKeytext "I:"
                
                'SendKeys I1, False
                MySendKeytext I1
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                
                'SendKeys I3, False
                MySendKeytext I3
                'SendKeys "{+}{(}3", False
                MySendKeytext "+(3"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{I}{^}2-2", False
                MySendKeytext "I^2-2"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{I}{^}3{)}", False
                MySendKeytext "I^3"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (i4 - I3), False
                MySendKeytext (i4 - I3)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 2:IfEnd{ENTER}", False
                MySendKeytext "Goto 2:IfEnd"
                MySendKey vbKeyReturn
                '                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                '                "(K-" & K1 & ")/(" & (K2 - K1) & ")→I:" & _
                '                I1 & "→A:" & _
                '                I3 & "+(3×I^2-2×I^3)×(" & (i4 - I3) & ")→B:" & _
                '                "Goto 2:IfEnd↙"
            Else
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys hengpo(n1, 0), False
                MySendKeytext hengpo(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                
                'SendKeys I1, False
                MySendKeytext I1
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                
                'SendKeys I3, False
                MySendKeytext I3
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 2:IfEnd{ENTER}", False
                MySendKeytext "Goto 2:IfEnd"
                MySendKey vbKeyReturn
                '                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                '                I1 & "→A:" & _
                '                I3 & "→B:" & _
                '                "Goto 2:IfEnd↙"
            End If
        End If
    Next
    'SendKeys "Lbl 2:If T", False
    MySendKeytext "Lbl 2:If T"
    'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
    SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "0:Then A", False
    MySendKeytext "0:Then A"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "I:Else B", False
    MySendKeytext "I:Else B"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "I:IfEnd{ENTER}", False
    MySendKeytext "I:IfEnd"
    MySendKey vbKeyReturn
    
    'SendKeys "H-I", False
    MySendKeytext "H-I"
    'SendKeys "%+k{DOWN 2}{ENTER}", False '×
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "{(}Abs {(}T", False
    MySendKeytext "(Abs (T"
    'SendKeys "%+k{DOWN 2}{ENTER}", False '×
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "sin {(}P{)}{)}-", False
    MySendKeytext "sin (P))-"
    'SendKeys Val(Text43), False
    MySendKeytext Val(Text43)
    'SendKeys "{)}", False
    MySendKeytext ")"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
    vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "G:""G="":G", False
    MySendKeytext "G:""G="":G"
    DoEvents                                                                    '把控制权交还给主程序
    '    Print #13, "Lbl 2:If T≤0:Then A→I:Else B→I:IfEnd↙"
    '    Print #13, "H-I×(Abs (T×sin (P))-" & Val(Text43) & ")→G:“G=”:G"
    'fa-124程序编写完毕，弹出对话框：
    Me.SetFocus
    MsgBox "CASIO FA-124高程程序输出完毕！", vbOKOnly + vbExclamation, "导出CASIO fx-9750G/9850G/9860G程序"
End Sub


Private Sub Command70_Click()
    '生成卡西欧9860计算器主程序
    On Error Resume Next
    '查找FA-124程序编辑窗口句柄
    Dim Sx As String
    Dim ckbt As String, cklm As String
    Sx = GetOpenWindowNames("CASIO FA-124 - ")
    '    MsgBox Sx
    Dim hwnd As Long
    Dim hWnd2 As Long, hwnd3 As Long, hwnd4 As Long
    hwnd = FindWindow(vbNullString, Sx)                                         '通过完整标题名找句柄
    hWnd2 = FindWindowEx(hwnd, 0, "MDIClient", vbNullString)
    If hWnd2 <> 0 Then
        EnumChildWindows hWnd2, AddressOf EnumChildProc, ByVal 0&
        '        MsgBox Module1.cklm
        hwnd3 = FindWindowEx(hWnd2, 0, Module1.cklm, vbNullString)              '模块级变量cklm必须带上父变量model1.cklm
    End If
    hwnd4 = FindWindowEx(hwnd3, 0, "Edit", vbNullString)
    '    MsgBox hwnd4
    If hwnd4 = 0 Then
        MsgBox "请先打开CASIO FA-124传输软件，调出程序编辑窗口，并把光标定位到要输入的位置！", vbOKOnly + vbExclamation, "导出CASIO fx-9750G/9850G/9860G程序"
        Exit Sub
    End If
    Me.WindowState = 1                                                          '主程序最小化
    
    If Dir(App.Path & "\closeIME.exe") <> "" Then Shell App.Path & "\closeIME.exe", vbHide
    
    '把casio fa-124程序编辑框带到前台并关闭输入法
    '    SetWindowPos hwnd, 0, 0, 0, 0, 0, 3     '设置置顶'Private Const HWND_TOP = 0
    '    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE     '取消置顶
    '以下代码激活父窗体：
    SetForegroundWindow hwnd
    '以下代码激活子窗体：
    'Private Const WM_MOUSEACTIVATE = &H21
    SendMessage hwnd4, WM_MOUSEACTIVATE, 0, 0
    Sleep 2000                                                                  '等待焦点切换到FA-124程序编辑框
    
    '    '以下代码把前台进程的输入法关闭：
    '    '-------------------------------------------------------------------------------------------------------
    '    Dim hwnd As Long                    '前面已经定义过了，不要定义
    On Error Resume Next
    Dim hCurThread As Long
    Dim sCaption As String
    Dim sBuffer As String
    Dim hCurKBDLayout As Long
    hwnd = GetForegroundWindow
    sCaption = Space(255)
    GetWindowText hwnd, sCaption, 255
    If InStr(sCaption, Chr(0)) Then
        sCaption = Left(sCaption, InStr(sCaption, Chr(0)) - 1)
    End If
    hCurThread = GetWindowThreadProcessId(hwnd, ByVal 0&)
    hCurKBDLayout = GetKeyboardLayout(hCurThread)
    If ImmIsIME(hCurKBDLayout) = 1 Then
        sBuffer = Space(255)
        RetCount = ImmGetDescription(ByVal hCurKBDLayout, sBuffer, 255)
        sBuffer = Left(sBuffer, RetCount)
        '        sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        sBuffer = "English(American)"
    End If
    Debug.Print sBuffer
    If sBuffer <> "English(American)" And sBuffer <> "" Then SendKeyCTRLplus vbKeySpace
    Sleep 2000                                                                  '延迟1秒确保ctrl+空格生效
    '-------------------------------------------------------------------------------------------------------
    
    '修正大小写状态
    If GetKeyState(vbKeyCapital) = 1 Then                                       '如果CAPSLOCK被按下
        'SendKeys "{CAPSLOCK}", False    '就发送{capslock}键
        MySendKey vbKeyCapital                                                  '就发送{capslock}键
    End If
    
    
    '开始写程序了：
    '变量K回显：
    'SendKeys "Deg:""K=""", False
    MySendKeytext "Deg:""K="""
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "Str 1:1", False
    MySendKeytext "Str 1:1"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "%+k{UP 5}{ENTER 2}", False 'θ
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn, vbKeyReturn
    'SendKeys "K", False
    MySendKeytext "K"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "%+k{UP 6}{ENTER}", False  'r
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
    'SendKeys ":Prog ""A1"":", False
    MySendKeytext ":Prog ""A1"":"
    'SendKeys "%+k{UP 6}{ENTER}", False  'r
    SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "K{ENTER}", False
    MySendKeytext "K"
    MySendKey vbKeyReturn
    '    Print #17, "？K：↙"
    
    kuandu(0, 0) = kuandu(1, 0)
    kuandu(0, 1) = kuandu(1, 1)
    kuandu(0, 2) = kuandu(1, 2)
    For n1 = 1 To Val(Text97)
        K1 = CDec(kuandu(n1 - 1, 0))                                            '转换为decimal类型
        K2 = CDec(kuandu(n1, 0))
        L1 = CDec(kuandu(n1 - 1, 1))
        L2 = CDec(kuandu(n1, 1))
        If Val(kuandu(n1, 2)) = 0 Then                                          '如果=0就是线性渐变
            If L2 <> L1 And K2 <> K1 Then
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys kuandu(n1, 0), False
                MySendKeytext kuandu(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys L1, False
                MySendKeytext L1
                'SendKeys "{+}{(}K-", False
                MySendKeytext "+(K-"
                'SendKeys K1, False
                MySendKeytext K1
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (L2 - L1), False
                MySendKeytext (L2 - L1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (K2 - K1), False
                MySendKeytext (K2 - K1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                'SendKeys "Goto 3:IfEnd{ENTER}", False
                MySendKeytext "Goto 3:IfEnd"
                MySendKey vbKeyReturn
                '                Print #17, "If K≤" & kuandu(n1, 0) & "：Then " & _
                '                L1 & "+(K-" & K1 & ")×(" & (L2 - L1) & ")/(" & (K2 - K1) & ")→A：" & _
                '                "Goto 3：IfEnd↙"
            Else
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys kuandu(n1, 0), False
                MySendKeytext kuandu(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys L1, False
                MySendKeytext L1
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                'SendKeys "Goto 3:IfEnd{ENTER}", False
                MySendKeytext "Goto 3:IfEnd"
                MySendKey vbKeyReturn
                '                Print #17, "If K≤" & kuandu(n1, 0) & "：Then " & _
                '                L1 & "→A：" & _
                '                "Goto 3：IfEnd↙"
            End If
        ElseIf Val(kuandu(n1, 2)) = 1 Then                                      '如果是抛物线渐变
            If L2 <> L1 And K1 <> K2 Then
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys kuandu(n1, 0), False
                MySendKeytext kuandu(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys "{(}K-", False
                MySendKeytext "(K-"
                'SendKeys K1, False
                MySendKeytext K1
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (K2 - K1), False
                MySendKeytext (K2 - K1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "I:", False
                MySendKeytext "I:"
                
                'SendKeys L1, False
                MySendKeytext L1
                'SendKeys "{+}{(}3", False
                MySendKeytext "+(3"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{I}{^}2-2", False
                MySendKeytext "I^2-2"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{I}{^}3{)}", False
                MySendKeytext "I^3)"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (L2 - L1), False
                MySendKeytext (L2 - L1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                'SendKeys "Goto 3:IfEnd{ENTER}", False
                MySendKeytext "Goto 3:IfEnd"
                MySendKey vbKeyReturn
                '                Print #17, "If K≤" & kuandu(n1, 0) & "：Then " & _
                '                "(K-" & K1 & ")/(" & (K2 - K1) & ")→I：" & _
                '                L1 & "+(3×I^2-2×I^3)×(" & (L2 - L1) & ")→A：" & _
                '                "Goto 3：IfEnd↙"
            Else
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys kuandu(n1, 0), False
                MySendKeytext kuandu(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys L1, False
                MySendKeytext L1
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "A:", False
                MySendKeytext "A:"
                'SendKeys "Goto 3:IfEnd{ENTER}", False
                MySendKeytext "Goto 3:IfEnd"
                MySendKey vbKeyReturn
                '                Print #17, "If K≤" & kuandu(n1, 0) & "：Then " & _
                '                L1 & "→A：" & _
                '                "Goto 3：IfEnd↙"
            End If
        End If
    Next
    'SendKeys "Lbl 3:{ENTER}", False
    MySendKeytext "Lbl 3:"
    MySendKey vbKeyReturn
    '    Print #17, "Lbl 3：↙"
    
    kuandu2(0, 0) = kuandu2(1, 0)
    kuandu2(0, 1) = kuandu2(1, 1)
    kuandu2(0, 2) = kuandu2(1, 2)
    For n1 = 1 To Val(Text102)
        K1 = CDec(kuandu2(n1 - 1, 0))                                           '转换为decimal类型
        K2 = CDec(kuandu2(n1, 0))
        R1 = CDec(kuandu2(n1 - 1, 1))
        R2 = CDec(kuandu2(n1, 1))
        If Val(kuandu2(n1, 2)) = 0 Then                                         '如果=0就是线性渐变
            If R2 <> R1 And K2 <> K1 Then
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys kuandu2(n1, 0), False
                MySendKeytext kuandu2(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys R1, False
                MySendKeytext R1
                'SendKeys "{+}{(}K-", False
                MySendKeytext "+(K-"
                'SendKeys K1, False
                MySendKeytext K1
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (R2 - R1), False
                MySendKeytext (R2 - R1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (K2 - K1), False
                MySendKeytext (K2 - K1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 4:IfEnd{ENTER}", False
                MySendKeytext "Goto 4:IfEnd"
                MySendKey vbKeyReturn
                '                Print #17, "If K≤" & kuandu2(n1, 0) & "：Then " & _
                '                R1 & "+(K-" & K1 & ")×(" & (R2 - R1) & ")/(" & (K2 - K1) & ")→B：" & _
                '                "Goto 4：IfEnd↙"
            Else
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys kuandu2(n1, 0), False
                MySendKeytext kuandu2(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys R1, False
                MySendKeytext R1
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 4:IfEnd{ENTER}", False
                MySendKeytext "Goto 4:IfEnd"
                MySendKey vbKeyReturn
                '                Print #17, "If K≤" & kuandu2(n1, 0) & "：Then " & _
                '                R1 & "→B：" & _
                '                "Goto 4：IfEnd↙"
            End If
        ElseIf Val(kuandu2(n1, 2)) = 1 Then                                     '如果是抛物线渐变
            If R2 <> R1 And K1 <> K2 Then
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys kuandu2(n1, 0), False
                MySendKeytext kuandu2(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys "{(}K-", False
                MySendKeytext "(K-"
                'SendKeys K1, False
                MySendKeytext K1
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 3}{ENTER}", False '÷
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (K2 - K1), False
                MySendKeytext (K2 - K1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "I:", False
                MySendKeytext "I:"
                
                'SendKeys R1, False
                MySendKeytext R1
                'SendKeys "{+}{(}3", False
                MySendKeytext "+(3"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{I}{^}2-2", False
                MySendKeytext "I^2-2"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{I}{^}3{)}", False
                MySendKeytext "I^3)"
                'SendKeys "%+k{DOWN 2}{ENTER}", False '×
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "{(}", False
                MySendKeytext "("
                'SendKeys (R2 - R1), False
                MySendKeytext (R2 - R1)
                'SendKeys "{)}", False
                MySendKeytext ")"
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 4:IfEnd{ENTER}", False
                MySendKeytext "Goto 4:IfEnd"
                MySendKey vbKeyReturn
                '                Print #17, "If K≤" & kuandu2(n1, 0) & "：Then " & _
                '                "(K-" & K1 & ")/(" & (K2 - K1) & ")→I：" & _
                '                R1 & "+(3×I^2-2×I^3)×(" & (R2 - R1) & ")→B：" & _
                '                "Goto 4：IfEnd↙"
            Else
                'SendKeys "If K", False
                MySendKeytext "If K"
                'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys kuandu2(n1, 0), False
                MySendKeytext kuandu2(n1, 0)
                'SendKeys ":Then ", False
                MySendKeytext ":Then "
                'SendKeys R1, False
                MySendKeytext R1
                'SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                'SendKeys "B:", False
                MySendKeytext "B:"
                'SendKeys "Goto 4:IfEnd{ENTER}", False
                MySendKeytext "Goto 4:IfEnd"
                MySendKey vbKeyReturn
                '                Print #17, "If K≤" & kuandu2(n1, 0) & "：Then " & _
                '                R1 & "→B：" & _
                '                "Goto 4：IfEnd↙"
            End If
        End If
    Next
    'SendKeys "Lbl 4:""A="":A", False
    MySendKeytext "Lbl 4:""A="":A"
    'SendKeys "%+p{DOWN 4}{ENTER}", False '◢
    SendKeyALTplus vbKeyP, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "Lbl 4:""B="":B", False
    MySendKeytext "Lbl 4:""B="":B"
    'SendKeys "%+p{DOWN 4}{ENTER}", False '◢
    SendKeyALTplus vbKeyP, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    
    'SendKeys "-A", False
    MySendKeytext "-A"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "List 26[1]:", False       '把左幅宽度保存在List26[1]中
    MySendKeytext "List 26[1]:"
    'SendKeys "B", False
    MySendKeytext "B"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "List 26[2]:", False       '把右幅宽度保存在List26[2]中
    MySendKeytext "List 26[2]:"
    
    'SendKeys "-A", False
    MySendKeytext "-A"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "T:Prog ""1ZKGC""", False
    MySendKeytext "T:Prog ""1ZKGC"""
    'SendKeys "%+p{DOWN 4}{ENTER}", False '◢
    SendKeyALTplus vbKeyP, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    
    'SendKeys """C=""?", False
    MySendKeytext """C=""?"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "C:If T", False
    MySendKeytext "C:If T"
    'SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
    SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "0:Then List 26[1]-{(}G-C{)}", False
    MySendKeytext "0:Then List 26[1]-(G-C)"
    'SendKeys "%+k{DOWN 2}{ENTER}", False '×
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "1.5", False
    MySendKeytext "1.5"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "W:Else List 26[2]{+}{(}G-C{)}", False
    MySendKeytext "W:Else List 26[2]+(G-C)"
    'SendKeys "%+k{DOWN 2}{ENTER}", False '×
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "1.5", False
    MySendKeytext "1.5"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "W:", False
    MySendKeytext "W:"
    'SendKeys "IfEnd{ENTER}", False
    MySendKeytext "IfEnd"
    MySendKey vbKeyReturn
    
    'SendKeys """W="":W", False
    MySendKeytext """W="":W"
    'SendKeys "%+p{DOWN 4}{ENTER}", False '◢
    SendKeyALTplus vbKeyP, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "W", False
    MySendKeytext "W"
    'SendKeys "%+k{DOWN 24}{ENTER}", False '→
    SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
    'SendKeys "T:Prog ""1HXZK""", False
    MySendKeytext "T:Prog ""1HXZK"""
    DoEvents                                                                    '交还控制权给主程序
    '    Print #17, "Lbl 4：“A=”：A◢“B=”：B◢-A→T：Prog“ZKGC”◢"
    '    Print #17, "？C：If T≤0：Then -A-（G-C）×1.5→W：Else B+（G-C）×1.5→W：IfEnd↙"
    '    Print #17, "“W=”：W◢W→T：Prog“ZK”"
    'fa-124程序编写完毕，弹出对话框：
    Me.SetFocus
    MsgBox "CASIO FA-124宽度程序输出完毕！", vbOKOnly + vbExclamation, "导出CASIO fx-9750G/9850G/9860G程序"
End Sub

Private Sub Command71_Click()
    Form4.Show
End Sub

'----------------------------------------------------------------------------------------------------------
Private Sub Form_Load()
    Saveflag = 0                                                                '在Form_Load事件里预设SaveFlag=0未存盘
    TextPartNumber = 1
    Combo2.Clear
    Combo2.AddItem "直线"
    Combo2.AddItem "圆"
    Combo2.AddItem "缓1+圆+缓2"
    Combo2.AddItem "布劳斯曲线"
    Combo2.AddItem "麦克康纳尔曲线"
    Combo2.ListIndex = 0
    Text24 = 3
    Text33 = 1
    Text97 = 1
    Text102 = 1
    角度运算
    直线2
    载入序号
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Saveflag = 1 Or TEXTKI = 0 Then
        Dim ix As Integer
        For ix = Forms.Count - 1 To 1 Step -1
            Unload Forms(ix)
        Next
        End                                                                     '如果曲线参数已保存就直接GAME OVER......
    Else
        If MsgBox("平曲线参数未保存,确定退出吗?", vbYesNo, "警告!") = vbYes Then '询问是否退出
            For ix = Forms.Count - 1 To 1 Step -1
                Unload Forms(ix)
            Next
            End                                                                 '如果单击确定就GMAE OVER,单击取消就返回
        End If
    End If
End Sub

Sub 角度运算()
    J1 = CDec(Val(Text48))
    J2 = CDec(Val(Text49))
    度分秒化为弧度 J1, J11
    度分秒化为弧度 J2, J22
    J3 = J11 + J22
    Text52 = Round(J3, 8)
    Text51 = Round(J3 * 180 / 3.14159265358979, 8)
    弧度化为度分秒 J3
    Text50 = Round(J3, 5)
End Sub

Sub 直线2()
    m = CDec(Val(Text57))
    n = CDec(Val(Text58))
    i = CDec(Val(Text59))
    J = CDec(Val(Text60))
    度分秒化为弧度 J, J1
    x = m + i * Cos(J1)
    y = n + i * Sin(J1)
    Text61 = Round(x, 3)
    Text62 = Round(y, 3)
End Sub

Sub 载入序号()
    Dim I1 As Double                                                            '循环用的变量
    Combo6.Clear
    For I1 = Val(Text64) To Val(Text64) + Val(Text65) - 1
        Combo6.AddItem I1
    Next I1
    If Combo6.ListCount > 0 Then Combo6.ListIndex = 0
    Combo5.Clear
    If banben = 0 Then
        njdh = Val(Text64) + Val(Text65) - 2
    Else
        njdh = Val(Text64) + Val(Text65) - 1
    End If
    For I1 = Val(Text64) + 1 To njdh
        Combo5.AddItem I1
    Next I1
    If Combo5.ListCount > 0 Then Combo5.ListIndex = 0
End Sub
Sub 角度转换()
    J1 = CDec(Val(Text53)) * 3.14159265358979 / 180
    弧度化为度分秒 J1
    Text54 = Round(J1, 5)
    J2 = CDec(Val(Text55))
    度分秒化为弧度 J2, J22
    J22 = Round(J22 * 180 / 3.14159265358979, 8)
    Text56 = J22
End Sub




Private Sub Option7_Click()
    Rx = CDec(Val(Text8))
    Ax = CDec(Val(Text7))
    If Ax <> 0 And Rx <> 0 Then
        度分秒化为弧度 F, F1                                                    'F----度.分秒--弧度
        ix = Rx / Abs(Rx)                                                       '判断左右偏的变量，右偏R取正，左偏取负
        Hx = Ax ^ 2 / Abs(Rx)                                                   '完整的缓和曲线总长
    End If
    
    If Option7.Value = True Then
        Label1(134) = "HY点里程KHY"
        '        Label1(26) = "ZH点里程"
        '        Label1(27) = "ZH点X坐标C"
        '        Label1(28) = "ZH点Y坐标D"
        '        Label1(29) = "ZH点切线方位角F"
        Text12 = Round(Val(Text107) - Hx, 3)
        Text13 = Val(Text107)
    Else
        Label1(134) = "YH点里程KYH"
        '        Label1(26) = "HZ点里程"
        '        Label1(27) = "HZ点X坐标C"
        '        Label1(28) = "HZ点Y坐标D"
        '        Label1(29) = "HZ点切线方位角F"
        Text12 = Text107
        Text13 = Round(Val(Text107) + Hx, 3)
    End If
End Sub

Private Sub Option8_Click()
    Rx = CDec(Val(Text8))
    Ax = CDec(Val(Text7))
    If Ax <> 0 And Rx <> 0 Then
        度分秒化为弧度 F, F1                                                    'F----度.分秒--弧度
        ix = Rx / Abs(Rx)                                                       '判断左右偏的变量，右偏R取正，左偏取负
        Hx = Ax ^ 2 / Abs(Rx)                                                   '完整的缓和曲线总长
    End If
    
    If Option7.Value = True Then
        Label1(134) = "HY点里程KHY"
        '        Label1(26) = "ZH点里程"
        '        Label1(27) = "ZH点X坐标C"
        '        Label1(28) = "ZH点Y坐标D"
        '        Label1(29) = "ZH点切线方位角F"
        Text12 = Round(Val(Text107) - Hx, 3)
        Text13 = Val(Text107)
    Else
        Label1(134) = "YH点里程KYH"
        '        Label1(26) = "HZ点里程"
        '        Label1(27) = "HZ点X坐标C"
        '        Label1(28) = "HZ点Y坐标D"
        '        Label1(29) = "HZ点切线方位角F"
        Text12 = Text107
        Text13 = Round(Val(Text107) + Hx, 3)
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    直线2
End Sub

Private Sub Combo8_Click()
    载入宽度参数2
End Sub
Sub 载入宽度参数2()
    Text101 = CDec(Val(kuandu2(Val(Combo8), 0)))                                'kuandu2(段数,0列)存储分段终点里程KI
    Text94 = CDec(Val(kuandu2(Val(Combo8), 1)))                                 'kuandu2(段数,1列)存储变坡点右幅宽度
    If Val(kuandu2(Val(Combo8), 2)) = 0 Then Option6.Value = True
    If Val(kuandu2(Val(Combo8), 2)) = 1 Then Option5.Value = True
End Sub

Private Sub Command26_Click()
    '求特征点里程和点名保存在数组tezheng中，第0列=里程，第1列=点名
    Dim tezheng(2000, 2) As Variant
    Erase tezheng                                                               '清空数组内容，便于多次输入
    ntz = 1
    tezheng(ntz, 0) = Val(TextQD)
    tezheng(ntz, 1) = "QD"
    ntz = ntz + 1
    For nd = 1 To Val(TextPartNumber)
        If factor(nd, 0) = "直线" Then
            tezheng(ntz, 0) = Val(factor(nd, 1))
            If factor(nd + 1, 0) = "圆" Then tezheng(ntz, 1) = "ZY"
            If factor(nd + 1, 0) = "缓1+圆+缓2" Then tezheng(ntz, 1) = "ZH"
            If factor(nd + 1, 0) = "布劳斯曲线" Then tezheng(ntz, 1) = "ZH"
            If factor(nd + 1, 0) = "麦克康纳尔曲线" Then tezheng(ntz, 1) = "ZH"
            If factor(nd + 1, 0) = "" Then tezheng(ntz, 1) = "ZD"
            ntz = ntz + 1
        ElseIf factor(nd, 0) = "圆" Then
            tezheng(ntz, 0) = Val(factor(nd, 1))
            If factor(nd + 1, 0) = "直线" Then tezheng(ntz, 1) = "YZ"
            If factor(nd + 1, 0) = "缓1+圆+缓2" Then tezheng(ntz, 1) = "YH"
            If factor(nd + 1, 0) = "布劳斯曲线" Then tezheng(ntz, 1) = "YH"
            If factor(nd + 1, 0) = "麦克康纳尔曲线" Then tezheng(ntz, 1) = "YH"
            If factor(nd + 1, 0) = "" Then tezheng(ntz, 1) = "ZD"
            ntz = ntz + 1
        ElseIf factor(nd, 0) = "缓1+圆+缓2" Then
            tezheng(ntz, 0) = Val(factor(nd, 2))
            tezheng(ntz, 1) = "HY"
            ntz = ntz + 1
            If factor(nd, 9) <> 0 Then
                tezheng(ntz, 0) = Val(factor(nd, 2)) + Val(factor(nd, 9))
                tezheng(ntz, 1) = "YH"
                ntz = ntz + 1
            End If
            If factor(nd, 1) <> factor(nd, 2) + factor(nd, 9) Then
                tezheng(ntz, 0) = Val(factor(nd, 1))
                If factor(nd + 1, 0) = "直线" Then tezheng(ntz, 1) = "HZ"
                If factor(nd + 1, 0) = "圆" Then tezheng(ntz, 1) = "HY"
                If factor(nd + 1, 0) = "缓1+圆+缓2" And factor(nd, 1) = factor(nd + 1, 2) Then ntz = ntz - 1
                If factor(nd + 1, 0) = "缓1+圆+缓2" And factor(nd, 1) <> factor(nd + 1, 2) Then tezheng(ntz, 1) = "HH"
                If factor(nd + 1, 0) = "布劳斯曲线" Then tezheng(ntz, 1) = "HH"
                If factor(nd + 1, 0) = "麦克康纳尔曲线" Then tezheng(ntz, 1) = "HH"
                ntz = ntz + 1
            End If
            If factor(nd + 1, 0) = "" Then tezheng(ntz - 1, 1) = "ZD"
        ElseIf factor(nd, 0) = "布劳斯曲线" Then
            tezheng(ntz, 0) = Val(factor(nd, 1))
            If factor(nd + 1, 0) = "圆" Then tezheng(ntz, 1) = "HY"
            If factor(nd + 1, 0) = "直线" Then tezheng(ntz, 1) = "HZ"
            If factor(nd + 1, 0) = "" Then tezheng(ntz, 1) = "ZD"
            ntz = ntz + 1
        ElseIf factor(nd, 0) = "麦克康纳尔曲线" Then
            tezheng(ntz, 0) = Val(factor(nd, 1))
            If factor(nd + 1, 0) = "圆" Then tezheng(ntz, 1) = "HY"
            If factor(nd + 1, 0) = "直线" Then tezheng(ntz, 1) = "HZ"
            If factor(nd + 1, 0) = "" Then tezheng(ntz, 1) = "ZD"
            ntz = ntz + 1
        End If
    Next
    
    msg = "输出特征点里程坐标吗？"
    queren = MsgBox(msg, vbYesNoCancel + vbDefaultButton2, "输出特征点确认")
    If queren = vbYes Then
        shuchu = 1
    ElseIf queren = vbNo Then
        shuchu = 0
    Else
        GoTo Label1
    End If
    
    '输出到excel
    On Error Resume Next
    Set appexcel = GetObject(, "Excel.application")
    If appexcel Is Nothing Then Set appexcel = CreateObject("excel.application")
    If Err Then                                                                 '如果没有安装excel，就用wps表格打开
        Set appexcel = GetObject(, "et.application")
        If appexcel Is Nothing Then Set appexcel = CreateObject("et.application")
    End If
    appexcel.Visible = True
    Set exwbook = appexcel.Workbooks.Add
    Set exsheet = exwbook.Worksheets("sheet1")
    exsheet.Activate                                                            '激活
    appexcel.cells.Select
    appexcel.Selection.HorizontalAlignment = xlCenter
    appexcel.Selection.VerticalAlignment = xlCenter
    appexcel.sheets("sheet1").Name = "坐标表"
    appexcel.Range("a2") = " 里程"
    appexcel.Range("b2") = "中桩X坐标"
    appexcel.Range("c2") = "中桩Y坐标"
    appexcel.Range("d2") = "  方位角"
    appexcel.Range("e2") = "边桩X坐标"
    appexcel.Range("f2") = "边桩Y坐标"
appexcel.Columns("A:A").columnwidth = 9.25
appexcel.Columns("B:F").columnwidth = 13
appexcel.Range("A1:f1").Select
    appexcel.Selection.merge
    appexcel.Range("A1") = Text2.text
    appexcel.cells(1, 1).Font.Bold = True
    appexcel.cells(1, 1).Font.Size = 16
    Dim jj As Double
    jj = 3
    ntz4 = 1
    For K = CDec(Val(TEXTKstart)) To CDec(Val(TEXTKend)) Step CDec(Val(TEXToffset))
        T = CDec(Val(TEXTT))
        P = CDec(Val(TEXTP))
        S = CDec(Val(TEXTS))
        Z = CDec(Val(TEXTZ))
        计算 K, x, y, Q, T, P, S, Z, G, H
        On Error GoTo Label1
        If shuchu = 0 Then
            appexcel.cells(jj, 1) = K
            appexcel.cells(jj, 2) = x
            appexcel.cells(jj, 3) = y
            appexcel.cells(jj, 4) = Q
            appexcel.cells(jj, 5) = G
            appexcel.cells(jj, 6) = H
            jj = jj + 1
        ElseIf shuchu = 1 Then
            '比较K和特征点里程谁最小，求出最小值保存在Kmin中，并计算里程Kmin里程的坐标
            If K < tezheng(ntz4, 0) Then
                Kmin = K
                计算 Kmin, x, y, Q, T, P, S, Z, G, H
                appexcel.cells(jj, 1) = K
                appexcel.cells(jj, 2) = x
                appexcel.cells(jj, 3) = y
                appexcel.cells(jj, 4) = Q
                appexcel.cells(jj, 5) = G
                appexcel.cells(jj, 6) = H
                jj = jj + 1
            Else                                                                '如果K大，找出所有小于等于K的所有特征点并输出
                Do While K >= tezheng(ntz4, 0) And ntz4 <= ntz - 1
                    Kmin = Val(tezheng(ntz4, 0))
                    T = CDec(Val(TEXTT))
                    P = CDec(Val(TEXTP))
                    S = CDec(Val(TEXTS))
                    Z = CDec(Val(TEXTZ))
                    计算 Kmin, x, y, Q, T, P, S, Z, G, H
                    appexcel.cells(jj, 1) = tezheng(ntz4, 1) & tezheng(ntz4, 0)
                    appexcel.cells(jj, 2) = x
                    appexcel.cells(jj, 3) = y
                    appexcel.cells(jj, 4) = Q
                    appexcel.cells(jj, 5) = G
                    appexcel.cells(jj, 6) = H
                    appexcel.Range(appexcel.cells(jj, 1), appexcel.cells(jj, 6)).Font.Bold = True
                    ntz4 = ntz4 + 1
                    jj = jj + 1
                Loop
                If K <> tezheng(ntz4 - 1, 0) Then                               '输出非特征点，不能与已输出的特征点重复
                    T = CDec(Val(TEXTT))
                    P = CDec(Val(TEXTP))
                    S = CDec(Val(TEXTS))
                    Z = CDec(Val(TEXTZ))
                    计算 K, x, y, Q, T, P, S, Z, G, H
                    appexcel.cells(jj, 1) = K
                    appexcel.cells(jj, 2) = x
                    appexcel.cells(jj, 3) = y
                    appexcel.cells(jj, 4) = Q
                    appexcel.cells(jj, 5) = G
                    appexcel.cells(jj, 6) = H
                    jj = jj + 1
                End If
            End If
        End If
    Next K
    
appexcel.Range("A2:F" + LTrim(Str(jj - 1))).Select
    appexcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    appexcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With appexcel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    appexcel.cells(1, 7).Select
    
    '设置顶端标题行和页脚的第几页共几页
    With appexcel.activesheet.PageSetup
        .PrintTitleRows = "$1:$2"
        '        .PrintTitleColumns = ""
    End With
    '    appexcel.ActiveSheet.PageSetup.PrintArea = ""
    With appexcel.activesheet.PageSetup
        '        .LeftHeader = ""
        '        .CenterHeader = ""
        '        .RightHeader = ""
        '        .LeftFooter = ""
        .CenterFooter = "第 &P 页，共 &N 页"
        '        .RightFooter = ""
        .LeftMargin = appexcel.application.InchesToPoints(0.748031496062992)
        .RightMargin = appexcel.application.InchesToPoints(0.748031496062992)
        .TopMargin = appexcel.application.InchesToPoints(0.866141732283465)
        .BottomMargin = appexcel.application.InchesToPoints(0.866141732283465)
        .HeaderMargin = appexcel.application.InchesToPoints(0.511811023622047)
        .FooterMargin = appexcel.application.InchesToPoints(0.511811023622047)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
    End With
Label1:
End Sub

Private Sub Command27_Click()
    On Error Resume Next
    Set appexcel = GetObject(, "Excel.application")
    If appexcel Is Nothing Then Set appexcel = CreateObject("excel.application")
    If Err Then                                                                 '如果没有安装excel，就用wps表格打开
        Set appexcel = GetObject(, "et.application")
        If appexcel Is Nothing Then Set appexcel = CreateObject("et.application")
    End If
    appexcel.Visible = True
    Set exwbook = appexcel.Workbooks.Add
    Set exsheet = exwbook.Worksheets("sheet1")
    exsheet.Activate                                                            '激活
    appexcel.cells.Select
    appexcel.Selection.HorizontalAlignment = xlCenter
    appexcel.Selection.VerticalAlignment = xlCenter
    appexcel.sheets("sheet1").Name = "坐标表"
    appexcel.Range("a2") = " 里程"
    appexcel.Range("b2") = "中桩X坐标"
    appexcel.Range("c2") = "中桩Y坐标"
    appexcel.Range("d2") = "  方位角"
    appexcel.Range("e2") = "边桩X坐标"
    appexcel.Range("f2") = "边桩Y坐标"
appexcel.Columns("A:A").columnwidth = 6
appexcel.Columns("B:F").columnwidth = 13
appexcel.Range("A1:f1").Select
    appexcel.Selection.merge
    appexcel.Range("a1") = Text21.text
    appexcel.cells(1, 1).Font.Bold = True
    appexcel.cells(1, 1).Font.Size = 16
    Dim jj As Double
    jj = 3
    For K = CDec(Val(Text12)) To CDec(Val(Text13)) Step CDec(Val(Text14))
        缓和曲线 K, x, y, Q, G, H
        On Error GoTo Label1
        appexcel.cells(jj, 1) = K
        appexcel.cells(jj, 2) = x
        appexcel.cells(jj, 3) = y
        appexcel.cells(jj, 4) = Q
        appexcel.cells(jj, 5) = G
        appexcel.cells(jj, 6) = H
        jj = jj + 1
    Next K
    
appexcel.Range("A2:F" + LTrim(Str(jj - 1))).Select
    appexcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    appexcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With appexcel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    appexcel.cells(1, 10).Select
    
    '设置顶端标题行和页脚的第几页共几页
    'With appexcel.ActiveSheet.PageSetup
    '        .PrintTitleRows = "$1:$3"
    '        .PrintTitleColumns = ""
    '    End With
    '    appexcel.ActiveSheet.PageSetup.PrintArea = ""
    '    With appexcel.ActiveSheet.PageSetup
    '        .LeftHeader = ""
    '        .CenterHeader = ""
    '        .RightHeader = ""
    '        .LeftFooter = ""
    '        .CenterFooter = "第 &P 页，共 &N 页"
    '        .RightFooter = ""
    '        .LeftMargin = appexcel.Application.InchesToPoints(0.748031496062992)
    '        .RightMargin = appexcel.Application.InchesToPoints(0.748031496062992)
    '        .TopMargin = appexcel.Application.InchesToPoints(0.866141732283465)
    '        .BottomMargin = appexcel.Application.InchesToPoints(0.866141732283465)
    '        .HeaderMargin = appexcel.Application.InchesToPoints(0.511811023622047)
    '        .FooterMargin = appexcel.Application.InchesToPoints(0.511811023622047)
    '        .PrintHeadings = False
    '        .PrintGridlines = False
    '        .PrintComments = xlPrintNoComments
    '        .CenterHorizontally = True
    '        .CenterVertically = False
    '        .Orientation = xlPortrait
    '        .Draft = False
    '        .PaperSize = xlPaperA4
    '        .FirstPageNumber = xlAutomatic
    '        .Order = xlDownThenOver
    '        .BlackAndWhite = False
    '        .Zoom = 100
    '        .PrintErrors = xlPrintErrorsDisplayed
    '    End With
Label1:
End Sub

Private Sub Command28_Click()
    On Error Resume Next
    Set appexcel = GetObject(, "Excel.application")
    If appexcel Is Nothing Then Set appexcel = CreateObject("excel.application")
    If Err Then                                                                 '如果没有安装excel，就用wps表格打开
        Set appexcel = GetObject(, "et.application")
        If appexcel Is Nothing Then Set appexcel = CreateObject("et.application")
    End If
    appexcel.Visible = True
    Set wbmybook = appexcel.Workbooks.Add
    Set exsheet = wbmybook.Worksheets("sheet1")
    exsheet.Activate                                                            '激活
    appexcel.cells.Select
    appexcel.Selection.HorizontalAlignment = xlCenter
    appexcel.Selection.VerticalAlignment = xlCenter
    appexcel.sheets("sheet1").Name = "标高表"
    appexcel.Range("A2") = "中桩里程"
    appexcel.Range("B2") = "偏距"
    appexcel.Range("C2") = "偏角"
    appexcel.Range("D2") = "边桩里程"
    appexcel.Range("E2") = "设计高程"
    appexcel.Range("F2") = "左幅横坡"
    appexcel.Range("G2") = "右幅横坡"
    appexcel.Range("H2") = "左边桩高程"
    appexcel.Range("I2") = "右边桩高程"
appexcel.Columns("A:A").columnwidth = 9
appexcel.Columns("B:B").columnwidth = 7
appexcel.Columns("C:C").columnwidth = 4.5
appexcel.Columns("D:D").columnwidth = 9
appexcel.Columns("E:G").columnwidth = 8
appexcel.Columns("H:I").columnwidth = 10
appexcel.Columns("J:J").columnwidth = 3
'appexcel.Columns("K:L").columnwidth = 85
appexcel.Range("A1:I1").Select
    appexcel.Selection.merge
    appexcel.Range("A1") = Text41.text
    appexcel.cells(1, 1).Font.Bold = True
    appexcel.cells(1, 1).Font.Size = 16
    Dim jj As Double
    jj = 3
    On Error GoTo Label1
    For K = CDec(Val(Text37)) To CDec(Val(Text38)) Step CDec(Val(Text39))
        计算设计高程 K, J, y, R, E, F, H
        T = CDec(-Abs(Val(Text35)))
        P = CDec(Val(Text44))
        计算边桩高程 K, T, P, KG, H, Iz, Gz
        T = CDec(Abs(Val(Text35)))
        P = CDec(Val(Text44))
        计算边桩高程 K, T, P, KG, H, Iy, Gy
        appexcel.cells(jj, 1) = K
        appexcel.cells(jj, 2) = CDec(Val(Text35))
        appexcel.cells(jj, 3) = P
        appexcel.cells(jj, 4) = KG
        appexcel.cells(jj, 5) = H
        appexcel.cells(jj, 6) = Iz
        appexcel.cells(jj, 7) = Iy
        appexcel.cells(jj, 8) = Gz
        appexcel.cells(jj, 9) = Gy
        jj = jj + 1
    Next K
    
    '设置页面边框
    'appexcel.Range(appexcel.cells(2, 1), appexcel.cells(jj - 1, 8)).Select
appexcel.Range("A2:I" + LTrim(Str(jj - 1))).Select
    appexcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    appexcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With appexcel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    appexcel.cells(1, 11).Select
    '设置打印区域
    appexcel.activesheet.PageSetup.PrintArea = "$A:$I"
    
    
    '设置顶端标题行和页脚的第几页共几页
    'With appexcel.ActiveSheet.PageSetup
    '        .PrintTitleRows = "$1:$2"
    '        .PrintTitleColumns = ""
    '    End With
    '    appexcel.ActiveSheet.PageSetup.PrintArea = ""
    '    With appexcel.ActiveSheet.PageSetup
    '        .LeftHeader = ""
    '        .CenterHeader = ""
    '        .RightHeader = ""
    '        .LeftFooter = ""
    '        .CenterFooter = "第 &P 页，共 &N 页"
    '        .RightFooter = ""
    '        .LeftMargin = appexcel.Application.InchesToPoints(0.748031496062992)
    '        .RightMargin = appexcel.Application.InchesToPoints(0.748031496062992)
    '        .TopMargin = appexcel.Application.InchesToPoints(0.866141732283465)
    '        .BottomMargin = appexcel.Application.InchesToPoints(0.866141732283465)
    '        .HeaderMargin = appexcel.Application.InchesToPoints(0.511811023622047)
    '        .FooterMargin = appexcel.Application.InchesToPoints(0.511811023622047)
    '        .PrintHeadings = False
    '        .PrintGridlines = False
    '        .PrintComments = xlPrintNoComments
    '        .CenterHorizontally = True
    '        .CenterVertically = False
    '        .Orientation = xlPortrait
    '        .Draft = False
    '        .PaperSize = xlPaperA4
    '        .FirstPageNumber = xlAutomatic
    '        .Order = xlDownThenOver
    '        .BlackAndWhite = False
    '        .Zoom = 100
    '        .PrintErrors = xlPrintErrorsDisplayed
    '    End With
    
Label1: Close
End Sub

Private Sub Command29_Click()
    Shell "calc.exe", vbNormalFocus
End Sub

'----------------------------------------------------------------------------------------------------------
Private Sub Command3_Click()
    Shell "notepad.exe " & """" & App.Path + "\更新记录.txt" & """", vbMaximizedFocus
End Sub

Private Sub Command30_Click()
    Text57 = TextX2
    Text58 = TextY2
End Sub

Private Sub Command31_Click()
    Text60 = Text22
End Sub

Private Sub Command32_Click()
    With CommonDialog1
        .FileName = ""
        .CancelError = False
        .DialogTitle = "从文件载入交点参数..."
        .DefaultExt = "sct"
        .Filter = "输入的交点参数(*.sct)|*.sct"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text1 = .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 1
    
    Dim kk() As String
    Dim str1 As String
    Line Input #1, str1
    kk = Split(str1, ",")
    If UBound(kk) = 0 Then GoTo Label1
    
    banben = 0                                                                  '新版本格式：
    njd = Val(kk(0))
    Text65 = njd
    njdh = Val(kk(1))
    Text64 = njdh
    Erase factor2                                                               '清空交点参数数组便于多次载入
    For n1 = njdh To njdh + njd - 1
        Line Input #1, str1
        kk = Split(str1, ",")
        If UBound(kk) <= 5 Then
            factor2(n1, 0) = Val(kk(0))
            factor2(n1, 1) = Val(kk(1))
            factor2(n1, 2) = Val(kk(2))
            factor2(n1, 3) = Val(kk(3))
            factor2(n1, 4) = Val(kk(4))
            factor2(n1, 5) = Val(kk(5))
        Else
            factor2(n1, 0) = Val(kk(0))
            factor2(n1, 1) = Val(kk(1))
            factor2(n1, 2) = Val(kk(2))
            factor2(n1, 3) = Val(kk(3))
            factor2(n1, 4) = Val(kk(4))
            factor2(n1, 5) = Val(kk(5))
            factor2(n1, 7) = Val(kk(6))
            factor2(n1, 8) = Val(kk(7))
            factor2(n1, 9) = Val(kk(8))
        End If
        
        '  Input #1, factor2(n1, 0), factor2(n1, 1), factor2(n1, 2), factor2(n1, 3), factor2(n1, 4), factor2(n1, 5) _
        '  , factor2(n1, 7), factor2(n1, 8), factor2(n1, 9)
        'Kjd，X,Y,R,Ls1，Ls2,曲线长，类型，A1,A2
    Next n1
    Close #1
    Saveflag = 1
    载入交点参数
    载入序号
    Exit Sub
    
Label1:                                                                         '老版本格式：
    banben = 1
    njd = Val(kk(0))
    Text65 = njd
    Input #1, njdh
    Text64 = njdh
    '清空交点参数数组便于多次载入
    Erase factor2
    For n1 = njdh To njdh + njd - 1
        Input #1, n1
        For n2 = 0 To 6
            Input #1, factor2(n1, n2)
        Next n2
    Next n1
    Close #1
    Saveflag = 1
    载入交点参数
    载入序号
End Sub

Private Sub Command33_Click()
    Dim n1, n2 As Integer                                                       '循环用的变量
    With CommonDialog1
        .FileName = ""
        .CancelError = False
        .DialogTitle = "保存交点参数到文件..."
        .DefaultExt = "sct"
        .Filter = "输出的交点参数(*.sct)|*.sct"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text1 = .FileTitle
    End With
    Open CommonDialog1.FileName For Output As 2
    
    If banben = 0 Then                                                          '如果是新版本,默认值
        Write #2, Val(Text65), Val(Text64)
        For n1 = Val(Text64) To Val(Text64) + Val(Text65) - 1
            Write #2, factor2(n1, 0), factor2(n1, 1), factor2(n1, 2), Val(factor2(n1, 3)), Val(factor2(n1, 4)) _
            , Val(factor2(n1, 5)), Val(factor2(n1, 7)), Val(factor2(n1, 8)), Val(factor2(n1, 9))
            'Kjd，X,Y,R,Ls1，Ls2,类型，A1，A2
        Next n1
    End If
    
    If banben = 1 Then                                                          '如果是老版本
        If MsgBox("交点参数建议转换为新版本，现在转换吗？", vbQuestion + vbYesNo + vbDefaultButton1, "交点参数版本转换") _
            = vbYes Then                                                        '转换为新版本
            njdh = Val(Text65) + Val(Text64) - 1
            If Val(factor2(njdh, 3)) = 0 Then
                njd = Val(Text65)
            Else
                njd = Val(Text65) + 1
            End If
            Write #2, njd, Val(Text64)
            For n1 = Val(Text64) To Val(Text64) + Val(Text65) - 1
                Write #2, factor2(n1, 0), factor2(n1, 1), factor2(n1, 2), factor2(n1, 3), factor2(n1, 4), factor2(n1, 4) _
                , Val(factor2(n1, 7)), Val(factor2(n1, 8)), Val(factor2(n1, 9))
                'Kjd，X,Y,R,Ls1，Ls2,类型，A1，A2
            Next n1
            K1 = factor2(njdh - 1, 0)
            x1 = factor2(njdh - 1, 1)
            y1 = factor2(njdh - 1, 2)
            K2 = factor2(njdh, 0)
            x2 = factor2(njdh, 1)
            y2 = factor2(njdh, 2)
            R = CDec(Val(factor2(njdh, 3)))
            Ls1 = CDec(Val(factor2(njdh, 4)))
            T1 = CDec(Val(factor2(njdh, 5)))
            Lz = CDec(Val(factor2(njdh, 6)))
            
            If x2 <> x1 And y2 <> y1 Then
                Q = Atn((y2 - y1) / (x2 - x1)) + (1 - Abs(x2 - x1) / (x2 - x1)) * pai / 2 + (1 + Abs(x2 - x1) / (x2 - x1)) * (1 - Abs(y2 - y1) / (y2 - y1)) * pai / 2
                弧度化为度分秒 Q
            ElseIf x2 = x1 Then                                                 '如果X2=X1 会被零除，要单独计算
                If y2 > y1 Then Q = 90
                If y2 < y1 Then Q = 270
            ElseIf y2 = y1 Then                                                 '如果Y2=Y1 会被零除，要单独计算
                If x2 > x1 Then Q = 0
                If x2 < x1 Then Q = 180
            End If
            度分秒化为弧度 Q, Qh                                                '首先计算P1到P2的方位角
            If R <> 0 Then ap2 = (Lz - Ls1) / R                                 '弧度
            K3 = K2 - T1 + Lz
            X3 = Round(x2 + T1 * Cos(Qh + ap2), 3)
            Y3 = Round(y2 + T1 * Sin(Qh + ap2), 3)
            Write #2, K3, X3, Y3, 0, 0, 0
            banben = 0
            '再次载入参数
            Close #2
            Open CommonDialog1.FileName For Input As #16
            Dim kk() As String
            Dim str1 As String
            Line Input #16, str1
            kk = Split(str1, ",")
            '新版本格式：
            njd = Val(kk(0))
            Text65 = njd
            njdh = Val(kk(1))
            Text64 = njdh
            Erase factor2                                                       '清空交点参数数组便于多次载入
            For n1 = njdh To njdh + njd - 1
                Input #16, factor2(n1, 0), factor2(n1, 1), factor2(n1, 2), factor2(n1, 3), factor2(n1, 4), factor2(n1, 5)
                'Kjd，X,Y,R,Ls1，Ls2
            Next n1
            Close #16
            Saveflag = 1
            载入交点参数
            载入序号
            
        Else                                                                    '如果不转换，就直接保存为老版本
            Write #2, Val(Text65)
            Write #2, Val(Text64)
            For n1 = Val(Text64) To Val(Text64) + Val(Text65) - 1
                Write #2, n1
                For n2 = 0 To 6
                    Print #2, factor2(n1, n2)
                Next n2
            Next n1
        End If
    End If
    Close #2
    Saveflag = 1
End Sub
Sub 计算方位角(x1, y1, x2, y2, ss, Qh)
    ss = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
    If x2 <> x1 And y2 <> y1 Then
        Q = Atn((y2 - y1) / (x2 - x1)) + (1 - Abs(x2 - x1) / (x2 - x1)) * pai / 2 + (1 + Abs(x2 - x1) / (x2 - x1)) * (1 - Abs(y2 - y1) / (y2 - y1)) * pai / 2
        弧度化为度分秒 Q
    ElseIf x2 = x1 Then                                                         '如果X2=X1 会被零除，要单独计算
        If y2 > y1 Then Q = 90
        If y2 < y1 Then Q = 270
    ElseIf y2 = y1 Then                                                         '如果Y2=Y1 会被零除，要单独计算
        If x2 > x1 Then Q = 0
        If x2 < x1 Then Q = 180
    End If
    度分秒化为弧度 Q, Qh                                                        '首先计算P1到P2的方位角
End Sub
Private Sub Command59_Click()
    For I1 = Val(Text64) + 1 To Val(Text64) + Val(Text65) - 2
        K1 = CDec(Val(factor2(I1 - 1, 0)))
        x1 = CDec(Val(factor2(I1 - 1, 1)))
        y1 = CDec(Val(factor2(I1 - 1, 2)))
        K2 = CDec(Val(factor2(I1, 0)))
        x2 = CDec(Val(factor2(I1, 1)))
        y2 = CDec(Val(factor2(I1, 2)))
        x33 = CDec(Val(factor2(I1 + 1, 1)))
        y33 = CDec(Val(factor2(I1 + 1, 2)))
        计算方位角 x1, y1, x2, y2, ss1, Qh1
        计算方位角 x2, y2, x33, y33, ss2, qh2
        If x33 = 0 And y33 = 0 Then
            MsgBox "输入交点" & I1 + 1 & "的坐标之后才能判断左右偏。建议输入全部交点坐标之后再判断。", vbOKOnly, "检验左右偏"
            Exit Sub
        End If
        If Sin(qh2 - Qh1) > 0 And Val(factor2(I1, 3)) < 0 Or Sin(qh2 - Qh1) < 0 And Val(factor2(I1, 3)) > 0 Then
            '检验方位角之差的正弦值判断左右偏
            If Combo6.ListCount > 0 Then Combo6.ListIndex = I1 - Val(Text64)
            msg = "交点" & I1 & "的半径R应为" & -factor2(I1, 3) & "，现在修改吗？"
            queren = MsgBox(msg, vbYesNo + vbDefaultButton1 + vbCritical, "检验左右偏")
            If queren = vbYes Then
                Text66 = -factor2(I1, 3)
                factor2(I1, 3) = -factor2(I1, 3)
                With Text66
                    .SetFocus                                                   '让text获得焦点
                    .SelStart = 0
                    .SelLength = Len(.text)
                End With
            End If
            Exit For
        End If
    Next
    If I1 = Val(Text64) + Val(Text65) - 1 Then
        MsgBox "没有发现半径R的错误。", vbInformation + vbOKOnly, "检验半径R正负号"
    End If
End Sub

Private Sub Command60_Click()
    I1 = Val(Combo6) - 1
    If I1 >= Val(Text64) And I1 < Val(Text64) + Val(Text65) - 1 Then
        计算交点里程 I1, K33
        Text73 = Round(K33, 3)
        'Else
        'MsgBox "输入XY坐标后才能计算交点" & Val(Text64) + 2 & "-" & Val(Text64) + Val(Text65) - 1 & _
        '"的里程。建议输入所有交点坐标后从交点" & Val(Text64) + 2 & "开始依次推算。", vbInformation + vbOKOnly, "推算里程"
    End If
End Sub
Sub 计算交点里程(I1, K33)
    '先预设版本=1
    K1 = CDec(Val(factor2(I1 - 1, 0)))
    x1 = CDec(Val(factor2(I1 - 1, 1)))
    y1 = CDec(Val(factor2(I1 - 1, 2)))
    K2 = CDec(Val(factor2(I1, 0)))
    x2 = CDec(Val(factor2(I1, 1)))
    y2 = CDec(Val(factor2(I1, 2)))
    x33 = CDec(Val(Text72))
    y33 = CDec(Val(Text71))
    R = CDec(Val(factor2(I1, 3)))
    Ls1 = CDec(Val(factor2(I1, 4)))
    Ls2 = Ls1
    T1 = CDec(Val(factor2(I1, 5)))
    T2 = T1
    Lz = CDec(Val(factor2(I1, 6)))
    If R <> 0 Then ap = (Lz - Ls1) / R * 180 / pai
    apD = Fix(ap)                                                               '度
    apM = (ap - Fix(ap)) * 60                                                   '分
    apS = (apM - Fix(apM)) * 60                                                 '秒
    ap = apD + Fix(apM) / 100 + apS / 10000                                     '度.分秒
    aa = Sqr((x1 - x33) ^ 2 + (y1 - y33) ^ 2)                                   '根号
    bb = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    cc = Sqr((x2 - x33) ^ 2 + (y2 - y33) ^ 2)
    
    If banben = 0 Then                                                          '如果是新版本
        Ls2 = CDec(Val(factor2(I1, 5)))
        'If R <> 0 Then
        'b1 = Ls1 / 2 / Abs(R)
        'b2 = Ls2 / 2 / Abs(R)
        'p1 = Ls1 ^ 2 / 24 / Abs(R) - Ls1 ^ 4 / 2688/ Abs(R ^ 3)
        'p2 = Ls2 ^ 2 / 24 / Abs(R) - Ls2 ^ 4 / 2688 / Abs(R ^ 3)
        'm1 = Ls1 / 2 - Ls1 ^ 3 / (240 * R ^ 2) + Ls1 ^ 5 / (34560 * R ^ 4)
        'm2 = Ls2 / 2 - Ls2 ^ 3 / (240 * R ^ 2) + Ls2 ^ 5 / (34560 * R ^ 4)
        'End If
        
        If Val(factor2(I1, 7)) = 0 Then                                         '如果没有被选中就是完全缓和曲线
            A1 = Sqr(Abs(R * Ls1))
            A2 = Sqr(Abs(R * Ls2))
        Else                                                                    '如果选中就是不完全缓和曲线
            A1 = Val(factor2(I1, 8))
            A2 = Val(factor2(I1, 9))
        End If
        计算bpm R, Rq, Ls1, A1, b1, p1, m1
        计算bpm R, Rz, Ls2, A2, b2, p2, m2
        If bb * cc <> 0 And x33 <> 0 Then ap = pai - arccos((bb * bb + cc * cc - aa * aa) / (2 * bb * cc))
        If Sin(ap) <> 0 And Tan(ap) <> 0 Then
            T1 = Round((Abs(R) + p2) / Sin(ap) - (Abs(R) + p1) / Tan(ap) + m1, 3)
            T2 = Round((Abs(R) + p1) / Sin(ap) - (Abs(R) + p2) / Tan(ap) + m2, 3)
        End If
        Lz = Round(Ls1 + Ls2 + (ap - b1 - b2) * Abs(R), 3)                      'ap,b1,b2单位均为弧度
        弧度化为度分秒 ap
    End If
    qq = T1 + T2 - Lz
    If I1 = Val(Text64) Then qq = 0
    K33 = Round(K2 + cc - qq, 3)                                                '交点的里程推算
End Sub
Function arccos(x)
    arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function
Sub 计算bpm(R, Rq, Ls1, A1, b1, p1, m1)
    If A1 > 0 Then                                                              '如果半径向圆曲线变小
        Rq = 0
        Hs1 = A1 ^ 2 / Abs(R)
        Lsx = Hs1 - Ls1
        b1 = (Hs1 ^ 2 - Lsx ^ 2) / 2 / A1 ^ 2
        y0 = (Hs1 ^ 3 - Lsx ^ 3) / 6 / A1 ^ 2 - (Hs1 ^ 7 - Lsx ^ 7) / 336 / A1 ^ 6 + _
        (Hs1 ^ 11 - Lsx ^ 11) / 42240 / A1 ^ 10 - (Hs1 ^ 15 - Lsx ^ 15) / 9676800 / A1 ^ 14
        x0 = (Hs1 - Lsx) - (Hs1 ^ 5 - Lsx ^ 5) / 40 / A1 ^ 4 + (Hs1 ^ 9 - Lsx ^ 9) / 3456 / A1 ^ 8 _
        - (Hs1 ^ 13 - Lsx ^ 13) / 599040 / A1 ^ 12
        If x0 <> 0 Then alpha1 = Atn(y0 / x0) - Lsx ^ 2 / 2 / A1 ^ 2
        x01 = Sqr(x0 ^ 2 + y0 ^ 2) * Cos(alpha1)
        y01 = Sqr(x0 ^ 2 + y0 ^ 2) * Sin(alpha1)
        p1 = y01 - Abs(R) * (1 - Cos(b1))
        m1 = x01 - Abs(R) * Sin(b1)
    ElseIf A1 < 0 Then                                                          '如果半径向圆曲线变大
        Rq = A1 ^ 2 * R / (A1 ^ 2 + Ls1 * Abs(R))                               '起点半径
        Hs1 = A1 ^ 2 / Abs(Rq)
        Lsx = Hs1 - Ls1
        b1 = Abs(Hs1 ^ 2 - Lsx ^ 2) / 2 / A1 ^ 2
        y0 = (Hs1 ^ 3 - Lsx ^ 3) / 6 / A1 ^ 2 - (Hs1 ^ 7 - Lsx ^ 7) / 336 / A1 ^ 6 + _
        (Hs1 ^ 11 - Lsx ^ 11) / 42240 / A1 ^ 10 - (Hs1 ^ 15 - Lsx ^ 15) / 9676800 / A1 ^ 14
        x0 = (Hs1 - Lsx) - (Hs1 ^ 5 - Lsx ^ 5) / 40 / A1 ^ 4 + (Hs1 ^ 9 - Lsx ^ 9) / 3456 / A1 ^ 8 _
        - (Hs1 ^ 13 - Lsx ^ 13) / 599040 / A1 ^ 12
        If x0 <> 0 Then alpha1 = Hs1 ^ 2 / 2 / A1 ^ 2 - Atn(y0 / x0)
        x01 = Sqr(x0 ^ 2 + y0 ^ 2) * Cos(alpha1)
        y01 = Sqr(x0 ^ 2 + y0 ^ 2) * Sin(alpha1)
        p1 = y01 - Abs(R) * (1 - Cos(b1))
        m1 = x01 - Abs(R) * Sin(b1)
    Else
        b1 = 0
        p1 = 0
        m1 = 0
    End If
End Sub
Sub 计算特征点(I1, R, Ls1, Ls2, A1, A2, T1, T2, Lz, ap, K3, X3, Y3, Q3, K4, X4, Y4, Q4, K5, X5, Y5, Q5, K6, X6, Y6, Q6)
    '计算factor2的I1分段特征点参数，factor2数组是全局数组
    K1 = CDec(Val(factor2(I1 - 1, 0)))
    x1 = CDec(Val(factor2(I1 - 1, 1)))
    y1 = CDec(Val(factor2(I1 - 1, 2)))
    K2 = CDec(Val(factor2(I1, 0)))
    x2 = CDec(Val(factor2(I1, 1)))
    y2 = CDec(Val(factor2(I1, 2)))
    x33 = CDec(Val(factor2(I1 + 1, 1)))
    y33 = CDec(Val(factor2(I1 + 1, 2)))
    R = CDec(Val(factor2(I1, 3)))
    Ls1 = CDec(Val(factor2(I1, 4)))
    Ls2 = Ls1
    T1 = CDec(Val(factor2(I1, 5)))
    T2 = T1
    Lz = CDec(Val(factor2(I1, 6)))
    If R <> 0 Then ap = (Lz - Ls1) / R * 180 / pai
    apD = Fix(ap)                                                               '度
    apM = (ap - Fix(ap)) * 60                                                   '分
    apS = (apM - Fix(apM)) * 60                                                 '秒
    ap = Round(apD + Fix(apM) / 100 + apS / 10000, 5)                           '度.分秒
    
    If banben = 0 Then                                                          '如果是新版本
        Ls2 = CDec(Val(factor2(I1, 5)))
        'If R <> 0 Then
        'b1 = Ls1 / 2 / Abs(R)
        'b2 = Ls2 / 2 / Abs(R)
        'p1 = Ls1 ^ 2 / 24 / Abs(R) - Ls1 ^ 4 / 2688/ Abs(R ^ 3)
        'p2 = Ls2 ^ 2 / 24 / Abs(R) - Ls2 ^ 4 / 2688 / Abs(R ^ 3)
        'm1 = Ls1 / 2 - Ls1 ^ 3 / (240 * R ^ 2) + Ls1 ^ 5 / (34560 * R ^ 4)
        'm2 = Ls2 / 2 - Ls2 ^ 3 / (240 * R ^ 2) + Ls2 ^ 5 / (34560 * R ^ 4)
        'End If
        If Val(factor2(I1, 7)) = 0 Then                                         '如果没有被选中就是完全缓和曲线
            A1 = Sqr(Abs(R * Ls1))
            A2 = Sqr(Abs(R * Ls2))
        Else                                                                    '如果选中就是不完全缓和曲线
            A1 = Val(factor2(I1, 8))
            A2 = Val(factor2(I1, 9))
        End If
        计算bpm R, Rq, Ls1, A1, b1, p1, m1
        计算bpm R, Rz, Ls2, A2, b2, p2, m2
        aa = Sqr((x1 - x33) ^ 2 + (y1 - y33) ^ 2)                               '根号
        bb = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
        cc = Sqr((x2 - x33) ^ 2 + (y2 - y33) ^ 2)
        If bb * cc <> 0 And x33 <> 0 Then ap = pai - arccos((bb * bb + cc * cc - aa * aa) / (2 * bb * cc))
        If Sin(ap) <> 0 And Tan(ap) <> 0 Then
            T1 = Round((Abs(R) + p2) / Sin(ap) - (Abs(R) + p1) / Tan(ap) + m1, 3)
            T2 = Round((Abs(R) + p1) / Sin(ap) - (Abs(R) + p2) / Tan(ap) + m2, 3)
        End If
        Lz = Round(Ls1 + Ls2 + (ap - b1 - b2) * Abs(R), 3)                      'ap,b1,b2单位均为弧度
        弧度化为度分秒 ap
    End If
    
    '先计算ZH（ZY）点：
    'pai = 3.14159265358979
    If x2 <> x1 And y2 <> y1 Then
        Q = Atn((y2 - y1) / (x2 - x1)) + (1 - Abs(x2 - x1) / (x2 - x1)) * pai / 2 + _
        (1 + Abs(x2 - x1) / (x2 - x1)) * (1 - Abs(y2 - y1) / (y2 - y1)) * pai / 2
        弧度化为度分秒 Q
    ElseIf x2 = x1 Then                                                         '如果X2=X1 会被零除，要单独计算
        If y2 > y1 Then Q = 90
        If y2 < y1 Then Q = 270
    ElseIf y2 = y1 Then                                                         '如果Y2=Y1 会被零除，要单独计算
        If x2 > x1 Then Q = 0
        If x2 < x1 Then Q = 180
    End If
    度分秒化为弧度 Q, Qh                                                        '首先计算P1到P2的方位角
    K3 = K2 - T1
    X3 = x2 - T1 * Cos(Qh)
    Y3 = y2 - T1 * Sin(Qh)
    Q3 = Q                                                                      '单位：度.分秒
    
    '计算HY点：
    If A1 <> 0 And R <> 0 Then
        F1 = Qh
        W = A1 / Abs(A1)                                                        'W=1顺向计算，W=-1反向计算
        i = R / Abs(R)                                                          '判断左右偏的变量，右偏R取正，左偏取负
        If A1 > 0 Then
            Hs1 = A1 ^ 2 / Abs(R)
            Lsx = Hs1 - Ls1
            E = F1 - Lsx ^ 2 * i / (2 * A1 ^ 2)                                 '弧度
        Else
            Hs1 = A1 ^ 2 / Abs(Rq)
            Lsx = Hs1 - Ls1
            E = F1 + Hs1 ^ 2 * i / (2 * A1 ^ 2)                                 '弧度
        End If
        U = (Hs1 - Lsx) - (Hs1 ^ 5 - Lsx ^ 5) / 40 / A1 ^ 4 + (Hs1 ^ 9 - Lsx ^ 9) / 3456 / A1 ^ 8 _
        - (Hs1 ^ 13 - Lsx ^ 13) / 599040 / A1 ^ 12
        U = U * W
        V = (Hs1 ^ 3 - Lsx ^ 3) / 6 / A1 ^ 2 - (Hs1 ^ 7 - Lsx ^ 7) / 336 / A1 ^ 6 + _
        (Hs1 ^ 11 - Lsx ^ 11) / 42240 / A1 ^ 10 - (Hs1 ^ 15 - Lsx ^ 15) / 9676800 / A1 ^ 14
        V = V * W
        K4 = K3 + Ls1
        X4 = X3 + W * U * Cos(E) - i * V * Sin(E)
        Y4 = Y3 + W * U * Sin(E) + i * V * Cos(E)
        Q4 = F1 + (Hs1 ^ 2 - Lsx ^ 2) * i / (2 * A1 ^ 2)                        '切线方位角，单位弧度
        弧度化为度分秒 Q4
    Else
        K4 = K3
        X4 = X3
        Y4 = Y3
        Q4 = Q3                                                                 '单位：度.分秒
    End If
    
    '计算YH点：
    If R <> 0 Then
        度分秒化为弧度 Q4, Q4h                                                  'F----度.分秒--弧度
        L = Lz - Ls1 - Ls2
        If banben = 0 Then L = Lz - Ls1 - Ls2                                   '如果是新版本
        W = Q4h + L / (2 * R)                                                   'W----弧度
        i = 2 * R * Sin(L / (2 * R))                                            'I为弦长
        K5 = K4 + L
        X5 = X4 + i * Cos(W)
        Y5 = Y4 + i * Sin(W)
        Q5 = Q4h + L / R                                                        'Q----弧度
        弧度化为度分秒 Q5
    End If
    
    '计算HZ点：
    
    If x2 <> x33 And y2 <> y33 Then
        Q = Atn((y33 - y2) / (x33 - x2)) + (1 - Abs(x33 - x2) / (x33 - x2)) * pai / 2 + _
        (1 + Abs(x33 - x2) / (x33 - x2)) * (1 - Abs(y33 - y2) / (y33 - y2)) * pai / 2
        弧度化为度分秒 Q
    ElseIf x2 = x33 Then                                                        '如果X2=X1 会被零除，要单独计算
        If y2 > y33 Then Q = 90
        If y2 < y33 Then Q = 270
    ElseIf y2 = y33 Then                                                        '如果Y2=Y1 会被零除，要单独计算
        If x2 > x33 Then Q = 0
        If x2 < x33 Then Q = 180
    End If
    度分秒化为弧度 Q, Qh                                                        '首先计算P1到P2的方位角
    If R <> 0 Then
        K6 = K3 + Lz
        X6 = x2 + T2 * Cos(Qh)
        Y6 = y2 + T2 * Sin(Qh)
        Q6 = Q                                                                  '单位：度.分秒
    Else
        K6 = K3
        X6 = x2
        Y6 = y2
        Q6 = Q                                                                  '单位：度.分秒
    End If
    
    
    'a = Sqr(Abs(R * Ls1))
    'If banben = 0 Then a = Sqr(Abs(R * Ls2))
    'If a <> 0 Then
    'W = -1
    'I = R / Abs(R)    '判断左右偏的变量，右偏R取正，左偏取负
    'L = -Ls1  'L=K到ZH（HZ）点的弧长
    'If banben = 0 Then L = -Ls2
    '度分秒化为弧度 Q5, Q5h
    'E = Q5h - a ^ 2 * W * I / (2 * R ^ 2)   '弧度
    'U = L - L ^ 5 / (40 * a ^ 4) + L ^ 9 / (3456 * a ^ 8) - L ^ 13 / (599040 * a ^ 12)
    'V = L ^ 3 / (6 * a ^ 2) - L ^ 7 / (336 * a ^ 6) + L ^ 11 / (42240 * a ^ 10) - L ^ 15 / (9676800 * a ^ 14)
    'K6 = K3 + Lz
    'X6 = X5 + W * U * Cos(E) - I * V * Sin(E) '中桩的X坐标
    'Y6 = Y5 + W * U * Sin(E) + I * V * Cos(E) '中桩的Y坐标
    'Q6 = E     '切线方位角，单位弧度
    '弧度化为度分秒 Q6
    'Else
    'K6 = K5
    'X6 = X5
    'Y6 = Y5
    'Q6 = Q5 '单位：度.分秒
    'End If
End Sub
Private Sub Combo5_Click()
    I1 = Val(Combo5)
    计算特征点 I1, R, Ls1, Ls2, A1, A2, T1, T2, Lz, ap, K3, X3, Y3, Q3, K4, X4, Y4, Q4, K5, X5, Y5, Q5, K6, X6, Y6, Q6
    '把特征点显示在当前文本框内 ，R A Ls1 T1不变，B O Ls2 T2改变。Lz需要根据Ls1，Ls2，R来推算。
    Text91(0) = R                                                               '半径R
    Text89(0) = Round(A1, 4)                                                    '回旋参数A
    Text89(1) = Round(A2, 4)                                                    '回旋参数B
    Text90(0) = Round(Lz - Ls1 - Ls2, 3)                                        '圆曲线长O
    Text91(1) = Ls1                                                             '缓1长Ls1
    Text89(2) = Ls2                                                             '缓2长Ls2
    Text89(3) = Round(T1, 3)                                                    '切线长T1
    Text90(1) = Round(T2, 3)                                                    '切线长T2
    Text91(2) = Round(Lz, 3)                                                    '曲线总长
    Text91(3) = Round(ap, 6)                                                    '偏角
    
    If R <> 0 And Ls1 <> 0 Or Ls2 <> 0 Then                                     '如果有缓和曲线
        Label2 = "ZH点"
        Label3 = "HY点"
        Label4.Visible = True
        Label5.Visible = True
        Text86.Visible = True
        Text78.Visible = True
        Text77.Visible = True
        Text76.Visible = True
        Text85.Visible = True
        Text84.Visible = True
        Text83.Visible = True
        Text79.Visible = True
        Text88 = Round(K3, 3)
        Text82 = Round(X3, 4)
        Text81 = Round(Y3, 4)
        Text80 = Round(Q3, 5)
        Text87 = Round(K4, 3)
        Text75 = Round(X4, 4)
        Text74 = Round(Y4, 4)
        Text69 = Round(Q4, 5)                                                   '单位：度.分秒
        Text86 = Round(K5, 3)
        Text78 = Round(X5, 4)
        Text77 = Round(Y5, 4)
        Text76 = Round(Q5, 5)
        Text85 = Round(K6, 3)
        Text84 = Round(X6, 4)
        Text83 = Round(Y6, 4)
        Text79 = Round(Q6, 5)
    ElseIf Ls1 = 0 And R <> 0 Then                                              '如果没有缓和曲线
        Label2 = "ZY点"
        Label3 = "YZ点"
        Label4.Visible = False
        Label5.Visible = False
        Text86.Visible = False
        Text78.Visible = False
        Text77.Visible = False
        Text76.Visible = False
        Text85.Visible = False
        Text84.Visible = False
        Text83.Visible = False
        Text79.Visible = False
        Text88 = Round(K3, 3)
        Text82 = Round(X3, 4)
        Text81 = Round(Y3, 4)
        Text80 = Round(Q3, 5)
        Text87 = Round(K5, 3)
        Text75 = Round(X5, 4)
        Text74 = Round(Y5, 4)
        Text69 = Round(Q5, 5)                                                   '单位：度.分秒
    ElseIf R = 0 Then
        Label2 = "JD前"
        Label3 = "JD后"
        Label4.Visible = False
        Label5.Visible = False
        Text86.Visible = False
        Text78.Visible = False
        Text77.Visible = False
        Text76.Visible = False
        Text85.Visible = False
        Text84.Visible = False
        Text83.Visible = False
        Text79.Visible = False
        Text88 = Round(K3, 3)
        Text82 = Round(X3, 4)
        Text81 = Round(Y3, 4)
        Text80 = Round(Q3, 5)
        Text87 = Round(K6, 3)
        Text75 = Round(X6, 4)
        Text74 = Round(Y6, 4)
        Text69 = Round(Q6, 5)                                                   '单位：度.分秒
    Else
        Text88 = ""
        Text82 = ""
        Text81 = ""
        Text80 = ""
        Text87 = ""
        Text75 = ""
        Text74 = ""
        Text69 = ""
        Text86 = ""
        Text78 = ""
        Text77 = ""
        Text76 = ""
        Text85 = ""
        Text84 = ""
        Text83 = ""
        Text79 = ""
    End If
End Sub


Private Sub Command35_Click()                                                   '从圆曲线推算HY点参数
    '    TEXTKI = Text88
    '    TEXTJ = Text88
    '    TEXTC = Text82
    '    TEXTD = Text81
    '    TEXTF = Text80
    Form3.Left = 5500
    Form3.Top = 5700
    Form3.Show
End Sub

Private Sub Command36_Click()
    '    TEXTKI = Text85
    '    TEXTJ = Text87
    '    TEXTC = Text75
    '    TEXTD = Text74
    '    TEXTF = Text69
    '    TEXTR = Text91(0)
    '    TEXTA = Text89(0)
    '    TEXTB = Text89(1)
    '    TEXTO = Text90(0)
    '检验J1～Jn是否按照从小到大排列
    For n1 = 1 To Partnumber
        If shu(n1 - 1, 1) > shu(n1, 1) Then                                     '检验终点里程是否有序，如果有错误就修改
            msg = "第" & n1 & "段变坡点里程Ki不符合从小到大的顺序，现在修改吗？"
            queren = MsgBox(msg, vbYesNo + vbDefaultButton1 + vbCritical, "检验变坡点里程")
            If queren = vbYes Then
                Combo3.ListIndex = n1 - 1
                With Text29
                    .SetFocus                                                   '让text获得焦点
                    .SelStart = 0
                    .SelLength = Len(.text)
                End With
            End If
            Exit For
        End If
    Next
    If n1 = Partnumber + 1 Then
        MsgBox "没有发现变坡点里程的错误。", vbInformation + vbOKOnly, "检验终点里程"
    End If
End Sub

Private Sub Command37_Click()
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "导出全部特征点到文件..."
        .DefaultExt = "scp"
        .Filter = "坐标结果(*.scp)|*.scp"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
    End With
    Open CommonDialog1.FileName For Output As 8
    For I1 = Val(Text64) + 1 To Val(Text64) + Val(Text65) - 1
        计算特征点 I1, R, Ls1, Ls2, A1, A2, T1, T2, Lz, ap, K3, X3, Y3, Q3, K4, X4, Y4, Q4, K5, X5, Y5, Q5, K6, X6, Y6, Q6
        Write #8, "交点", I1, factor2(I1, 0), factor2(I1, 1), factor2(I1, 2), R, Ls1, Ls2, T1, T2, Lz, ap
        If Ls1 <> 0 Or Ls2 <> 0 And R <> 0 Then                                 '如果有缓和曲线
            Write #8, "ZH点", Round(K3, 3), Round(X3, 4), Round(Y3, 4), Round(Q3, 5)
            Write #8, "HY点", Round(K4, 3), Round(X4, 4), Round(Y4, 4), Round(Q4, 5)
            Write #8, "YH点", Round(K5, 3), Round(X5, 4), Round(Y5, 4), Round(Q5, 5)
            Write #8, "HZ点", Round(K6, 3), Round(X6, 4), Round(Y6, 4), Round(Q6, 5)
        ElseIf Ls1 * Ls2 = 0 And R <> 0 Then                                    '如果没有缓和曲线
            Write #8, "ZY点", Round(K3, 3), Round(X3, 4), Round(Y3, 4), Round(Q3, 5)
            Write #8, "YZ点", Round(K5, 3), Round(X5, 4), Round(Y5, 4), Round(Q5, 5)
        End If
        Write #8, "-------------------------------------------------------------------------"
    Next I1
    Close #8
    msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
    If MsgBox(msg, vbYesNo) = vbYes Then
        Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbNormalFocus
    End If
End Sub

Private Sub Command38_Click()                                                   '生成卡西欧9860计算器主程序
    '查找FA-124程序编辑窗口句柄
    Dim Sx As String
    Dim ckbt As String, cklm As String
    Sx = GetOpenWindowNames("CASIO FA-124 - ")
    '    MsgBox Sx
    Dim hwnd As Long
    Dim hWnd2 As Long, hwnd3 As Long, hwnd4 As Long
    hwnd = FindWindow(vbNullString, Sx)                                         '通过完整标题名找句柄
    hWnd2 = FindWindowEx(hwnd, 0, "MDIClient", vbNullString)
    If hWnd2 <> 0 Then
        EnumChildWindows hWnd2, AddressOf EnumChildProc, ByVal 0&
        '        MsgBox Module1.cklm
        hwnd3 = FindWindowEx(hWnd2, 0, Module1.cklm, vbNullString)              '模块级变量cklm必须带上父变量model1.cklm
    End If
    hwnd4 = FindWindowEx(hwnd3, 0, "Edit", vbNullString)
    '    MsgBox hwnd4
    If hwnd4 = 0 Then
        MsgBox "请先打开CASIO FA-124传输软件，调出程序编辑窗口，并把光标定位到要输入的位置！", vbOKOnly + vbExclamation, "导出CASIO fx-9750G/9850G/9860G程序"
        Exit Sub
    End If
    
    '    MsgBox "hwnd4=" & hwnd4
    
    If Dir(App.Path & "\closeIME.exe") <> "" Then Shell App.Path & "\closeIME.exe", vbHide
    
    msg1 = MsgBox("正算=已知里程K求XY，反算=已知XY求里程K，需要进行反算吗？" & vbCrLf & "是=只需要正算，一个主程序；否=需要正反算，正反算主程序+中桩子程序，取消=退出", vbYesNoCancel, "正反算确认")
    If msg1 = vbCancel Then
        Exit Sub                                                                '如果取消直接退出该过程。
    ElseIf msg1 = vbYes Then
        Me.WindowState = 1                                                      '主程序最小化
    Else                                                                        'msg1=vbNO
        '        Me.WindowState = 1              '主程序最小化
        Dim file6 As String
        Me.SetFocus
        file6 = UCase(Left(InputBox("请输入程序名，确保为6个英文以内的字符(例如1ZKXY)：", "请输入程序名"), 6))
        If file6 = "" Then Exit Sub
    End If
    
    '把casio fa-124程序编辑框带到前台并关闭输入法
    '    SetWindowPos hwnd, 0, 0, 0, 0, 0, 3     '设置置顶'Private Const HWND_TOP = 0
    '    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE     '取消置顶
    '以下代码激活父窗体：
    SetForegroundWindow hwnd
    '以下代码激活子窗体：
    'Private Const WM_MOUSEACTIVATE = &H21
    SendMessage hwnd4, WM_MOUSEACTIVATE, 0, 0
    Sleep 2000                                                                  '等待焦点切换到FA-124程序编辑框
    
    'Set WshShell = CreateObject("WScript.Shell")
    'MySendKey "1{+}"   '代替用法，测试WIN7 WIN10无效。
    
    
    '    '以下代码把前台进程的输入法关闭：
    '    '-------------------------------------------------------------------------------------------------------
    '    Dim hwnd As Long                    '前面已经定义过了，不要定义
    On Error Resume Next
    Dim hCurThread As Long
    Dim sCaption As String
    Dim sBuffer As String
    Dim hCurKBDLayout As Long
    hwnd = GetForegroundWindow
    sCaption = Space(255)
    GetWindowText hwnd, sCaption, 255
    If InStr(sCaption, Chr(0)) Then
        sCaption = Left(sCaption, InStr(sCaption, Chr(0)) - 1)
    End If
    hCurThread = GetWindowThreadProcessId(hwnd, ByVal 0&)
    hCurKBDLayout = GetKeyboardLayout(hCurThread)
    If ImmIsIME(hCurKBDLayout) = 1 Then
        sBuffer = Space(255)
        RetCount = ImmGetDescription(ByVal hCurKBDLayout, sBuffer, 255)
        sBuffer = Left(sBuffer, RetCount)
        '        sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        sBuffer = "English(American)"
    End If
    If sBuffer <> "English(American)" And sBuffer <> "" Then MySendKey "^ ", False
    Sleep 1000                                                                  '延迟1秒确保ctrl+空格生效
    '-------------------------------------------------------------------------------------------------------
    
    '修正大小写状态
    If GetKeyState(vbKeyCapital) = 1 Then                                       '如果CAPSLOCK被按下
        MySendKey vbKeyCapital                                                  '就发送{capslock}键
    End If
    
    '开始写程序了：
    If msg1 = vbYes Then                                                        '只进行正算，1个程序
        '        SendKeys "Deg:""K=""", False
        MySendKeytext "Deg:""K="""
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "Str 1:1", False
        MySendKeytext "Str 1:1"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "%+k{UP 5}{ENTER 2}", False 'θ
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn, vbKeyReturn
        '        SendKeys "K", False
        MySendKeytext "K"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "%+k{UP 6}{ENTER}", False 'r
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
        '        SendKeys ":Prog ""A1"":", False
        MySendKeytext ":Prog ""A1"":"
        '        SendKeys "%+k{UP 6}{ENTER}", False 'r
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "K{ENTER}", False
        SendKeySHIFTplus vbKeyK, vbKeyReturn
        For n1 = 1 To Val(TextPartNumber)
            If factor(n1, 0) = "直线" Then
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys factor(n1, 1), False
                MySendKeytext factor(n1, 1)
                '                SendKeys ":Then ", False
                MySendKeytext ":Then "
                '                SendKeys factor(n1, 2), False
                MySendKeytext factor(n1, 2)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "J:", False
                MySendKeytext "J:"
                '                SendKeys factor(n1, 3), False
                MySendKeytext factor(n1, 3)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "C:", False
                MySendKeytext "C:"
                '                SendKeys factor(n1, 4), False
                MySendKeytext factor(n1, 4)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "D:", False
                MySendKeytext "D:"
                '                SendKeys Int(factor(n1, 5)), False
                MySendKeytext Int(factor(n1, 5))
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100, False
                MySendKeytext Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####"), False
                MySendKeytext Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####")
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "F:", False
                MySendKeytext "F:"
                '                SendKeys "Prog ""ZX"":", False
                MySendKeytext "Prog ""ZX"":"
                '                SendKeys "Goto 1:IfEnd{ENTER}", False
                MySendKeytext "Goto 1:IfEnd"
                MySendKey vbKeyReturn
                '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & "Prog ""ZX"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "圆" Then
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys factor(n1, 1), False
                MySendKeytext factor(n1, 1)
                '                SendKeys ":Then ", False
                MySendKeytext ":Then "
                '                SendKeys factor(n1, 2), False
                MySendKeytext factor(n1, 2)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "J:", False
                MySendKeytext "J:"
                '                SendKeys factor(n1, 3), False
                MySendKeytext factor(n1, 3)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "C:", False
                MySendKeytext "C:"
                '                SendKeys factor(n1, 4), False
                MySendKeytext factor(n1, 4)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "D:", False
                MySendKeytext "D:"
                '                SendKeys Int(factor(n1, 5)), False
                MySendKeytext Int(factor(n1, 5))
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100, False
                MySendKeytext Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####"), False
                MySendKeytext Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####")
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "F:", False
                MySendKeytext "F:"
                '                SendKeys factor(n1, 6), False
                MySendKeytext factor(n1, 6)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "R:", False
                MySendKeytext "R:"
                '                SendKeys "Prog ""YUAN"":", False
                MySendKeytext "Prog ""YUAN"":"
                '                SendKeys "Goto 1:IfEnd{ENTER}", False
                MySendKeytext "Goto 1:IfEnd"
                MySendKey vbKeyReturn
                '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & "Prog ""YUAN"":" _
                '                & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "缓1+圆+缓2" Then
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys factor(n1, 1), False
                MySendKeytext factor(n1, 1)
                '                SendKeys ":Then ", False
                MySendKeytext ":Then "
                '                SendKeys factor(n1, 2), False
                MySendKeytext factor(n1, 2)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "J:", False
                MySendKeytext "J:"
                '                SendKeys factor(n1, 3), False
                MySendKeytext factor(n1, 3)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "C:", False
                MySendKeytext "C:"
                '                SendKeys factor(n1, 4), False
                MySendKeytext factor(n1, 4)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "D:", False
                MySendKeytext "D:"
                '                SendKeys Int(factor(n1, 5)), False
                MySendKeytext Int(factor(n1, 5))
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100, False
                MySendKeytext Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####"), False
                MySendKeytext Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####")
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "F:", False
                MySendKeytext "F:"
                '                SendKeys factor(n1, 6), False
                MySendKeytext factor(n1, 6)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "R:", False
                MySendKeytext "R:"
                '                SendKeys factor(n1, 7), False
                MySendKeytext factor(n1, 7)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "A:", False
                MySendKeytext "A:"
                '                SendKeys factor(n1, 8), False
                MySendKeytext factor(n1, 8)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "B:", False
                MySendKeytext "B:"
                '                SendKeys factor(n1, 9), False
                MySendKeytext factor(n1, 9)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "O:", False
                MySendKeytext "O:"
                '                SendKeys "Prog ""PQX"":", False
                MySendKeytext "Prog ""PQX"":"
                '                SendKeys "Goto 1:IfEnd{ENTER}", False
                MySendKeytext "Goto 1:IfEnd"
                MySendKey vbKeyReturn
                '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & factor(n1, 7) & "→A:" _
                '                & factor(n1, 8) & "→B:" & factor(n1, 9) & "→O:" & "Prog ""PQX"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "布劳斯曲线" Then
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys factor(n1, 1), False
                MySendKeytext factor(n1, 1)
                '                SendKeys ":Then ", False
                MySendKeytext ":Then "
                '                SendKeys factor(n1, 2), False
                MySendKeytext factor(n1, 2)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "J:", False
                MySendKeytext "J:"
                '                SendKeys factor(n1, 3), False
                MySendKeytext factor(n1, 3)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "C:", False
                MySendKeytext "C:"
                '                SendKeys factor(n1, 4), False
                MySendKeytext factor(n1, 4)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "D:", False
                MySendKeytext "D:"
                '                SendKeys Int(factor(n1, 5)), False
                MySendKeytext Int(factor(n1, 5))
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100, False
                MySendKeytext Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####"), False
                MySendKeytext Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####")
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "F:", False
                MySendKeytext "F:"
                '                SendKeys factor(n1, 6), False
                MySendKeytext factor(n1, 6)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "R:", False
                MySendKeytext "R:"
                '                SendKeys factor(n1, 7), False
                MySendKeytext factor(n1, 7)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "A:", False
                MySendKeytext "A:"
                '                SendKeys "Prog ""BLOSS"":", False
                MySendKeytext "Prog ""BLOSS"":"
                '                SendKeys "Goto 1:IfEnd{ENTER}", False
                MySendKeytext "Goto 1:IfEnd"
                MySendKey vbKeyReturn
                '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & factor(n1, 7) & "→A:" _
                '                & "Prog ""BLOSS"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "麦克康纳尔曲线" Then
                Ls = CDec(Val(factor(n1, 1)) - Val(factor(n1, 2)))              '对于MC1，Ls=分段终点Khy-Kzh
                If Abs(Ls) < 0.05 Then Ls = Val(factor(n1, 1)) - Val(factor(n1 - 1, 1))
                v0 = Format(CDec(Val(factor(n1, 7)) / 3.6), "##0.####")         '设计平衡车速m/s，(120-300)/3.6
                t0 = Format(CDec(Atn(Val(factor(n1, 8)))) * 180 / pai, "##0.####") '缓和曲线起点处的路面横向倾角atn(0.7%-0.8%)，弧度
                g0 = CDec((Val(factor(n1, 9))))                                 '重力加速度m/s2
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys factor(n1, 1), False
                MySendKeytext factor(n1, 1)
                '                SendKeys ":Then ", False
                MySendKeytext ":Then "
                '                SendKeys factor(n1, 2), False
                MySendKeytext factor(n1, 2)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "J:", False
                MySendKeytext "J:"
                '                SendKeys factor(n1, 3), False
                MySendKeytext factor(n1, 3)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "C:", False
                MySendKeytext "C:"
                '                SendKeys factor(n1, 4), False
                MySendKeytext factor(n1, 4)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "D:", False
                MySendKeytext "D:"
                '                SendKeys Int(factor(n1, 5)), False
                MySendKeytext Int(factor(n1, 5))
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100, False
                MySendKeytext Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####"), False
                MySendKeytext Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####")
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "F:", False
                MySendKeytext "F:"
                '                SendKeys factor(n1, 6), False
                MySendKeytext factor(n1, 6)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "R:", False
                MySendKeytext "R:"
                '                SendKeys v0, False
                MySendKeytext v0
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "A:", False
                MySendKeytext "A:"
                '                SendKeys t0, False
                MySendKeytext t0
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "B:", False
                MySendKeytext "B:"
                '                SendKeys Ls, False
                MySendKeytext Ls
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "O:", False
                MySendKeytext "O:"
                '                SendKeys "Prog ""MC"":", False
                MySendKeytext "Prog ""MC"":Goto 1:IfEnd"
                MySendKey vbKeyReturn
                '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & v0 & "→A:" _
                '                & t0 & "→B:" & Ls & "→O:Prog ""MC"":Goto 1:IfEnd↙"
            End If
        Next
        '        SendKeys "Lbl 1:Prog ""CEYI""", False '最后一个MySendKey等待发送完毕
        '        SendKeys "Lbl 1:Prog ""CEYI""", True '最后一个MySendKey等待发送完毕
        MySendKeytext "Lbl 1:Prog ""CEYI"""
        DoEvents
        'fa-124程序编写完毕，弹出对话框：
        Me.SetFocus
        MsgBox "CASIO FA-124坐标程序输出完毕！", vbOKOnly + vbExclamation, "导出CASIO fx-9750G/9850G/9860G程序"
    Else
        '正反算和中桩共3个程序，放在一起，手工在FA-124中复制粘贴，拆分成3个即可。
        
        '正算主程序：
        '        SendKeys "Deg:""K=""", False
        MySendKeytext "Deg:""K="""
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "Str 1:1", False
        MySendKeytext "Str 1:1"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "%+k{UP 5}{ENTER 2}", False 'θ
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn, vbKeyReturn
        '        SendKeys "K", False
        MySendKeytext "K"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "%+k{UP 6}{ENTER}", False 'r
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
        '        SendKeys ":Prog ""A1"":", False
        MySendKeytext ":Prog ""A1"":"
        '        SendKeys "%+k{UP 6}{ENTER}", False 'r
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "K{ENTER}", False
        MySendKeytext "K"
        MySendKey vbKeyReturn
        '        SendKeys "Prog """
        MySendKeytext "Prog """
        '        SendKeys file6
        MySendKeytext file6
        '        SendKeys "ZZ"":Prog ""CEYI"""
        MySendKeytext "ZZ"":Prog ""CEYI"""
        '        SendKeys "{enter 4}"
        MySendKey vbKeyReturn, vbKeyReturn, vbKeyReturn, vbKeyReturn
        '        Print #12, "?K:"
        '        Print #12, "Prog """ & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "ZZ"":Prog ""CEYI"""
        '        Print #12, "========================================================================================="
        
        '反算主程序：
        '以下代码为M变量回显
        '        SendKeys """M=""", False
        MySendKeytext """M="""
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "Str 1:1", False
        MySendKeytext "Str 1:1"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "%+k{UP 5}{ENTER 2}", False 'θ
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn, vbKeyReturn
        '        SendKeys "M", False
        MySendKeytext "M"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "%+k{UP 6}{ENTER}", False 'r
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
        '        SendKeys ":Prog ""A1"":", False
        MySendKeytext ":Prog ""A1"":"
        '        SendKeys "%+k{UP 6}{ENTER}", False 'r
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "M{ENTER}", False
        MySendKeytext "M"
        MySendKey vbKeyReturn
        '以下代码为N变量回显
        '        SendKeys """N=""", False
        MySendKeytext """N="""
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "Str 1:2", False
        MySendKeytext "Str 1:2"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "%+k{UP 5}{ENTER 2}", False 'θ
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn, vbKeyReturn
        '        SendKeys "N", False
        MySendKeytext "N"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "%+k{UP 6}{ENTER}", False 'r
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
        '        SendKeys ":Prog ""A1"":", False
        MySendKeytext ":Prog ""A1"":"
        '        SendKeys "%+k{UP 6}{ENTER}", False 'r
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "N{ENTER}", False
        MySendKeytext "N"
        MySendKey vbKeyReturn
        '以下代码为K变量回显
        '        SendKeys """K=""", False
        MySendKeytext """K="""
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "Str 1:3", False
        MySendKeytext "Str 1:3"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "%+k{UP 5}{ENTER 2}", False 'θ
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn, vbKeyReturn
        '        SendKeys "K", False
        MySendKeytext "K"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "%+k{UP 6}{ENTER}", False 'r
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
        '        SendKeys ":Prog ""A1"":", False
        MySendKeytext ":Prog ""A1"":"
        '        SendKeys "%+k{UP 6}{ENTER}", False 'r
        SendKeyALTplus vbKeyK, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyUp, vbKeyReturn
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "K{ENTER}", False
        SendKeySHIFTplus vbKeyK, vbKeyReturn
        '开始反算
        '        SendKeys "Lbl 1:Prog """, False
        MySendKeytext "Lbl 1:Prog """
        '        SendKeys file6, False
        MySendKeytext file6
        '        SendKeys "ZZ""{ENTER}", False
        MySendKeytext "ZZ"""
        MySendKey vbKeyReturn
        '        SendKeys "Pol{(}M-X,N-Y{)}:List Ans[1]", False
        MySendKeytext "Pol(M-X,N-Y):List Ans[1]"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "I:List Ans[2]", False
        MySendKeytext "I:List Ans[2]"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "J:J<0", False
        MySendKeytext "J:J<0"
        '        SendKeys "%+pj{DOWN 2}{ENTER}", False '=>
        SendKeyALTplus vbKeyP, vbKeyJ, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "J{+}360", False
        MySendKeytext "J+360"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "J{ENTER}", False
        SendKeySHIFTplus vbKeyJ, vbKeyReturn
        '        SendKeys "If Abs {(}I", False
        MySendKeytext "If Abs (I"
        '        SendKeys "%+k{DOWN 2}{ENTER}", False '×
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
        MySendKeytext "0.382"
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "cos {(}J-Q{)}{)}>0.005:Then K{+}I"
        MySendKeytext "cos (J-Q))>0.005:Then K+I"
        '        SendKeys "%+k{DOWN 2}{ENTER}", False '×
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "0.382"
        MySendKeytext "0.382"
        '        SendKeys "%+k{DOWN 2}{ENTER}", False '×
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "cos {(}J-Q{)}", False
        MySendKeytext "cos (J-Q)"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "K:Goto 1:{ENTER}", False
        MySendKeytext "K:Goto 1:"
        MySendKey vbKeyReturn
        '反算完毕，计算中边桩
        '        SendKeys "Else {ENTER}", False
        MySendKeytext "Else "
        MySendKey vbKeyReturn
        '        SendKeys """K="":K", False
        MySendKeytext """K="":K"
        '        SendKeys "%+p{DOWN 4}{ENTER}", False '◢
        SendKeyALTplus vbKeyP, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "I", False
        MySendKeytext "I"
        '        SendKeys "%+k{DOWN 2}{ENTER}", False '×
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "sin {(}J-Q{)}", False
        MySendKeytext "sin (J-Q)"
        '        SendKeys "%+k{DOWN 3}{ENTER}", False '÷
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "Abs {(}sin {(}J-Q{)}{)}", False
        MySendKeytext "Abs (sin (J-Q))"
        '        SendKeys "%+k{DOWN 24}{ENTER}", False '→
        SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
        vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "T:""T="":T", False
        MySendKeytext "T:""T="":T"
        '        SendKeys "%+p{DOWN 4}{ENTER}", False '◢
        SendKeyALTplus vbKeyP, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
        '        SendKeys "IfEnd{ENTER}", False
        MySendKeytext "IfEnd"
        MySendKey vbKeyReturn
        '        SendKeys "Prog """, False
        MySendKeytext "Prog """
        '        SendKeys file6, False
        MySendKeytext file6
        '        SendKeys "ZS""", False
        MySendKeytext "ZS"""
        '        SendKeys "{ENTER 4}"
        MySendKey vbKeyReturn, vbKeyReturn, vbKeyReturn, vbKeyReturn
        '        Print #12, "反算主程序:" & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "FS"
        '        Print #12, "?M:?N:?K:"
        '        Print #12, "Lbl 1:Prog """ & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "ZZ"":"
        '        Print #12, "Pol(M-X,N-Y):        (9750和9860计算器需附加List Ans[1]→I:List Ans[2]→J:J<0=>J+360→J:)"
        '        Print #12, "If Abs (I*0.382*cos (J-Q))>0.005:Then "
        '        Print #12, "K+I*0.382*cos (J-Q)→K:Goto 1:"
        '        Print #12, "Else ""K="":K◢"
        '        Print #12, "I*sin (J-Q)/Abs (sin (J-Q))→T:""T="":T◢"
        '        Print #12, "IfEnd"
        '        Print #12, "Prog """ & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "ZS"""
        '        Print #12, "========================================================================================="
        
        
        '中桩子程序
        For n1 = 1 To Val(TextPartNumber)
            If factor(n1, 0) = "直线" Then
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys factor(n1, 1), False
                MySendKeytext factor(n1, 1)
                '                SendKeys ":Then ", False
                MySendKeytext ":Then "
                '                SendKeys factor(n1, 2), False
                MySendKeytext factor(n1, 2)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "J:", False
                MySendKeytext "J:"
                '                SendKeys factor(n1, 3), False
                MySendKeytext factor(n1, 3)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "C:", False
                MySendKeytext "C:"
                '                SendKeys factor(n1, 4), False
                MySendKeytext factor(n1, 4)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "D:", False
                MySendKeytext "D:"
                '                SendKeys Int(factor(n1, 5)), False
                MySendKeytext Int(factor(n1, 5))
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100, False
                MySendKeytext Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####"), False
                MySendKeytext Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####")
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "F:", False
                MySendKeytext "F:"
                '                SendKeys "Prog ""ZX"":", False
                MySendKeytext "Prog ""ZX"":"
                '                SendKeys "Goto 1:IfEnd{ENTER}", False
                MySendKeytext "Goto 1:IfEnd"
                MySendKey vbKeyReturn
                '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & "Prog ""ZX"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "圆" Then
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys factor(n1, 1), False
                MySendKeytext factor(n1, 1)
                '                SendKeys ":Then ", False
                MySendKeytext ":Then "
                '                SendKeys factor(n1, 2), False
                MySendKeytext factor(n1, 2)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "J:", False
                MySendKeytext "J:"
                '                SendKeys factor(n1, 3), False
                MySendKeytext factor(n1, 3)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "C:", False
                MySendKeytext "C:"
                '                SendKeys factor(n1, 4), False
                MySendKeytext factor(n1, 4)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "D:", False
                MySendKeytext "D:"
                '                SendKeys Int(factor(n1, 5)), False
                MySendKeytext Int(factor(n1, 5))
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100, False
                MySendKeytext Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####"), False
                MySendKeytext Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####")
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "F:", False
                MySendKeytext "F:"
                '                SendKeys factor(n1, 6), False
                MySendKeytext factor(n1, 6)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "R:", False
                MySendKeytext "R:"
                '                SendKeys "Prog ""YUAN"":", False
                MySendKeytext "Prog ""YUAN"":"
                '                SendKeys "Goto 1:IfEnd{ENTER}", False
                MySendKeytext "Goto 1:IfEnd"
                MySendKey vbKeyReturn
                '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & "Prog ""YUAN"":" _
                '                & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "缓1+圆+缓2" Then
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys factor(n1, 1), False
                MySendKeytext factor(n1, 1)
                '                SendKeys ":Then ", False
                MySendKeytext ":Then "
                '                SendKeys factor(n1, 2), False
                MySendKeytext factor(n1, 2)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "J:", False
                MySendKeytext "J:"
                '                SendKeys factor(n1, 3), False
                MySendKeytext factor(n1, 3)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "C:", False
                MySendKeytext "C:"
                '                SendKeys factor(n1, 4), False
                MySendKeytext factor(n1, 4)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "D:", False
                MySendKeytext "D:"
                '                SendKeys Int(factor(n1, 5)), False
                MySendKeytext Int(factor(n1, 5))
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100, False
                MySendKeytext Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####"), False
                MySendKeytext Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####")
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "F:", False
                MySendKeytext "F:"
                '                SendKeys factor(n1, 6), False
                MySendKeytext factor(n1, 6)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "R:", False
                MySendKeytext "R:"
                '                SendKeys factor(n1, 7), False
                MySendKeytext factor(n1, 7)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "A:", False
                MySendKeytext "A:"
                '                SendKeys factor(n1, 8), False
                MySendKeytext factor(n1, 8)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "B:", False
                MySendKeytext "B:"
                '                SendKeys factor(n1, 9), False
                MySendKeytext factor(n1, 9)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "O:", False
                MySendKeytext "O:"
                '                SendKeys "Prog ""PQX"":", False
                MySendKeytext "Prog ""PQX"":"
                '                SendKeys "Goto 1:IfEnd{ENTER}", False
                MySendKeytext "Goto 1:IfEnd"
                MySendKey vbKeyReturn
                '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & factor(n1, 7) & "→A:" _
                '                & factor(n1, 8) & "→B:" & factor(n1, 9) & "→O:" & "Prog ""PQX"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "布劳斯曲线" Then
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys factor(n1, 1), False
                MySendKeytext factor(n1, 1)
                '                SendKeys ":Then ", False
                MySendKeytext ":Then "
                '                SendKeys factor(n1, 2), False
                MySendKeytext factor(n1, 2)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "J:", False
                MySendKeytext "J:"
                '                SendKeys factor(n1, 3), False
                MySendKeytext factor(n1, 3)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "C:", False
                MySendKeytext "C:"
                '                SendKeys factor(n1, 4), False
                MySendKeytext factor(n1, 4)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "D:", False
                MySendKeytext "D:"
                '                SendKeys Int(factor(n1, 5)), False
                MySendKeytext Int(factor(n1, 5))
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100, False
                MySendKeytext Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####"), False
                MySendKeytext Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####")
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "F:", False
                MySendKeytext "F:"
                '                SendKeys factor(n1, 6), False
                MySendKeytext factor(n1, 6)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "R:", False
                MySendKeytext "R:"
                '                SendKeys factor(n1, 7), False
                MySendKeytext factor(n1, 7)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "A:", False
                MySendKeytext "A:"
                '                SendKeys "Prog ""BLOSS"":", False
                MySendKeytext "Prog ""BLOSS"":"
                '                SendKeys "Goto 1:IfEnd{ENTER}", False
                MySendKeytext "Goto 1:IfEnd"
                MySendKey vbKeyReturn
                '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & factor(n1, 7) & "→A:" _
                '                & "Prog ""BLOSS"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "麦克康纳尔曲线" Then
                Ls = CDec(Val(factor(n1, 1)) - Val(factor(n1, 2)))              '对于MC1，Ls=分段终点Khy-Kzh
                If Abs(Ls) < 0.05 Then Ls = Val(factor(n1, 1)) - Val(factor(n1 - 1, 1))
                v0 = Format(CDec(Val(factor(n1, 7)) / 3.6), "##0.####")         '设计平衡车速m/s，(120-300)/3.6
                t0 = Format(CDec(Atn(Val(factor(n1, 8)))) * 180 / pai, "##0.####") '缓和曲线起点处的路面横向倾角atn(0.7%-0.8%)，弧度
                g0 = CDec((Val(factor(n1, 9))))                                 '重力加速度m/s2
                '                SendKeys "If K", False
                MySendKeytext "If K"
                '                SendKeys "%+pr{DOWN 5}{ENTER}", False '≤
                SendKeyALTplus vbKeyP, vbKeyR, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys factor(n1, 1), False
                MySendKeytext factor(n1, 1)
                '                SendKeys ":Then ", False
                MySendKeytext ":Then "
                '                SendKeys factor(n1, 2), False
                MySendKeytext factor(n1, 2)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "J:", False
                MySendKeytext "J:"
                '                SendKeys factor(n1, 3), False
                MySendKeytext factor(n1, 3)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "C:", False
                MySendKeytext "C:"
                '                SendKeys factor(n1, 4), False
                MySendKeytext factor(n1, 4)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "D:", False
                MySendKeytext "D:"
                '                SendKeys Int(factor(n1, 5)), False
                MySendKeytext Int(factor(n1, 5))
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100, False
                MySendKeytext Int(factor(n1, 5) * 100) - Int(factor(n1, 5)) * 100
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####"), False
                MySendKeytext Format(factor(n1, 5) * 10000 - Int(factor(n1, 5) * 100) * 100, "#0.####")
                '                SendKeys "%+oa{DOWN 3}{ENTER}", False '°
                SendKeyALTplus vbKeyO, vbKeyA, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "F:", False
                MySendKeytext "F:"
                '                SendKeys factor(n1, 6), False
                MySendKeytext factor(n1, 6)
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "R:", False
                MySendKeytext "R:"
                '                SendKeys v0, False
                MySendKeytext v0
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "A:", False
                MySendKeytext "A:"
                '                SendKeys t0, False
                MySendKeytext t0
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "B:", False
                MySendKeytext "B:"
                '                SendKeys Ls, False
                MySendKeytext Ls
                '                SendKeys "%+k{DOWN 24}{ENTER}", False '→
                SendKeyALTplus vbKeyK, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, _
                vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyDown, vbKeyReturn
                '                SendKeys "O:", False
                MySendKeytext "O:"
                '                SendKeys "Prog ""MC"":", False
                MySendKeytext "Prog ""MC"":Goto 1:IfEnd"
                MySendKey vbKeyReturn
                '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & v0 & "→A:" _
                '                & t0 & "→B:" & Ls & "→O:Prog ""MC"":Goto 1:IfEnd↙"
            End If
        Next
        '       MySendKey "Lbl 1:", False        '最后一个MySendKey等待发送完毕
        '        SendKeys "Lbl 1:", True         '最后一个MySendKey等待发送完毕
        MySendKeytext "Lbl 1:"
        DoEvents                                                                '把控制权交给主程序
        'fa-124程序编写完毕，弹出对话框：
        Me.SetFocus
        MsgBox "CASIO FA-124程序输出完毕！请参考CASIO fx-5800P正反算程序清单，把以上程序在FA-124软件中复制粘贴依次拆分成三个程序,分别放在各自的程序名里面：(已用三个空行区分开)" & vbCrLf & _
        file6 & "ZS(正算)" & vbCrLf & file6 & "FS(反算)" & vbCrLf & file6 & "ZZ(中桩)", vbOKOnly + vbExclamation, "导出CASIO fx-9750G/9850G/9860G程序"
        '    Me.WindowState = 0                  '主程序窗口恢复正常大小
        
        '        Print #12, "中桩子程序:" & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "ZZ"
        '        For n1 = 1 To Val(TextPartNumber)
        '            If factor(n1, 0) = "直线" Then
        '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
        '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & "Prog ""ZX"":" & "Goto 1:IfEnd↙"
        '            ElseIf factor(n1, 0) = "圆" Then
        '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
        '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & "Prog ""YUAN"":" _
        '                & "Goto 1:IfEnd↙"
        '            ElseIf factor(n1, 0) = "缓1+圆+缓2" Then
        '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
        '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & factor(n1, 7) & "→A:" _
        '                & factor(n1, 8) & "→B:" & factor(n1, 9) & "→O:" & "Prog ""PQX"":" & "Goto 1:IfEnd↙"
        '            ElseIf factor(n1, 0) = "布劳斯曲线" Then
        '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
        '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & factor(n1, 7) & "→A:" _
        '                & "Prog ""BLOSS"":" & "Goto 1:IfEnd↙"
        '            ElseIf factor(n1, 0) = "麦克康纳尔曲线" Then
        '                Ls = CDec(Val(factor(n1, 1)) - Val(factor(n1, 2))) '对于MC1，Ls=分段终点Khy-Kzh
        '                If Abs(Ls) < 0.05 Then Ls = Val(factor(n1, 1)) - Val(factor(n1 - 1, 1))
        '                v0 = Format(CDec(Val(factor(n1, 7)) / 3.6), "##0.####") '设计平衡车速m/s，(120-300)/3.6
        '                t0 = Format(CDec(Atn(Val(factor(n1, 8)))) * 180 / pai, "##0.####") '缓和曲线起点处的路面横向倾角atn(0.7%-0.8%)，弧度
        '                g0 = CDec((Val(factor(n1, 9)))) '重力加速度m/s2
        '                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
        '                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & v0 & "→A:" _
        '                & t0 & "→B:" & Ls & "→O:Prog ""MC"":Goto 1:IfEnd↙"
    End If
End Sub

Private Sub Command39_Click()
    Dim n1, n2 As Integer                                                       '循环用的变量
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "保存平曲线参数到文件..."
        .DefaultExt = "ini"
        .Filter = "输出的曲线参数(*.ini)|*.ini"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
        'Text2 = .FileTitle
    End With
    
    njdh = CDec(Val(Text64))                                                    '起始交点号
    njd = CDec(Val(Text65))                                                     '交点数目，包括起点及终点
    id = 1                                                                      '平曲线分段序号
    
    For n1 = njdh + 1 To njdh + njd - 2
        计算特征点 n1 - 1, R, Ls1, Ls2, A1, A2, T1, T2, Lz, ap, K3, X3, Y3, Q3, K4, X4, Y4, Q4, K5, X5, Y5, Q5, K6, X6, Y6, Q6
        R1p = R
        K6p = K6
        Ls2p = Ls2
        A2p = A2
        If n1 - 1 = CDec(Val(Text64)) Then K6p = factor2(n1 - 1, 0)             '如果n1-1=起点号，K6p取起点里程
        K2 = CDec(Val(factor2(n1, 0)))
        计算特征点 n1 + 1, R, Ls1, Ls2, A1, A2, T1, T2, Lz, ap, K3, X3, Y3, Q3, K4, X4, Y4, Q4, K5, X5, Y5, Q5, K6, X6, Y6, Q6
        R2p = R
        K3p = K3
        Ls1p = Ls1
        A1p = A1
        计算特征点 n1, R, Ls1, Ls2, A1, A2, T1, T2, Lz, ap, K3, X3, Y3, Q3, K4, X4, Y4, Q4, K5, X5, Y5, Q5, K6, X6, Y6, Q6
        
        If Abs(K3 - K6p) > 0.015 Then                                           '直线段
            factor3(id, 0) = "直线"                                             '第I分段的第0列存储曲线类型
            factor3(id, 1) = K3                                                 '第I分段的第1列存储分段终点
            factor3(id, 2) = K3                                                 '第I分段的第2列存储起算点里程
            factor3(id, 3) = Round(X3, 4)                                       '第I分段的第3列存储起算点X坐标
            factor3(id, 4) = Round(Y3, 4)                                       '第I分段的第3列存储起算点Y坐标
            factor3(id, 5) = Round(Q3, 5)                                       '第I分段的第3列存储起算点方位角
            factor3(id, 6) = 0                                                  '第I分段的第6列存储圆半径
            factor3(id, 7) = 0                                                  '第I分段的第7列存储第一回旋参数
            factor3(id, 8) = 0                                                  '第I分段的第8列存储第二回旋参数
            factor3(id, 9) = 0                                                  '第I分段的第9列存储圆曲线长
            id = id + 1                                                         '分段数加一，准备计算缓1+圆+缓2
        End If
        
        
        If Ls1 = 0 And Ls2 = 0 And R <> 0 Then                                  '类型：圆
            factor3(id, 0) = "圆"                                               '第I分段的第0列存储曲线类型
        End If
        
        
        If Ls1 <> 0 Or Ls2 <> 0 And R <> 0 Then                                 '类型：缓1+圆+缓2
            factor3(id, 0) = "缓1+圆+缓2"                                       '第I分段的第0列存储曲线类型
        End If
        
        
        If R <> 0 Then
            factor3(id, 1) = K3 + Lz                                            '第I分段的第1列存储分段终点
            factor3(id, 2) = K4                                                 '第I分段的第2列存储起算点里程
            factor3(id, 3) = Round(X4, 4)                                       '第I分段的第3列存储起算点X坐标
            factor3(id, 4) = Round(Y4, 4)                                       '第I分段的第4列存储起算点Y坐标
            factor3(id, 5) = Round(Q4, 5)                                       '第I分段的第5列存储起算点方位角
            factor3(id, 6) = R                                                  '第I分段的第6列存储圆半径
            factor3(id, 7) = Round(A1, 5)                                       '第I分段的第7列存储第一回旋参数
            factor3(id, 8) = Round(A2, 5)                                       '第I分段的第8列存储第二回旋参数
            factor3(id, 9) = Lz - Ls1 - Ls2                                     '第I分段的第9列存储圆曲线长
            If Abs(K3 - K6p) < 0.015 And Ls1 = 0 And Ls2p <> 0 Or Ls1 <> 0 And Ls2p = 0 And R * R1p > 0 Then
                If A2p <> 0 Then factor3(id, 7) = Round(Abs(A2p), 5)            '第I分段的第7列存储第一回旋参数
            End If
            If Abs(K3p - K6) < 0.015 And Ls2 = 0 And Ls1p <> 0 Or Ls2 <> 0 And Ls1p = 0 And R * R2p > 0 Then
                If A1p <> 0 Then factor3(id, 8) = Round(Abs(A1p), 5)            '第I分段的第8列存储第二回旋参数
                If Abs(R) <= Abs(R2p) Then
                    factor3(id, 1) = K3 + Lz + Ls1p                             '第I分段的第1列存储分段终点
                Else
                    factor3(id, 1) = K3 + Lz - Ls2                              '第I分段的第1列存储分段终点
                End If
            End If
            id = id + 1
        End If
    Next n1
    
    
    If Abs(factor2(njdh + njd - 1, 0) - K6) > 0.015 Then                        '如果存在末直线段
        factor3(id, 0) = "直线"                                                 '第I分段的第0列存储曲线类型
        factor3(id, 1) = factor2(njdh + njd - 1, 0)                             '第I分段的第1列存储分段终点
        factor3(id, 2) = factor2(njdh + njd - 1, 0)                             '第I分段的第2列存储起算点里程
        factor3(id, 3) = factor2(njdh + njd - 1, 1)                             '第I分段的第3列存储起算点X坐标
        factor3(id, 4) = factor2(njdh + njd - 1, 2)                             '第I分段的第3列存储起算点Y坐标
        factor3(id, 5) = Round(Q6, 5)                                           '第I分段的第3列存储起算点方位角
        factor3(id, 6) = 0                                                      '第I分段的第6列存储圆半径
        factor3(id, 7) = 0                                                      '第I分段的第7列存储第一回旋参数
        factor3(id, 8) = 0                                                      '第I分段的第8列存储第二回旋参数
        factor3(id, 9) = 0                                                      '第I分段的第9列存储圆曲线长
        id = id + 1                                                             '分段数加一，准备计算缓1+圆+缓2
    End If
    
    Open CommonDialog1.FileName For Output As 9
    Write #9, id - 1
    Write #9, Val(factor2(Val(njdh), 0))
    For n1 = 1 To id - 1
        Write #9, n1
        For n2 = 0 To 9
            Print #9, factor3(n1, n2)
        Next n2
    Next n1
    Print #9, Val(TEXToffset)
    Close #9
    msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
    If MsgBox(msg, vbYesNo) = vbYes Then
        Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbNormalFocus
    End If
End Sub

Private Sub Command40_Click()
    If TEXTKstart <> "" And TEXTKend <> "" And TEXTKI <> "" And CDec(Val(TEXToffset)) <> 0 Then '满足条件才计算
        With CommonDialog1
            .FileName = ""                                                      '先清空文件名
            .CancelError = False
            .DialogTitle = "生成中桩坐标的AUTOCAD脚本文件（Z=0）..."
            .DefaultExt = "scr"
            .Filter = "AUTOCAD脚本文件(*.scr)|*.scr"
            .ShowSave
            If Len(.FileName) = 0 Then: Close: Exit Sub
        End With
        Open CommonDialog1.FileName For Output As 5
        Print #5, "setvar"
        Print #5, "osnapcoord"
        Print #5, "1"                                                           '设为1也可以，2=键盘输入优先，脚本例外
        Print #5, "UCS"                                                         '在世界坐标系下画图
        Print #5, "W"
        'Print #5, "UCS"    '在大地测量坐标系下画图
        'Print #5, "N"
        'Print #5, "3"
        'Print #5, "0,0"
        'Print #5, "0,1"
        'Print #5, "1,0"
        
        '新建图层？？？中桩
        Print #5, "-layer"                                                      '首先新建图层
        Print #5, "m"                                                           'm=生成并设置为当前层
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "中桩" '新建图层名称为？？？中桩
        Print #5, "c"                                                           '修改图层颜色
        Print #5, Text116                                                       '1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "中桩" '输入需要修改颜色的图层里程标注
        Print #5, ""                                                            '结束图层设置命令
        '画线
        Print #5, "line"
        Kstart = CDec(Val(TEXTKstart))
        Kend = CDec(Val(TEXTKend))
        T = CDec(Val(TEXTT))
        P = CDec(Val(TEXTP))
        S = CDec(Val(TEXTS))
        Z = CDec(Val(TEXTZ))
        offset = CDec(Val(TEXToffset))
        For K = CDec(Val(TEXTKstart)) To CDec(Val(TEXTKend)) Step CDec(Val(TEXToffset))
            计算 K, x, y, Q, T, P, S, Z, G, H
            Write #5, y, x                                                      '在世界坐标系下画图XY必须互换，Z=0
        Next K
        计算 Kend, x, y, Q, T, P, S, Z, G, H
        Write #5, y, x
        Print #5, ""                                                            '结束line命令
        
        '新建图层里程标注
        Print #5, "-layer"                                                      '首先新建图层
        Print #5, "m"                                                           'm=生成并设置为当前层
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "里程标注" '新建图层名称为里程标注
        Print #5, "c"                                                           '修改图层颜色
        Print #5, Text116                                                       '1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "里程标注" '输入需要修改颜色的图层里程标注
        Print #5, ""                                                            '结束图层设置命令
        '设置文字样式
        Height1 = Val(TEXToffset) / 10                                          '设置文字高度
        length1 = Val(TEXToffset) / 10                                          '设置引线长度
        length2 = Val(TEXToffset) / 25                                          '设置引线偏移距离
        Print #5, "-style"                                                      '建立文字样式
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "里程标注" '文字样式名称
        Print #5, "simfang.ttf"                                                 '字体名称
        Write #5, Height1                                                       '文字高度
        Print #5, "1"                                                           '宽度比例
        Print #5, "0"                                                           '倾斜角度
        Print #5, "n"                                                           '反向显示文字=否
        Print #5, "n"                                                           '颠倒显示文字=否
        
        Kstart = CDec(Val(TEXTKstart))
        Kend = CDec(Val(TEXTKend))
        offset = CDec(Val(TEXToffset))
        For K = CDec(Val(TEXTKstart)) To CDec(Val(TEXTKend)) Step CDec(Val(TEXToffset))
            计算 K, x, y, Q, T, P, S, Z, G, H
            度分秒化为弧度 Q, QHD                                               '把度.分秒格式转化为弧度保存在QHD中
            QHD = 360 - QHD * 180 / 3.14159265358979                            '把弧度转化成度
            Print #5, "line"
            Write #5, y, x                                                      '在世界坐标系下画图XY必须互换，Z=0
            Print #5, "@" & length1 & "<" & Format(QHD, "#0.#######")
            Print #5, ""                                                        '结束line命令
            '标注文字
            Print #5, "text"                                                    '输入单行文字text
            Print #5, "j"                                                       '修改文字格式
            Print #5, "ml"                                                      '左中对齐
            Print #5, "@" & length2 & "<" & Format(QHD, "#0.#######")           '文字起点位置
            Print #5, Format(QHD, "#0.#######")                                 '文字标注方向
            Write #5, K                                                         '里程
            '注意：单行文字命令自动结束的！不需要用回车来结束
        Next K
        
        '求特征点里程和点名保存在数组tezheng中，第0列=里程，第1列=点名
        Dim tezheng(2000, 2) As Variant
        Erase tezheng                                                           '清空数组内容，便于多次输入
        ntz = 1
        tezheng(ntz, 0) = Val(TextQD)
        tezheng(ntz, 1) = "QD"
        ntz = ntz + 1
        For nd = 1 To Val(TextPartNumber)
            If factor(nd, 0) = "直线" Then
                tezheng(ntz, 0) = Val(factor(nd, 1))
                If factor(nd + 1, 0) = "圆" Then tezheng(ntz, 1) = "ZY"
                If factor(nd + 1, 0) = "缓1+圆+缓2" Then tezheng(ntz, 1) = "ZH"
                If factor(nd + 1, 0) = "布劳斯曲线" Then tezheng(ntz, 1) = "ZH"
                If factor(nd + 1, 0) = "麦克康纳尔曲线" Then tezheng(ntz, 1) = "ZH"
                If factor(nd + 1, 0) = "" Then tezheng(ntz, 1) = "ZD"
                ntz = ntz + 1
            ElseIf factor(nd, 0) = "圆" Then
                tezheng(ntz, 0) = Val(factor(nd, 1))
                If factor(nd + 1, 0) = "直线" Then tezheng(ntz, 1) = "YZ"
                If factor(nd + 1, 0) = "缓1+圆+缓2" Then tezheng(ntz, 1) = "YH"
                If factor(nd + 1, 0) = "布劳斯曲线" Then tezheng(ntz, 1) = "YH"
                If factor(nd + 1, 0) = "麦克康纳尔曲线" Then tezheng(ntz, 1) = "YH"
                If factor(nd + 1, 0) = "" Then tezheng(ntz, 1) = "ZD"
                ntz = ntz + 1
            ElseIf factor(nd, 0) = "缓1+圆+缓2" Then
                tezheng(ntz, 0) = Val(factor(nd, 2))
                tezheng(ntz, 1) = "HY"
                ntz = ntz + 1
                If factor(nd, 9) <> 0 Then
                    tezheng(ntz, 0) = Val(factor(nd, 2)) + Val(factor(nd, 9))
                    tezheng(ntz, 1) = "YH"
                    ntz = ntz + 1
                End If
                If Abs(factor(nd, 1) - factor(nd, 2) - factor(nd, 9)) > 0.015 Then '如果Ls2不等于零
                    tezheng(ntz, 0) = Val(factor(nd, 1))
                    If factor(nd + 1, 0) = "直线" Then tezheng(ntz, 1) = "HZ"
                    If factor(nd + 1, 0) = "圆" Then tezheng(ntz, 1) = "HY"
                    If factor(nd + 1, 0) = "缓1+圆+缓2" And Abs(factor(nd, 1) - factor(nd + 1, 2)) < 0.015 Then ntz = ntz - 1
                    If factor(nd + 1, 0) = "缓1+圆+缓2" And Abs(factor(nd, 1) - factor(nd + 1, 2)) > 0.015 Then tezheng(ntz, 1) = "HH"
                    If factor(nd + 1, 0) = "布劳斯曲线" Then tezheng(ntz, 1) = "HH"
                    If factor(nd + 1, 0) = "麦克康纳尔曲线" Then tezheng(ntz, 1) = "HH"
                    ntz = ntz + 1
                End If
                If factor(nd + 1, 0) = "" Then tezheng(ntz - 1, 1) = "ZD"
            ElseIf factor(nd, 0) = "布劳斯曲线" Then
                tezheng(ntz, 0) = Val(factor(nd, 1))
                If factor(nd + 1, 0) = "圆" Then tezheng(ntz, 1) = "HY"
                If factor(nd + 1, 0) = "直线" Then tezheng(ntz, 1) = "HZ"
                If factor(nd + 1, 0) = "" Then tezheng(ntz, 1) = "ZD"
                ntz = ntz + 1
            ElseIf factor(nd, 0) = "麦克康纳尔曲线" Then
                tezheng(ntz, 0) = Val(factor(nd, 1))
                If factor(nd + 1, 0) = "圆" Then tezheng(ntz, 1) = "HY"
                If factor(nd + 1, 0) = "直线" Then tezheng(ntz, 1) = "HZ"
                If factor(nd + 1, 0) = "" Then tezheng(ntz, 1) = "ZD"
                ntz = ntz + 1
            End If
        Next
        
        'MsgBox "ntz=" & ntz-1
        '新建图层特征点标注
        Print #5, "-layer"                                                      '首先新建图层
        Print #5, "m"                                                           'm=生成并设置为当前层
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "特征点标注" '新建图层名称为里程标注
        Print #5, "c"                                                           '修改图层颜色
        Print #5, "3"                                                           '1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "特征点标注" '输入需要修改颜色的图层里程标注
        Print #5, ""                                                            '结束图层设置命令
        
        '设置文字样式
        Height1 = 30 / 10                                                       '设置文字高度
        length1 = 30 / 5                                                        '设置引线长度
        length2 = 30 / 25                                                       '设置引线偏移距离
        Print #5, "-style"                                                      '建立文字样式
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "特征点标注" '文字样式名称
        Print #5, "simfang.ttf"                                                 '字体名称
        Write #5, Height1                                                       '文字高度
        Print #5, "1"                                                           '宽度比例
        Print #5, "0"                                                           '倾斜角度
        Print #5, "n"                                                           '反向显示文字=否
        Print #5, "n"                                                           '颠倒显示文字=否
        
        For ntz2 = 1 To ntz - 1
            K = Val(tezheng(ntz2, 0))
            计算 K, x, y, Q, T, P, S, Z, G, H
            度分秒化为弧度 Q, QHD                                               '把度.分秒格式转化为弧度保存在QHD中
            QHD = (-QHD) * 180 / 3.14159265358979 + 360                         '把弧度转化成度
            Print #5, "line"
            Write #5, y, x                                                      '在世界坐标系下画图XY必须互换，Z=0
            Print #5, "@" & length1 & "<" & Format(QHD, "#0.#######")
            Print #5, ""                                                        '结束line命令
            '标注文字
            Print #5, "text"                                                    '输入单行文字text
            Print #5, "j"                                                       '修改文字格式
            Print #5, "ml"                                                      '左中对齐
            Print #5, "@" & length2 & "<" & Format(QHD, "#0.#######")           '文字起点位置
            Print #5, Format(QHD, "#0.#######")                                 '文字标注方向
            Print #5, tezheng(ntz2, 1) & K                                      '里程
            '注意：单行文字命令自动结束的！不需要用回车来结束
        Next ntz2
        
        '标注AR参数
        For ntz3 = 1 To ntz - 2
            K = (Val(tezheng(ntz3, 0)) + Val(tezheng(ntz3 + 1, 0))) / 2
            计算 K, x, y, Q, T, P, S, Z, G, H                                   'Q单位度.分秒,K所在的区间保存在number全局变量中
            度分秒化为弧度 Q, QHD                                               '把度.分秒格式转化为弧度保存在QHD中
            T = -length2 * 1.2
            P = 90
            S = 0
            Z = 0
            度分秒化为弧度 P, p1
            度分秒化为弧度 Z, Z1
            U = x + T * Cos(QHD + p1)                                           '第一偏距、偏角决定的点的X坐标，中间变量
            V = y + T * Sin(QHD + p1)                                           '第一偏距、偏角决定的点的Y坐标，中间变量
            G = U + S * Cos(QHD + Z1)                                           '第二偏距、偏角决定的点的X坐标，最终结果
            H = V + S * Sin(QHD + Z1)                                           '第二偏距、偏角决定的点的Y坐标，最终结果
            G = Round(G, 3)
            H = Round(H, 3)
            QHD = 450 - QHD * 180 / 3.14159265358979                            '把弧度转化成度
            '标注文字AR
            Print #5, "text"                                                    '输入单行文字text
            Print #5, "j"                                                       '修改文字格式
            Print #5, "BC"                                                      '左中对齐
            Print #5, H & "," & G                                               '文字起点位置
            Print #5, Format(QHD, "#0.#######")                                 '文字标注方向
            If factor(number, 0) = "直线" Then
                Print #5, "F=" & factor(number, 5)
            ElseIf factor(number, 0) = "圆" Then
                Print #5, "R=" & factor(number, 6)
            ElseIf factor(number, 0) = "布劳斯曲线" Then
                Print #5, "Ls=" & factor(number, 7)
            ElseIf factor(number, 0) = "麦克康纳尔曲线" Then
                If Abs(Val(factor(number, 1)) - Val(factor(number, 2))) > 0.05 Then
                    Print #5, "Ls=" & Val(factor(number, 1)) - Val(factor(number, 2))
                Else
                    Print #5, "Ls=" & Val(factor(number, 1)) - Val(factor(number - 1, 1))
                End If
            ElseIf factor(number, 0) = "缓1+圆+缓2" Then
                If K < factor(number, 2) Then
                    Print #5, "A1=" & factor(number, 7)                         '第一缓和曲线
                ElseIf K < factor(number, 2) + factor(number, 9) Then
                    Print #5, "R=" & factor(number, 6) & ",O=" & factor(number, 9) '圆曲线半径和圆曲线长
                Else
                    Print #5, "A2=" & factor(number, 8)                         '第二缓和曲线
                End If
            End If
            '注意：单行文字命令自动结束的！不需要用回车来结束
        Next
        
        
        
        '设置当前图层为0层
        Print #5, "-layer"                                                      '设置图层
        Print #5, "s"                                                           '设置当前图层
        Print #5, "0"                                                           '设置当前图层为0层
        Print #5, ""                                                            '结束-layer命令
        Print #5, "zoom"
        Print #5, "e"
        Close #5
        msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
        If MsgBox(msg, vbYesNo) = vbYes Then
            Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
        End If
    Else                                                                        '如果TEXTKstart为空或TEXTKend为空或TextKi为空或间距为零，就提示重新输入参数
        MsgBox "请输全相关参数,并确保间距不为零且KQD≤Kstart≤Kend≤Ki(max)!", vbCritical, "警告！"
    End If
End Sub

Private Sub Command41_Click()
    anniu = 1                                                                   '如果按下了生成cad脚本按钮就把anniu全局变量赋值为1，在ceyi子程序里通过判断该变量给TPSZ赋值。
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "生成中桩坐标的AUTOCAD脚本文件（Z=0）..."
        .DefaultExt = "scr"
        .Filter = "AUTOCAD脚本文件(*.scr)|*.scr"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
    End With
    Open CommonDialog1.FileName For Output As 11
    Print #11, "setvar"
    Print #11, "osnapcoord"
    Print #11, "1"                                                              '设为1也可以，2=键盘输入优先，脚本例外
    Print #11, "UCS"                                                            '在世界坐标系下画图
    Print #11, "W"
    Print #11, "pdmode"
    Print #11, "32"
    Print #11, "pdsize"
    Print #11, "1"
    
    For ndh = 1 To ndianhao - 1
        K = Val(factor4(ndh, 1))
        T = CDec(Val(factor4(ndh, 2)))
        P = CDec(Val(factor4(ndh, 3)))
        S = CDec(Val(factor4(ndh, 4)))
        Z = CDec(Val(factor4(ndh, 5)))
        计算 K, x, y, Q, T, P, S, Z, G, H
        If CDec(Val(QB)) >= 0 And CDec(Val(QB)) <= 360 Then                     '平行布置要利用桥梁基准切线方位角QB计算,QB在[0,360]之间为平行布置
            Q = QB                                                              'QB单位是度.分秒
            度分秒化为弧度 Q, Q                                                 '把单位转化为弧度进行下一步的计算
            CEYI x, y, Q, T, P, S, Z, G, H
        Else
            '径向布置直接利用K点的切向方位角Q计算
        End If
        Print #11, "point"
        Write #11, H, G                                                         '在世界坐标系下画图XY必须互换，Z=0
    Next ndh
    Print #11, "zoom"
    Print #11, "a"
    Print #11, ""
    Close #11
    msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
    If MsgBox(msg, vbYesNo) = vbYes Then
        Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbNormalFocus
    End If
    anniu = 0
End Sub

Private Sub Command43_Click()
    On Error GoTo Label1
    Dim jj As Double
    Dim str1 As String
    Dim kk() As String
    Dim fenge As String
    fenge = Chr(38) & Chr(34) & "," & Chr(34) & Chr(38)                         '相当于&","&
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "从文件载入竖曲线参数..."
        .DefaultExt = "zas"
        .Filter = "输入的曲线参数(*.zap;*.zas;*.xls;*.xlsx)|*.zap;*.zas;*.xls;*.xlsx|任意交角文件(*.zap)|*.zap|正交多偏距文件(*.zas)|*.zas|excel单偏距文件(*.xls;*.xlsx)|*.xls;*.xlsx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text92 = .FileTitle
    End With
    If Right(CommonDialog1.FileTitle, 3) = "zap" Then                           '如果扩展名是zap
        hangshu = 0
        Open CommonDialog1.FileName For Input As 11
        Do Until EOF(11)
            Line Input #11, shuoming1
            hangshu = hangshu + 1
        Loop
        Seek #11, 1                                                             '文件指针跳转到文件的第一个字节
        msg = MsgBox("单行输出标高？是=单行，否=双行，取消=退出。", vbDefaultButton1 + vbYesNoCancel, "单双行输出到excel")
        If msg = vbCancel Then Close: Exit Sub                                  '如果点了取消就关闭文件并退出子程序
        Line Input #11, shuoming1
        Line Input #11, shuoming2
        Line Input #11, shuoming3
        'If Right(CommonDialog1.FileTitle, 3) = "zap" Then                       '如果扩展名是zap
        LIE = 8
        On Error Resume Next                                                    '如果没有发现excel进程就继续执行下一句：创建excel进程
        Set appexcel = GetObject(, "Excel.application")
        If appexcel Is Nothing Then Set appexcel = CreateObject("excel.application")
        If Err Then                                                             '如果没有安装excel，就用wps表格打开
            Set appexcel = GetObject(, "et.application")
            If appexcel Is Nothing Then Set appexcel = CreateObject("et.application")
        End If
        appexcel.Visible = True
        Set wbmybook = appexcel.Workbooks.Add
        Set exsheet = wbmybook.Worksheets("sheet1")
        exsheet.Activate                                                        '激活
        appexcel.cells.Select
        appexcel.Selection.HorizontalAlignment = xlCenter
        appexcel.Selection.VerticalAlignment = xlCenter
appexcel.Columns("A:A").Select
        appexcel.Selection.NumberFormatLocal = "@"                              '设置第一列为文本格式
        appexcel.sheets("sheet1").Name = "标高表"
        appexcel.Range("A2") = "点号"
        appexcel.Range("B2") = "中桩里程"
        appexcel.Range("C2") = "偏距"
        appexcel.Range("D2") = "偏角"
        appexcel.Range("E2") = "边桩里程"
        appexcel.Range("F2") = "设计高程"
        appexcel.Range("G2") = "横坡"
        appexcel.Range("H2") = "边桩高程"
appexcel.Columns("A:G").columnwidth = 9
appexcel.Range("A1:H1").Select
        appexcel.Selection.merge
        appexcel.Range("A1") = Text92.text
        appexcel.cells(1, 1).Font.Bold = True
        appexcel.cells(1, 1).Font.Size = 16
        jj = 3
        If msg = vbNo Then                                                      '合并单元格
appexcel.Range("A3:A4").merge
appexcel.Range("A3:A4").AutoFill Destination:=appexcel.Range("A3:H4"), Type:=0
appexcel.Range("A3:H4").AutoFill Destination:=appexcel.Range("A3:H" & hangshu * 2 - 4), Type:=0
appexcel.Rows("3:" & hangshu * 2 - 4).RowHeight = 11.25
        End If
        
        
        Do While EOF(11) = False
            Line Input #11, str1                                                '分行读入ktp
            kk = Split(str1, ",")
            If UBound(kk) = 3 Then                                              '依次存储点号，K，T，P
                K = Val(kk(1))
                T = Val(kk(2))
                P = Val(kk(3))
            ElseIf UBound(kk) = 2 Then                                          '依次存储K，T，P
                K = Val(kk(0))
                T = Val(kk(1))
                P = Val(kk(2))
            End If
            计算设计高程 K, J, y, R, E, F, H
            计算边桩高程 K, T, P, KG, H, i, G
            On Error GoTo Label1
            If UBound(kk) = 3 Then
                appexcel.cells(jj, 1) = kk(0)
            Else
                appexcel.cells(jj, 1) = jj - 2
            End If
            appexcel.cells(jj, 2) = K
            appexcel.cells(jj, 3) = T
            appexcel.cells(jj, 4) = P
            appexcel.cells(jj, 5) = KG
            appexcel.cells(jj, 6) = H
            appexcel.cells(jj, 7) = i
            appexcel.cells(jj, 8) = G
            If msg = vbYes Then                                                 '单行输出
                jj = jj + 1
            Else                                                                '双行输出
                jj = jj + 2
            End If
        Loop
        Close #11
    ElseIf Right(CommonDialog1.FileTitle, 3) = "zas" Then                       '如果扩展名是zas
        hangshu = 0
        Open CommonDialog1.FileName For Input As 11
        Do Until EOF(11)
            Line Input #11, shuoming1
            hangshu = hangshu + 1
        Loop
        Seek #11, 1                                                             '文件指针跳转到文件的第一个字节
        msg = MsgBox("单行输出标高？是=单行，否=双行，取消=退出。", vbDefaultButton1 + vbYesNoCancel, "单双行输出到excel")
        If msg = vbCancel Then Close: Exit Sub                                  '如果点了取消就关闭文件并退出子程序
        Line Input #11, shuoming1
        Line Input #11, shuoming2
        Line Input #11, shuoming3
        LIE = 8                                                                 '列=8
        On Error Resume Next                                                    '如果没有发现excel进程就继续执行下一句：创建excel进程
        Set appexcel = GetObject(, "Excel.application")
        If appexcel = "" Then Set appexcel = CreateObject("excel.application")
        appexcel.Visible = True
        Set wbmybook = appexcel.Workbooks.Add
        Set exsheet = wbmybook.Worksheets("sheet1")
        exsheet.Activate                                                        '激活
        appexcel.cells.Select
        appexcel.Selection.HorizontalAlignment = xlCenter
        appexcel.Selection.VerticalAlignment = xlCenter
        appexcel.sheets("sheet1").Name = "标高表"
        appexcel.Range("A2") = "里程"
        appexcel.Range("B2") = "左宽"
        appexcel.Range("C2") = "右宽"
        appexcel.Range("D2") = "左幅标高"
        appexcel.Range("E2") = "中标高"
        appexcel.Range("F2") = "右幅高程"
        appexcel.Range("G2") = "左幅横坡"
        appexcel.Range("H2") = "右幅横坡"
        appexcel.Range("J2") = "生成设计线sjx文件(以填方1:1.5坡脚为例)"
appexcel.Columns("A:H").columnwidth = 9
appexcel.Columns("I:I").columnwidth = 3
appexcel.Columns("J:J").columnwidth = 80
appexcel.Range("A1:H1").Select
        appexcel.Selection.merge
        appexcel.Range("A1") = Text92.text
        appexcel.cells(1, 1).Font.Bold = True
        appexcel.cells(1, 1).Font.Size = 16
        jj = 3
        If msg = vbNo Then                                                      '合并单元格
appexcel.Range("A3:A4").merge
appexcel.Range("A3:A4").AutoFill Destination:=appexcel.Range("A3:H4"), Type:=0
appexcel.Range("A3:H4").AutoFill Destination:=appexcel.Range("A3:H" & hangshu * 2 - 4), Type:=0
appexcel.Rows("3:" & hangshu * 2 - 4).RowHeight = 10
appexcel.Range("J3:J4").merge
        End If
        Do While EOF(11) = False
            Line Input #11, str1                                                '分行读入k,左右宽度
            kk = Split(str1, ",")
            If UBound(kk) >= 0 Then
                For I1 = 0 To UBound(kk)
                    K = Val(kk(0))
                    On Error GoTo Label1
                    appexcel.cells(jj, 1) = kk(0)
                    
                    If UBound(kk) >= 2 Then
                        appexcel.cells(jj, 2) = kk(1)
                        appexcel.cells(jj, 3) = kk(UBound(kk))
                        计算设计高程 K, J, y, R, E, F, H
                        T = Val(kk(1))
                        P = 90
                        计算边桩高程 K, T, P, KG, H, i, G
                        appexcel.cells(jj, 4) = G
                        appexcel.cells(jj, 5) = H
                        appexcel.cells(jj, 7) = Round(i, 4)
                        T = Val(kk(UBound(kk)))
                        P = 90
                        计算边桩高程 K, T, P, KG, H, i, G
                        appexcel.cells(jj, 6) = G
                        appexcel.cells(jj, 8) = Round(i, 4)
                    End If
                    
                Next
            End If
            If msg = vbYes Then                                                 '单行输出
                jj = jj + 1
            Else                                                                '双行输出
                jj = jj + 2
            End If
        Loop
        
        '生成sjx文件
        If Val(Text43) <> 0 Then
            appexcel.cells(3, 10) = "=A3" & fenge & "E3" & fenge & _
            "B3-15" & fenge & "round(d3-E3-10,3)" & fenge & _
            "B3" & fenge & "round(D3-E3,3)" & fenge & _
            -Val(Text43) & fenge & 0 & fenge & Val(Text43) & fenge & 0 & fenge & _
            "C3" & fenge & "round(F3-E3,3)" & fenge & _
            "C3+15" & fenge & "round(F3-E3-10,3)"
        Else
            appexcel.cells(3, 10) = "=A3" & fenge & "E3" & fenge & _
            "B3-15" & fenge & "round(D3-E3-10,3)" & fenge & _
            "B3" & fenge & "round(D3-E3,3)" & fenge & _
0 & fenge & 0 & fenge _
            & "C3" & fenge & "round(F3-E3,3)" & fenge & _
            "C3+15" & fenge & "round(F3-E3-10,3)"
        End If
        If msg = vbYes Then                                                     '单行
appexcel.Range("j3:j3").AutoFill Destination:=appexcel.Range("j3:j" & jj - 1), Type:=0
        ElseIf msg = vbNo Then                                                  '双行
appexcel.Range("j3:j4").AutoFill Destination:=appexcel.Range("j3:j" & hangshu * 2 - 4), Type:=0
        End If
        
        Close #11
    Else                                                                        '如果扩展名是xls或者xlsx
        Dim xlapp As Object
        Dim xlBook As Object
        Dim xlSheet As Object
        On Error Resume Next
        Set xlapp = GetObject(, "Excel.application")
        If xlapp Is Nothing Then Set xlapp = CreateObject("excel.application")
        If Err Then                                                             '如果没有安装excel，就用wps表格打开
            Set xlapp = GetObject(, "et.application")
            If xlapp Is Nothing Then Set xlapp = CreateObject("et.application")
        End If
        Set xlBook = xlapp.Workbooks(CommonDialog1.FileTitle)                   '假设xlbook为选定的xls工作簿
        If xlBook Is Nothing Then Set xlBook = xlapp.Workbooks.Open(CommonDialog1.FileName)
        '如果xlbook不是选定的工作簿就打开那个工作簿
        Set xlSheet = xlBook.activesheet
        置顶 Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)    '调用模块里的api把excel进程置顶
        xlapp.Visible = True
        jj = 3
        ndianhao = 1
Label2:
        K = (xlSheet.cells(jj, 2))
        T = (xlSheet.cells(jj, 3))
        P = (xlSheet.cells(jj, 4))
        计算设计高程 K, J, y, R, E, F, H
        计算边桩高程 K, T, P, KG, H, i, G
        On Error GoTo Label1
        xlSheet.cells(jj, 5) = KG
        xlSheet.cells(jj, 6) = H
        xlSheet.cells(jj, 7) = i
        xlSheet.cells(jj, 8) = G
        jj = jj + xlSheet.cells(jj, 1).MergeArea.Rows.Count                     '单元格的合并区域的行计数
        ndianhao = ndianhao + 1
        If xlSheet.cells(jj, 1) <> "" And xlSheet.cells(jj, 2) <> "" Then GoTo Label2
        Set xlapp = Nothing
    End If
    
    appexcel.Range(appexcel.cells(2, 1), appexcel.cells(jj - 1, LIE)).Select
'appexcel.Range("A2:H" + LTrim(Str(jj - 1))).Select
    appexcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    appexcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With appexcel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    appexcel.cells(1, 7).Select
    
    
    '设置打印区域
    appexcel.activesheet.PageSetup.PrintArea = "$A:$H"
    
    '设置顶端标题行和页脚的第几页共几页
    'With appexcel.ActiveSheet.PageSetup
    '        .PrintTitleRows = "$1:$2"
    '        .PrintTitleColumns = ""
    '    End With
    '    appexcel.ActiveSheet.PageSetup.PrintArea = "$A:$H"
    '    With appexcel.ActiveSheet.PageSetup
    '        .LeftHeader = ""
    '        .CenterHeader = ""
    '        .RightHeader = ""
    '        .LeftFooter = ""
    '        .CenterFooter = "第 &P 页，共 &N 页"
    '        .RightFooter = ""
    '        .LeftMargin = appexcel.Application.InchesToPoints(0.748031496062992)
    '        .RightMargin = appexcel.Application.InchesToPoints(0.748031496062992)
    '        .TopMargin = appexcel.Application.InchesToPoints(0.866141732283465)
    '        .BottomMargin = appexcel.Application.InchesToPoints(0.866141732283465)
    '        .HeaderMargin = appexcel.Application.InchesToPoints(0.511811023622047)
    '        .FooterMargin = appexcel.Application.InchesToPoints(0.511811023622047)
    '        .PrintHeadings = False
    '        .PrintGridlines = False
    '        .PrintComments = xlPrintNoComments
    '        .CenterHorizontally = True
    '        .CenterVertically = False
    '        .Orientation = xlPortrait
    '        .Draft = False
    '        .PaperSize = xlPaperA4
    '        .FirstPageNumber = xlAutomatic
    '        .Order = xlDownThenOver
    '        .BlackAndWhite = False
    '        .Zoom = 100
    '        .PrintErrors = xlPrintErrorsDisplayed
    '    End With
    
Label1: Close
End Sub

Private Sub Command42_Click()
    anniu = 1                                                                   '如果按下了导出到excel按钮就把anniu全局变量赋值为1，在ceyi子程序里通过判断该变量给TPSZ赋值。
    Dim jj As Double
    Dim ex As Object
    Dim exwbook As Object
    Dim exsheet As Object
    
    On Error Resume Next
    Set appexcel = GetObject(, "Excel.application")
    If appexcel Is Nothing Then Set appexcel = CreateObject("excel.application")
    If Err Then                                                                 '如果没有安装excel，就用wps表格打开
        Set appexcel = GetObject(, "et.application")
        If appexcel Is Nothing Then Set appexcel = CreateObject("et.application")
    End If
    appexcel.Visible = True
    Set exwbook = appexcel.Workbooks.Add
    Set exsheet = exwbook.Worksheets("sheet1")
    exsheet.Activate                                                            '激活
    appexcel.sheets("sheet1").Name = "坐标表"
    appexcel.cells.Select
    appexcel.Selection.HorizontalAlignment = xlCenter
    appexcel.Selection.VerticalAlignment = xlCenter
    appexcel.cells(1, 1).Font.Bold = True
    appexcel.cells(1, 1).Font.Size = 16
appexcel.Columns("A:A").Select
    appexcel.Selection.NumberFormatLocal = "@"                                  '设置第一列为文本格式
appexcel.Columns("A:A").columnwidth = 6
appexcel.Columns("B:F").columnwidth = 9.25
appexcel.Columns("G:H").columnwidth = 11
appexcel.Range("A1:H1").Select
    appexcel.Selection.merge
    appexcel.Range("a1") = Text93.text
appexcel.Range("A2:D2").Select
    appexcel.Selection.merge
    appexcel.Range("a2") = "基准切线方位角(度.分秒)"
appexcel.Range("E2:H2").Select
    appexcel.Selection.merge
    If Val(QB) >= 0 And Val(QB) <= 360 Then
        appexcel.Range("E2") = Val(QB)
    Else
        appexcel.Range("E2") = "各点中线切线方位角Q"
    End If
    appexcel.Range("a3") = "点号"
    appexcel.Range("b3") = "里程"
    appexcel.Range("c3") = "第一偏距T"
    appexcel.Range("d3") = "第一偏角P"
    appexcel.Range("e3") = "第二偏距S"
    appexcel.Range("f3") = "第二偏角Z"
    appexcel.Range("G3") = "X"
    appexcel.Range("H3") = "Y"
    
    jj = 4
    For ndh = 1 To ndianhao - 1
        K = Val(factor4(ndh, 1))
        T = CDec(Val(factor4(ndh, 2)))
        P = CDec(Val(factor4(ndh, 3)))
        S = CDec(Val(factor4(ndh, 4)))
        Z = CDec(Val(factor4(ndh, 5)))
        计算 K, x, y, Q, T, P, S, Z, G, H
        If Val(QB) >= 0 And Val(QB) <= 360 Then                                 '平行布置要利用桥梁基准切线方位角QB计算,QB在[0,360]之间为平行布置
            Q = QB                                                              'QB单位是度.分秒
            度分秒化为弧度 Q, Q                                                 '把单位转化为弧度进行下一步的计算
            CEYI x, y, Q, T, P, S, Z, G, H
        Else
            '径向布置直接利用K点的切向方位角Q计算
        End If
        On Error GoTo Label1
        appexcel.cells(jj, 1) = factor4(ndh, 0)
        appexcel.cells(jj, 2) = K
        appexcel.cells(jj, 3) = T
        appexcel.cells(jj, 4) = P
        appexcel.cells(jj, 5) = S
        appexcel.cells(jj, 6) = Z
        appexcel.cells(jj, 7) = G
        appexcel.cells(jj, 8) = H
        jj = jj + 1
    Next ndh
    
    'appexcel.Range(appexcel.cells(2, 1), appexcel.cells(jj - 1, 8)).Select
appexcel.Range("A2:H" + LTrim(Str(jj - 1))).Select
    appexcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    appexcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With appexcel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    appexcel.cells(1, 10).Select
    
    '设置顶端标题行和页脚的第几页共几页
    With appexcel.activesheet.PageSetup
        .PrintTitleRows = "$1:$3"
        .PrintTitleColumns = ""
    End With
    appexcel.activesheet.PageSetup.PrintArea = ""
    With appexcel.activesheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = "第 &P 页，共 &N 页"
        .RightFooter = ""
        .LeftMargin = appexcel.application.InchesToPoints(0.748031496062992)
        .RightMargin = appexcel.application.InchesToPoints(0.748031496062992)
        .TopMargin = appexcel.application.InchesToPoints(0.866141732283465)
        .BottomMargin = appexcel.application.InchesToPoints(0.866141732283465)
        .HeaderMargin = appexcel.application.InchesToPoints(0.511811023622047)
        .FooterMargin = appexcel.application.InchesToPoints(0.511811023622047)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
    End With
    
Label1:
    anniu = 0
End Sub

Private Sub Command44_Click()
    On Error GoTo Label1
    Dim dianhao As String
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "从文件载入平曲线参数..."
        .DefaultExt = "wsc"
        .Filter = "输入的曲线参数(*.wsc;*.xls;*.xlsx)|*.wsc;*.xls;*.xlsx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text93 = .FileTitle
    End With
    If Right(CommonDialog1.FileName, 3) = "wsc" Then
        Open CommonDialog1.FileName For Input As 10
        Line Input #10, shuoming1
        Line Input #10, shuoming2
        Line Input #10, shuoming3
        Input #10, QB
        Q = QB
        ndianhao = 1
        '清空所求点里程数组便于多次载入
        Erase factor4
        Do While EOF(10) = False
            Input #10, dianhao, factor4(ndianhao, 1), factor4(ndianhao, 2), factor4(ndianhao, 3), factor4(ndianhao, 4), factor4(ndianhao, 5)
            factor4(ndianhao, 0) = dianhao                                      '首先把点号读入到变量dianhao中，然后赋值给变体数组，这样输入不用加引号。
            ndianhao = ndianhao + 1
        Loop
        Close #10
    Else
        '如果是excel格式
        anniu = 1
        Dim xlapp As Object
        Dim xlBook As Object
        Dim xlSheet As Object
        On Error Resume Next
        Set xlapp = GetObject(, "Excel.application")
        If xlapp Is Nothing Then Set xlapp = CreateObject("excel.application")
        If Err Then                                                             '如果没有安装excel，就用wps表格打开
            Set xlapp = GetObject(, "et.application")
            If xlapp Is Nothing Then Set xlapp = CreateObject("et.application")
        End If
        Set xlBook = xlapp.Workbooks(CommonDialog1.FileTitle)                   '假设xlbook为选定的xls工作簿
        If xlBook Is Nothing Then Set xlBook = xlapp.Workbooks.Open(CommonDialog1.FileName) '如果xlbook不是选定的工作簿就打开那个工作簿
        Set xlSheet = xlBook.activesheet
        置顶 Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)    '调用模块里的api把excel进程置顶
        xlapp.Visible = True
        QB = xlSheet.cells(2, 5)
        jj = 4
        ndianhao = 1
        Erase factor4
Label2:
        K = (xlSheet.cells(jj, 2))
        T = (xlSheet.cells(jj, 3))
        P = (xlSheet.cells(jj, 4))
        S = (xlSheet.cells(jj, 5))
        Z = (xlSheet.cells(jj, 6))
        factor4(ndianhao, 0) = xlSheet.cells(jj, 1)                             '点号
        factor4(ndianhao, 1) = xlSheet.cells(jj, 2)                             'K
        factor4(ndianhao, 2) = xlSheet.cells(jj, 3)                             'T
        factor4(ndianhao, 3) = xlSheet.cells(jj, 4)                             'P
        factor4(ndianhao, 4) = xlSheet.cells(jj, 5)                             'S
        factor4(ndianhao, 5) = xlSheet.cells(jj, 6)                             'Z
        计算 K, x, y, Q, T, P, S, Z, G, H
        If Val(QB) >= 0 And Val(QB) <= 360 Then                                 '平行布置要利用桥梁基准切线方位角QB计算,QB在[0,360]之间为平行布置
            Q = QB                                                              'QB单位是度.分秒
            度分秒化为弧度 Q, Q                                                 '把单位转化为弧度进行下一步的计算
            CEYI x, y, Q, T, P, S, Z, G, H
        Else
            '径向布置直接利用K点的切向方位角Q计算
        End If
        On Error GoTo Label1
        xlSheet.cells(jj, 7) = G
        xlSheet.cells(jj, 8) = H
        jj = jj + xlSheet.cells(jj, 1).MergeArea.Rows.Count                     '单元格的合并区域的行计数
        ndianhao = ndianhao + 1
        If xlSheet.cells(jj, 1) <> "" And xlSheet.cells(jj, 2) <> "" Then GoTo Label2
        Set xlapp = Nothing
    End If
Label1: anniu = 0
End Sub

Private Sub Command45_Click()
    Dim file1 As String
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "编辑现有的wsc文件..."
        .DefaultExt = "wsc"
        .Filter = "输入的曲线参数(*.wsc)|*.wsc"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        file1 = .FileName
    End With
    If Dir(file1) = "" Then
        '如果文件不存在就先把三行说明行写入文件
        Open CommonDialog1.FileName For Output As 1
        Print #1, "说明：该文件前3行是说明行不要删除，从第4行开始是数据行。数据文件格式如下："
        Print #1, "第4行输入桥梁涵洞基准切线方位角，格式为度.分秒，例如230度04分05秒输入230.0405，[0,360]度之间表示平行布置，否则径向布置。"
        Print #1, "下一行输入点号,里程,第一偏距,第一偏角,第二偏距,第二偏角,以英文逗号隔开，回车后输入下一行。"
        Close #1
    End If
    '如果文件存在就直接打开：
    Shell "notepad.exe " & """" & file1 & """", vbMaximizedFocus
End Sub

Private Sub Command46_Click()
    Dim file1 As String
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "编辑现有的zap文件..."
        .DefaultExt = "zap"
        .Filter = "输入的曲线参数(*.zap)|*.zap"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        file1 = .FileName
    End With
    If Dir(file1) = "" Then
        '如果文件不存在就先把三行说明行写入文件
        Open CommonDialog1.FileName For Output As 1
        Print #1, "说明：该文件前3行是说明行不要删除，从第4行开始是数据行。数据文件格式如下："
        Print #1, "依次输入点号,里程,偏距,偏角,以英文逗号隔开，回车后输入下一行。"
        Print #1, "注意偏角输入的格式为度.分秒，例如80度04分05秒输入80.0405。"
        Close #1
    End If
    '如果文件存在就直接打开：
    Shell "notepad.exe " & """" & file1 & """", vbMaximizedFocus
End Sub

Private Sub Command47_Click()
    '生成卡西欧5800计算器主程序
    msg1 = MsgBox("正算=已知里程K求XY，反算=已知XY求里程K，需要进行反算吗？" & vbCrLf & "是=只需要正算，一个主程序；否=需要正反算，正反算主程序+中桩子程序，取消=退出", vbYesNoCancel, "正反算确认")
    If msg1 = vbCancel Then
        Exit Sub                                                                '如果取消直接退出该过程。
    ElseIf msg1 = vbYes Then                                                    '只进行正算，1个程序
        With CommonDialog1
            .FileName = ""                                                      '先清空文件名
            .CancelError = False
            .DialogTitle = "生成卡西欧5800可编程计算器程序..."
            .DefaultExt = "txt"
            .Filter = "卡西欧5800可编程计算器程序(*.txt)|*.txt"
            .ShowSave
            If Len(.FileName) = 0 Then: Close: Exit Sub
            'Text2 = .FileTitle
        End With
        Open CommonDialog1.FileName For Output As 12
        Print #12, Text2
        Print #12, "计算全线中边桩坐标主程序。子程序ZX,HUAN,YUAN,PQX不可缺少。所有角度F格式是度分秒要正确输入！！！"
        Print #12, "K为所求点里程，XYGH中边桩坐标，Q切线方位角，TPSZ第一偏距偏角第二偏距偏角。"
        Print #12, "Deg:?K↙"
        For n1 = 1 To Val(TextPartNumber)
            If factor(n1, 0) = "直线" Then
                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & "Prog ""ZX"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "圆" Then
                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & "Prog ""YUAN"":" _
                & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "缓1+圆+缓2" Then
                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & factor(n1, 7) & "→A:" _
                & factor(n1, 8) & "→B:" & factor(n1, 9) & "→O:" & "Prog ""PQX"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "布劳斯曲线" Then
                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & factor(n1, 7) & "→A:" _
                & "Prog ""BLOSS"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "麦克康纳尔曲线" Then
                Ls = CDec(Val(factor(n1, 1)) - Val(factor(n1, 2)))              '对于MC1，Ls=分段终点Khy-Kzh
                If Abs(Ls) < 0.05 Then Ls = Val(factor(n1, 1)) - Val(factor(n1 - 1, 1))
                v0 = Format(CDec(Val(factor(n1, 7)) / 3.6), "##0.####")         '设计平衡车速m/s，(120-300)/3.6
                t0 = Format(CDec(Atn(Val(factor(n1, 8)))) * 180 / pai, "##0.####") '缓和曲线起点处的路面横向倾角atn(0.7%-0.8%)，弧度
                g0 = CDec((Val(factor(n1, 9))))                                 '重力加速度m/s2
                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & v0 & "→A:" _
                & t0 & "→B:" & Ls & "→O:Prog ""MC"":Goto 1:IfEnd↙"
            End If
        Next
        Print #12, "Lbl 1:Prog ""CEYI"""
        Close #12
    Else                                                                        '正反算和中桩共3个程序
        With CommonDialog1
            .FileName = ""                                                      '先清空文件名
            .CancelError = False
            .DialogTitle = "生成卡西欧5800可编程计算器程序..."
            .DefaultExt = "txt"
            .Filter = "卡西欧5800可编程计算器程序(*.txt)|*.txt"
            .ShowSave
            If Len(.FileName) = 0 Then: Close: Exit Sub
            'Text2 = .FileTitle
        End With
        Open CommonDialog1.FileName For Output As 12
        Print #12, Text2
        Print #12, "计算全线中边桩坐标主程序。子程序ZX,HUAN,YUAN,PQX不可缺少。所有角度F格式是度分秒要正确输入！！！"
        Print #12, "K为所求点里程，XYGH中边桩坐标，Q切线方位角，TPSZ第一偏距偏角第二偏距偏角。"
        Print #12, "正算主程序:" & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "ZS"
        Print #12, "?K:"
        Print #12, "Prog """ & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "ZZ"":Prog ""CEYI"""
        Print #12, vbCrLf & vbCrLf
        
        
        
        Print #12, "反算主程序:" & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "FS"
        Print #12, "?M:?N:?K:"
        Print #12, "Lbl 1:Prog """ & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "ZZ"":"
        Print #12, "Pol(M-X,N-Y):        (9750和9860计算器需附加List Ans[1]→I:List Ans[2]→J:J<0=>J+360→J:)"
        Print #12, "If Abs (I*cos (J-Q))>0.005:Then "
        Print #12, "K+I*0.382*cos (J-Q)→K:Goto 1:"
        Print #12, "Else ""K="":K◢"
        Print #12, "I*sin (J-Q)/Abs (sin (J-Q))→T:""T="":T◢"
        Print #12, "IfEnd:↙"
        Print #12, "Prog """ & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "ZS"""
        Print #12, vbCrLf & vbCrLf
        
        
        
        Print #12, "中桩子程序:" & Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "ZZ"
        For n1 = 1 To Val(TextPartNumber)
            If factor(n1, 0) = "直线" Then
                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & "Prog ""ZX"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "圆" Then
                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & "Prog ""YUAN"":" _
                & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "缓1+圆+缓2" Then
                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & factor(n1, 7) & "→A:" _
                & factor(n1, 8) & "→B:" & factor(n1, 9) & "→O:" & "Prog ""PQX"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "布劳斯曲线" Then
                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & factor(n1, 7) & "→A:" _
                & "Prog ""BLOSS"":" & "Goto 1:IfEnd↙"
            ElseIf factor(n1, 0) = "麦克康纳尔曲线" Then
                Ls = CDec(Val(factor(n1, 1)) - Val(factor(n1, 2)))              '对于MC1，Ls=分段终点Khy-Kzh
                If Abs(Ls) < 0.05 Then Ls = Val(factor(n1, 1)) - Val(factor(n1 - 1, 1))
                v0 = Format(CDec(Val(factor(n1, 7)) / 3.6), "##0.####")         '设计平衡车速m/s，(120-300)/3.6
                t0 = Format(CDec(Atn(Val(factor(n1, 8)))) * 180 / pai, "##0.####") '缓和曲线起点处的路面横向倾角atn(0.7%-0.8%)，弧度
                g0 = CDec((Val(factor(n1, 9))))                                 '重力加速度m/s2
                Print #12, "If K≤" & factor(n1, 1) & ":Then " & factor(n1, 2) & "→J:" & factor(n1, 3) & "→C:" _
                & factor(n1, 4) & "→D:" & Format(factor(n1, 5) * 10000, "##0°##°#0.####°") & "→F:" & factor(n1, 6); "→R:" & v0 & "→A:" _
                & t0 & "→B:" & Ls & "→O:Prog ""MC"":Goto 1:IfEnd↙"
            End If
        Next
        Print #12, "Lbl 1:"
        Close #12
    End If
    msg2 = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
    If MsgBox(msg2, vbYesNo, "查看CASIO 5800程序清单") = vbYes Then
        Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
    End If
End Sub

Private Sub Command48_Click()
    '生成卡西欧5800计算器主程序
    On Error Resume Next
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "生成卡西欧5800可编程计算器程序..."
        .DefaultExt = "txt"
        .Filter = "卡西欧5800可编程计算器程序(*.txt)|*.txt"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
        'Text2 = .FileTitle
    End With
    Open CommonDialog1.FileName For Output As 13
    Print #13, Text41
    Print #13, "计算全线任意交角中边桩高程主程序，子程序GC不可缺少。边桩高程考虑线性超高和三次抛物线超高。"
    Print #13, "K所求点里程，T偏距左负右正，P偏角，H设计高程，G边桩高程。"
    Print #13, "?K:?T:?P↙"
    Print #13, "K+T*cos(P)→V↙"
    For n1 = 2 To Val(Text24) - 1
        Print #13, "If V≤" & shu(n1, 0) & ":Then " & shu(n1, 1) & "→J:" & shu(n1, 2) & "→Y:" _
        & shu(n1, 3) & "→R:" & Format(shu(n1, 4) * 100, "0.#####") & "→E:" & Format(shu(n1, 5) * 100, "0.#####"); "→F:" & "Goto 1:IfEnd↙"
    Next
    Print #13, "Lbl 1:Prog ""GC""◢"
    '边桩高程计算，考虑超高
    hengpo(0, 0) = hengpo(1, 0)
    hengpo(0, 1) = hengpo(1, 1)
    hengpo(0, 2) = hengpo(1, 2)
    'If K2 <> K1 Then I = I1 + (KG - K1) * (I2 - I1) / (K2 - K1)
    'If K2 = K1 Then I = I1
    'If K4 <> K3 Then I = I3 + (KG - K3) * (I4 - I3) / (K4 - K3)
    'If K4 = K3 Then I = I3
    'ts = Val(Text43)
    'G = H - I * (Abs(T * Sin(P1)) - Abs(ts))
    
    For n1 = 1 To Val(Text33)
        K1 = CDec(hengpo(n1 - 1, 0))
        K2 = CDec(hengpo(n1, 0))
        K3 = CDec(hengpo(n1 - 1, 0))
        K4 = CDec(hengpo(n1, 0))
        I1 = CDec(hengpo(n1 - 1, 1))
        i2 = CDec(hengpo(n1, 1))
        I3 = CDec(hengpo(n1 - 1, 2))
        i4 = CDec(hengpo(n1, 2))
        If hengpo(n1, 3) = 0 Then                                               '如果是线性超高
            If i2 <> I1 And I3 <> i4 And K2 <> K1 And K3 <> K4 Then
                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                I1 & "+(V-" & K1 & ")*(" & (i2 - I1) & ")/(" & (K2 - K1) & ")→A:" & _
                I3 & "+(V-" & K3 & ")*(" & (i4 - I3) & ")/(" & (K4 - K3) & ")→B:" & _
                "Goto 2:IfEnd↙"
            ElseIf i2 <> I1 And I3 = i4 And K2 <> K1 Then
                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                I1 & "+(V-" & K1 & ")*(" & (i2 - I1) & ")/(" & (K2 - K1) & ")→A:" & _
                I3 & "→B:" & _
                "Goto 2:IfEnd↙"
            ElseIf i2 = I1 And I3 <> i4 And K3 <> K4 Then
                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                I1 & "→A:" & _
                I3 & "+(V-" & K3 & ")*(" & (i4 - I3) & ")/(" & (K4 - K3) & ")→B:" & _
                "Goto 2:IfEnd↙"
            Else
                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                I1 & "→A:" & _
                I3 & "→B:" & _
                "Goto 2:IfEnd↙"
            End If
        ElseIf hengpo(n1, 3) = 1 Then                                           '如果是三次抛物线超高
            If i2 <> I1 And I3 <> i4 And K2 <> K1 And K3 <> K4 Then
                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                "(K-" & K1 & ")/(" & (K2 - K1) & ")→I:" & _
                I1 & "+(3×I^2-2×I^3)×(" & (i2 - I1) & ")→A:" & _
                I3 & "+(3×I^2-2×I^3)×(" & (i4 - I3) & ")→B:" & _
                "Goto 2:IfEnd↙"
            ElseIf i2 <> I1 And I3 = i4 And K2 <> K1 Then
                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                "(K-" & K1 & ")/(" & (K2 - K1) & ")→I:" & _
                I1 & "+(3×I^2-2×I^3)×(" & (i2 - I1) & ")→A:" & _
                I3 & "→B:" & _
                "Goto 2:IfEnd↙"
            ElseIf i2 = I1 And I3 <> i4 And K3 <> K4 Then
                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                "(K-" & K1 & ")/(" & (K2 - K1) & ")→I:" & _
                I1 & "→A:" & _
                I3 & "+(3×I^2-2×I^3)×(" & (i4 - I3) & ")→B:" & _
                "Goto 2:IfEnd↙"
            Else
                Print #13, "If K≤" & hengpo(n1, 0) & ":Then " & _
                I1 & "→A:" & _
                I3 & "→B:" & _
                "Goto 2:IfEnd↙"
            End If
        End If
    Next
    Print #13, "Lbl 2:If T≤0:Then A→I:Else B→I:IfEnd↙"
    Print #13, "H-I×(Abs (T×sin (P))-" & Val(Text43) & ")→G:""G="":G"
    Close #13
    msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗?"
    If MsgBox(msg, vbYesNo) = vbYes Then
        Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
    End If
End Sub

Private Sub Combo7_Click()
    载入宽度参数
End Sub
Sub 载入宽度参数()
    Text96 = CDec(Val(kuandu(Val(Combo7), 0)))                                  'kuandu(段数,0列)存储分段终点里程KI
    Text95 = CDec(Val(kuandu(Val(Combo7), 1)))                                  'kuandu(段数,1列)存储变坡点左幅宽度
    'Text94 = CDec(Val(kuandu(Val(Combo7), 2)))   'kuandu(段数,2列)存储变坡点右幅宽度
    If Val(kuandu(Val(Combo7), 2)) = 0 Then Option4.Value = True
    If Val(kuandu(Val(Combo7), 2)) = 1 Then Option3.Value = True
End Sub

Private Sub Command49_Click()
    保存宽度参数
    If Combo7.ListIndex < Combo7.ListCount - 1 Then Combo7.ListIndex = Combo7.ListIndex + 1
End Sub
Sub 保存宽度参数()
    kuandu(Val(Combo7), 0) = CDec(Val(Text96))                                  'kuandu(段数,0列)存储分段终点里程KI
    kuandu(Val(Combo7), 1) = CDec(Val(Text95))                                  'kuandu(段数,1列)存储变坡点左横坡
    'kuandu(Val(Combo7), 2) = CDec(Val(Text94))  'kuandu(段数,2列)存储变坡点右横坡
    If Option4.Value = True Then kuandu(Val(Combo7), 2) = 0                     'kuandu(段数,3列)=0线形，=1抛物线
    If Option3.Value = True Then kuandu(Val(Combo7), 2) = 1                     'kuandu(段数,3列)=0线形，=1抛物线
End Sub
Private Sub Command50_Click()
    '保存宽度参数
    Dim n1, n2 As Integer                                                       '循环用的变量
    With CommonDialog1
        .FileName = ""
        .CancelError = False
        .DialogTitle = "保存宽度分段参数到文件..."
        .DefaultExt = "kdu"
        .Filter = "宽度参数(*.kdu)|*.kdu"
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
    End With
    Open CommonDialog1.FileName For Output As 15
    Write #15, Val(Text97), Val(Text104), Val(Text106), Val(Text98)
    For n1 = 1 To Val(Text97)
        Write #15, kuandu(n1, 0), kuandu(n1, 1), kuandu(n1, 2)
    Next n1
    Write #15, Val(Text102)
    For n1 = 1 To Val(Text102)
        Write #15, kuandu2(n1, 0), kuandu2(n1, 1), kuandu2(n1, 2)
    Next n1
    Close #15
End Sub

Private Sub Command51_Click()
    '载入宽度分段数据
    Dim n1, n2 As Integer                                                       '循环用的变量
    With CommonDialog1
        .FileName = ""
        .CancelError = False
        .DialogTitle = "从文件载入宽度分段参数..."
        .DefaultExt = "kdu"
        .Filter = "宽度参数(*.kdu)|*.kdu"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        Text99 = .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 14
    Input #14, partnumber3, qidian, jianju, xiaoshu
    Text97 = partnumber3
    Text104 = qidian
    Text106 = jianju
    Text98 = xiaoshu
    Erase kuandu                                                                '清空数组便于多次载入
    For n1 = 1 To partnumber3
        Input #14, kuandu(n1, 0), kuandu(n1, 1), kuandu(n1, 2)
    Next n1
    
    Input #14, partnumber4
    Text102 = partnumber4
    Erase kuandu2                                                               '清空数组便于多次载入
    For n2 = 1 To partnumber4
        Input #14, kuandu2(n2, 0), kuandu2(n2, 1), kuandu2(n2, 2)
    Next n2
    载入宽度参数
    载入宽度参数2
    If kuandu(n1 - 1, 0) >= kuandu2(n2 - 1, 0) Then
        Text105 = kuandu(n1 - 1, 0)
    Else
        Text105 = kuandu2(n2 - 1, 0)
    End If
    'If EOF(14) = False Then
    'Input #1, jianju
    'Text106 = jianju
    'End If
    Close #14
    'Saveflag = 1
End Sub
Private Sub Command52_Click()                                                   '编辑里程文件
    Dim file1 As String
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "编辑现有的lch文件..."
        .DefaultExt = "lch"
        .Filter = "输入的曲线参数(*.lch)|*.lch"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        file1 = .FileName
    End With
    If Dir(file1) = "" Then
        '如果文件不存在就先把三行说明行写入文件
        Open CommonDialog1.FileName For Output As 1
        Print #1, "说明：该文件前3行是说明行不能删除，从第4行开始是数据行。数据文件格式如下："
        Print #1, "里程1"
        Print #1, "里程2（以此类推，回车输入下一行）"
        Close #1
    End If
    '如果文件存在就直接打开：
    Shell "notepad.exe " & """" & file1 & """", vbMaximizedFocus
End Sub

Private Sub Command53_Click()
    On Error GoTo Label1
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "从文件载入里程参数..."
        .DefaultExt = "lch"
        .Filter = "输入的里程参数(*.lch;*.xls;*.xlsx)|*.lch;*.xls;*.xlsx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text100 = .FileTitle
    End With
    Dim fenge As String
    fenge = Chr(38) & Chr(34) & "," & Chr(34) & Chr(38)                         'fenge相当于&","&
    
    If Right(CommonDialog1.FileTitle, 3) = "lch" Then                           '如果扩展名是lch
        hangshu = 0
        Open CommonDialog1.FileName For Input As 16
        Do Until EOF(16)
            Line Input #16, shuoming1
            hangshu = hangshu + 1
        Loop
        Seek #16, 1                                                             '文件指针跳转到文件的第一个字节
        msg = MsgBox("单行输出宽度？是=单行，否=双行，取消=退出。", vbDefaultButton1 + vbYesNoCancel, "单双行输出到excel")
        If msg = vbCancel Then Close #16: Exit Sub                              '如果点了取消就关闭文件并退出子程序
        Line Input #16, shuoming1
        Line Input #16, shuoming2
        Line Input #16, shuoming3
        ndianhao2 = 1
        On Error Resume Next                                                    '如果没有发现excel进程就继续执行下一句：创建excel进程
        Set appexcel = GetObject(, "Excel.application")
        If appexcel Is Nothing Then Set appexcel = CreateObject("excel.application")
        If Err Then                                                             '如果没有安装excel，就用wps表格打开
            Set appexcel = GetObject(, "et.application")
            If appexcel Is Nothing Then Set appexcel = CreateObject("et.application")
        End If
        appexcel.Visible = True
        Set wbmybook = appexcel.Workbooks.Add
        Set exsheet = wbmybook.Worksheets("sheet1")
        exsheet.Activate                                                        '激活
        appexcel.cells.HorizontalAlignment = xlCenter
        appexcel.cells.VerticalAlignment = xlCenter
        appexcel.sheets("sheet1").Name = "宽度统计表"
        appexcel.Range("A2") = "里程"
        appexcel.Range("B2") = "左宽"
        appexcel.Range("C2") = "右宽"
        appexcel.Range("H2") = "左宽2"
        appexcel.Range("I2") = "右宽2"
        appexcel.Range("J2") = "生成zas文件"
appexcel.Columns("A:A").columnwidth = 9
appexcel.Columns("B:C").columnwidth = 7.13
appexcel.Columns("D:D").columnwidth = 9
appexcel.Columns("E:F").columnwidth = 7.13
appexcel.Columns("G:G").columnwidth = 9
appexcel.Columns("H:I").columnwidth = 7.13
appexcel.Columns("J:J").columnwidth = 25.75
appexcel.Range("A1:C1").merge
        appexcel.Range("A1") = Text99.text
        appexcel.cells(1, 1).Font.Bold = True
        appexcel.cells(1, 1).Font.Size = 16
        
        Dim jj As Double
        jj = 3
        If msg = vbNo Then                                                      '合并单元格
appexcel.Range("A3:A4").merge
appexcel.Range("A3:A4").AutoFill Destination:=appexcel.Range("A3:C4"), Type:=0
appexcel.Range("A3:C4").AutoFill Destination:=appexcel.Range("A3:C" & hangshu * 2 - 4), Type:=0
appexcel.Range("H3:H4").merge
appexcel.Range("H3:H4").AutoFill Destination:=appexcel.Range("H3:J4"), Type:=0
appexcel.Rows("3:" & hangshu * 2 - 4).RowHeight = 11.25
        End If
        
        Do While EOF(16) = False
            Input #16, KG
            Call 计算宽度(KG, Lx, Rx)
            On Error GoTo Label1
            appexcel.cells(jj, 1) = KG
            appexcel.cells(jj, 2) = Lx
            appexcel.cells(jj, 3) = Rx
            If msg = vbYes Then                                                 '单行输出
                jj = jj + 1
            Else                                                                '双行输出
                jj = jj + 2
            End If
            ndianhao2 = ndianhao2 + 1
        Loop
        
        '求宽度并生成zas文件
        
        appexcel.cells(3, 8) = "=B" & 3
        appexcel.cells(3, 9) = "=C" & 3
        If Val(Text43) <> 0 Then
            appexcel.cells(3, 10) = "=A" & 3 & fenge & "-H" & 3 & fenge & -Val(Text43) & fenge & Val(Text43) & fenge & "I" & 3
        Else
            appexcel.cells(3, 10) = "=A" & 3 & fenge & "-H" & 3 & fenge & 0 & fenge & "I" & 3
        End If
        If msg = vbYes Then
appexcel.Range("H3:J3").AutoFill Destination:=appexcel.Range("H3:J" & jj - 1), Type:=0
        ElseIf msg = vbNo Then
appexcel.Range("H3:J4").AutoFill Destination:=appexcel.Range("H3:J" & hangshu * 2 - 4), Type:=0
        End If
        
        Close #16
        
        
        '设置里程格式
        appexcel.Range(appexcel.cells(3, 1), appexcel.cells(jj - 1, 1)).NumberFormatLocal = "!" & Text111.text
        '设置页面边框，居中
        appexcel.Range(appexcel.cells(2, 1), appexcel.cells(jj - 1, 8)).Select
appexcel.Range("A2:C" + LTrim(Str(jj - 1))).Select
        appexcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        appexcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With appexcel.Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With appexcel.Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With appexcel.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With appexcel.Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With appexcel.Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With appexcel.Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        '设置第二个标高面的边框
appexcel.Range("H2:J" + LTrim(Str(jj - 1))).Select
        appexcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        appexcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With appexcel.Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With appexcel.Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With appexcel.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With appexcel.Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With appexcel.Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With appexcel.Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        appexcel.cells(2, 12).Select
        
        '设置顶端标题行和页脚的第几页共几页
        'With appexcel.ActiveSheet.PageSetup
        '        .PrintTitleRows = "$1:$2"
        '        .PrintTitleColumns = ""
        '    End With
        '    appexcel.ActiveSheet.PageSetup.PrintArea = ""
        '    With appexcel.ActiveSheet.PageSetup
        '        .LeftHeader = ""
        '        .CenterHeader = ""
        '        .RightHeader = ""
        '        .LeftFooter = ""
        '        .CenterFooter = "第 &P 页，共 &N 页"
        '        .RightFooter = ""
        '        .LeftMargin = appexcel.Application.InchesToPoints(0.748031496062992)
        '        .RightMargin = appexcel.Application.InchesToPoints(0.748031496062992)
        '        .TopMargin = appexcel.Application.InchesToPoints(0.866141732283465)
        '        .BottomMargin = appexcel.Application.InchesToPoints(0.866141732283465)
        '        .HeaderMargin = appexcel.Application.InchesToPoints(0.511811023622047)
        '        .FooterMargin = appexcel.Application.InchesToPoints(0.511811023622047)
        '        .PrintHeadings = False
        '        .PrintGridlines = False
        '        .PrintComments = xlPrintNoComments
        '        .CenterHorizontally = True
        '        .CenterVertically = False
        '        .Orientation = xlPortrait
        '        .Draft = False
        '        .PaperSize = xlPaperA4
        '        .FirstPageNumber = xlAutomatic
        '        .Order = xlDownThenOver
        '        .BlackAndWhite = False
        '        .Zoom = 100
        '        .PrintErrors = xlPrintErrorsDisplayed
        '    End With
        
    Else                                                                        '如果扩展名是xls或者xlsx
        Dim xlapp As Object
        Dim xlBook As Object
        Dim xlSheet As Object
        On Error Resume Next
        '        Set xlapp = GetObject(CommonDialog1.FileName).application
        '        Set xlBook = xlapp.Workbooks.Open(CommonDialog1.FileName)
        '        Set xlSheet = xlBook.activesheet
        Set xlapp = GetObject(, "Excel.application")
        If xlapp Is Nothing Then Set xlapp = CreateObject("excel.application")
        If Err Then                                                             '如果没有安装excel，就用wps表格打开
            Set xlapp = GetObject(, "et.application")
            If xlapp Is Nothing Then Set xlapp = CreateObject("et.application")
        End If
        Set xlBook = xlapp.Workbooks(CommonDialog1.FileTitle)                   '假设xlbook为选定的xls工作簿
        If xlBook Is Nothing Then Set xlBook = xlapp.Workbooks.Open(CommonDialog1.FileName)
        '如果xlbook不是选定的工作簿就打开那个工作簿
        Set xlSheet = xlBook.activesheet
        置顶 Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)    '调用模块里的api把excel进程置顶
        xlapp.Visible = True
        jj = 3
        ndianhao = 1
        
Label2:
        KG = xlSheet.cells(jj, 1)
        Call 计算宽度(KG, Lx, Rx)
        On Error GoTo Label1
        xlSheet.cells(jj, 2) = Lx
        xlSheet.cells(jj, 3) = Rx
        jj = jj + xlSheet.cells(jj, 1).MergeArea.Rows.Count                     '单元格的合并区域的行计数
        ndianhao = ndianhao + 1
        If xlSheet.cells(jj, 1) <> "" Then GoTo Label2
        
        '求宽度并生成zas文件
        xlSheet.cells(3, 8) = "=B" & 3
        xlSheet.cells(3, 9) = "=C" & 3
        If Val(Text43) <> 0 Then
            xlSheet.cells(3, 10) = "=A" & 3 & fenge & "-H" & 3 & fenge & -Val(Text43) & fenge & Val(Text43) & fenge & "I" & 3
        Else
            xlSheet.cells(3, 10) = "=A" & 3 & fenge & "-H" & 3 & fenge & 0 & fenge & "I" & 3
        End If
        If xlSheet.cells(3, 1).MergeArea.Rows.Count = 1 Then
xlSheet.Range("H3:J3").AutoFill Destination:=xlSheet.Range("H3:J" & jj - 1), Type:=0
        Else
xlSheet.Range("H3:J4").AutoFill Destination:=xlSheet.Range("H3:J" & hangshu * 2 - 4), Type:=0
        End If
        
        '        '设置里程格式
        '        xlSheet.Range(xlSheet.cells(3, 1), xlSheet.cells(jj - 1, 1)).NumberFormatLocal = "!" & Text111.text
        '        '设置页面边框，居中
        '        xlSheet.Range(xlSheet.cells(2, 1), xlSheet.cells(jj - 1, 8)).Select
'xlSheet.Range("A2:C" + LTrim(Str(jj - 1))).Select
        '        xlSheet.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        '        xlSheet.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        '        With xlSheet.Selection.Borders(xlEdgeLeft)
        '            .LineStyle = xlContinuous
        '            .Weight = xlMedium
        '            .ColorIndex = xlAutomatic
        '        End With
        '        With xlSheet.Selection.Borders(xlEdgeTop)
        '            .LineStyle = xlContinuous
        '            .Weight = xlMedium
        '            .ColorIndex = xlAutomatic
        '        End With
        '        With xlSheet.Selection.Borders(xlEdgeBottom)
        '            .LineStyle = xlContinuous
        '            .Weight = xlMedium
        '            .ColorIndex = xlAutomatic
        '        End With
        '        With xlSheet.Selection.Borders(xlEdgeRight)
        '            .LineStyle = xlContinuous
        '            .Weight = xlMedium
        '            .ColorIndex = xlAutomatic
        '        End With
        '        With xlSheet.Selection.Borders(xlInsideVertical)
        '            .LineStyle = xlContinuous
        '            .Weight = xlthin
        '            .ColorIndex = xlAutomatic
        '        End With
        '        With xlSheet.Selection.Borders(xlInsideHorizontal)
        '            .LineStyle = xlContinuous
        '            .Weight = xlthin
        '            .ColorIndex = xlAutomatic
        '        End With
        
        '设置第二个标高面的边框
        xlapp.Range("H2") = "左宽2"
        xlapp.Range("I2") = "右宽2"
        xlapp.Range("J2") = "生成zas文件"
xlapp.Range("H2:J" + LTrim(Str(jj - 1))).Select
        xlapp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        xlapp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With xlapp.Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With xlapp.Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With xlapp.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With xlapp.Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With xlapp.Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        With xlapp.Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlthin
            .ColorIndex = xlAutomatic
        End With
        xlapp.cells(2, 12).Select
        Set xlapp = Nothing
    End If
Label1: Close
End Sub

Private Sub Command54_Click()
    KG = Val(Text104)
    Call 计算宽度(KG, Lx, Rx)
    Text110 = Lx
    Text103 = Rx
End Sub

Sub 计算宽度(KG, Lx, Rx)
    kuandu(0, 0) = kuandu(1, 0)
    kuandu(0, 1) = kuandu(1, 1)
    kuandu(0, 2) = kuandu(1, 2)
    求区间3 KG, number3                                                         '计算出KG落在宽度分段的哪个区间上，保存在变量number3中
    K1 = CDec(kuandu(number3 - 1, 0))                                           '转换为decimal类型
    K2 = CDec(kuandu(number3, 0))
    L1 = CDec(kuandu(number3 - 1, 1))
    L2 = CDec(kuandu(number3, 1))
    If Val(kuandu(number3, 2)) = 0 Then                                         '线性渐变
        If K2 <> K1 Then Lx = L1 + (KG - K1) * (L2 - L1) / (K2 - K1)
        If K2 = K1 Then Lx = L1
    ElseIf Val(kuandu(number3, 2)) = 1 Then                                     '抛物线渐变
        If K2 <> K1 Then Kx = (KG - K1) / (K2 - K1): Lx = L1 + (3 * Kx ^ 2 - 2 * Kx ^ 3) * (L2 - L1)
        If K2 = K1 Then Lx = L1
    End If
    
    kuandu2(0, 0) = kuandu2(1, 0)
    kuandu2(0, 1) = kuandu2(1, 1)
    kuandu2(0, 2) = kuandu2(1, 2)
    求区间4 KG, number4                                                         '计算出KG落在宽度分段的哪个区间上，保存在变量number4中
    K1 = CDec(kuandu2(number4 - 1, 0))                                          '转换为decimal类型
    K2 = CDec(kuandu2(number4, 0))
    R1 = CDec(kuandu2(number4 - 1, 1))
    R2 = CDec(kuandu2(number4, 1))
    If Val(kuandu2(number4, 2)) = 0 Then                                        '线性渐变
        If K2 <> K1 Then Rx = R1 + (KG - K1) * (R2 - R1) / (K2 - K1)
        If K2 = K1 Then Rx = R1
    ElseIf Val(kuandu2(number4, 2)) = 1 Then                                    '抛物线渐变
        If K2 <> K1 Then Kx = (KG - K1) / (K2 - K1): Rx = R1 + (3 * Kx ^ 2 - 2 * Kx ^ 3) * (R2 - R1)
        If K2 = K1 Then Rx = L1
    End If
    
    weishu = Val(Text98)
    Lx = Round(Lx, weishu)
    Rx = Round(Rx, weishu)
End Sub

'----------------------------------------------------------------------------------------------------------
Sub 求区间3(K, number)                                                          '求K所在的宽度区间，保存在number变量里
    number = 1                                                                  '先赋个初值
    If kuandu(1, 0) < kuandu(2, 0) Then
        Do While K > CDec(Val(kuandu(number, 0))) And number < Val(Text97)
            number = number + 1
        Loop                                                                    '跳出循环后，Number为K所在的区间
    End If
    If kuandu(1, 0) > kuandu(2, 0) Then
        Do While K < CDec(Val(kuandu(number, 0))) And number < Val(Text97)
            number = number + 1
        Loop                                                                    '跳出循环后，Number为K所在的区间
    End If
    If Combo7.ListCount > 0 Then Combo7.ListIndex = number - 1
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 求区间4(K, number)                                                          '求K所在的宽度区间，保存在number变量里
    number = 1                                                                  '先赋个初值
    If kuandu2(1, 0) < kuandu2(2, 0) Then
        Do While K > CDec(Val(kuandu2(number, 0))) And number <= Val(Text102)
            number = number + 1
        Loop                                                                    '跳出循环后，Number为K所在的区间
    End If
    If kuandu2(1, 0) > kuandu2(2, 0) Then
        Do While K < CDec(Val(kuandu2(number, 0))) And number <= Val(Text102)
            number = number + 1
        Loop                                                                    '跳出循环后，Number为K所在的区间
    End If
    If Combo8.ListCount > 0 Then Combo8.ListIndex = number - 1
End Sub


Private Sub Command55_Click()
    '生成卡西欧5800计算器主程序
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "生成卡西欧5800可编程计算器程序..."
        .DefaultExt = "txt"
        .Filter = "卡西欧5800可编程计算器程序(*.txt)|*.txt"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
        'Text2 = .FileTitle
    End With
    Open CommonDialog1.FileName For Output As 17
    Print #17, Text99
    Print #17, "计算全线路基宽度主程序，宽度分为线性渐变和三次抛物线渐变两种方式。"
    Print #17, "K所求点里程，A左宽，B右宽。"
    Print #17, "算出宽度后默认把-A保存到T（负值），可手工输入右宽（正值），调用“ZKGC”程序默认计算左幅边桩高程G，输入实测边桩高程C，根据高差、坡度和顶宽计算路基底宽W。"
    Print #17, "把W保存到第一偏距T中，再调用“ZK”程序计算边桩坐标。"
    Print #17, "2→Dimz：↙"                                                    '新增两个扩充变量Z[1]用来存储左幅宽度，Z[2]用来存储右幅宽度
    Print #17, "?K：↙"
    
    kuandu(0, 0) = kuandu(1, 0)
    kuandu(0, 1) = kuandu(1, 1)
    kuandu(0, 2) = kuandu(1, 2)
    For n1 = 1 To Val(Text97)
        K1 = CDec(kuandu(n1 - 1, 0))                                            '转换为decimal类型
        K2 = CDec(kuandu(n1, 0))
        L1 = CDec(kuandu(n1 - 1, 1))
        L2 = CDec(kuandu(n1, 1))
        If Val(kuandu(n1, 2)) = 0 Then                                          '如果=0就是线性渐变
            If L2 <> L1 And K2 <> K1 Then
                Print #17, "If K≤" & kuandu(n1, 0) & "：Then " & _
                L1 & "+(K-" & K1 & ")×(" & (L2 - L1) & ")/(" & (K2 - K1) & ")→A：" & _
                "Goto 3：IfEnd↙"
            Else
                Print #17, "If K≤" & kuandu(n1, 0) & "：Then " & _
                L1 & "→A：" & _
                "Goto 3：IfEnd↙"
            End If
        ElseIf Val(kuandu(n1, 2)) = 1 Then                                      '如果是抛物线渐变
            If L2 <> L1 And K1 <> K2 Then
                Print #17, "If K≤" & kuandu(n1, 0) & "：Then " & _
                "(K-" & K1 & ")/(" & (K2 - K1) & ")→I：" & _
                L1 & "+(3×I^2-2×I^3)×(" & (L2 - L1) & ")→A：" & _
                "Goto 3：IfEnd↙"
            Else
                Print #17, "If K≤" & kuandu(n1, 0) & "：Then " & _
                L1 & "→A：" & _
                "Goto 3：IfEnd↙"
            End If
        End If
    Next
    Print #17, "Lbl 3：↙"
    
    kuandu2(0, 0) = kuandu2(1, 0)
    kuandu2(0, 1) = kuandu2(1, 1)
    kuandu2(0, 2) = kuandu2(1, 2)
    For n1 = 1 To Val(Text102)
        K1 = CDec(kuandu2(n1 - 1, 0))                                           '转换为decimal类型
        K2 = CDec(kuandu2(n1, 0))
        R1 = CDec(kuandu2(n1 - 1, 1))
        R2 = CDec(kuandu2(n1, 1))
        If Val(kuandu2(n1, 2)) = 0 Then                                         '如果=0就是线性渐变
            If R2 <> R1 And K2 <> K1 Then
                Print #17, "If K≤" & kuandu2(n1, 0) & "：Then " & _
                R1 & "+(K-" & K1 & ")×(" & (R2 - R1) & ")/(" & (K2 - K1) & ")→B：" & _
                "Goto 4：IfEnd↙"
            Else
                Print #17, "If K≤" & kuandu2(n1, 0) & "：Then " & _
                R1 & "→B：" & _
                "Goto 4：IfEnd↙"
            End If
        ElseIf Val(kuandu2(n1, 2)) = 1 Then                                     '如果是抛物线渐变
            If R2 <> R1 And K1 <> K2 Then
                Print #17, "If K≤" & kuandu2(n1, 0) & "：Then " & _
                "(K-" & K1 & ")/(" & (K2 - K1) & ")→I：" & _
                R1 & "+(3×I^2-2×I^3)×(" & (R2 - R1) & ")→B：" & _
                "Goto 4：IfEnd↙"
            Else
                Print #17, "If K≤" & kuandu2(n1, 0) & "：Then " & _
                R1 & "→B：" & _
                "Goto 4：IfEnd↙"
            End If
        End If
    Next
    
    Print #17, "Lbl 4：“A=”：A◢“B=”：B◢-A→Z[1]：B→Z[2]：-A→T：Prog“1ZKGC”◢"
    Print #17, "？C：If T≤0：Then Z[1]-（G-C）×1.5→W：Else Z[2]+（G-C）×1.5→W：IfEnd↙"
    Print #17, "“W=”：W◢W→T：Prog“1HXZK”"
    Close #17
    msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
    If MsgBox(msg, vbYesNo) = vbYes Then
        Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
    End If
End Sub

Private Sub Command56_Click()
    For n1 = 1 To Val(TextPartNumber)
        If CDec(factor(n1 - 1, 1)) > CDec(factor(n1, 1)) Then                   '检验终点里程是否有序，如果有错误就修改
            msg = "第" & n1 & "段分段终点里程Ki不符合从小到大的顺序，现在修改吗？"
            queren = MsgBox(msg, vbYesNo + vbDefaultButton1 + vbCritical, "检验终点里程")
            If queren = vbYes Then
                Combo1.ListIndex = n1 - 1
                With TEXTKI
                    .SetFocus                                                   '让text获得焦点
                    .SelStart = 0
                    .SelLength = Len(.text)
                End With
            End If
            Exit For
        End If
        
        If factor(n1, 0) = "缓1+圆+缓2" And factor(n1 + 1, 0) = "缓1+圆+缓2" Then
            If Abs(factor(n1, 1) - factor(n1 + 1, 2)) <= 0.003 And _
                Abs(factor(n1, 6)) > Abs(factor(n1 + 1, 6)) Then                '分段终点错误
                Combo1.ListIndex = n1 - 1
                msg = "第" & n1 & "段分段终点里程Ki错误，现在修改吗？"
                queren = MsgBox(msg, vbYesNo + vbDefaultButton1 + vbCritical, "检验终点里程")
                If queren = vbYes Then
                    With TEXTKI
                        .SetFocus                                               '让text获得焦点
                        .SelStart = 0
                        .SelLength = Len(.text)
                        .text = CDec(factor(n1, 2)) + CDec(factor(n1, 9))
                    End With
                    factor(n1, 1) = CDec(factor(n1, 2)) + CDec(factor(n1, 9))
                End If
                Exit For
            ElseIf Abs(factor(n1, 1) - factor(n1, 2) - factor(n1, 9)) <= 0.003 And _
                Abs(factor(n1, 1) - factor(n1 + 1, 2)) > 0.003 And _
                Abs(factor(n1 + 1, 6)) > Abs(factor(n1, 6)) Then                '分段终点错误
                Combo1.ListIndex = n1 - 1
                msg = "第" & n1 & "段分段终点里程Ki错误，现在修改吗？"
                queren = MsgBox(msg, vbYesNo + vbDefaultButton1 + vbCritical, "检验终点里程")
                If queren = vbYes Then
                    With TEXTKI
                        .SetFocus                                               '让text获得焦点
                        .SelStart = 0
                        .SelLength = Len(.text)
                        .text = CDec(factor(n1 + 1, 2))
                    End With
                    factor(n1, 1) = CDec(factor(n1 + 1, 2))
                End If
                Exit For
            End If
        End If
        
        If factor(n1, 0) = "缓1+圆+缓2" And factor(n1 + 1, 0) = "圆" Then
            If Abs(factor(n1, 6)) > Abs(factor(n1 + 1, 6)) Then                 '分段终点错误
                Combo1.ListIndex = n1 - 1
                msg = "第" & n1 & "段分段终点里程Ki错误，现在修改吗？"
                queren = MsgBox(msg, vbYesNo + vbDefaultButton1 + vbCritical, "检验终点里程")
                If queren = vbYes Then
                    With TEXTKI
                        .SetFocus                                               '让text获得焦点
                        .SelStart = 0
                        .SelLength = Len(.text)
                        .text = CDec(factor(n1, 2)) + CDec(factor(n1, 9))
                    End With
                    factor(n1, 1) = CDec(factor(n1, 2)) + CDec(factor(n1, 9))
                    factor(n1 + 1, 0) = "缓1+圆+缓2"
                    factor(n1 + 1, 7) = factor(n1, 8)                           'A
                    ' factor(n1 + 1, 8) = factor(n1, 8) 'B
                End If
                Exit For
            End If
        End If
        
        If factor(n1, 0) = "圆" And factor(n1 + 1, 0) = "缓1+圆+缓2" Then
            If Abs(factor(n1 + 1, 6)) > Abs(factor(n1, 6)) Then                 '分段终点错误
                Combo1.ListIndex = n1 - 1
                msg = "第" & n1 & "段分段终点里程Ki错误，现在修改吗？"
                queren = MsgBox(msg, vbYesNo + vbDefaultButton1 + vbCritical, "检验终点里程")
                If queren = vbYes Then
                    Combo2.ListIndex = 2                                        '"缓1+圆+缓2"
                    TEXTA.text = CDec(factor(n1 + 1, 7))
                    TEXTB.text = CDec(factor(n1 + 1, 7))
                    TEXTO.text = CDec(factor(n1, 1)) - CDec(factor(n1, 2))
                    factor(n1, 0) = "缓1+圆+缓2"
                    factor(n1, 7) = CDec(factor(n1 + 1, 7))                     'A
                    ' factor(n1, 8) = CDec(factor(n1 + 1, 7)) 'B
                    factor(n1, 9) = CDec(factor(n1, 1)) - CDec(factor(n1, 2))
                    With TEXTKI
                        .SetFocus                                               '让text获得焦点
                        .SelStart = 0
                        .SelLength = Len(.text)
                        .text = CDec(factor(n1 + 1, 2))
                    End With
                    factor(n1, 1) = CDec(factor(n1 + 1, 2))
                End If
                Exit For
            End If
        End If
    Next
    If n1 = Val(TextPartNumber) + 1 Then
        MsgBox "没有发现分段终点里程的错误。", vbInformation + vbOKOnly, "检验终点里程"
    End If
End Sub

Private Sub Command57_Click()
    '二路归并算法,把二维有序数列kuandu,kuandu2第0列归并进入有序数列tezheng
    Dim tezheng(2000) As Variant                                                '存储变坡特征点里程
    ia = 1
    ib = 1
    ic = 1
    iamax = Val(Text97)                                                         '根据实际情况修改为有序数列的个数
    ibmax = Val(Text102)                                                        '根据实际情况修改为有序数列的个数
    Do While ia <= iamax And ib <= ibmax
        If kuandu(ia, 0) < kuandu2(ib, 0) Then                                  '取小值放入c数组
            tezheng(ic) = kuandu(ia, 0)
            ia = ia + 1
            ic = ic + 1
        ElseIf kuandu(ia, 0) > kuandu2(ib, 0) Then                              '取小值放入c数组
            tezheng(ic) = kuandu2(ib, 0)
            ib = ib + 1
            ic = ic + 1
        Else                                                                    '重复的数字只放一次
            tezheng(ic) = kuandu(ia, 0)
            ia = ia + 1
            ib = ib + 1
            ic = ic + 1
        End If
    Loop
    If ia <= iamax Then
        For ix = ia To iamax
            tezheng(ic) = kuandu(ix, 0)
            ic = ic + 1
        Next
    End If
    If ib <= ibmax Then
        For ix = ib To ibmax
            tezheng(ic) = kuandu2(ix, 0)
            ic = ic + 1
        Next
    End If
    
    msg = "输出特征点里程坐标吗？"
    queren = MsgBox(msg, vbYesNoCancel + vbDefaultButton2, "输出特征点确认")
    If queren = vbYes Then
        shuchu = 1
    ElseIf queren = vbNo Then
        shuchu = 0
    Else
        GoTo Label1                                                             '退出子程序
    End If
    
    On Error Resume Next
    Set appexcel = GetObject(, "Excel.application")
    If appexcel Is Nothing Then Set appexcel = CreateObject("excel.application")
    If Err Then                                                                 '如果没有安装excel，就用wps表格打开
        Set appexcel = GetObject(, "et.application")
        If appexcel Is Nothing Then Set appexcel = CreateObject("et.application")
    End If
    appexcel.Visible = True
    Set wbmybook = appexcel.Workbooks.Add
    Set exsheet = wbmybook.Worksheets("sheet1")
    exsheet.Activate                                                            '激活
    
    appexcel.cells.HorizontalAlignment = xlCenter
    appexcel.cells.VerticalAlignment = xlCenter
    appexcel.sheets("sheet1").Name = "宽度统计表"
    appexcel.Range("A2") = "里程"
    appexcel.Range("B2") = "左宽"
    appexcel.Range("C2") = "右宽"
    appexcel.Range("H2") = "左宽2"
    appexcel.Range("I2") = "右宽2"
    appexcel.Range("J2") = "生成zas文件"
appexcel.Columns("A:A").columnwidth = 9
appexcel.Columns("B:C").columnwidth = 7.13
appexcel.Columns("D:D").columnwidth = 9
appexcel.Columns("E:F").columnwidth = 7.13
appexcel.Columns("G:G").columnwidth = 9
appexcel.Columns("H:I").columnwidth = 7.13
appexcel.Columns("J:J").columnwidth = 25.75
appexcel.Range("A1:C1").merge
    appexcel.Range("A1") = Text99.text
    appexcel.cells(1, 1).Font.Bold = True
    appexcel.cells(1, 1).Font.Size = 16
    Dim fenge As String
    fenge = Chr(38) & Chr(34) & "," & Chr(34) & Chr(38)                         '相当于&","&
    
    Dim jj As Double
    jj = 3
    n1 = 1
    For K = CDec(Val(Text104)) To CDec(Val(Text105)) Step CDec(Val(Text106))
        On Error GoTo Label1
        If shuchu = 0 Then
            Call 计算宽度(K, Lx, Rx)
            appexcel.cells(jj, 1) = K
            appexcel.cells(jj, 2) = Lx
            appexcel.cells(jj, 3) = Rx
            jj = jj + 1
        ElseIf shuchu = 1 Then
            If K < CDec(Val(tezheng(n1))) Then                                  '如果K小，直接用K计算
                Call 计算宽度(K, Lx, Rx)
                appexcel.cells(jj, 1) = K
                appexcel.cells(jj, 2) = Lx
                appexcel.cells(jj, 3) = Rx
                jj = jj + 1
            Else                                                                '如果K大，那么就搜索出小于等于K的所有特征点并计算
                Do While CDec(Val(tezheng(n1))) <= K And n1 <= ic - 1
                    Call 计算宽度(CDec(Val(tezheng(n1))), Lx, Rx)
                    appexcel.cells(jj, 1) = Val(tezheng(n1))                    '不能直接用KT，否则会在excel中显示货币型
                    appexcel.cells(jj, 2) = Lx
                    appexcel.cells(jj, 3) = Rx
                    appexcel.Range(appexcel.cells(jj, 1), appexcel.cells(jj, 3)).Font.Bold = True
                    n1 = n1 + 1
                    jj = jj + 1
                Loop
                If K <> CDec(Val(tezheng(n1 - 1))) Then                         '输出非特征点，不能与已输出的特征点重复
                    Call 计算宽度(K, Lx, Rx)
                    appexcel.cells(jj, 1) = K
                    appexcel.cells(jj, 2) = Lx
                    appexcel.cells(jj, 3) = Rx
                    jj = jj + 1
                End If
            End If
        End If
    Next K
    
    '求宽度并生成zas文件
    appexcel.cells(3, 8) = "=B" & 3
    appexcel.cells(3, 9) = "=C" & 3
    If Val(Text43) <> 0 Then
        appexcel.cells(3, 10) = "=A" & 3 & fenge & "-H" & 3 & fenge & -Val(Text43) & fenge & Val(Text43) & fenge & "I" & 3
    Else
        appexcel.cells(3, 10) = "=A" & 3 & fenge & "-H" & 3 & fenge & 0 & fenge & "I" & 3
    End If
appexcel.Range("H3:J3").AutoFill Destination:=appexcel.Range("H3:J" & jj - 1), Type:=0
    
    '设置里程格式
    appexcel.Range(appexcel.cells(3, 1), appexcel.cells(jj - 1, 1)).NumberFormatLocal = "!" & Text111.text
    'appexcel.Range(appexcel.cells(2, 1), appexcel.cells(jj - 1, 8)).Select
appexcel.Range("A2:C" + LTrim(Str(jj - 1))).Select
    appexcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    appexcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With appexcel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    '设置第二个标高面的边框
appexcel.Range("H2:J" + LTrim(Str(jj - 1))).Select
    appexcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    appexcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With appexcel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    appexcel.cells(2, 12).Select
    
    '设置顶端标题行和页脚的第几页共几页
    'With appexcel.ActiveSheet.PageSetup
    '        .PrintTitleRows = "$1:$2"
    '        .PrintTitleColumns = ""
    '    End With
    '    appexcel.ActiveSheet.PageSetup.PrintArea = ""
    '    With appexcel.ActiveSheet.PageSetup
    '        .LeftHeader = ""
    '        .CenterHeader = ""
    '        .RightHeader = ""
    '        .LeftFooter = ""
    '        .CenterFooter = "第 &P 页，共 &N 页"
    '        .RightFooter = ""
    '        .LeftMargin = appexcel.Application.InchesToPoints(0.748031496062992)
    '        .RightMargin = appexcel.Application.InchesToPoints(0.748031496062992)
    '        .TopMargin = appexcel.Application.InchesToPoints(0.866141732283465)
    '        .BottomMargin = appexcel.Application.InchesToPoints(0.866141732283465)
    '        .HeaderMargin = appexcel.Application.InchesToPoints(0.511811023622047)
    '        .FooterMargin = appexcel.Application.InchesToPoints(0.511811023622047)
    '        .PrintHeadings = False
    '        .PrintGridlines = False
    '        .PrintComments = xlPrintNoComments
    '        .CenterHorizontally = True
    '        .CenterVertically = False
    '        .Orientation = xlPortrait
    '        .Draft = False
    '        .PaperSize = xlPaperA4
    '        .FirstPageNumber = xlAutomatic
    '        .Order = xlDownThenOver
    '        .BlackAndWhite = False
    '        .Zoom = 100
    '        .PrintErrors = xlPrintErrorsDisplayed
    '    End With
    
    '代码运行后,恢复Excel原来的状态;将下面的代码放在代码的末尾
    'appexcel.ScreenUpdating = screenUpdateState
    'appexcel.DisplayStatusBar = statusBarState
    'appexcel.Calculation = calcState
    'appexcel.EnableEvents = eventsState
    appexcel.activesheet.DisplayPageBreaks = displayPageBreaksState             '注:这是工作表级的设置
    
Label1:
End Sub

Private Sub Command58_Click()
    保存宽度参数2
    If Combo8.ListIndex < Combo8.ListCount - 1 Then Combo8.ListIndex = Combo8.ListIndex + 1
End Sub
Sub 保存宽度参数2()
    kuandu2(Val(Combo8), 0) = CDec(Val(Text101))                                'kuandu2(段数,0列)存储分段终点里程KI
    kuandu2(Val(Combo8), 1) = CDec(Val(Text94))                                 'kuandu2(段数,1列)存储变坡点右幅横坡
    If Option6.Value = True Then kuandu2(Val(Combo8), 2) = 0                    'kuandu2(段数,2列)=0线形，=1抛物线
    If Option5.Value = True Then kuandu2(Val(Combo8), 2) = 1                    'kuandu2(段数,2列)=0线形，=1抛物线
End Sub

Private Sub Text102_Change()
    Dim I1                                                                      '循环用的变量
    Combo8.Clear
    For I1 = 1 To (Val(Text102))
        Combo8.AddItem I1
    Next I1
    If Combo8.ListCount > 0 Then Combo8.ListIndex = 0
End Sub

Private Sub Text40_Change()
    Text46 = Val(Text40) - Val(Text47)
End Sub

Private Sub Text47_Change()
    Text46 = Val(Text40) - Val(Text47)
End Sub

Private Sub Text48_Change()
    角度运算
End Sub

Private Sub Text49_Change()
    角度运算
End Sub


Private Sub Text53_Change()
    角度转换
End Sub

Private Sub Text55_Change()
    角度转换
End Sub

Private Sub Text57_Change()
    直线2
End Sub

Private Sub Text58_Change()
    直线2
End Sub

Private Sub Text59_Change()
    直线2
End Sub

Private Sub Text60_Change()
    直线2
End Sub

Private Sub Text64_Change()
    载入序号
End Sub

Private Sub Text65_Change()
    载入序号
End Sub

Private Sub Combo6_Click()
    载入交点参数
    隐藏多余交点参数
End Sub

Private Sub Command34_Click()
    保存交点参数
    With Combo6
        If .ListIndex < .ListCount - 1 Then .ListIndex = .ListIndex + 1
    End With
End Sub

Sub 载入交点参数()
    Text73 = CDec(Val(factor2(Val(Combo6), 0)))                                 '第0列存储交点里程
    Text72 = CDec(Val(factor2(Val(Combo6), 1)))                                 '第1列存储交点X
    Text71 = CDec(Val(factor2(Val(Combo6), 2)))                                 '第2列存储交Y
    Text66 = CDec(Val(factor2(Val(Combo6), 3)))                                 '第3列存储圆曲线半径,右负左正
    Text67 = CDec(Val(factor2(Val(Combo6), 4)))                                 '第4列存储缓和曲线长
    Text70 = CDec(Val(factor2(Val(Combo6), 5)))                                 '第5列存储切线长
    Text68 = CDec(Val(factor2(Val(Combo6), 6)))                                 '第6列存储曲线长度
    Check1.Value = Val(factor2(Val(Combo6), 7))                                 '第7列存储是否完全
    Text109 = factor2(Val(Combo6), 8)                                           '第8列存储A1
    Text108 = factor2(Val(Combo6), 9)                                           '第9列存储A2
End Sub

Sub 保存交点参数()
    factor2(Val(Combo6), 0) = CDec(Val(Text73))                                 '第0列存储交点里程
    factor2(Val(Combo6), 1) = CDec(Val(Text72))                                 '第1列存储交点X
    factor2(Val(Combo6), 2) = CDec(Val(Text71))                                 '第2列存储交Y
    factor2(Val(Combo6), 3) = CDec(Val(Text66))                                 '第3列存储圆曲线半径,右负左正
    factor2(Val(Combo6), 4) = CDec(Val(Text67))                                 '第4列存储缓和曲线长
    factor2(Val(Combo6), 5) = CDec(Val(Text70))                                 '第5列存储切线长
    factor2(Val(Combo6), 6) = CDec(Val(Text68))                                 '第6列存储曲线长度
    factor2(Val(Combo6), 7) = Check1.Value                                      '第7列存储是否完全
    factor2(Val(Combo6), 8) = CDec(Val(Text109))                                '第8列存储A1
    factor2(Val(Combo6), 9) = CDec(Val(Text108))                                '第9列存储A2
End Sub

Sub 隐藏多余交点参数()
    If Combo6 = Val(Text64) Then
        LabelT(0).Visible = False
        Label6.Visible = False
        Label7(0).Visible = False
        Label8.Visible = False
        Text70.Visible = False
        Text66.Visible = False
        Text67.Visible = False
        Text68.Visible = False
        Label8.Visible = False
        Text68.Visible = False
        'Check1.Visible = False
    Else
        LabelT(0).Visible = True
        Label6.Visible = True
        Label7(0).Visible = True
        Label8.Visible = True
        Text70.Visible = True
        Text66.Visible = True
        Text67.Visible = True
        Text68.Visible = True
        Label8.Visible = True
        Text68.Visible = True
        'Check1.Visible = True
    End If
    
    If banben = 0 Then                                                          '新版本
        LabelT(0) = "缓2长Ls2"
        Label8.Visible = False                                                  '曲线总长
        Text68.Visible = False                                                  '曲线总长
    ElseIf banben = 1 And Combo6 <> Text64 Then                                 '老版本
        LabelT(0) = "切线长T1"
        Label8.Visible = True                                                   '曲线总长
        Text68.Visible = True                                                   '曲线总长
    End If
    
    If banben = 0 And Combo6 = Text64 Then
        Label1(103) = "起点里程"
    ElseIf banben = 0 And Combo6 = Val(Text64) + Val(Text65) - 1 Then
        Label1(103) = "终点里程"
        LabelT(0).Visible = False
        Label6.Visible = False
        Label7(0).Visible = False
        Label8.Visible = False
        Text70.Visible = False
        Text66.Visible = False
        Text67.Visible = False
        Text68.Visible = False
    Else
        Label1(103) = "交点里程"
    End If
End Sub

Private Sub Text97_Change()
    Dim I1                                                                      '循环用的变量
    Combo7.Clear
    For I1 = 1 To (Val(Text97))
        Combo7.AddItem I1
    Next I1
    If Combo7.ListCount > 0 Then Combo7.ListIndex = 0
End Sub


'----------------------------------------------------------------------------------------------------------
Private Sub TextPartNumber_Change()
    Dim I1 As Double                                                            '循环用的变量
    Combo1.Clear
    For I1 = 1 To (Val(TextPartNumber))
        Combo1.AddItem I1
    Next I1
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Combo1_Click()
    载入平曲线参数
    隐藏多余参数
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 载入平曲线参数()
    Combo2 = factor(Val(Combo1), 0)                                             'factor(段数,0列)存储曲线类型
    TEXTKI = CDec(Val(factor(Val(Combo1), 1)))
    TEXTJ = CDec(Val(factor(Val(Combo1), 2)))
    TEXTC = CDec(Val(factor(Val(Combo1), 3)))
    TEXTD = CDec(Val(factor(Val(Combo1), 4)))
    TEXTF = CDec(Val(factor(Val(Combo1), 5)))
    TEXTR = CDec(Val(factor(Val(Combo1), 6)))
    TEXTA = CDec(Val(factor(Val(Combo1), 7)))
    TEXTB = CDec(Val(factor(Val(Combo1), 8)))
    TEXTO = CDec(Val(factor(Val(Combo1), 9)))
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 隐藏多余参数()
    If Combo2 = "直线" Then
        TEXTR.Visible = False
        TEXTA.Visible = False
        TEXTB.Visible = False
        TEXTO.Visible = False
        LabelR.Visible = False
        LabelA.Visible = False
        LabelB.Visible = False
        LabelO.Visible = False
        Label1(5) = "起算点里程J"
        Label1(6) = "起算点X坐标C"
        Label1(7) = "起算点Y坐标D"
        Label1(8) = "起算点切线方位角F"
    ElseIf Combo2 = "圆" Then
        TEXTR.Visible = True
        TEXTA.Visible = False
        TEXTB.Visible = False
        TEXTO.Visible = False
        LabelR.Visible = True
        LabelA.Visible = False
        LabelB.Visible = False
        LabelO.Visible = False
        Label1(5) = "起算点里程J"
        Label1(6) = "起算点X坐标C"
        Label1(7) = "起算点Y坐标D"
        Label1(8) = "起算点切线方位角F"
    ElseIf Combo2 = "缓1+圆+缓2" Then
        TEXTR.Visible = True
        TEXTA.Visible = True
        TEXTB.Visible = True
        TEXTO.Visible = True
        LabelR.Visible = True
        LabelA.Visible = True
        LabelB.Visible = True
        LabelO.Visible = True
        Label1(5) = "HY点里程J"
        Label1(6) = "HY点X坐标C"
        Label1(7) = "HY点Y坐标D"
        Label1(8) = "HY点切线方位角F"
        LabelA = "第一回旋参数A"
        LabelB = "第二回旋参数B"
        LabelO = "圆曲线长O"
    ElseIf Combo2 = "布劳斯曲线" Then
        LabelB.Visible = False
        LabelO.Visible = False
        TEXTB.Visible = False
        TEXTO.Visible = False
        TEXTR.Visible = True
        TEXTA.Visible = True
        LabelR.Visible = True
        LabelA.Visible = True
        Label1(5) = "ZH/HZ点里程J"
        Label1(6) = "ZH/HZ点X坐标C"
        Label1(7) = "ZH/HZ点Y坐标D"
        Label1(8) = "ZH/HZ点切线方位角F"
        LabelA = "缓和曲线长Ls"
    ElseIf Combo2 = "麦克康纳尔曲线" Then
        LabelR.Visible = True
        LabelA.Visible = True
        LabelB.Visible = True
        LabelO.Visible = True
        TEXTR.Visible = True
        TEXTA.Visible = True
        TEXTB.Visible = True
        TEXTO.Visible = True
        Label1(5) = "ZH/HZ点里程J"
        Label1(6) = "ZH/HZ点X坐标C"
        Label1(7) = "ZH/HZ点Y坐标D"
        Label1(8) = "ZH/HZ点切线方位角F"
        LabelA = "平衡车速km/h"
        LabelB = "ZH/HZ点横坡"
        LabelO = "重力加速度"
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Combo2_Click()
    隐藏多余参数
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command9_Click()                                                    '输入的事件
    保存平曲线参数
    If Combo1.ListIndex < Combo1.ListCount - 1 Then Combo1.ListIndex = Combo1.ListIndex + 1
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 保存平曲线参数()
    factor(Val(Combo1), 0) = Combo2                                             'factor(段数,0列)存储曲线类型
    factor(Val(Combo1), 1) = CDec(Val(TEXTKI))                                  'factor(段数,1列)存储分段终点里程KI
    factor(Val(Combo1), 2) = CDec(Val(TEXTJ))                                   'factor(段数,2列)存储起算点里程J
    factor(Val(Combo1), 3) = CDec(Val(TEXTC))                                   'factor(段数,3列)存储起算点X坐标C
    factor(Val(Combo1), 4) = CDec(Val(TEXTD))                                   'factor(段数,4列)存储起算点Y坐标D
    factor(Val(Combo1), 5) = CDec(Val(TEXTF))                                   'factor(段数,5列)存储起算点的切线方位角F
    factor(Val(Combo1), 6) = CDec(Val(TEXTR))                                   'factor(段数,6列)存储圆半径R
    factor(Val(Combo1), 7) = CDec(Val(TEXTA))                                   'factor(段数,7列)存储第一回旋参数A
    factor(Val(Combo1), 8) = CDec(Val(TEXTB))                                   'factor(段数,8列)存储第二回旋参数B
    factor(Val(Combo1), 9) = CDec(Val(TEXTO))                                   'factor(段数,9列)存储圆曲线长O
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command1_Click()                                                    '从文件载入平曲线参数...
    Dim n1, n2 As Integer                                                       '循环用的变量
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "从文件载入平曲线参数..."
        .DefaultExt = "ini"
        .Filter = "输入的曲线参数(*.ini)|*.ini"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text2 = .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 1
    Input #1, Partnumber
    TextPartNumber.text = Partnumber
    Input #1, Kqd                                                               '把文件中的分段数、起点里程、分段序号读入到相应的变量中
    TextQD.text = Kqd
    '先清空平曲线参数数组,避免载入第二个平曲线计算受到影响！
    Erase factor
    For n1 = 1 To Partnumber
        Input #1, n1
        For n2 = 0 To 9
            Input #1, factor(n1, n2)
        Next n2
    Next n1
    
    载入平曲线参数
    隐藏多余参数
    TEXTKstart = TextQD                                                         '为起算点、终点里程、间距赋初值，默认计算全线坐标
    TEXTKend = factor(Val(TextPartNumber), 1)
    If EOF(1) = False Then
        Input #1, jianju
        TEXToffset = jianju
    End If
    
    Close #1
    Saveflag = 1
    
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command2_Click()                                                    '保存平曲线参数到文件...
    Dim n1, n2 As Integer                                                       '循环用的变量
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "保存平曲线参数到文件..."
        .DefaultExt = "ini"
        .Filter = "输出的曲线参数(*.ini)|*.ini"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text2 = .FileTitle
    End With
    Open CommonDialog1.FileName For Output As 2
    Partnumber = Val(TextPartNumber)
    Kqd = Val(TextQD)
    Write #2, Val(TextPartNumber)
    Write #2, Val(TextQD)
    For n1 = 1 To Val(TextPartNumber)
        Write #2, n1
        For n2 = 0 To 9
            Print #2, factor(n1, n2)
        Next n2
    Next n1
    Write #2, Val(TEXToffset)
    Close #2
    Saveflag = 1
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command6_Click()                                                    '计算起点按钮的单击事件
    Kstart = Val(TEXTKstart)
    Kqd = Val(TextQD)
    Partnumber = Val(TextPartNumber)
    T = CDec(Val(TEXTT))
    P = CDec(Val(TEXTP))
    S = CDec(Val(TEXTS))
    Z = CDec(Val(TEXTZ))
    'If Kstart >= Kqd And Kstart <= Val(factor(Partnumber, 1)) Then
    计算 Kstart, x, y, Q, T, P, S, Z, G, H
    TEXTX = x
    TEXTY = y
    TEXTQ = Q
    TEXTG = G
    TEXTH = H
    'Else
    'TEXTKstart = Val(InputBox("K必须在" & Kqd & "到" & Val(factor(Partnumber, 1)) & "之间！请重新输入："))
    'Kstart = Val(TEXTKstart)
    'End If
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command11_Click()                                                   '导出成全站仪CSV格式按钮的单击事件
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "添加到全站仪的CSV文件..."
        .DefaultExt = "scp"
        .Filter = "坐标结果(*.scp)|*.scp"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
    End With
    Open CommonDialog1.FileName For Append As 3
    Kstart = CDec(Val(TEXTKstart))
    Kend = CDec(Val(TEXTKend))
    Kqd = CDec(Val(TextQD))
    Partnumber = Val(TextPartNumber)
    T = CDec(Val(TEXTT))
    P = CDec(Val(TEXTP))
    S = CDec(Val(TEXTS))
    Z = CDec(Val(TEXTZ))
    If Val(TEXToffset) <> 0 And Kqd <= Kstart And Kstart <= Kend And Kend <= CDec(Val(factor(Partnumber, 1))) Then
        '执行多点计算
        For K = Kstart To Kend Step CDec(Val(TEXToffset))
            计算 K, x, y, Q, T, P, S, Z, G, H
            TEXTX = x
            TEXTY = y
            TEXTQ = Q
            TEXTG = G
            TEXTH = H
            x = Format(x, ",+00000000.000")
            y = Format(y, ",+00000000.000")
            Print #3, K & x & y & ",+00000000.000,,"
        Next K
        msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
        If MsgBox(msg, vbYesNo) = vbYes Then
            Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbNormalFocus
        End If
    Else                                                                        '如果间距为零或超限就执行单点计算
        计算 Kstart, x, y, Q, T, P, S, Z, G, H
        TEXTX = x
        TEXTY = y
        TEXTQ = Q
        TEXTG = G
        TEXTH = H
        x = Format(x, ",+00000000.000")
        y = Format(y, ",+00000000.000")
        Print #3, Kstart & x & y & ",+00000000.000,,"
        msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
        If MsgBox(msg, vbYesNo) = vbYes Then
            Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbNormalFocus
        End If
    End If
    Close #3
End Sub

Private Sub Command5_Click()
    '    Shell "start " & """" & App.Path & "\高速公路坐标高程计算软件说明.doc" & """", vbMaximizedFocus
    ShellExecute Me.hwnd, "open", App.Path & "\高速公路坐标高程计算软件说明.doc", vbNullString, vbNullString, 1
End Sub

'----------------------------------------------------------------------------------------------------------
'Private Sub Command55_Click()                                                    '生成中桩坐标的AUTOCAD脚本文件...
'    If TEXTKstart <> "" And TEXTKend <> "" And TEXTKI <> "" And Val(TEXToffset) <> 0 Then '满足条件才计算
'        With CommonDialog1
'            .FileName = ""                                                      '先清空文件名
'            .CancelError = False
'            .DialogTitle = "生成中桩坐标的AUTOCAD脚本文件（Z=K）..."
'            .DefaultExt = "scr"
'            .Filter = "AUTOCAD脚本文件(*.scr)|*.scr"
'            .ShowSave
'            If Len(.FileName) = 0 Then: Close: Exit Sub
'        End With
'        Open CommonDialog1.FileName For Output As 5
'        Print #5, "setvar"
'        Print #5, "osnapcoord"
'        Print #5, "1"                                                           '设为1也可以，2=键盘输入优先，脚本例外
'        Print #5, "UCS"                                                         '在世界坐标系下画图
'        Print #5, "W"
'        'Print #5, "UCS"    '在大地测量坐标系下画图
'        'Print #5, "N"
'        'Print #5, "3"
'        'Print #5, "0,0"
'        'Print #5, "0,1"
'        'Print #5, "1,0"
'        Print #5, "line"
'        Kstart = CDec(Val(TEXTKstart))
'        Kend = CDec(Val(TEXTKend))
'        offset = CDec(Val(TEXToffset))
'        T = CDec(Val(TEXTT))
'        P = CDec(Val(TEXTP))
'        S = CDec(Val(TEXTS))
'        Z = CDec(Val(TEXTZ))
'        For K = CDec(Val(TEXTKstart)) To CDec(Val(TEXTKend)) Step CDec(Val(TEXToffset))
'            计算 K, x, y, Q, T, P, S, Z, G, H
'            Write #5, y, x, K                                                   '在世界坐标系下画图XY必须互换，把里程写入Z坐标，便于查看
'        Next K
'        Print #5, ""                                                            '结束line命令
'        '显示所有图形
'        Print #5, "zoom"
'        Print #5, "e"
'        Close #5
'
'        msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
'        If MsgBox(msg, vbYesNo) = vbYes Then
'            Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
'        End If
'    Else                                                                        '如果TEXTKstart为空或TEXTKend为空或TextKi为空或间距为零，就提示重新输入参数
'        MsgBox "请输全相关参数,并确保间距不为零且KQD≤Kstart≤Kend≤Ki(max)!", vbCritical, "警告！"
'    End If
'End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command7_Click()                                                    '导出全部坐标到文件...
    If TEXTKstart <> "" And TEXTKend <> "" And TEXTKI <> "" And Val(TEXToffset) <> 0 Then
        With CommonDialog1
            .FileName = ""                                                      '先清空文件名
            .CancelError = False
            .DialogTitle = "导出全部坐标到文件..."
            .DefaultExt = "txt"
            .Filter = "坐标结果(*.txt)|*.txt"
            .ShowSave
            If Len(.FileName) = 0 Then: Close: Exit Sub
        End With
        T = CDec(Val(TEXTT))
        P = CDec(Val(TEXTP))
        S = CDec(Val(TEXTS))
        Z = CDec(Val(TEXTZ))
        Open CommonDialog1.FileName For Output As 7
        Print #7, "K,X,Y,Q,G,H"                                                 '便于一一对应
        For K = CDec(Val(TEXTKstart)) To CDec(Val(TEXTKend)) Step CDec(Val(TEXToffset))
            计算 K, x, y, Q, T, P, S, Z, G, H
            Write #7, K, x, y, Q, G, H
        Next K
        Close #7
        msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
        If MsgBox(msg, vbYesNo) = vbYes Then
            Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
        End If
    Else                                                                        '如果TEXTKstart为空或TEXTKend为空或TextKi为空或间距为零，就提示重新输入参数
        MsgBox "请输全相关参数,并确保间距不为零且KQD≤Kstart≤Kend≤Ki(max)!", vbCritical, "警告！"
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 计算(K, x, y, Q, T, P, S, Z, G, H)
    求区间 K, number                                                            '计算出K落在平曲线哪个区间上，保存在变量Number中
    If factor(number, 0) = "直线" Then
        直线 K, x, y, Q, T, P, S, Z, G, H
    ElseIf factor(number, 0) = "圆" Then
        圆 K, x, y, Q, T, P, S, Z, G, H
    ElseIf factor(number, 0) = "缓1+圆+缓2" Then
        平曲线 K, x, y, Q, T, P, S, Z, G, H
    ElseIf factor(number, 0) = "布劳斯曲线" Then
        布劳斯曲线 K, x, y, Q, T, P, S, Z, G, H
    ElseIf factor(number, 0) = "麦克康纳尔曲线" Then
        麦克康纳尔曲线 K, x, y, Q, T, P, S, Z, G, H
    End If
    xiaoshu = Val(Text114)
    xiaoshu2 = xiaoshu + 2
    x = Round(x, xiaoshu)                                                       '小数位取舍
    y = Round(y, xiaoshu)
    Q = Round(Q, xiaoshu2)
    G = Round(G, xiaoshu)
    H = Round(H, xiaoshu)
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 求区间(K, number)                                                           '计算出K落在平曲线哪个区间上，保存在变量Number中
    number = 1                                                                  '先赋个初值
    If factor(1, 1) < factor(2, 1) Then
        Do While Val(K) > Val(factor(number, 1)) And number < Val(TextPartNumber)
            number = number + 1
        Loop
    End If
    If factor(1, 1) > factor(2, 1) Then
        Do While Val(K) < Val(factor(number, 1)) And number < Val(TextPartNumber)
            number = number + 1
        Loop
    End If
    If Combo1.ListCount > 0 Then Combo1.ListIndex = number - 1
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 直线(K, x, y, Q, T, P, S, Z, G, H)
    求区间 K, number                                                            '计算出K落在平曲线的哪个区间上，保存在变量Number中
    J = CDec(Val(factor(number, 2)))
    C = CDec(Val(factor(number, 3)))
    D = CDec(Val(factor(number, 4)))
    F = CDec(Val(factor(number, 5)))
    度分秒化为弧度 F, F1                                                        '调用子程序:F----度.分秒--弧度
    i = K - J
    x = C + i * Cos(F1)                                                         'F1--------弧度
    y = D + i * Sin(F1)
    Q = F1                                                                      'Q-----弧度
    CEYI x, y, Q, T, P, S, Z, G, H
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 圆(K, x, y, Q, T, P, S, Z, G, H)
    求区间 K, number                                                            '计算出K落在平曲线的哪个区间上，保存在变量Number中
    J = CDec(Val(factor(number, 2)))
    C = CDec(Val(factor(number, 3)))
    D = CDec(Val(factor(number, 4)))
    F = CDec(Val(factor(number, 5)))
    R = CDec(Val(factor(number, 6)))
    If R <> 0 Then
        L = K - J
        度分秒化为弧度 F, F1                                                    'F----度.分秒--弧度
        W = F1 + L / (2 * R)                                                    'W----弧度
        i = 2 * R * Sin(L / (2 * R))                                            'I为弦长
        x = C + i * Cos(W)
        y = D + i * Sin(W)
        Q = F1 + L / R                                                          'Q----弧度
        CEYI x, y, Q, T, P, S, Z, G, H
    Else                                                                        'R=0就提示重新输入参数R
        MsgBox ("第" & number & "分段的R不能为零,请修改参数并按输入按钮。"), vbOKOnly, "提示"
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 平曲线(K, x, y, Q, T, P, S, Z, G, H)
    求区间 K, number                                                            '计算出K落在平曲线的哪个区间上，保存在变量Number中
    J = CDec(Val(factor(number, 2)))
    C = CDec(Val(factor(number, 3)))
    D = CDec(Val(factor(number, 4)))
    F = CDec(Val(factor(number, 5)))
    R = CDec(Val(factor(number, 6)))
    A = CDec(Val(factor(number, 7)))
    B = CDec(Val(factor(number, 8)))
    O = CDec(Val(factor(number, 9)))
    If CDec(Val(K)) < CDec(Val(J)) Then
        缓 K, x, y, Q, T, P, S, Z, G, H
    ElseIf CDec(Val(K)) <= CDec(Val(J)) + CDec(Val(O)) Then
        圆 K, x, y, Q, T, P, S, Z, G, H
    Else
        缓 K, x, y, Q, T, P, S, Z, G, H
    End If
End Sub
Sub 布劳斯曲线(K, x, y, Q, T, P, S, Z, G, H)
    求区间 K, number                                                            '计算出K落在平曲线的哪个区间上，保存在变量Number中
    J = CDec(Val(factor(number, 2)))
    C = CDec(Val(factor(number, 3)))
    D = CDec(Val(factor(number, 4)))
    F = CDec(Val(factor(number, 5)))
    R = CDec(Val(factor(number, 6)))
    A = CDec(Val(factor(number, 7)))
    If R <> 0 Then i = R / Abs(R)
    W = 1
    If K < J Then W = -1
    L = Abs(K - J)
    度分秒化为弧度 F, E                                                         'F----度.分秒--弧度
    If A <> 0 And R <> 0 Then
        '幂级数法
        'U = L - L ^ 7 / (14 * R ^ 2 * a ^ 4) + L ^ 8 / (16 * R ^ 2 * a ^ 5) - L ^ 9 / (72 * R ^ 2 * a ^ 6) + L ^ 13 / (312 * R ^ 4 * a ^ 8) - L ^ 14 / (168 * R ^ 4 * a ^ 9) + L ^ 15 / (240 * R ^ 4 * a ^ 10) - L ^ 16 / (768 * R ^ 4 * a ^ 11) + L ^ 17 / (6528 * R ^ 4 * a ^ 12)
        'V = L ^ 4 / (4 * R * a ^ 2) - L ^ 5 / (10 * R * a ^ 3) - L ^ 10 / (60 * R ^ 3 * a ^ 6) + L ^ 11 / (44 * R ^ 3 * a ^ 7) - L ^ 12 / (96 * R ^ 3 * a ^ 8) + L ^ 13 / (624 * R ^ 3 * a ^ 9) + L ^ 16 / (1920 * R ^ 5 * a ^ 10) - L ^ 17 / (816 * R ^ 5 * a ^ 11) + L ^ 18 / (864 * R ^ 5 * a ^ 12) - L ^ 19 / (1824 * R ^ 5 * a ^ 13) + L ^ 20 / (7680 * R ^ 5 * a ^ 14) - L ^ 21 / (80640 * R ^ 5 * a ^ 15)
        '辛普森法求定积分近似值
        L1 = 0
        L2 = Abs(K - J)
        Ls = CDec(Val(factor(number, 7)))
        Call simpson(L1, L2, R, Ls, U, V)
        x = C + W * U * Cos(E) - i * V * Sin(E)
        y = D + W * U * Sin(E) + i * V * Cos(E)
        Q = E + (L ^ 3 / (R * Ls ^ 2) - L ^ 4 / (2 * R * Ls ^ 3)) * W * i       '单位弧度
        CEYI x, y, Q, T, P, S, Z, G, H
    End If
End Sub

Sub simpson(L1, L2, R, Ls, U, V)
    Step = 16                                                                   'step必须是正偶数2,4,6,8,10...
    Do
        U2 = U
        V2 = V
        HL = (L2 - L1) / Step
        L = L1
        Four = 0
        Two = 0
        Four2 = 0
        Two2 = 0
        For id = 1 To Step / 2
            L = L + HL
            Four = Four + Cos(Fun(R, Ls, L))
            Four2 = Four2 + Sin(Fun(R, Ls, L))
            L = L + HL
            Two = Two + Cos(Fun(R, Ls, L))
            Two2 = Two2 + Sin(Fun(R, Ls, L))
        Next id
        Two = Two - Cos(Fun(R, Ls, L))
        Two2 = Two2 - Sin(Fun(R, Ls, L))
        U = (Cos(Fun(R, Ls, L1)) + Cos(Fun(R, Ls, L2)) + 4 * Four + 2 * Two) * HL / 3
        V = (Sin(Fun(R, Ls, L1)) + Sin(Fun(R, Ls, L2)) + 4 * Four2 + 2 * Two2) * HL / 3
        'Print Step, U, V
        Step = Step + 2
    Loop While Abs(U2 - U) > 0.0001 Or Abs(V2 - V) > 0.0001
End Sub

Private Function Fun(R, Ls, L)
    '布劳斯曲线切线方位角与弧长的函数关系τ=f(l)
    If R * Ls <> 0 Then Fun = (L ^ 3 / (R * Ls * Ls) - L ^ 4 / (2 * R * Ls ^ 3))
End Function

'----------------------------------------------------------------------------------------------------------
Sub 缓(K, x, y, Q, T, P, S, Z, G, H)
    Dim W, i As Integer                                                         'W=1为缓1，W=-1缓2；I=1右偏，I=2左偏
    求区间 K, number                                                            '计算出K落在平曲线的哪个区间上，保存在变量Number中
    J = CDec(Val(factor(number, 2)))
    C = CDec(Val(factor(number, 3)))
    D = CDec(Val(factor(number, 4)))
    F = CDec(Val(factor(number, 5)))
    R = CDec(Val(factor(number, 6)))
    A = CDec(Val(factor(number, 7)))
    B = CDec(Val(factor(number, 8)))
    O = CDec(Val(factor(number, 9)))
    W = 1                                                                       '先假设为缓1
    If CDec(Val(K)) >= CDec(Val(J)) + CDec(Val(O)) Then W = -1
    G = A * (1 + W) / 2 + B * (1 - W) / 2                                       'W=1取缓1的回旋参数A，W=-1取缓2的回旋参数B
    If G <> 0 And R <> 0 Then
        度分秒化为弧度 F, F1                                                    'F----度.分秒--弧度
        i = R / Abs(R)                                                          '判断左右偏的变量，右偏R取正，左偏取负
        H = G ^ 2 / Abs(R)                                                      '完整的缓和曲线总长
        L = H - Abs(J + O * (1 - W) / 2 - K)                                    'L=K到ZH（HZ）点的弧长
        E = F1 + O * (1 - W) / (2 * R) - G ^ 2 * W * i / (2 * R ^ 2)            '弧度
        '        U = (L - H) - (L ^ 5 - H ^ 5) / (40 * G ^ 4) + (L ^ 9 - H ^ 9) / (3456 * G ^ 8) _
        '        - (L ^ 13 - H ^ 13) / (599040 * G ^ 12) + (L ^ 17 - H ^ 17) / (175472640 * G ^ 16)
        '
        '        V = (L ^ 3 - H ^ 3) / (6 * G ^ 2) - (L ^ 7 - H ^ 7) / (336 * G ^ 6) + (L ^ 11 - H ^ 11) / (42240 * G ^ 10) _
        '        - (L ^ 15 - H ^ 15) / (9676800 * G ^ 14) + (L ^ 19 - H ^ 19) / (3158507520.001 * G ^ 18)
        
        '辛普森法求定积分近似值
        L1 = H
        L2 = L
        Call simpson2(L1, L2, G, U, V)
        
        x = C + R * (1 - W) * Sin(O / (2 * R)) * Cos(F1 + O / (2 * R)) + W * U * Cos(E) - i * V * Sin(E) '中桩的X坐标
        y = D + R * (1 - W) * Sin(O / (2 * R)) * Sin(F1 + O / (2 * R)) + W * U * Sin(E) + i * V * Cos(E) '中桩的Y坐标
        Q = E + L ^ 2 * W * i / (2 * G ^ 2)                                     '切线方位角，单位弧度
        CEYI x, y, Q, T, P, S, Z, G, H
    Else                                                                        'R=0就提示重新输入参数R
        'MsgBox ("第" & number & "分段的ABR不能为零,请修改参数并按输入按钮。"), vbOKOnly, "提示"
    End If
End Sub

Sub simpson2(L1, L2, G, U, V)
    Step = 16                                                                   'step必须是正偶数2,4,6,8,10...
    Do
        U2 = U
        V2 = V
        HL = (L2 - L1) / Step
        L = L1
        Four = 0
        Two = 0
        Four2 = 0
        Two2 = 0
        For id = 1 To Step / 2
            L = L + HL
            Four = Four + Cos(Fun2(G, L))
            Four2 = Four2 + Sin(Fun2(G, L))
            L = L + HL
            Two = Two + Cos(Fun2(G, L))
            Two2 = Two2 + Sin(Fun2(G, L))
        Next id
        Two = Two - Cos(Fun2(G, L))
        Two2 = Two2 - Sin(Fun2(G, L))
        U = (Cos(Fun2(G, L1)) + Cos(Fun2(G, L2)) + 4 * Four + 2 * Two) * HL / 3
        V = (Sin(Fun2(G, L1)) + Sin(Fun2(G, L2)) + 4 * Four2 + 2 * Two2) * HL / 3
        'Print Step, U, V
        Step = Step + 2
    Loop While Abs(U2 - U) > 0.0001 Or Abs(V2 - V) > 0.0001
End Sub

Private Function Fun2(G, L)
    '缓和曲线切线方位角与弧长的函数关系τ=f(l)
    If G <> 0 Then Fun2 = (L * L) / (2 * G * G)
End Function
Sub 麦克康纳尔曲线(K, x, y, Q, T, P, S, Z, G, H)
    求区间 K, number                                                            '计算出K落在平曲线的哪个区间上，保存在变量Number中
    J = CDec(Val(factor(number, 2)))
    C = CDec(Val(factor(number, 3)))
    D = CDec(Val(factor(number, 4)))
    F = CDec(Val(factor(number, 5)))
    R = CDec(Val(factor(number, 6)))
    Ls = CDec(Val(factor(number, 1)) - Val(factor(number, 2)))                  '对于MC1，Ls=分段终点Khy-Kzh
    If Abs(Ls) < 0.05 Then Ls = Val(factor(number, 1)) - Val(factor(number - 1, 1))
    '对于MC2,Ls=分段终点Khz-Kyh
    v0 = CDec(Val(factor(number, 7)) / 3.6)                                     '设计平衡车速m/s，(120-300)/3.6
    t0 = CDec(Atn(Val(factor(number, 8))))                                      '缓和曲线起点处的路面横向倾角atn(0.7%-0.8%)，弧度
    g0 = CDec((Val(factor(number, 9))))                                         '重力加速度m/s2
    If R <> 0 Then i = R / Abs(R)
    W = 1
    If K < J Then W = -1
    L = Abs(K - J)
    度分秒化为弧度 F, E                                                         'F----度.分秒--弧度
    If Ls <> 0 And R <> 0 Then
        '以下进行二次积分，计算x=∫[0,L]cosτdl，y=∫[0,L]sinτdl，
        If IsEmpty(dL) Then                                                     '如果还没有给dL赋值就赋值
            dL = InputBox("ΔL一般取值0.4～0.5m之间，调整步距ΔL直到HY点或YH点的X坐标与设计相吻合即可。", "请输入麦克康纳尔曲线计算步距", 0.5)
            If dL = "" Or dL = 0 Then Exit Sub
        End If
        n0 = Int(L / dL) + 1                                                    '每隔50cm左右分成一段，分段计算Δx
        'Do
        U2 = U
        V2 = V
        U = 0
        V = 0
        For id = 1 To n0
            '-------------------------------------------------------
            L1 = 0
            L2 = (L / n0) * (id)
            Q = NILgdGauss(L1, L2, 0.00001, R, Ls, v0, t0, g0)                  '高斯求积公式计算任意一点切线方位角
            '            If 0 <= L2 And L2 < Ls / 4 Then
            '                Q = NILgdGauss(0, L2, 0.00001, R, Ls, v0, t0, g0)
            '            ElseIf Ls / 4 <= L2 And Ls < Ls * 3 / 4 Then
            '                Q = NILgdGauss(0, Ls / 4, 0.00001, R, Ls, v0, t0, g0) + NILgdGauss(Ls / 4, L2, 0.00001, R, Ls, v0, t0, g0)
            '            Else
            '                Q = NILgdGauss(0, Ls / 4, 0.00001, R, Ls, v0, t0, g0) + NILgdGauss(Ls / 4, Ls * 3 / 4, 0.00001, R, Ls, v0, t0, g0) + NILgdGauss(Ls * 3 / 4, L2, 0.00001, R, Ls, v0, t0, g0)
            '            End If
            'Call simpson3(L1, L2, R, Ls, v0, t0, g0, Q)                     '辛普森法则计算任意一点的切线方位角
            '-------------------------------------------------------
            '            Q = 0
            '            For id2 = 1 To id
            '                Ri = v0 ^ 2 / g0 / Tan(t0 + phi(R, Ls, v0, t0, g0, L / n0 * id2))
            '                If id2 < id Then
            '                    Q = Q + L / n0 / Ri
            '                Else
            '                    Q = Q + L / n0 / Ri / 2
            '                End If
            '            Next
            '-------------------------------------------------------
            U = U + (L / n0) * Cos(Q)                                           '计算切线支距x坐标
            V = V + (L / n0) * Sin(Q)                                           '计算切线支距y坐标
        Next
        'Q = Q + L / n0 / Ri / 2'计算切线支距坐标系下的切线方位角
        n0 = n0 + 1                                                             '等分数量增加
        'Loop While Abs(U2 - U) > 0.001 Or Abs(V2 - V) > 0.001
        x = C + W * U * Cos(E) - i * V * Sin(E)
        y = D + W * U * Sin(E) + i * V * Cos(E)
        Q = E + Q * W * i                                                       '单位弧度
        CEYI x, y, Q, T, P, S, Z, G, H
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 模块名：NIModule.bas
' 函数名：NILgdGauss
' 功能： 用勒让德－高斯求积法计算一维积分，本函数需要调用计算函数f(x)值的函数Func，其形式为：
'          Function Func(x As Double) As Double
' 参数： a     - Double型变量，积分下限
'          b     - Double型变量，积分上限，要求 b>a
'          eps   - Double型变量，积分精度要求
' 返回值：Double型，积分值
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function NILgdGauss(A, B, eps, R, Ls, v0, t0, g0)
    Dim m As Long, i As Integer, J As Integer
    Dim S As Double, P As Double, ep As Double, H As Double, aa As Double, bb As Double, W As Double, x As Double, G As Double
    Dim T(5) As Double, C(5) As Double
    
    ' 勒让德－高斯求积系数
    T(1) = -0.9061798459
    T(2) = -0.5384693101
    T(3) = 0#
    T(4) = 0.5384693101
    T(5) = 0.9061798459
    
    C(1) = 0.2369268851
    C(2) = 0.4786286705
    C(3) = 0.5688888889
    C(4) = 0.4786286705
    C(5) = 0.2369268851
    
    ' 初值
    m = 1
    H = B - A
    S = Abs(0.001 * H)
    P = 1E+35
    ep = eps + 1#
    
    ' 循环计算
    While ((ep >= eps) And (Abs(H) > S))
        G = 0#
        
        For i = 1 To m
            aa = A + (i - 1#) * H
            bb = A + i * H
            W = 0#
            
            For J = 1 To 5
                L = ((bb - aa) * T(J) + (bb + aa)) / 2#
                W = W + Func(R, Ls, v0, t0, g0, L) * C(J)
            Next J
            
            G = G + W
        Next i
        
        G = G * H / 2#
        ep = Abs(G - P) / (1# + Abs(G))
        P = G
        m = m + 1
        H = (B - A) / m
    Wend
    NILgdGauss = G                                                              ' 返回结果
End Function
Private Function Func(R, Ls, v0, t0, g0, L)
    '麦克康纳尔曲线切线方位角与弧长的函数关系τ=f(l)
    phimax = Atn(v0 ^ 2 / g0 / R) - t0                                          'phimax为圆曲线上的侧摆角
    C0 = 32 * phimax / Ls ^ 3                                                   '计算因子
    If L >= 0 And L < Ls / 4 Then
        Func = g0 / v0 ^ 2 * Tan(t0 + C0 * L ^ 3 / 6)
    ElseIf L >= Ls / 4 And L < 3 * Ls / 4 Then
        Func = g0 / v0 ^ 2 * Tan(t0 + C0 * (-L ^ 3 / 6 + Ls * L ^ 2 / 4 - Ls ^ 2 * L / 16 + Ls ^ 3 / 192))
    ElseIf L >= 3 * Ls / 4 And L <= Ls Then
        Func = g0 / v0 ^ 2 * Tan(t0 + C0 * (L ^ 3 / 6 - Ls * L ^ 2 / 2 + Ls ^ 2 * L / 2 - 13 * Ls ^ 3 / 96))
    End If
End Function

Sub simpson3(L1, L2, R, Ls, v0, t0, g0, Q)
    '计算任意一点的切线方位角，辛普森法则求定积分近似值
    Step = 16                                                                   'step必须是正偶数2,4,6,8,10...
    Do
        Q2 = Q
        HL = (L2 - L1) / Step
        L = L1
        Four = 0
        Two = 0
        For id = 1 To Step / 2
            L = L + HL
            Four = Four + (Fun3(R, Ls, v0, t0, g0, L))
            L = L + HL
            Two = Two + (Fun3(R, Ls, v0, t0, g0, L))
        Next id
        Two = Two - (Fun3(R, Ls, v0, t0, g0, L))
        Q = ((Fun3(R, Ls, v0, t0, g0, L1)) + (Fun3(R, Ls, v0, t0, g0, L2)) + 4 * Four + 2 * Two) * HL / 3
        Step = Step + 2
    Loop While Abs(Q2 - Q) > 0.00001
End Sub

Private Function Fun3(R, Ls, v0, t0, g0, L)
    '麦克康纳尔曲线切线方位角与弧长的函数关系τ=f(l)
    phimax = Atn(v0 ^ 2 / g0 / R) - t0                                          'phimax为圆曲线上的侧摆角
    C0 = 32 * phimax / Ls ^ 3                                                   '计算因子
    If L >= 0 And L < Ls / 4 Then
        Fun3 = g0 / v0 ^ 2 * Tan(t0 + C0 * L ^ 3 / 6)
    ElseIf L >= Ls / 4 And L < 3 * Ls / 4 Then
        Fun3 = g0 / v0 ^ 2 * Tan(t0 + C0 * (-L ^ 3 / 6 + Ls * L ^ 2 / 4 - Ls ^ 2 * L / 16 + Ls ^ 3 / 192))
    ElseIf L >= 3 * Ls / 4 And L <= Ls Then
        Fun3 = g0 / v0 ^ 2 * Tan(t0 + C0 * (L ^ 3 / 6 - Ls * L ^ 2 / 2 + Ls ^ 2 * L / 2 - 13 * Ls ^ 3 / 96))
    End If
End Function
Private Function phi(R, Ls, v0, t0, g0, L)
    '麦克康纳尔曲线切线方位角与弧长的函数关系τ=f(l)
    phimax = Atn(v0 ^ 2 / g0 / R) - t0                                          'phimax为圆曲线上的侧摆角
    C0 = 32 * phimax / Ls ^ 3                                                   '计算因子
    If L >= 0 And L < Ls / 4 Then
        phi = C0 * L ^ 3 / 6
    ElseIf L >= Ls / 4 And L < 3 * Ls / 4 Then
        phi = C0 * (-L ^ 3 / 6 + Ls * L ^ 2 / 4 - Ls ^ 2 * L / 16 + Ls ^ 3 / 192)
    ElseIf L >= 3 * Ls / 4 And L <= Ls Then
        phi = C0 * (L ^ 3 / 6 - Ls * L ^ 2 / 2 + Ls ^ 2 * L / 2 - 13 * Ls ^ 3 / 96)
    End If
End Function
'----------------------------------------------------------------------------------------------------------
Sub CEYI(x, y, Q, T, P, S, Z, G, H)                                             '计算边桩的坐标
    If Val(Q) < 0 Then                                                          '转化到[0,2pai)之间
        Q = Q + 2 * 3.14159265358979
    ElseIf Val(Q) >= 2 * 3.14159265358979 Then
        Q = Q - 2 * 3.14159265358979
    End If
    '    T = CDec(Val(TEXTT))
    '    P = CDec(Val(TEXTP))
    '    S = CDec(Val(TEXTS))
    '    Z = CDec(Val(TEXTZ))
    '    If anniu = 1 Then
    '        T = CDec(Val(factor4(ndh, 2)))
    '        P = CDec(Val(factor4(ndh, 3)))
    '        S = CDec(Val(factor4(ndh, 4)))
    '        Z = CDec(Val(factor4(ndh, 5)))
    '    End If
    度分秒化为弧度 P, p1
    度分秒化为弧度 Z, Z1
    U = x + T * Cos(Q + p1)                                                     '第一偏距、偏角决定的点的X坐标，中间变量
    V = y + T * Sin(Q + p1)                                                     '第一偏距、偏角决定的点的Y坐标，中间变量
    G = U + S * Cos(Q + Z1)                                                     '第二偏距、偏角决定的点的X坐标，最终结果
    H = V + S * Sin(Q + Z1)                                                     '第二偏距、偏角决定的点的Y坐标，最终结果
    xiaoshu = Val(Text114)
    xiaoshu2 = xiaoshu + 2
    G = Round(G, xiaoshu)
    H = Round(H, xiaoshu)
    弧度化为度分秒 Q                                                            '调用子程序把单位转化为度.分秒的格式
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 度分秒化为弧度(F, F1)                                                       '把F(度.分秒)的格式转化为度，然后转化为弧度
    F1 = Fix(F) + (Fix(F * 100) - Fix(F) * 100) / 60 + (F * 10000 - Fix(F * 100) * 100) / 3600 '转化为度
    F1 = F1 * 3.14159265358979 / 180                                            '再转化为弧度
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 弧度化为度分秒(Q)                                                           '把Q(弧度)的格式转化为Q(度)，然后转化为Q（度.分秒）
    'pai = 3.14159265358979
    If Q < 0 Then
        Q = Q + 2 * pai
    ElseIf Q >= 2 * pai Then
        Q = Q - 2 * pai
    End If
    Q = Q * 180 / pai                                                           '度
    Q = Round(Q, 10)
    QD = Fix(Q)                                                                 '度
    QM = (Q - Fix(Q)) * 60                                                      '分
    QS = (QM - Fix(QM)) * 60                                                    '秒
    Q = QD + Fix(QM) / 100 + QS / 10000                                         '度.分秒
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command4_Click()                                                    '退出按钮的事件
    If Saveflag = 1 Or TEXTKI = 0 Then
        Dim ix As Integer
        End                                                                     '如果曲线参数已保存就直接GAME OVER......
    Else
        If MsgBox("平曲线参数未保存,确定退出吗?", vbYesNo, "警告!") = vbYes Then '询问是否退出
            For ix = Forms.Count - 1 To 1 Step -1
                Unload Forms(ix)
            Next
            End                                                                 '如果单击确定就GMAE OVER,单击取消就返回
        End If
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub TEXTQ_Change()
    If TEXTQ <> Q Then                                                          '注意textq=q的赋值语句总会触发TEXTQ_Change事件，为了防止CEYI被再次调用，必须判断！
        x = CDec(Val(TEXTX))
        y = CDec(Val(TEXTY))
        Q = CDec(Val(TEXTQ))                                                    '单位为度.分秒
        度分秒化为弧度 Q, Q                                                     '把单位转化为弧度进行下一步的计算
        T = CDec(Val(TEXTT))
        P = CDec(Val(TEXTP))
        S = CDec(Val(TEXTS))
        Z = CDec(Val(TEXTZ))
        CEYI x, y, Q, T, P, S, Z, G, H
        TEXTG = G
        TEXTH = H
    End If                                                                      '如果CEYI被第二次调用的话，因为Q=round(Q,5)的取舍，G，H将会得出差1-2毫米的不同的结果
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 缓和曲线(K, x, y, Q, G, H)                                                  '缓和曲线坐标计算程序
    Dim W As Integer                                                            'W=1取缓1，W=-1取缓2
    Khy = CDec(Val(Text107))
    J = CDec(Val(Text6))                                                        '加载所有的曲线参数
    C = CDec(Val(Text5))
    D = CDec(Val(Text4))
    F = CDec(Val(Text3))
    R = CDec(Val(Text8))
    A = CDec(Val(Text7))
    If Option7.Value = True Then
        W = 1
    Else
        W = -1
    End If
    
    If A <> 0 And R <> 0 Then
        度分秒化为弧度 F, F1                                                    'F----度.分秒--弧度
        i = R / Abs(R)                                                          '判断左右偏的变量，右偏R取正，左偏取负
        Kzh = Round(Khy - A ^ 2 / Abs(R) * W, 3)
        L = Abs(J - Kzh)                                                        'L=K到ZH（HZ）点的弧长
        H = Abs(K - Kzh)                                                        'H=K到ZH（HZ）点的弧长
        E = F1 - L ^ 2 * W * i / (2 * A ^ 2)                                    '弧度
        U = (L - H) - (L ^ 5 - H ^ 5) / (40 * A ^ 4) + (L ^ 9 - H ^ 9) / (3456 * A ^ 8) _
        - (L ^ 13 - H ^ 13) / (599040 * A ^ 12)
        '+ (L ^ 17 - H ^ 17) / (175472640 * a ^ 16)
        U = -U
        V = (L ^ 3 - H ^ 3) / (6 * A ^ 2) - (L ^ 7 - H ^ 7) / (336 * A ^ 6) + (L ^ 11 - H ^ 11) / (42240 * A ^ 10) _
        - (L ^ 15 - H ^ 15) / (9676800 * A ^ 14)
        '+ (L ^ 19 - H ^ 19) / (3158507520.001 * a ^ 18)
        V = -V
        
        '辛普森法求定积分近似值
        'L1 = H
        'L2 = L
        'Call simpson2(L1, L2, G, U, V)
        
        x = C + W * U * Cos(E) - i * V * Sin(E)                                 '中桩的X坐标
        y = D + W * U * Sin(E) + i * V * Cos(E)                                 '中桩的Y坐标
        Lx = L + Abs(K - J)
        Q = E + H ^ 2 * W * i / (2 * A ^ 2)                                     '切线方位角，单位弧度
        'Q = F1 + (Hs1 ^ 2 - Lx ^ 2) * I / (2 * A1 ^ 2)  '切线方位角，单位弧度
        CEYI2 x, y, Q, G, H
        
        'If CDec(Val(R)) <> 0 And CDec(Val(a)) <> 0 Then
        '度分秒化为弧度 F, F1  'F----度.分秒--弧度
        'W = 1   'W=1取缓1的回旋参数A，W=-1取缓2的回旋参数B
        'If CDec(Val(K)) < CDec(Val(J)) Then
        'W = -1
        'End If
        'I = R / Abs(R)    '判断左右偏的变量，右偏R取正，左偏取负
        'L = K - J    'L=K到ZH（HZ）点的弧长
        'E = F1 + (1 - W) * 3.14159265358979 / 2     'E------弧度
        'U = L - L ^ 5 / (40 * a ^ 4) + L ^ 9 / (3456 * a ^ 8) - L ^ 13 / (599040 * a ^ 12) _
        '+ L ^ 17 / (175472640 * a ^ 16)
        'V = L ^ 3 / (6 * a ^ 2) - L ^ 7 / (336 * a ^ 6) + L ^ 11 / (42240 * a ^ 10) _
        '- L ^ 15 / (9676800 * a ^ 14) + L ^ 19 / (3158507520.001 * a ^ 18)
        'X = C + W * U * Cos(E) - I * V * Sin(E)
        'Y = D + W * U * Sin(E) + I * V * Cos(E)
        'Q = F1 + L ^ 2 * W * I / (2 * a ^ 2)      '切线方位角，单位弧度
        'CEYI2 X, Y, Q, G, H
    Else                                                                        'R=0或A=0就提示重新输入参数AR
        MsgBox ("A和R不能为零,请修改参数。"), vbOKOnly, "提示"
    End If
    xiaoshu = Val(Text114)
    xiaoshu2 = xiaoshu + 2
    x = Round(x, xiaoshu)                                                       '小数位取舍
    y = Round(y, xiaoshu)
    Q = Round(Q, xiaoshu2)
End Sub
'----------------------------------------------------------------------------------------------------------
Sub CEYI2(x, y, Q, G, H)                                                        '边桩的坐标的计算
    T = CDec(Val(Text9))                                                        '继续加载部分曲线参数
    P = CDec(Val(Text10))
    S = CDec(Val(Text11))
    Z = CDec(Val(Text15))
    If CDec(Val(Q)) < 0 Then                                                    '转化到[0,2pai)之间
        Q = Q + 2 * 3.14159265358979
    ElseIf CDec(Val(Q)) >= 2 * 3.14159265358979 Then
        Q = Q - 2 * 3.14159265358979
    End If
    度分秒化为弧度 P, p1
    度分秒化为弧度 Z, Z1
    U = x + T * Cos(Q + p1)                                                     '第一偏距、偏角决定的点的X坐标，中间变量
    V = y + T * Sin(Q + p1)                                                     '第一偏距、偏角决定的点的Y坐标，中间变量
    G = U + S * Cos(Q + Z1)                                                     '第二偏距、偏角决定的点的X坐标，中间变量
    H = V + S * Sin(Q + Z1)                                                     '第二偏距、偏角决定的点的Y坐标，中间变量
    G = Round(G, 3)
    H = Round(H, 3)
    弧度化为度分秒 Q                                                            '调用子程序把单位转化为度.分秒的格式
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command14_Click()                                                   '从文件载入缓和曲线参数
    With CommonDialog1
        .FileName = ""
        .CancelError = False
        .DialogTitle = "从文件载入缓和曲线参数..."
        .DefaultExt = "inf"
        .Filter = "输入的曲线参数(*.inf)|*.inf"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text21 = .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 1
    Input #1, J
    Input #1, C
    Input #1, D
    Input #1, F
    Input #1, R
    Input #1, A
    If Not EOF(1) Then
        Input #1, Khy
        Input #1, W
        Text107 = Khy
        If W = -1 Then
            Option8.Value = True
        Else
            Option7.Value = True
        End If
        Input #1, jianju
    End If
    Text6 = J
    Text5 = C
    Text4 = D
    Text3 = F
    Text8 = R
    Text7 = A
    If jianju <> 0 Then Text14 = jianju
    
    Rx = R
    Ax = A
    If Ax <> 0 And Rx <> 0 Then
        度分秒化为弧度 F, F1                                                    'F----度.分秒--弧度
        ix = Rx / Abs(Rx)                                                       '判断左右偏的变量，右偏R取正，左偏取负
        Hx = Ax ^ 2 / Abs(Rx)                                                   '完整的缓和曲线总长
    End If
    
    If Option7.Value = True Then
        Label1(134) = "HY点里程KHY"
        Text12 = Round(Val(Text107) - Hx, 3)
        Text13 = Val(Text107)
    Else
        Label1(134) = "YH点里程KYH"
        Text12 = Text107
        Text13 = Round(Val(Text107) + Hx, 3)
    End If
    
    Close #1
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command13_Click()                                                   '保存缓和曲线参数到文件
    With CommonDialog1
        .FileName = ""
        .DialogTitle = "保存缓和曲线参数到文件..."
        .DefaultExt = "inf"
        .Filter = "输出的曲线参数(*.inf)|*.inf"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text21 = .FileTitle
    End With
    Open CommonDialog1.FileName For Output As 2
    J = Text6
    C = Text5
    D = Text4
    F = Text3
    R = Text8
    A = Text7
    Print #2, J
    Print #2, C
    Print #2, D
    Print #2, F
    Print #2, R
    Print #2, A
    Print #2, Val(Text107)
    If Option7.Value = True Then Print #2, 1
    If Option8.Value = True Then Print #2, -1
    Print #2, Val(Text14)
    Close #2
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command15_Click()                                                   '查看可计算的范围
    Rx = CDec(Val(Text8))
    Ax = CDec(Val(Text7))
    If Ax <> 0 And Rx <> 0 Then
        度分秒化为弧度 F, F1                                                    'F----度.分秒--弧度
        ix = Rx / Abs(Rx)                                                       '判断左右偏的变量，右偏R取正，左偏取负
        Hx = Ax ^ 2 / Abs(Rx)                                                   '完整的缓和曲线总长
    End If
    
    If Option7.Value = True Then
        Label1(134) = "HY点里程KHY"
        Text12 = Round(Val(Text107) - Hx, 3)
        Text13 = Val(Text107)
    Else
        Label1(134) = "YH点里程KYH"
        Text12 = Text107
        Text13 = Round(Val(Text107) + Hx, 3)
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command17_Click()
    Dim L As Double
    R = CDec(Val(Text8))
    A = CDec(Val(Text7))
    If R <> 0 And A <> 0 Then
        L = A ^ 2 / R
        J = CDec(Val(Text6))
        Kstart = CDec(Val(Text12))
        缓和曲线 Kstart, x, y, Q, G, H
        Text18 = x
        Text17 = y
        Text16 = Q
        Text19 = G
        Text20 = H
    Else                                                                        'R=0就提示重新输入参数R
        MsgBox "A和R不能为零,请修改参数！", vbCritical, "警告！"
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command16_Click()                                                   '生成中桩坐标的AUTOCAD脚本文件
    If Text12 <> "" And Text6 <> "" And Val(Text14) <> 0 Then
        With CommonDialog1
            .FileName = ""
            .CancelError = False
            .DialogTitle = "生成中桩坐标的AUTOCAD脚本文件..."
            .DefaultExt = "scr"
            .Filter = "AUTOCAD脚本文件(*.scr)|*.scr"
            .ShowSave
            If Len(.FileName) = 0 Then: Close: Exit Sub
        End With
        Open CommonDialog1.FileName For Output As 5
        Print #5, "osnapcoord"
        Print #5, "1"                                                           '设为1也可以，2=键盘输入优先，脚本例外
        Print #5, "UCS"
        Print #5, "W"
        'Print #5, "UCS"
        'Print #5, "N"
        'Print #5, "3"
        'Print #5, "0,0"
        'Print #5, "0,1"
        'Print #5, "1,0"
        
        '新建图层？？？中桩
        Print #5, "-layer"                                                      '首先新建图层
        Print #5, "m"                                                           'm=生成并设置为当前层
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "中桩" '新建图层名称为？？？中桩
        Print #5, "c"                                                           '修改图层颜色
        Print #5, Text116                                                       '1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "中桩" '输入需要修改颜色的图层里程标注
        Print #5, ""                                                            '结束图层设置命令
        
        '画线
        Print #5, "line"
        For K = CDec(Val(Text12)) To CDec(Val(Text13)) Step CDec(Val(Text14))
            缓和曲线 K, x, y, Q, G, H
            Write #5, y, x                                                      '大地坐标系下，XY必须互换
        Next K
        K = CDec(Val(Text13))
        缓和曲线 K, x, y, Q, G, H
        Write #5, y, x                                                          '大地坐标系下，XY必须互换
        Print #5, ""                                                            '结束line命令
        
        '标注里程
        Print #5, "-layer"                                                      '首先新建图层
        Print #5, "m"                                                           'm=生成并设置为当前层
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "里程标注" '新建图层名称为里程标注
        Print #5, "c"                                                           '修改图层颜色
        Print #5, Text116                                                       '1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "里程标注" '输入需要修改颜色的图层里程标注
        Print #5, ""                                                            '结束图层设置命令
        '设置文字样式
        Height1 = Val(Text14) / 10                                              '设置文字高度
        length1 = Val(Text14) / 10                                              '设置引线长度
        length2 = Val(Text14) / 25                                              '设置引线偏移距离
        Print #5, "-style"                                                      '建立文字样式
        Print #5, Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4) & "里程标注" '文字样式名称
        Print #5, "simfang.ttf"                                                 '字体名称
        Write #5, Height1                                                       '文字高度
        Print #5, "1"                                                           '宽度比例
        Print #5, "0"                                                           '倾斜角度
        Print #5, "n"                                                           '反向显示文字=否
        Print #5, "n"                                                           '颠倒显示文字=否
        
        For K = CDec(Val(Text12)) To CDec(Val(Text13)) Step CDec(Val(Text14))
            缓和曲线 K, x, y, Q, G, H
            度分秒化为弧度 Q, QHD                                               '把度.分秒格式转化为弧度保存在QHD中
            QHD = (-QHD) * 180 / 3.14159265358979                               '把弧度转化成度
            Print #5, "line"
            Write #5, y, x                                                      '在世界坐标系下画图XY必须互换，Z=0
            Print #5, "@" & length1 & "<" & Format(QHD, "#0.#######")
            Print #5, ""                                                        '结束line命令
            '标注文字
            Print #5, "text"                                                    '输入单行文字text
            Print #5, "j"                                                       '修改文字格式
            Print #5, "ml"                                                      '左中对齐
            Print #5, "@" & length2 & "<" & Format(QHD, "#0.#######")           '文字起点位置
            Print #5, Format(QHD, "#0.#######")                                 '文字标注方向
            Write #5, K                                                         '里程
            '注意：单行文字命令自动结束的！不需要用回车来结束
        Next K
        '设置当前图层为0层
        Print #5, "-layer"                                                      '设置图层
        Print #5, "s"                                                           '设置当前图层
        Print #5, "0"                                                           '设置当前图层为0层
        Print #5, ""                                                            '结束-layer命令
        Print #5, "zoom"
        Print #5, "e"
        
        Close #5
        msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
        If MsgBox(msg, vbYesNo) = vbYes Then
            Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbNormalFocus
        End If
    Else
        MsgBox "请重新输入各参数，确保间距不为零！", vbCritical, "警告！"
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command18_Click()                                                   '导出全部坐标到文件
    If Text12 <> "" And Text6 <> "" And Val(Text14) <> 0 Then
        With CommonDialog1
            .FileName = ""
            .CancelError = True
            .DialogTitle = "导出全部坐标到文件..."
            .DefaultExt = "wtx"
            .Filter = "坐标结果(*.wtx)|*.wtx"
            .ShowSave
            If Len(.FileName) = 0 Then: Close: Exit Sub
        End With
        Open CommonDialog1.FileName For Output As 7
        For K = CDec(Val(Text12)) To CDec(Val(Text13)) Step CDec(Val(Text14))
            缓和曲线 K, x, y, Q, G, H
            Write #7, K, x, y, Q, G, H
        Next K
        Close #7
        msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
        If MsgBox(msg, vbYesNo) = vbYes Then
            Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
        End If
    Else
        MsgBox "请重新输入各参数，确保间距不为零！", vbCritical, "警告！"
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command10_Click()                                                   '计算坐标和方位角按钮
    Text57 = TextX1
    Text58 = TextY1
End Sub
Sub 方位角()
    'pai = 3.14159265358979
    x1 = CDec(Val(TextX1))
    y1 = CDec(Val(TextY1))
    x2 = CDec(Val(TextX2))
    y2 = CDec(Val(TextY2))
    S = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
    If x2 <> x1 And y2 <> y1 Then
        Q = Atn((y2 - y1) / (x2 - x1)) + (1 - Abs(x2 - x1) / (x2 - x1)) * pai / 2 + (1 + Abs(x2 - x1) / (x2 - x1)) * (1 - Abs(y2 - y1) / (y2 - y1)) * pai / 2
        弧度化为度分秒 Q
    ElseIf x2 = x1 Then                                                         '如果X2=X1 会被零除，要单独计算
        If y2 > y1 Then Q = 90
        If y2 < y1 Then Q = 270
    ElseIf y2 = y1 Then                                                         '如果Y2=Y1 会被零除，要单独计算
        If x2 > x1 Then Q = 0
        If x2 < x1 Then Q = 180
    End If
    Text23 = Round(S, 7)
    Text22 = Round(Q, 5)
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Text22_click()                                                      '一单击就选中所有的文字，便于复制粘贴
    With Text22
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Text23_click()                                                      '一单击就选中所有的文字，便于复制粘贴
    With Text23
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Text24_Change()                                                     '竖曲线段数改变
    Dim I1                                                                      '循环用的变量
    Combo3.Clear
    For I1 = 1 To (Val(Text24))
        Combo3.AddItem I1
    Next I1
    If Combo3.ListCount > 0 Then Combo3.ListIndex = 0
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Text33_Change()                                                     '横坡段数改变
    Dim I1                                                                      '循环用的变量
    Combo4.Clear
    For I1 = 1 To (Val(Text33))
        Combo4.AddItem I1
    Next I1
    If Combo4.ListCount > 0 Then Combo4.ListIndex = 0
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command12_Click()                                                   '保存竖曲线参数
    Dim n1, n2 As Integer                                                       '循环用的变量
    With CommonDialog1
        .FileName = ""
        .CancelError = False
        .DialogTitle = "保存竖曲线参数到文件..."
        .DefaultExt = "css"
        .Filter = "输出的竖曲线参数(*.css)|*.css"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text41 = .FileTitle
    End With
    Open CommonDialog1.FileName For Output As 2
    Partnumber = Val(Text24)
    Write #2, Val(Text24)
    '    Kqd = Val(TextKQD)
    '    Write #2, Val(TextKQD)
    
    
    '新版本格式：
    For n1 = 1 To Val(Text24)
        Write #2, shu(n1, 1), shu(n1, 2), shu(n1, 3)
    Next n1
    
    Write #2, Val(Text39)
    Write #2, Val(Text43)
    Write #2, Val(Text47)
    
    Close #2
    Saveflag = 1
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Combo4_click()                                                      '单击分段列表就载入横坡参数
    载入横坡参数
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command22_Click()                                                   '载入横坡参数按钮
    Dim n1, n2 As Integer                                                       '循环用的变量
    With CommonDialog1
        .FileName = ""
        .CancelError = False
        .DialogTitle = "从文件载入横坡参数..."
        .DefaultExt = "hpo"
        .Filter = "横坡参数(*.hpo)|*.hpo"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        Text42 = .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 1
    
    Dim kk() As String
    Dim str1 As String
    Line Input #1, str1
    kk = Split(str1, ",")
    If UBound(kk) = 0 Then GoTo Label1
    '新版本格式：
    Text33 = Val(kk(0))                                                         '分段数目
    Text34 = Val(kk(1))                                                         '起点里程
    partnumber1 = Val(kk(0))
    Erase hengpo                                                                '清空数组便于多次载入
    For n1 = 1 To partnumber1
        Input #1, hengpo(n1, 0), hengpo(n1, 1), hengpo(n1, 2), hengpo(n1, 3)
    Next n1
    Close #1
    载入横坡参数
    Exit Sub
    
Label1:                                                                         '老版本格式
    Text33 = Val(kk(0))
    partnumber1 = Val(kk(0))
    Input #1, Kqd1                                                              '把文件中的分段数、起点里程读入到相应的变量中
    Text34 = Kqd1
    Erase hengpo                                                                '清空数组便于多次载入
    For n1 = 1 To partnumber1
        Input #1, n1
        For n2 = 0 To 2
            Input #1, hengpo(n1, n2)
        Next n2
    Next n1
    Close #1
    载入横坡参数
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 载入横坡参数()
    Text32 = CDec(Val(hengpo(Val(Combo4), 0)))                                  'hengpo(段数,0列)存储分段终点里程KI
    Text31 = CDec(Val(hengpo(Val(Combo4), 1)))                                  'hengpo(段数,1列)存储变坡点左横坡
    Text30 = CDec(Val(hengpo(Val(Combo4), 2)))                                  'hengpo(段数,2列)存储变坡点右横坡
    If hengpo(Val(Combo4), 3) = 0 Then Option1.Value = True                     '线性渐变
    If hengpo(Val(Combo4), 3) = 1 Then Option2.Value = True                     '抛物线渐变
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command20_Click()
    保存横坡参数
    If Combo4.ListIndex < Combo4.ListCount - 1 Then Combo4.ListIndex = Combo4.ListIndex + 1
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 保存横坡参数()
    hengpo(Val(Combo4), 0) = CDec(Val(Text32))                                  'hengpo(段数,0列)存储分段终点里程KI
    hengpo(Val(Combo4), 1) = CDec(Val(Text31))                                  'hengpo(段数,1列)存储变坡点左横坡
    hengpo(Val(Combo4), 2) = CDec(Val(Text30))                                  'hengpo(段数,2列)存储变坡点右横坡
    If Option1.Value = True Then hengpo(Val(Combo4), 3) = 0                     'hengpo第3列=0线性渐变，=1抛物线渐变
    If Option2.Value = True Then hengpo(Val(Combo4), 3) = 1
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command21_Click()                                                   '保存横坡参数到文件
    Dim n1, n2 As Integer                                                       '循环用的变量
    With CommonDialog1
        .FileName = ""
        .CancelError = False
        .DialogTitle = "保存横坡参数到文件..."
        .DefaultExt = "hpo"
        .Filter = "横坡参数(*.hpo)|*.hpo"
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
    End With
    Open CommonDialog1.FileName For Output As 2
    partnumber1 = Val(Text33)
    Kqd1 = Val(Text34)
    Write #2, Val(Text33), Val(Text34)
    For n1 = 1 To Val(Text33)
        Write #2, hengpo(n1, 0), hengpo(n1, 1), hengpo(n1, 2), hengpo(n1, 3)
    Next n1
    Close #2
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command19_Click()
    保存竖曲线参数
    With Combo3
        If .ListIndex < .ListCount - 1 Then .ListIndex = .ListIndex + 1
    End With
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 保存竖曲线参数()
    On Error Resume Next                                                        '如果被零除就跳过
    shu(Val(Combo3), 1) = CDec(Val(Text29))                                     'shu(段数,1列)存储起算点里程J
    shu(Val(Combo3), 2) = CDec(Val(Text28))                                     'shu(段数,2列)存储起算点里程Y
    shu(Val(Combo3), 3) = CDec(Val(Text27))                                     'shu(段数,3列)存储半径R
    If Val(Combo3) >= 3 Then                                                    '从第三个边坡点开始计算分段终点，第一第二坡度，第一变坡点只做定向用
        shu(Val(Combo3) - 1, 4) = (shu(Val(Combo3) - 1, 2) - shu(Val(Combo3) - 2, 2)) / (shu(Val(Combo3) - 1, 1) - shu(Val(Combo3) - 2, 1)) 'shu(段数,4列)存储第一坡度,上坡取正,下坡取负
        shu(Val(Combo3) - 1, 5) = (shu(Val(Combo3), 2) - shu(Val(Combo3) - 1, 2)) / (shu(Val(Combo3), 1) - shu(Val(Combo3) - 1, 1)) 'shu(段数,5列)存储第二坡度,上坡取正,下坡取负
        shu(Val(Combo3) - 1, 0) = shu(Val(Combo3) - 1, 1) + Abs((shu(Val(Combo3) - 1, 4) - shu(Val(Combo3) - 1, 5)) * shu(Val(Combo3) - 1, 3)) / 2 'shu(段数,0列)存储分段终点里程KI
        If Val(Combo3) = Val(Text24) Then                                       '最后一个变坡点，终点里程取J+T和KZD的最大值
            If shu(Val(Combo3), 1) > shu(Val(Combo3) - 1, 0) Then
                shu(Val(Combo3) - 1, 0) = shu(Val(Combo3), 1)
                shu(Val(Combo3), 0) = shu(Val(Combo3), 1)
            End If
        End If
        '    Else                                                                        '第一段
        '        shu(Val(Combo3), 0) = CDec(Val(TextKQD))                                'shu(段数,0列)存储分段终点里程KI
        '        shu(Val(Combo3), 4) = CDec(Val(TEXTE))                                  'shu(段数,4列)存储第一坡度,不带%号,上坡取正,下坡取负
        '        shu(Val(Combo3), 5) = CDec(Val(Text26))                                 'shu(段数,5列)存储第二坡度,不带%号,上坡取正,下坡取负
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 载入竖曲线参数()
    On Error Resume Next                                                        '如果被零除就跳过
    Text29 = CDec(Val(shu(Val(Combo3), 1)))                                     'shu(段数,1列)存储起算点里程J
    Text28 = CDec(Val(shu(Val(Combo3), 2)))                                     'shu(段数,2列)存储起算点里程Y
    Text27 = CDec(Val(shu(Val(Combo3), 3)))                                     'shu(段数,3列)存储半径R
    Text25 = CDec(Val(shu(Val(Combo3), 0)))                                     'shu(段数,0列)存储分段终点里程KI
    TEXTE = Format((Val(shu(Val(Combo3), 4))) * 100, "0.#####")                 'shu(段数,4列)存储第一坡度,不带%号,上坡取正,下坡取负
    Text26 = Format((Val(shu(Val(Combo3), 5))) * 100, "0.#####")                'shu(段数,5列)存储第二坡度,不带%号,上坡取正,下坡取负
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Combo3_click()                                                      '载入竖曲线参数
    载入竖曲线参数
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command8_Click()                                                    '载入竖曲线参数按钮
    Dim n1, n2 As Integer                                                       '循环用的变量
    With CommonDialog1
        .FileName = ""
        .CancelError = False
        .DialogTitle = "从文件载入竖曲线参数..."
        .DefaultExt = "css"
        .Filter = "输入的竖曲线参数(*.css)|*.css"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text41 = .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 1
    Input #1, Partnumber
    Text24.text = Partnumber
    '    Input #1, Kqd                                                               '把文件中的分段数、起点里程、分段序号读入到相应的变量中
    '    TextKQD.Text = Kqd
    '清空数组便于多次载入
    Erase shu
    On Error Resume Next
    Dim kk() As String
    Dim str1 As String
    Erase kk
    Line Input #1, str1
    kk = Split(str1, ",")
    
    '老版本格式
    If UBound(kk) = 0 Then
        banben = 3.3
        Text24.text = Partnumber + 2
        For n1 = 1 To Partnumber
            Input #1, n1
            For n2 = 0 To 5
                Input #1, shu(n1 + 1, n2)
            Next n2
        Next n1
        'T2 = Abs((shu(2, 4) - shu(2, 5)) / 100 * shu(2, 3)) / 2
        'shu(1, 1) = shu(2, 1) - T2
        shu(1, 0) = Val(kk(0))
        shu(1, 1) = Val(kk(0))
        shu(1, 2) = shu(2, 2) - shu(2, 4) / 100 * (shu(2, 1) - shu(1, 1))
        shu(Partnumber + 2, 0) = shu(Partnumber + 1, 0)
        shu(Partnumber + 2, 1) = shu(Partnumber + 1, 0)
        shu(Partnumber + 2, 2) = shu(Partnumber + 1, 2) + shu(Partnumber + 1, 5) / 100 * (shu(Partnumber + 2, 1) - shu(Partnumber + 1, 1))
        '        tn = Abs((shu(Partnumber - 1, 4) - shu(Partnumber - 1, 5)) / 100 * shu(Partnumber - 1, 3)) / 2
        '        shu(Partnumber, 1) = shu(Partnumber - 1, 1) + tn
        '        shu(Partnumber, 2) = shu(Partnumber - 1, 2) + shu(Partnumber - 1, 5) / 100 * tn
    End If
    
    '新版本格式
    If UBound(kk) > 0 Then
        banben = 3.4
        Close #1
        Open CommonDialog1.FileName For Input As 1
        Line Input #1, Partnumber
        '        Line Input #1, Kqd
        '        shu(1, 0) = Kqd
        '        For n1 = 1 To Partnumber
        '            Input #1, shu(n1, 0), shu(n1, 1), shu(n1, 2), shu(n1, 3), shu(n1, 4), shu(n1, 5)
        '        Next n1
        For n1 = 1 To Partnumber
            Input #1, shu(n1, 1), shu(n1, 2), shu(n1, 3)
        Next n1
        shu(1, 0) = shu(1, 1)
    End If
    
    '计算KI，E，F
    For fd = 3 To Val(Text24)                                                   '从第三个边坡点开始计算分段终点，第一第二坡度，第一变坡点只做定向用
        shu(fd - 1, 4) = (shu(fd - 1, 2) - shu(fd - 2, 2)) / (shu(fd - 1, 1) - shu(fd - 2, 1)) 'shu(段数,4列)存储第一坡度,上坡取正,下坡取负
        shu(fd - 1, 5) = (shu(fd, 2) - shu(fd - 1, 2)) / (shu(fd, 1) - shu(fd - 1, 1)) 'shu(段数,5列)存储第二坡度,上坡取正,下坡取负
        shu(fd - 1, 0) = Format(shu(fd - 1, 1) + Abs((shu(fd - 1, 4) - shu(fd - 1, 5)) * shu(fd - 1, 3)) / 2, "0.000") 'shu(段数,0列)存储分段终点里程KI
        '    Else                                                                        '第一段
        '        shu(fd, 0) = CDec(Val(TextKQD))                                'shu(段数,0列)存储分段终点里程KI
        '        shu(fd, 4) = CDec(Val(TEXTE))                                  'shu(段数,4列)存储第一坡度,不带%号,上坡取正,下坡取负
        '        shu(fd, 5) = CDec(Val(Text26))                                 'shu(段数,5列)存储第二坡度,不带%号,上坡取正,下坡取负
    Next
    '最后两个变坡点，终点里程取J+T和KZD的最大值
    shu(Val(Text24) - 1, 0) = shu(Val(Text24), 1)
    shu(Val(Text24), 0) = shu(Val(Text24), 1)
    
    
    载入竖曲线参数
    Text37 = shu(1, 1)                                                          '为起算点、终点、间距赋初值，默认计算全线标高
    Text38 = shu(Val(Text24), 0)
    If EOF(1) = False Then
        Input #1, jianju
        Text39 = jianju
        Input #1, ts
        Text43 = ts
        Input #1, chazhi
        Text47 = chazhi
    End If
    Close #1
    Saveflag = 1
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 计算设计高程(K, J, y, R, E, F, H)
    On Error Resume Next
    Dim W As Double                                                             '   '
    Dim Q As Double                                                             '切线长
    Dim U As Double
    求区间2 K, number                                                           '计算出K落在竖曲线的哪个区间上，保存在变量Number中
    J = CDec(Val(shu(number, 1)))
    y = CDec(Val(shu(number, 2)))
    R = CDec(Val(shu(number, 3)))
    E = CDec(Val(shu(number, 4)))
    F = CDec(Val(shu(number, 5)))
    i = CDec(Val(TEXTI))
    W = Abs(E - F)
    Q = W * Abs(R) / 2
    '    If R <> 0 And W <> 0 Then
    U = W / (E - F)
    If K <= J - Q Then
        H = y - E * (J - K)
    ElseIf K <= J Then
        H = y - E * (J - K) - U * (K - J + Q) ^ 2 / (2 * R)
    ElseIf K <= J + Q Then
        H = y + F * (K - J) - U * (J + Q - K) ^ 2 / (2 * R)
    Else
        H = y + F * (K - J)
    End If
    xiaoshu = Val(Text115)
    H = Round(H, xiaoshu)
    'Else                                                                            'R=0或W=0就提示重新输入参数
    '    MsgBox "第" & number & "分段的半径R和坡度差(E-F)不能为零,请修改参数并按输入按钮。", vbCritical, "警告！"
    'End If
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 计算边桩高程(K, T, P, KG, H, i, G)
    On Error Resume Next
    度分秒化为弧度 P, p1
    KG = K + T * Cos(p1)
    计算设计高程 KG, J, y, R, E, F, H                                           '计算边桩对应里程的设计高程
    hengpo(0, 0) = CDec(Val(Text34))
    hengpo(0, 1) = hengpo(1, 1)
    hengpo(0, 2) = hengpo(1, 2)
    求区间1 KG, number1                                                         '计算出KG落在横坡的哪个区间上，保存在变量Number1中
    If number1 <= Val(Text33) Then
        'KG >= Val(Text34) And
        '2011年2月23日修改
        K1 = CDec(Val(hengpo(number1 - 1, 0)))
        K2 = CDec(Val(hengpo(number1, 0)))
        If T < 0 Then
            I1 = CDec(Val(hengpo(number1 - 1, 1)))
            i2 = CDec(Val(hengpo(number1, 1)))
        Else
            I1 = CDec(Val(hengpo(number1 - 1, 2)))
            i2 = CDec(Val(hengpo(number1, 2)))
        End If
        
        If Option1.Value = True Then
            If K2 <> K1 Then i = I1 + (KG - K1) * (i2 - I1) / (K2 - K1)
            If K2 = K1 Then i = I1
        ElseIf Option2.Value = True Then
            If K2 <> K1 Then Kx = (KG - K1) / (K2 - K1): i = I1 + (3 * Kx ^ 2 - 2 * Kx ^ 3) * (i2 - I1)
            If K2 = K1 Then i = I1
        End If
        
        ts = CDec(Val(Text43))
        G = H - i * (Abs(T * Sin(p1)) - Abs(ts))
        
    Else
        i = 0
        G = 0
    End If
    xiaoshu = Val(Text115)
    i = Round(i, xiaoshu + 3)
    TEXTI = i
    G = Round(G, xiaoshu)
    KG = Round(KG, xiaoshu)
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 求区间2(K, number)                                                          '求K所在的竖曲线区间，保存在number变量里
    number = 1                                                                  '先赋个初值
    If Val(shu(1, 0)) < Val(shu(2, 0)) Then
        Do While K > CDec(Val(shu(number, 0))) And number < Val(Text24)
            number = number + 1
        Loop                                                                    '跳出循环后，Number为K所在的区间
    End If
    If Val(shu(1, 0)) > Val(shu(2, 0)) Then
        Do While K < CDec(Val(shu(number, 0))) And number < Val(Text24)
            number = number + 1
        Loop                                                                    '跳出循环后，Number为K所在的区间
    End If
    If Combo3.ListCount > 0 Then Combo3.ListIndex = number - 1
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 求区间1(K, number1)                                                         '求K所在的横坡区间，保存在number1变量里
    number1 = 1                                                                 '先赋个初值
    If Val(hengpo(1, 0)) < Val(hengpo(2, 0)) Then
        Do While K > CDec(Val(hengpo(number1, 0))) And number1 < Val(Text33)
            number1 = number1 + 1
        Loop                                                                    '跳出循环后，Number1为K所在的区间
    End If
    If Val(hengpo(1, 0)) > Val(hengpo(2, 0)) Then
        Do While K < CDec(Val(hengpo(number1, 0))) And number1 < Val(Text33)
            number1 = number1 + 1
        Loop                                                                    '跳出循环后，Number1为K所在的区间
    End If
    If Combo4.ListCount > 0 Then Combo4.ListIndex = number1 - 1
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command23_Click()                                                   '计算高程按钮
    Dim Ki As Double
    Kstart = CDec(Val(Text37))
    '    Kqd = CDec(Val(TextKQD))
    Kqd = shu(1, 1)
    Partnumber = CDec(Val(Text24))
    'If Kstart >= Kqd And Kstart <= CDec(Val(shu(Partnumber, 0))) Then
    T = CDec(Val(Text35))
    P = CDec(Val(Text44))
    计算边桩高程 Kstart, T, P, KG, H, i, G
    Text45 = KG
    Text36 = H
    TEXTI = Format(i, "0.#####")
    Text40 = Format(G, "0.#####")
    
    If Kstart < CDec(Val(Text34)) Then
        MsgBox "从" & Kstart & "到" & Val(Text34) & "之间无横坡参数，只能计算中桩高程！", vbInformation, "提示"
    ElseIf Kstart > hengpo(Val(Text33), 0) Then
        MsgBox "从" & hengpo(Val(Text33), 0) & "到" & Kstart & "之间无横坡参数，只能计算中桩高程！", vbInformation, "提示"
    End If
    
    'Else
    'Ki = CDec(Val(shu(Partnumber, 0)))
    'Text37 = Val(InputBox("K必须在" & Kqd & "到" & Ki & "之间！请重新输入："))
    'Kstart = CDec(Val(Text37))
    'End If
    
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command25_Click()                                                   '生成中桩高程的AUTOCAD脚本文件
    '    If CDec(Val(Text37)) >= CDec(Val(TextKQD)) And Val(Text38) <= shu(Val(Text24), 0) And Text25 <> "" And CDec(Val(Text39)) <> 0 Then
    If CDec(Val(Text37)) >= shu(1, 0) And Val(Text38) <= shu(Val(Text24), 0) And Text25 <> "" And CDec(Val(Text39)) <> 0 Then
        With CommonDialog1
            .FileName = ""
            .CancelError = False
            .DialogTitle = "生成中桩高程的AUTOCAD脚本文件..."
            .DefaultExt = "scr"
            .Filter = "AUTOCAD脚本文件(*.scr)|*.scr"
            .ShowSave
            If Len(.FileName) = 0 Then Exit Sub
        End With
        Open CommonDialog1.FileName For Output As 5
        Print #5, "setvar"
        Print #5, "osnapcoord"
        Print #5, "1"                                                           '设为1=键盘输入优先也可以，2=键盘输入优先(脚本例外)
        Print #5, "UCS"
        Print #5, "W"
        Print #5, "line"
        Kstart = CDec(Val(Text37))
        Kend = CDec(Val(Text38))
        offset = CDec(Val(Text39))
        For K = Kstart To Kend Step offset
            计算设计高程 K, J, y, R, E, F, H
            Write #5, K, H
            'Write #5, K, H, K    '把里程K写入Z坐标是为了可以在AUTOCAD中仅Scale X轴时，可以从Z轴看出其里程
        Next K
        Print #5, ""                                                            '结束line命令
        
        '标注里程
        Print #5, "-layer"                                                      '首先新建图层
        Print #5, "m"                                                           'm=生成并设置为当前层
        Print #5, "里程标注"                                                    '新建图层名称为里程标注
        Print #5, "c"                                                           '修改图层颜色
        Print #5, "1"                                                           '1红色 2黄色 3绿色 4青色 5蓝色 6洋红 7白 8灰，可以查看cad颜色索引
        Print #5, "里程标注"                                                    '输入需要修改颜色的图层里程标注
        Print #5, ""                                                            '结束图层设置命令
        
        '设置文字样式
        Height1 = Val(Text39) / 10                                              '设置文字高度
        'length1 = Val(Text39) / 10 '设置引线长度
        length2 = Val(Text39) / 25                                              '设置引线偏移距离
        Print #5, "-style"                                                      '建立文字样式
        Print #5, "里程标注"                                                    '文字样式名称
        Print #5, "simfang.ttf"                                                 '字体名称
        Write #5, Height1                                                       '文字高度
        Print #5, "1"                                                           '宽度比例
        Print #5, "0"                                                           '倾斜角度
        Print #5, "n"                                                           '反向显示文字=否
        Print #5, "n"                                                           '颠倒显示文字=否
        
        Kstart = CDec(Val(Text37))
        Kend = CDec(Val(Text38))
        offset = CDec(Val(Text39))
        For K = Kstart To Kend Step offset
            计算设计高程 K, J, y, R, E, F, H
            Print #5, "line"
            Write #5, K, H
            Write #5, K, 0
            Print #5, ""                                                        '结束line命令
            '标注设计高程
            Print #5, "text"                                                    '输入单行文字text
            Print #5, "j"                                                       '修改文字格式
            Print #5, "mr"                                                      '右中对齐
            Write #5, K, -2 * length2                                           '文字起点位置
            Write #5, 90                                                        '文字标注方向
            Write #5, H                                                         '设计高程
            '注意：单行文字命令自动结束的！不需要用回车来结束
            '标注里程
            Print #5, "text"                                                    '输入单行文字text
            Print #5, "j"                                                       '修改文字格式
            Print #5, "mr"                                                      '右中对齐
            Write #5, K, -6 * Height1 - 2 * length2                             '文字起点位置
            Write #5, 90                                                        '文字标注方向
            Write #5, K                                                         '里程
            '注意：单行文字命令自动结束的！不需要用回车来结束
        Next K
        
        '画水平线h=0
        Print #5, "line"
        Write #5, Kstart - offset - 12 * Height1, 0
        Write #5, K, 0
        Print #5, ""                                                            '结束line命令
        '画水平线h=-6倍文字高度
        Print #5, "line"
        Write #5, Kstart - offset - 12 * Height1, -6 * Height1
        Write #5, K, -6 * Height1
        Print #5, ""                                                            '结束line命令
        '画水平线h=-12倍文字高度
        Print #5, "line"
        Write #5, Kstart - offset - 12 * Height1, -12 * Height1
        Write #5, K, -12 * Height1
        Print #5, ""                                                            '结束line命令
        '画竖直线h=H+12倍文字高度
        Print #5, "line"
        Write #5, Kstart - offset, -12 * Height1
        Write #5, Kstart - offset, H + 12 * Height1
        Print #5, ""                                                            '结束line命令
        '画竖直线h=H+12倍文字高度
        Print #5, "line"
        Write #5, Kstart - offset - 12 * Height1, -12 * Height1
        Write #5, Kstart - offset - 12 * Height1, H + 12 * Height1
        Write #5, K, H + 12 * Height1
        Write #5, K, -12 * Height1
        Print #5, ""                                                            '结束line命令
        
        '文字标注：标注设计高程，里程等
        Height1 = Val(Text39) / 10                                              '设置文字高度
        'length1 = Val(Text39) / 10 '设置引线长度
        length2 = Val(Text39) / 25                                              '设置引线偏移距离
        Print #5, "-style"                                                      '建立文字样式
        Print #5, "文字标注"                                                    '文字样式名称
        Print #5, "simfang.ttf"                                                 '字体名称
        Write #5, 2 * Height1                                                   '文字高度
        Print #5, "1"                                                           '宽度比例
        Print #5, "0"                                                           '倾斜角度
        Print #5, "n"                                                           '反向显示文字=否
        Print #5, "n"                                                           '颠倒显示文字=否
        
        '标注设计高程文字
        Print #5, "text"                                                        '输入单行文字text
        Print #5, "j"                                                           '修改文字格式
        Print #5, "ml"                                                          '左中对齐
        Write #5, Kstart - offset - 12 * Height1 + length2, -3 * Height1        '文字起点位置
        Print #5, "0"
        Print #5, "设计高程"                                                    '标注完毕
        Print #5, "text"                                                        '输入单行文字text
        Print #5, "j"                                                           '修改文字格式
        Print #5, "ml"                                                          '左中对齐
        Write #5, Kstart - offset - 12 * Height1 + length2, -9 * Height1        '文字起点位置
        Print #5, "0"
        Print #5, "桩    号"                                                    '标注完毕
        
        '设置当前图层为0层
        Print #5, "-layer"                                                      '设置图层
        Print #5, "s"                                                           '设置当前图层
        Print #5, "0"                                                           '设置当前图层为0层
        Print #5, ""                                                            '结束-layer命令
        Print #5, "zoom"
        Print #5, "e"
        
        Close #5
        msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
        If MsgBox(msg, vbYesNo) = vbYes Then
            Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbMaximizedFocus '有空格的目录需要加双引号
        End If
    Else                                                                        '如果起算点里程为空或终点里程为空或分段终点里程为空或计算间距为零，就提示重新输入参数
        MsgBox "请输全相关参数,并确保间距不为零且KQD≤Kstart≤Kend≤Ki(max)!", vbCritical, "警告！"
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub Command24_Click()                                                   '导出设计标高及横坡上点的标高按钮
    '    If CDec(Val(Text37)) >= CDec(Val(TextKQD)) And Val(Text38) <= shu(Val(Text24), 0) And Text25 <> "" And CDec(Val(Text39)) <> 0 Then
    If CDec(Val(Text37)) >= shu(1, 0) And Val(Text38) <= shu(Val(Text24), 0) And Text25 <> "" And CDec(Val(Text39)) <> 0 Then
        With CommonDialog1
            .FileName = ""
            .CancelError = False
            .DialogTitle = "导出设计标高及横坡上点的标高到文件..."
            .DefaultExt = "exc"
            .Filter = "坐标结果(*.exc)|*.exc"
            .ShowSave
            If Len(.FileName) = 0 Then Exit Sub
        End With
        Open CommonDialog1.FileName For Output As 7
        Print #7, "K, H, I, T, G"
        For K = CDec(Val(Text37)) To CDec(Val(Text38)) Step CDec(Val(Text39))
            计算设计高程 K, J, y, R, E, F, H
            T = CDec(Val(Text35))
            P = CDec(Val(Text44))
            计算边桩高程 K, T, P, KG, H, i, G
            Write #7, K, T, P, KG, H, i, G
        Next K
        Close #7
        If Val(Text34) > Val(Text37) Then
            MsgBox "从" & Val(Text37) & "到" & Val(Text34) & "之间无横坡参数，只能计算中桩高程！", vbInformation, "提示"
        ElseIf hengpo(Val(Text33), 0) < Val(Text38) Then
            MsgBox "从" & hengpo(Val(Text33), 0) & "到" & Val(Text38) & "之间无横坡参数，只能计算中桩高程！", vbInformation, "提示"
        End If
        msg = "文件保存在" & CommonDialog1.FileName & "，用记事本查看吗？"
        If MsgBox(msg, vbYesNo) = vbYes Then
            Shell "notepad.exe " & """" & CommonDialog1.FileName & """", vbNormalFocus
        End If
    Else                                                                        '如果所求点里程在竖曲线起讫里程之外，就提示重新输入参数
        MsgBox "所求点里程在竖曲线起讫里程之外，请重新输入所求点起点、终点里程及其他数值！", vbCritical, "警告！"
    End If
End Sub
'----------------------------------------------------------------------------------------------------------

Private Sub TextX1_Change()
    方位角
End Sub

Private Sub TextX2_Change()
    方位角
End Sub

Private Sub TextY1_Change()
    方位角
End Sub

Private Sub TextY2_Change()
    方位角
End Sub


Sub 计算交点(Xa, Ya, Xb, Yb, Xc, Yc, Xd, Yd, tt, ss, Xjd, Yjd)
    '计算设计线与地面线的交点
    If ((Xb - Xa) * (Yc - Yd) - (Xc - Xd) * (Yb - Ya)) <> 0 Then
        tt = ((Xc - Xa) * (Yc - Yd) - (Xc - Xd) * (Yc - Ya)) / ((Xb - Xa) * (Yc - Yd) - (Xc - Xd) * (Yb - Ya))
        ss = ((Xb - Xa) * (Yc - Ya) - (Xc - Xa) * (Yb - Ya)) / ((Xb - Xa) * (Yc - Yd) - (Xc - Xd) * (Yb - Ya))
    End If
    Xjd = Round(Xa + tt * (Xb - Xa), 3)
    Yjd = Round(Ya + tt * (Yb - Ya), 3)
End Sub


Private Sub Command64_Click()
    '载入参数并生成导出到Excel
    On Error GoTo Label1
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "从文件载入设计线参数..."
        .DefaultExt = "sjx"
        .Filter = "设计线参数(*.sjx)|*.sjx|地面线参数(*.dmx)|*.dmx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text109 = .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 17
    Line Input #17, shuoming1
    Line Input #17, shuoming2
    Line Input #17, shuoming3
    
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "从文件载入地面线参数..."
        .DefaultExt = "dmx"
        .Filter = "地面线参数(*.dmx)|*.dmx|设计线参数(*.sjx)|*.sjx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        Text109 = Text109 & "  " & .FileTitle
    End With
    Open CommonDialog1.FileName For Input As 18
    Line Input #18, shuoming1
    Line Input #18, shuoming2
    Line Input #18, shuoming3
    
    On Error Resume Next                                                        '如果没有发现excel进程就继续执行下一句：创建excel进程
    Set appexcel = GetObject(, "Excel.application")
    If appexcel Is Nothing Then Set appexcel = CreateObject("excel.application")
    If Err Then                                                                 '如果没有安装excel，就用wps表格打开
        Set appexcel = GetObject(, "et.application")
        If appexcel Is Nothing Then Set appexcel = CreateObject("et.application")
    End If
    appexcel.Visible = True
    Set wbmybook = appexcel.Workbooks.Add
    Set exsheet = wbmybook.Worksheets("sheet1")
    exsheet.Activate                                                            '激活
    On Error GoTo Label1
    appexcel.cells.Select
    appexcel.Selection.HorizontalAlignment = xlCenter
    appexcel.Selection.VerticalAlignment = xlCenter
    appexcel.sheets("sheet1").Name = "面积计算表"
appexcel.Range("A1:I1").merge
    appexcel.Range("A1") = "土石方数量表"
    appexcel.cells(1, 1).Font.Bold = True
    appexcel.cells(1, 1).Font.Size = 16
    appexcel.Range("A2") = "桩号"
    appexcel.Range("B2") = "设计" & vbCrLf & "高程"
    appexcel.Range("C2") = "中桩" & vbCrLf & "高程"
    appexcel.Range("D2") = "高差"
    appexcel.Range("E2") = "横断面积" & vbCrLf & "(平方米)"
    appexcel.Range("F2") = "平均面积" & vbCrLf & "(平方米)"
    appexcel.Range("G2") = "距离" & vbCrLf & "(米)"
    appexcel.Range("H2") = "体积" & vbCrLf & "(方)"
    appexcel.Range("I2") = "备注"
appexcel.Columns("A:A").columnwidth = 9
appexcel.Columns("B:D").columnwidth = 7
appexcel.Columns("E:F").columnwidth = 8.5
appexcel.Columns("G:G").columnwidth = 5
appexcel.Columns("H:H").columnwidth = 9
appexcel.Columns("I:I").columnwidth = 9
    jj = 3
    
    Dim str1 As String
    Dim str2 As String
    Dim sjx() As String
    Dim dmx() As String
    Line Input #17, str2                                                        '首先读入设计线的第一个里程文件准备比较
    sjx = Split(str2, ",")
    Do While EOF(18) = False                                                    '读完地面线文件
        Line Input #18, str1
        dmx = Split(str1, ",")
        If UBound(dmx) = -1 Then Exit Do                                        '如果str1是空值，那么ubound(dmx)=-1,直接引用dmx(0)会造成数组下标越界，应该直接退出循环
        
        Do While Not EOF(17) And UBound(sjx) >= 0                               '读设计线到结束且不为空值
            If Val(sjx(0)) < Val(dmx(0)) Then                                   '如果设计线里程小就继续读设计线
                Line Input #17, str2
                sjx = Split(str2, ",")
            Else
                Exit Do                                                         '如果设计里程大就退出循环
            End If
        Loop
        
        If UBound(sjx) >= 0 And UBound(dmx) >= 0 Then
            If CDec(Val(sjx(0))) = CDec(Val(dmx(0))) Then                       '如果在设计线中找到对应的桩号就计算
                '求左幅坡脚线
                For I1 = 2 To UBound(sjx) - 3 Step 2
                    If UBound(sjx) >= 4 Then
                        If UBound(sjx) Mod 2 = 1 Then                           '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                            Xa = Val(sjx(I1))
                            Ya = Val(sjx(1)) + Val(sjx(I1 + 1))
                            Xb = Val(sjx(I1 + 2))
                            Yb = Val(sjx(1)) + Val(sjx(I1 + 3))
                        Else                                                    '如果是里程1,平距1，高程1，平距2，高程2格式
                            Xa = Val(sjx(I1 - 1))
                            Ya = Val(sjx(I1))
                            Xb = Val(sjx(I1 + 1))
                            Yb = Val(sjx(I1 + 2))
                        End If
                        Xc = Val(dmx(1))
                        Yc = Val(dmx(2))
                        If UBound(dmx) >= 4 Then
                            Xd = Val(dmx(3))
                            Yd = Val(dmx(4))
                        Else
                            Xd = Xc - 13
                            Yd = Yc
                        End If
                        Call 计算交点(Xa, Ya, Xb, Yb, Xc, Yc, Xd, Yd, tt1, ss1, Xjd1, Yjd1) '计算左幅交点
                        If tt1 >= 0 And tt1 < 1 Then                            '要确保设计线足够长才有交点
                            Exit For
                        End If
                    End If
                Next
                
                '求右幅坡脚线
                For I1 = UBound(sjx) - 1 To 2 Step -2
                    If UBound(sjx) >= 4 Then
                        If UBound(sjx) Mod 2 = 1 Then                           '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                            Xa = Val(sjx(I1 - 2))
                            Ya = Val(sjx(1)) + Val(sjx(I1 - 1))
                            Xb = Val(sjx(I1))
                            Yb = Val(sjx(1)) + Val(sjx(I1 + 1))
                        Else                                                    '如果是里程1,平距1，高程1，平距2，高程2格式
                            Xa = Val(sjx(I1 - 2))
                            Ya = Val(sjx(I1 - 1))
                            Xb = Val(sjx(I1))
                            Yb = Val(sjx(I1 + 1))
                        End If
                        i2 = UBound(dmx)
                        If UBound(dmx) >= 4 Then
                            Xc = Val(dmx(i2 - 3))
                            Yc = Val(dmx(i2 - 2))
                        Else
                            Xc = Val(dmx(1)) + 13
                            Yc = Val(dmx(2))
                        End If
                        Xd = Val(dmx(i2 - 1))
                        Yd = Val(dmx(i2))
                        Call 计算交点(Xa, Ya, Xb, Yb, Xc, Yc, Xd, Yd, tt1, ss1, xjd2, Yjd2) '计算右幅交点
                        If tt1 >= 0 And tt1 < 1 Then                            '要确保设计线足够长才有交点
                            Exit For
                        End If
                    End If
                Next
                
                If UBound(dmx) = 2 Then hd = Val(dmx(2))
                For I1 = 2 To UBound(dmx) - 2 Step 2                            '求中桩地面高程
                    If dmx(I1 - 1) = 0 Or dmx(I1 - 1) < 0 And dmx(I1 + 1) > 0 Then
                        hd = dmx(I1) + Abs(Val(dmx(I1 - 1))) * (dmx(I1 + 2) - dmx(I1)) / (dmx(I1 + 1) + Abs(Val(dmx(I1 - 1))))
                        Exit For
                    End If
                Next
                
                '计算填方段或挖方段面积
                smax = UBound(sjx)
                dmax = UBound(dmx)
                Dim aa(200)                                                     '先把多边形顶点xy存放到aa中，有I1个顶点
                aa(0) = Xjd1
                aa(1) = Yjd1
                For I1 = 4 To UBound(sjx) - 2 Step 2
                    If UBound(sjx) Mod 2 = 1 Then                               '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                        aa(I1 - 2) = Val(sjx(I1))
                        aa(I1 - 1) = Val(sjx(1)) + Val(sjx(I1 + 1))
                    Else                                                        '如果是里程1,平距1，高程1，平距2，高程2格式
                        aa(I1 - 2) = Val(sjx(I1 - 1))
                        aa(I1 - 1) = Val(sjx(I1))
                    End If
                Next
                aa(I1 - 2) = xjd2
                aa(I1 - 1) = Yjd2
                For i2 = UBound(dmx) - 1 To 1 Step -2
                    If Val(dmx(i2)) > Xjd1 And Val(dmx(i2)) < xjd2 Then
                        aa(I1) = Val(dmx(i2))
                        aa(I1 + 1) = Val(dmx(i2 + 1))
                        I1 = I1 + 2
                    End If
                Next
                aa(I1) = Xjd1
                aa(I1 + 1) = Yjd1
                '已知坐标求多边形面积：不能计算半填半挖段。
                ss = 0
                For i2 = 0 To I1 - 2 Step 2
                    ss = ss + aa(i2) * aa(i2 + 3) - aa(i2 + 1) * aa(i2 + 2)
                Next
                ss = Round(Abs(ss) / 2, 3)
                appexcel.cells(jj, 1) = Val(dmx(0))                             '里程
                '添加设计高程，中桩高程，高差
                If UBound(sjx) Mod 2 = 1 Then                                   '如果是里程1,设计高程1，平距1，高差1，平距2，高差2格式
                    appexcel.cells(jj, 2) = Val(sjx(1))                         '设计高程
                Else                                                            '如果是里程1,平距1，高程1，平距2，高程2格式
                    For i4 = 1 To UBound(sjx) - 2
                        If Val(sjx(i4)) <= 0 And Val(sjx(i4 + 2)) > 0 Then
                            appexcel.cells(jj, 2) = Val(sjx(i4 + 1))            '设计高程
                            Exit For
                        End If
                    Next
                End If
                appexcel.cells(jj, 3) = hd                                      '中桩高程
                '第4列存储高差
                appexcel.cells(jj, 5) = ss                                      '断面面积
                jj = jj + 1
            End If
        End If
    Loop
    Close (17)
    Close (18)
    Close (19)
    
    'For I3 = 4 To jj - 1
    'appexcel.cells(I3, 3) = "=1*round((B" & I3 - 1 & "+B" & I3 & ")/ 2,3)"
    'appexcel.cells(I3, 4) = "=1*round((A" & I3 & "-A" & I3 - 1 & "),3)"
    'appexcel.cells(I3, 5) = "=1*round(C" & I3 & "*D" & I3 & ",3)"
    'Next
    
    '填写平均面积、距离、体积计算公式
    appexcel.cells(3, 4) = "=B3-C3"
    appexcel.cells(4, 6) = "=1*round((E3+E4)/2,3)"
    appexcel.cells(4, 7) = "=1*ROUND((A4-A3),3)"
    appexcel.cells(4, 8) = "=1*ROUND(F4*G4,3)"
appexcel.Range("D3:D3").AutoFill Destination:=appexcel.Range("D3:D" & jj - 1), Type:=0
appexcel.Range("F4:H4").AutoFill Destination:=appexcel.Range("F4:H" & jj - 1), Type:=0
    
    
    appexcel.Range(appexcel.cells(3, 1), appexcel.cells(jj - 1, 1)).NumberFormatLocal = "!" & Text111.text
    appexcel.cells(jj, 1) = "合计"
    appexcel.cells(jj, 7) = "=sum(G4:G" & jj - 1 & ")"
    appexcel.cells(jj, 8) = "=sum(H4:H" & jj - 1 & ")"
    
    
    '以下设置边框：
    appexcel.Range(appexcel.cells(2, 1), appexcel.cells(jj, 9)).Select
    appexcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    appexcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With appexcel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    With appexcel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlthin
        .ColorIndex = xlAutomatic
    End With
    
    appexcel.cells(2, 10).Select
Label1: Close
End Sub


Private Sub Command62_Click()
    '编辑设计线参数
    Dim file1 As String
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "编辑现有的设计线文件..."
        .DefaultExt = "sjx"
        .Filter = "设计线参数(*.sjx)|*.sjx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        file1 = .FileName
    End With
    If Dir(file1) = "" Then
        '如果文件不存在就先把三行说明行写入文件
        Open CommonDialog1.FileName For Output As 1
        Print #1, "说明：前3行是说明行不能删除，第4行之后是数据行。平距左负右正。"
        Print #1, "里程1,设计高程1,平距1,高差1,平距2,高差2……,以英文逗号隔开，回车后输入下一行。"
        Print #1, "里程2,设计高程1,平距1,高差1,平距2,高差2……"
        Close #1
    End If
    '如果文件存在就直接打开：
    Shell "notepad.exe " & """" & file1 & """", vbMaximizedFocus
End Sub

Private Sub Command63_Click()
    '编辑地面线
    Dim file1 As String
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "编辑现有的地面线文件..."
        .DefaultExt = "dmx"
        .Filter = "原地面参数(*.dmx)|*.dmx"
        .ShowOpen
        If Len(.FileName) = 0 Then: Close: Exit Sub
        file1 = .FileName
    End With
    If Dir(file1) = "" Then
        '如果文件不存在就先把三行说明行写入文件
        Open CommonDialog1.FileName For Output As 1
        Print #1, "说明：前3行是说明行不能删除，第4行之后是数据行。平距左负右正。"
        Print #1, "里程1,平距1,高程1,平距2,高程2……,以英文逗号隔开，回车后输入下一行。"
        Print #1, "里程2,平距1,高程1,平距2,高程2……"
        Close #1
    End If
    '如果文件存在就直接打开：
    Shell "notepad.exe " & """" & file1 & """", vbMaximizedFocus
End Sub
