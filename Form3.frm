VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form3 
   Caption         =   "从圆曲线推算HY点参数"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   6435
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      Height          =   375
      Left            =   5760
      TabIndex        =   31
      Top             =   3240
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   81000
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text21 
      Height          =   300
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "注意A不能取零"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "输入到主程序"
      Height          =   375
      Left            =   4320
      TabIndex        =   30
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "推算HY点"
      Height          =   375
      Left            =   3240
      TabIndex        =   29
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存"
      Height          =   375
      Left            =   2400
      TabIndex        =   28
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "载入"
      Height          =   375
      Left            =   1560
      TabIndex        =   27
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Text20 
      Height          =   300
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "J"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text19 
      Height          =   300
      Left            =   1605
      TabIndex        =   23
      ToolTipText     =   "C"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      Height          =   300
      Left            =   3090
      TabIndex        =   24
      ToolTipText     =   "D"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text17 
      Height          =   300
      Left            =   4590
      TabIndex        =   25
      ToolTipText     =   "必须为度.分秒的格式"
      Top             =   2760
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4260
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "一点法PTR"
      TabPicture(0)   =   "Form3.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TEXTJ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TEXTC"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TEXTD"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TEXTF"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TEXTA"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LabelA"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "两点法PPR"
      TabPicture(1)   =   "Form3.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(9)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(10)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(11)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(23)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(26)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Text3"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text4"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Text6"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text7"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Text8"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text1"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "三点法PPP"
      TabPicture(2)   =   "Form3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(4)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(12)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(13)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(14)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label1(15)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label1(16)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label1(17)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label1(18)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label1(24)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label1(25)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label1(27)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Text5"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Text9"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Text10"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Text11"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Text12"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Text13"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Text14"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Text15"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Text16"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Text22"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Text23"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).ControlCount=   23
      Begin VB.TextBox Text23 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   -70200
         TabIndex        =   32
         ToolTipText     =   "D"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text22 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   -70200
         TabIndex        =   33
         ToolTipText     =   "D"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Height          =   300
         Left            =   -74760
         TabIndex        =   19
         ToolTipText     =   "J"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         Height          =   300
         Left            =   -73275
         TabIndex        =   20
         ToolTipText     =   "C"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text14 
         Height          =   300
         Left            =   -71790
         TabIndex        =   21
         ToolTipText     =   "D"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text13 
         Height          =   300
         Left            =   -74760
         TabIndex        =   16
         ToolTipText     =   "J"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Height          =   300
         Left            =   -73275
         TabIndex        =   17
         ToolTipText     =   "C"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   300
         Left            =   -71790
         TabIndex        =   18
         ToolTipText     =   "D"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   300
         Left            =   -74760
         TabIndex        =   13
         ToolTipText     =   "J"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   300
         Left            =   -73275
         TabIndex        =   14
         ToolTipText     =   "C"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   300
         Left            =   -71790
         TabIndex        =   15
         ToolTipText     =   "D"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "注意A不能取零"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "J"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   300
         Left            =   1605
         TabIndex        =   10
         ToolTipText     =   "C"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   300
         Left            =   3120
         TabIndex        =   11
         ToolTipText     =   "D"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "J"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   1605
         TabIndex        =   7
         ToolTipText     =   "C"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   3090
         TabIndex        =   8
         ToolTipText     =   "D"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TEXTJ 
         Height          =   300
         Left            =   -74880
         TabIndex        =   1
         ToolTipText     =   "J"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TEXTC 
         Height          =   300
         Left            =   -73395
         TabIndex        =   2
         ToolTipText     =   "C"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TEXTD 
         Height          =   300
         Left            =   -71910
         TabIndex        =   3
         ToolTipText     =   "D"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TEXTF 
         Height          =   300
         Left            =   -74880
         TabIndex        =   4
         ToolTipText     =   "必须为度.分秒的格式"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TEXTA 
         Height          =   300
         Left            =   -73440
         TabIndex        =   5
         ToolTipText     =   "注意A不能取零"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "圆心Y"
         Height          =   255
         Index           =   27
         Left            =   -69840
         TabIndex        =   64
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "圆心X"
         Height          =   255
         Index           =   25
         Left            =   -69840
         TabIndex        =   63
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "HY点里程手工输入。"
         Height          =   255
         Index           =   26
         Left            =   1680
         TabIndex        =   62
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "请确保J1<J2<J3"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   24
         Left            =   -70200
         TabIndex        =   61
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "请确保J1<J2"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   23
         Left            =   4680
         TabIndex        =   60
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "P3点里程J3"
         Height          =   255
         Index           =   18
         Left            =   -74520
         TabIndex        =   54
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "P3点X坐标C3"
         Height          =   255
         Index           =   17
         Left            =   -73200
         TabIndex        =   53
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "P3点Y坐标D3"
         Height          =   255
         Index           =   16
         Left            =   -71640
         TabIndex        =   52
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "P2点里程J2"
         Height          =   255
         Index           =   15
         Left            =   -74520
         TabIndex        =   51
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "P2点X坐标C2"
         Height          =   255
         Index           =   14
         Left            =   -73200
         TabIndex        =   50
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "P2点Y坐标D2"
         Height          =   255
         Index           =   13
         Left            =   -71640
         TabIndex        =   49
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "P1点里程J1"
         Height          =   255
         Index           =   12
         Left            =   -74520
         TabIndex        =   48
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "P1点X坐标C1"
         Height          =   255
         Index           =   4
         Left            =   -73200
         TabIndex        =   47
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "P1点Y坐标D1"
         Height          =   255
         Index           =   0
         Left            =   -71640
         TabIndex        =   46
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "圆半径R"
         Height          =   255
         Left            =   195
         TabIndex        =   45
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "P2点里程J2"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   44
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "P2点X坐标C2"
         Height          =   255
         Index           =   10
         Left            =   1680
         TabIndex        =   43
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "P2点Y坐标D2"
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   42
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "P1点里程J1"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   41
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "P1点X坐标C1"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   40
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "P1点Y坐标D1"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   39
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "P1点里程J1"
         Height          =   255
         Index           =   5
         Left            =   -74640
         TabIndex        =   38
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "P1点X坐标C1"
         Height          =   255
         Index           =   6
         Left            =   -73320
         TabIndex        =   37
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "P1点Y坐标D1"
         Height          =   255
         Index           =   7
         Left            =   -71760
         TabIndex        =   36
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "P1点切线方位角T1"
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   35
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label LabelA 
         Caption         =   "圆半径R"
         Height          =   255
         Left            =   -73200
         TabIndex        =   34
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      Caption         =   "圆半径R"
      Height          =   255
      Left            =   360
      TabIndex        =   59
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "输入HY点里程"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   22
      Left            =   240
      TabIndex        =   58
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "HY点X坐标"
      Height          =   255
      Index           =   21
      Left            =   1680
      TabIndex        =   57
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "HY点Y坐标"
      Height          =   255
      Index           =   20
      Left            =   3240
      TabIndex        =   56
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "HY点切线方位角F"
      Height          =   255
      Index           =   19
      Left            =   4560
      TabIndex        =   55
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'定义一下常量
Const pai = 3.14159265358979
Private Sub Command1_Click()
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "载入已知点参数到文件..."
        .DefaultExt = "yzd"
        .Filter = "载入点参数(*.yzd)|*.yzd"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
    End With
    Open CommonDialog1.FileName For Input As 2
    Input #2, leixing
    If leixing = "PTR" Then
        Input #2, p1, c1, d1, T1, R, Khy
        SSTab1.Tab = 0
        TEXTJ = p1
        TEXTC = c1
        TEXTD = d1
        TEXTF = T1
        TEXTA = R
        Text20 = Khy
    ElseIf leixing = "PPR" Then
        Input #2, p1, c1, d1, p2, c2, d2, R, Khy
        SSTab1.Tab = 1
        Text4 = p1
        Text3 = c1
        Text2 = d1
        Text8 = p2
        Text7 = c2
        Text6 = d2
        Text1 = R
        Text20 = Khy
    Else
        Input #2, p1, c1, d1, p2, c2, d2, p3, c3, d3, Khy
        SSTab1.Tab = 2
        Text10 = p1
        Text9 = c1
        Text5 = d1
        Text13 = p2
        Text12 = c2
        Text11 = d2
        Text16 = p3
        Text15 = c3
        Text14 = d3
        Text20 = Khy
    End If
    Close #2
End Sub

Private Sub Command2_Click()
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "保存已知点参数到文件..."
        .DefaultExt = "yzd"
        .Filter = "已知点参数(*.yzd)|*.yzd"
        .ShowSave
        If Len(.FileName) = 0 Then: Close: Exit Sub
    End With
    Open CommonDialog1.FileName For Output As 2
    If SSTab1.Tab = 0 Then
        Write #2, "PTR"
        Write #2, Val(TEXTJ), Val(TEXTC), Val(TEXTD), Val(TEXTF), Val(TEXTA), Val(Text20)
    ElseIf SSTab1.Tab = 1 Then
        Write #2, "PPR"
        Write #2, Val(Text4), Val(Text3), Val(Text2), Val(Text8), Val(Text7), Val(Text6), Val(Text1), Val(Text20)
    Else
        Write #2, "PPP"
        Write #2, Val(Text10), Val(Text9), Val(Text5), Val(Text13), Val(Text12), Val(Text11), Val(Text16), Val(Text15), Val(Text14), Val(Text20)
    End If
    Close #2
End Sub

Private Sub Command3_Click()
    If SSTab1.Tab = 0 Then                                                      'PTR一点法
        p1 = Val(TEXTJ)
        c1 = Val(TEXTC)
        d1 = Val(TEXTD)
        T1 = Val(TEXTF)
        R = Val(TEXTA)
        Khy = Val(Text20)
        J = p1
        C = c1
        D = d1
        F = T1
        R = R
        If R <> 0 Then
            L = Khy - J
            度分秒化为弧度 F, F1                                                'F----度.分秒--弧度
            W = F1 + L / (2 * R)                                                'W----弧度
            i = 2 * R * Sin(L / (2 * R))                                        'I为弦长
            x = C + i * Cos(W)
            y = D + i * Sin(W)
            Q = F1 + L / R
            弧度化为度分秒 Q
        End If
        Text19 = Round(x, 4)
        Text18 = Round(y, 4)
        Text17 = Round(Q, 5)
        Text21 = R
    ElseIf SSTab1.Tab = 1 Then                                                  'PPR两点法
        p1 = Val(Text4)
        c1 = Val(Text3)
        d1 = Val(Text2)
        p2 = Val(Text8)
        c2 = Val(Text7)
        d2 = Val(Text6)
        R = Val(Text1)
        Khy = Val(Text20)
        x1 = c1
        y1 = d1
        x2 = c2
        y2 = d2
        计算方位角 x1, y1, x2, y2, ss, Qh
        J = p1
        C = c1
        D = d1
        If R <> 0 Then F1 = Qh - (p2 - p1) / (2 * R)
        R = R
        If R <> 0 Then
            L = Khy - J
            '度分秒化为弧度 F, F1                                                'F----度.分秒--弧度
            W = F1 + L / (2 * R)                                                'W----弧度
            i = 2 * R * Sin(L / (2 * R))                                        'I为弦长
            x = C + i * Cos(W)
            y = D + i * Sin(W)
            Q = F1 + L / R
            弧度化为度分秒 Q
        End If
        Text19 = Round(x, 4)
        Text18 = Round(y, 4)
        Text17 = Round(Q, 5)
        Text21 = R
    Else                                                                        'PPP三点法
        p1 = Val(Text10)
        c1 = Val(Text9)
        d1 = Val(Text5)
        p2 = Val(Text13)
        c2 = Val(Text12)
        d2 = Val(Text11)
        p3 = Val(Text16)
        c3 = Val(Text15)
        d3 = Val(Text14)
        Khy = Val(Text20)
        x1 = c1
        y1 = d1
        x2 = c2
        y2 = d2
        X3 = c3
        Y3 = d3
        x0 = ((Y3 - y1) * (y2 * y2 - y1 * y1 + x2 * x2 - x1 * x1) + (y2 - y1) * (y1 * y1 - Y3 * Y3 + x1 * x1 - X3 * X3)) / (2 * (x2 - x1) * (Y3 - y1) - 2 * (X3 - x1) * (y2 - y1))
        y0 = ((X3 - x1) * (x2 * x2 - x1 * x1 + y2 * y2 - y1 * y1) + (x2 - x1) * (x1 * x1 - X3 * X3 + y1 * y1 - Y3 * Y3)) / (2 * (y2 - y1) * (X3 - x1) - 2 * (Y3 - y1) * (x2 - x1))
        R = Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
        Text23 = Round(x0, 4)
        Text22 = Round(y0, 4)
        计算方位角 x1, y1, x2, y2, ss1, Qh1
        计算方位角 x2, y2, X3, Y3, ss2, qh2
        If Sin(qh2 - Qh1) > 0 Then                                              '利用正弦值判断左右偏
            R = R
        Else
            R = -R
        End If
        '以下转化为两点法计算坐标
        J = p1
        C = c1
        D = d1
        If R <> 0 Then F1 = Qh1 - (p2 - p1) / (2 * R)
        R = R
        If R <> 0 Then
            L = Khy - J
            '度分秒化为弧度 F, F1                                                'F----度.分秒--弧度
            W = F1 + L / (2 * R)                                                'W----弧度
            i = 2 * R * Sin(L / (2 * R))                                        'I为弦长
            x = C + i * Cos(W)
            y = D + i * Sin(W)
            Q = F1 + L / R
            弧度化为度分秒 Q
        End If
        Text19 = Round(x, 4)
        Text18 = Round(y, 4)
        Text17 = Round(Q, 5)
        Text21 = Round(R, 3)
    End If
End Sub

Private Sub Command4_Click()
    Form1.TEXTJ = Text20
    Form1.TEXTC = Text19
    Form1.TEXTD = Text18
    Form1.TEXTF = Text17
    Form1.TEXTR = Text21
End Sub

Private Sub Command5_Click()
    Me.Hide
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

