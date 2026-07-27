VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "已知坐标反算里程偏距"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   3210
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   "批量 反算"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "反算"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "反算偏距："
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "反算里程："
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "大致里程K："
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "坐标Y："
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "坐标X："
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim factor(1000, 9) As Variant  '在通用部分中申明平曲线参数数组：类型，Ki，J，C，D，F，R，A，B，O
Const pai = 3.14159265358979

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Command2_Click()
    mm = Val(Text1)
    nn = Val(Text2)
    kk = Val(Text3)
Label1:
    计算 kk, x, y, Q
    '度分秒化为弧度 Q, Qh
    x1 = x
    y1 = y
    x2 = mm
    y2 = nn
    ii = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
    If x2 <> x1 And y2 <> y1 Then
        jj = Atn((y2 - y1) / (x2 - x1)) + (1 - Abs(x2 - x1) / (x2 - x1)) * pai / 2 + (1 + Abs(x2 - x1) / (x2 - x1)) * (1 - Abs(y2 - y1) / (y2 - y1)) * pai / 2
    ElseIf x2 = x1 Then                                                         '如果X2=X1 会被零除，要单独计算
        If y2 > y1 Then jj = pai / 2
        If y2 < y1 Then jj = pai / 2 * 3
    ElseIf y2 = y1 Then                                                         '如果Y2=Y1 会被零除，要单独计算
        If x2 > x1 Then jj = 0
        If x2 < x1 Then jj = pai
    End If
    If Abs(ii * Cos(jj - Q)) > 0.001 Then
        kk = kk + ii * 0.382 * Cos(jj - Q)
        '        Call MsgBox("kk=" & kk & "ss=" & ii * 0.382 * Cos(jj - Q))
        GoTo Label1
    Else
        xiaoshu = Val(Form1.Text114)
        Text4 = Round(kk, xiaoshu)
        If jj - Q <> 0 Then Text5 = Round(ii * Sin(jj - Q) / Abs(Sin(jj - Q)), xiaoshu)
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 计算(K, x, y, Q)
    求区间 K, number                                                            '计算出K落在平曲线哪个区间上，保存在变量Number中
    If factor(number, 0) = "直线" Then
        直线 K, x, y, Q
    ElseIf factor(number, 0) = "圆" Then
        圆 K, x, y, Q
    ElseIf factor(number, 0) = "缓1+圆+缓2" Then
        平曲线 K, x, y, Q
    ElseIf factor(number, 0) = "布劳斯曲线" Then
        布劳斯曲线 K, x, y, Q
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 求区间(K, number)                                                           '计算出K落在平曲线哪个区间上，保存在变量Number中
    number = 1                                                                  '先赋个初值
    If factor(1, 1) < factor(2, 1) Then
        Do While Val(K) > Val(factor(number, 1)) And number < Val(Form1.TextPartNumber)
            number = number + 1
        Loop
    End If
    If factor(1, 1) > factor(2, 1) Then
        Do While Val(K) < Val(factor(number, 1)) And number < Val(Form1.TextPartNumber)
            number = number + 1
        Loop
    End If
    'If Combo1.ListCount > 0 Then Combo1.ListIndex = number - 1
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 直线(K, x, y, Q)
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
    'CEYI x, Y, Q, T, P, S, Z, G, H
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 圆(K, x, y, Q)
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
        'CEYI x, Y, Q, T, P, S, Z, G, H
    Else                                                                        'R=0就提示重新输入参数R
        MsgBox ("第" & number & "分段的R不能为零,请修改参数并按输入按钮。"), vbOKOnly, "提示"
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Sub 平曲线(K, x, y, Q)
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
        缓 K, x, y, Q
    ElseIf CDec(Val(K)) <= CDec(Val(J)) + CDec(Val(O)) Then
        圆 K, x, y, Q
    Else
        缓 K, x, y, Q
    End If
End Sub
Sub 布劳斯曲线(K, x, y, Q)
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
        'U = L - L ^ 7 / (14 * R ^ 2 * a ^ 4) + L ^ 8 / (16 * R ^ 2 * a ^ 5) - L ^ 9 / (72 * R ^ 2 * a ^ 6) + L ^ 13 / (312 * R ^ 4 * a ^ 8) - L ^ 14 / (168 * R ^ 4 * a ^ 9) + L ^ 15 / (240 * R ^ 4 * a ^ 10) - L ^ 16 / (768 * R ^ 4 * a ^ 11) + L ^ 17 / (6528 * R ^ 4 * a ^ 12)
        'V = L ^ 4 / (4 * R * a ^ 2) - L ^ 5 / (10 * R * a ^ 3) - L ^ 10 / (60 * R ^ 3 * a ^ 6) + L ^ 11 / (44 * R ^ 3 * a ^ 7) - L ^ 12 / (96 * R ^ 3 * a ^ 8) + L ^ 13 / (624 * R ^ 3 * a ^ 9) + L ^ 16 / (1920 * R ^ 5 * a ^ 10) - L ^ 17 / (816 * R ^ 5 * a ^ 11) + L ^ 18 / (864 * R ^ 5 * a ^ 12) - L ^ 19 / (1824 * R ^ 5 * a ^ 13) + L ^ 20 / (7680 * R ^ 5 * a ^ 14) - L ^ 21 / (80640 * R ^ 5 * a ^ 15)
        
        '辛普森法求定积分近似值
        L1 = 0
        L2 = Abs(K - J)
        Ls = CDec(Val(factor(number, 7)))
        Call simpson(L1, L2, R, Ls, U, V)
        
        x = C + W * U * Cos(E) - i * V * Sin(E)
        y = D + W * U * Sin(E) + i * V * Cos(E)
        Q = E + (L ^ 3 / (R * A ^ 2) - L ^ 4 / (2 * R * A ^ 3)) * W * i         '单位弧度
        'CEYI x, Y, Q, T, P, S, Z, G, H
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
    '布劳斯曲线
    If R * Ls <> 0 Then Fun = (L ^ 3 / (R * Ls * Ls) - L ^ 4 / (2 * R * Ls ^ 3))
End Function

'----------------------------------------------------------------------------------------------------------
Sub 缓(K, x, y, Q)
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
        U = (L - H) - (L ^ 5 - H ^ 5) / (40 * G ^ 4) + (L ^ 9 - H ^ 9) / (3456 * G ^ 8) _
        - (L ^ 13 - H ^ 13) / (599040 * G ^ 12) + (L ^ 17 - H ^ 17) / (175472640 * G ^ 16)
        
        V = (L ^ 3 - H ^ 3) / (6 * G ^ 2) - (L ^ 7 - H ^ 7) / (336 * G ^ 6) + (L ^ 11 - H ^ 11) / (42240 * G ^ 10) _
        - (L ^ 15 - H ^ 15) / (9676800 * G ^ 14) + (L ^ 19 - H ^ 19) / (3158507520.001 * G ^ 18)
        
        '辛普森法求定积分近似值
        'L1 = H
        'L2 = L
        'Call simpson2(L1, L2, G, U, V)
        
        x = C + R * (1 - W) * Sin(O / (2 * R)) * Cos(F1 + O / (2 * R)) + W * U * Cos(E) - i * V * Sin(E) '中桩的X坐标
        y = D + R * (1 - W) * Sin(O / (2 * R)) * Sin(F1 + O / (2 * R)) + W * U * Sin(E) + i * V * Cos(E) '中桩的Y坐标
        Q = E + L ^ 2 * W * i / (2 * G ^ 2)                                     '切线方位角，单位弧度
        'CEYI x, Y, Q, T, P, S, Z, G, H
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
    '缓和曲线
    If G <> 0 Then Fun2 = (L * L) / (2 * G * G)
End Function

'----------------------------------------------------------------------------------------------------------
Sub CEYI(x, y, Q, T, P, S, Z, G, H)                                             '计算边桩的坐标
    If Val(Q) < 0 Then                                                          '转化到[0,2pai)之间
        Q = Q + 2 * 3.14159265358979
    ElseIf Val(Q) >= 2 * 3.14159265358979 Then
        Q = Q - 2 * 3.14159265358979
    End If
    'T = CDec(Val(TEXTT))
    'P = CDec(Val(TEXTP))
    'S = CDec(Val(TEXTS))
    'Z = CDec(Val(TEXTZ))
    'If anniu = 1 Then
    'T = CDec(Val(factor4(ndh, 2)))
    'P = CDec(Val(factor4(ndh, 3)))
    'S = CDec(Val(factor4(ndh, 4)))
    'Z = CDec(Val(factor4(ndh, 5)))
    'End If
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

Private Sub Command3_Click()
    Dim dianhao As String
    With CommonDialog1
        .FileName = ""                                                          '先清空文件名
        .CancelError = False
        .DialogTitle = "载入需要反算的excel坐标..."
        .DefaultExt = "xls"
        .Filter = "输入的曲线参数(*.xls)|*.xls"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
    End With
    Dim xlapp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    On Error Resume Next
    Set xlapp = GetObject(CommonDialog1.FileName).application
    Set xlBook = xlapp.Workbooks.Open(CommonDialog1.FileName)
    Set xlSheet = xlBook.activesheet
    xlapp.Visible = True
    rw = 3
Label2:
    mm = xlSheet.cells(rw, 1)
    nn = xlSheet.cells(rw, 2)
    kk = xlSheet.cells(rw, 3)                                                   '第三列输入大致里程
Label1:
    计算 kk, x, y, Q
    '度分秒化为弧度 Q, Qh
    x1 = x
    y1 = y
    x2 = mm
    y2 = nn
    ii = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
    If x2 <> x1 And y2 <> y1 Then
        jj = Atn((y2 - y1) / (x2 - x1)) + (1 - Abs(x2 - x1) / (x2 - x1)) * pai / 2 + (1 + Abs(x2 - x1) / (x2 - x1)) * (1 - Abs(y2 - y1) / (y2 - y1)) * pai / 2
    ElseIf x2 = x1 Then                                                         '如果X2=X1 会被零除，要单独计算
        If y2 > y1 Then jj = pai / 2
        If y2 < y1 Then jj = pai / 2 * 3
    ElseIf y2 = y1 Then                                                         '如果Y2=Y1 会被零除，要单独计算
        If x2 > x1 Then jj = 0
        If x2 < x1 Then jj = pai
    End If
    If Abs(ii * Cos(jj - Q)) > 0.001 Then
        kk = kk + ii * 0.382 * Cos(jj - Q)
        GoTo Label1
    Else
        xiaoshu = Val(Form1.Text114)
        xlapp.cells(rw, 4) = Round(kk, xiaoshu)
        If jj - Q <> 0 Then xlapp.cells(rw, 5) = Round(ii * Sin(jj - Q) / Abs(Sin(jj - Q)), xiaoshu)
    End If
    rw = rw + xlSheet.cells(rw, 1).MergeArea.Rows.Count                         '单元格的合并区域的行计数
    If xlSheet.cells(rw, 1) <> 0 And xlSheet.cells(rw, 2) <> 0 Then GoTo Label2
    Set xlapp = Nothing
End Sub
