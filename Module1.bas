Attribute VB_Name = "Module1"
Option Explicit
Public factor(1000, 9) As Variant                                               '在通用部分中申明平曲线参数数组：类型，Ki，J，C，D，F，R，A，B，O
'该模块传递一个公共数组，平曲线参数的数组。
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2

'Declare Function GetDesktopWindow Lib "user32" () As Long
'Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'Public Const GW_CHILD = 5
'Public Const GW_HWNDNEXT = 2

'Public SY As String
'Public bt As String
'Public 句柄 As Long
'Public 标题 As String
'Public 类名 As String
'Public hm As Integer
Public ckbt As String * 50                                                      '窗口标题
Public cklm As String * 50

'Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'public Const SWP_NOSIZE = &H1
'public Const SWP_NOMOVE = &H2
'public Const HWND_TOPMOST = -1
'public Const HWND_NOTOPMOST = -2
'public Const HWND_TOP = 0
'public Const SWP_NOZORDER = &H8
'public Const SWP_SHOWWINDOWS = &H40
'public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'public Const SW_HIDE = 0
'public Const SW_SHOW = 5
'public Const SW_RESTORE = 9
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_MOUSEACTIVATE = &H21
Public Const MK_LBUTTON = &H1
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const SW_SHOWNORMAL = 1
'public Const WM_KEYDOWN = &H100
'public Const VK_NUMPAD0 = &H60
'public Const VK_NUMPAD3 = &H63
'public Const VK_NUMPAD4 = &H64
'public Const VK_F5 = &H74
'public Const VK_F8 = &H77
'public Const VK_F10 = &H79
'public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
'以下几个API用于关闭前台进程输入法：
'-------------------------------------------------------------------------------------------------------
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal ubuflen As Long) As Long
'Public Declare Function ImmAssociateContext Lib "imm32.dll" (ByVal hwnd As Long, ByVal himc As Long) As Long
'Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
'Public Declare Function ImmDisableIME Lib "imm32.dll" (ByVal idThread As Long) As Boolean

'-------------------------------------------------------------------------------------------------------
'以下几个API设置capslock状态
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
'Public Const KEYEVENTF_EXTENDEDKEY = &H1
'Public Const KEYEVENTF_KEYUP = &H2

Public Function 置顶(Title As String) As Long
    '直接调用置顶函数，然后传一个窗口标题进去就可以获得其窗口句柄了.修改后包含title的第一个结果作为查找结果
    Dim tmp As String
    Dim hwnd As Long
    Dim lngProcID As Long
    Dim strTitle As String * 255                                                '//用来存储窗口的标题
    Dim GetHandle2 As Long
    hwnd = GetDesktopWindow()                                                   '//取得桌面窗口
    hwnd = GetWindow(hwnd, GW_CHILD)                                            '//取得桌面窗口的第一个子窗口
    Do While hwnd <> 0                                                          '//通过循环来枚举所有的窗口
        GetWindowText hwnd, strTitle, Len(strTitle)                             '//取得下一个窗口的标题，并写入到列表框中
        If Left$(strTitle, 1) <> vbNullChar Then                                '如果strtitle第一个字符不是空字符就继续
            If InStr(strTitle, Title) > 0 Then
                GetHandle2 = hwnd
                Exit Do
            End If
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)                                     '//调用GetWindow函数，来取得下一个窗口
    Loop
    SetForegroundWindow hwnd
    'BringWindowToTop hwnd
    'SetWindowPos hwnd, 0, 0, 0, 0, 0, 3
End Function

Public Function GetHandle(Title As String) As Long
    '直接调用GetHandle函数，然后传一个窗口标题进去就可以获得其窗口句柄了.
    Dim tmp As String
    Dim hwnd As Long
    Dim lngProcID As Long
    Dim strTitle As String * 255                                                '//用来存储窗口的标题
    hwnd = GetDesktopWindow()                                                   '//取得桌面窗口
    hwnd = GetWindow(hwnd, GW_CHILD)                                            '//取得桌面窗口的第一个子窗口
    Do While hwnd <> 0                                                          '//通过循环来枚举所有的窗口
        GetWindowText hwnd, strTitle, Len(strTitle)                             '//取得下一个窗口的标题，并写入到列表框中
        If Left$(strTitle, 1) <> vbNullChar Then                                '如果strtitle第一个字符不是空字符就继续
            '            tmp = Left$(strTitle, InStr(1, strTitle, vbNullChar))               'tmp=strtitle左起不是空值的字符串
            '            If Left(tmp, Len(Title)) = Title Then
            '                GetHandle = hwnd
            '            End If
            If InStr(strTitle, Title) > 0 Then
                GetHandle = hwnd
                Exit Do
            End If
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        '调用GetWindow函数，来取得下一个窗口
    Loop
End Function

'模糊查找标题获取桌面窗口的句柄
'程序调用举例:
'    模糊查找“新建”这样的标题：
'Dim Sx As String
'    Sx = GetOpenWindowNames("新建")
'    返回值Sx就是查找的全部标题
Public Function GetOpenWindowNames(Title As String) As String
    GetOpenWindowNames = Title                                                  '赋个初值，如果找不到会赋空值，影响后续的句柄查找
    Dim lngDeskTopHandle As Long
    Dim lngHand As Long, i As Long, S2 As String
    Dim strName As String * 255
    Dim lngWindowCount As Long
    lngDeskTopHandle = GetDesktopWindow()
    lngHand = GetWindow(lngDeskTopHandle, GW_CHILD)
    lngWindowCount = 1
    Do While lngHand <> 0
        i = GetWindowText(lngHand, strName, Len(strName))
        S2 = Left(strName, i)
        If InStr(1, S2, Title, vbTextCompare) > 0 Then                          '模糊查找
            GetOpenWindowNames = S2
            Exit Function
        End If
        lngHand = GetWindow(lngHand, GW_HWNDNEXT)
    Loop
End Function


' 枚举子窗口用法：
'if hwnd<>0 then EnumChildWindows hwnd, AddressOf EnumChildProc, ByVal 0&
Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    GetWindowText hwnd, ckbt, 50
    GetClassName hwnd, cklm, 50
    If InStr(cklm, "Afx:400000:b") > 0 Then                                     '如果类名含有需要的关键字
        EnumChildProc = 0                                                       '停止枚举子句柄
    Else
        EnumChildProc = 1                                                       '继续枚举子句柄
    End If
End Function

'Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
'   Dim WshShell As Object
'   Set WshShell = CreateObject("wscript.shell")
'   WshShell.Sendkeys CStr(text), wait
'   Set WshShell = Nothing
'End Sub
