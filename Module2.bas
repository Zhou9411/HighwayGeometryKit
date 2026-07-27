Attribute VB_Name = "Module2"
Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
'Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Type GENERALINPUT
    dwType As Long
    xi(0 To 23) As Byte
End Type

Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

Const INPUT_KEYBOARD = 1
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Const KEYEVENTF_down = &H0
Public Const KEYEVENTF_KEYUP = &H2
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

'调用方法举例：
'MySendKey vbKeyA '向前台窗口发送字符A
Sub MySendKey(ParamArray arg() As Variant)
    '参数bkey传入要模拟按键的虚拟码即可模拟按下指定键
    Dim GInput(0 To 100) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    
    '    KInput.wVk = bkey                   '你要模拟的按键
    '    KInput.dwFlags = 0                  '按下键标志
    '    GInput(0).dwType = INPUT_KEYBOARD
    '    CopyMemory GInput(0).xi(0), KInput, Len(KInput) '这个函数用来把内存中KInput的数据复制到GInput
    '    KInput.wVk = bkey
    '    KInput.dwFlags = KEYEVENTF_KEYUP    ' 释放按键
    '    GInput(1).dwType = INPUT_KEYBOARD   ' 表示该消息为键盘消息
    '    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    '    '以上工作把按下键和释放键共2条键盘消息加入到GInput数据结构中
    '    SendInput 2, GInput(0), Len(GInput(0)) '把上面的按键消息插入到windows消息队列
    
    nx = 0
    For Each x In arg
        KInput.wVk = x                                                          '你要模拟的按键
        KInput.dwFlags = 0                                                      '按下键标志
        GInput(nx).dwType = INPUT_KEYBOARD
        CopyMemory GInput(nx).xi(0), KInput, Len(KInput)                        '这个函数用来把内存中KInput的数据复制到GInput
        Call SendInput(1, GInput(nx), Len(GInput(nx)))                          '把上面的按键消息插入到windows消息队列
        Sleep 1
        KInput.wVk = x
        KInput.dwFlags = KEYEVENTF_KEYUP                                        ' 释放按键
        GInput(nx + 1).dwType = INPUT_KEYBOARD                                  ' 表示该消息为键盘消息
        CopyMemory GInput(nx + 1).xi(0), KInput, Len(KInput)
        '以上工作把按下键和释放键共2条键盘消息加入到GInput数据结构中
        Call SendInput(1, GInput(nx + 1), Len(GInput(nx + 1)))                  '把上面的按键消息插入到windows消息队列
        Sleep 1
        nx = nx + 1
    Next x
End Sub

'调用方法举例：
'MySendKeytext "12345" '向前台窗口发送字符12345
Sub MySendKeytext(text As Variant)
    '参数bkey传入要模拟按键的虚拟码即可模拟按下指定键
    Dim GInput(0 To 200) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    For nx = 1 To Len(text)
        tx = Mid(text, nx, 1)
        '        Debug.Print Mid(text, nx, 1) & VkKeyScan(Asc(Mid(text, nx, 1)))
        If Asc(tx) > 64 And Asc(tx) < 91 Or Asc(tx) = 33 Or Asc(tx) = 34 Or Asc(tx) = 35 Or Asc(tx) = 36 Or Asc(tx) = 37 Or Asc(tx) = 38 _
            Or Asc(tx) = 40 Or Asc(tx) = 41 Or Asc(tx) = 42 Or Asc(tx) = 43 _
            Or Asc(tx) = 58 Or Asc(tx) = 60 Or Asc(tx) = 62 Or Asc(tx) = 63 Or Asc(tx) = 64 Or Asc(tx) = 94 Or Asc(tx) = 95 _
            Or Asc(tx) = 123 Or Asc(tx) = 124 Or Asc(tx) = 125 Or Asc(tx) = 126 Then
            'MsgBox "大写字母"
            '33=! 34双引号，35=#,36=$,37=%,38=&,40=(,41=),42=*,43=+,58冒号,60=<,62=>，63=?,64=@,94=^,95=_,123={,125=},124=|,126=~
            
            '按下SHIFT
            KInput.wVk = vbKeyShift                                             '你要模拟的按键
            KInput.dwFlags = 0                                                  '按下键标志
            GInput(0).dwType = INPUT_KEYBOARD
            CopyMemory GInput(0).xi(0), KInput, Len(KInput)                     '这个函数用来把内存中KInput的数据复制到GInput
            Call SendInput(1, GInput(0), Len(GInput(0)))                        '把上面的按键消息插入到windows消息队列
            Sleep 1
            
            KInput.wVk = VkKeyScan(Asc(Mid(text, nx, 1)))                       '你要模拟的按键
            KInput.dwFlags = 0                                                  '按下键标志
            GInput(nx).dwType = INPUT_KEYBOARD
            CopyMemory GInput(nx).xi(0), KInput, Len(KInput)                    '这个函数用来把内存中KInput的数据复制到GInput
            Call SendInput(1, GInput(nx), Len(GInput(nx)))                      '把上面的按键消息插入到windows消息队列
            Sleep 1
            KInput.wVk = VkKeyScan(Asc(Mid(text, nx, 1)))
            KInput.dwFlags = KEYEVENTF_KEYUP                                    ' 释放按键
            GInput(nx + 1).dwType = INPUT_KEYBOARD                              ' 表示该消息为键盘消息
            CopyMemory GInput(nx + 1).xi(0), KInput, Len(KInput)
            Call SendInput(1, GInput(nx + 1), Len(GInput(nx + 1)))              '把上面的按键消息插入到windows消息队列
            Sleep 1
            
            
            '弹起SHIFT
            KInput.wVk = vbKeyShift
            KInput.dwFlags = KEYEVENTF_KEYUP                                    ' 释放按键
            GInput(nx + 2).dwType = INPUT_KEYBOARD                              ' 表示该消息为键盘消息
            CopyMemory GInput(nx + 2).xi(0), KInput, Len(KInput)                '这个函数用来把内存中KInput的数据复制到GInput
            Call SendInput(1, GInput(nx + 2), Len(GInput(nx + 2)))              '把上面的按键消息插入到windows消息队列
            Sleep 1
            
        Else
            'Asc(tx) > 96 And Asc(tx) < 123 Then
            'MsgBox "小写字母"
            KInput.wVk = VkKeyScan(Asc(Mid(text, nx, 1)))                       '你要模拟的按键
            KInput.dwFlags = 0                                                  '按下键标志
            GInput(nx).dwType = INPUT_KEYBOARD
            CopyMemory GInput(nx).xi(0), KInput, Len(KInput)                    '这个函数用来把内存中KInput的数据复制到GInput
            Call SendInput(1, GInput(nx), Len(GInput(nx)))                      '把上面的按键消息插入到windows消息队列
            Sleep 1
            KInput.wVk = VkKeyScan(Asc(Mid(text, nx, 1)))
            KInput.dwFlags = KEYEVENTF_KEYUP                                    ' 释放按键
            GInput(nx + 1).dwType = INPUT_KEYBOARD                              ' 表示该消息为键盘消息
            CopyMemory GInput(nx + 1).xi(0), KInput, Len(KInput)
            Call SendInput(1, GInput(nx + 1), Len(GInput(nx + 1)))              '把上面的按键消息插入到windows消息队列
            Sleep 1
            'MsgBox "非字母"，暂未细分。
        End If
        
    Next nx
End Sub



'把一个按键消息分成两步，第一步按下，第二步弹起，调用方法举例1：
'    MySendKeyStep vbKeyControl, 0'按下
'    Sleep 1
'    MySendKeyStep vbKeyA, 0'按下
'    Sleep 1
'    MySendKeyStep vbKeyA, 1'弹起
'    Sleep 1
'    MySendKeyStep vbKeyControl, 1'弹起
'举例2：
'    MySendKeyStep vbKeyF3, 0 '按下
'    Sleep 1
'    MySendKeyStep vbKeyF3, 1 '弹起
Sub MySendKeyStep(bkey As Integer, press)
    '参数bkey传入要模拟按键的虚拟码即可模拟按下指定键
    Dim GInput(0 To 1) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    If press = 0 Then                                                           '如果按键按下
        KInput.wVk = bkey                                                       '你要模拟的按键
        KInput.dwFlags = 0                                                      '按下键标志
        GInput(0).dwType = INPUT_KEYBOARD
        CopyMemory GInput(0).xi(0), KInput, Len(KInput)                         '这个函数用来把内存中KInput的数据复制到GInput
        Call SendInput(1, GInput(0), Len(GInput(0)))                            '把上面的按键消息插入到windows消息队列
    Else                                                                        '如果按键弹起
        KInput.wVk = bkey
        KInput.dwFlags = KEYEVENTF_KEYUP                                        ' 释放按键
        GInput(0).dwType = INPUT_KEYBOARD                                       ' 表示该消息为键盘消息
        CopyMemory GInput(0).xi(0), KInput, Len(KInput)
        Call SendInput(1, GInput(0), Len(GInput(0)))                            '把上面的按键消息插入到windows消息队列
    End If
End Sub

'调用方法举例：
'SendKeyCTRLplus vbKeyA 'ctrl+a
Sub SendKeyCTRLplus(bkey As Integer)
    '参数bkey传入要模拟按键的虚拟码即可模拟按下指定键
    Dim GInput(0 To 100) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    '按下ctrl
    KInput.wVk = vbKeyControl                                                   '你要模拟的按键
    KInput.dwFlags = 0                                                          '按下键标志
    GInput(0).dwType = INPUT_KEYBOARD
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)                             '这个函数用来把内存中KInput的数据复制到GInput
    Call SendInput(1, GInput(0), Len(GInput(0)))                                '把上面的按键消息插入到windows消息队列
    Sleep 1
    
    '按下bkey
    KInput.wVk = bkey                                                           '你要模拟的按键
    KInput.dwFlags = 0                                                          '按下键标志
    GInput(1).dwType = INPUT_KEYBOARD
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)                             '这个函数用来把内存中KInput的数据复制到GInput
    Call SendInput(1, GInput(1), Len(GInput(1)))                                '把上面的按键消息插入到windows消息队列
    Sleep 1
    
    '弹起bkey
    KInput.wVk = bkey
    KInput.dwFlags = KEYEVENTF_KEYUP                                            ' 释放按键
    GInput(2).dwType = INPUT_KEYBOARD                                           ' 表示该消息为键盘消息
    CopyMemory GInput(2).xi(0), KInput, Len(KInput)                             '这个函数用来把内存中KInput的数据复制到GInput
    Call SendInput(1, GInput(2), Len(GInput(2)))                                '把上面的按键消息插入到windows消息队列
    Sleep 1
    
    '弹起ctrl
    KInput.wVk = vbKeyControl
    KInput.dwFlags = KEYEVENTF_KEYUP                                            ' 释放按键
    GInput(3).dwType = INPUT_KEYBOARD                                           ' 表示该消息为键盘消息
    CopyMemory GInput(3).xi(0), KInput, Len(KInput)                             '这个函数用来把内存中KInput的数据复制到GInput
    Call SendInput(1, GInput(3), Len(GInput(3)))                                '把上面的按键消息插入到windows消息队列
End Sub

'调用方法举例：
'SendKeyALTplus vbKeyH, vbKeyA 'ALT+H,A
Sub SendKeyALTplus(bkey As Integer, ParamArray arg() As Variant)
    'Sub SendKeyALTplus(bkey As Integer, bkey2 As Variant)
    '参数bkey传入要模拟按键的虚拟码即可模拟按下指定键
    Dim GInput(0 To 100) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    '按下alt
    KInput.wVk = vbKeyMenu                                                      '你要模拟的按键
    KInput.dwFlags = 0                                                          '按下键标志
    GInput(0).dwType = INPUT_KEYBOARD
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)                             '这个函数用来把内存中KInput的数据复制到GInput
    Call SendInput(1, GInput(0), Len(GInput(0)))                                '把上面的按键消息插入到windows消息队列
    Sleep 1
    
    '按下bkey
    KInput.wVk = bkey                                                           '你要模拟的按键
    KInput.dwFlags = 0                                                          '按下键标志
    GInput(1).dwType = INPUT_KEYBOARD
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)                             '这个函数用来把内存中KInput的数据复制到GInput
    Call SendInput(1, GInput(1), Len(GInput(1)))                                '把上面的按键消息插入到windows消息队列
    Sleep 1
    
    '弹起bkey
    KInput.wVk = bkey
    KInput.dwFlags = KEYEVENTF_KEYUP                                            ' 释放按键
    GInput(2).dwType = INPUT_KEYBOARD                                           ' 表示该消息为键盘消息
    CopyMemory GInput(2).xi(0), KInput, Len(KInput)                             '这个函数用来把内存中KInput的数据复制到GInput
    Call SendInput(1, GInput(2), Len(GInput(2)))                                '把上面的按键消息插入到windows消息队列
    Sleep 1
    
    '弹起alt
    KInput.wVk = vbKeyMenu
    KInput.dwFlags = KEYEVENTF_KEYUP                                            ' 释放按键
    GInput(3).dwType = INPUT_KEYBOARD                                           ' 表示该消息为键盘消息
    CopyMemory GInput(3).xi(0), KInput, Len(KInput)                             '这个函数用来把内存中KInput的数据复制到GInput
    Call SendInput(1, GInput(3), Len(GInput(3)))                                '把上面的按键消息插入到windows消息队列
    Sleep 1
    
    nx = 4
    For Each x In arg
        KInput.wVk = x                                                          '你要模拟的按键
        KInput.dwFlags = 0                                                      '按下键标志
        GInput(nx).dwType = INPUT_KEYBOARD
        CopyMemory GInput(nx).xi(0), KInput, Len(KInput)                        '这个函数用来把内存中KInput的数据复制到GInput
        Call SendInput(1, GInput(nx), Len(GInput(nx)))                          '把上面的按键消息插入到windows消息队列
        Sleep 1
        KInput.wVk = x
        KInput.dwFlags = KEYEVENTF_KEYUP                                        ' 释放按键
        GInput(nx + 1).dwType = INPUT_KEYBOARD                                  ' 表示该消息为键盘消息
        CopyMemory GInput(nx + 1).xi(0), KInput, Len(KInput)
        '以上工作把按下键和释放键共2条键盘消息加入到GInput数据结构中
        Call SendInput(1, GInput(nx + 1), Len(GInput(nx + 1)))                  '把上面的按键消息插入到windows消息队列
        Sleep 1
        nx = nx + 1
    Next x
    
    '    '按下bkey2
    '    KInput.wVk = bkey2                  '你要模拟的按键
    '    KInput.dwFlags = 0                  '按下键标志
    '    GInput(4).dwType = INPUT_KEYBOARD
    '    CopyMemory GInput(4).xi(0), KInput, Len(KInput) '这个函数用来把内存中KInput的数据复制到GInput
    '    Call SendInput(1, GInput(4), Len(GInput(4))) '把上面的按键消息插入到windows消息队列
    '    Sleep 1
    '
    '    '弹起bkey2
    '    KInput.wVk = bkey2
    '    KInput.dwFlags = KEYEVENTF_KEYUP    ' 释放按键
    '    GInput(5).dwType = INPUT_KEYBOARD   ' 表示该消息为键盘消息
    '    CopyMemory GInput(5).xi(0), KInput, Len(KInput) '这个函数用来把内存中KInput的数据复制到GInput
    '    Call SendInput(1, GInput(5), Len(GInput(5))) '把上面的按键消息插入到windows消息队列
End Sub

'调用方法举例：
'SendKeySHIFTplus vbKeyH, vbKeyA 'SHIFT+H,A
Sub SendKeySHIFTplus(ParamArray arg() As Variant)
    'Sub SendKeySHIFTplus(bkey As Integer, bkey2 As Variant)
    '参数bkey传入要模拟按键的虚拟码即可模拟按下指定键
    Dim GInput(0 To 100) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    '按下SHIFT
    KInput.wVk = vbKeyShift                                                     '你要模拟的按键
    KInput.dwFlags = 0                                                          '按下键标志
    GInput(0).dwType = INPUT_KEYBOARD
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)                             '这个函数用来把内存中KInput的数据复制到GInput
    Call SendInput(1, GInput(0), Len(GInput(0)))                                '把上面的按键消息插入到windows消息队列
    Sleep 1
    
    '    '按下bkey
    '    KInput.wVk = bkey                   '你要模拟的按键
    '    KInput.dwFlags = 0                  '按下键标志
    '    GInput(1).dwType = INPUT_KEYBOARD
    '    CopyMemory GInput(1).xi(0), KInput, Len(KInput) '这个函数用来把内存中KInput的数据复制到GInput
    '    Call SendInput(1, GInput(1), Len(GInput(1))) '把上面的按键消息插入到windows消息队列
    '    Sleep 1
    '
    '    '弹起bkey
    '    KInput.wVk = bkey
    '    KInput.dwFlags = KEYEVENTF_KEYUP    ' 释放按键
    '    GInput(2).dwType = INPUT_KEYBOARD   ' 表示该消息为键盘消息
    '    CopyMemory GInput(2).xi(0), KInput, Len(KInput) '这个函数用来把内存中KInput的数据复制到GInput
    '    Call SendInput(1, GInput(2), Len(GInput(2))) '把上面的按键消息插入到windows消息队列
    '    Sleep 1
    
    nx = 1
    For Each x In arg
        KInput.wVk = x                                                          '你要模拟的按键
        KInput.dwFlags = 0                                                      '按下键标志
        GInput(nx).dwType = INPUT_KEYBOARD
        CopyMemory GInput(nx).xi(0), KInput, Len(KInput)                        '这个函数用来把内存中KInput的数据复制到GInput
        Call SendInput(1, GInput(nx), Len(GInput(nx)))                          '把上面的按键消息插入到windows消息队列
        Sleep 1
        KInput.wVk = x
        KInput.dwFlags = KEYEVENTF_KEYUP                                        ' 释放按键
        GInput(nx + 1).dwType = INPUT_KEYBOARD                                  ' 表示该消息为键盘消息
        CopyMemory GInput(nx + 1).xi(0), KInput, Len(KInput)
        '以上工作把按下键和释放键共2条键盘消息加入到GInput数据结构中
        Call SendInput(1, GInput(nx + 1), Len(GInput(nx + 1)))                  '把上面的按键消息插入到windows消息队列
        Sleep 1
        nx = nx + 1
    Next x
    
    '弹起SHIFT
    KInput.wVk = vbKeyShift
    KInput.dwFlags = KEYEVENTF_KEYUP                                            ' 释放按键
    GInput(nx + 2).dwType = INPUT_KEYBOARD                                      ' 表示该消息为键盘消息
    CopyMemory GInput(nx + 2).xi(0), KInput, Len(KInput)                        '这个函数用来把内存中KInput的数据复制到GInput
    Call SendInput(1, GInput(nx + 2), Len(GInput(nx + 2)))                      '把上面的按键消息插入到windows消息队列
    Sleep 1
End Sub

