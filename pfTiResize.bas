Attribute VB_Name = "modTiResize"
Option Explicit
'API Declarations used for subclassing.
' You can find more o these (lower) in the API Viewer.  Here
' they are used only for resizing the left and right
''Public Const HTCLIENT                   As Integer = 1
''Public Const HTCAPTION                  As Integer = 2
''Public Const HTSYSMENU                  As Integer = 3
''Public Const HTGROWBOX                  As Integer = 4
''Public Const HTMENU                     As Integer = 5
''Public Const HTHSCROLL                  As Integer = 6
''Public Const HTVSCROLL                  As Integer = 7
''Public Const HTMINBUTTON                As Integer = 8
''Public Const HTMAXBUTTON                As Integer = 9
Public Const HTLEFT                     As Integer = 10
Public Const HTRIGHT                    As Integer = 11
Public Const HTTOP                      As Integer = 12
Public Const HTTOPLEFT                  As Integer = 13
Public Const HTTOPRIGHT                 As Integer = 14
Public Const HTBOTTOM                   As Integer = 15
Public Const HTBOTTOMLEFT               As Integer = 16
Public Const HTBOTTOMRIGHT              As Integer = 17
''Public Const HTBORDER                   As Integer = 18
Public Const WM_NCLBUTTONDOWN           As Long = &HA1
Public Const WM_SYSCOMMAND              As Long = &H112
Public Const SC_MOVE                    As Long = &HF010&
''Public Const SC_SIZE                    As Long = &HF000&
Public Const WM_SIZING                  As Long = &H214
Public Const WM_PRINTCLIENT             As Long = &H318
Public Const WM_PRINT                   As Long = &H317
Public Const WM_CTLCOLOREDIT            As Long = &H133
Public Const WM_CAPTURECHANGED          As Long = &H215
Public Const WM_GETMINMAXINFO           As Long = &H24
Public Const WM_ERASEBKGND              As Long = &H14
'This message is fired whilst the window is being resized. The lParam of the message points to a RECT structure containing the desired position of the window. Any modifications you make to this rectangle are passed back to Windows, which moves or sizes the window directly to the size and position you specify.
Public Const WM_MOVING                  As Long = &H216
'This message works the same way as WM_SIZING except it is fired whilst the window is being moved.
Public Const WM_ENTERSIZEMOVE           As Long = &H231
'This message is fired when your window is about to start moving or sizing.
Public Const WM_EXITSIZEMOVE            As Long = &H232
'This message is fired when a moving or sizing operation on your window has completed.
Public Const WM_SIZE                    As Long = &H5
'This message is fired whenever your window has its size changed by the SetWindowPos function, for example when windows minimizes, maximizes or restores your window, or when you call a VB function which changes the size of the window.
'The sample application shows how you can subclass these messages for a window and respond correctly to them, providing the following new
Public Const WM_MOVE                    As Long = &H3
'Constants for GetWindowLong() and SetWindowLong() APIs.
Public Const GWL_WNDPROC                As Long = (-4)
Public Const GWL_USERDATA               As Long = (-21)
''Public Const WM_MENUSELECT              As Long = &H11F
Public Const WM_PARENTNOTIFY            As Long = &H210
Public Const WM_MOUSEACTIVATE           As Long = &H21
Public Const WM_NOTIFY                  As Long = &H4E&
''Public Const WM_HSCROLL                 As Long = &H114
''Public Const WM_VSCROLL                 As Long = &H115
''Public Const NM_RCLICK                  As Integer = -5
''Public Const WM_LBUTTONDBLCLK           As Long = &H203
Public Const WM_LBUTTONDOWN             As Long = &H201
Public Const WM_LBUTTONUP               As Long = &H202
''Public Const WM_MBUTTONDBLCLK           As Long = &H209
''Public Const WM_MBUTTONDOWN             As Long = &H207
''Public Const WM_MBUTTONUP               As Long = &H208
''Public Const WM_RBUTTONDBLCLK           As Long = &H206
Public Const WM_RBUTTONDOWN             As Long = &H204
Public Const WM_RBUTTONUP               As Long = &H205
Public Const WM_MOUSEFIRST              As Long = &H200
Public Const WM_PAINT                   As Long = &HF
Public Const WM_COMMAND                 As Long = &H111
Public Const WM_SETCURSOR               As Long = &H20
Public Const WM_SETFOCUS                As Long = &H7
Public Const WM_SETHOTKEY               As Long = &H32
Public Const WM_SETREDRAW               As Long = &HB
Public Const WM_SETTEXT                 As Long = &HC
Public Const WM_SHOWWINDOW              As Long = &H18
Public Const WM_WINDOWPOSCHANGED        As Long = &H47
Public Const WM_WINDOWPOSCHANGING       As Long = &H46
Public Const WM_CHILDACTIVATE           As Long = &H22
Public Const WM_NCCALCSIZE              As Long = &H83
Public Const WM_NCCREATE                As Long = &H81
Public Const WM_NCDESTROY               As Long = &H82
Public Const WM_NCHITTEST               As Long = &H84
Public Const WM_NCACTIVATE              As Long = &H86
Public Const WM_NCLBUTTONUP             As Long = &HA2
Public Const WM_NCMBUTTONDOWN           As Long = &HA7
Public Const WM_NCMOUSEMOVE             As Long = &HA0
Public Const WM_NCPAINT                 As Long = &H85
Public Const WM_NCRBUTTONDOWN           As Long = &HA4
Public Type RECT
    Left                                  As Long
    Top                                   As Long
    Right                                 As Long
    Bottom                                As Long
End Type
Public Type Subclass
    hwnd                                  As Long
    ProcessId                             As Long
End Type
'Used to hold a reference to the control to call its procedure.
'NOTE: "UserControl1" is the UserControl.Name Property at
'      design-time of the .CTL file.
'      ('As Object' or 'As Control' does not work)
Private ctlShadowControl                As pfTiResize
'Used as a pointer to the UserData section of a window.
Public mWndSubClass(1)                  As Subclass
'Used as a pointer to the UserData section of a window.
Private ptrObject                       As Long
''Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                       ByVal wMsg As Long, _
                                                                       ByVal wParam As Long, _
                                                                       lParam As Any) As Long
''Public Declare Function GetBoundsRect Lib "gdi32" (ByVal Hdc As Long, lprcBounds As RECT, ByVal Flags As Long) As Long
''Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
''Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
''Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
''Public Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                    pSrc As Any, _
                                                                    ByVal ByteLen As Long)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                           ByVal nIndex As Long, _
                                                                           ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                           ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                             ByVal hwnd As Long, _
                                                                             ByVal Msg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long
''Public Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
'The address of this function is used for subclassing.
'Messages will be sent here and then forwarded to the
'UserControl's WindowProc function. The HWND determines
'to which control the message is sent.
Private END_OF_DECLARATION_DUMMY_WILL_BE_DELETED As Boolean
Public Function HiWord(Param As Long) As Integer

  Dim WordHex As String
  Dim offset  As Long

    WordHex = Hex$(Param)
    offset = Len(WordHex) - 4
    If offset > 0 Then
        HiWord = CInt("&H" & Left$(WordHex, offset))
     Else
        HiWord = 0
    End If

End Function

Public Function LoWord(Param As Long) As Integer

  Dim WordHex As String

    WordHex = Hex$(Param)
    LoWord = CInt("&H" & Right$(WordHex, 4))

End Function

Public Function MakeLong(LoWord As Integer, _
                         HiWord As Integer) As Long

  Dim nLoWord As Long

    'Creates a Long value using Low and High integers
    'Useful when converting code from C++
    If LoWord% < 0 Then
        nLoWord& = LoWord% + &H10000
     Else
        nLoWord& = LoWord%
    End If
    MakeLong& = CLng(nLoWord&) Or (HiWord% * &H10000)

End Function

Public Function MakeWord(LoByte As Byte, _
                         HiByte As Byte) As Integer

  Dim nLoByte As Integer

    'Creates an integer value using Low and High bytes
    'Useful when converting code from C++
    If LoByte < 0 Then
        nLoByte = LoByte + &H100
     Else
        nLoByte = LoByte
    End If
    MakeWord = CInt(nLoByte) Or (HiByte * &H100)

End Function

Public Function SubWndProc(ByVal hwnd As Long, _
                           ByVal Msg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

    On Error Resume Next
    'Get pointer to the control's VTable from the
    'window's UserData section. The VTable is an internal
    'structure that contains pointers to the methods and
    'properties of the control.
    ptrObject = GetWindowLong(mWndSubClass(0).hwnd, GWL_USERDATA)
    'Copy the memory that points to the VTable of our original
    'control to the shadow copy of the control you use to
    'call the original control's WindowProc Function.
    'This way, when you call the method of the shadow control,
    'you are actually calling the original controls' method.
    CopyMemory ctlShadowControl, ptrObject, 4
    'Call the WindowProc function in the instance of the UserControl.
    SubWndProc = ctlShadowControl.WindowProc(hwnd, Msg, wParam, lParam)
    'Destroy the Shadow Control Copy
    CopyMemory ctlShadowControl, 0&, 4
    Set ctlShadowControl = Nothing
    On Error GoTo 0

End Function

':) Roja's VB Code Fixer V1.0.99 (6/18/2003 8:12:41 PM) 122 + 79 = 201 Lines Thanks Ulli for inspiration and lots of code.

