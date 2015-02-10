Attribute VB_Name = "ApiCalls"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias _
  "RtlMoveMemory" (dest As Any, Source As Any, _
  ByVal numBytes As Long)

Public Type API_PointType
    X As Long
    Y As Long
End Type
Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'hwnd
'The handle of the window to move and resize.
'x
'The x-coordinate to position the upper-left corner of the window at.
'y
'The y-coordinate to position the upper-left corner of the window at.
'nWidth
'The width in pixels to resize the window to.
'nHeight
'The height in pixels to resize the window to.
'bRepaint
'If 1, updates the screen to display the window at its new position. If 0, does not update the screen to reflect the move (the
'window will appear to be unmoved but will actually be at its new location!).
Public Declare Function ClientToScreen& Lib "user32" (ByVal hWnd As Long, lPPoint As API_PointType)
Declare Function SendMessageBynum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessageBystring Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const EM_GETLINE = &HC4
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function FloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lPRect As API_Rect, ByVal hBrush As Long) As Long
Public Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lPrcFrom As API_Rect, lPrcTo As API_Rect) As Long
Public Declare Function Drawcaption Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, PcRect As API_Rect, ByVal un As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As API_Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawEscape Lib "gdi32" (ByVal hDC As Long, ByVal nEscape As Long, ByVal cbinput As Long, ByVal lPszInData As String) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lPRect As API_Rect) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lPRect As API_Rect, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lPDrawStateproc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lPStr As String, ByVal nCount As Long, lPRect As API_Rect, ByVal wFormat As Long) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal N As Long, lPRect As API_Rect, ByVal un As Long, lPDrawTextParams As DRAWTEXTPARAMS) As Long
Public Declare Function ElliPse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lPRect As API_Rect, ByVal hBrush As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function FillPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long 'Used in DrawLine Function Below
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lPPoint As POINTAPI) As Long 'Used in DrawLine Function Below

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRoP As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRoP As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Const RDW_UPDATENOW As Long = &H100
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Any, ByVal fuRedraw As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Boolean

Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H4400328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086


'Functions more commonly used my PcMemDc
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lPDriverName As String, _
           lPDeviceName As Any, lPOutPut As Any, lPInitData As Any) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lPDriverName As String, ByVal lPDeviceName As String, ByVal lPOutPut As String, lPInitData As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lPObject As Any) As Long
'System Metrics
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    'nIndex is one of followingSPecifies the system metric or configuration setting to retrieve. All SM_CX* values are widths. All SM_CY* values are heights. The following values are defined:
''SM_ARRANGE
''Flags specifying how the system arranged minimized windows. For more information about minimized windows, see the following Remarks section.
''SM_CLEANBOOT
''Value that specifies how the system was started:
''0 Normal boot
''1 Fail-safe boot
''2 Fail-safe with network boot
''Fail-safe boot (also called SafeBoot) byPasses the user’s startup files.
''SM_CMOUSEBUTTONS
''Number of buttons on mouse, or zero if no mouse is installed.
''SM_CXBORDER,
''SM_CYBORDER
''The width and height, in Pixels, of a window border. This is equivalent to the SM_CXEDGE value for windows with the 3-D look.
''SM_CXCURSOR,
''SM_CYCURSOR
''Width and height, in Pixels, of a cursor. These are the cursor dimensions supported by the current display driver. The system cannot create cursors of other sizes.
''SM_CXDLGFRAME,
''SM_CYDLGFRAME
''Same as SM_CXFIXEDFRAME and SM_CYFIXEDFRAME.
''SM_CXDOUBLECLK,
''SM_CYDOUBLECLK
''Width and height, in Pixels, of the rectangle around the location of a first click in a double-click sequence. The second click must occur within this rectangle for the system to consider the two clicks a double-click. (The two clicks must also occur within a specified time.)
''SM_CXDRAG,
''SM_CYDRAG
''Width and height, in Pixels, of a rectangle centered on a drag Point to allow for limited movement of the mouse Pointer before a drag oPeration begins. This allows the user to click and release the mouse button easily without unintentionally starting a drag oPeration.
''SM_CXEDGE,
''SM_CYEDGE
''Dimensions, in Pixels, of a 3-D border. These are the 3-D counterParts of SM_CXBORDER and SM_CYBORDER.
''SM_CXFIXEDFRAME,
''SM_CYFIXEDFRAME
''Thickness, in Pixels, of the frame around the Perimeter of a window that has a caption but is not sizable. SM_CXFIXEDFRAME is the width of the horizontal border and SM_CYFIXEDFRAME is the height of the vertical border.
''Same as SM_CXDLGFRAME and SM_CYDLGFRAME.
''SM_CXFRAME,
''SM_CYFRAME
''Same as SM_CXSIZEFRAME and SM_CYSIZEFRAME.
''SM_CXFULLSCREEN , SM_CYFULLSCREEN
''Width and height of the client area for a full-screen window. To get the coordinates of the Portion of the screen not obscured by the tray, call the SystemparametersInfo function with the SPI_GETWORKAREA value.


''SM_CXHTHUMB
''Width, in Pixels, of the thumb box in a horizontal scroll bar.
''SM_CXICON,
''SM_CYICON
''The default width and height, in Pixels, of an icon. These values are tyPically 32x32, but can vary depending on the installed display hardware.
''The LoadIcon function can only load icons of these dimensions.
''SM_CXICONSPACING , SM_CYICONSPACING
''Dimensions, in Pixels, of a grid cell for items in large icon view. Each item fits into a rectangle of this size when arranged. These values are always greater than or equal to SM_CXICON and SM_CYICON.
''SM_CXMAXIMIZED,
''SM_CYMAXIMIZED
''Default dimensions, in Pixels, of a maximized toP-level window.
''SM_CXMAXTRACK,
''SM_CYMAXTRACK
''Default maximum dimensions, in Pixels, of a window that has a caption and sizing borders. The user cannot drag the window frame to a size larger than these dimensions. A window can override these values by Processing the WM_GETMINMAXINFO message.
''SM_CXMENUCHECK=''Dimensions, in Pixels, of the default menu check-mark bitmap.
''SM_CYMENUCHECK=''Dimensions, in Pixels, of the default menu check-mark bitmap.
''SM_CXMENUSIZE,
''SM_CYMENUSIZE
''Dimensions, in Pixels, of menu bar buttons, such as multiple document (MIDI) child close.
''SM_CXMIN,
''SM_CYMIN
''Minimum width and height, in Pixels, of a window.
''SM_CXMINIMIZED,
''SM_CYMINIMIZED
''Dimensions, in Pixels, of a normal minimized window.
''SM_CXMINSPACING
''SM_CYMINSPACING
''Dimensions, in Pixels, of a grid cell for minimized windows. Each minimized window fits into a rectangle this size when arranged. These values are always greater than or equal to SM_CXMINIMIZED and SM_CYMINIMIZED.
''SM_CXMINTRACK , SM_CYMINTRACK
''Minimum tracking width and height, in Pixels, of a window. The user cannot drag the window frame to a size smaller than these dimensions. A window can override these values by Processing the WM_GETMINMAXINFO message.
''SM_CXSCREEN,
''SM_CYSCREEN
''Width and height, in Pixels, of the screen.
''SM_CXSIZE,
''SM_CYSIZE
''Width and height, in Pixels, of a button in a window's caption or title bar.
''SM_CXSIZEFRAME,SM_CYSIZEFRAME Thickness, in Pixels, of the sizing border around the Perimeter of a window that can be resized. SM_CXSIZEFRAME is the width of the horizontal border and SM_CYSIZEFRAME is the height of the vertical border. ''Same as SM_CXFRAME and SM_CYFRAME.
''SM_CXSMICON,
''SM_CYSMICON
''Recommended dimensions, in Pixels, of a small icon. Small icons tyPically appear in window captions and in small icon view.
''SM_CXSMSIZE
''SM_CYSMSIZE
''Dimensions, in Pixels, of small caption buttons.
''SM_CXVSCROLL , SM_CYVSCROLL
''Width, in Pixels, of a vertical scroll bar; and height, in Pixels, of the arrow bitmap on a vertical scroll bar.

''SM_CYKANJIWINDOW
''For double-byte character set versions of Windows, height, in Pixels, of the Kanji window at the bottom of the screen.
''SM_CYMENU
''Height, in Pixels, of single-line menu bar.
''SM_CYSMCAPTION
''Height, in Pixels, of a small caption.
''SM_CYVTHUMB
''Height , in Pixels, of the thumb box in a vertical scroll bar.
''SM_DBCSENABLED
''TRUE or nonzero if the double-byte character set (DBCS) version of USER.EXE is installed; FALSE, or zero otherwise.
''SM_DEBUG
''TRUE or nonzero if the debugging version of USER.EXE is installed; FALSE, or zero, otherwise.
''SM_MENUDROPALIGNMENT
''TRUE, or nonzero if droP-down menus are right-aligned relative to the corresponding menu-bar item; FALSE, or zero if they are left-aligned.
''SM_MIDEASTENABLED
''TRUE if the system is enabled for Hebrew/Arabic languages.
''SM_MOUSEPRESENT
''TRUE or nonzero if a mouse is installed; FALSE, or zero, otherwise.
''SM_MOUSEWHEELPRESENT
''Windows NT only: TRUE or nonzero if a mouse with a wheel is installed; FALSE, or zero, otherwise.
''SM_NETWORK
''The least significant bit is set if a network is Present; otherwise, it is cleared. The other bits are reserved for future use.
''SM_PENWINDOWS
''TRUE or nonzero if the Microsoft Windows for Pen comPuting extensions are installed; zero, or FALSE, otherwise.
''SM_SECURE
''TRUE if security is Present, FALSE otherwise.
''SM_SHOWSOUNDS
''TRUE or nonzero if the user requires an Application to Present information visually in situations where it would otherwise Present the information only in audible form; FALSE, or zero, otherwise.
''SM_SLOWMACHINE
''TRUE if the comPuter has a low-end (slow) Processor, FALSE otherwise.
''SM_SWAPBUTTON
''TRUE or nonzero if the meanings of the left and right mouse buttons are swapped; FALSE, or zero, otherwise.
'Private Const SM_CXSCREEN = 0
'Private Const SM_CYSCREEN = 1
'Private Const SM_CXVSCROLL = 2 'Width, in Pixels, of a vertical scroll bar; and height, in Pixels, of the arrow bitmap on a vertical scroll bar.
Public Const SM_CYHSCROLL = 3 'Width, in Pixels, of the arrow bitmap on a horizontal scroll bar; and height, in Pixels, of a horizontal scroll bar.
Public Const SM_CYCAPTION = 4  'Height, in Pixels, of normal caption area.
'Private Const SM_CXBORDER = 5
'Private Const SM_CYBORDER = 6
'Private Const SM_CXDLGFRAME = 7
'Private Const SM_CYDLGFRAME = 8
'Private Const SM_CYVTHUMB = 9
'Private Const SM_CXHTHUMB = 10
'Private Const SM_CXICON = 11
'Private Const SM_CYICON = 12
'Private Const SM_CXCURSOR = 13
'Private Const SM_CYCURSOR = 14
Public Const SM_CYMENU = 15
'Private Const SM_CXFULLSCREEN = 16
'Private Const SM_CYFULLSCREEN = 17
'Private Const SM_CYKANJIWINDOW = 18
'Private Const SM_MOUSEPRESENT = 19
'Private Const SM_CYVSCROLL = 20
'Private Const SM_CXHSCROLL = 21'Width, in Pixels, of the arrow bitmap on a horizontal scroll bar; and height, in Pixels, of a horizontal scroll bar.
'Private Const SM_DEBUG = 22
'Private Const SM_SWAPBUTTON = 23
'Private Const SM_RESERVED1 = 24
'Private Const SM_RESERVED2 = 25
'Private Const SM_RESERVED3 = 26
'Private Const SM_RESERVED4 = 27
'Private Const SM_CXMIN = 28
'Private Const SM_CYMIN = 29
'Private Const SM_CXSIZE = 30
'Private Const SM_CYSIZE = 31
'Private Const SM_CXFRAME = 32' Thickness, in Pixels, of the frame around the Perimeter of a window that has a caption but is not sizable. SM_CXFIXEDFRAME is the width of the horizontal border and SM_CYFIXEDFRAME is the height of the vertical border.
'Private Const SM_CYFRAME = 33' Thickness, in Pixels, of the frame around the Perimeter of a window that has a caption but is not sizable. SM_CXFIXEDFRAME is the width of the horizontal border and SM_CYFIXEDFRAME is the height of the vertical border.
'Private Const SM_CXMINTRACK = 34
'Private Const SM_CYMINTRACK = 35
'Private Const SM_CXDOUBLECLK = 36
'Private Const SM_CYDOUBLECLK = 37
'Private Const SM_CXICONSPACING = 38
'Private Const SM_CYICONSPACING = 39
'Private Const SM_MENUDROPALIGNMENT = 40
'Private Const SM_PENWINDOWS = 41
'Private Const SM_DBCSENABLED = 42
'Private Const SM_CMOUSEBUTTONS = 43
'Private Const SM_CMETRICS = 44

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As Long) As Long

Public Const DT_BOTTOM = &H8
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_CALCRECT = &H400
Public Const DT_WORDBREAK = &H10
Public Const DT_VCENTER = &H4
Public Const DT_TOP = &H0
Public Const DT_TABSTOP = &H80
Public Const DT_SINGLELINE = &H20
Public Const DT_RIGHT = &H2
Public Const DT_NOCLIP = &H100
Public Const DT_INTERNAL = &H1000
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_EXPANDTABS = &H40
Public Const DT_CHARSTREAM = 4
Public Const DT_EDITCONTROL = &H2000&
Public Const DT_PATH_ELLIPSIS = &H4000&
Public Const DT_END_ELLIPSIS = &H8000&
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000

Private Const CB_FINDSTRING = &H14C


Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'------------**scrollbars ---------------------------*
Const GWL_STYLE = (-16)
Const WS_VSCROLL = &H200000
Const WS_HSCROLL = &H100000

'---------------------------------------------------------------
' Procedure .....Fill Region With SPecified Color
'---------------------------------------------------------------
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal colorCode As Long, _
    ByVal fillType As Long) As Long
Const FLOODFILLBORDER = 0
Const FLOODFILLSURFACE = 1

'END

'Type Delcares
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmplanes As Long
    bmBitspixel As Integer
    bmBits As Long
End Type
Public Type API_Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function ClipCursor Lib "user32" (lPRect As API_Rect) As Long
Private Declare Function ClipCursorByNum Lib "user32" Alias "ClipCursor" (ByVal _
  lPRect As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lPPoint As API_PointType) As Long
Public Declare Function GetCursor Lib "user32" () As Long     'Gets a handle to the current mouse cursor
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long          'bShow=0 will hide,1 will show
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lPRect As API_Rect) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, _
  lPRect As API_Rect, ByVal bErase As Long) As Long
Private Declare Function InvalidateRectByNum Lib "user32" Alias _
  "InvalidateRect" (ByVal hWnd As Long, ByVal lPRect As Long, _
  ByVal bErase As Long) As Long
  
'Send a mouse click
'Simulate a left click mouse_event( MOUSEEVENTF_LEFTDOWN, dx=0, dy=0, 0, 0 )
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUp = &H4 '  left button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Public Const MOUSEEVENTF_MIDDLEUp = &H40 '  middle button up
Public Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Public Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Public Const MOUSEEVENTF_RIGHTUp = &H10 '  right button up
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Sub SetWindowAsChild(Hwnd_Parent As Long, Hwnd_Child As Long)
  SetParent Hwnd_Child, Hwnd_Parent
End Sub
Sub FormAlwaysOnTop(Frm As Form, action As Boolean)
  Const HWND_TOPMOST = -1
  Const HWND_NOTOPMOST = -2
  Const SWP_NOMOVE = 2
  Const SWP_NOSIZE = 1

' When called by a form: If action <> 0 makes the form float (always on toP)
' If action = 0 "unfloats" the window.
  Dim wFlags&, Ans&
  wFlags = SWP_NOMOVE Or SWP_NOSIZE
  If action Then
   Ans = SetWindowPos(Frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, wFlags)
  Else
   Ans = SetWindowPos(Frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, wFlags)
  End If
  If Ans = 0 Then MsgBox ("API ERROR"): MsgBox ("Critical Problem")
End Sub

