Attribute VB_Name = "Tooltip"
Option Explicit
' Beware of word-wrap when copying this code!
' API calls necessary for creating and destroying tool tip windows.
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

' A RECT user defined type. This is used for setting the bounds of the tool tip window.
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

' A TOOLINFO user defined type. This is used for setting 'all of the necessary
' flags when creating a tool tip window.
Private Type TOOLINFO
cbSize As Long
uFlags As Long
hwnd As Long
uid As Long
RECT As RECT
hinst As Long
lpszText As String
lParam As Long
End Type

' A constant used in conjunction with the CreateWindowEx 'API. It indicates to use the default value.
Private Const CW_USEDEFAULT = &H80000000

' Constants used with the SetWindowPosition API.
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_BOTTOM = 1

' Constants for setting the style of the tool tip window.
Private Const WS_POPUP = &H80000000
Private Const WS_EX_TOPMOST = &H8&

' A constant used with the SendMessage API to define 'private messages.
Private Const WM_USER = &H400

' Messages used for setting the duration time of tool 'tips.
' Not used here.
Private Const TTDT_AUTOMATIC = 0
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TTDT_RESHOW = 1

' All of the flags for tool tip windows.
Private Const TTF_ABSOLUTE = &H80
Private Const TTF_CENTERTIP = &H2
Private Const TTF_DI_SETITEM = &H8000
Private Const TTF_IDISHWND = &H1
Private Const TTF_RTLREADING = &H4
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_TRANSPARENT = &H100

' All of the available messages for tool tip windows.
Private Const TTM_ACTIVATE = (WM_USER + 1)
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ADDTOOLW = (WM_USER + 50)
Private Const TTM_ADJUSTRECT = (WM_USER + 31)
Private Const TTM_DELTOOLA = (WM_USER + 5)
Private Const TTM_DELTOOLW = (WM_USER + 51)
Private Const TTM_ENUMTOOLSA = (WM_USER + 14)
Private Const TTM_ENUMTOOLSW = (WM_USER + 58)
Private Const TTM_GETBUBBLESIZE = (WM_USER + 30)
Private Const TTM_GETCURRENTTOOLA = (WM_USER + 15)
Private Const TTM_GETCURRENTTOOLW = (WM_USER + 59)
Private Const TTM_GETDELAYTIME = (WM_USER + 21)
Private Const TTM_GETMARGIN = (WM_USER + 27)
Private Const TTM_GETMAXTIPWIDTH = (WM_USER + 25)
Private Const TTM_GETTEXTA = (WM_USER + 11)
Private Const TTM_GETTEXTW = (WM_USER + 56)
Private Const TTM_GETTIPBKCOLOR = (WM_USER + 22)
Private Const TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
Private Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Private Const TTM_GETTOOLINFOA = (WM_USER + 8)
Private Const TTM_GETTOOLINFOW = (WM_USER + 53)
Private Const TTM_HITTESTA = (WM_USER + 10)
Private Const TTM_HITTESTW = (WM_USER + 55)
Private Const TTM_NEWTOOLRECTA = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW = (WM_USER + 52)
Private Const TTM_POP = (WM_USER + 28)
Private Const TTM_RELAYEVENT = (WM_USER + 7)
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTM_SETMARGIN = (WM_USER + 26)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLEA = (WM_USER + 32)
Private Const TTM_SETTITLEW = (WM_USER + 33)
Private Const TTM_SETTOOLINFOA = (WM_USER + 9)
Private Const TTM_SETTOOLINFOW = (WM_USER + 54)
Private Const TTM_TRACKACTIVATE = (WM_USER + 17)
Private Const TTM_TRACKPOSITION = (WM_USER + 18)
Private Const TTM_UPDATE = (WM_USER + 29)
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_UPDATETIPTEXTW = (WM_USER + 57)
Private Const TTM_WINDOWFROMPOINT = (WM_USER + 16)

' Constants for setting the style of the tool tip window.
'
' Always tip, even if the parent window is not active.
Private Const TTS_ALWAYSTIP = &H1
'
' Use the balloon style tool tip.
Private Const TTS_BALLOON = &H40
'
' Win98 and up - do not use sliding tool tips.
Private Const TTS_NOANIMATE = &H10
'
' Win2K and up - do not fade in tool tips.
Private Const TTS_NOFADE = &H20
'
' Prevents windows from removing any ampersand characters 'in the tool tip
' string. Without this flag, Windows will automatically 'remove ampersand
' characters from the string. This is done to allow the 'same string to be
' used as the tool tip text, and as the caption of a 'control.
Private Const TTS_NOPREFIX = &H2


' The two different tool tip classes.
Private Const TOOLTIPS_CLASS = "tooltips_class"
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

' A long for containing the hwnd (window handle) of the 'tool tip window created in
' this demo. This would have to be an array of longs if 'we were to create
' tool tip windows for multiple windows/controls.
Dim hwndTT As Long

Sub colocatip(Controle As Control, strPntr As String, Form As Form)
' Declare a toolinfo UDT variable.
Dim ti As TOOLINFO

' A rect UDT variable used to indicate the bounds of the 'tool tip window's
' region.
Dim RECT As RECT

' A long used to pass to the toolinfo UDT as an ID for 'the tool tip window.
' Does not really do anything in this demo, but is used 'for clarity.
Dim uid As Long
uid = 0

' A string that contains the text of the tool tip.


' A long to contain the return values of the API calls.
Dim RetVal As Long


' Create a tool tip window, and capture it's hwnd in the 'form-wide hwndTT long.
hwndTT = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASSA, vbNullString, WS_POPUP Or TTS_NOPREFIX Or TTS_BALLOON, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, Form.hwnd, 0, App.hInstance, 0)

' Use the setwindowpos API to set the window position of 'the tool tip window.
SetWindowPos hwndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE

' Determine the rectangle of the control we are adding 'the tool tip to. This
' will be the region in which the tool tip window will be 'activated.
GetClientRect Controle.hwnd, RECT

' Set all of the necessary info in the toolinfo UDT.
'
' The size of the toolinfo UDT in bytes. Must be set!
ti.cbSize = Len(ti)
'
' The flags that we want to pass to the tool tip. 'TTF_CENTERTIP is not
' necessary, but centers the tool tip to the window it is 'being applied to
' (when possible). TTF_SUBCLASS tells the tool tip window 'to subclass the
' window it is being applied to. This is the best route 'to take in VB, so
' subclassing by the developer is not necessary.
ti.uFlags = TTF_CENTERTIP Or TTF_SUBCLASS
'
' The hwnd of the control having the tool tip applied.
ti.hwnd = Controle.hwnd
'
' The instance of the app the tool tip applies to.
ti.hinst = App.hInstance
'
' The ID (hwnd) of the tool tip window. Not necessary 'unless the window is
' created using the TTF_IDISHWND flag.
ti.uid = uid
'
' A pointer to the tool tip text.
ti.lpszText = strPntr
'
' The coordinates that specify the tool tip window's 'region of activation.
ti.RECT.Left = RECT.Left
ti.RECT.Top = RECT.Top
ti.RECT.Right = RECT.Right
ti.RECT.Bottom = RECT.Bottom


' Send a message to the tool tip window telling it to 'apply itself to the
' control it is being added to.
RetVal = SendMessage(hwndTT, TTM_ADDTOOLA, 0, ti)

' Send a message to the tool tip window telling it to set 'the maximum tip
' width, to allow line breaking.
RetVal = SendMessage(hwndTT, TTM_SETMAXTIPWIDTH, 0, 80)

' Send messages to the tool tip window telling it what 'it's fore and back
' colours are.
RetVal = SendMessage(hwndTT, TTM_SETTIPBKCOLOR, RGB(255, 255, 255), 0)
RetVal = SendMessage(hwndTT, TTM_SETTIPTEXTCOLOR, RGB(0, 0, 150), 0)

' Send a message to the tool tip window telling it to 'update itself
' (to reflect the new fore and back colours).
RetVal = SendMessage(hwndTT, TTM_UPDATETIPTEXTA, 0, ti)


End Sub

