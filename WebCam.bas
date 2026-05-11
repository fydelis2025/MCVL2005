Attribute VB_Name = "WebCam"
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "USER32" () As Long
Public Const WM_CAP_DRIVER_CONNECT As Long = 1034
Public Const WM_CAP_DRIVER_DISCONNECT As Long = 1035
Public Const WM_CAP_GRAB_FRAME As Long = 1084
Public Const WM_CAP_EDIT_COPY As Long = 1054
Public Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
Public Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066
Public Const WM_CLOSE = &H10
Public mCapHwnd As Long
