Attribute VB_Name = "conexao"
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias _
"capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal _
x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, _
ByVal nlD As Long) As Long


