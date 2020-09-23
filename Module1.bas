Attribute VB_Name = "Module1"
Public mymod As Integer
Public MaxId(11) As Integer
Public duz_dev_ediyor As Boolean
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
   ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
' You can find more o these (lower) in the API Viewer.  Here
' they are used only for resizing the left and right
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public myobject As Object
Public myy, myx As Single
Public MYOBJEXT2 As String

Public MOV As Boolean
Public start2 As Boolean
Public myprop() As String
Public mytip As Integer
Public birak As Boolean
Public xobj As Object
Public mysel As Boolean
Public myerr As String
