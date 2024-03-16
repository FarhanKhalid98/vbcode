Attribute VB_Name = "Declarations"
Option Explicit
Public ObjRegistry As New SoftinnRegistry.Registry
Public Declare Function SetWindowText Lib "User32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "User32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public CN As ADODB.Connection
Public CNR As ADODB.Connection
Public ParaOutID As String
Public ParaOutIDL As String
Public ParaPass As String
Public Char As Object
Public vZone As Boolean
Public vSector As Boolean
Public vEmployee As Boolean
Public vOrganization As Boolean
Public vSessionID As Byte
Public vQty As Boolean
Public vShowRetailPrice As Boolean
Public Sub SetConnection(ConnObject As ADODB.Connection)
  Set CN = ConnObject
End Sub

Public Sub SetChar(c As Object)
  Set Char = c
End Sub



