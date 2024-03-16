Attribute VB_Name = "Declarations"
Option Explicit
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public CN As ADODB.Connection
Public vReadOnly As Boolean
Public ParaPass As String

Public Sub SetConnection(ConnObject As ADODB.Connection)
  ' If ParaPass <> EncryptStr("›· ‹ÔÙÌÓ€ﬂÿ‡", False) Then Exit Sub
  Set CN = ConnObject
End Sub

