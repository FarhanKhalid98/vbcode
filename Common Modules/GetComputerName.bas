Attribute VB_Name = "GetComputerName"
Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function LocalComputerName() As String
Dim strBuffer As String
  Dim lngBufSize As Long
  Dim lngStatus As Long
  
  lngBufSize = 255
  strBuffer = String$(lngBufSize, " ")
  lngStatus = GetComputerName(strBuffer, lngBufSize)
  If lngStatus <> 0 Then
      LocalComputerName = Left(strBuffer, lngBufSize)
  End If
End Function
