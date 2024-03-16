Attribute VB_Name = "Declarations"
Option Explicit
Public ObjRegistry As New SoftinnRegistry.Registry
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
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
Public vQty As Boolean


Public Sub SetConnection(ConnObject As ADODB.Connection)
  Set CN = ConnObject
End Sub

Public Sub SetChar(c As Object)
  Set Char = c
End Sub

Public Function EStr(myString As String, Flag As Boolean) As String
   Dim i As Integer
   Dim myStr   As String
   For i = 1 To Len(myString)
       If Flag = True Then
           myStr = Chr(Asc(Mid(myString, i, i)) + 60 + Asc(i))
       ElseIf Flag = False Then
           myStr = Chr(Asc(Mid(myString, i, i)) - 60 - Asc(i))
       End If
       EStr = EStr & myStr
   Next
End Function

'Private Sub Command1_Click()
'    Dim sSecret     As String
'    sSecret = ToHexDump(RC4("a message here", "password"))
'    Debug.Print sSecret
'    Debug.Print RC4(FromHexDump(sSecret), "password")
'End Sub

Public Function RC4(sText As String, sKey As String) As String
    Dim baS(0 To 255) As Byte
    Dim baK(0 To 255) As Byte
    Dim bytSwap     As Byte
    Dim lI          As Long
    Dim lJ          As Long
    Dim lIdx        As Long

    For lIdx = 0 To 255
        baS(lIdx) = lIdx
        baK(lIdx) = Asc(Mid$(sKey, 1 + (lIdx Mod Len(sKey)), 1))
    Next
    For lI = 0 To 255
        lJ = (lJ + baS(lI) + baK(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
    Next
    lI = 0
    lJ = 0
    For lIdx = 1 To Len(sText)
        lI = (lI + 1) Mod 256
        lJ = (lJ + baS(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
        RC4 = RC4 & Chr$((pvXor(baS((CLng(baS(lI)) + baS(lJ)) Mod 256), Asc(Mid$(sText, lIdx, 1)))))
    Next
End Function

Private Function pvXor(ByVal lI As Long, ByVal lJ As Long) As Long
    If lI = lJ Then
        pvXor = lJ
    Else
        pvXor = lI Xor lJ
    End If
End Function

Public Function ToHexDump(sText As String) As String
    Dim lIdx As Long
    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex(Asc(Mid(sText, lIdx, 1))), 2)
    Next
End Function

Public Function FromHexDump(sText As String) As String
    Dim lIdx As Long
    For lIdx = 1 To Len(sText) Step 2
        FromHexDump = FromHexDump & Chr$(CLng("&H" & Mid(sText, lIdx, 2)))
    Next
End Function


