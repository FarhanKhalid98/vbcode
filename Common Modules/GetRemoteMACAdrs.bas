Attribute VB_Name = "GetRemoteMACAdrs"
 Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
'Determining a Local or Remote MAC Address via SendARP
'Distributor Okan CELEN
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const NO_ERROR = 0

Private Declare Function inet_addr Lib "wsock32.dll" _
  (ByVal s As String) As Long

Private Declare Function SendARP Lib "iphlpapi.dll" _
  (ByVal DestIP As Long, _
   ByVal SrcIP As Long, _
   pMacAddr As Long, _
   PhyAddrLen As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (dst As Any, _
   src As Any, _
   ByVal bcount As Long)

Public Function GetRemoteMACAddress(ByVal sRemoteIP As String) As Boolean
   Dim dwRemoteIP As Long
   Dim pMacAddr As Long
   Dim bpMacAddr() As Byte
   Dim PhyAddrLen As Long
   Dim cnt As Long
   Dim tmp As String
   Dim tmpStr As String
  'convert the string IP into
  'an unsigned long value containing
  'a suitable binary representation
  'of the Internet address given
   dwRemoteIP = inet_addr(sRemoteIP)
   'Text2.Text = dwRemoteIP
   'GetRemoteMACAddress = True
   'Exit Function
   If dwRemoteIP <> 0 Then
     'set PhyAddrLen to 6
      PhyAddrLen = 6
     'retrieve the remote MAC address
      If SendARP(dwRemoteIP, 0&, pMacAddr, PhyAddrLen) = NO_ERROR Then
         If pMacAddr <> 0 And PhyAddrLen <> 0 Then
           'returned value is a long pointer
           'to the mac address, so copy data
           'to a byte array
            ReDim bpMacAddr(0 To PhyAddrLen - 1)
            CopyMemory bpMacAddr(0), pMacAddr, ByVal PhyAddrLen
           'loop through array to build string
            For cnt = 0 To PhyAddrLen - 1
               If bpMacAddr(cnt) = 0 Then
                  tmp = "00"
               Else
                  tmp = Hex$(bpMacAddr(cnt))
               End If
               If (Len(tmp) < 2) Then tmp = "0" & tmp
               tmpStr = tmpStr & tmp
               'If (cnt < 5) Then tmpStr = tmpStr & "-"
            Next
            Address = tmpStr & " "
            GetRemoteMACAddress = True
            Exit Function
         End If
      End If  'SendARP
   End If  'dwRemoteIP
   GetRemoteMACAddress = False
End Function
