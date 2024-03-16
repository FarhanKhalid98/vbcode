Attribute VB_Name = "Declarations"
Option Explicit
Public ObjRegistry As New SoftinnRegistry.Registry
Public ObjUserSecurity As UserSecurity.ClsUserSecurity
Public Declare Function SetWindowText Lib "User32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpstring As String) As Long
Public Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "User32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public CN As ADODB.Connection
Public ParaCustID, ParaCustName As String
Public ConnString  As String
Public ParaOutID As String
Public ParaPass As String
Public LogoPath As String
Public WaterMarkPath As String
Public vUser As Byte
Public vSessionID As Byte
'Public Char As Object
Public vBm As Variant

Public Sub SetConnection(ConnObject As ADODB.Connection)
   Set CN = ConnObject
   ConnString = CN.ConnectionString
'   CN.CommandTimeout = 600
'   CN.ConnectionTimeout = 600
End Sub

Public Sub SetChar(c As Object)
'   Set Char = c
End Sub

Public Sub SetLogo(vStr As String)
   LogoPath = vStr
End Sub

Public Sub SetWaterMark(vStr As String)
   WaterMarkPath = vStr
End Sub

Public Sub SetSecurityReference(SecObject As UserSecurity.ClsUserSecurity)
   Set ObjUserSecurity = SecObject
   vUser = ObjUserSecurity.UserNo
   vSessionID = ObjUserSecurity.SessionID
End Sub



Public Sub ActivityLog(FormType As String, Mode As EntryMode, Optional Key1 As Long = 0, Optional Key2 As Date = "01-01-1900", Optional Key3 As String = "")
   Dim vSQL As String
   vSQL = "Exec ProdActivityLog '" & FormType & "'," & ObjUserSecurity.UserNo & "," & Mode & "," & Key1 & ",'" & Key2 & "','" & Key3 & "'"
   'vSQL = "INSERT into ActivityLog(userno,FormType,EntryDate,Description,isnew,isedit,isdelete) values(" & ObjUserSecurity.UserNo & ",'" & FormType & "',getdate(),'" & Desc & "'," & IIf(Mode = eAdd, 1, 0) & "," & IIf(Mode = eEdit, 1, 0) & "," & IIf(Mode = eDelete, 1, 0) & ")"
   CN.Execute vSQL
End Sub

Public Sub ActivityLogBin(vTempID As String, vFormNo As Integer, vActionNO As Integer, vid As Long, vDate As Date, vDesc As String)
   Dim vSQL As String
   If ObjRegistry.UseBin = True Then
      vSQL = "insert into " & vBinDataBase & ".dbo.ActivityLogBin(ActivityDate,ActionNo,userno,FormNo, TempID, TransactionID,TransactionDate,TransactionInfo) values(getdate()," & vActionNO & "," & ObjUserSecurity.UserNo & ",'" & vFormNo & "'," & IIf(vTempID = "", "Null", "'" & vTempID & "'") & "," & IIf(vid = 0, "Null", vid) & "," & IIf(vDate = "01/01/1900", "Null", "'" & vDate & "'") & ",'" & vDesc & "')"
      CN.Execute vSQL
   End If
End Sub

Public Sub DeleteTempActivityLogBin(vTempID As String)
   Dim vSQL As String
   If ObjRegistry.UseBin = True Then
      vSQL = "Delete " & vBinDataBase & ".dbo.Activitylogbin Where ActionNo = " & eAddTempRecord & " and TempID = '" & vTempID & "'"
      CN.Execute (vSQL)
   End If
End Sub



