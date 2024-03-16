Attribute VB_Name = "GetConnected"
Option Explicit
   Private Const LOCALE_SSHORTDATE = &H1F
   Private Const LOCALE_SDATE = &H1D

   Private Const WM_SETTINGCHANGE = &H1A
   'same as the old WM_WININICHANGE
   Private Const HWND_BROADCAST = &HFFFF&
   Private Declare Function GetLocaleInfo Lib "kernel32" Alias _
      "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As _
      Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
   Private Declare Function SetLocaleInfo Lib "kernel32" Alias _
       "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As _
       Long, ByVal lpLCData As String) As Boolean
   Private Declare Function PostMessage Lib "user32" Alias _
       "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
       ByVal wParam As Long, ByVal lParam As Long) As Long
   Private Declare Function GetSystemDefaultLCID Lib "kernel32" _
          () As Long
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
Public ObjAccounts As New AccountForms.Forms
Public ObjAccountReports As New AccountReports.Forms
Public ObjBank As New BankForms.Forms
Public ObjBankReports As New BankReports.Reports
Public ObjDefinition As New DefinitionForms.Forms
Public ObjListReport As New ListReports.Reports
Public ObjProduction As New ProductionForms.Forms
Public ObjProductionReport As New ProductionReports.Reports
Public ObjPurchase As New PurchaseForms.Forms
Public ObjPurchaseReports As New PurchaseReports.Reports
Public ObjOthers As New OtherForms.Forms
Public ObjOtherReports As New OtherReports.Reports
Public ObjSale As New SaleForms.Forms
Public ObjSaleReports As New SaleReports.Reports
Public ObjStock As New StockForms.Forms
Public ObjStockReports As New StockReports.Reports
Public ObjUserSecurity As New UserSecurity.ClsUserSecurity
'Public Char1 As Object
'Public Animation As Variant
'Public AgentLoaded As Boolean
'Public CharName As String
Public LogoName As String
Public WaterMarkName As String
Public User1 As Byte
Public objFSO As New Scripting.FileSystemObject
Public vConnStr As String
Public vStr As String, vAddress As String, vLogo As String
Public NotVisibility As Boolean
Public vEncryptionString As String
Dim TempCon As New ADODB.Connection

'Public Sub CharPopMnu()
''    Char1.Commands.Visible = True
''    For Each Animation In Desktop.Agent1.Characters("Char1").AnimationNames
''        Char1.Commands.Add Animation, Animation, , True, True
''    Next
'End Sub

Public Sub OpenConnection()
   On Error GoTo ErrorHandler
   Dim vConnTime As String
   Dim vConnString As String
'   vConnString = "Provider=SQLOLEDB.1;User ID=softinn;Password=soft;Initial Catalog=" & vConnStr
   vConnString = "Provider=SQLOLEDB.1;User ID=sa;Initial Catalog=" & vConnStr
'    vConnString = "Provider=SQLOLEDB.1;Trusted_Connection=true;Integrated Security=False;Persist Security Info=False;User ID=softinn;Initial Catalog=" & vConnStr
   If TempCon.State = adStateOpen Then TempCon.Close
   TempCon.Open vConnString
'   With TempCon.Execute("Select ConnectionTimeOut from Registry")
      vConnTime = 0 'IIf(IsNull(!ConnectionTimeout), 60, !ConnectionTimeout)
'   End With
   If TempCon.State = adStateOpen Then TempCon.Close
   If CN.State = adStateOpen Then CN.Close
   CN.ConnectionTimeout = Val(vConnTime)
   CN.Open vConnString
   
   
   
   CN.CursorLocation = adUseClient
'   CN.CommandTimeout = Val(vConnTime)
   If CNR.State = adStateOpen Then CNR.Close
   CNR.ConnectionTimeout = 0
   CNR.Open vConnString
   CNR.CursorLocation = adUseClient
   CNR.CommandTimeout = 0
   
   
  
   Exit Sub
ErrorHandler:
   If Err.Number = -2147217843 Then
      vConnString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Initial Catalog=" & vConnStr
'      CN.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Initial Catalog=" & vConnStr
      Resume Next
   End If
   If Err.Number = -2147467259 Then
'      vConnString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Initial Catalog=" & vConnStr
      If objFSO.FileExists(App.Path & "\database\SuperSoftv1_Data.mdf") Then
         If TempCon.State = adStateOpen Then TempCon.Close
         TempCon.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Initial Catalog=master;data source=.;"
         TempCon.Execute "EXEC sp_attach_db 'SuperSoftv1','" & App.Path & "\DataBase\SuperSoftv1_Data.MDF','" & App.Path & "\DataBase\SuperSoftv1_Log.LDF'"
         Resume Next
      End If
   End If
   Call ShowErrorMessage
End Sub

Public Sub Main()
   On Error GoTo ErrorHandler
   'Dim objIniFile As File
   Splash.Show
   Splash.LblProgress.Value = 5
   Splash.LblStatus.Caption = "Connecting with the Database..."
   DoEvents
   Set CN = New ADODB.Connection
   Set CNR = New ADODB.Connection
   If Not objFSO.FileExists(App.Path & "\Config.ini") Then
      FrmConfig.Show vbModal
   End If
   vStr = EncryptStr("›· ‹ÔÙÌÓ€ﬂÿ‡", False)
   vTmp = App.Path & "\" 'objFSO.GetSpecialFolder(2)
   Dim vString As String
   
'   Open App.Path & "\Config.ini" For Append As #1
'   Print #1, "farhan khalid"
'   Close #1
   
   Open App.Path & "\Config.ini" For Input As #1
   Line Input #1, vConnStr
   
'   Do Until EOF(1)
'      Line Input #1, vString
'      Debug.Print vString
'   Loop
   Close #1
   
   Open App.Path & "\Bin.ini" For Input As #1
   Line Input #1, vBinDataBase
   Close #1
   
   Call OpenConnection
   Splash.LblProgress.Value = 20
   Splash.LblStatus.Caption = "Connection established with the Database..."
   DoEvents
   ''''''''' security check ''''''''''
   'MsgBox EncryptStr(MacAddress(), True) & vbCrLf & CN.Execute("select * from Court").Fields(0).Value
'   Dim a  As String, b As String
'   'a = EncryptStr("00065B4109DA", True)
'   a = CN.Execute("select * from Court").Fields(0).Value
'   vAddress = MacAddress()
'   Dim B1 As String, vFlag As Boolean
'   vFlag = False
'   NotVisibility = False
'   While (vAddress <> "" And vFlag = False)
'      B1 = Left(vAddress, InStr(1, vAddress, " ") - 1)
'      b = EncryptStr(B1, True)
'      vFlag = IIf(a = b, True, False)
'      vAddress = Replace(vAddress, B1 & " ", "")
'   Wend
   Dim a  As String, b As String, vAddress1 As String, vPermament As Boolean
   vAddress = MacAddress()
   vAddress1 = vAddress
   Dim B1 As String, vFlag As Boolean
   vFlag = False
   vPermament = Abs(CN.Execute("Select Value From sysindex Where SrNo = 5").RecordCount)
   If vPermament = False Then
      Call CheckDateFormat
      NotVisibility = Not FunSecurityCheck
   End If
   While (vAddress <> "" And vFlag = False)
      B1 = Left(vAddress, InStr(1, vAddress, " ") - 1)
      b = B1
      With CN.Execute("Select * from Court where SID = '" & EStr(ToHexDump(RC4(B1, EStr("‡›’‰⁄‡·", False))), True) & "'")
         If .RecordCount > 0 Then
            a = RC4(FromHexDump(EStr(!SID, False)), EStr("‡›’‰⁄‡·", False))
            'convert 'ToHexDump(RC4("","password"))
            'deconvert RC4(FromHexDump(sSecret), "password")
            If RC4(FromHexDump(EStr(!Type, False)), EStr("‡›’‰⁄‡·", False)) = RC4(FromHexDump(EStr("••†∂§§¨´¨¶¶ü", False)), EStr("‡›’‰⁄‡·", False)) Then 'Ser
               NotVisibility = False
            ElseIf RC4(FromHexDump(EStr(!Type, False)), EStr("‡›’‰⁄‡·", False)) = RC4(FromHexDump(EStr("•†°±°∑", False)), EStr("‡›’‰⁄‡·", False)) Then  'VN
               NotVisibility = False
            ElseIf RC4(FromHexDump(EStr(!Type, False)), EStr("‡›’‰⁄‡·", False)) = RC4(FromHexDump(EStr("¶¶†≤§¢¨©¨†¶ù", False)), EStr("‡›’‰⁄‡·", False)) Then  'Lap
               NotVisibility = True
            End If
         End If
         If vPermament = True Then
            NotVisibility = False
         End If
         .Close
      End With
      vFlag = IIf(a = b, True, False)
      vAddress = Replace(vAddress, B1 & " ", "")
   Wend
'   If vFlag = False Then
'      MsgBox "SuperSoft could not recognize you as a Legal User of this Copy. Please Contact the Vendor." + B1, vbCritical + vbOKOnly, "Soft Inn" 'vaddress = " & vAddress1 & " a = " & a
'      End
'   End If
'   '''''''''''
   SubInitilizeLogo
   SubInitilizeWaterMark
'   ''''''''''''''
   Splash.LblProgress.Value = 30
   Splash.LblStatus.Caption = "Initializing the Account Forms..."
   ObjAccounts.Initialize vStr, vTmp
   ObjAccounts.InitializeClass CN
   ObjAccounts.Bin vBinDataBase
   
   DoEvents
   ''''''''''''''''''''
   Splash.LblProgress.Value = 35
   Splash.LblStatus.Caption = "Initializing the Bank Forms..."
   ObjBank.Initialize vStr, vTmp
   ObjBank.InitializeClass CN
   DoEvents
   ''''''''''''''''''''
   Splash.LblProgress.Value = 40
   Splash.LblStatus.Caption = "Initializing the Bank Reports..."
   ObjBankReports.Initialize vStr, vTmp
   ObjBankReports.InitializeClass CNR
   DoEvents
   '''''''''''''''''''
   Splash.LblProgress.Value = 45
   Splash.LblStatus.Caption = "Initializing the Definition Forms..."
   ObjDefinition.Initialize vStr, vTmp
   ObjDefinition.InitializeClass CN
   ObjDefinition.Bin vBinDataBase
   DoEvents
   '''''''''''''''''''''
   Splash.LblProgress.Value = 50
   Splash.LblStatus.Caption = "Initializing the Purchase Reports..."
   ObjPurchaseReports.Initialize vStr, vTmp
   ObjPurchaseReports.InitializeClass CNR
   ObjAccounts.Bin vBinDataBase
   DoEvents
   '''''''''''''''''''''
   Splash.LblProgress.Value = 55
   Splash.LblStatus.Caption = "Initializing the Sales Report..."
   ObjSaleReports.Initialize vStr, vTmp
   ObjSaleReports.InitializeClass CNR
   DoEvents
   '''''''''''''''''''''
   Splash.LblProgress.Value = 60
   Splash.LblStatus.Caption = "Initializing the User Security..."
   ObjUserSecurity.Bin vBinDataBase
   ObjUserSecurity.Initialize vStr, vTmp
   ObjUserSecurity.InitializeClass CN
   
   DoEvents
   ''''''''''''''''''''''''
   Splash.LblProgress.Value = 65
   Splash.LblStatus.Caption = "Initializing the Registry..."
   ObjRegistry.Initialize vStr, vTmp
   ObjRegistry.InitializeClass CN
   ObjRegistry.InitializeStatus NotVisibility
   ObjRegistry.RefreshRegistry
   ''''''''''''''''''''''''
   Splash.LblProgress.Value = 65
   Splash.LblStatus.Caption = "Initializing the Purchase..."
   ObjPurchase.Initialize vStr, vTmp
   ObjPurchase.InitializeClass CN
   ObjPurchase.Bin vBinDataBase
   DoEvents
   ''''''''''''''''''''''''
   Splash.LblProgress.Value = 70
   Splash.LblStatus.Caption = "Initializing the Sale..."
   ObjSale.Initialize EncryptStr("«ŸÏÌÎÎÔÚ", False), vTmp
   ObjSale.InitializeClass CN
   ObjSale.Bin vBinDataBase
   ObjSale.InitializeLogo LogoName
   ObjSale.InitializeWaterMark WaterMarkName
   SubInitilizeWaterMark
   DoEvents
   ''''''''''''''''''''''''
   Splash.LblProgress.Value = 75
   Splash.LblStatus.Caption = "Initializing the Production..."
   ObjProduction.Initialize vStr, vTmp
   ObjProduction.InitializeClass CN
   DoEvents
   ''''''''''''''''''''''''
   Splash.LblProgress.Value = 80
   Splash.LblStatus.Caption = "Initializing the Production Reports..."
   ObjProductionReport.Initialize vStr, vTmp
   ObjProductionReport.InitializeClass CNR
   DoEvents
   ''''''''''''''''''''''''
   Splash.LblProgress.Value = 85
   Splash.LblStatus.Caption = "Initializing the Stock..."
   ObjStock.Initialize vStr, vTmp
   ObjStock.InitializeClass CN
   ObjStock.Bin vBinDataBase
   DoEvents
   ''''''''''''''''''''''''
   Splash.LblProgress.Value = 90
   Splash.LblStatus.Caption = "Initializing the Stock Report..."
   ObjStockReports.Initialize vStr, vTmp
   ObjStockReports.InitializeClass CNR
   DoEvents
   ''''''''''''''''''''
   Splash.LblProgress.Value = 95
   Splash.LblStatus.Caption = "Initializing the List Reports..."
   ObjListReport.Initialize vStr, vTmp
   ObjListReport.InitializeClass CNR
   DoEvents
   ''''''''''''''''''''
   Splash.LblProgress.Value = 100
   Splash.LblStatus.Caption = "Initializing the Account Reports..."
   ObjAccountReports.Initialize vStr, vTmp
   ObjAccountReports.InitializeClass CNR
   DoEvents
   ''''''''''''''''''''
   Splash.LblProgress.Value = 100
   Splash.LblStatus.Caption = "Initializing the Other ..."
   ObjOthers.Initialize vStr, vTmp
   ObjOthers.InitializeClass CN
   ObjOthers.Bin vBinDataBase
   DoEvents
   ''''''''''''''''''''
   Splash.LblProgress.Value = 100
   Splash.LblStatus.Caption = "Initializing the Other Reports..."
   ObjOtherReports.Initialize vStr, vTmp
   ObjOtherReports.InitializeClass CNR
   DoEvents
   ''''''''''''''''''''''''
   Unload Splash
   Desktop.Show
   Exit Sub
ErrorHandler:
'   If Err.Number = -2147217843 Then
'      CN.Open "Provider=SQLOLEDB;Integrated Security=SSPI;Initial Catalog=" & vConnStr
'      Resume Next
'   End If
   Call ShowErrorMessage
   End
End Sub

Private Sub CheckDateFormat()
   On Error GoTo ErrorHandler
   Dim Symbol As String
   Dim iRet1 As Long
   Dim iRet2 As Long
   Dim lpLCDataVar As String
   Dim Pos As Integer
   Dim Locale As Long
   
   Locale = GetUserDefaultLCID()
   'LOCALE_SSHORTDATE
   'LOCALE_SDATE
   iRet1 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, lpLCDataVar, 0)
   Symbol = String$(iRet1, 0)
   
   iRet2 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol, iRet1)
   Pos = InStr(Symbol, Chr$(0))
   If Pos > 0 Then
      Symbol = Left$(Symbol, Pos - 1)
      If Symbol <> "MM/dd/yyyy" Then
         Dim dwLCID As Long
         dwLCID = GetSystemDefaultLCID()
         If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, "MM/dd/yyyy") = False Then
            MsgBox "Failed"
            Exit Sub
         End If
         PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
      End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
                  
Sub SubInitilizeLogo()
   On Error GoTo ErrorHandler
   strsql = "SELECT * FROM CompanyLogo"
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
   If Rs.RecordCount = 0 Then LogoName = "": Exit Sub
   DataFile = 1
    
   Open App.Path & "\SI.Bmp" For Binary Access Write As DataFile
      Fl = Rs!pic.ActualSize ' Length of data in file
      If Fl = 0 Then Close DataFile: Exit Sub
      Chunks = Fl \ ChunkSize
      Fragment = Fl Mod ChunkSize
      ReDim Chunk(Fragment)
      Chunk() = Rs!pic.GetChunk(Fragment)
      Put DataFile, , Chunk()
      For i = 1 To Chunks
         ReDim Buffer(ChunkSize)
         Chunk() = Rs!pic.GetChunk(ChunkSize)
         Put DataFile, , Chunk()
      Next i
   Close DataFile
   LogoName = App.Path & "\SI.Bmp"
   Rs.Close
   Set Rs = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Sub SubInitilizeWaterMark()
   On Error GoTo ErrorHandler
   strsql = "SELECT * FROM WaterMark"
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
   If Rs.RecordCount = 0 Then WaterMarkName = "": Exit Sub
   DataFile = 1
    
   Open "C:\WaterMark.Bmp" For Binary Access Write As DataFile
      Fl = Rs!pic.ActualSize ' Length of data in file
      If Fl = 0 Then Close DataFile: Exit Sub
      Chunks = Fl \ ChunkSize
      Fragment = Fl Mod ChunkSize
      ReDim Chunk(Fragment)
      Chunk() = Rs!pic.GetChunk(Fragment)
      Put DataFile, , Chunk()
      For i = 1 To Chunks
         ReDim Buffer(ChunkSize)
         Chunk() = Rs!pic.GetChunk(ChunkSize)
         Put DataFile, , Chunk()
      Next i
   Close DataFile
   WaterMarkName = "C:\WaterMark.Bmp"
   Rs.Close
   Set Rs = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Public Function DBExists() As Boolean
    On Error GoTo ErrorHandler
    With CN.Execute("Exec SP_Databases")
        .Find "Database_Name = 'SuperSoftv1'", , adSearchForward, 1
        DBExists = Not .EOF
    End With
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

