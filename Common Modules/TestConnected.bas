Attribute VB_Name = "TestConnected"
Option Explicit
'Public ObjDefinition As New DefintionForms.Forms
'Public ObjTransection As New TransectionForms.Forms
'Public ObjAccounts As New AccountForms.Forms
'Public ObjAccountReports As New AccountReports.Forms
'Public ObjTransectionReports As New TransectionReports.Forms
'Public ObjListReports As New ListReports.Forms
Public User1 As Byte

Public Sub Main()
   Dim objFSO As New Scripting.FileSystemObject
   Dim objIniFile As File
   
'   Splash.Show
'   Splash.LblProgress.Value = 5
'   Splash.LblStatus.Caption = "Connecting with the Database..."
'   DoEvents
   
   Dim vConnStr As String
   Set CN = New ADODB.Connection
   Open App.Path & "\Config.ini" For Input As #1
   Input #1, vConnStr
   Close #1
   'CN.Open "Driver=SQL Server;" & vConnStr & "uid=sa;"
   CN.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & vConnStr
   CN.CursorLocation = adUseClient
   'DefProducts.Show
   User = 1
   'FrmSaleInvoice.Show
   SchProduct.Show
   'FrmPurchaseInvoice.Show
   'FrmPurchaseReturnInvoice.Show
   'FrmOpeningStock.Show1
   'Form1.Show
       
'   Splash.LblProgress.Value = 25
'   Splash.LblStatus.Caption = "Connection established with the Database..."
'   DoEvents
'   '''''''''''''''''''
'   Splash.LblProgress.Value = 35
'   Splash.LblStatus.Caption = "Initializing the Definition Forms..."
'   ObjDefinition.InitializeClass CN
'   DoEvents
'   '''''''''''''''''''''
'   Splash.LblProgress.Value = 45
'   Splash.LblStatus.Caption = "Initializing the Transections..."
'   ObjTransection.InitializeClass CN
'   DoEvents
'   '''''''''''''''''''''
'   Splash.LblProgress.Value = 55
'   Splash.LblStatus.Caption = "Initializing the Account Forms..."
'   ObjAccounts.InitializeClass CN
'   DoEvents
'   '''''''''''''''''''''
'   Splash.LblProgress.Value = 65
'   Splash.LblStatus.Caption = "Initializing the Account Reports..."
'   ObjAccountReports.InitializeClass CN
'   DoEvents
'   '''''''''''''''''''''
'   Splash.LblProgress.Value = 75
'   Splash.LblStatus.Caption = "Initializing the Transection Reports..."
'   ObjTransectionReports.InitializeClass CN
'   DoEvents
'   '''''''''''''''''''''
'   Splash.LblProgress.Value = 85
'   Splash.LblStatus.Caption = "Initializing the List Reports..."
'   ObjListReports.InitializeClass CN
'   DoEvents
'   ''''''''''''''''''''''''
'   Unload Splash
'   Desktop.Show
End Sub

