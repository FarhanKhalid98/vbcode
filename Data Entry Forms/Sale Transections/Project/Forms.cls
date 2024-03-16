VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Forms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub Initialize(Con As String, vPath As String)
    ParaPass = Con
    vTmp = vPath
    SetDllDirectory App.path
'    LoadLibrary App.path & "\libmysql.dll"
End Sub

Public Sub InitializeClass(objCN As ADODB.Connection)
   If ParaPass <> EncryptStr("«ŸÏÌÎÎÔÚ", False) Then Exit Sub
   Declarations.SetConnection objCN
   
   If cn.Execute("Select Value From sysindex Where SrNo = 5").RecordCount = 1 Then
      If EStr(cn.Execute("Select dbo.Value('5')").Fields(0).Value, False) = EStr("Ω”·›“‡ÿ‚È", False) Then Exit Sub 'for permament basis
   End If
   
   Dim vDate As String
   vDate = EStr(cn.Execute("Select dbo.Value('4')").Fields(0), False)
   If DateDiff("d", Format(Date, "MM/dd/yyyy"), vDate) < 1 Then '"03/10/2015"
      cn.Execute "update sysindexs set RegistryKey = 'InvType' where RegistryKey = 'InvTypeVisible'"
   End If
End Sub

Public Sub InitializeSecurity(objSec As UserSecurity.ClsUserSecurity)
    Declarations.SetSecurityReference objSec
End Sub

Public Sub InitializeChar(c As Object)
    Declarations.SetChar c
End Sub

Public Sub InitializeLogo(vStr As String)
    Declarations.SetLogo vStr
End Sub

Public Sub InitializeWaterMark(vStr As String)
    Declarations.SetWaterMark vStr
End Sub

'Public Sub RefreshRegistry()
'   ObjRegistry.RefreshRegistry
'End Sub

Public Sub SaleInvoiceForm()
    Dim a As New FrmSaleInvoiceDist
    a.Show
End Sub

Public Sub SaleInvoicePOSForm()
    Dim a As New FrmSaleInvoicePOS
    a.Show
End Sub

'Public Sub SaleInvoicePandGForm()
'    Dim a As New FrmSaleInvoicePandG
'    a.Show
'End Sub

'Public Sub ServiceInvoiceForm()
'    Dim a As New FrmServiceInvoice
'    a.Show
'End Sub

Public Sub SaleOrderForm()
    Dim a As New FrmSaleOrder
    a.Show
End Sub

Public Sub SaleOrderPOSForm()
    Dim a As New FrmSaleOrderPOS
    a.Show
End Sub

Public Sub SaleReturnInvoiceForm()
    Dim a As New FrmSaleReturnInvoice
    a.Show
End Sub

Public Sub SaleReturnInvoicePOSForm()
    Dim a As New FrmSaleReturnInvoicePOS
    a.Show
End Sub
Public Sub SaleInvoiceTouchForm()
    Dim a As New FrmSaleInvoiceTouch
    a.Show
End Sub

Public Sub RecoveryCustomerWiseForm()
    Dim a As New FrmRecoveryCustomerWise
    a.Show
End Sub

Public Sub RecoveryInvoiceWiseForm()
    Dim a As New FrmRecoveryInvoiceWise
    a.Show
End Sub

Public Sub ReplacementInvoiceForm()
    Dim a As New FrmReplacementInvoice
    a.Show
End Sub

Public Sub CustomOrderForm()
    Dim a As New FrmCustomOrderBooking
    a.Show
End Sub

Public Sub CustomOrderDeliveryForm()
    Dim a As New FrmCustomOrderDelivery
    a.Show
End Sub

Public Sub CustomOrderReturnForm()
    Dim a As New FrmCustomOrderReturn
    a.Show
End Sub
Public Sub ReceivingInvoicePrint()
    Dim a As New FrmReceivingInvoicePrint
    a.Show
End Sub
Public Sub PostSaleForm()
    Dim a As New FrmPostSelectedSales
    a.Show
End Sub

Public Sub Bin(objBin As String)
   vBinDataBase = objBin
End Sub

