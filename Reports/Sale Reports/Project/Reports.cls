VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Initialize(Con As String, vPath As String)
    ParaPass = Con
    vTmp = vPath
End Sub

Public Sub InitializeClass(objCN As ADODB.Connection)
   If ParaPass <> EncryptStr("›· ‹ÔÙÌÓ€ﬂÿ‡", False) Then Exit Sub
   Declarations.SetConnection objCN
End Sub

Public Sub InitializeSecurity(objSec As UserSecurity.ClsUserSecurity)
    Declarations.SetSecurityReference objSec
End Sub


Public Sub GroupWiseSaleDetailReport()
'    RptGroupWiseSaleDetail.Show
End Sub

Public Sub GroupWiseSaleSummaryReport()
'    RptGroupWiseSaleSummary.Show
End Sub

Public Sub GroupWiseSaleReturnDetailReport()
'    RptGroupWiseSaleReturnDetail.Show
End Sub

Public Sub GroupWiseSaleReturnSummaryReport()
'    RptGroupWiseSaleReturnSummary.Show
End Sub

Public Sub UserWiseSaleDetailReport()
'    RptUserWiseSaleDetail.Show
End Sub

Public Sub UserWiseSaleSummaryReport()
'    RptUserWiseSaleSummary.Show
End Sub

Public Sub UserWiseSaleReturnDetailReport()
'   RptUserWiseSaleReturnDetail.Show
End Sub

Public Sub UserWiseSaleReturnSummaryReport()
'    RptUserWiseSaleReturnSummary.Show
End Sub

Public Sub CustomerWiseSaleDetailReport()
'    RptCustomerWiseSaleDetail.Show
End Sub

Public Sub CustomerWiseSaleSummaryReport()
'    RptCustomerWiseSaleSummary.Show
End Sub

Public Sub CustomerWiseSaleReturnDetailReport()
'    RptCustomerWiseSaleReturnDetail.Show
End Sub

Public Sub CustomerWiseSaleReturnSummaryReport()
'    RptCustomerWiseSaleReturnSummary.Show
End Sub

Public Sub SaleDetailCustomerWiseReport()
    RptSaleDetailCustomerWise.Show
End Sub

Public Sub SaleDetailDateWiseReport()
    RptSaleDetailDateWise.Show
End Sub

Public Sub SaleRegisterReport()
    RptSaleRegister.Show
End Sub

Public Sub SaleRegisterSerailWiseReport()
    RptSaleRegisterSerailWise.Show
End Sub

Public Sub SaleDetailStoreWiseReport()
    RptSaleDetailStoreWise.Show
End Sub

Public Sub ProductLedgerReport()
    RptProductLedger.Show
End Sub

Public Sub HotAndColdProductsReport()
    RptHotAndColdProducts.Show
End Sub

Public Sub HotAndColdCustomersReport()
    RptHotAndColdCustomers.Show
End Sub

Public Sub DeadProductsReport()
    RptDeadProducts.Show
End Sub

Public Sub DateWiseSaleExpenseReport()
    RptDateWiseSaleExpense.Show
End Sub

Public Sub SaleExpenseStoreWiseReport()
    RptSaleExpenseStoreWise.Show
End Sub

Public Sub CustomOrderBalanceReport()
    RptCustomOrderBalance.Show
End Sub

Public Sub ProductWiseSaleStoreWiseReport()
    RptProductWiseSaleStoreWise.Show
End Sub

Public Sub CustomerSaleBillsReport()
    RptCustomerSaleBills.Show
End Sub

Public Sub SaleOrderRegisterReport()
    RptSaleOrderRegister.Show
End Sub

Public Sub SalestaxRegisterReport()
    RptSalestaxRegister.Show
End Sub
