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
End Sub

Public Sub InitializeClass(objCN As ADODB.Connection)
   If ParaPass <> EncryptStr("›· ‹ÔÙÌÓ€ﬂÿ‡", False) Then Exit Sub
   Declarations.SetConnection objCN
End Sub

Public Sub AccountPayablesReport()
    RptAccountPayables.Show
End Sub

Public Sub AccountReceivableReport()
    RptAccountReceivables.Show
End Sub

Public Sub RoceverySheetReport()
   RptRoceverySheet.Show
End Sub

Public Sub BillWiseProfitReport()
    RptBillWiseProfitNew.Show
End Sub

Public Sub CashBookReport()
    RptCashBook.Show
End Sub

Public Sub CashFlowReport()
    RptCashFlow.Show
End Sub
Public Sub AgeingReport()
   RptAgeingReport.Show
End Sub

Public Sub LedgerReport()
    RptLedgerNew.Show
End Sub

Public Sub LedgerDetailReport()
    RptAccountLedgerDetail.Show
End Sub
Public Sub JVRegister()
    RptJVRegister.Show
End Sub
Public Sub CustomerLedgerReport()
    RptCustomerLedger.Show
End Sub

Public Sub TrialBalanceReport()
   RptTrialBalanceNew.ShowDetailed = True
   RptTrialBalanceNew.Show
End Sub

Public Sub PLStatementReport()
   RptPLStatement.Show
End Sub

Public Sub PLStatementOrgReport()
   RptPLStatementOrg.Show
End Sub

Public Sub ProfitReport()
   RptProfit.Show
End Sub

Public Sub DateWiseProfitReport()
    RptDateWiseProfitNew.Show
End Sub

Public Sub DateWiseProfitStoreWiseReport()
    RptDateWiseProfitStoreWise.Show
End Sub

Public Sub ProductWiseProfitReport()
    RptProductWiseProfitNew.Show
End Sub

Public Sub ProductWiseProfitStoreWiseReport()
    RptProductWiseProfitStoreWise.Show
End Sub

Public Sub AccountStatusReport()
    RptAccountStatus.Show
End Sub

Public Sub AccountSummaryReport()
    RptAccount.Show
End Sub

Public Sub DailyActivityReport()
    RptDailyActivity.Show
End Sub

Public Sub ExpenseReport()
    RptExpense.Show
End Sub

Public Sub BallaneSheetReport()
    RptBalanceSheet.Show
End Sub

Public Sub ProfitRegisterReport()
    RptProfitRegister.Show
End Sub

Public Sub RecoveryRegisterReport()
    RptRecoveryRegister.Show
End Sub
Public Sub ReceivedRegisterReport()
    RptReceivedRegister.Show
End Sub
Public Sub PaymentRegisterReport()
    RptPaymentRegister.Show
End Sub
Public Sub AccountsBalancesDiff()
    FrmAccountsBalancesDiff.Show
End Sub

