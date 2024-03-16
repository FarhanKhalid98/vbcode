VERSION 5.00
Begin VB.Form DesktopReport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   11220
   ClientLeft      =   0
   ClientTop       =   675
   ClientWidth     =   15360
   Icon            =   "DesktopReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "DesktopReport.frx":0ECA
   ScaleHeight     =   11220
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image ImgDesktopButton 
      Height          =   705
      Index           =   0
      Left            =   1050
      Tag             =   "MniSaleInvoice"
      Top             =   2580
      Width           =   1545
   End
   Begin VB.Image ImgStart 
      Height          =   405
      Left            =   0
      Top             =   8580
      Width           =   1050
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11640
      Top             =   90
      Width           =   315
   End
   Begin VB.Image ImgMin 
      Height          =   315
      Left            =   11295
      Top             =   105
      Width           =   315
   End
   Begin VB.Menu MnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnu100 
         Caption         =   ""
      End
   End
   Begin VB.Menu MnuAccounts 
      Caption         =   "Accounts Reports"
      Tag             =   "1MnuReport"
      Begin VB.Menu MniAccountReceivable 
         Caption         =   "A/C Receivable"
         Tag             =   "2MnuAccounts"
      End
      Begin VB.Menu MniAccountPayable 
         Caption         =   "A/C Payable "
         Tag             =   "2MnuAccounts"
      End
      Begin VB.Menu MniReceivedRegister 
         Caption         =   "Received Register"
      End
      Begin VB.Menu MniPaymentRegister 
         Caption         =   "Payment Register"
      End
      Begin VB.Menu MniJVRegister 
         Caption         =   "JV Register"
      End
      Begin VB.Menu MniAgeingReport 
         Caption         =   "Ageing Report"
      End
      Begin VB.Menu MniRecoverySheet 
         Caption         =   "Recovery Sheet"
      End
      Begin VB.Menu MniRecoveryRegister 
         Caption         =   "Recovery Register"
      End
      Begin VB.Menu MniCashBook 
         Caption         =   "Cash Book"
         Tag             =   "2MnuAccounts"
      End
      Begin VB.Menu MniCashFlow 
         Caption         =   "Cash Flow"
      End
      Begin VB.Menu MniDailyActivity 
         Caption         =   "Daily Activity"
      End
      Begin VB.Menu MniLedgerDetail 
         Caption         =   "Ledger Detail"
      End
      Begin VB.Menu MniExpenseReport 
         Caption         =   "Expense Report"
      End
      Begin VB.Menu MniLedger 
         Caption         =   "Ledger"
         Tag             =   "2MnuAccounts"
      End
      Begin VB.Menu MniTrialBalance 
         Caption         =   "Trial Balance"
         Tag             =   "2MnuAccounts"
      End
      Begin VB.Menu MniBalanceSheet 
         Caption         =   "Balance Sheet"
      End
      Begin VB.Menu MniProfit 
         Caption         =   "P/L Statement"
         Tag             =   "2MnuAccounts"
      End
      Begin VB.Menu MniDateWiseProfit 
         Caption         =   "Date Wise Profit"
      End
      Begin VB.Menu MniDateWiseProfitStoreWise 
         Caption         =   "Date Wise Profit Store Wise"
      End
      Begin VB.Menu MniProfitRegister 
         Caption         =   "Profit Register"
      End
      Begin VB.Menu MniAccountStatus 
         Caption         =   "Account Status"
      End
      Begin VB.Menu MniAccountsBalancesDiff 
         Caption         =   "Accounts Ballances Diff"
      End
      Begin VB.Menu Line5 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuBankReports 
      Caption         =   "Bank Reports"
      Tag             =   "1MnuReport"
      Begin VB.Menu MniBankCashDepositReport 
         Caption         =   "Bank Cash Deposit Report"
         Tag             =   "2MnuBankReports"
      End
      Begin VB.Menu MniBankChequeDepositReport 
         Caption         =   "Bank Cheque Deposit Report"
         Tag             =   "2MnuBankReports"
      End
      Begin VB.Menu MniBankChequeIssuanceReport 
         Caption         =   "Bank Cheque Issuance Report"
         Tag             =   "2MnuBankReports"
      End
      Begin VB.Menu MniBankChequeReceiveReport 
         Caption         =   "Bank Cheque Receive Report"
         Tag             =   "2MnuBankReports"
      End
      Begin VB.Menu Line6 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuListReport 
      Caption         =   "List Report"
      Tag             =   "1MnuReport"
      Begin VB.Menu MniGroupWiseProductPackInfoList 
         Caption         =   "Group Wise Product Packing Info List"
      End
      Begin VB.Menu MniMemberList 
         Caption         =   "Member List"
      End
      Begin VB.Menu MniCustomerList 
         Caption         =   "Customer List"
         Tag             =   "2MnuListReport"
      End
      Begin VB.Menu MniProductList 
         Caption         =   "Product List"
         Tag             =   "2MnuListReport"
      End
      Begin VB.Menu MniProductPriceList 
         Caption         =   "Product Price List"
      End
      Begin VB.Menu MniProductNotIncludedList 
         Caption         =   "Product Not Included List"
      End
      Begin VB.Menu MniDeadProductList 
         Caption         =   "Dead Product List"
      End
      Begin VB.Menu MniStoreList 
         Caption         =   "Store List"
         Tag             =   "2MnuListReport"
      End
      Begin VB.Menu MniCompanyList 
         Caption         =   "Company List"
      End
      Begin VB.Menu MniGroupList 
         Caption         =   "Group List"
      End
      Begin VB.Menu MniSubGroupList 
         Caption         =   "Sub Group List"
         Tag             =   "2MnuListReport"
      End
      Begin VB.Menu MniEmployeeList 
         Caption         =   "Employee List"
      End
      Begin VB.Menu MniVendorList 
         Caption         =   "Vendor List"
         Tag             =   "2MnuListReport"
      End
      Begin VB.Menu Line7 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuProductionReports 
      Caption         =   "Production Reports"
      Tag             =   "1MnuReport"
      Begin VB.Menu MniFinishedProduct 
         Caption         =   "Finished Product"
         Tag             =   "2MnuProductionReport"
      End
      Begin VB.Menu MniProductsUsed 
         Caption         =   "Products Used"
         Tag             =   "2MnuProductionReport"
      End
      Begin VB.Menu MniManufacturedProduct 
         Caption         =   "Manufactured Product"
      End
      Begin VB.Menu MniProductionRegister 
         Caption         =   "Production Register"
      End
      Begin VB.Menu MniDateWiseProductionInOut 
         Caption         =   "Date Wise Production In Out"
      End
      Begin VB.Menu mnu50 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuOtherReports 
      Caption         =   "Other Reports"
      Begin VB.Menu MniAdminClosingReport 
         Caption         =   "Admin Closing Report"
      End
      Begin VB.Menu MniHotandColdMembers 
         Caption         =   "Hot and Cold Members"
      End
      Begin VB.Menu MniSleepingMembers 
         Caption         =   "Sleeping Members"
      End
      Begin VB.Menu MniMeterReadingRegister 
         Caption         =   "Meter Reading Register"
      End
      Begin VB.Menu MniCustomerDemandRegister 
         Caption         =   "Customer Demand Register"
      End
      Begin VB.Menu MniEmployeeAttendanceReport 
         Caption         =   "Employee Attendance Report"
      End
      Begin VB.Menu MniSalaryEmployeeWise 
         Caption         =   "Salary Employee Wise"
      End
      Begin VB.Menu MniProductDifference 
         Caption         =   "Product Difference"
      End
      Begin VB.Menu Line25 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuPurchaseReports 
      Caption         =   "Purchase Reports"
      Tag             =   "2MnuProductionReport"
      Begin VB.Menu MniVenderPurchaseBills 
         Caption         =   "Vender Purchase Bills"
         Tag             =   "2MnuPurchaseReports"
      End
      Begin VB.Menu MniProductsNotInPurchase 
         Caption         =   "Products Not In Purchase"
      End
      Begin VB.Menu Line8 
         Caption         =   "-"
      End
      Begin VB.Menu MniPurchaseRegister 
         Caption         =   "Purchase Register"
      End
      Begin VB.Menu MniPurchaseRegisterSerailWise 
         Caption         =   "Purchase Register Serail Wise"
      End
      Begin VB.Menu MniPurchaseOrderRegister 
         Caption         =   "Purchase Order Register"
      End
      Begin VB.Menu mnu7 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuSalesReports 
      Caption         =   "Sales Reports"
      Begin VB.Menu MniCustomerSaleBills 
         Caption         =   "Customer Sale Bills"
      End
      Begin VB.Menu MniCustomOrderBalance 
         Caption         =   "Custom Order Balance"
      End
      Begin VB.Menu MniProductLedger 
         Caption         =   "Product Ledger"
      End
      Begin VB.Menu MniDateWiseSaleExpense 
         Caption         =   "Date Wise Sale Expense"
      End
      Begin VB.Menu MniSaleExpenseStoreWise 
         Caption         =   "Sale Expense Store Wise"
      End
      Begin VB.Menu MniHotandColdCustomers 
         Caption         =   "Hot and Cold Customers"
      End
      Begin VB.Menu MniHotandColdProducts 
         Caption         =   "Hot and Cold Products"
      End
      Begin VB.Menu MniDeadProducts 
         Caption         =   "Dead Products"
      End
      Begin VB.Menu mnu8 
         Caption         =   "-"
      End
      Begin VB.Menu MniSaleRegister 
         Caption         =   "Sale Register"
      End
      Begin VB.Menu MniSaleRegisterSerailWise 
         Caption         =   "Sale Register Serial Wise"
      End
      Begin VB.Menu MniSalesTaxRegister 
         Caption         =   "SalesTax Register"
      End
      Begin VB.Menu MniSaleOrderRegister 
         Caption         =   "Sale Order Register"
      End
      Begin VB.Menu mnu9 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuStockReports 
      Caption         =   "Stock Reports"
      Begin VB.Menu MniProductStockValue 
         Caption         =   "Product Stock Value"
      End
      Begin VB.Menu MniProductStockSummary 
         Caption         =   "Product Stock Summary"
      End
      Begin VB.Menu MniProductStockSummaryStoreWise 
         Caption         =   "Product Stock Summary Store Wise (Branch)"
      End
      Begin VB.Menu MniStockWastage 
         Caption         =   "Stock Wastage Report"
      End
      Begin VB.Menu MniStockAdjustmentValueReport 
         Caption         =   "Stock Adjustment Value"
      End
      Begin VB.Menu MniOpeningStockRegister 
         Caption         =   "Opening Stock Register"
      End
      Begin VB.Menu MniProductAnalysisReport 
         Caption         =   "Product Analysis Report"
      End
      Begin VB.Menu MniProductAnalysisStoreValueReport 
         Caption         =   "Product Analysis Store Value Report"
      End
      Begin VB.Menu mnu10 
         Caption         =   "-"
      End
      Begin VB.Menu MniStoreWiseStockTransferDetail 
         Caption         =   "Stock Transfer Detail Store Wise"
      End
      Begin VB.Menu MniDailyDemandList 
         Caption         =   "Daily Demand List"
      End
      Begin VB.Menu MniDemandList 
         Caption         =   "Demand List"
      End
      Begin VB.Menu MniDemandListStoreWise 
         Caption         =   "Demand List Store Wise"
      End
      Begin VB.Menu MniPriceVariationList 
         Caption         =   "Price Variation List"
      End
      Begin VB.Menu MniStockValueRegister 
         Caption         =   "Stock Value Register"
      End
      Begin VB.Menu MniStockRegister 
         Caption         =   "Stock Register"
      End
      Begin VB.Menu mnu11 
         Caption         =   "-"
      End
      Begin VB.Menu MniCurrentStockWastage 
         Caption         =   "Current Stock Wastage"
      End
      Begin VB.Menu MniCurrentStockExpiryValue 
         Caption         =   "Current Stock Expiry Value"
      End
      Begin VB.Menu MniBatchExpiryReport 
         Caption         =   "Batch Expiry Report"
      End
      Begin VB.Menu mnu12 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuBarcode 
      Caption         =   "BarCode"
      Begin VB.Menu MniMultipleBarCodePrinting 
         Caption         =   "Multiple BarCode Printing"
      End
      Begin VB.Menu MniMultiBarcodesDetail 
         Caption         =   "Multiple Barcode Detail Printing"
      End
      Begin VB.Menu MniSingleBarCodePrinting 
         Caption         =   "SingleBarCodePrinting"
      End
      Begin VB.Menu mnu13 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuLogOut 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "DesktopReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   SetWindowText Me.hwnd, "Desktop Report"
   ShowPicture Me, 1
   Call EnableMenus
End Sub
 
Public Sub EnableMenus()
    On Error GoTo ErrorHandler
    Dim mnu As Object, sql As String
    If ObjUserSecurity.IsAdministrator = False Then
      sql = "Select t.TaskKey From Tasks t inner join (Select Distinct TaskKey From UserTasks Where UserNo = " & ObjUserSecurity.UserNo & ") u on t.TaskKey = u.TaskKey where isLocked = 0"
    Else
      sql = "Select TaskKey From Tasks Where isLocked = 0"
    End If
               
    With CN.Execute(sql)
        For Each mnu In Me.Controls
            If TypeOf mnu Is Menu And UCase(mnu.Name) Like "MNI*" Then
               If mnu.Checked = False Then mnu.Visible = True
               .Filter = "TaskKey = '" & mnu.Name & "'"
               If .RecordCount = 0 Then mnu.Visible = False
            End If
        Next
    End With
    Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub
Private Sub MniAccountPayable_Click()
   If MniAccountPayable.Visible Then ObjAccountReports.AccountPayablesReport
End Sub

Private Sub MniAccountReceivable_Click()
   If MniAccountReceivable.Visible Then ObjAccountReports.AccountReceivableReport
End Sub

Private Sub MniAccountsDefaultSetting_Click()
   If MniAccountsDefaultSetting.Visible Then ObjAccounts.AccountsDefaultSettingForm
End Sub

Private Sub MniAccountsBalancesDiff_Click()
   If MniAccountsBalancesDiff.Visible Then ObjAccountReports.AccountsBalancesDiff
End Sub

Private Sub MniAccountStatus_Click()
   If MniAccountStatus.Visible Then ObjAccountReports.AccountStatusReport
End Sub
Private Sub MniAdminClosingReport_Click()
   If MniAdminClosingReport.Visible Then ObjOtherReports.AdminClosingReport
End Sub

Private Sub MniAgeingReport_Click()
   If MniAgeingReport.Visible Then ObjAccountReports.AgeingReport
End Sub

Private Sub MniBalanceSheet_Click()
   If MniBalanceSheet.Visible Then ObjAccountReports.BallaneSheetReport
End Sub

Private Sub MniBankCashDepositReport_Click()
   If MniBankCashDepositReport.Visible Then ObjBankReports.BankCashDepositReport
End Sub

Private Sub MniBankChequeDepositReport_Click()
   If MniBankChequeDepositReport.Visible Then ObjBankReports.BankChequeDepositReport
End Sub

Private Sub MniBankChequeIssuanceReport_Click()
   If MniBankChequeIssuanceReport.Visible Then ObjBankReports.BankChequeIssuanceReport
End Sub

Private Sub MniBankChequeReceiveReport_Click()
   If MniBankChequeReceiveReport.Visible Then ObjBankReports.BankChequeReceiveReport
End Sub
Private Sub MniBatchExpiryReport_Click()
   If MniBatchExpiryReport.Visible Then ObjStockReports.BatchExpiryReport
End Sub
Private Sub MniCompanyList_Click()
   If MniCompanyList.Visible Then ObjListReport.CompanyListReport
End Sub
Private Sub MniCustomerDemandRegister_Click()
   If MniCustomerDemandRegister.Visible Then ObjOtherReports.CustomerDemandRegisterReport
End Sub

Private Sub MniCustomerLedger_Click()
'   If MniCustomerLedger.Visible Then ObjAccountReports.CustomerLedgerReport
End Sub

Private Sub MniCustomerSaleBills_Click()
   If MniCustomerSaleBills.Visible Then ObjSaleReports.CustomerSaleBillsReport
End Sub
Private Sub MniCustomOrderBalance_Click()
   If MniCustomOrderBalance.Visible Then ObjSaleReports.CustomOrderBalanceReport
End Sub
Private Sub MniDailyActivity_Click()
   If MniDailyActivity.Visible Then ObjAccountReports.DailyActivityReport
End Sub

Private Sub MniDailyDemandList_Click()
    If MniDailyDemandList.Visible Then ObjStockReports.DailyDemandListCompanyWiseReport
End Sub

Private Sub MniDateWiseProductionInOut_Click()
   If MniDateWiseProductionInOut.Visible Then ObjProductionReport.DateWiseProductionInOutReport
End Sub

Private Sub MniDateWiseProfitStoreWise_Click()
   If MniDateWiseProfitStoreWise.Visible Then ObjAccountReports.DateWiseProfitStoreWiseReport
End Sub

Private Sub MniDeadProductList_Click()
   If MniDeadProductList.Visible Then ObjListReport.DeadProductListReport
End Sub

Private Sub MniDeadProducts_Click()
   If MniDeadProducts.Visible Then ObjSaleReports.DeadProductsReport
End Sub

Private Sub MniDemandListStoreWise_Click()
   If MniDemandListStoreWise.Visible Then ObjStockReports.DemandListStoreWiseReport
End Sub

Private Sub MniEmployeeAttendanceReport_Click()
   If MniEmployeeAttendanceReport.Visible Then ObjOtherReports.EmployeeAttendanceReport
End Sub

Private Sub MniEmployeeList_Click()
   If MniEmployeeList.Visible Then ObjListReport.EmployeeListReport
End Sub

Private Sub MniExpenseReport_Click()
    If MniDailyActivity.Visible Then ObjAccountReports.ExpenseReport
End Sub

Private Sub MniGroupList_Click()
   If MniGroupList.Visible Then ObjListReport.GroupListReport
End Sub

Private Sub MniGroupWiseProductPackInfoList_Click()
If MniGroupWiseProductPackInfoList.Visible Then ObjListReport.GroupWiseProductPackInfoListReport
End Sub

Private Sub MniHotandColdCustomers_Click()
   If MniHotandColdCustomers.Visible Then ObjSaleReports.HotAndColdCustomersReport
End Sub

Private Sub MniHotandColdMembers_Click()
   If MniHotandColdMembers.Visible Then ObjOtherReports.HotAndColdMembersReport
End Sub

Private Sub MniHotandColdProducts_Click()
   If MniHotandColdProducts.Visible Then ObjSaleReports.HotAndColdProductsReport
End Sub

Private Sub MniJVRegister_Click()
   If MniLedgerDetail.Visible Then ObjAccountReports.JVRegister
End Sub

Private Sub MniLedgerDetail_Click()
   If MniLedgerDetail.Visible Then ObjAccountReports.LedgerDetailReport
End Sub

Private Sub MniManufacturedProduct_Click()
   If MniManufacturedProduct.Visible Then ObjProductionReport.ManufacturedProductReport
End Sub

Private Sub MniMemberList_Click()
   If MniMemberList.Visible Then ObjListReport.MemberListReport
End Sub

Private Sub MniMeterReadingRegister_Click()
   If MniMeterReadingRegister.Visible Then ObjOtherReports.MeterReadingRegisterReport
End Sub

Private Sub MniMultiBarcodesDetail_Click()
   If MniMultiBarcodesDetail.Visible Then ObjPurchase.MultiBarcodesDetailForm
End Sub

Private Sub MniOpeningStockRegister_Click()
If MniOpeningStockRegister.Visible Then ObjStockReports.OpeningStockRegisterReport
End Sub

Private Sub MniPaymentRegister_Click()
   If MniPaymentRegister.Visible Then ObjAccountReports.PaymentRegisterReport
End Sub

Private Sub MniPriceVariationList_Click()
   If MniPriceVariationList.Visible Then ObjStockReports.PriceVariationListReport
End Sub

Private Sub MniProductAnalysisReport_Click()
   If MniProductAnalysisReport.Visible Then ObjStockReports.ProductAnalysisReport
End Sub

Private Sub MniProductAnalysisStoreValueReport_Click()
   If MniProductAnalysisStoreValueReport.Visible Then ObjStockReports.ProductAnalysisStoreValueReport
End Sub

Private Sub MniProductDifference_Click()
   If MniProductDifference.Visible Then ObjOtherReports.ProductDifferenceReport
End Sub
Private Sub MniProductionRegister_Click()
   If MniProductionRegister.Visible Then ObjProductionReport.ProductionRegisterReport
End Sub

Private Sub MniProductNotIncludedList_Click()
   If MniProductNotIncludedList.Visible Then ObjListReport.ProductNotIncludedListReport
End Sub
Private Sub MniProductPriceList_Click()
   If MniProductPriceList.Visible Then ObjListReport.ProductPriceListReport
End Sub

Private Sub MniProductsNotInPurchase_Click()
   If MniProductsNotInPurchase.Visible Then ObjPurchaseReports.ProductNotInPurchaseReport
End Sub

Private Sub MniProductStockSummaryStoreWise_Click()
   If MniProductStockSummaryStoreWise.Visible Then ObjStockReports.ProductStockSummaryStoreWiseReport
End Sub

Private Sub MniProductStockValue_Click()
   If MniProductStockValue.Visible Then ObjStockReports.ProductStockValueReport
End Sub

Private Sub MniProfitRegister_Click()
   If MniProfitRegister.Visible Then ObjAccountReports.ProfitRegisterReport
End Sub

Private Sub MniReceivedRegister_Click()
   If MniReceivedRegister.Visible Then ObjAccountReports.ReceivedRegisterReport
End Sub

Private Sub MniRecoveryRegister_Click()
   If MniRecoveryRegister.Visible Then ObjAccountReports.RecoveryRegisterReport
End Sub

Private Sub MniRecoverySheet_Click()
   If MniRecoverySheet.Visible Then ObjAccountReports.RoceverySheetReport
End Sub
Private Sub MniPurchaseOrderRegister_Click()
   If MniPurchaseOrderRegister.Visible Then ObjPurchaseReports.PurchaseOrderRegisterReport
End Sub

Private Sub MniPurchaseRegister_Click()
   If MniPurchaseRegister.Visible Then ObjPurchaseReports.PurchaseRegisterReport
End Sub

Private Sub MniPurchaseRegisterSerailWise_Click()
   If MniPurchaseRegisterSerailWise.Visible Then ObjPurchaseReports.PurchaseSerialWiseReport
End Sub
Private Sub MniSalaryEmployeeWise_Click()
   If MniSalaryEmployeeWise.Visible Then ObjOtherReports.SalaryEmployeeWiseReport
End Sub
Private Sub MniSaleExpenseStoreWise_Click()
   If MniSaleExpenseStoreWise.Visible Then ObjSaleReports.SaleExpenseStoreWiseReport
End Sub
Private Sub MniSaleOrderRegister_Click()
   If MniSaleOrderRegister.Visible Then ObjSaleReports.SaleOrderRegisterReport
End Sub

Private Sub MniSaleRegister_Click()
   If MniSaleRegister.Visible Then ObjSaleReports.SaleRegisterReport
End Sub

Private Sub MniSaleRegisterSerailWise_Click()
   If MniSaleRegisterSerailWise.Visible Then ObjSaleReports.SaleRegisterSerailWiseReport
End Sub

Private Sub MniSalesTaxRegister_Click()
   If MniSalesTaxRegister.Visible Then ObjSaleReports.SalestaxRegisterReport
End Sub

Private Sub MniSingleBarCodePrinting_Click()
   If MniMultipleBarCodePrinting.Visible Then ObjDefinition.SingleBarcodeForm
End Sub

Private Sub MniSleepingMembers_Click()
   If MniSleepingMembers.Visible Then ObjOtherReports.SleepingMembersReport
End Sub
Private Sub MniStockRegister_Click()
   If MniStockRegister.Visible Then ObjStockReports.StockRegisterReport
End Sub

Private Sub MniStockValueRegister_Click()
   If MniStockValueRegister.Visible Then ObjStockReports.StockValueRegisterReport
End Sub
Private Sub MniStockAdjustmentValueReport_Click()
   If MniStockAdjustmentValueReport.Visible Then ObjStockReports.StockAdjustmentValueReport
End Sub
Private Sub MniCashBook_Click()
   If MniCashBook.Visible Then ObjAccountReports.CashBookReport
End Sub
Private Sub MniCashFlow_Click()
   If MniCashFlow.Visible Then ObjAccountReports.CashFlowReport
End Sub
Private Sub MniDateWiseProfit_Click()
   If MniDateWiseProfit.Visible Then ObjAccountReports.DateWiseProfitReport
End Sub

Private Sub MniDateWiseSaleExpense_Click()
   If MniDateWiseSaleExpense.Visible Then ObjSaleReports.DateWiseSaleExpenseReport
End Sub

Private Sub MniDemandList_Click()
   If MniDemandList.Visible Then ObjStockReports.DemandListReport
End Sub

Private Sub MniProductLedger_Click()
   If MniProductLedger.Visible Then ObjSaleReports.ProductLedgerReport
End Sub

Private Sub MniProductStockSummary_Click()
   If MniProductStockSummary.Visible Then ObjStockReports.ProductStockSummaryReport
End Sub

Private Sub MniProductsUsed_Click()
   If MniProductsUsed.Visible Then ObjProductionReport.ProductsUsedReport
End Sub

Private Sub MniCurrentStockExpiryValue_Click()
   If MniCurrentStockExpiryValue.Visible Then ObjStockReports.CurrentExpiryValueReport
End Sub

Private Sub MniCurrentStockWastage_Click()
   If MniCurrentStockWastage.Visible Then ObjStockReports.CurrentStockWastageReport
End Sub

Private Sub MniCustomerList_Click()
   If MniCustomerList.Visible Then ObjListReport.CustomerListReport
End Sub

Private Sub MniLedger_Click()
   If MniLedger.Visible Then ObjAccountReports.LedgerReport
End Sub

Private Sub MniProductList_Click()
   If MniProductList.Visible Then ObjListReport.ProductListReport
End Sub

Private Sub MniProfit_Click()
   If MniProfit.Visible Then ObjAccountReports.PLStatementOrgReport
End Sub

Private Sub MniStockWastage_Click()
    If MniStockWastage.Visible Then ObjStockReports.StockWastageReport
End Sub

Private Sub MniStoreList_Click()
   If MniStoreList.Visible Then ObjListReport.StoreListReport
End Sub
Private Sub MniStoreWiseStockTransferDetail_Click()
   If MniStoreWiseStockTransferDetail.Visible Then ObjStockReports.StoreWiseStockTransferDetailReport
End Sub

Private Sub MniSubGroupList_Click()
   If MniSubGroupList.Visible Then ObjListReport.SubGroupListReport
End Sub

Private Sub MniTrialBalance_Click()
   If MniTrialBalance.Visible Then ObjAccountReports.TrialBalanceReport
End Sub

Private Sub MniVendorList_Click()
   If MniVendorList.Visible Then ObjListReport.VendorListReport
End Sub

Private Sub MniVenderPurchaseBills_Click()
   If MniVenderPurchaseBills.Visible Then ObjPurchaseReports.VenderPurchaseBillsReport
End Sub

Private Sub MniMultipleBarCodePrinting_Click()
   If MniMultipleBarCodePrinting.Visible Then ObjDefinition.MultiBarcodesForm
End Sub

Private Sub MnuLogOut_Click()
   Unload Me
End Sub
Public Sub DesktopShortcutsRecport(ShortcutName As String)
   Select Case ShortcutName
      Case "MniAccountPayable"
         Call MniAccountPayable_Click
      Case "MniAccountReceivable"
         Call MniAccountReceivable_Click
      Case "MniAccountDefaultSettings"
'         Call MniAccountDefaultSettings_Click
      Case "MniActivityLog"
'         Call MniActivityLog_Click
      Case "MniAddSkin"
'         Call MniAddSkin_Click
      Case "MniAdminClosing"
'         Call MniAdminClosing_Click
      Case "MniAdminClosingReport"
         Call MniAdminClosingReport_Click
      Case "MniAdvances"
'         Call MniAdvances_Click
      Case "MniAssignTasks"
'         Call MniAssignTasks_Click
      Case "MniAutoBackup"
'         Call MniAutoBackup_Click
      Case "MniBackup"
'         Call MniBackup_Click
      Case "MniBalanceSheet"
         Call MniBalanceSheet_Click
      Case "MniBankCashDepositReport"
         Call MniBankCashDepositReport_Click
      Case "MniBankChequeDepositReport"
         Call MniBankChequeDepositReport_Click
      Case "MniBankChequeIssuanceReport"
         Call MniBankChequeIssuanceReport_Click
      Case "MniBankChequeReceiveReport"
         Call MniBankChequeReceiveReport_Click
      Case "MniBankMachines"
'         Call MniBankMachines_Click
      Case "MniBatchExpiryReport"
         Call MniBatchExpiryReport_Click
      Case "MniBrands"
'         Call MniBrands_Click
      Case "MniCashBook"
         Call MniCashBook_Click
      Case "MniCashDeposite"
'         Call MniCashDeposite_Click
      Case "MniCashFlow"
         Call MniCashFlow_Click
      Case "MniCashPaymentVoucher"
'         Call MniCashPaymentVoucher_Click
      Case "MniCashReceiveVoucher"
'         Call MniCashReceiveVoucher_Click
      Case "MniChangePrice"
'         Call MniChangePrice_Click
      Case "MniChartofAccounts"
'         Call MniChartofAccounts_Click
      Case "MniChequeDeposit"
'         Call MniChequeDeposit_Click
      Case "MniChequeIssuance"
'         Call MniChequeIssuance_Click
      Case "MniChequeIssuanceReconcilation"
'         Call MniChequeIssuanceReconcilation_Click
      Case "MniChequeReceive"
'         Call MniChequeReceive_Click
      Case "MniChequeReceiveReconcilation"
'         Call MniChequeReceiveReconcilation_Click
      Case "MniChangeCategory"
'         Call MniChangeCategory_Click
      Case "MniCompany"
'         Call MniCompany_Click
      Case "MniCompanyInformation"
'         Call MniCompanyInformation_Click
      Case "MniCurrentStockExpiryValue"
         Call MniCurrentStockExpiryValue_Click
      Case "MniCurrentStockWastage"
         Call MniCurrentStockWastage_Click
      Case "MniCustomOrderBalance"
         Call MniCustomOrderBalance_Click
      Case "MniCustomOrderBooking"
'         Call MniCustomOrderBooking_Click
      Case "MniCustomOrderDelivery"
'         Call MniCustomOrderDelivery_Click
      Case "MniCustomOrderPurchase"
'         Call MniCustomOrderPurchase_Click
      Case "MniCustomOrderReturn"
'         Call MniCustomOrderReturn_Click
      Case "MniCustomProductsAndMeasurements"
'         Call MniCustomProductsAndMeasurements_Click
      Case "MniCustomerDemand"
'         Call MniCustomerDemand_Click
      Case "MniCustomerDemandRegister"
         Call MniCustomerDemandRegister_Click
      Case "MniCustomerList"
         Call MniCustomerList_Click
      Case "MniCustomerSaleBills"
         Call MniCustomerSaleBills_Click
      Case "MniCustomers"
''         Call MniCustomers_Click
      Case "MniDailyActivity"
         Call MniDailyActivity_Click
      Case "MniDateWiseProductionInOut"
         Call MniDateWiseProductionInOut_Click
      Case "MniDateWiseProfit"
         Call MniDateWiseProfit_Click
      Case "MniDateWiseProfitStoreWise"
         Call MniDateWiseProfitStoreWise_Click
      Case "MniDateWiseSaleExpense"
         Call MniDateWiseSaleExpense_Click
      Case "MniDeadProducts"
         Call MniDeadProducts_Click
      Case "MniDefineStockLimit"
'         Call MniDefineStockLimit_Click
      Case "MniDemandList"
         Call MniDemandList_Click
      Case "MniDemandListStoreWise"
         Call MniDemandListStoreWise_Click
      Case "MniDepartments"
'         Call MniDepartments_Click
      Case "MniDesignations"
'         Call MniDesignations_Click
      Case "MniDisputeInvoice"
'         Call MniDisputeInvoice_Click
      Case "MniEmployeeAttendanceAll"
'         Call MniEmployeeAttendanceAll_Click
      Case "MniEmployeeAttendanceIn"
'         Call MniEmployeeAttendanceIn_Click
      Case "MniEmployeeAttendanceOut"
'         Call MniEmployeeAttendanceOut_Click
      Case "MniEmployeeAttendanceReport"
         Call MniEmployeeAttendanceReport_Click
      Case "MniEmployeeLeave"
'         Call MniEmployeeLeave_Click
      Case "MniEmployeeList"
         Call MniEmployeeList_Click
      Case "MniEmployees"
'         Call MniEmployees_Click
      Case "MniExpSetting"
'         Call MniExpSetting_Click
      Case "MniExpiryDamageClaimInvoice"
'         Call MniExpiryDamageClaimInvoice_Click
      Case "MniExpiryDamageInvoice"
'         Call MniExpiryDamageInvoice_Click
      Case "MniFinishedProduct"
'         Call MniFinishedProduct_Click
      Case "MniFormulaInfo"
'         Call MniFormulaInfo_Click
      Case "MniGatePassIn"
'         Call MniGatePassIn_Click
      Case "MniGatePassOut"
'         Call MniGatePassOut_Click
      Case "MniGroups"
'         Call MniGroups_Click
      Case "MniHoliday"
'         Call MniHoliday_Click
      Case "MniHotandColdMembers"
         Call MniHotandColdMembers_Click
      Case "MniHotandColdProducts"
         Call MniHotandColdProducts_Click
      Case "MniInstallmentCustomers"
'         Call MniInstallmentCustomers_Click
      Case "MniJournalVoucher"
'         Call MniJournalVoucher_Click
      Case "MniLedger"
         Call MniLedger_Click
      Case "MniLedgerDetail"
         Call MniLedgerDetail_Click
      Case "MniLiftInvoice"
'         Call MniLiftInvoice_Click
      Case "MniLoans"
'         Call MniLoans_Click
      Case "MniLockAccounts"
      Case "MniManufacturedProduct"
         Call MniManufacturedProduct_Click
      Case "MniManufacturedProducts"
'         Call MniManufacturedProducts_Click
      Case "MniManufacturedReturn"
'         Call MniManufacturedReturn_Click
      Case "MniMemberList"
         Call MniMemberList_Click
      Case "MniMemberTypes"
'         Call MniMemberTypes_Click
      Case "MniMembers"
'         Call MniMembers_Click
      Case "MniMembersDiscount"
'         Call MniMembersDiscount_Click
      Case "MniMeterReadingRegister"
         Call MniMeterReadingRegister_Click
      Case "MniMeterReadings"
'         Call MniMeterReadings_Click
      Case "MniMonthlyIncomeExpense"
'         Call MniMonthlyIncomeExpense_Click
      Case "MniMultipleBarCodePrinting"
         Call MniMultipleBarCodePrinting_Click
      Case "MniOpeningAccountsOrganization"
'         Call MniOpeningAccountsOrganization_Click
      Case "MniOpeningProduct"
'         Call MniOpeningProduct_Click
      Case "MniOpeningProductVerification"
'         Call MniOpeningProductVerification_Click
      Case "MniOpeningStock"
'         Call MniOpeningStock_Click
      Case "MniOrganization"
'         Call MniOrganization_Click
      Case "MniPLSetting"
'         Call MniPLSetting_Click
      Case "MniProfit"
         Call MniProfit_Click
      Case "MniPacking"
'         Call MniPacking_Click
      Case "MniPaymentInvoiceWise"
'         Call MniPaymentInvoiceWise_Click
      Case "MniPaymentVender"
'         Call MniPaymentVender_Click
      Case "MniPettyCash"
'         Call MniPettyCash_Click
      Case "MniPettyCashVerification"
'         Call MniPettyCashVerification_Click
      Case "MniProductAnalysisReport"
         Call MniProductAnalysisReport_Click
      Case "MniProductAnalysisStoreValueReport"
         Call MniProductAnalysisStoreValueReport_Click
      Case "MniProductDifference"
         Call MniProductDifference_Click
      Case "MniProductLedger"
         Call MniProductLedger_Click
      Case "MniProductList"
         Call MniProductList_Click
      Case "MniProductNotIncludedList"
         Call MniProductNotIncludedList_Click
      Case "MniProductOffer"
'         Call MniProductOffer_Click
      Case "MniProductPriceList"
         Call MniProductPriceList_Click
      Case "MniProductStockSummary"
         Call MniProductStockSummary_Click
      Case "MniProductStockSummaryStoreWise"
         Call MniProductStockSummaryStoreWise_Click
      Case "MniProductStockValue"
         Call MniProductStockValue_Click
      Case "MniProductWiseSaleStoreWise"
'         Call MniProductWiseSaleStoreWise_Click
      Case "MniProductionIN"
'         Call MniProductionIN_Click
      Case "MniProductionOut"
'         Call MniProductionOut_Click
      Case "MniProductionRegister"
         Call MniProductionRegister_Click
      Case "MniProducts"
'         Call MniProducts_Click
      Case "MniProductsNotInPurchase"
         Call MniProductsNotInPurchase_Click
      Case "MniProductsUsed"
         Call MniProductsUsed_Click
      Case "MniProfitRegister"
         Call MniProfitRegister_Click
      Case "MniPurchaseInvoice"
'         Call MniPurchaseInvoice_Click
      Case "MniPurchaseOrder"
'         Call MniPurchaseOrder_Click
      Case "MniPurchaseRegister"
         Call MniPurchaseRegister_Click
      Case "MniPurchaseRegisterSerailWise"
         Call MniPurchaseRegisterSerailWise_Click
      Case "MniPurchaseReturnInvoice"
'         Call MniPurchaseReturnInvoice_Click
      Case "MniRecoveryCustomer"
'         Call MniRecoveryCustomer_Click
      Case "MniRecoveryInvoiceWise"
'         Call MniRecoveryInvoiceWise_Click
      Case "MniRecoverySheet"
         Call MniRecoverySheet_Click
      Case "MniReferences"
'         Call MniReferences_Click
      Case "MniReplacementInvoice"
'         Call MniReplacementInvoice_Click
      Case "MniRestore"
'         Call MniRestore_Click
      Case "MniSalary"
'         Call MniSalary_Click
      Case "MniSalaryEmployeeWise"
         Call MniSalaryEmployeeWise_Click
      Case "MniSaleCounters"
'         Call MniSaleCounters_Click
      Case "MniSaleExpenseStoreWise"
         Call MniSaleExpenseStoreWise_Click
      Case "MniSaleInvoice"
'         Call MniSaleInvoice_Click
      Case "MniSaleInvoicePOS"
'         Call MniSaleInvoicePOS_Click
      Case "MniSaleOrder"
'         Call MniSaleOrder_Click
      Case "MniSaleOrderPOS"
'         Call MniSaleOrderPOS_Click
      Case "MniSaleRegister"
         Call MniSaleRegister_Click
      Case "MniSaleRegisterSerailWise"
         Call MniSaleRegisterSerailWise_Click
      Case "MniSaleReturnInvoice"
'         Call MniSaleReturnInvoice_Click
      Case "MniSaleReturnInvoicePOS"
'         Call MniSaleReturnInvoicePOS_Click
      Case "MniSectors"
'         Call MniSectors_Click
      Case "MniServiceInvoice"
'         Call MniServiceInvoice_Click
      Case "MniServiceProducts"
'         Call MniServiceProducts_Click
      Case "MniShifts"
'         Call MniShifts_Click
      Case "MniSkinSelection"
'         Call MniSkinSelection_Click
      Case "MniSleepingMembers"
         Call MniSleepingMembers_Click
      Case "MniSoftwareDefaultSettings"
'         Call MniSoftwareDefaultSettings_Click
      Case "MniStockAdjustment"
'         Call MniStockAdjustment_Click
      Case "MniStockAdjustmentValueReport"
         Call MniStockAdjustmentValueReport_Click
      Case "MniStockLimitStoreWise"
'         Call MniStockLimitStoreWise_Click
      Case "MniStockPendingTransfer"
'         Call MniStockPendingTransfer_Click
      Case "MniStockRegister"
         Call MniStockRegister_Click
      Case "MniStoreWiseStockTransferDetail"
         Call MniStoreWiseStockTransferDetail_Click
      Case "MniStockTransferInvoice"
'         Call MniStockTransferInvoice_Click
      Case "MniStockValueRegister"
         Call MniStockValueRegister_Click
      Case "MniStockWastageInvoice"
'         Call MniStockWastageInvoice_Click
      Case "MniStoreList"
         Call MniStoreList_Click
      Case "MniStores"
'         Call MniStores_Click
      Case "MniSubGroup"
'         Call MniSubGroup_Click
      Case "MniSubGroupList"
         Call MniSubGroupList_Click
      Case "MniTables"
'         Call MniTables_Click
      Case "MniTrialBalance"
         Call MniTrialBalance_Click
      Case "MniPackageDealInfo"
'         Call MniPackageDealInfo_Click
      Case "MniUnits"
'         Call MniUnits_Click
      Case "MniUserDefaultSettings"
'         Call MniUserDefaultSettings_Click
      Case "MniUsers"
'         Call MniUsers_Click
      Case "MniVenderPurchaseBills"
         Call MniVenderPurchaseBills_Click
      Case "MniVendors"
'         Call MniVendors_Click
      Case "MniVendorList"
         Call MniVendorList_Click
      Case "MniZones"
'         Call MniZones_Click
   End Select
End Sub

