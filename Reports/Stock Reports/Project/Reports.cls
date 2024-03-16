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

Public Sub CurrentExpiryValueReport()
    RptCurrentExpiryValue.Show
End Sub

Public Sub CurrentStockReport()
'    RptCurrentStockGroupWise.Show
End Sub

Public Sub OpeningStockReport()
'    RptOpeningStock.Show
End Sub

Public Sub ProductAnalysisReport()
    RptProductAnalysis.Show
End Sub

Public Sub ProductAnalysisStoreValueReport()
    RptProductAnalysisStoreValue.Show
End Sub

Public Sub StockValueReport()
'    RptStockValue.Show
End Sub

Public Sub DemandListReport()
    RptDemandList.Show
End Sub

Public Sub DemandListStoreWiseReport()
    RptDemandListStoreWise.Show
End Sub

Public Sub CurrentStockStoreWiseReport()
'    RptCurrentStockStoreWise.Show
End Sub

Public Sub CurrentStockValueReport()
'    RptCurrentStockValue.Show
End Sub

Public Sub CurrentStockValueGroupWiseReport()
'    RptCurrentStockValueGroupWise.Show
End Sub

Public Sub CurrentStockValueStoreWiseReport()
'    RptCurrentStockValueStoreWise.Show
End Sub

Public Sub CurrentStockWastageReport()
    RptCurrentStockWastage.Show
End Sub

Public Sub StoreWiseStockTransferDetailReport()
    RptStoreWiseStockTransferDetail.Show
End Sub

Public Sub CurrentStockCompanyWiseReport()
'    RptCurrentStockCompanyWise.Show
End Sub

Public Sub ProductStockSummaryReport()
    RptProductStockSummary.Show
End Sub

Public Sub ProductStockSummaryStoreWiseReport()
    RptProductStockSummaryStoreWise.Show
End Sub

Public Sub ProductStockValueReport()
    RptProductStockValue.Show
End Sub

Public Sub StockAdjustmentReport()
    'RptStockAdjustment.Show
End Sub

Public Sub StockAdjustmentValueReport()
    RptStockAdjustmentValue.Show
End Sub

Public Sub BatchExpiryReport()
    RptBatchExpiry.Show
End Sub

Public Sub StockValueRegisterReport()
    RptStockValueRegister.Show
End Sub

Public Sub StockRegisterReport()
    RptStockRegister.Show
End Sub

Public Sub PriceVariationListReport()
    RptPriceVariationList.Show
End Sub

Public Sub StockWastageReport()
    RptStockWastage.Show
End Sub

Public Sub DailyDemandListCompanyWiseReport()
    RptDailyDemandListCompanyWise.Show
End Sub
Public Sub OpeningStockRegisterReport()
    RptOpeningStockRegister.Show
End Sub

