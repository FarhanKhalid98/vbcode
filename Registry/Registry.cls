VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = True
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

Public Sub InitializeClass(Conn As ADODB.Connection)
   If ParaPass <> EncryptStr("›· ‹ÔÙÌÓ€ﬂÿ‡", False) Then Exit Sub
   SetConnection Conn
End Sub

Public Sub InitializeStatus(vData As Boolean)
   vReadOnly = vData
End Sub

Private Function GetRegistryValue(key As String) As String
   'On Error GoTo ErrorHandler
    GetRegistryValue = CN.Execute("Select isnull(dbo.FunRValue('" & key & "'),'empty')").Fields(0)
    If GetRegistryValue = "empty" Then
        MsgBox "The Registry Key '" & key & "' not found on your system. Please contact your software vendor", vbCritical, "Registry Not Found"
        GetRegistryValue = Empty
   End If
'   Exit Function
'ErrorHandler:
'   Call ShowErrorMessage
End Function

Public Sub RefreshRegistry()
   'Company
   CompanyName = GetRegistryValue("CompanyName")
   CompanyShortName = GetRegistryValue("CompanyShortName")
   CompanyAddress = GetRegistryValue("CompanyAddress")
   CompanyCity = GetRegistryValue("CompanyCity")
   CompanyPhoneNo = GetRegistryValue("CompanyPhoneNo")
   CompanyEMail = GetRegistryValue("CompanyEmail")
   
   'Hidden
   DevelopedBy = EncryptStr(CN.Execute("Select dbo.Value('1')").Fields(0), False) 'GetRegistryValue("DevelopedBy")
   UpdateDatabase = GetRegistryValue("UpdateDatabase")
   DuplicateCode = GetRegistryValue("DuplicateCode")
   LastRateVisible = GetRegistryValue("LastRateVisible")
   AutoPrintinInvoices = GetRegistryValue("AutoPrintinInvoices")
   MainStoreID = GetRegistryValue("MainStoreID")
   AddSpace = GetRegistryValue("AddSpace")
   LineSpacing = GetRegistryValue("LineSpacing")
   AllowSameStore = GetRegistryValue("AllowSameStore")
   ShowLastPriceOption = GetRegistryValue("ShowLastPriceOption")
   ShowWeightedAvgOption = GetRegistryValue("ShowWeightedAvgOption")
   ShowMovingAvgOption = GetRegistryValue("ShowMovingAvgOption")
   AutoEnterBeforeQty = GetRegistryValue("AutoEnterBeforeQty")
   AutoPrintSaleOrder = GetRegistryValue("AutoPrintSaleOrder")
   PrintKitchenInoices = GetRegistryValue("PrintKitchenInoices")
   EmptyEnterGotoSave = GetRegistryValue("EmptyEnterGotoSave")
   InvType = GetRegistryValue("InvTypeVisible")
   ShowOrganizationWiseStock = GetRegistryValue("ShowOrganizationWiseStock")
   PreviewSaleInoice = GetRegistryValue("PreviewSaleInoice")
   CustomerTypeVisible = GetRegistryValue("CustomerTypeVisible")
   ShowRetailPriceStockRegister = GetRegistryValue("ShowRetailPriceStockRegister")
   ShowMultiBranches = GetRegistryValue("ShowMultiBranches")
   ShowDisc2 = GetRegistryValue("ShowDisc2")
   
   
   'Allow
   NegativeSale = GetRegistryValue("NegativeSale")
   SystemDate = GetRegistryValue("SystemDate")
   SaleInProduction = GetRegistryValue("SaleInProduction")
   PrintHeadersSaleInvoice = GetRegistryValue("PrintHeadersSaleInvoice")
   LaserPrintofSaleInvoice = GetRegistryValue("LaserPrintofSaleInvoice")
   ProperCase = GetRegistryValue("ProperCase")
   DiscAllowed = GetRegistryValue("DiscAllowed")
   CostVisible = GetRegistryValue("CostVisible")
   CashReceived = GetRegistryValue("CashReceived")
   ProductSearchOpenInPreviousState = GetRegistryValue("ProductSearchOpenInPreviousState")
   ChangePrice = GetRegistryValue("ChangePrice")
   AutoApplyPartyLastPrice = GetRegistryValue("AutoApplyPartyLastPrice")
   AutoApplyPartyLastDiscount = GetRegistryValue("AutoApplyPartyLastDiscount")
   AllowUrduProduct = GetRegistryValue("AllowUrduProduct")
   AllowUrduRemarks = GetRegistryValue("AllowUrduRemarks")
   AllowFBRContinuousNo = GetRegistryValue("AllowFBRContinuousNo")
   
   AlertAllocateProduct = GetRegistryValue("AlertAllocateProduct")
   SeperateProductWithPrice = GetRegistryValue("SeperateProductWithPrice")
   SeperateProductInPOS = GetRegistryValue("SeperateProductInPOS")
   AllowNegativeStockInBarcodes = GetRegistryValue("AllowNegativeStockInBarcodes")
   AllowEmployeProductWise = GetRegistryValue("AllowEmployeProductWise")
   AllowStoreProductWise = GetRegistryValue("AllowStoreProductWise")
   AllowOrderByCodeinInvoices = GetRegistryValue("AllowOrderByCodeinInvoices")
   AllowManualBillNoValidation = GetRegistryValue("AllowManualBillNoValidation")
   AllowContinuousBillNo = GetRegistryValue("AllowContinuousBillNo")
   AllowMonthlyBillNo = GetRegistryValue("AllowMonthlyBillNo")
   AllowDailyBillNo = GetRegistryValue("AllowDailyBillNo")
   AllowBothPackingsareSame = GetRegistryValue("AllowBothPackingsareSame")
   
   SetEnterKeyGridStockAdjustment = GetRegistryValue("SetEnterKeyGridStockAdjustment")
   SaveAsNewBill = GetRegistryValue("SaveAsNewBill")
   AfterRowEditFocusNextGridLine = GetRegistryValue("AfterRowEditFocusNextGridLine")
   HideClearButton = GetRegistryValue("HideClearButton")
   AutoEnterQtyintoGridSaleInvoice = GetRegistryValue("AutoEnterQtyintoGridSaleInvoice")
   AllowSMSOnSave = GetRegistryValue("AllowSMSOnSave")
   AllowSMSOnDelete = GetRegistryValue("AllowSMSOnDelete")
   AllowSMSOnClear = GetRegistryValue("AllowSMSOnClear")
   AllowSMSWithDetail = GetRegistryValue("AllowSMSWithDetail")
   AllowSMSOnLogin = GetRegistryValue("AllowSMSOnLogin")
   AllowSMSThroughDevice = GetRegistryValue("AllowSMSThroughDevice")
   ChangeQtyOnChangedPrice = GetRegistryValue("ChangeQtyOnChangedPrice")
   HeaderInfoNotClear = GetRegistryValue("HeaderInfoNotClear")
   AutoMoveGridWhenSerialEntered = GetRegistryValue("AutoMoveGridWhenSerialEntered")
   SerialCompulsoryinInvoice = GetRegistryValue("SerialCompulsoryinInvoice")
   EmployeeMandatory = GetRegistryValue("EmployeeMandatory")
   TableIDMandatory = GetRegistryValue("TableIDMandatory")
   UpdateStockSaleBodyInsert = GetRegistryValue("UpdateStockSaleBodyInsert")
   EmployeeCommision = GetRegistryValue("EmployeeCommision")
   SalePriceLessThanPurchase = GetRegistryValue("SalePriceLessThanPurchase")
   ChangeQtyPack = GetRegistryValue("ChangeQtyPack")
   EitherPackORLooseEnter = GetRegistryValue("EitherPackORLooseEnter")
   IsRoundFigure = GetRegistryValue("IsRoundFigure")
   IsSingleBarcode = GetRegistryValue("IsSingleBarcode")
   SectorCompulsory = GetRegistryValue("SectorCompulsory")
   RemarksCompulsory = GetRegistryValue("RemarksCompulsory")
   DivideRetailWithPacking = GetRegistryValue("DivideRetailWithPacking")
   DisableQuantityinPOS = GetRegistryValue("DisableQuantityinPOS")
   UseMultipleStore = GetRegistryValue("UseMultipleStore")
   PLSamePR = GetRegistryValue("PLSamePR")
   AdminClosingSaveWhenUserClosingSaved = GetRegistryValue("AdminClosingSaveWhenUserClosingSaved")
   LockPurPrice = GetRegistryValue("LockPurPrice")
   ProductSearchWithStore = GetRegistryValue("ProductSearchWithStore")
   ShowAllStoreStock = GetRegistryValue("ShowAllStoreStock")
   UsePurPrice = GetRegistryValue("UsePurPrice")
   ChangeTransactionDate = GetRegistryValue("ChangeTransactionDate")
   UseBin = GetRegistryValue("UseBin")
   AttendanceNextDayOut = GetRegistryValue("AttendanceNextDayOut")
   AllowDiscountOnSaleDistribution = GetRegistryValue("AllowDiscountOnSaleDistribution")
   DefaultCustomer = GetRegistryValue("DefaultCustomer")
   ClientDate = GetRegistryValue("ClientDate")
   IsAllowBarCodeQtyInpurchaseQty = GetRegistryValue("IsAllowBarCodeQtyInpurchaseQty")
   TermAllowZero = GetRegistryValue("TermAllowZero")
   'Show Hide
   Tag = GetRegistryValue("Tag")
   QuantityinBarcodes = GetRegistryValue("QuantityinBarcodes")
   BatchNoVisible = GetRegistryValue("BatchNoVisible")
   StoreVisible = GetRegistryValue("StoreVisible")
   CustomerTypeVisible = GetRegistryValue("CustomerTypeVisible")
   ManualBillNoVisible = GetRegistryValue("ManualBillNoVisible")
   TableVisible = GetRegistryValue("TableVisible")
   OrganizationVisible = GetRegistryValue("OrganizationVisible")
   RemarksVisible = GetRegistryValue("RemarksVisible")
   MemberVisible = GetRegistryValue("MemberVisible")
   SaleOrderVisible = GetRegistryValue("SaleOrderVisible")
   FreightVisible = GetRegistryValue("FreightVisible")
   EmpVisible = GetRegistryValue("EmpVisible")
   PreviousBalanceVisible = GetRegistryValue("PreviousBalanceVisible")
   HidePurchaseAmount = GetRegistryValue("HidePurchaseAmount")
   HideSaleAmount = GetRegistryValue("HideSaleAmount")
   ShowBankInTransection = GetRegistryValue("ShowBankInTransection")
   ShowSerial = GetRegistryValue("ShowSerial")
   ShowCode = GetRegistryValue("ShowCode")
   HideAutoPrint = GetRegistryValue("HideAutoPrint")
   ShowWarrantyinSaleInvoice = GetRegistryValue("ShowWarrantyinSaleInvoice")
   ShowLastInvoiceMsgAtSave = GetRegistryValue("ShowLastInvoiceMsgAtSave")
   ShowRetailinPurchaseReturnPrint = GetRegistryValue("ShowRetailinPurchaseReturnPrint")
   ShowRawMaterialProductInSaleInvoices = GetRegistryValue("ShowRawMaterialProductInSaleInvoices")
   ShowWholeSaleMargin = GetRegistryValue("ShowWholeSaleMargin")
   ShowPromiseDateInSalaPurchase = GetRegistryValue("ShowPromiseDateInSalaPurchase")
   ShowSyllabus = GetRegistryValue("ShowSyllabus")
   ShowProdProfit = GetRegistryValue("ShowProdProfit")
   ShowStockFromTableGridDataMovement = GetRegistryValue("ShowStockFromTableGridDataMovement")
   ShowBarcodeProductSearch = GetRegistryValue("ShowBarcodeProductSearch")
   OrganizationMandatory = GetRegistryValue("OrganizationMandatory")
   isShowPublisher = GetRegistryValue("isShowPublisher")
   isShowListPrice = GetRegistryValue("isShowListPrice")
   IsShowPrintYesOrNo = GetRegistryValue("IsShowPrintYesOrNo")
   isShowDepartment = GetRegistryValue("isShowDepartment")
   isShowSubDepartment = GetRegistryValue("isShowSubDepartment")
   isShowSeason = GetRegistryValue("isShowSeason")
   isShowItemDesc = GetRegistryValue("isShowItemDesc")
   isShowOther = GetRegistryValue("isShowOther")
   isShowVendor = GetRegistryValue("isShowVendor")
   ShowColourSize = GetRegistryValue("ShowColourSize")
   IsGrossQty = GetRegistryValue("IsGrossQty")
   ShowSavedStock = GetRegistryValue("ShowSavedStock")
   ShowAllPrices = GetRegistryValue("ShowAllPrices")
   ShowSession = GetRegistryValue("ShowSession")
   ShowTradeOffer = GetRegistryValue("ShowTradeOffer")
   ShowExpiryInvoice = GetRegistryValue("ShowExpiryInvoice")
   ShowBonus = GetRegistryValue("ShowBonus")
   ShowOffer = GetRegistryValue("ShowOffer")
   ShowSaleTax = GetRegistryValue("ShowSaleTax")
   ShowSC = GetRegistryValue("ShowSC")
   ShowBatchPrint = GetRegistryValue("ShowBatchPrint")
   ShowBarCodeQty = GetRegistryValue("ShowBarCodeQty")
   ShowChangePriceOnSavePI = GetRegistryValue("ShowChangePriceOnSavePI")
   ShowChangeRetailInPurchaseInvoice = GetRegistryValue("ShowChangeRetailInPurchaseInvoice")
   ShowGrandTotalinSearch = GetRegistryValue("ShowGrandTotalinSearch")
   ShowDispatchDate = GetRegistryValue("ShowDispatchDate")
   ShowReSale = GetRegistryValue("ShowReSale")
   ShowDiscPurPrice = GetRegistryValue("ShowDiscPurPrice")
   ShowStockPriceChecker = GetRegistryValue("ShowStockPriceChecker")
   AllowNegativeOrder = GetRegistryValue("AllowNegativeOrder")
   ShowAddBarCode = GetRegistryValue("ShowAddBarCode")
   ShowPurPrice = GetRegistryValue("ShowPurPrice")
   SearchCodeInGrid = GetRegistryValue("SearchCodeInGrid")
   UsePasswordForm = GetRegistryValue("UsePasswordForm")
   ShowPurchaseProfit = GetRegistryValue("ShowPurchaseProfit")
   ShowBarcodeDesc = GetRegistryValue("ShowBarcodeDesc")
   ShowDiscPrice = GetRegistryValue("ShowDiscPrice")
   ShowDuplicatePrint = GetRegistryValue("ShowDuplicatePrint")
   RunnngLastPrice = GetRegistryValue("RunnngLastPrice")
   
   'Default Values
   OrganizationID = GetRegistryValue("OrganizationID")
   StoreID = GetRegistryValue("StoreID")
   BankMachineID = GetRegistryValue("BankMachineID")
   NoofPrints = GetRegistryValue("NoofPrints")
   MemberMin = GetRegistryValue("MemberMin")
   MemberMax = GetRegistryValue("MemberMax")
   DeviceName = GetRegistryValue("DeviceName")
   DeviceName2 = GetRegistryValue("DeviceName2")
   DriverName = GetRegistryValue("DriverName")
   Port = GetRegistryValue("Port")
   ProdDesc1 = GetRegistryValue("ProdDesc1")
   PackingChargesPer = GetRegistryValue("PackingChargesPer")
   HourDifference = GetRegistryValue("HourDifference")
   EmployeeLateRelaxTime = GetRegistryValue("EmployeeLateRelaxTime")
   BarCodePrefix = GetRegistryValue("BarCodePrefix")
   SearchDateDifference = GetRegistryValue("SearchDateDifference")
   x = GetRegistryValue("X")
   Y = GetRegistryValue("Y")
   CostX = GetRegistryValue("CostX")
   CostY = GetRegistryValue("CostY")
   BlankFooter = GetRegistryValue("BlankFooter")
   ConnectionTimeout = GetRegistryValue("ConnectionTimeOut")
   Statement = GetRegistryValue("Statement")

   OwnerMobileNo = GetRegistryValue("OwnerMobileNo")
   PrefixPhoneNo = GetRegistryValue("PrefixPhoneNo")
   CustomerSalesMessage = GetRegistryValue("CustomerSalesMessage")
   WebLinkForSMS = GetRegistryValue("WebLinkForSMS")
   CurrentDateDataEntry = GetRegistryValue("CurrentDateDataEntry")
   isEntryDate = GetRegistryValue("IsEntryDate")
   
   FromDate = GetRegistryValue("FromDate")
   ToDate = GetRegistryValue("ToDate")
   IsPortrait = GetRegistryValue("IsPortrait")
   IsLegal = GetRegistryValue("IsLegal")
   ChargesName = GetRegistryValue("ChargesName")
   ShowHistoryofAllCustomer = GetRegistryValue("ShowHistoryofAllCustomer")
   
   TimeWiseReport = GetRegistryValue("TimeWiseReport")
   
   CheckStockOnSave = GetRegistryValue("CheckStockOnSave")
   AdminClssingFinePerOnShort = GetRegistryValue("AdminClssingFinePerOnShort")
   GridRowHeight = GetRegistryValue("GridRowHeight")
   RoundfigureInSearchForm = GetRegistryValue("RoundfigureInSearchForm")
   
   FromEmail = GetRegistryValue("FromEmail")
   ToEmail = GetRegistryValue("ToEmail")
   EmailPwd = GetRegistryValue("EmailPwd")
   SMTPServerAddress = GetRegistryValue("SMTPServerAddress")
   PortNo = GetRegistryValue("PortNo")
   UseEmail = GetRegistryValue("UseEmail")
   ExportReportASPDF = GetRegistryValue("ExportReportASPDF")
   ActivityActionNo = GetRegistryValue("ActivityActionNo")
   ByDefaultActionNo = GetRegistryValue("ByDefaultActionNo")
   FBRGroupName = GetRegistryValue("FBRGroupName")
   IsFBRedit = GetRegistryValue("IsFBRedit")
   
'    BankAccounts = GetRegistryValue("BankAccounts")
'    BothQty = GetRegistryValue("BothQty")
'    CounterSaleAmount = GetRegistryValue("CounterSaleAmount")
'    ChangePriceAtSave = GetRegistryValue("ChangePriceAtSave")
'    ShowCostinSaleInvoice = GetRegistryValue("ShowCostinSaleInvoice")
'    ShowAllInvoiceInSearch = GetRegistryValue("ShowAllInvoiceInSearch")
    ReadOnlyStatus = vReadOnly
End Sub

'local variable(s) to hold property value(s)
Public Property Let AlertAllocateProduct(ByVal vNewSetting As Boolean)
   mvarAlertAllocateProduct = vNewSetting
End Property

Public Property Get AlertAllocateProduct() As Boolean
   AlertAllocateProduct = mvarAlertAllocateProduct
End Property

Public Property Let AllowOrderByCodeinInvoices(ByVal vNewSetting As Boolean)
   mvarAllowOrderByCodeinInvoices = vNewSetting
End Property

Public Property Get AllowOrderByCodeinInvoices() As Boolean
   AllowOrderByCodeinInvoices = mvarAllowOrderByCodeinInvoices
End Property

Public Property Let AllowContinuousBillNo(ByVal vNewSetting As Boolean)
   mvarAllowContinuousBillNo = vNewSetting
End Property

Public Property Get AllowContinuousBillNo() As Boolean
   AllowContinuousBillNo = mvarAllowContinuousBillNo
End Property

Public Property Let AllowMonthlyBillNo(ByVal vNewSetting As Boolean)
   mvarAllowMonthlyBillNo = vNewSetting
End Property

Public Property Get AllowMonthlyBillNo() As Boolean
   AllowMonthlyBillNo = mvarAllowMonthlyBillNo
End Property

Public Property Let AllowDailyBillNo(ByVal vNewSetting As Boolean)
   mvarAllowDailyBillNo = vNewSetting
End Property

Public Property Get AllowDailyBillNo() As Boolean
   AllowDailyBillNo = mvarAllowDailyBillNo
End Property

Public Property Let SetEnterKeyGridStockAdjustment(ByVal vNewSetting As Boolean)
   mvarSetEnterKeyGridStockAdjustment = vNewSetting
End Property

Public Property Get SetEnterKeyGridStockAdjustment() As Boolean
   SetEnterKeyGridStockAdjustment = mvarSetEnterKeyGridStockAdjustment
End Property

Public Property Let SaveAsNewBill(ByVal vNewSetting As Boolean)
   mvarSaveAsNewBill = vNewSetting
End Property

Public Property Get SaveAsNewBill() As Boolean
   SaveAsNewBill = mvarSaveAsNewBill
End Property
Public Property Let HideClearButton(ByVal vNewSetting As Boolean)
   mvarHideClearButton = vNewSetting
End Property
Public Property Get HideClearButton() As Boolean
   HideClearButton = mvarHideClearButton
End Property
Public Property Let AfterRowEditFocusNextGridLine(ByVal vNewSetting As Boolean)
   mvarAfterRowEditFocusNextGridLine = vNewSetting
End Property
Public Property Get AfterRowEditFocusNextGridLine() As Boolean
   AfterRowEditFocusNextGridLine = mvarAfterRowEditFocusNextGridLine
End Property
Public Property Let AutoEnterQtyintoGridSaleInvoice(ByVal vNewSetting As Boolean)
   mvarAutoEnterQtyintoGridSaleInvoice = vNewSetting
End Property
Public Property Get AutoEnterQtyintoGridSaleInvoice() As Boolean
   AutoEnterQtyintoGridSaleInvoice = mvarAutoEnterQtyintoGridSaleInvoice
End Property

Public Property Let AllowManualBillNoValidation(ByVal vNewSetting As Boolean)
   mvarAllowManualBillNoValidation = vNewSetting
End Property

Public Property Get AllowManualBillNoValidation() As Boolean
   AllowManualBillNoValidation = mvarAllowManualBillNoValidation
End Property

Public Property Let AllowStoreProductWise(ByVal vNewSetting As Boolean)
   mvarAllowStoreProductWise = vNewSetting
End Property

Public Property Get AllowStoreProductWise() As Boolean
   AllowStoreProductWise = mvarAllowStoreProductWise
End Property

Public Property Let AllowEmployeProductWise(ByVal vNewSetting As Boolean)
   mvarAllowEmployeProductWise = vNewSetting
End Property

Public Property Get AllowEmployeProductWise() As Boolean
   AllowEmployeProductWise = mvarAllowEmployeProductWise
End Property
        
Public Property Let ShowManufacturingProductInInvoices(ByVal vNewSetting As Boolean)
   mvarShowManufacturingProductInInvoices = vNewSetting
End Property

Public Property Get ShowManufacturingProductInInvoices() As Boolean
   ShowManufacturingProductInInvoices = mvarShowManufacturingProductInInvoices
End Property

Public Property Let ShowRawMaterialProductInSaleInvoices(ByVal vNewSetting As Boolean)
   mvarShowRawMaterialProductInSaleInvoices = vNewSetting
End Property

Public Property Get ShowRawMaterialProductInSaleInvoices() As Boolean
   ShowRawMaterialProductInSaleInvoices = mvarShowRawMaterialProductInSaleInvoices
End Property

Public Property Let PreviewSaleInoice(ByVal vNewSetting As Boolean)
   mvarPreviewSaleInoice = vNewSetting
End Property

Public Property Get PreviewSaleInoice() As Boolean
   PreviewSaleInoice = mvarPreviewSaleInoice
End Property

Public Property Let ShowWarrantyinSaleInvoice(ByVal vNewSetting As Boolean)
   mvarShowWarrantyinSaleInvoice = vNewSetting
End Property

Public Property Get ShowWarrantyinSaleInvoice() As Boolean
   ShowWarrantyinSaleInvoice = mvarShowWarrantyinSaleInvoice
End Property

Public Property Let ShowRetailinPurchaseReturnPrint(ByVal vNewSetting As Boolean)
   mvarShowRetailinPurchaseReturnPrint = vNewSetting
End Property

Public Property Get ShowRetailinPurchaseReturnPrint() As Boolean
   ShowRetailinPurchaseReturnPrint = mvarShowRetailinPurchaseReturnPrint
End Property

Public Property Let ShowOrganizationWiseStock(ByVal vNewSetting As Boolean)
   mvarShowOrganizationWiseStock = vNewSetting
End Property

Public Property Get ShowOrganizationWiseStock() As Boolean
   ShowOrganizationWiseStock = mvarShowOrganizationWiseStock
End Property

Public Property Let AllowNegativeStockInBarcodes(ByVal vNewSetting As Boolean)
   mvarAllowNegativeStockInBarcodes = vNewSetting
End Property

Public Property Get AllowNegativeStockInBarcodes() As Boolean
   AllowNegativeStockInBarcodes = mvarAllowNegativeStockInBarcodes
End Property

Public Property Let SeperateProductInPOS(ByVal vNewSetting As Boolean)
   mvarSeperateProductInPOS = vNewSetting
End Property

Public Property Get SeperateProductInPOS() As Boolean
   SeperateProductInPOS = mvarSeperateProductInPOS
End Property

Public Property Let SeperateProductWithPrice(ByVal vNewSetting As Boolean)
   mvarSeperateProductWithPrice = vNewSetting
End Property

Public Property Get SeperateProductWithPrice() As Boolean
   SeperateProductWithPrice = mvarSeperateProductWithPrice
End Property

Public Property Let CompanyName(ByVal vData As String)
   mvarCompanyName = vData
End Property

Public Property Get CompanyName() As String
   CompanyName = mvarCompanyName
End Property

Public Property Let CompanyShortName(ByVal vData As String)
   mvarCompanyShortName = vData
End Property

Public Property Get CompanyShortName() As String
   CompanyShortName = mvarCompanyShortName
End Property

Public Property Let CompanyAddress(ByVal vData As String)
   mvarCompanyAddress = vData
End Property

Public Property Get CompanyAddress() As String
   CompanyAddress = mvarCompanyAddress
End Property

Public Property Let CompanyCity(ByVal vData As String)
   mvarCompanyCity = vData
End Property

Public Property Get CompanyCity() As String
   CompanyCity = mvarCompanyCity
End Property

Public Property Let CompanyPhoneNo(ByVal vData As String)
   mvarCompanyPhoneNo = vData
End Property

Public Property Get CompanyPhoneNo() As String
   CompanyPhoneNo = mvarCompanyPhoneNo
End Property

Public Property Let CompanyEMail(ByVal vData As String)
   mvarCompanyEMail = vData
End Property

Public Property Get CompanyEMail() As String
   CompanyEMail = mvarCompanyEMail
End Property

Public Property Let ConnectionTimeout(ByVal vData As String)
   mvarConnectionTimeout = vData
End Property

Public Property Get ConnectionTimeout() As String
   ConnectionTimeout = mvarConnectionTimeout
End Property

Public Property Let BlankFooter(ByVal vData As String)
   mvarBlankFooter = vData
End Property

Public Property Get BlankFooter() As String
   BlankFooter = mvarBlankFooter
End Property

Public Property Let SearchDateDifference(ByVal vData As String)
   mvarSearchDateDifference = vData
End Property

Public Property Get SearchDateDifference() As String
   SearchDateDifference = mvarSearchDateDifference
End Property

Public Property Let ProdDesc1(ByVal vData As String)
   mvarProdDesc1 = vData
End Property

Public Property Get ProdDesc1() As String
   ProdDesc1 = mvarProdDesc1
End Property

Public Property Let StoreID(ByVal vData As String)
   mvarStoreID = vData
End Property

Public Property Get StoreID() As String
   StoreID = mvarStoreID
End Property
      
Public Property Let EmptyEnterGotoSave(ByVal vNewSetting As Boolean)
   mvarEmptyEnterGotoSave = vNewSetting
End Property

Public Property Get EmptyEnterGotoSave() As Boolean
   EmptyEnterGotoSave = mvarEmptyEnterGotoSave
End Property

Public Property Let AutoPrintSaleOrder(ByVal vNewSetting As Boolean)
   mvarAutoPrintSaleOrder = vNewSetting
End Property

Public Property Get AutoPrintSaleOrder() As Boolean
   AutoPrintSaleOrder = mvarAutoPrintSaleOrder
End Property

Public Property Let PrintKitchenInoices(ByVal vNewSetting As Boolean)
   mvarPrintKitchenInoices = vNewSetting
End Property

Public Property Get PrintKitchenInoices() As Boolean
   PrintKitchenInoices = mvarPrintKitchenInoices
End Property

Public Property Get InvType() As Boolean
   InvType = mvarInvType
End Property

Public Property Let InvType(ByVal vNewSetting As Boolean)
   mvarInvType = vNewSetting
End Property

Public Property Let AutoEnterBeforeQty(ByVal vNewSetting As Boolean)
   mvarAutoEnterBeforeQty = vNewSetting
End Property

Public Property Get AutoEnterBeforeQty() As Boolean
   AutoEnterBeforeQty = mvarAutoEnterBeforeQty
End Property

Public Property Let ShowLastPriceOption(ByVal vNewSetting As Boolean)
   mvarShowLastPriceOption = vNewSetting
End Property

Public Property Get ShowLastPriceOption() As Boolean
   ShowLastPriceOption = mvarShowLastPriceOption
End Property

Public Property Let ShowWeightedAvgOption(ByVal vNewSetting As Boolean)
   mvarShowWeightedAvgOption = vNewSetting
End Property

Public Property Get ShowWeightedAvgOption() As Boolean
   ShowWeightedAvgOption = mvarShowWeightedAvgOption
End Property

Public Property Let ShowMovingAvgOption(ByVal vNewSetting As Boolean)
   mvarShowMovingAvgOption = vNewSetting
End Property

Public Property Get ShowMovingAvgOption() As Boolean
   ShowMovingAvgOption = mvarShowMovingAvgOption
End Property

Public Property Let DiscAllowed(ByVal vNewSetting As Boolean)
   mvarDiscAllowed = vNewSetting
End Property

Public Property Get DiscAllowed() As Boolean
   DiscAllowed = mvarDiscAllowed
End Property

Public Property Let ProductSearchOpenInPreviousState(ByVal vNewSetting As Boolean)
   mvarProductSearchOpenInPreviousState = vNewSetting
End Property

Public Property Get ProductSearchOpenInPreviousState() As Boolean
   ProductSearchOpenInPreviousState = mvarProductSearchOpenInPreviousState
End Property

Public Property Let AllowSameStore(ByVal vNewSetting As Boolean)
   mvarAllowSameStore = vNewSetting
End Property

Public Property Get AllowSameStore() As Boolean
   AllowSameStore = mvarAllowSameStore
End Property

Public Property Let ProperCase(ByVal vNewSetting As Boolean)
   mvarProperCase = vNewSetting
End Property

Public Property Get ProperCase() As Boolean
   ProperCase = mvarProperCase
End Property

Public Property Let SystemDate(ByVal vNewSetting As Boolean)
   mvarSystemDate = vNewSetting
End Property

Public Property Get SystemDate() As Boolean
   SystemDate = mvarSystemDate
End Property

Public Property Let SaleOrderVisible(ByVal vNewSetting As Boolean)
   mvarSaleOrderVisible = vNewSetting
End Property

Public Property Get SaleOrderVisible() As Boolean
   SaleOrderVisible = mvarSaleOrderVisible
End Property

Public Property Let TableVisible(ByVal vNewSetting As Boolean)
   mvarTableVisible = vNewSetting
End Property

Public Property Get TableVisible() As Boolean
   TableVisible = mvarTableVisible
End Property

Public Property Let SaleInProduction(ByVal vNewSetting As Boolean)
   mvarSaleInProduction = vNewSetting
End Property

Public Property Get SaleInProduction() As Boolean
   SaleInProduction = mvarSaleInProduction
End Property

Public Property Let BatchNoVisible(ByVal vNewSetting As Boolean)
   mvarBatchNoVisible = vNewSetting
End Property

Public Property Get BatchNoVisible() As Boolean
   BatchNoVisible = mvarBatchNoVisible
End Property

Public Property Let CustomerTypeVisible(ByVal vNewSetting As Boolean)
   mvarCustomerTypeVisible = vNewSetting
End Property

Public Property Get CustomerTypeVisible() As Boolean
   CustomerTypeVisible = mvarCustomerTypeVisible
End Property

Public Property Let StoreVisible(ByVal vNewSetting As Boolean)
   mvarStoreVisible = vNewSetting
End Property

Public Property Get StoreVisible() As Boolean
   StoreVisible = mvarStoreVisible
End Property

Public Property Let AutoApplyPartyLastPrice(ByVal vNewSetting As Boolean)
   mvarAutoApplyPartyLastPrice = vNewSetting
End Property

Public Property Get AutoApplyPartyLastPrice() As Boolean
   AutoApplyPartyLastPrice = mvarAutoApplyPartyLastPrice
End Property
        
Public Property Let AutoApplyPartyLastDiscount(ByVal vNewSetting As Boolean)
   mvarAutoApplyPartyLastDiscount = vNewSetting
End Property

Public Property Get AutoApplyPartyLastDiscount() As Boolean
   AutoApplyPartyLastDiscount = mvarAutoApplyPartyLastDiscount
End Property

Public Property Let AllowFBRContinuousNo(ByVal vNewSetting As Boolean)
   mvarAllowFBRContinuousNo = vNewSetting
End Property

Public Property Get AllowFBRContinuousNo() As Boolean
   AllowFBRContinuousNo = mvarAllowFBRContinuousNo
End Property

Public Property Let AllowUrduProduct(ByVal vNewSetting As Boolean)
   mvarAllowUrduProduct = vNewSetting
End Property

Public Property Get AllowUrduProduct() As Boolean
   AllowUrduProduct = mvarAllowUrduProduct
End Property

Public Property Let AllowUrduRemarks(ByVal vNewSetting As Boolean)
   mvarAllowUrduRemarks = vNewSetting
End Property

Public Property Get AllowUrduRemarks() As Boolean
   AllowUrduRemarks = mvarAllowUrduRemarks
End Property

Public Property Let NegativeSale(ByVal vNewSetting As Boolean)
   mvarNegativeSale = vNewSetting
End Property

Public Property Get NegativeSale() As Boolean
   NegativeSale = mvarNegativeSale
End Property

Public Property Let AddSpace(ByVal vNewSetting As Boolean)
   mvarAddSpace = vNewSetting
End Property

Public Property Get AddSpace() As Boolean
   AddSpace = mvarAddSpace
End Property

Public Property Let BankMachineID(ByVal vData As String)
   mvarBankMachineID = vData
End Property

Public Property Get BankMachineID() As String
   BankMachineID = mvarBankMachineID
End Property

Public Property Let CashReceived(ByVal vNewSetting As Boolean)
   mvarCashReceived = vNewSetting
End Property

Public Property Get CashReceived() As Boolean
   CashReceived = mvarCashReceived
End Property

Public Property Let DuplicateCode(ByVal vNewSetting As Boolean)
   mvarDuplicateCode = vNewSetting
End Property

Public Property Get DuplicateCode() As Boolean
   DuplicateCode = mvarDuplicateCode
End Property

Public Property Let CostVisible(ByVal vNewSetting As Boolean)
   mvarCostVisible = vNewSetting
End Property

Public Property Get CostVisible() As Boolean
   CostVisible = mvarCostVisible
End Property

Public Property Let LineSpacing(ByVal vNewSetting As Boolean)
   mvarLineSpacing = vNewSetting
End Property

Public Property Get LineSpacing() As Boolean
   LineSpacing = mvarLineSpacing
End Property
    
Public Property Let Statement(ByVal vData As String)
   mvarStatement = vData
End Property

Public Property Get Statement() As String
   Statement = mvarStatement
End Property
    
Public Property Let ChangePrice(ByVal vNewSetting As Boolean)
   mvarChangePrice = vNewSetting
End Property

Public Property Get ChangePrice() As Boolean
   ChangePrice = mvarChangePrice
End Property

Public Property Let FreightVisible(ByVal vNewSetting As Boolean)
   mvarFreightVisible = vNewSetting
End Property

Public Property Get FreightVisible() As Boolean
   FreightVisible = mvarFreightVisible
End Property

Public Property Let LastRateVisible(ByVal vNewSetting As Boolean)
   mvarLastRateVisible = vNewSetting
End Property

Public Property Get LastRateVisible() As Boolean
   LastRateVisible = mvarLastRateVisible
End Property

Public Property Let EmpVisible(ByVal vNewSetting As Boolean)
   mvarEmpVisible = vNewSetting
End Property

Public Property Get EmpVisible() As Boolean
   EmpVisible = mvarEmpVisible
End Property
   
Public Property Let HidePurchaseAmount(ByVal vNewSetting As Boolean)
   mvarHidePurchaseAmount = vNewSetting
End Property

Public Property Get HidePurchaseAmount() As Boolean
   HidePurchaseAmount = mvarHidePurchaseAmount
End Property

Public Property Let HideSaleAmount(ByVal vNewSetting As Boolean)
   mvarHideSaleAmount = vNewSetting
End Property

Public Property Get HideSaleAmount() As Boolean
   HideSaleAmount = mvarHideSaleAmount
End Property

Public Property Let ShowBankInTransection(ByVal vNewSetting As Boolean)
   mvarShowBankInTransection = vNewSetting
End Property

Public Property Get ShowBankInTransection() As Boolean
   ShowBankInTransection = mvarShowBankInTransection
End Property

Public Property Let ShowSerial(ByVal vNewSetting As Boolean)
   mvarShowSerial = vNewSetting
End Property

Public Property Get ShowSerial() As Boolean
   ShowSerial = mvarShowSerial
End Property

Public Property Let ShowLastInvoiceMsgAtSave(ByVal vNewSetting As Boolean)
   mvarShowLastInvoiceMsgAtSave = vNewSetting
End Property

Public Property Get ShowLastInvoiceMsgAtSave() As Boolean
   ShowLastInvoiceMsgAtSave = mvarShowLastInvoiceMsgAtSave
End Property

Public Property Let ShowCode(ByVal vNewSetting As Boolean)
   mvarShowCode = vNewSetting
End Property

Public Property Get ShowCode() As Boolean
   ShowCode = mvarShowCode
End Property

Public Property Let HideAutoPrint(ByVal vNewSetting As Boolean)
   mvarHideAutoPrint = vNewSetting
End Property

Public Property Get HideAutoPrint() As Boolean
   HideAutoPrint = mvarHideAutoPrint
End Property

Public Property Let MemberVisible(ByVal vNewSetting As Boolean)
   mvarMemberVisible = vNewSetting
End Property

Public Property Get MemberVisible() As Boolean
   MemberVisible = mvarMemberVisible
End Property

Public Property Let LaserPrintofSaleInvoice(ByVal vNewSetting As Boolean)
   mvarLaserPrintofSaleInvoice = vNewSetting
End Property

Public Property Get LaserPrintofSaleInvoice() As Boolean
   LaserPrintofSaleInvoice = mvarLaserPrintofSaleInvoice
End Property
    
Public Property Let PrintHeadersSaleInvoice(ByVal vNewSetting As Boolean)
   mvarPrintHeadersSaleInvoice = vNewSetting
End Property

Public Property Get PrintHeadersSaleInvoice() As Boolean
   PrintHeadersSaleInvoice = mvarPrintHeadersSaleInvoice
End Property
    
Public Property Let MainStoreID(ByVal vData As String)
   mvarMainStoreID = vData
End Property

Public Property Get MainStoreID() As String
   MainStoreID = mvarMainStoreID
End Property
    
Public Property Let NoofPrints(ByVal vData As String)
   mvarNoofPrints = vData
End Property

Public Property Get NoofPrints() As String
   NoofPrints = mvarNoofPrints
End Property
    
Public Property Let MemberMin(ByVal vData As String)
   mvarMemberMin = vData
End Property

Public Property Get MemberMin() As String
   MemberMin = mvarMemberMin
End Property

Public Property Let MemberMax(ByVal vData As String)
   mvarMemberMax = vData
End Property

Public Property Get MemberMax() As String
   MemberMax = mvarMemberMax
End Property
    
Public Property Let QuantityinBarcodes(ByVal vNewSetting As Boolean)
   mvarQuantityinBarcodes = vNewSetting
End Property

Public Property Get QuantityinBarcodes() As Boolean
   QuantityinBarcodes = mvarQuantityinBarcodes
End Property

Public Property Let Tag(ByVal vNewSetting As Boolean)
   mvarTag = vNewSetting
End Property

Public Property Get Tag() As Boolean
   Tag = mvarTag
End Property
    
Public Property Let OrderStatement(ByVal vData As String)
   mvarOrderStatement = vData
End Property

Public Property Get OrderStatement() As String
   OrderStatement = mvarOrderStatement
End Property
    
Public Property Let DeviceName(ByVal vData As String)
   mvarDeviceName = vData
End Property

Public Property Get DeviceName() As String
   DeviceName = mvarDeviceName
End Property
     
Public Property Let DriverName(ByVal vData As String)
   mvarDriverName = vData
End Property

Public Property Get DriverName() As String
   DriverName = mvarDriverName
End Property
    
Public Property Let Port(ByVal vData As String)
   mvarPort = vData
End Property

Public Property Get Port() As String
   Port = mvarPort
End Property

Public Property Let DeviceName2(ByVal vData As String)
   mvarDeviceName2 = vData
End Property

Public Property Get DeviceName2() As String
   DeviceName2 = mvarDeviceName2
End Property

Public Property Let DriverName2(ByVal vData As String)
   mvarDriverName2 = vData
End Property

Public Property Get DriverName2() As String
   DriverName2 = mvarDriverName2
End Property

Public Property Let Port2(ByVal vData As String)
   mvarPort2 = vData
End Property

Public Property Get Port2() As String
   Port2 = mvarPort2
End Property

Public Property Let ManualBillNoVisible(ByVal vNewSetting As Boolean)
   mvarManualBillNoVisible = vNewSetting
End Property

Public Property Get ManualBillNoVisible() As Boolean
   ManualBillNoVisible = mvarManualBillNoVisible
End Property

Public Property Let PreviousBalanceVisible(ByVal vNewSetting As Boolean)
   mvarPreviousBalanceVisible = vNewSetting
End Property

Public Property Get PreviousBalanceVisible() As Boolean
   PreviousBalanceVisible = mvarPreviousBalanceVisible
End Property

Public Property Let x(ByVal vData As String)
   mvarX = vData
End Property

Public Property Get x() As String
   x = mvarX
End Property

Public Property Let Y(ByVal vData As String)
   mvarY = vData
End Property

Public Property Get Y() As String
   Y = mvarY
End Property

Public Property Let CostX(ByVal vData As String)
   mvarCostX = vData
End Property

Public Property Get CostX() As String
   CostX = mvarCostX
End Property

Public Property Let CostY(ByVal vData As String)
   mvarCostY = vData
End Property

Public Property Get CostY() As String
   CostY = mvarCostY
End Property

Public Property Let RemarksVisible(ByVal vNewSetting As Boolean)
   mvarRemarksVisible = vNewSetting
End Property

Public Property Get RemarksVisible() As Boolean
   RemarksVisible = mvarRemarksVisible
End Property

Public Property Let AutoPrintinInvoices(ByVal vNewSetting As Boolean)
   mvarAutoPrintinInvoices = vNewSetting
End Property

Public Property Get AutoPrintinInvoices() As Boolean
   AutoPrintinInvoices = mvarAutoPrintinInvoices
End Property

Public Property Let PackingChargesPer(ByVal vData As String)
   mvarPackingChargesPer = vData
End Property

Public Property Get PackingChargesPer() As String
   PackingChargesPer = mvarPackingChargesPer
End Property

Public Property Let HourDifference(ByVal vData As String)
   mvarHourDifference = vData
End Property

Public Property Get HourDifference() As String
   HourDifference = mvarHourDifference
End Property

Public Property Let EmployeeLateRelaxTime(ByVal vData As String)
   mvarEmployeeLateRelaxTime = vData
End Property

Public Property Get EmployeeLateRelaxTime() As String
   EmployeeLateRelaxTime = mvarEmployeeLateRelaxTime
End Property

Public Property Let AdminClssingFinePerOnShort(ByVal vData As String)
   mvarAdminClssingFinePerOnShort = vData
End Property

Public Property Get AdminClssingFinePerOnShort() As String
   AdminClssingFinePerOnShort = mvarAdminClssingFinePerOnShort
End Property

Public Property Let GridRowHeight(ByVal vData As String)
   mvarGridRowHeight = vData
End Property

Public Property Get GridRowHeight() As String
   GridRowHeight = mvarGridRowHeight
End Property

Public Property Let BarCodePrefix(ByVal vData As String)
   mvarBarCodePrefix = vData
End Property

Public Property Get BarCodePrefix() As String
   BarCodePrefix = mvarBarCodePrefix
End Property

Public Property Let RoundfigureInSearchForm(ByVal vData As String)
   mvarRoundfigureInSearchForm = vData
End Property

Public Property Get RoundfigureInSearchForm() As String
   RoundfigureInSearchForm = mvarRoundfigureInSearchForm
End Property

Public Property Let OrganizationVisible(ByVal vNewSetting As Boolean)
   mvarOrganizationVisible = vNewSetting
End Property

Public Property Get OrganizationVisible() As Boolean
  OrganizationVisible = mvarOrganizationVisible
End Property

Public Property Let ShowWholeSaleMargin(ByVal vNewSetting As Boolean)
   mvarShowWholeSaleMargin = vNewSetting
End Property

Public Property Get ShowWholeSaleMargin() As Boolean
   ShowWholeSaleMargin = mvarShowWholeSaleMargin
End Property

Public Property Let ShowPromiseDateInSalaPurchase(ByVal vNewSetting As Boolean)
   mvarShowPromiseDateInSalaPurchase = vNewSetting
End Property

Public Property Get ShowPromiseDateInSalaPurchase() As Boolean
   ShowPromiseDateInSalaPurchase = mvarShowPromiseDateInSalaPurchase
End Property

Public Property Let ShowSyllabus(ByVal vNewSetting As Boolean)
   mvarShowSyllabus = vNewSetting
End Property
Public Property Get ShowSyllabus() As Boolean
   ShowSyllabus = mvarShowSyllabus
End Property

Public Property Let ShowProdProfit(ByVal vNewSetting As Boolean)
   mvarShowProdProfit = vNewSetting
End Property
Public Property Get ShowProdProfit() As Boolean
   ShowProdProfit = mvarShowProdProfit
End Property

Public Property Let OrganizationID(ByVal vData As String)
   mvarOrganizationID = vData
End Property

Public Property Get OrganizationID() As String
   OrganizationID = mvarOrganizationID
End Property
   
Public Property Let DevelopedBy(ByVal vData As String)
   mvarDevelopedBy = vData
End Property

Public Property Get DevelopedBy() As String
   DevelopedBy = mvarDevelopedBy
End Property

Public Property Let BankAccounts(ByVal vData As String)
   mvarBankAccounts = vData
End Property

Public Property Get BankAccounts() As String
   BankAccounts = mvarBankAccounts
End Property

Public Property Let BothQty(ByVal vData As String)
   mvarBothQty = vData
End Property

Public Property Get BothQty() As String
   BothQty = mvarBothQty
End Property

Public Property Let OwnerMobileNo(ByVal vData As String)
   mvarOwnerMobileNo = vData
End Property

Public Property Get OwnerMobileNo() As String
   OwnerMobileNo = mvarOwnerMobileNo
End Property

Public Property Let PrefixPhoneNo(ByVal vData As String)
   mvarPrefixPhoneNo = vData
End Property

Public Property Get PrefixPhoneNo() As String
   PrefixPhoneNo = mvarPrefixPhoneNo
End Property


Public Property Let FromEmail(ByVal vData As String)
   mvarFromEmail = vData
End Property

Public Property Get FromEmail() As String
   FromEmail = mvarFromEmail
End Property

Public Property Let ToEmail(ByVal vData As String)
   mvarToEmail = vData
End Property

Public Property Get ToEmail() As String
   ToEmail = mvarToEmail
End Property

Public Property Let EmailPwd(ByVal vData As String)
   mvarEmailPwd = vData
End Property

Public Property Get EmailPwd() As String
   EmailPwd = mvarEmailPwd
End Property
Public Property Let ActivityActionNo(ByVal vData As String)
   mvarActivityActionNo = vData
End Property

Public Property Get ActivityActionNo() As String
   ActivityActionNo = mvarActivityActionNo
End Property

Public Property Let ByDefaultActionNo(ByVal vData As String)
   mvarByDefaultActionNo = vData
End Property

Public Property Get ByDefaultActionNo() As String
   ByDefaultActionNo = mvarByDefaultActionNo
End Property
Public Property Let ChargesName(ByVal vData As String)
   mvarChargesName = vData
End Property

Public Property Get ChargesName() As String
   ChargesName = mvarChargesName
End Property

Public Property Let SMTPServerAddress(ByVal vData As String)
   mvarSMTPServerAddress = vData
End Property

Public Property Get SMTPServerAddress() As String
   SMTPServerAddress = mvarSMTPServerAddress
End Property

Public Property Let CustomerSalesMessage(ByVal vData As String)
   mvarCustomerSalesMessage = vData
End Property

Public Property Get PortNo() As String
   PortNo = mvarPortNo
End Property

Public Property Let PortNo(ByVal vData As String)
   mvarPortNo = vData
End Property

Public Property Let UseEmail(ByVal vData As Boolean)
   mvarUseEmail = vData
End Property

Public Property Get UseEmail() As Boolean
   UseEmail = mvarUseEmail
End Property
Public Property Let FBRGroupName(ByVal vData As Boolean)
   mvarFBRGroupName = vData
End Property

Public Property Get FBRGroupName() As Boolean
   FBRGroupName = mvarFBRGroupName
End Property
Public Property Let IsFBRedit(ByVal vData As Boolean)
   mvarIsFBRedit = vData
End Property

Public Property Get IsFBRedit() As Boolean
   IsFBRedit = mvarIsFBRedit
End Property

Public Property Let ExportReportASPDF(ByVal vData As Boolean)
   mvarExportReportASPDF = vData
End Property

Public Property Get ExportReportASPDF() As Boolean
   ExportReportASPDF = mvarExportReportASPDF
End Property

Public Property Get CustomerSalesMessage() As String
   CustomerSalesMessage = mvarCustomerSalesMessage
End Property
Public Property Let WebLinkForSMS(ByVal vData As String)
   mvarWebLinkForSMS = vData
End Property

Public Property Get WebLinkForSMS() As String
   WebLinkForSMS = mvarWebLinkForSMS
End Property

Public Property Let UpdateDatabase(ByVal vData As String)
   mvarUpdateDatabase = vData
End Property

Public Property Get UpdateDatabase() As String
   UpdateDatabase = mvarUpdateDatabase
End Property

Public Property Let ReadOnlyStatus(ByVal vData As Boolean)
   mvarReadOnly = vData
End Property

Public Property Get ReadOnlyStatus() As Boolean
   ReadOnlyStatus = mvarReadOnly
End Property

Public Property Let CounterSaleAmount(ByVal vData As Boolean)
   mvarCounterSaleAmount = vData
End Property

Public Property Get CounterSaleAmount() As Boolean
   CounterSaleAmount = mvarCounterSaleAmount
End Property

Public Property Let ChangePriceAtSave(ByVal vData As Boolean)
   mvarChangePriceAtSave = vData
End Property
Public Property Get ShowHistoryofAllCustomer() As Boolean
   ShowHistoryofAllCustomer = mvarShowHistoryofAllCustomer
End Property

Public Property Let ShowHistoryofAllCustomer(ByVal vData As Boolean)
   mvarShowHistoryofAllCustomer = vData
End Property

Public Property Get ChangePriceAtSave() As Boolean
   ChangePriceAtSave = mvarChangePriceAtSave
End Property

Public Property Let ShowCostinSaleInvoice(ByVal vData As Boolean)
   mvarShowCostinSaleInvoice = vData
End Property

Public Property Get ShowCostinSaleInvoice() As Boolean
   ShowCostinSaleInvoice = mvarShowCostinSaleInvoice
End Property

Public Property Let ShowAllInvoiceInSearch(ByVal vData As Boolean)
   mvarShowAllInvoiceInSearch = vData
End Property

Public Property Get ShowAllInvoiceInSearch() As Boolean
   ShowAllInvoiceInSearch = mvarShowAllInvoiceInSearch
End Property
Public Property Let AllowSMSOnSave(ByVal vData As Boolean)
   mvarAllowSMSOnSave = vData
End Property

Public Property Get AllowSMSOnSave() As Boolean
   AllowSMSOnSave = mvarAllowSMSOnSave
End Property
Public Property Let AllowSMSOnDelete(ByVal vData As Boolean)
   mvarAllowSMSOnDelete = vData
End Property

Public Property Get AllowSMSOnDelete() As Boolean
   AllowSMSOnDelete = mvarAllowSMSOnDelete
End Property

Public Property Let AllowSMSOnClear(ByVal vData As Boolean)
   mvarAllowSMSOnClear = vData
End Property
Public Property Get AllowSMSOnClear() As Boolean
   AllowSMSOnClear = mvarAllowSMSOnClear
End Property
Public Property Let AllowSMSOnLogin(ByVal vData As Boolean)
   mvarAllowSMSOnLogin = vData
End Property
Public Property Get AllowSMSOnLogin() As Boolean
   AllowSMSOnLogin = mvarAllowSMSOnLogin
End Property

Public Property Let AllowSMSWithDetail(ByVal vData As Boolean)
   mvarAllowSMSWithDetail = vData
End Property

Public Property Get AllowSMSWithDetail() As Boolean
   AllowSMSWithDetail = mvarAllowSMSWithDetail
End Property
Public Property Let AllowSMSThroughDevice(ByVal vData As Boolean)
   mvarAllowSMSThroughDevice = vData
End Property

Public Property Get AllowSMSThroughDevice() As Boolean
   AllowSMSThroughDevice = mvarAllowSMSThroughDevice
End Property
Public Property Let ChangeQtyOnChangedPrice(ByVal vData As Boolean)
   mvarAllowChangeQtyOnChangedPrice = vData
End Property

Public Property Get ChangeQtyOnChangedPrice() As Boolean
   ChangeQtyOnChangedPrice = mvarAllowChangeQtyOnChangedPrice
End Property
Public Property Let HeaderInfoNotClear(ByVal vData As Boolean)
   mvarHeaderInfoNotClear = vData
End Property

Public Property Get HeaderInfoNotClear() As Boolean
   HeaderInfoNotClear = mvarHeaderInfoNotClear
End Property

Public Property Let AutoMoveGridWhenSerialEntered(ByVal vData As Boolean)
   mvarAutoMoveGridWhenSerialEntered = vData
End Property

Public Property Get AutoMoveGridWhenSerialEntered() As Boolean
   AutoMoveGridWhenSerialEntered = mvarAutoMoveGridWhenSerialEntered
End Property

Public Property Let SerialCompulsoryinInvoice(ByVal vData As Boolean)
   mvarSerialCompulsoryinInvoice = vData
End Property

Public Property Get SerialCompulsoryinInvoice() As Boolean
   SerialCompulsoryinInvoice = mvarSerialCompulsoryinInvoice
End Property

Public Property Let isEntryDate(ByVal vData As Boolean)
   mvarIsEntryDate = vData
End Property

Public Property Get CurrentDateDataEntry() As Boolean
   CurrentDateDataEntry = mvarCurrentDateDataEntry
End Property

Public Property Let CurrentDateDataEntry(ByVal vData As Boolean)
   mvarCurrentDateDataEntry = vData
End Property

Public Property Get isEntryDate() As Boolean
   isEntryDate = mvarIsEntryDate
End Property
Public Property Let FromDate(ByVal vData As String)
   mvarFromDate = vData
End Property

Public Property Get FromDate() As String
   FromDate = mvarFromDate
End Property
Public Property Let ToDate(ByVal vData As String)
   mvarToDate = vData
End Property

Public Property Get ToDate() As String
   ToDate = mvarToDate
End Property

Public Property Let ShowStockFromTableGridDataMovement(ByVal vData As Boolean)
   mvarShowStockFromTableGridDataMovement = vData
End Property

Public Property Get ShowStockFromTableGridDataMovement() As Boolean
   ShowStockFromTableGridDataMovement = mvarShowStockFromTableGridDataMovement
End Property

Public Property Let ShowBarcodeProductSearch(ByVal vData As Boolean)
   mvarShowBarcodeProductSearch = vData
End Property

Public Property Get ShowBarcodeProductSearch() As Boolean
   ShowBarcodeProductSearch = mvarShowBarcodeProductSearch
End Property
Public Property Let OrganizationMandatory(ByVal vData As Boolean)
   mvarOrganizationMandatory = vData
End Property

Public Property Get OrganizationMandatory() As Boolean
   OrganizationMandatory = mvarOrganizationMandatory
End Property

Public Property Let EmployeeMandatory(ByVal vData As Boolean)
   mvarEmployeeMandatory = vData
End Property

Public Property Get EmployeeMandatory() As Boolean
   EmployeeMandatory = mvarEmployeeMandatory
End Property

Public Property Let TableIDMandatory(ByVal vData As Boolean)
   mvarTableIDMandatory = vData
End Property

Public Property Get TableIDMandatory() As Boolean
   TableIDMandatory = mvarTableIDMandatory
End Property

Public Property Let UpdateStockSaleBodyInsert(ByVal vData As Boolean)
   mvarUpdateStockSaleBodyInsert = vData
End Property

Public Property Get UpdateStockSaleBodyInsert() As Boolean
   UpdateStockSaleBodyInsert = mvarUpdateStockSaleBodyInsert
End Property

Public Property Let isShowPublisher(ByVal vData As Boolean)
   mvarisShowPublisher = vData
End Property

Public Property Get isShowPublisher() As Boolean
   isShowPublisher = mvarisShowPublisher
End Property

Public Property Let isShowListPrice(ByVal vData As Boolean)
   mvarisShowListPrice = vData
End Property

Public Property Get isShowListPrice() As Boolean
   isShowListPrice = mvarisShowListPrice
End Property

Public Property Let IsShowPrintYesOrNo(ByVal vData As Boolean)
   mvarIsShowPrintYesOrNo = vData
End Property

Public Property Get IsShowPrintYesOrNo() As Boolean
   IsShowPrintYesOrNo = mvarIsShowPrintYesOrNo
End Property


Public Property Let isShowDepartment(ByVal vData As Boolean)
   mvarisShowDepartment = vData
End Property

Public Property Get isShowDepartment() As Boolean
   isShowDepartment = mvarisShowDepartment
End Property

Public Property Let isShowSubDepartment(ByVal vData As Boolean)
   mvarisShowSubDepartment = vData
End Property

Public Property Get isShowSubDepartment() As Boolean
   isShowSubDepartment = mvarisShowSubDepartment
End Property
Public Property Let isShowSeason(ByVal vData As Boolean)
   mvarisShowSeason = vData
End Property

Public Property Get isShowSeason() As Boolean
   isShowSeason = mvarisShowSeason
End Property
Public Property Let isShowItemDesc(ByVal vData As Boolean)
   mvarisShowItemDesc = vData
End Property

Public Property Get isShowItemDesc() As Boolean
   isShowItemDesc = mvarisShowItemDesc
End Property

Public Property Let isShowOther(ByVal vData As Boolean)
   mvarisShowOther = vData
End Property

Public Property Get isShowOther() As Boolean
   isShowOther = mvarisShowOther
End Property

Public Property Let isShowVendor(ByVal vData As Boolean)
   mvarisShowVendor = vData
End Property

Public Property Get isShowVendor() As Boolean
   isShowVendor = mvarisShowVendor
End Property

Public Property Let ShowColourSize(ByVal vData As Boolean)
   mvarShowColourSize = vData
End Property

Public Property Get ShowColourSize() As Boolean
   ShowColourSize = mvarShowColourSize
End Property

Public Property Let AllowBothPackingsareSame(ByVal vData As Boolean)
   mvarAllowBothPackingsareSame = vData
End Property

Public Property Get AllowBothPackingsareSame() As Boolean
   AllowBothPackingsareSame = mvarAllowBothPackingsareSame
End Property

Public Property Let ShowSavedStock(ByVal vData As Boolean)
   mvarShowSavedStock = vData
End Property

Public Property Get ShowSavedStock() As Boolean
   ShowSavedStock = mvarShowSavedStock
End Property

Public Property Let ShowAllPrices(ByVal vData As Boolean)
   mvarShowAllPrices = vData
End Property

Public Property Get ShowAllPrices() As Boolean
   ShowAllPrices = mvarShowAllPrices
End Property


Public Property Get ShowSession() As Boolean
   ShowSession = mvarShowSession
End Property

Public Property Let ShowSession(ByVal vData As Boolean)
   mvarShowSession = vData
End Property

Public Property Get ShowTradeOffer() As Boolean
   ShowTradeOffer = mvarShowTradeOffer
End Property

Public Property Let ShowTradeOffer(ByVal vData As Boolean)
   mvarShowTradeOffer = vData
End Property

Public Property Get ShowExpiryInvoice() As Boolean
   ShowExpiryInvoice = mvarShowExpiryInvoice
End Property

Public Property Let ShowExpiryInvoice(ByVal vData As Boolean)
   mvarShowExpiryInvoice = vData
End Property

Public Property Get IsPortrait() As Boolean
   IsPortrait = mvarIsPortrait
End Property

Public Property Let IsPortrait(ByVal vData As Boolean)
   mvarIsPortrait = vData
End Property

Public Property Get IsLegal() As Boolean
   IsLegal = mvarIsLegal
End Property

Public Property Let IsLegal(ByVal vData As Boolean)
   mvarIsLegal = vData
End Property
Public Property Get EmployeeCommision() As Boolean
   EmployeeCommision = mvarEmployeeCommision
End Property

Public Property Let EmployeeCommision(ByVal vData As Boolean)
   mvarEmployeeCommision = vData
End Property
Public Property Get SalePriceLessThanPurchase() As Boolean
   SalePriceLessThanPurchase = mvarSalePriceLessThanPurchase
End Property

Public Property Let SalePriceLessThanPurchase(ByVal vData As Boolean)
   mvarSalePriceLessThanPurchase = vData
End Property


Public Property Let IsGrossQty(ByVal vData As Boolean)
   mvarIsGrossQty = vData
End Property

Public Property Get IsGrossQty() As Boolean
   IsGrossQty = mvarIsGrossQty
End Property

Public Property Let ShowBonus(ByVal vData As Boolean)
   mvarShowBonus = vData
End Property

Public Property Get ShowBonus() As Boolean
   ShowBonus = mvarShowBonus
End Property

Public Property Let ShowOffer(ByVal vData As Boolean)
   mvarShowOffer = vData
End Property

Public Property Get ShowOffer() As Boolean
   ShowOffer = mvarShowOffer
End Property

Public Property Let ShowSaleTax(ByVal vData As Boolean)
   mvarShowSaleTax = vData
End Property

Public Property Get ShowSaleTax() As Boolean
   ShowSaleTax = mvarShowSaleTax
End Property

Public Property Let ShowSC(ByVal vData As Boolean)
   mvarShowSC = vData
End Property

Public Property Get ShowSC() As Boolean
   ShowSC = mvarShowSC
End Property

Public Property Let ShowBatchPrint(ByVal vData As Boolean)
   mvarShowBatchPrint = vData
End Property

Public Property Get ShowBatchPrint() As Boolean
   ShowBatchPrint = mvarShowBatchPrint
End Property

Public Property Let ChangeQtyPack(ByVal vData As Boolean)
   mvarChangeQtyPack = vData
End Property

Public Property Get ShowBarCodeQty() As Boolean
   ShowBarCodeQty = mvarShowBarCodeQty
End Property

Public Property Let ShowBarCodeQty(ByVal vData As Boolean)
   mvarShowBarCodeQty = vData
End Property

Public Property Get ShowChangePriceOnSavePI() As Boolean
   ShowChangePriceOnSavePI = mvarShowChangePriceOnSavePI
End Property

Public Property Let ShowChangePriceOnSavePI(ByVal vData As Boolean)
   mvarShowChangePriceOnSavePI = vData
End Property

Public Property Get ShowChangeRetailInPurchaseInvoice() As Boolean
   ShowChangeRetailInPurchaseInvoice = mvarShowChangeRetailInPurchaseInvoice
End Property

Public Property Let ShowChangeRetailInPurchaseInvoice(ByVal vData As Boolean)
   mvarShowChangeRetailInPurchaseInvoice = vData
End Property

Public Property Get ShowGrandTotalinSearch() As Boolean
   ShowGrandTotalinSearch = mvarShowGrandTotalinSearch
End Property

Public Property Let ShowGrandTotalinSearch(ByVal vData As Boolean)
   mvarShowGrandTotalinSearch = vData
End Property

Public Property Get ChangeQtyPack() As Boolean
   ChangeQtyPack = mvarChangeQtyPack
End Property

Public Property Let EitherPackORLooseEnter(ByVal vData As Boolean)
   mvarEitherPackORLooseEnter = vData
End Property

Public Property Get EitherPackORLooseEnter() As Boolean
   EitherPackORLooseEnter = mvarEitherPackORLooseEnter
End Property

Public Property Get IsRoundFigure() As Boolean
   IsRoundFigure = mvarIsRoundFigure
End Property

Public Property Let IsRoundFigure(ByVal vData As Boolean)
   mvarIsRoundFigure = vData
End Property

Public Property Get IsSingleBarcode() As Boolean
   IsSingleBarcode = mvarIsSingleBarcode
End Property

Public Property Let IsSingleBarcode(ByVal vData As Boolean)
   mvarIsSingleBarcode = vData
End Property

Public Property Get SectorCompulsory() As Boolean
   SectorCompulsory = mvarSectorCompulsory
End Property

Public Property Let SectorCompulsory(ByVal vData As Boolean)
   mvarSectorCompulsory = vData
End Property

Public Property Get RemarksCompulsory() As Boolean
   RemarksCompulsory = mvarRemarksCompulsory
End Property

Public Property Let RemarksCompulsory(ByVal vData As Boolean)
   mvarRemarksCompulsory = vData
End Property

Public Property Get DivideRetailWithPacking() As Boolean
   DivideRetailWithPacking = mvarDivideRetailWithPacking
End Property

Public Property Let DivideRetailWithPacking(ByVal vData As Boolean)
   mvarDivideRetailWithPacking = vData
End Property

Public Property Get DisableQuantityinPOS() As Boolean
   DisableQuantityinPOS = mvarDisableQuantityinPOS
End Property

Public Property Let DisableQuantityinPOS(ByVal vData As Boolean)
   mvarDisableQuantityinPOS = vData
End Property

Public Property Let UseMultipleStore(ByVal vData As Boolean)
   mvarUseMultipleStore = vData
End Property

Public Property Get UseMultipleStore() As Boolean
   UseMultipleStore = mvarUseMultipleStore
End Property
Public Property Let PLSamePR(ByVal vData As Boolean)
   mvarPLSamePR = vData
End Property
Public Property Get PLSamePR() As Boolean
   PLSamePR = mvarPLSamePR
End Property
Public Property Let AdminClosingSaveWhenUserClosingSaved(ByVal vData As Boolean)
   mvarAdminClosingSaveWhenUserClosingSaved = vData
End Property
Public Property Get AdminClosingSaveWhenUserClosingSaved() As Boolean
   AdminClosingSaveWhenUserClosingSaved = mvarAdminClosingSaveWhenUserClosingSaved
End Property

Public Property Let LockPurPrice(ByVal vData As Boolean)
   mvarLockPurPrice = vData
End Property
Public Property Get LockPurPrice() As Boolean
   LockPurPrice = mvarLockPurPrice
End Property

Public Property Let ProductSearchWithStore(ByVal vData As Boolean)
   mvarProductSearchWithStore = vData
End Property
Public Property Get ProductSearchWithStore() As Boolean
   ProductSearchWithStore = mvarProductSearchWithStore
End Property

Public Property Let ShowAllStoreStock(ByVal vData As Boolean)
   mvarShowAllStoreStock = vData
End Property
Public Property Get ShowAllStoreStock() As Boolean
   ShowAllStoreStock = mvarShowAllStoreStock
End Property

Public Property Let SearchCodeInGrid(ByVal vData As Boolean)
   mvarSearchCodeInGrid = vData
End Property
Public Property Get UsePurPrice() As Boolean
   UsePurPrice = mvarUsePurPrice
End Property

Public Property Let UsePurPrice(ByVal vData As Boolean)
   mvarUsePurPrice = vData
End Property

Public Property Get ChangeTransactionDate() As Boolean
   ChangeTransactionDate = mvarChangeTransactionDate
End Property

Public Property Let ChangeTransactionDate(ByVal vData As Boolean)
   mvarChangeTransactionDate = vData
End Property
Public Property Get UseBin() As Boolean
   UseBin = mvarUseBin
End Property

Public Property Let UseBin(ByVal vData As Boolean)
   mvarUseBin = vData
End Property
Public Property Get AttendanceNextDayOut() As Boolean
   AttendanceNextDayOut = mvarAttendanceNextDayOut
End Property

Public Property Let AllowDiscountOnSaleDistribution(ByVal vData As Boolean)
   mvarAllowDiscountOnSaleDistribution = vData
End Property

Public Property Get AllowDiscountOnSaleDistribution() As Boolean
   AllowDiscountOnSaleDistribution = mvarAllowDiscountOnSaleDistribution
End Property

Public Property Let DefaultCustomer(ByVal vData As Boolean)
   mvarDefaultCustomer = vData
End Property

Public Property Get DefaultCustomer() As Boolean
   DefaultCustomer = mvarDefaultCustomer
End Property

Public Property Let ClientDate(ByVal vData As Boolean)
   mvarClientDate = vData
End Property

Public Property Get ClientDate() As Boolean
   ClientDate = mvarClientDate
End Property

Public Property Let IsAllowBarCodeQtyInpurchaseQty(ByVal vData As Boolean)
   mvarIsAllowBarCodeQtyInpurchaseQty = vData
End Property

Public Property Get IsAllowBarCodeQtyInpurchaseQty() As Boolean
   IsAllowBarCodeQtyInpurchaseQty = mvarIsAllowBarCodeQtyInpurchaseQty
End Property

Public Property Let AttendanceNextDayOut(ByVal vData As Boolean)
   mvarAttendanceNextDayOut = vData
End Property

Public Property Get SearchCodeInGrid() As Boolean
   SearchCodeInGrid = mvarSearchCodeInGrid
End Property

Public Property Let UsePasswordForm(ByVal vData As Boolean)
   mvarUsePasswordForm = vData
End Property
Public Property Get UsePasswordForm() As Boolean
   UsePasswordForm = mvarUsePasswordForm
End Property

Public Property Let ShowPurchaseProfit(ByVal vData As Boolean)
   mvarShowPurchaseProfit = vData
End Property

Public Property Get ShowPurchaseProfit() As Boolean
   ShowPurchaseProfit = mvarShowPurchaseProfit
End Property

Public Property Let ShowBarcodeDesc(ByVal vData As Boolean)
   mvarShowBarcodeDesc = vData
End Property

Public Property Get ShowBarcodeDesc() As Boolean
   ShowBarcodeDesc = mvarShowBarcodeDesc
End Property

Public Property Let ShowDiscPrice(ByVal vData As Boolean)
   mvarShowDiscPrice = vData
End Property
Public Property Get ShowDiscPrice() As Boolean
   ShowDiscPrice = mvarShowDiscPrice
End Property

Public Property Let ShowDuplicatePrint(ByVal vData As Boolean)
   mvarShowDuplicatePrint = vData
End Property

Public Property Get ShowDuplicatePrint() As Boolean
   ShowDuplicatePrint = mvarShowDuplicatePrint
End Property

Public Property Let RunnngLastPrice(ByVal vData As Boolean)
   mvarRunnngLastPrice = vData
End Property

Public Property Get RunnngLastPrice() As Boolean
   RunnngLastPrice = mvarRunnngLastPrice
End Property


Public Property Get ShowDispatchDate() As Boolean
   ShowDispatchDate = mvarShowDispatchDate
End Property

Public Property Let ShowDispatchDate(ByVal vData As Boolean)
   mvarShowDispatchDate = vData
End Property

Public Property Get ShowReSale() As Boolean
   ShowReSale = mvarShowReSale
End Property

Public Property Let ShowReSale(ByVal vData As Boolean)
   mvarShowReSale = vData
End Property

Public Property Get ShowDiscPurPrice() As Boolean
   ShowDiscPurPrice = mvarShowDiscPurPrice
End Property

Public Property Let ShowDiscPurPrice(ByVal vData As Boolean)
   mvarShowDiscPurPrice = vData
End Property

Public Property Get ShowStockPriceChecker() As Boolean
   ShowStockPriceChecker = mvarShowStockPriceChecker
End Property

Public Property Let ShowStockPriceChecker(ByVal vData As Boolean)
   mvarShowStockPriceChecker = vData
End Property

Public Property Get AllowNegativeOrder() As Boolean
   AllowNegativeOrder = mvarAllowNegativeOrder
End Property
Public Property Let AllowNegativeOrder(ByVal vData As Boolean)
   mvarAllowNegativeOrder = vData
End Property

Public Property Get ShowPurPrice() As Boolean
   ShowPurPrice = mvarShowPurPrice
End Property
Public Property Let ShowPurPrice(ByVal vData As Boolean)
   mvarShowPurPrice = vData
End Property

Public Property Get ShowAddBarCode() As Boolean
   ShowAddBarCode = mvarShowAddBarCode
End Property
Public Property Let ShowAddBarCode(ByVal vData As Boolean)
   mvarShowAddBarCode = vData
End Property


Public Property Get ShowRetailPriceStockRegister() As Boolean
   ShowRetailPriceStockRegister = mvarShowRetailPriceStockRegister
End Property

Public Property Let ShowRetailPriceStockRegister(ByVal vData As Boolean)
   mvarShowRetailPriceStockRegister = vData
End Property

Public Property Get ShowMultiBranches() As Boolean
   ShowMultiBranches = mvarShowMultiBranches
End Property

Public Property Let ShowMultiBranches(ByVal vData As Boolean)
   mvarShowMultiBranches = vData
End Property

Public Property Get TimeWiseReport() As Boolean

   TimeWiseReport = mvarTimeWiseReport
End Property

Public Property Let TimeWiseReport(ByVal vData As Boolean)
   mvarTimeWiseReport = vData
End Property
Public Property Get CheckStockOnSave() As Boolean
   CheckStockOnSave = mvarCheckStockOnSave
End Property

Public Property Let CheckStockOnSave(ByVal vData As Boolean)
   mvarCheckStockOnSave = vData
End Property
Public Property Get ShowDisc2() As Boolean
   ShowDisc2 = mvarShowDisc2
End Property

Public Property Let ShowDisc2(ByVal vData As Boolean)
   mvarShowDisc2 = vData
End Property
Public Property Get TermAllowZero() As Boolean
   TermAllowZero = mvarTermAllowZero
End Property

Public Property Let TermAllowZero(ByVal vData As Boolean)
   mvarTermAllowZero = vData
End Property


