Attribute VB_Name = "MdlRegistry"
Option Explicit
Public mvarCompanyName As String 'local copy
Public mvarCompanyShortName As String 'local copy
Public mvarCompanyAddress As String 'local copy
Public mvarCompanyCity As String 'local copy
Public mvarCompanyPhoneNo As String 'local copy
Public mvarCompanyEMail As String 'local copy

Public mvarStoreID As String 'local copy
Public mvarSeperateProductWithPrice As Boolean 'local copy
Public mvarSeperateProductInPOS As Boolean 'local copy
Public mvarAllowNegativeStockInBarcodes As Boolean 'local copy
Public mvarBatchNoVisible As Boolean 'local copy
Public mvarStoreVisible As Boolean 'local copy
Public mvarCustomerTypeVisible As Boolean 'local copy
Public mvarOrganizationID As String 'local copy
Public mvarOrganizationVisible As Boolean 'local copy
Public mvarTableVisible As Boolean 'local copy
Public mvarSaleOrderVisible As Boolean 'local copy
Public mvarSystemDate As Boolean 'local copy
Public mvarSaleInProduction As Boolean 'local copy
Public mvarProperCase As Boolean 'local copy
Public mvarDiscAllowed As Boolean 'local copy
Public mvarProductSearchOpenInPreviousState As Boolean 'local copy
Public mvarShowRetailinPurchaseReturnPrint As Boolean 'local copy
Public mvarPreviewSaleInoice As Boolean 'local copy
Public mvarProdDesc1 As String 'local copy
Public mvarSearchDateDifference As String 'local copy
Public mvarConnectionTimeout As String 'local copyFemp
Public mvarBlankFooter As String 'local copy
Public mvarAutoPrintSaleOrder As Boolean 'local copy
Public mvarPrintKitchenInoices As Boolean 'local copy
Public mvarAutoEnterBeforeQty As Boolean 'local copy
Public mvarEmptyEnterGotoSave As Boolean 'local copy
Public mvarInvType As Boolean 'local copy
Public mvarShowWarrantyinSaleInvoice As Boolean 'local copy
Public mvarIsSingleBarcode As Boolean 'local copy
Public mvarShowPurchaseProfit As Boolean


Public mvarHideAutoPrint As Boolean 'local copy
Public mvarAllowOrderByCodeinInvoices As Boolean 'local copy
Public mvarAllowUrduProduct As Boolean 'local copy
Public mvarAllowFBRContinuousNo As Boolean 'local copy
Public mvarAllowEmployeProductWise As Boolean 'local copy
Public mvarAllowStoreProductWise As Boolean 'local copy
Public mvarAllowManualBillNoValidation As Boolean 'local copy
Public mvarAllowContinuousBillNo As Boolean 'local copy
Public mvarAllowMonthlyBillNo As Boolean 'local copy
Public mvarAllowDailyBillNo As Boolean 'local copy
Public mvarSetEnterKeyGridStockAdjustment As Boolean 'local copy
Public mvarSaveAsNewBill As Boolean 'local copy
Public mvarAfterRowEditFocusNextGridLine As Boolean
Public mvarHideClearButton As Boolean
Public mvarAllowUrduRemarks As Boolean 'local copy

Public mvarShowSerial As Boolean 'local copy
Public mvarShowCode As Boolean 'local copy
Public mvarShowLastInvoiceMsgAtSave As Boolean 'local copy

Public mvarShowWeightedAvgOption As Boolean 'local copy
Public mvarShowMovingAvgOption As Boolean 'local copy
Public mvarShowLastPriceOption As Boolean 'local copy
Public mvarShowRawMaterialProductInSaleInvoices As Boolean 'local copy
Public mvarShowManufacturingProductInInvoices As Boolean 'local copy

Public mvarShowOrganizationWiseStock As Boolean 'local copy
  
Public mvarAllowSameStore As Boolean 'local copy
Public mvarNegativeSale As Boolean 'local copy
Public mvarDevelopedBy  As String 'local copy
Public mvarBankAccounts  As String 'local copy
Public mvarBothQty As String 'local copy
Public mvarUpdateDatabase As String 'local copy
Public mvarChangePriceAtSave As String 'local copy
Public mvarReadOnly  As Boolean 'local copy
Public mvarCounterSaleAmount As String 'local copy
Public mvarShowCostinSaleInvoice As String 'local copy
Public mvarShowAllInvoiceInSearch As String 'local copy
Public mvarShowChangeRetailInPurchaseInvoice As String 'local copy

Public mvarBankMachineID As String 'local copy
Public mvarCashReceived As String 'local copy
Public mvarDuplicateCode As String 'local copy
Public mvarAddSpace As String 'local copy
Public mvarCostVisible As String 'local copy
Public mvarLineSpacing As String 'local copy
Public mvarStatement As String 'local copy
Public mvarChangePrice As String 'local copy
Public mvarLastRateVisible As String 'local copy
Public mvarEmpVisible As String 'local copy
Public mvarHidePurchaseAmount As String 'local copy
Public mvarHideSaleAmount As String 'local copy
Public mvarShowBankInTransection As String 'local copy
Public mvarMemberVisible As String 'local copy
Public mvarLaserPrintofSaleInvoice As String 'local copy
Public mvarPrintHeadersSaleInvoice As String 'local copy
Public mvarMainStoreID As String 'local copy
Public mvarNoofPrints As String 'local copy
Public mvarMemberMin As String 'local copy
Public mvarMemberMax As String 'local copy
Public mvarTag As String 'local copy
Public mvarQuantityinBarcodes As String 'local copy
Public mvarFreightVisible As String 'local copy
Public mvarOrderStatement As String 'local copy
Public mvarDeviceName, mvarDeviceName2 As String 'local copy
Public mvarDriverName, mvarDriverName2 As String 'local copy
Public mvarPort, mvarPort2 As String 'local copy
Public mvarManualBillNoVisible As String 'local copy
Public mvarPreviousBalanceVisible As String 'local copy
Public mvarX As String 'local copy
Public mvarY As String 'local copy
Public mvarRemarksVisible As String 'local copy
Public mvarAutoPrintinInvoices As String 'local copy
Public mvarHourDifference As String 'local copy
Public mvarPackingChargesPer As String 'local copy
Public mvarAutoApplyPartyLastPrice As String 'local copy
Public mvarAutoApplyPartyLastDiscount As String 'local copy
Public mvarAlertAllocateProduct As String 'local copy

Public mvarOwnerMobileNo As String 'local copy
Public mvarPrefixPhoneNo As String 'local copy
Public mvarCustomerSalesMessage As String 'local copy
Public mvarWebLinkForSMS As String 'local copy
Public mvarChargesName As String

Public mvarAutoEnterQtyintoGridSaleInvoice As String 'local copy

Public mvarShowWholeSaleMargin As Boolean
Public mvarShowPromiseDateInSalaPurchase As Boolean
Public mvarShowSyllabus As Boolean
Public mvarShowProdProfit As Boolean


Public mvarAllowSMSOnSave As Boolean
Public mvarAllowSMSOnDelete As Boolean
Public mvarAllowSMSOnClear As Boolean
Public mvarAllowSMSOnLogin As Boolean
Public mvarAllowSMSWithDetail As Boolean
Public mvarAllowSMSThroughDevice As Boolean

Public mvarAllowChangeQtyOnChangedPrice As Boolean
Public mvarHeaderInfoNotClear As Boolean

Public mvarAutoMoveGridWhenSerialEntered As Boolean
Public mvarSerialCompulsoryinInvoice As Boolean

Public mvarCurrentDateDataEntry As Boolean
Public mvarIsEntryDate As Boolean
Public mvarFromDate As String
Public mvarToDate As String

Public mvarShowStockFromTableGridDataMovement As Boolean
Public mvarShowBarcodeProductSearch As Boolean
Public mvarOrganizationMandatory As Boolean
Public mvarEmployeeMandatory As Boolean
Public mvarTableIDMandatory As Boolean
Public mvarisShowPublisher As Boolean
Public mvarisShowListPrice As Boolean
Public mvarisShowDepartment As Boolean
Public mvarisShowSubDepartment As Boolean
Public mvarisShowSeason As Boolean
Public mvarisShowItemDesc As Boolean
Public mvarisShowOther As Boolean
Public mvarisShowVendor As Boolean
Public mvarShowColourSize As Boolean
Public mvarIsGrossQty As Boolean
Public mvarShowSavedStock As Boolean
Public mvarShowAllPrices As Boolean
Public mvarShowSession As Boolean
Public mvarUpdateStockSaleBodyInsert As Boolean
Public mvarEmployeeCommision As Boolean
Public mvarShowTradeOffer As Boolean
Public mvarSalePriceLessThanPurchase As Boolean
Public mvarShowExpiryInvoice As Boolean
Public mvarIsPortrait As Boolean
Public mvarIsLegal As Boolean

Public mvarShowBonus As Boolean
Public mvarShowOffer As Boolean
Public mvarShowSaleTax As Boolean
Public mvarShowSC As Boolean
Public mvarShowBatchPrint As Boolean
Public mvarChangeQtyPack As Boolean
Public mvarEitherPackORLooseEnter As Boolean
Public mvarShowBarCodeQty As Boolean
Public mvarShowHistoryofAllCustomer As Boolean
Public mvarIsRoundFigure As Boolean
Public mvarShowChangePriceOnSavePI As Boolean
Public mvarShowGrandTotalinSearch As Boolean
Public mvarRemarksCompulsory As Boolean
Public mvarSectorCompulsory As Boolean
Public mvarDivideRetailWithPacking As Boolean
Public mvarDisableQuantityinPOS As Boolean
Public mvarShowDispatchDate As Boolean
Public mvarUseMultipleStore As Boolean
Public mvarShowRetailPriceStockRegister As Boolean
Public mvarTimeWiseReport As Boolean
Public mvarShowReSale As Boolean
Public mvarPLSamePR As Boolean
Public mvarBarCodePrefix As String 'local copy
Public mvarEmployeeLateRelaxTime As String 'local copy
Public mvarAdminClosingSaveWhenUserClosingSaved As Boolean
Public mvarCostX As String 'local copy
Public mvarCostY As String 'local copy
Public mvarCheckStockOnSave As Boolean
Public mvarLockPurPrice As Boolean
Public mvarShowDiscPurPrice As Boolean
Public mvarShowStockPriceChecker As Boolean
Public mvarAllowNegativeOrder As Boolean
Public mvarShowAddBarCode As Boolean
Public mvarAdminClssingFinePerOnShort As String
Public mvarShowPurPrice As Boolean
Public mvarGridRowHeight As String
Public mvarProductSearchWithStore As Boolean
Public mvarShowAllStoreStock As Boolean
Public mvarSearchCodeInGrid As Boolean
Public mvarUsePasswordForm As Boolean
Public mvarUsePurPrice As Boolean
Public mvarRoundfigureInSearchForm As String
Public mvarAllowBothPackingsareSame As Boolean 'local copy
Public mvarChangeTransactionDate As Boolean 'local copy
Public mvarUseBin As Boolean

Public mvarUseEmail As Boolean
Public mvarExportReportASPDF As Boolean
Public mvarFromEmail As String
Public mvarToEmail As String
Public mvarEmailPwd As String
Public mvarSMTPServerAddress As String
Public mvarPortNo As String
Public mvarActivityActionNo As String
Public mvarByDefaultActionNo As String

Public mvarShowBarcodeDesc As Boolean
Public mvarShowDiscPrice As Boolean
Public mvarAttendanceNextDayOut As Boolean
Public mvarAllowDiscountOnSaleDistribution As Boolean
Public mvarShowMultiBranches As Boolean
Public mvarDefaultCustomer As Boolean
Public mvarClientDate As Boolean
Public mvarIsAllowBarCodeQtyInpurchaseQty As Boolean
Public mvarIsShowPrintYesOrNo As Boolean
Public mvarShowDisc2 As Boolean
Public mvarTermAllowZero As Boolean
Public mvarShowDuplicatePrint As Boolean
Public mvarRunnngLastPrice As Boolean
Public mvarFBRGroupName As Boolean
Public mvarIsFBRedit As Boolean
