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

Public Sub InitializeSecurity(objSec As UserSecurity.ClsUserSecurity)
    Declarations.SetSecurityReference objSec
End Sub

Public Sub OrganizationForm()
    DefOrganization.Show
End Sub

Public Sub BankMachinesForm()
    DefBankMachines.Show
End Sub

Public Sub ChangeCategoriesForm()
    FrmChangeCategories.Show
End Sub

Public Sub MembersForm()
    DefMembers.Show
End Sub

Public Sub MemberTypesForm()
    DefMemberTypes.Show
End Sub

Public Sub MembersDiscountForm()
    FrmMembersDiscount.Show
End Sub
Public Sub CommisionDiscForm()
    DefCommissionDiscRange.Show
End Sub
Public Sub ExpiryDayColorForm()
    DefExpiryDayColor.Show
End Sub

Public Sub BrandsForm()
    DefBrands.Show
End Sub

Public Sub CompaniesForm()
    DefCompanies.Show
End Sub

Public Sub CustomersForm()
    DefCustomers.Show
End Sub

Public Sub ChangePriceForm()
    FrmChangePrice.Show
End Sub

Public Sub DepartmentsForm()
    DefDepartments.Show
End Sub

Public Sub EmpDepartmentsForm()
    DefEmpDepartments.Show
End Sub
Public Sub DesignationsForm()
    DefDesignations.Show
End Sub

Public Sub EmpolyeesForm()
    DefEmpolyees.Show
End Sub

Public Sub GroupsForm()
    DefGroups.Show
End Sub

Public Sub SetCurrentStockForm()
    FrmSetCurrentStock.Show
End Sub

Public Sub ChangeProductNameForm()
'    FrmChangeProductName.Show
End Sub

Public Sub OpeningStockForm()
    FrmOpeningStockInvoice.Show
End Sub

Public Sub PackingsForm()
    DefPackings.Show
End Sub

Public Sub ProductsForm()
    DefProducts.Show
End Sub
Public Sub ProductsDetailForm()
    DefProductsDetail.Show
End Sub

Public Sub FormulaInfoForm()
    DefFormulaInfo.Show
End Sub

Public Sub ProductProcessInfoForm()
    'DefProductProcessInfo.Show
End Sub

Public Sub StoresForm()
    DefStores.Show
End Sub

Public Sub SubGroupsForm()
    DefSubGroups.Show
End Sub

Public Sub RacksForm()
    DefRacks.Show
End Sub

Public Sub UnionInfoForm()
'    DefUnionInfo.Show
End Sub

Public Sub VendorsForm()
    DefVendors.Show
End Sub

Public Sub SingleBarcodeForm()
    FrmSingleBarcodes.Show
End Sub

Public Sub ServiceProductsForm()
    DefSProducts.Show
End Sub

Public Sub MultiBarcodesForm()
    Dim a As New FrmMultiBarcodes
    a.Show
End Sub

Public Sub UnitsForm()
    DefUnits.Show
End Sub

Public Sub ZonesForm()
    DefZones.Show
End Sub

Public Sub SectorsForm()
    DefSectors.Show
End Sub

Public Sub SalaryForm()
'    FrmSalary.Show
End Sub

Public Sub CustomProductsandMeasurementsForm()
   DefCustomProductsandMeasurements.Show
End Sub

Public Sub CustomerTypesForm()
   DefCustomerTypes.Show
End Sub

Public Sub StockAdjustmentForm()
   FrmStockAdjustment.Show
End Sub

Public Sub StockLimitStoreWiseForm()
   FrmDefineStockLimitStoreWise.Show
End Sub

Public Sub DefineStockLimitForm()
   FrmDefineStockLimit.Show
End Sub

Public Sub OpeningBatchExpiryForm()
    FrmOpeningBatchExpiry.Show
End Sub

Public Sub ProductOfferForm()
    DefProductOffer.Show
End Sub

Public Sub ReferencesForm()
   DefReferences.Show
End Sub

Public Sub InstallmentCustomersForm()
   DefInstallmentCustomers.Show
End Sub

Public Sub TablesForm()
   DefTables.Show
End Sub

Public Sub ShiftsForm()
   DefShifts.Show
End Sub

Public Sub SaleCountersForm()
   DefSaleCounters.Show
End Sub

Public Sub PackageDealInfoForm()
    DefPackageDealInfo.Show
End Sub

Public Sub SchoolForm()
    DefSchool.Show
End Sub

Public Sub ClassesForm()
    DefClasses.Show
End Sub

Public Sub CoursesForm()
    DefCourses.Show
End Sub

Public Sub SyllabusForm()
    Defsyllabus.Show
End Sub

Public Sub SessionsForm()
    DefSessions.Show
End Sub

Public Sub PurchaseDiscountsForm()
   FrmPurchaseDiscounts.Show
End Sub

Public Sub Bin(objBin As String)
   vBinDataBase = objBin
End Sub
