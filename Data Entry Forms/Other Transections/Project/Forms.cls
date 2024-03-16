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
Public Sub Bin(objBin As String)
   vBinDataBase = objBin
End Sub

Public Sub InitializeSecurity(objSec As UserSecurity.ClsUserSecurity)
    Declarations.SetSecurityReference objSec
End Sub

Public Sub InitializeChar(c As Object)
    Declarations.SetChar c
End Sub

Public Sub EmployeeAttendanceInForm()
    FrmEmployeeAttendanceIn.Show
End Sub

Public Sub EmployeeAttendanceOutForm()
    FrmEmployeeAttendanceOut.Show
End Sub

Public Sub EmployeeLeaveForm()
    FrmEmployeeLeave.Show
End Sub

Public Sub HolidayForm()
    FrmHoliday.Show
End Sub

Public Sub SalaryForm()
   FrmSalary.Show
End Sub

Public Sub PettyCashForm()
    Dim a As New FrmPettyCash
    a.Show
End Sub

Public Sub UserClosingForm()
    Dim a As New FrmUserClosing
    a.Show
End Sub

Public Sub AdminClosingForm()
    Dim a As New FrmAdminClosing
    a.Show
End Sub

Public Sub OpeningProductsForm()
    Dim a As New FrmOpeningProducts
    a.Show
End Sub

Public Sub MeterReadingsForm()
   FrmMeterReadings.Show
End Sub

Public Sub PettyCashVerificationForm()
   FrmPettyCashVerification.Show vbModal
End Sub

Public Sub OpeningProductsVerificationForm()
   FrmOpeningProductsVerification.Show vbModal
End Sub

Public Sub CustomerDemandForm()
   FrmCustomerDemand.Show
End Sub

Public Sub MonthlyIncomeExpenseForm()
   FrmMonthlyIncomeExpense.Show
End Sub

Public Sub EmployeeAttendanceForm()
   FrmEmployeeAttendance.Show
End Sub

Public Sub BanquetOrderForm()
   FrmBanquetOrder.Show
End Sub

Public Sub BanquetInvoiceForm()
   FrmBanquetInvoice.Show
End Sub

