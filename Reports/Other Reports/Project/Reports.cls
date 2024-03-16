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

Public Sub AdminClosingReport()
   RptAdminClosing.Show
End Sub

Public Sub SalaryEmployeeWiseReport()
    RptSalaryEmpWise.Show
End Sub

Public Sub EmployeeAttendanceReport()
    RptEmpAttendance.Show
End Sub

Public Sub HotAndColdMembersReport()
    RptHotAndColdMembers.Show
End Sub

Public Sub SalesActivityReport()
'    RptSalesActivity.Show
End Sub

Public Sub SleepingMembersReport()
    RptSleepingMembers.Show
End Sub

Public Sub CashActivityReport()
    RptCashActivity.Show
End Sub

Public Sub CustomerDemandRegisterReport()
    RptCustomerDemandRegister.Show
End Sub

Public Sub MeterReadingRegisterReport()
    RptMeterReadingRegister.Show
End Sub

Public Sub MembersLetterReport()
    RptMembersLetter.Show
End Sub

Public Sub EnvelopReport()
    RptEnvelop.Show
End Sub

Public Sub ProductDifferenceReport()
    RptProductDifference.Show
End Sub

