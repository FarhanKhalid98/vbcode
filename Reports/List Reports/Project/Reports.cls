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

Public Sub CompanyListReport()
    RptCompanyList.Show
End Sub

Public Sub CustomerListReport()
    RptCustomerList.Show
End Sub

Public Sub GroupListReport()
    RptGroupList.Show
End Sub

Public Sub GroupWiseProductPackInfoListReport()
    RptGroupWiseProductPackInfoList.Show
End Sub

Public Sub OpeningStockListReport()
    RptOpeningStockList.Show
End Sub

Public Sub ProductListReport()
    RptProductList.Show
End Sub

Public Sub DeadProductListReport()
    RptDeadProductList.Show
End Sub

Public Sub MemberListReport()
    RptMemberList.Show
End Sub

Public Sub ProductPriceListReport()
    RptProductPriceList.Show
End Sub

Public Sub StoreListReport()
    RptStoreList.Show
End Sub

Public Sub SubGroupListReport()
    RptSubGroupList.Show
End Sub

Public Sub VendorListReport()
    RptVendorList.Show
End Sub

Public Sub EmployeeListReport()
    RptEmployeeList.Show
End Sub

Public Sub ProductNotIncludedListReport()
    RptProductNotIncludedList.Show
End Sub

