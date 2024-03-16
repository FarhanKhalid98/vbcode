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

Public Sub PurchaseRegisterReport()
   RptPurchaseRegister.Show
End Sub

Public Sub PurchaseOrderRegisterReport()
   RptPurchaseOrderRegister.Show
End Sub

Public Sub PurchaseSerialWiseReport()
   RptPurchaseSerialWise.Show
End Sub

Public Sub PendingBillsToVendorReport()
'    RptPendingBillsToVendor.Show
End Sub

Public Sub ProductWisePurchaseDetailReport()
    RptProductWisePurchaseDetail.Show
End Sub

Public Sub VenderPurchaseBillsReport()
   RptVenderPurchaseBills.Show
End Sub

Public Sub ProductNotInPurchaseReport()
   RptProductNotInPurchase.Show
End Sub
