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

Public Sub Initialize(Con As String)
    ParaPass = Con
End Sub

Public Sub InitializeClass(objCN As ADODB.Connection)
   If ParaPass <> EncryptStr("›· ‹ÔÙÌÓ€ﬂÿ‡", False) Then Exit Sub
   Declarations.SetConnection objCN
End Sub

Public Sub InitializeSecurity(objSec As UserSecurity.ClsUserSecurity)
    Declarations.SetSecurityReference objSec
End Sub

Public Sub StockIssueToSalemanForm()
    Dim a As New FrmStockIssueToSaleman
    a.Show
End Sub

Public Sub StockReturnFromSalemanForm()
    Dim a As New FrmStockReturnFromSaleman
    a.Show
End Sub

Public Sub RecoveryInvoiceWiseForm()
'    Dim a As New FrmRecoveryInvoiceWise
'    a.Show
End Sub

Public Sub RecoveryCustomerWiseForm()
'    Dim a As New FrmRecoveryCustomerWise
'    a.Show
End Sub

