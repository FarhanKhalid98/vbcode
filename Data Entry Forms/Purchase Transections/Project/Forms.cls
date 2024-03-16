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

Public Sub InitializeChar(c As Object)
    Declarations.SetChar c
End Sub

Public Sub PurchaseInvoiceForm()
    Dim a As New FrmPurchaseInvoice
    a.Show
End Sub

Public Sub PurchaseOrderForm()
    Dim a As New FrmPurchaseOrder
    a.Show
End Sub

Public Sub PurchaseReturnInvoiceForm()
    Dim a As New FrmPurchaseReturnInvoice
    a.Show
End Sub

Public Sub PaymentVenderWiseForm()
    Dim a As New FrmPaymentVenderWise
    a.Show
End Sub

Public Sub PaymentInvoiceWiseForm()
    Dim a As New FrmPaymentInvoiceWise
    a.Show
End Sub

Public Sub PurchasePendingForm()
    Dim a As New SchPurchasePending
    a.Show
End Sub

Public Sub CustomOrderPurchaseForm()
    FrmCustomOrderPurchase.Show
End Sub

Public Sub MultiBarcodesDetailForm()
    FrmMultiBarcodesDetail.Show
End Sub

Public Sub GRNForm()
    Dim a As New FrmGRN
    a.Show
End Sub
Public Sub Bin(objBin As String)
   vBinDataBase = objBin
End Sub

