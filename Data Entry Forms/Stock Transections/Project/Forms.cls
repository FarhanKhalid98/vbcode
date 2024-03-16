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

Public Sub ExpiryDamageClaimInvoiceForm()
    Dim a As New FrmExpiryDamageClaimInvoice
    a.Show
End Sub

Public Sub ExpiryDamageInvoiceForm()
    Dim a As New FrmExpiryDamageInvoice
    a.Show
End Sub

Public Sub StockTransferForm()
    Dim a As New FrmStockTransfer
    a.Show
End Sub

Public Sub LiftInvoiceForm()
    Dim a As New FrmLiftInvoice
    a.Show
End Sub

Public Sub DisputeInvoiceForm()
    Dim a As New FrmDisputeInvoice
    a.Show
End Sub

Public Sub StockWastageInvoiceForm()
    Dim a As New FrmStockWastageInvoice
    a.Show
End Sub

Public Sub StockPendingTransferForm()
    Dim a As New SchStockPendingTransfer
    a.Show
End Sub

Public Sub PriceVariationForm()
    Dim a As New FrmPriceVariation
    a.Show
End Sub

Public Sub Bin(objBin As String)
   vBinDataBase = objBin
End Sub
