Attribute VB_Name = "MdlDeclaration"
Option Explicit
Public ObjRegistry As New SoftinnRegistry.Registry
Public CN As New ADODB.Connection
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public mvarUserNo As Integer 'local copy
Public mvarUserName As String  'local copy
Public mvarIsAdministrator As Boolean 'local copy
Public mvarIsManager As Boolean 'local copy
Public mvarIsEdit As Boolean 'local copy
Public mvarIsDelete As Boolean 'local copy
Public mvarIsChangeRetail As Boolean 'local copy
Public mvarChangePriceSaleInvoice As Boolean 'local copy
Public mvarIsLoginSuccess As Boolean
Public mvarIsCreditSale As Boolean
Public mvarIsDisableCreditSale As Boolean
Public mvarShowPurchasePriceInInvoice As Boolean
Public mvarShowSumInSearchSaleInvoice As Boolean
Public mvarSalePriceMustBeLessThanPurchase As Boolean
Public mvarNotEditingAfterPrinting As Boolean
Public mvarChangePriceFormOpenAsLogin As Boolean
Public mvarShowPrice As Boolean
Public ParaPass As String
Public mvarOrganizationID As Integer
Public mvarSessionID As Integer
Public mvarIsEditDefination As Boolean
Public mvarChangeDate As Boolean
Public mvarOpenForm As Boolean
Public mvarLastPurchasePrice As Boolean
Public mvarWeightedPrice As Boolean
Public mvarWSPrice As Boolean
Public mvarIsEditClosingInvoice As Boolean
Public mvarAllowMaximmDiscPer As Double
Public mvarSaleRePrint As Double
Public mvarPurchaseRePrint As Double
Public mvarNoofPurPrints As Integer
Public mvarNoofPrints As Integer
Public mvarAllowDiscount As Boolean
Public mvarShowStock As Boolean

Public Sub SetConnection(ConnObject As ADODB.Connection)
  Set CN = ConnObject
End Sub
