VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsUserSecurity"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "The Class to manage the Security of Users"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
'Private mvarUserNo As Integer 'local copy
'Private mvarIsAdministrator As Boolean 'local copy
'Private mvarIsLoginSuccess As Boolean
'Private colLoadedObjects As Collection

Public Function TaskAllowance(TaskKey As String) As Boolean
  On Error GoTo ErrorHandler
  If CN.Execute("Select * From UserTasks where taskkey = '" & TaskKey & "' And userno = " & Me.UserNo).RecordCount > 0 Then
    TaskAllowance = True
  Else
    If CN.Execute("Select * From Tasks where taskkey = '" & TaskKey & "'").RecordCount = 0 Then
        MsgBox "Task Key '" & TaskKey & "' was not found. Please contact your Software vendor for Support", vbCritical, "Tasks not available"
    End If
    TaskAllowance = False
  End If
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Public Function LogIn() As Boolean
  On Error GoTo ErrorHandler
  frmLogin.Show vbModal
  If frmLogin.ParaOutUserNo <> 0 Then
    mvarIsLoginSuccess = True
    With CN.Execute("Select userno, UserName, isadministrator, isManager,  isEdit, isEditDefination, IsDelete, IsChangeRetail, ChangePriceSaleInvoice, IsReadOnly, ShowPrice, LastPurchasePrice, WeightedPrice, WSPrice, IsEditClosingInvoice from users where userno = " & frmLogin.ParaOutUserNo)
      Me.UserNo = !UserNo
      Me.UserName = !UserName
      Me.IsAdministrator = !IsAdministrator
      Me.IsManager = !IsManager
      Me.IsEdit = !IsEdit
      Me.IsEditDefination = !IsEditDefination
      Me.IsDelete = !IsDelete
      Me.IsChangeRetail = !IsChangeRetail
      Me.ChangePriceSaleInvoice = !ChangePriceSaleInvoice
      Me.ShowPrice = Abs(!ShowPrice)
      Me.LastPurchasePrice = Abs(!LastPurchasePrice)
      Me.WeightedPrice = Abs(!WeightedPrice)
      Me.WSPrice = Abs(!WSPrice)
      Me.IsEditClosingInvoice = Abs(!IsEditClosingInvoice)
      Me.SessionID = frmLogin.ParaOutSessionID
      If ObjRegistry.ReadOnlyStatus = False Then
         ObjRegistry.ReadOnlyStatus = !IsReadOnly
      End If
    End With
    With CN.Execute("Select OrganizationID, CreditSale, isnull(SaleRePrint,0) SaleRePrint , isnull(PurchaseRePrint,0)  PurchaseRePrint, DisableCreditSale, ShowPurchasePriceInInvoice, ShowSumInSearchSaleInvoice, SalePriceMustBeLessThanPurchase, NotEditingAfterPrinting, ChangePriceFormOpenAsLogin, AllowMaximmDiscPer, ChangeDate, OpenForm, isnull(NoofPrints,0) as NoofPrints, isnull(NoofPurPrints,0) as NoofPurPrints, AllowDiscount, ShowStock  from UserRegistry where userno = " & frmLogin.ParaOutUserNo)
      If .RecordCount > 0 Then
'         Me.OrganizationID = IIf(IsNull(!OrganizationID), "0", !OrganizationID)
         Me.IsCreditSale = !CreditSale
         Me.IsDisableCreditSale = !DisableCreditSale
         Me.ShowPurchasePriceInInvoice = !ShowPurchasePriceInInvoice
         Me.ShowSumInSearchSaleInvoice = !ShowSumInSearchSaleInvoice
         Me.SalePriceMustBeLessThanPurchase = !SalePriceMustBeLessThanPurchase
         Me.ChangePriceFormOpenAsLogin = !ChangePriceFormOpenAsLogin
         Me.NotEditingAfterPrinting = !NotEditingAfterPrinting
         Me.ChangeDate = !ChangeDate
         Me.OpenForm = !OpenForm
         Me.SaleRePrint = Abs(!SaleRePrint)
         Me.PurchaseRePrint = Abs(!PurchaseRePrint)
         Me.AllowMaximmDiscPer = !AllowMaximmDiscPer
         Me.NoofPrints = !NoofPrints
         Me.NoofPurPrints = !NoofPurPrints
         Me.AllowDiscount = !AllowDiscount
         Me.ShowStock = !ShowStock
      Else
         Me.IsCreditSale = False
         Me.IsDisableCreditSale = False
      End If
    End With
  Else
    mvarIsLoginSuccess = False
  End If
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Public Function LogInOK() As Boolean
  On Error GoTo ErrorHandler
    mvarIsLoginSuccess = True
    With CN.Execute("Select userno, username, isadministrator, isManager, isEdit, isEditDefination, IsDelete, IsChangeRetail,ShowPrice, isCreditSale from users where userno = " & UserNo)
      Me.UserNo = !UserNo
      Me.UserName = !UserName
      Me.IsAdministrator = !IsAdministrator
      Me.IsManager = !IsManager
      Me.IsEdit = !IsEdit
      Me.IsEditDefination = !IsEditDefination
      Me.IsDelete = !IsDelete
      Me.IsChangeRetail = !IsChangeRetail
      Me.ShowPrice = !ShowPrice
      Me.IsCreditSale = !IsCreditSale
      Me.SessionID = frmLogin.CmbSession.ItemData(frmLogin.CmbSession.ListIndex)
    End With
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

'Public Function Refresh() As Boolean
'   On Error GoTo ErrorHandler
'   With CN.Execute("Select userno, isadministrator, isEdit, IsDelete from users where userno = " & frmLogin.ParaOutUserNo)
'      Me.UserNo = !UserNo
'      Me.UserName = !UserName
'      Me.IsAdministrator = !IsAdministrator
'      Me.IsEdit = !IsEdit
'      Me.IsDelete = !IsDelete
'    End With
'    Exit Function
'ErrorHandler:
'  Call ShowErrorMessage
'End Function

Public Function LogOut() As Boolean
Attribute LogOut.VB_Description = "Breaks the session of the current user with the Application."
      Me.UserNo = 0
      Me.UserName = ""
      Me.IsAdministrator = False
      Me.IsManager = False
      Me.IsEdit = False
      Me.IsEditDefination = False
      Me.IsDelete = False
      Me.IsChangeRetail = False
      LogIn
End Function

Public Property Get IsLoginSuccess() As Boolean
    IsLoginSuccess = mvarIsLoginSuccess
End Property

Public Property Let IsAdministrator(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsAdministrator = 5
    mvarIsAdministrator = vData
End Property

Public Property Get IsAdministrator() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsAdministrator
    IsAdministrator = mvarIsAdministrator
End Property

Public Property Let IsCreditSale(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsAdministrator = 5
    mvarIsCreditSale = vData
End Property

Public Property Get IsCreditSale() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsAdministrator
    IsCreditSale = mvarIsCreditSale
End Property

Public Property Let IsDisableCreditSale(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsAdministrator = 5
    mvarIsDisableCreditSale = vData
End Property

Public Property Get IsDisableCreditSale() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsAdministrator
    IsDisableCreditSale = mvarIsDisableCreditSale
End Property
Public Property Let IsManager(ByVal vData As Boolean)
    mvarIsManager = vData
End Property

Public Property Get IsManager() As Boolean
    IsManager = mvarIsManager
End Property
Public Property Let OrganizationID(ByVal vData As Integer)
    mvarOrganizationID = vData
End Property

Public Property Get OrganizationID() As Integer
    OrganizationID = mvarOrganizationID
End Property

Public Property Let ShowSumInSearchSaleInvoice(ByVal vData As Boolean)
    mvarShowSumInSearchSaleInvoice = vData
End Property

Public Property Get ShowSumInSearchSaleInvoice() As Boolean
    ShowSumInSearchSaleInvoice = mvarShowSumInSearchSaleInvoice
End Property

Public Property Let ShowPurchasePriceInInvoice(ByVal vData As Boolean)
    mvarShowPurchasePriceInInvoice = vData
End Property

Public Property Get ShowPurchasePriceInInvoice() As Boolean
    ShowPurchasePriceInInvoice = mvarShowPurchasePriceInInvoice
End Property

Public Property Let SalePriceMustBeLessThanPurchase(ByVal vData As Boolean)
    mvarSalePriceMustBeLessThanPurchase = vData
End Property

Public Property Get SalePriceMustBeLessThanPurchase() As Boolean
    SalePriceMustBeLessThanPurchase = mvarSalePriceMustBeLessThanPurchase
End Property

Public Property Let ChangePriceFormOpenAsLogin(ByVal vData As Boolean)
    mvarChangePriceFormOpenAsLogin = vData
End Property

Public Property Get ChangePriceFormOpenAsLogin() As Boolean
    ChangePriceFormOpenAsLogin = mvarChangePriceFormOpenAsLogin
End Property


Public Property Let NotEditingAfterPrinting(ByVal vData As Boolean)
    mvarNotEditingAfterPrinting = vData
End Property

Public Property Get NotEditingAfterPrinting() As Boolean
    NotEditingAfterPrinting = mvarNotEditingAfterPrinting
End Property
Public Property Let ChangeDate(ByVal vData As Boolean)
    mvarChangeDate = vData
End Property

Public Property Get ChangeDate() As Boolean
    ChangeDate = mvarChangeDate
End Property

Public Property Let OpenForm(ByVal vData As Boolean)
    mvarOpenForm = vData
End Property

Public Property Get OpenForm() As Boolean
    OpenForm = mvarOpenForm
End Property

Public Property Let AllowDiscount(ByVal vData As Boolean)
    mvarAllowDiscount = vData
End Property

Public Property Get AllowDiscount() As Boolean
    AllowDiscount = mvarAllowDiscount
End Property

Public Property Let ShowStock(ByVal vData As Boolean)
    mvarShowStock = vData
End Property

Public Property Get ShowStock() As Boolean
    ShowStock = mvarShowStock
End Property

Public Property Let IsEdit(ByVal vData As Boolean)
    mvarIsEdit = vData
End Property

Public Property Get IsEdit() As Boolean
    IsEdit = mvarIsEdit
End Property

Public Property Let IsEditDefination(ByVal vData As Boolean)
    mvarIsEditDefination = vData
End Property

Public Property Get IsEditDefination() As Boolean
    IsEditDefination = mvarIsEditDefination
End Property

Public Property Let IsDelete(ByVal vData As Boolean)
    mvarIsDelete = vData
End Property

Public Property Get IsDelete() As Boolean
    IsDelete = mvarIsDelete
End Property

Public Property Let IsChangeRetail(ByVal vData As Boolean)
   mvarIsChangeRetail = vData
End Property

Public Property Get IsChangeRetail() As Boolean
   IsChangeRetail = mvarIsChangeRetail
End Property

Public Property Let ChangePriceSaleInvoice(ByVal vData As Boolean)
   mvarChangePriceSaleInvoice = vData
End Property

Public Property Get ShowPrice() As Boolean
   ShowPrice = mvarShowPrice
End Property
Public Property Let ShowPrice(ByVal vData As Boolean)
   mvarShowPrice = vData
End Property

Public Property Get WSPrice() As Boolean
   WSPrice = mvarWSPrice
End Property
Public Property Let WSPrice(ByVal vData As Boolean)
   mvarWSPrice = vData
End Property

Public Property Get IsEditClosingInvoice() As Boolean
   IsEditClosingInvoice = mvarIsEditClosingInvoice
End Property
Public Property Let IsEditClosingInvoice(ByVal vData As Boolean)
   mvarIsEditClosingInvoice = vData
End Property

Public Property Get SaleRePrint() As Boolean
   SaleRePrint = mvarSaleRePrint
End Property
Public Property Let SaleRePrint(ByVal vData As Boolean)
   mvarSaleRePrint = vData
End Property

Public Property Get PurchaseRePrint() As Boolean
   PurchaseRePrint = mvarPurchaseRePrint
End Property
Public Property Let PurchaseRePrint(ByVal vData As Boolean)
   mvarPurchaseRePrint = vData
End Property

Public Property Get WeightedPrice() As Boolean
   WeightedPrice = mvarWeightedPrice
End Property
Public Property Let WeightedPrice(ByVal vData As Boolean)
   mvarWeightedPrice = vData
End Property

Public Property Get LastPurchasePrice() As Boolean
   LastPurchasePrice = mvarLastPurchasePrice
End Property

Public Property Let LastPurchasePrice(ByVal vData As Boolean)
   mvarLastPurchasePrice = vData
End Property

Public Property Get ChangePriceSaleInvoice() As Boolean
   ChangePriceSaleInvoice = mvarChangePriceSaleInvoice
End Property


Public Sub Initialize(Con As String, vPath As String)
    ParaPass = Con
    vTmp = vPath
End Sub

Public Sub InitializeClass(AdodbConnection As ADODB.Connection)
   If ParaPass <> EncryptStr("›· ‹ÔÙÌÓ€ﬂÿ‡", False) Then Exit Sub
   SetConnection AdodbConnection
End Sub

Public Property Let UserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserNo = 5
    mvarUserName = vData
End Property

Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserNo
    UserName = mvarUserName
End Property

Public Property Let UserNo(ByVal vData As Integer)
Attribute UserNo.VB_Description = "The ID of the currently logged in User"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserNo = 5
    mvarUserNo = vData
End Property

Public Property Get UserNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserNo
    UserNo = mvarUserNo
End Property

Public Property Let NoofPurPrints(ByVal vData As Integer)
    mvarNoofPurPrints = vData
End Property

Public Property Get NoofPurPrints() As Integer
    NoofPurPrints = mvarNoofPurPrints
End Property

Public Property Let NoofPrints(ByVal vData As Integer)
    mvarNoofPrints = vData
End Property

Public Property Get NoofPrints() As Integer
    NoofPrints = mvarNoofPrints
End Property


Public Property Let SessionID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserNo = 5
    mvarSessionID = vData
End Property

Public Property Get SessionID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserNo
    SessionID = mvarSessionID
End Property

Public Property Let AllowMaximmDiscPer(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserNo = 5
    mvarAllowMaximmDiscPer = vData
End Property

Public Property Get AllowMaximmDiscPer() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserNo
    AllowMaximmDiscPer = mvarAllowMaximmDiscPer
End Property

Public Sub TaskAssignmentForm()
   DefAssignTasks.ParaInUserNo = Me.UserNo
   DefAssignTasks.Show vbModal
End Sub

Public Sub UsersForm()
   DefUsers.ParaInUserNo = Me.UserNo
   DefUsers.Show vbModal
End Sub

Public Sub ActivityLogForm()
   Dim a As New FrmActivityLogBin
   a.Show
End Sub
Public Sub OldActivityLogForm()
   Dim a As New FrmActivityLog
   a.Show
End Sub

Public Sub UserDefaultSettingsForm()
   FrmUserDefaultSettings.Show vbModal
End Sub

Public Sub ChangePasswordForm()
  DefChangePassword.ParaInUserNo = Me.UserNo
  DefChangePassword.Show vbModal
End Sub

Public Sub Bin(objBin As String)
   vBinDataBase = objBin
End Sub

