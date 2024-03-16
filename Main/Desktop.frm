VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Desktop 
   AutoRedraw      =   -1  'True
   ClientHeight    =   10650
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "Desktop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Desktop.frx":0ECA
   ScaleHeight     =   13484.28
   ScaleMode       =   0  'User
   ScaleWidth      =   34547.36
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTempTables 
      Height          =   1185
      Left            =   7920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "Desktop.frx":1A48D
      Top             =   5895
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox TxtProcList 
      Height          =   1185
      Left            =   7860
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "Desktop.frx":1AB90
      Top             =   4410
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6435
      Top             =   1530
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   9420
      Top             =   2010
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8280
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LblAutoBackup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4635
      TabIndex        =   7
      Top             =   10305
      UseMnemonic     =   0   'False
      Width           =   60
   End
   Begin VB.Image ImgDesktopButton 
      Height          =   705
      Index           =   0
      Left            =   1620
      Tag             =   "MniSaleInvoice"
      Top             =   2025
      Width           =   1545
   End
   Begin VB.Label LblCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   0
      Left            =   1740
      TabIndex        =   5
      Top             =   2025
      Width           =   885
      WordWrap        =   -1  'True
   End
   Begin VB.Shape ImgBorder 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   705
      Left            =   1050
      Top             =   2550
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10920
      TabIndex        =   3
      Top             =   9630
      Width           =   1035
   End
   Begin VB.Label LblUser 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8670
      TabIndex        =   2
      Top             =   9630
      Width           =   585
   End
   Begin VB.Label LblLoginTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8670
      TabIndex        =   1
      Top             =   9870
      Width           =   585
   End
   Begin VB.Label LblCompany 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10950
      TabIndex        =   0
      Top             =   9870
      UseMnemonic     =   0   'False
      Width           =   585
   End
   Begin VB.Image ImgStart 
      Height          =   405
      Left            =   0
      Top             =   8580
      Width           =   1050
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11640
      Top             =   90
      Width           =   315
   End
   Begin VB.Image ImgMin 
      Height          =   315
      Left            =   11295
      Top             =   105
      Width           =   315
   End
   Begin VB.Menu MnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnu100 
         Caption         =   ""
      End
   End
   Begin VB.Menu MnuDefinition 
      Caption         =   "Definition"
      Begin VB.Menu MniOrganization 
         Caption         =   "Organization"
      End
      Begin VB.Menu MniBrands 
         Caption         =   "Brands"
      End
      Begin VB.Menu MniCompany 
         Caption         =   "Company"
      End
      Begin VB.Menu MniGroups 
         Caption         =   "Groups"
      End
      Begin VB.Menu MniSubGroup 
         Caption         =   "Sub Group"
      End
      Begin VB.Menu MniRacks 
         Caption         =   "Racks"
      End
      Begin VB.Menu MniProducts 
         Caption         =   "Products"
      End
      Begin VB.Menu MniProductsDetail 
         Caption         =   "Products Detail"
      End
      Begin VB.Menu MniProductOffer 
         Caption         =   "Product Offer"
      End
      Begin VB.Menu MniChangePrice 
         Caption         =   "Change Price"
      End
      Begin VB.Menu MniChangeCategory 
         Caption         =   "Change Category"
      End
      Begin VB.Menu MniFormulaInfo 
         Caption         =   "Formula Info"
      End
      Begin VB.Menu MniPackageDealInfo 
         Caption         =   "Package Deal Info"
      End
      Begin VB.Menu MniCustomProductsAndMeasurements 
         Caption         =   "Custom Products and Measurements"
      End
      Begin VB.Menu MniStockLimitStoreWise 
         Caption         =   "Stock Limit Store Wise"
      End
      Begin VB.Menu MniDefineStockLimit 
         Caption         =   "Define Stock Limit"
      End
      Begin VB.Menu MniBankMachines 
         Caption         =   "Bank Machines"
      End
      Begin VB.Menu MniSaleCounters 
         Caption         =   "Sale Counters"
      End
      Begin VB.Menu MniTables 
         Caption         =   "Tables"
      End
      Begin VB.Menu MniPacking 
         Caption         =   "Packing"
      End
      Begin VB.Menu MniStores 
         Caption         =   "Stores"
      End
      Begin VB.Menu MniZones 
         Caption         =   "Zones"
      End
      Begin VB.Menu MniSectors 
         Caption         =   "Sectors"
      End
      Begin VB.Menu MniVendors 
         Caption         =   "Venders"
      End
      Begin VB.Menu MniCustomerTypes 
         Caption         =   "Customer Types"
      End
      Begin VB.Menu MniCustomers 
         Caption         =   "Customers"
      End
      Begin VB.Menu MniInstallmentCustomers 
         Caption         =   "Installment Customers"
      End
      Begin VB.Menu MniReferences 
         Caption         =   "References"
      End
      Begin VB.Menu MniMembers 
         Caption         =   "Members"
      End
      Begin VB.Menu MniMemberTypes 
         Caption         =   "Member Types"
      End
      Begin VB.Menu MniMembersDiscount 
         Caption         =   "Members Discount"
      End
      Begin VB.Menu MniCommisionDisc 
         Caption         =   "Commision Disc"
      End
      Begin VB.Menu MniExpiryDayColor 
         Caption         =   "Expiry Day Color"
      End
      Begin VB.Menu MniSchool 
         Caption         =   "School"
      End
      Begin VB.Menu MniClasses 
         Caption         =   "Classes"
      End
      Begin VB.Menu MniCourses 
         Caption         =   "Courses"
      End
      Begin VB.Menu MniSyllabus 
         Caption         =   "Syllabus"
      End
      Begin VB.Menu MniDepartments 
         Caption         =   "Departments"
      End
      Begin VB.Menu MniSessions 
         Caption         =   "Sessions"
      End
      Begin VB.Menu mnu34 
         Caption         =   "-"
      End
      Begin VB.Menu MniShifts 
         Caption         =   "Shifts"
      End
      Begin VB.Menu MniEmpDepartments 
         Caption         =   "Employee Departments"
      End
      Begin VB.Menu MniDesignations 
         Caption         =   "Designations"
      End
      Begin VB.Menu MniEmployees 
         Caption         =   "Employees"
      End
      Begin VB.Menu mnu33 
         Caption         =   "-"
      End
      Begin VB.Menu MniChartofAccounts 
         Caption         =   "Chart Of Accounts"
      End
      Begin VB.Menu MniLockAccounts 
         Caption         =   "Lock Accounts"
      End
      Begin VB.Menu MnuOpenings 
         Caption         =   "Openings"
         Begin VB.Menu MniOpeningStock 
            Caption         =   "Opening Stock"
         End
         Begin VB.Menu MniStockAdjustment 
            Caption         =   "Stock Adjustment"
         End
         Begin VB.Menu MniOpeningAccounts 
            Caption         =   "Opening Accounts"
         End
         Begin VB.Menu MniOpeningAccountsOrganization 
            Caption         =   "Opening Accounts Organization"
         End
         Begin VB.Menu MniExpSetting 
            Caption         =   "Exp Setting"
         End
         Begin VB.Menu MniPLSettingWithoutOrg 
            Caption         =   "P/L Setting Without Org"
         End
         Begin VB.Menu MniPLSetting 
            Caption         =   "P/L Setting"
         End
         Begin VB.Menu mnu22 
            Caption         =   "-"
         End
      End
   End
   Begin VB.Menu MnuInvoicing 
      Caption         =   "Invoicing"
      Begin VB.Menu MniPurchaseOrder 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu MniPurchaseInvoice 
         Caption         =   "Purchase Invoice"
      End
      Begin VB.Menu MniPurchaseReturnInvoice 
         Caption         =   "Purchase Return Invoice"
      End
      Begin VB.Menu MniGRN 
         Caption         =   "Goods Received Notes"
      End
      Begin VB.Menu MniPriceVariation 
         Caption         =   "Price Variation"
      End
      Begin VB.Menu mnu16 
         Caption         =   "-"
      End
      Begin VB.Menu MniCustomOrderPurchase 
         Caption         =   "Custom Order Purchase"
      End
      Begin VB.Menu MniCustomOrderBooking 
         Caption         =   "Custom Order Booking"
      End
      Begin VB.Menu MniCustomOrderDelivery 
         Caption         =   "Custom Order Delivery"
      End
      Begin VB.Menu MniCustomOrderReturn 
         Caption         =   "Custom Order Return"
      End
      Begin VB.Menu mnu25 
         Caption         =   "-"
      End
      Begin VB.Menu MniSaleOrderPOS 
         Caption         =   "Sale Order (POS)"
      End
      Begin VB.Menu MniSaleOrder 
         Caption         =   "Sale Order"
      End
      Begin VB.Menu MniSaleInvoiceTouch 
         Caption         =   "Sale Invoice Touch"
      End
      Begin VB.Menu MniSaleInvoicePOS 
         Caption         =   "Sale Invoice (POS)"
      End
      Begin VB.Menu MniSaleInvoice 
         Caption         =   "Sale Invoice"
         HelpContextID   =   1
      End
      Begin VB.Menu MniReceivingInvoicePrint 
         Caption         =   "Receiving Invoice Print"
      End
      Begin VB.Menu MniSaleReturnInvoicePOS 
         Caption         =   "Sale Return Invoice (POS)"
      End
      Begin VB.Menu MniSaleReturnInvoice 
         Caption         =   "Sale Return Invoice"
      End
      Begin VB.Menu MniReplacementInvoice 
         Caption         =   "Replacement Invoice"
      End
      Begin VB.Menu MniPostSale 
         Caption         =   "Post Sale"
      End
      Begin VB.Menu mnu17 
         Caption         =   "-"
      End
      Begin VB.Menu MniManufacturedProducts 
         Caption         =   "Manufactured Products"
      End
      Begin VB.Menu MniManufacturedReturn 
         Caption         =   "Manufactured Return"
      End
      Begin VB.Menu MniProductionIN 
         Caption         =   "Production IN"
      End
      Begin VB.Menu MniProductionOut 
         Caption         =   "Production OUT"
      End
      Begin VB.Menu mnu18 
         Caption         =   "-"
      End
      Begin VB.Menu MniStockTransferInvoice 
         Caption         =   "Stock Transfer Invoice"
      End
      Begin VB.Menu MniStockWastageInvoice 
         Caption         =   "Stock Wastage Invoice"
      End
      Begin VB.Menu MniExpiryDamageInvoice 
         Caption         =   "Expiry/Damage Invoice"
      End
      Begin VB.Menu MniExpiryDamageClaimInvoice 
         Caption         =   "Expiry/Damage Claim Invoice"
      End
      Begin VB.Menu MniLiftInvoice 
         Caption         =   "Lift Invoice"
      End
      Begin VB.Menu MniDisputeInvoice 
         Caption         =   "Dispute Invoice"
      End
   End
   Begin VB.Menu MnuTransection 
      Caption         =   "Transaction"
      Begin VB.Menu MniCashPaymentVoucher 
         Caption         =   "Cash Payment Voucher"
      End
      Begin VB.Menu MniCashReceiveVoucher 
         Caption         =   "Cash Received Voucher"
      End
      Begin VB.Menu MniJournalVoucher 
         Caption         =   "Journal Voucher"
      End
      Begin VB.Menu MniAdvances 
         Caption         =   "Advances"
      End
      Begin VB.Menu MniLoans 
         Caption         =   "Loans"
      End
      Begin VB.Menu mnu24 
         Caption         =   "-"
      End
      Begin VB.Menu MniRecoveryCustomer 
         Caption         =   "Recovery Customer"
      End
      Begin VB.Menu MniRecoveryInvoiceWise 
         Caption         =   "Recovery Invoice Wise"
      End
      Begin VB.Menu MniPaymentVender 
         Caption         =   "Payment Vender"
      End
      Begin VB.Menu MniPaymentInvoiceWise 
         Caption         =   "Payment Invoice Wise"
      End
      Begin VB.Menu mnu26 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuBank 
      Caption         =   "Bank"
      Begin VB.Menu MniChequeIssuance 
         Caption         =   "Cheque Issuance"
      End
      Begin VB.Menu MniChequeDeposit 
         Caption         =   "Cheque Deposit"
      End
      Begin VB.Menu MniChequeReceive 
         Caption         =   "Cheque Receive"
      End
      Begin VB.Menu MniCashDeposite 
         Caption         =   "Cash Deposite"
      End
      Begin VB.Menu MniChequeIssuanceReconcilation 
         Caption         =   "Cheque Issuance Reconcilation"
      End
      Begin VB.Menu MniChequeReceiveReconcilation 
         Caption         =   "Cheque Receive Reconcilation"
      End
      Begin VB.Menu mnu14 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuOthers 
      Caption         =   "Others"
      Begin VB.Menu MniPettyCash 
         Caption         =   "Petty Cash"
      End
      Begin VB.Menu MnuUserClosing 
         Caption         =   "User Closing"
      End
      Begin VB.Menu MniAdminClosing 
         Caption         =   "Admin Closing"
      End
      Begin VB.Menu MniMeterReadings 
         Caption         =   "Meter Readings"
      End
      Begin VB.Menu MniBanquetOrder 
         Caption         =   "Banquet Order"
      End
      Begin VB.Menu MniBanquetInvoice 
         Caption         =   "Banquet Invoice"
      End
      Begin VB.Menu MniGatePassIn 
         Caption         =   "Gate Pass In"
      End
      Begin VB.Menu MniGatePassOut 
         Caption         =   "Gate Pass Out"
      End
      Begin VB.Menu MniOpeningProduct 
         Caption         =   "Opening Product"
      End
      Begin VB.Menu MniEmployeeAttendanceAll 
         Caption         =   "Employee Attendance All"
      End
      Begin VB.Menu MniEmployeeAttendanceIn 
         Caption         =   "Employee Attendance In"
      End
      Begin VB.Menu MniEmployeeAttendanceOut 
         Caption         =   "Employee Attendance Out"
      End
      Begin VB.Menu MniEmployeeLeave 
         Caption         =   "Employee Leave"
      End
      Begin VB.Menu MniHoliday 
         Caption         =   "Holiday"
      End
      Begin VB.Menu MniSalary 
         Caption         =   "Salary"
      End
      Begin VB.Menu MniCustomerDemand 
         Caption         =   "Customer Demand"
      End
      Begin VB.Menu MniMonthlyIncomeExpense 
         Caption         =   "Monthly Income Expense"
      End
      Begin VB.Menu mnu30 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuUtility 
      Caption         =   "Utility"
      Tag             =   "0MnuMain"
      Begin VB.Menu MniPriceChecker 
         Caption         =   "Price Checker"
      End
      Begin VB.Menu MniSyncData 
         Caption         =   "Sync Data"
      End
      Begin VB.Menu MnuSMS 
         Caption         =   "SMS"
         Tag             =   "1MnuUtility"
      End
      Begin VB.Menu MniAutoBackup 
         Caption         =   "Auto Backup"
         Tag             =   "1MnuUtility"
      End
      Begin VB.Menu MniBackup 
         Caption         =   "Backup  Data"
         Tag             =   "1MnuUtility"
      End
      Begin VB.Menu MniRestore 
         Caption         =   "Restore  Data"
         Tag             =   "1MnuUtility"
      End
      Begin VB.Menu MnuDesktopSetting 
         Caption         =   "Desktop Setting"
         Tag             =   "1MnuUtility"
      End
      Begin VB.Menu MniSoftwareDefaultSettings 
         Caption         =   "Software Default Settings"
      End
      Begin VB.Menu MniAccountsDefaultSetting 
         Caption         =   "Accounts Default Setting"
      End
      Begin VB.Menu MniUserDefaultSettings 
         Caption         =   "User Default Settings"
      End
      Begin VB.Menu MniCompanyInformation 
         Caption         =   "Company Information"
      End
      Begin VB.Menu MniSkinSelection 
         Caption         =   "Skin Selection"
      End
      Begin VB.Menu MniAddSkin 
         Caption         =   "Add Skin"
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuUserSecurity 
      Caption         =   "User Security"
      Tag             =   "0MnuMain"
      Begin VB.Menu MniAssignTasks 
         Caption         =   "Assign Tasks"
         Tag             =   "1MnuUserSecurity"
      End
      Begin VB.Menu MnuChangePassword 
         Caption         =   "Change Password"
         Tag             =   "1MnuUserSecurity"
      End
      Begin VB.Menu MniUsers 
         Caption         =   "Users"
         Tag             =   "1MnuUserSecurity"
      End
      Begin VB.Menu MniActivityLog 
         Caption         =   "Activity Log"
      End
      Begin VB.Menu MniOldActivityLog 
         Caption         =   "Activity Log Old"
      End
      Begin VB.Menu Line3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuReports 
      Caption         =   "Reports"
   End
   Begin VB.Menu MnuLogout 
      Caption         =   "Logout"
      Tag             =   "0MnuMain"
   End
End
Attribute VB_Name = "Desktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Date
Dim Ind As Integer
Dim Flag As Boolean
Dim vMobileNo() As String, vMobile As String, sSql, vLessMargin As String

Private Sub Form_GotFocus()
   If ObjUserSecurity.ChangePriceFormOpenAsLogin Then
      If ObjRegistry.ShowWholeSaleMargin = True Then
         vLessMargin = "round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0)/isnull(sp.multiplier,1) - (PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1))*100/ isnull(WSPrice,1) end,3)"
      Else
         vLessMargin = "round(case when RetailPrice = 0 then 0 else (isnull(RetailPrice,0) - ((PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1)))*100/ isnull(RetailPrice,1) end,3)"
      End If
     
     sSql = " SELECT Top 1 p.ProductID, isnull(PackingName,'') as PackingName, isnull(pp.Multiplier,0) as Multiplier, isnull(SP.Multiplier,0) as SaleMultiplier, PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0) as PurchasePrice, isnull(ListPrice,0) as ListPrice, " & vLessMargin & " Margin " & " from Products p" & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 And " & vLessMargin & " <= 0"
      
      If Not CN.Execute(sSql).EOF Then
         If MniChangePrice.Visible Then ObjDefinition.ChangePriceForm
      End If
   End If
End Sub

'Private Sub Form_Click()
''  Call Triggers(Flag)
''  MsgBox Flag
''  Flag = Not Flag
''  Form1.Show
'End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
    ImgBorder.Visible = False
    Ind = 13
End Sub

Private Sub ImgDesktopButton_Click(Index As Integer)
   Dim mniDesktop As Boolean
   mniDesktop = False
   Dim mni As Control
      For Each mni In Me.Controls
         If TypeOf mni Is Menu Then
            If mni.Name = ImgDesktopButton(Index).Tag Then
               If Desktop.Controls(ImgDesktopButton(Index).Tag).Enabled Then
                  If ImgDesktopButton(Index).Tag <> "MnuChangePassword" Then DesktopShortcuts ImgDesktopButton(Index).Tag
                  Exit Sub
               End If
            End If
         End If
      Next
   If mniDesktop = False Then DesktopReport.DesktopShortcutsRecport ImgDesktopButton(Index).Tag
End Sub

Private Sub ImgDesktopButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'some effect
    ImgBorder.BorderColor = RGB(86, 114, 162)
End Sub

Private Sub ImgDesktopButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Ind <> Index Then
      ImgBorder.Left = ImgDesktopButton(Index).Left - 45
      ImgBorder.Top = ImgDesktopButton(Index).Top - 45
      ImgBorder.Visible = True
      ImgBorder.ZOrder 1
   End If
   Ind = Index
End Sub

Private Sub ImgDesktopButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'some effect
    ImgBorder.BorderColor = RGB(172, 205, 224)
End Sub

Public Sub EnableShortcuts()
   On Error GoTo ErrorHandler
   vStr = " select d.Position, isnull(d.TaskKey, t.TaskKey) as TaskKey, isnull([Caption], [Description]) as Caption" & vbCrLf _
      + " from" & vbCrLf _
      + " (select " & ObjUserSecurity.UserNo & "  as UserNo,isnull(s.Position, d.position) as Position, isnull(s.TaskKey, d.TaskKey) as  TaskKey, [Caption]" & vbCrLf _
      + " from (select * from DesktopShortcuts where userno = " & ObjUserSecurity.UserNo & " ) s" & vbCrLf _
      + " right outer join DesktopShortcutsDefault d on s.position = d.position )d" & vbCrLf _
      + " inner join tasks t on t.taskkey = d.taskkey" & vbCrLf _
      + IIf(ObjUserSecurity.IsAdministrator = True, "", " inner join usertasks u on u.userno = d.userno and d.taskkey = u.taskkey") & vbCrLf _
      + " where isLocked = 0 "
      
   Dim vCounter As Integer
   With CN.Execute(vStr)
      For vCounter = 0 To 11
         .Filter = "Position = " & vCounter + 1
         ImgDesktopButton(vCounter).Tag = IIf(.RecordCount = 0, "MnuChangePassword", !TaskKey)
         LblCaption(vCounter).Caption = IIf(.RecordCount = 0, "", !Caption)
      Next vCounter
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Public Sub EnableMenus()
    On Error GoTo ErrorHandler
    Dim mnu As Object, sql As String
    If ObjUserSecurity.IsAdministrator = False Then
      sql = "Select t.TaskKey From Tasks t inner join (Select Distinct TaskKey From UserTasks Where UserNo = " & ObjUserSecurity.UserNo & ") u on t.TaskKey = u.TaskKey where isLocked = 0"
    Else
      sql = "Select TaskKey From Tasks Where isLocked = 0"
    End If
               
    With CN.Execute(sql)
        For Each mnu In Me.Controls
            If TypeOf mnu Is Menu And UCase(mnu.Name) Like "MNI*" Then
               If mnu.Checked = False Then mnu.Visible = True
               .Filter = "TaskKey = '" & mnu.Name & "'"
               If .RecordCount = 0 Then mnu.Visible = False
            End If
        Next
    End With
    Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Public Sub EnterMenus()
   On Error GoTo ErrorHandler
   CN.Execute ("Delete from Tasks")
   Dim mnu As Object
   For Each mnu In Me.Controls
      If TypeOf mnu Is Menu And UCase(mnu.Name) Like "MNI*" Then
         If mnu.Checked = False Then
            CN.Execute ("INSERT INTO TASKS VALUES ('" & mnu.Name & "','','" & mnu.Caption & "',0)")
         End If
      End If
   Next
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub UserActivities()
   LblUser.Caption = "Login Name : " & ObjUserSecurity.UserName
   'SavePicture1
'   EnterMenus
   EnableShortcuts
   EnableMenus
   a = Now
   a = 0
   Call Timer1_Timer
   ObjAccounts.InitializeSecurity ObjUserSecurity
   ObjDefinition.InitializeSecurity ObjUserSecurity
   ObjBank.InitializeSecurity ObjUserSecurity
   ObjPurchase.InitializeSecurity ObjUserSecurity
'   ObjPurchaseReports.InitializeSecurity ObjUserSecurity
   ObjSale.InitializeSecurity ObjUserSecurity
'   ObjSaleReports.InitializeSecurity ObjUserSecurity
   ObjProduction.InitializeSecurity ObjUserSecurity
   ObjStock.InitializeSecurity ObjUserSecurity
   ObjOthers.InitializeSecurity ObjUserSecurity
   ObjSaleReports.InitializeSecurity ObjUserSecurity
   Call MniPettyCashVerification_Click
   Call MniOpeningProductVerification_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = vbAltMask And KeyCode = vbKeyF4 Then
      Unload Me
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyM Then
      ObjUserSecurity.LogIn
      If ObjUserSecurity.IsLoginSuccess = False Then Exit Sub
      If ObjUserSecurity.UserNo <> 1 Then Exit Sub
      DefMenuLocked.Show
   ElseIf Shift = vbAltMask + vbCtrlMask + vbShiftMask And KeyCode = vbKeyF11 Then
      ObjUserSecurity.LogIn
      If ObjUserSecurity.IsLoginSuccess = False Then Exit Sub
      If ObjUserSecurity.UserNo <> 1 Then Exit Sub
      FrmDevelopers.Show
   ElseIf Shift = vbAltMask + vbCtrlMask + vbShiftMask And KeyCode = vbKeyF12 Then
      ObjUserSecurity.LogIn
      If ObjUserSecurity.IsLoginSuccess = False Then Exit Sub
      If ObjUserSecurity.UserNo <> 1 Then Exit Sub
      FrmCompanyInformation.TxtName.Enabled = True
      FrmCompanyInformation.TxtCity.Enabled = True
      FrmCompanyInformation.Show
      LblCompany.Caption = "Licensed To : " & ObjRegistry.CompanyName  'CN.Execute("select companyName from company").Fields(0).Value
   ElseIf Shift = vbShiftMask And KeyCode = vbKeyF10 Then
      Me.PopupMenu MnuMain, , 0, ImgStart.Top
   ElseIf KeyCode = 93 Then
      Me.PopupMenu MnuMain
   ElseIf Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyF6 Then
      MsgBox CN.Execute("select getdate()").Fields(0).Value
   ElseIf Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyF7 Then
      MsgBox Now
   ElseIf Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyF8 Then
      Dim vMin
      vMin = CN.Execute("select CMin from Counter").Fields(0).Value
      MsgBox RC4(FromHexDump(EStr(CStr(vMin), False)), vEncryptionString), vbOKOnly + vbInformation, "Current Min"
   ElseIf Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyF9 Then
      Dim vCurrentDate As String
      vCurrentDate = CN.Execute("select CLog from Counter").Fields(0).Value
      MsgBox Format(RC4(FromHexDump(EStr(vCurrentDate, False)), vEncryptionString), "dd/MM/yyyy"), vbOKOnly + vbInformation, "Current Date"
   ElseIf Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyF10 Then
      Dim vExpiryDate As String
      vExpiryDate = CN.Execute("select ELog from Counter").Fields(0).Value
      MsgBox Format(RC4(FromHexDump(EStr(vExpiryDate, False)), vEncryptionString), "dd/MM/yyyy"), vbOKOnly + vbInformation, "Expiry Date"
   ElseIf KeyCode = vbKeyF11 Then
      Dim vFileName
      vFileName = CN.Execute("select filename from sysfiles").Fields(0).Value
      MsgBox Left(vFileName, Len(vFileName) - InStr(1, StrReverse(vFileName), "\")), vbOKOnly + vbInformation, "Database Path"
   ElseIf KeyCode = vbKeyF12 Then
      MsgBox "This Software is updated on " & Format(ObjRegistry.UpdateDatabase, "dd-MMM-yyyy"), vbOKOnly + vbInformation, "Information"
   ElseIf Shift = vbShiftMask And KeyCode = vbKeyF10 Then
      Me.PopupMenu MnuMain, , 0, ImgStart.Top
   ElseIf KeyCode = vbKeyF8 Then
      Form2.Show vbModal
   End If
End Sub

'Private Sub Agent1_BalloonHide(ByVal CharacterID As String)
'    'Char1.Play "Blink"
'End Sub
'
'Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
'    If Button = 1 Then
'        Char1.Play "Surprised"
'    End If
'End Sub
'
'Private Sub Agent1_Command(ByVal UserInput As Object)
'    Char1.Stop
'    Char1.Speak UserInput.Name
'    Char1.Play UserInput.Name
'End Sub
'
'Private Sub Agent1_DblClick(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
'    Char1.Play "Alert"
'    Char1.Speak "Stop it!"
'End Sub
'
'Private Sub Agent1_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
'    Char1.Speak "Thank you"
'End Sub
'
'Private Sub Agent1_DragStart(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
'    Char1.Play "Alert"
'    Char1.Speak "Put Me Down!"
'End Sub
'
'Private Sub Agent1_Hide(ByVal CharacterID As String, ByVal Cause As Integer)
'    'AgentActive = False
'End Sub
'
'Private Sub Agent1_Show(ByVal CharacterID As String, ByVal Cause As Integer)
'    'AgentActive = True
'    Char1.Play "Greet"
'End Sub

Private Sub UpdateFile()
   On Error Resume Next
   With CN.Execute("Select CurrentDateUpdate from Counter")
      If IsNull(.Fields(0).Value) Or Format(.Fields(0).Value, "M/d/yyyy") <> Format(Date, "M/d/yyyy") Then
         'execute procedure
         CN.Execute "exec ProdUpdateDaily"
         CN.Execute "update Counter set CurrentDateUpdate = '" & Date & "'"
      End If
      .Close
   End With
End Sub

Private Sub CheckAutoBackup()
   On Error Resume Next
   Dim vSQL As String
   vSQL = "select * from MSDB..SysJobs"
   With CN.Execute(vSQL)
      If .RecordCount = 0 Then
         LblAutoBackup.Caption = "Auto Backup is not set."
      End If
      .Close
   End With
   vSQL = "SELECT SJ.name FROM MSDB..SysJobHistory SJH RIGHT JOIN MSDB..SysJobs SJ ON SJ.Job_Id = SJH.job_id WHERE (SJ.name LIKE '%7th%' or SJ.name LIKE '%8th%' ) AND Step_ID = 0 and run_status = 1" & _
          "group by SJ.name having datediff(day,REPLACE(CONVERT(varchar,convert(datetime,convert(varchar,max(run_date))),102),'.','-'),getdate()) >= 2"
   With CN.Execute(vSQL)
      If .RecordCount > 0 Then
         LblAutoBackup.Caption = "Auto Backup is not working properly."
      End If
      .Close
   End With
End Sub

Private Sub LogFile()
   On Error Resume Next
   If CN.Execute("select count(*) from ActivityLog").Fields(0).Value < 60000 Then Exit Sub
'   If objFSO.FolderExists("C:\windows\Log") = False Then
'      objFSO.CreateFolder "C:\windows\Log"
'   End If
   'vconnstr = SuperSoftv1;Data Source=(local)
   Dim Min As String, Max As String
   Min = Format(CN.Execute("select min(EntryDate) from ActivityLog").Fields(0).Value, "yyyyMMdd")
   Max = Format(CN.Execute("select Max(EntryDate) from ActivityLog").Fields(0).Value, "yyyyMMdd")
   Call DTSPackage(Min + " To " + Max, Right(vConnStr, Len(vConnStr) - InStr(1, vConnStr, "=")), Left(vConnStr, InStr(1, vConnStr, ";") - 1))
   If objFSO.FileExists(App.Path & "\database\" + Min + " To " + Max + ".txt") Then
      CN.Execute "Delete from ActivityLog"
   End If
End Sub

Private Sub LoadDesktopControlArray()
  Dim i As Integer
  
  For i = 1 To 11
    Load ImgDesktopButton(i)
    Load LblCaption(i)
    ImgDesktopButton(i).Visible = True
    LblCaption(i).Visible = True
    ImgDesktopButton(i).Top = ImgDesktopButton(i - 1).Top + 960
    LblCaption(i).Top = ImgDesktopButton(i).Top
    ImgDesktopButton(i).Left = 3701
    LblCaption(i).Left = 3975
    If i >= 6 Then
      ImgDesktopButton(i).Top = 2565 + ((i Mod 6) * 960)
      LblCaption(i).Top = ImgDesktopButton(i).Top
      ImgDesktopButton(i).Left = 3701 + 4000
      LblCaption(i).Left = 3975 + 4000
    End If
  Next i
End Sub
  
Private Sub Form_Load()
   On Error Resume Next
   Ind = 10
   LoadDesktopControlArray
   LogFile
   CheckAutoBackup
   UpdateFile
   ImgBorder.BorderColor = RGB(172, 205, 224)
   SetWindowText Me.hwnd, "Desktop"
   ShowPicture Me, 1
   '/*********** Actual ************************/
   ObjUserSecurity.LogIn
   If ObjUserSecurity.IsLoginSuccess = False Then End
   '/*******************************************/
   '/*************** Testing *********************/
   'ObjUserSecurity.UserNo = 1
   'ObjUserSecurity.LogInOK
   '/*******************************************/
   User1 = ObjUserSecurity.UserNo
   'ObjDefinition.InitializeUser User
   'ObjAccounts.InitializeUser User
   
   Timer2.Enabled = Not NotVisibility
'   Agent1.Characters.Load "Char1", App.Path & "\Peedy.acs"
'   Set Char1 = Agent1.Characters("Char1")
'   CharPopMnu
'   ObjSale.InitializeChar Char1
'   ObjPurchase.InitializeChar Char1
'   ObjProduction.InitializeChar Char1
'   ObjStock.InitializeChar Char1
'   Char1.MoveTo 600, 40
'   Char1.Show
'   Char1.Speak "Asslam o Alaikum"

   Flag = False
   LblCompany.Caption = "Licensed To : " & ObjRegistry.CompanyName
   CN.Execute "Delete From TempNo where UserNo = " & ObjUserSecurity.UserNo
   UserActivities
   
'   Call Triggers(True)
   '
   '/************ Testing ***************/
   'ObjPurchaseReports.ProductWisePurchaseDetailReport
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   If MsgBox("Are you sure to Exit Software?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then
      Cancel = 1
   End If
   If ObjUserSecurity.UserNo <> 0 Then
      If ObjRegistry.UseBin = True Then
         sSql = "insert into " & vBinDataBase & ".dbo.ActivityLogBin(ActivityDate,ActionNo,userno,FormNo, TempID, TransactionID,TransactionDate,TransactionInfo) values(getdate()," & eLogOut & "," & ObjUserSecurity.UserNo & ",'" & eFrmLogOut & "'," & "Null" & "," & "Null" & "," & "Null" & ",'Log Out  (" & LocalComputerName & ") By Exit Software Directtly')"
         CN.Execute sSql
      End If
   End If
'   UnloadAll
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
   
   
'   Set frmObj.Name = "FrmSaleInvoicePOS"
'   If objForm.Visible = True Then
'      MsgBox "forms " & frmObj.Name
'   End If
   
   If Cancel = 0 Then
      '/******* Mobile SMS *************/
      If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnLogin Then
         vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               sSql = ObjUserSecurity.UserName & " LogOut at " & Now
               sSql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & sSql & "','')"
               CN.Execute sSql
            End If
         Next
      End If
   End If
   objFSO.DeleteFile Left(App.Path, 2) & "\*.tmp", True
   objFSO.DeleteFile App.Path & "\*.tmp", True
   CN.Close
   Set CN = Nothing
   Set ObjAccounts = Nothing
   Set ObjAccountReports = Nothing
   Set ObjDefinition = Nothing
   Set ObjListReport = Nothing
   Set ObjProduction = Nothing
   Set ObjPurchase = Nothing
   Set ObjPurchaseReports = Nothing
   Set ObjSale = Nothing
   Set ObjSaleReports = Nothing
   Set ObjStock = Nothing
   Set ObjStockReports = Nothing
   Set ObjUserSecurity = Nothing
   Set ObjOthers = Nothing
   Unload DesktopReport
End Sub

'Private Sub ImgCurrentStock_Click()
'   If MniCurrentStockGroupWise.Visible Then ObjStockReports.CurrentStockReport
'End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

'Private Sub ImgLedger_Click()
'   If MniLedger.Visible Then ObjAccountReports.LedgerReport
'End Sub

Private Sub ImgMin_Click()
   Me.WindowState = 1
End Sub

'Private Sub ImgPaymentVocuher_Click()
'   If MniCashPaymentVoucher.Visible Then ObjAccounts.DebitVoucherForm
'End Sub

'Private Sub ImgPurchaseDetail_Click()
'   If MniGroupWisePurchaseDetail.Visible Then ObjPurchaseReports.GroupWisePurchaseDetailReport
'End Sub

'Private Sub ImgPurchaseInvoice_Click()
'   If MniPurchaseInvoice.Visible Then ObjPurchase.PurchaseInvoiceForm
'End Sub

'Private Sub ImgSaleDetail_Click()
'   If MniGroupWiseSaleDetail.Visible Then ObjSaleReports.GroupWiseSaleDetailReport
'End Sub

'Private Sub ImgSaleInvoice_Click()
'   If MniSaleInvoice.Visible Then ObjSale.SaleInvoiceForm
'End Sub

Private Sub ImgStart_Click()
   'MsgBox ImgStart.Top
   Me.PopupMenu MnuMain, , 0, ImgStart.Top '6400  '350
End Sub

Private Sub MniAccountsDefaultSetting_Click()
   If MniAccountsDefaultSetting.Visible Then ObjAccounts.AccountsDefaultSettingForm
End Sub

Private Sub MniActivityLog_Click()
   If MniActivityLog.Visible Then ObjUserSecurity.ActivityLogForm
End Sub

Private Sub MniAddSkin_Click()
   If MniAddSkin.Visible Then FrmAddSkin.Show
End Sub

Private Sub MniAdminClosing_Click()
   If MniAdminClosing.Visible Then ObjOthers.AdminClosingForm
End Sub


Private Sub MniAdvances_Click()
   If MniAdvances.Visible Then ObjAccounts.AdvancesForm
End Sub

Private Sub MniAssignTasks_Click()
   If MniAssignTasks.Visible = False Then Exit Sub
'   If ObjUserSecurity.IsAdministrator = 0 And ObjUserSecurity.UserNo <> 1 Then Exit Sub
   If ObjUserSecurity.IsAdministrator = True Or ObjUserSecurity.IsManager = True Or ObjUserSecurity.UserNo = 1 Then ObjUserSecurity.TaskAssignmentForm
   
End Sub

Private Sub MniAutoBackup_Click()
   If MniAutoBackup.Visible Then FrmAutoBackupDefulatDB.Show
End Sub


Private Sub MniBankMachines_Click()
   If MniBankMachines.Visible Then ObjDefinition.BankMachinesForm
End Sub

Private Sub MniBanquetInvoice_Click()
   If MniBanquetInvoice.Visible Then ObjOthers.BanquetInvoiceForm
End Sub

Private Sub MniBanquetOrder_Click()
   If MniBanquetOrder.Visible Then ObjOthers.BanquetOrderForm
End Sub


Private Sub MniBillWiseProfit_Click()
'   If MniBillWiseProfit.Visible Then ObjAccountReports.BillWiseProfitReport
End Sub

Private Sub MniBrands_Click()
   If MniBrands.Visible Then ObjDefinition.BrandsForm
End Sub

Private Sub MniChangeCategory_Click()
   If MniChangeCategory.Visible Then ObjDefinition.ChangeCategoriesForm
End Sub

Private Sub MniClasses_Click()
   If MniClasses.Visible Then ObjDefinition.ClassesForm
End Sub

Private Sub MniCompanyInformation_Click()
   If MniCompanyInformation.Visible = False Then Exit Sub
   FrmCompanyInformation.TxtName.Enabled = False
   FrmCompanyInformation.TxtCity.Enabled = False
   FrmCompanyInformation.Show
End Sub


Private Sub MniCourses_Click()
   If MniCourses.Visible Then ObjDefinition.CoursesForm
End Sub

Private Sub MniCustomerDemand_Click()
   If MniCustomerDemand.Visible Then ObjOthers.CustomerDemandForm
End Sub


Private Sub MniCustomerTypes_Click()
   If MniCustomerTypes.Visible Then ObjDefinition.CustomerTypesForm
End Sub


Private Sub MniCustomOrderDelivery_Click()
   If MniCustomOrderDelivery.Visible Then ObjSale.CustomOrderDeliveryForm
End Sub

Private Sub MniCustomOrderPurchase_Click()
   If MniCustomOrderPurchase.Visible Then ObjPurchase.CustomOrderPurchaseForm
End Sub

Private Sub MniCustomOrderReturn_Click()
   If MniCustomOrderReturn.Visible Then ObjSale.CustomOrderReturnForm
End Sub


Private Sub MniDefineStockLimit_Click()
   If MniDefineStockLimit.Visible Then ObjDefinition.DefineStockLimitForm
End Sub


Private Sub MniDisputeInvoice_Click()
   If MniDisputeInvoice.Visible Then ObjStock.DisputeInvoiceForm
End Sub

Private Sub MniEmpDepartments_Click()
   If MniEmpDepartments.Visible Then ObjDefinition.EmpDepartmentsForm
End Sub

Private Sub MniEmployeeAttendanceAll_Click()
   If MniEmployeeAttendanceAll.Visible Then ObjOthers.EmployeeAttendanceForm
End Sub

Private Sub MniEmployeeAttendanceIn_Click()
   If MniEmployeeAttendanceIn.Visible Then ObjOthers.EmployeeAttendanceInForm
End Sub

Private Sub MniEmployeeAttendanceOut_Click()
   If MniEmployeeAttendanceOut.Visible Then ObjOthers.EmployeeAttendanceOutForm
End Sub


Private Sub MniExpiryDayColor_Click()
   If MniExpiryDayColor.Visible Then ObjDefinition.ExpiryDayColorForm
End Sub

Private Sub MniExpSetting_Click()
   If MniExpSetting.Visible Then ObjAccounts.ExpenseSettingForm
End Sub

Private Sub MniGatePassIn_Click()
   If MniGatePassIn.Visible Then ObjProduction.GatePassInForm
End Sub

Private Sub MniGatePassOut_Click()
   If MniGatePassOut.Visible Then ObjProduction.GatePassOutForm
End Sub




Private Sub MniGRN_Click()
    If MniGRN.Visible Then ObjPurchase.GRNForm
End Sub

Private Sub MniHoliday_Click()
   If MniHoliday.Visible Then ObjOthers.HolidayForm
End Sub

Private Sub MniEmployeeLeave_Click()
   If MniEmployeeLeave.Visible Then ObjOthers.EmployeeLeaveForm
End Sub


Private Sub MniInstallmentCustomers_Click()
   If MniInstallmentCustomers.Visible Then ObjDefinition.InstallmentCustomersForm
End Sub


Private Sub MniLiftInvoice_Click()
   If MniLiftInvoice.Visible Then ObjStock.LiftInvoiceForm
End Sub

Private Sub MniLoans_Click()
   If MniLoans.Visible Then ObjAccounts.LoansForm
End Sub


Private Sub MniManufacturedReturn_Click()
   If MniManufacturedReturn.Visible Then ObjProduction.ManufacturedReturnForm
End Sub



Private Sub MniMembers_Click()
   If MniMembers.Visible Then ObjDefinition.MembersForm
End Sub

Private Sub MniMembersDiscount_Click()
   If MniMembersDiscount.Visible Then ObjDefinition.MembersDiscountForm
End Sub
Private Sub MniCommisionDisc_Click()
   If MniCommisionDisc.Visible Then ObjDefinition.CommisionDiscForm
End Sub

Private Sub MniMemberTypes_Click()
   If MniMemberTypes.Visible = True Then ObjDefinition.MemberTypesForm
End Sub

Private Sub MniMeterReadings_Click()
   If MniMeterReadings.Visible Then ObjOthers.MeterReadingsForm
End Sub

Private Sub MniMonthlyIncomeExpense_Click()
   If MniMonthlyIncomeExpense.Visible Then ObjOthers.MonthlyIncomeExpenseForm
End Sub

Private Sub MniOldActivityLog_Click()
    If MniOldActivityLog.Visible Then ObjUserSecurity.OldActivityLogForm
End Sub

Private Sub MniOpeningAccounts_Click()
    If MniOpeningAccounts.Visible Then ObjAccounts.AccountsOpeningBalanceForm
End Sub

Private Sub MniOpeningAccountsOrganization_Click()
   If MniOpeningAccountsOrganization.Visible Then ObjAccounts.OrganizationalOpeningBalanceForm
End Sub

Private Sub MniOpeningProduct_Click()
   If MniOpeningProduct.Visible Then ObjOthers.OpeningProductsForm
End Sub

Private Sub MniOpeningProductVerification_Click()
   With CN.Execute("select ID, isVerify from OpeningProductHeader where ToUserNo = " & ObjUserSecurity.UserNo & " and EntryDate in (select max(EntryDate) from OpeningProductHeader where ToUserNo = " & ObjUserSecurity.UserNo & ")")
      If .RecordCount > 0 Then
         If !isverify = 0 Then
            ObjOthers.OpeningProductsVerificationForm
         End If
      End If
   End With
End Sub

Private Sub MniOrganization_Click()
   If MniOrganization.Visible Then ObjDefinition.OrganizationForm
End Sub
Private Sub MniPettyCashVerification_Click()
   With CN.Execute("select ID, isVerify from PettyCashHeader where ToUserNo = " & ObjUserSecurity.UserNo & " and EntryDate in (select max(EntryDate) from PettyCashHeader where ToUserNo = " & ObjUserSecurity.UserNo & ")")
      If .RecordCount > 0 Then
         If !isverify = 0 Then
            ObjOthers.PettyCashVerificationForm
         End If
      End If
   End With
End Sub

Private Sub MniPLSettingWithoutOrg_Click()
    If MniPLSettingWithoutOrg.Visible Then ObjAccounts.PLSettingsWithoutOrgForm
End Sub

Private Sub MniPostSale_Click()
   If MniPostSale.Visible Then ObjSale.PostSaleForm
End Sub

Private Sub MniPriceChecker_Click()
   FrmPriceChecker.Show
End Sub

Private Sub MniPriceVariation_Click()
   If MniPriceVariation.Visible Then ObjStock.PriceVariationForm
End Sub


Private Sub MniProductsDetail_Click()
   If MniProductsDetail.Visible Then ObjDefinition.ProductsDetailForm
End Sub

Private Sub MniProductionIN_Click()
   If MniProductionIN.Visible Then ObjProduction.ProductionInForm
End Sub

Private Sub MniProductionOut_Click()
   If MniProductionOut.Visible Then ObjProduction.ProductionOutForm
End Sub


Private Sub MniProductOffer_Click()
   If MniProductOffer.Visible Then ObjDefinition.ProductOfferForm
End Sub


Private Sub MniPurchaseOrder_Click()
   If MniPurchaseOrder.Visible Then ObjPurchase.PurchaseOrderForm
End Sub


Private Sub MniRacks_Click()
   If MniSubGroup.Visible Then ObjDefinition.RacksForm
End Sub

Private Sub MniReceivingInvoicePrint_Click()
    If MniReceivingInvoicePrint.Visible Then ObjSale.ReceivingInvoicePrint
End Sub

Private Sub MniRecoveryInvoiceWise_Click()
   If MniRecoveryInvoiceWise.Visible Then ObjSale.RecoveryInvoiceWiseForm
End Sub


Private Sub MniReferences_Click()
   If MniReferences.Visible Then ObjDefinition.ReferencesForm
End Sub


Private Sub MniSaleCounters_Click()
   If MniSaleCounters.Visible Then ObjDefinition.SaleCountersForm
End Sub


Private Sub MniSaleInvoicePOS_Click()
   If MniSaleInvoicePOS.Visible Then ObjSale.SaleInvoicePOSForm
End Sub

Private Sub MniSaleInvoiceTouch_Click()
If MniSaleInvoiceTouch.Visible Then ObjSale.SaleInvoiceTouchForm
End Sub

Private Sub MniSaleOrder_Click()
   If MniSaleOrder.Visible Then ObjSale.SaleOrderForm
End Sub

Private Sub MniSaleOrderPOS_Click()
   If MniSaleOrderPOS.Visible Then ObjSale.SaleOrderPOSForm
End Sub


Private Sub MniSaleReturnInvoicePOS_Click()
   If ObjRegistry.UsePasswordForm = True And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False Then
      If UsePasswordForm = False Then Exit Sub
   End If
   If MniSaleReturnInvoicePOS.Visible Then ObjSale.SaleReturnInvoicePOSForm
End Sub

Private Sub MniSchool_Click()
   If MniCourses.Visible Then ObjDefinition.SchoolForm
End Sub

Private Sub MniSectors_Click()
   If MniSectors.Visible Then ObjDefinition.SectorsForm
End Sub

Private Sub MniServiceInvoice_Click()
'   If MniServiceInvoice.Visible Then ObjSale.ServiceInvoiceForm
End Sub

Private Sub MniServiceProducts_Click()
'   If MniServiceProducts.Visible Then ObjDefinition.ServiceProductsForm
End Sub

Private Sub MniSessions_Click()
   If MniSessions.Visible Then ObjDefinition.SessionsForm
End Sub

Private Sub MniShifts_Click()
   If MniShifts.Visible Then ObjDefinition.ShiftsForm
End Sub

Private Sub MniSkinSelection_Click()
   If MniSkinSelection.Visible Then FrmSkinSelection.Show
End Sub


Private Sub MniSoftwareDefaultSettings_Click()
   If MniSoftwareDefaultSettings.Visible Then FrmSoftwareDefaultSettings.Show
End Sub


Private Sub MniStockLimitStoreWise_Click()
   If MniStockLimitStoreWise.Visible Then ObjDefinition.StockLimitStoreWiseForm
End Sub

Private Sub MniSyllabus_Click()
   If MniSyllabus.Visible Then ObjDefinition.SyllabusForm
End Sub

Private Sub MniSyncData_Click()
   If MniSyncData.Visible Then FrmSyncData.Show
End Sub

'Private Sub MniStockPendingTransfer_Click()
'   If MniStockPendingTransfer.Visible Then ObjStock.StockPendingTransferForm
'End Sub


Private Sub MniTables_Click()
   If MniTables.Visible Then ObjDefinition.TablesForm
End Sub

Private Sub MniZones_Click()
   If MniZones.Visible Then ObjDefinition.ZonesForm
End Sub

Private Sub MnuReports_Click()
   Dim a As New DesktopReport
    a.Show
End Sub

Private Sub MnuSMS_Click()
   FrmSMS.Show
End Sub

Private Sub MniCashDeposite_Click()
   If MniCashDeposite.Visible Then ObjBank.CashDepositForm
End Sub

Private Sub MniChangePrice_Click()
   If MniChangePrice.Visible Then ObjDefinition.ChangePriceForm
End Sub

Private Sub MniCustomOrderBooking_Click()
   If MniCustomOrderBooking.Visible Then ObjSale.CustomOrderForm
End Sub

Private Sub MniCustomProductsAndMeasurements_Click()
   If MniCustomProductsAndMeasurements.Visible Then ObjDefinition.CustomProductsandMeasurementsForm
End Sub

Private Sub MniDepartments_Click()
   If MniDepartments.Visible Then ObjDefinition.DepartmentsForm
End Sub

Private Sub MniDesignations_Click()
   If MniDesignations.Visible Then ObjDefinition.DesignationsForm
End Sub

Private Sub MniEmployees_Click()
   If MniEmployees.Visible Then ObjDefinition.EmpolyeesForm
End Sub

'Private Sub MniExportProducts_Click()
''   on error
'''   'Open "C:\Export.bat" For Output As #1
'''   'Print #1, "bcp """ & Left(vConnStr, InStr(1, vConnStr, ";") - 1) & ".dbo.Products"" out ""C:\Product.txt"" -c -C 850 -S""" & Replace(Right(vConnStr, Len(vConnStr) - InStr(1, vConnStr, "=")), ";", "") & """ -T "
'''   'Close #1
'''   Shell "bcp """ & Left(vConnStr, InStr(1, vConnStr, ";") - 1) & ".dbo.Products"" out """ & App.Path & "\Product.txt"" -c -C 850 -S""" & Replace(Right(vConnStr, Len(vConnStr) - InStr(1, vConnStr, "=")), ";", "") & """ -T ", vbHide
'''   'Shell "C:\Export.bat"
'
'
''   Me.MousePointer = vbHourglass
''   Call DTSSQLToXLS("G:\Test.XLS", Right(vConnStr, Len(vConnStr) - InStr(1, vConnStr, "=")), Left(vConnStr, InStr(1, vConnStr, ";") - 1))
''   Me.MousePointer = vbDefault
''   MsgBox "OK", vbInformation + vbOKOnly, "Information"
'   Me.MousePointer = vbHourglass
'   objFSO.DeleteFile App.Path & "\Data.xls", True
'   objFSO.CopyFile App.Path & "\Format.xls", App.Path & "\Data.xls", True
'   Call Export
'   Me.MousePointer = vbDefault
'   MsgBox "Export Data Successfully.", vbInformation + vbOKOnly, "Information"
'End Sub

Private Sub MniFinishedProduct_Click()
   ObjProductionReport.FinishedProductReport
End Sub

Private Sub MniFormulaInfo_Click()
   ObjDefinition.FormulaInfoForm
End Sub

'Private Sub MniImportProducts_Click()
'
''
''   Call DTSXLSToSQL("G:\Test", Right(vConnStr, Len(vConnStr) - InStr(1, vConnStr, "=")))
''   'CN.Execute "delete from TempProduct"
''   'Shell "bcp """ & Left(vConnStr, InStr(1, vConnStr, ";") - 1) & ".dbo.TempProduct"" in """ & App.Path & "\Product.txt"" -c -C 850 -S""" & Replace(Right(vConnStr, Len(vConnStr) - InStr(1, vConnStr, "=")), ";", "") & """ -T ", vbHide
''   'Timer2.Enabled = True
''
'   Me.MousePointer = vbHourglass
'   Dim CN1 As New ADODB.Connection
'   If CN1.State = adStateOpen Then CN1.Close
'   CN1.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Tempdb;Data Source=" & Right(vConnStr, Len(vConnStr) - InStr(1, vConnStr, "="))
'   CN1.CursorLocation = adUseClient
'   Dim vCommands() As String, i As Integer
'   vCommands = Split(TxtTempTables.Text, "GO")
'   For i = 0 To UBound(vCommands)
'       CN1.Execute (Replace(Replace(vCommands(i), Chr(255), ""), Chr(254), ""))
'   Next
'   Call Import
'   Me.MousePointer = vbDefault
'   MsgBox "Import Data Successfully.", vbInformation + vbOKOnly, "Information"
'End Sub

Private Sub MniLockAccounts_Click()
   If MniLockAccounts.Visible Then ObjAccounts.LockAccountsForm
End Sub

Private Sub MniOpeningStock_Click()
   If MniOpeningStock.Visible Then ObjDefinition.OpeningStockForm
End Sub

Private Sub MniPaymentInvoiceWise_Click()
   If MniPaymentInvoiceWise.Visible Then ObjPurchase.PaymentInvoiceWiseForm
End Sub

Private Sub MniPaymentVender_Click()
   If MniPaymentVender.Visible Then ObjPurchase.PaymentVenderWiseForm
End Sub

Private Sub MniPettyCash_Click()
   If MniPettyCash.Visible Then ObjOthers.PettyCashForm
End Sub

Private Sub MniPLSetting_Click()
   If MniPLSetting.Visible Then ObjAccounts.PLSettingsForm
End Sub


Private Sub MniRecoveryCustomer_Click()
   If MniRecoveryCustomer.Visible Then ObjSale.RecoveryCustomerWiseForm
End Sub

Private Sub MniReplacementInvoice_Click()
   If MniReplacementInvoice.Visible Then ObjSale.ReplacementInvoiceForm
End Sub

Private Sub MniSalary_Click()
   If MniSalary.Visible Then ObjOthers.SalaryForm
End Sub

Private Sub MniStockAdjustment_Click()
   If MniStockAdjustment.Visible Then ObjDefinition.StockAdjustmentForm
End Sub

Private Sub MniPackageDealInfo_Click()
   If MniPackageDealInfo.Visible Then ObjDefinition.PackageDealInfoForm
End Sub

Private Sub MniUnits_Click()
'   If MniUnits.Visible Then ObjDefinition.UnitsForm
End Sub

Private Sub MnuUserClosing_Click()
   If MnuUserClosing.Visible Then ObjOthers.UserClosingForm
End Sub

Private Sub MniUserDefaultSettings_Click()
   If MniUserDefaultSettings.Visible Then ObjUserSecurity.UserDefaultSettingsForm
End Sub

Private Sub MnuChangePassword_Click()
   If MnuChangePassword.Visible Then ObjUserSecurity.ChangePasswordForm
End Sub
 
Private Sub MniChartofAccounts_Click()
   If MniChartofAccounts.Visible Then ObjAccounts.ChartOfAccountsForm
End Sub

Private Sub MniChequeDeposit_Click()
   If MniChequeDeposit.Visible Then ObjBank.ChequeDepositForm
End Sub

Private Sub MniChequeIssuance_Click()
   If MniChequeIssuance.Visible Then ObjBank.ChequeIssuanceForm
End Sub

Private Sub MniChequeIssuanceReconcilation_Click()
   If MniChequeIssuanceReconcilation.Visible Then ObjBank.ChequeIssueReconciliationForm
End Sub

Private Sub MniChequeReceive_Click()
   If MniChequeReceive.Visible Then ObjBank.ChequeReceiveForm
End Sub

Private Sub MniChequeReceiveReconcilation_Click()
   If MniChequeReceiveReconcilation.Visible Then ObjBank.ChequeReceiveReconciliationForm
End Sub

Private Sub MniCompany_Click()
   If MniCompany.Visible Then ObjDefinition.CompaniesForm
End Sub

Private Sub MniCustomers_Click()
   If MniCustomers.Visible Then ObjDefinition.CustomersForm
End Sub

Private Sub MniExpiryDamageClaimInvoice_Click()
   If MniExpiryDamageClaimInvoice.Visible Then ObjStock.ExpiryDamageClaimInvoiceForm
End Sub

Private Sub MniExpiryDamageInvoice_Click()
   If MniExpiryDamageInvoice.Visible Then ObjStock.ExpiryDamageInvoiceForm
End Sub

Private Sub MniStockWastageInvoice_Click()
   If MniStockWastageInvoice.Visible Then ObjStock.StockWastageInvoiceForm
End Sub

Private Sub MniStores_Click()
   If MniStores.Visible Then ObjDefinition.StoresForm
End Sub

Private Sub MnuDesktopSetting_Click()
   If MnuDesktopSetting.Visible Then FrmDesktopSetting.Show
End Sub

Private Sub MnuLogOut_Click()
   On Error GoTo ErrorHandler
   If MsgBox("Are you sure to log out?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
   Me.Hide
     
   '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnLogin Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
      For i = 0 To UBound(vMobileNo)
         vMobile = "+92" + Right(vMobileNo(i), 10)
         If Len(vMobile) = 13 Then
            sSql = ObjUserSecurity.UserName & " LogOut at " & Now
            sSql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & sSql & "','')"
            CN.Execute sSql
         End If
      Next
   End If
   Unload DesktopReport
'   CN.Execute "Exec ProdActivityLog 'Logout'," & ObjUserSecurity.UserNo & ",1," & ObjUserSecurity.UserNo
   If ObjRegistry.UseBin = True Then
      sSql = "insert into " & vBinDataBase & ".dbo.ActivityLogBin(ActivityDate,ActionNo,userno,FormNo, TempID, TransactionID,TransactionDate,TransactionInfo) values(getdate()," & eLogOut & "," & ObjUserSecurity.UserNo & ",'" & eFrmLogOut & "'," & "Null" & "," & "Null" & "," & "Null" & ",'Log Out (" & LocalComputerName & ") Successfully at " & Now & "')"
      CN.Execute sSql
   End If
   ObjUserSecurity.LogOut
   If ObjUserSecurity.UserNo = 0 Then
      Unload DesktopReport
      Unload Me
      Exit Sub
   End If
   Me.Show
   UserActivities
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub MniBackup_Click()
   On Error GoTo ErrorHandler
   CD1.FileName = ""
   CD1.DialogTitle = "Enter Path To Take Backup"
   CD1.InitDir = App.Path
   CD1.Filter = "(Backup Files)|*.bak"
   CD1.ShowSave
   If CD1.FileName <> "" Then
      Me.MousePointer = vbHourglass
      Dim vConnStr As String
      Open App.Path & "\Config.ini" For Input As #1
      Input #1, vConnStr
      Close #1
      CNR.Execute " exec procBACKUPDATABASE '" & Left(vConnStr, InStr(1, vConnStr, ";") - 1) & "','" & CD1.FileName & "','" & Left(vConnStr, InStr(1, vConnStr, ";") - 1) & "'"
      Me.MousePointer = vbDefault
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
   Me.MousePointer = vbDefault
End Sub

Private Sub MniCashPaymentVoucher_Click()
   If MniCashPaymentVoucher.Visible Then ObjAccounts.DebitVoucherForm
End Sub

Private Sub MniCashReceiveVoucher_Click()
   If MniCashReceiveVoucher.Visible Then ObjAccounts.CreditVoucherForm
End Sub

Private Sub MniGroups_Click()
   If MniGroups.Visible Then ObjDefinition.GroupsForm
End Sub

Private Sub MniJournalVoucher_Click()
   If MniJournalVoucher.Visible Then ObjAccounts.JournalVoucherForm
End Sub

Private Sub MniManufacturedProducts_Click()
   If MniManufacturedProducts.Visible Then ObjProduction.ManufacturedProductsForm
End Sub

'Private Sub MniMergeProducts_Click()
'   If MniMergeProducts.Visible Then ObjDefinition.MergeProductsForm
'End Sub


'Private Sub MniOpeningAccounts_Click()
'   If MniOpeningAccounts.Visible Then ObjAccounts.AccountsOpeningBalanceForm
'End Sub

Private Sub MniPacking_Click()
   If MniPacking.Visible Then ObjDefinition.PackingsForm
End Sub

Private Sub MniProducts_Click()
   If MniProducts.Visible Then ObjDefinition.ProductsForm
End Sub

Private Sub MniPurchaseInvoice_Click()
   If MniPurchaseInvoice.Visible Then ObjPurchase.PurchaseInvoiceForm
End Sub

Private Sub MniPurchaseReturnInvoice_Click()
   If MniPurchaseReturnInvoice.Visible Then ObjPurchase.PurchaseReturnInvoiceForm
End Sub

Private Sub MniRestore_Click()
   On Error GoTo ErrorHandler
   CD1.FileName = ""
   CD1.DialogTitle = "Select File To Restore Database"
   CD1.InitDir = App.Path
   CD1.Filter = "(Backup Files)|*.bak"
   CD1.ShowOpen
   If CD1.FileName <> "" Then
      ObjUserSecurity.LogIn
      If ObjUserSecurity.IsLoginSuccess = False Then Exit Sub
      Me.MousePointer = vbHourglass
      Dim vConnStr As String
      Open App.Path & "\Config.ini" For Input As #1
      Input #1, vConnStr
      Close #1
      Shell "net stop server /y", vbHide
      CN.DefaultDatabase = "master"
      CN.Execute "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ProcRESTOREDATABASE]') and OBJECTPROPERTY(id, N'IsProcedure') = 1) drop procedure [dbo].[ProcRESTOREDATABASE]"
      CN.Execute TxtProcList.Text
      With CN.Execute("RESTORE FILELISTONLY FROM DISK = '" & CD1.FileName & "'")
         If .RecordCount > 0 Then
            CN.Execute "exec ProcRESTOREDATABASE '" & Left(vConnStr, InStr(1, vConnStr, ";") - 1) & "','" & App.Path & "\Database\" & "','" & CD1.FileName & "','" & Left(!LogicalName, Len(!LogicalName) - 5) & "'"
         End If
         .Close
      End With
      CN.DefaultDatabase = Left(vConnStr, InStr(1, vConnStr, ";") - 1)
      Shell "net start server /y", vbHide
      Me.MousePointer = vbDefault
      
      Me.Hide
      CN.Execute "Exec ProdActivityLog 'Logout'," & ObjUserSecurity.UserNo & ",1," & ObjUserSecurity.UserNo
      ObjUserSecurity.LogOut
      If ObjUserSecurity.UserNo = 0 Then Unload Me: Exit Sub
      Me.Show
      UserActivities
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
   CN.DefaultDatabase = Left(vConnStr, InStr(1, vConnStr, ";") - 1)
   Me.MousePointer = vbDefault
End Sub

Private Sub MniSaleInvoice_Click()
   If MniSaleInvoice.Visible Then ObjSale.SaleInvoiceForm
End Sub

Private Sub MniSaleReturnInvoice_Click()
   If ObjRegistry.UsePasswordForm = True And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False Then
      If UsePasswordForm = False Then
         MsgBox "Incorrect Password.", vbExclamation, "Alert"
         Exit Sub
      End If
   End If
   If MniSaleReturnInvoice.Visible Then ObjSale.SaleReturnInvoiceForm
End Sub

Private Sub MniSetCurrentStock_Click()
   If ObjUserSecurity.UserNo = 1 Then ObjDefinition.SetCurrentStockForm: Exit Sub
   With CN.Execute("select * from registry")
      If .RecordCount > 0 Then
         If !SetCurrentStock = True And ObjUserSecurity.IsAdministrator = True Then
            ObjDefinition.SetCurrentStockForm
         End If
      End If
      .Close
   End With
End Sub

Private Sub MniStockTransferInvoice_Click()
   If MniStockTransferInvoice.Visible Then ObjStock.StockTransferForm
End Sub

Private Sub MniSubGroup_Click()
   If MniSubGroup.Visible Then ObjDefinition.SubGroupsForm
End Sub

Private Sub MniUsers_Click()
   If MniUsers.Visible Then ObjUserSecurity.UsersForm
End Sub

Private Sub MniVendors_Click()
   If MniVendors.Visible Then ObjDefinition.VendorsForm
End Sub

Private Sub Timer1_Timer()
   LblLoginTime.Caption = "Login Time : " & Format(a, "h:mm")
   a = DateAdd("n", 1, a)
End Sub

'Private Sub Timer2_Timer()
'   Dim sSql As String
'   sSql = " Update Products Set ProductName = TempProduct.ProductName, " & vbCrLf _
'      + "      CompanyID = TempProduct.CompanyID, " & vbCrLf _
'      + "      GroupID = TempProduct.GroupID, " & vbCrLf _
'      + "      SubGroupID = TempProduct.SubGroupID, " & vbCrLf _
'      + "      PurPrice = TempProduct.PurPrice, " & vbCrLf _
'      + "      RetailPrice = TempProduct.RetailPrice, " & vbCrLf _
'      + "      PurchasePackingID = TempProduct.PurchasePackingID, " & vbCrLf _
'      + "      DiscPer = TempProduct.DiscPer, " & vbCrLf _
'      + "      DiscPC = TempProduct.DiscPC" & vbCrLf _
'      + "      FROM TempProduct" & vbCrLf _
'      + "      WHERE Products.ProductID = TempProduct.ProductID"
'   CN.Execute sSql
'   CN.Execute "ALTER TABLE Products DISABLE TRIGGER [ti_Products]"
'   CN.Execute "insert into Products select t.* from Products p full outer join TempProduct t on p.ProductID = t.ProductID where p.productid is null"
'   CN.Execute "insert into CurrentStock select t.ProductID, 0 as QtyLoose, 0 as Cost from CurrentStock s right outer join TempProduct t on t.ProductID = s.ProductID where s.productid is null"
'   CN.Execute "insert into CurrentStockStore select t.ProductID, 1 as Storeid, 0 as QtyLoose from CurrentStockStore s right outer join TempProduct t on t.ProductID = s.ProductID where s.productid is null"
'   CN.Execute "ALTER TABLE Products ENABLE TRIGGER [ti_Products]"
'   Me.MousePointer = vbDefault
'   MsgBox "Import Product Successfully.", vbOKOnly + vbInformation, "Alert"
'   Timer2.Enabled = False
'End Sub

Private Sub DesktopShortcuts(ShortcutName As String)
   Select Case ShortcutName
      Case "MniAccountPayable"
'         Call MniAccountPayable_Click
      Case "MniAccountReceivable"
'         Call MniAccountReceivable_Click
'      Case "MniAccountDefaultSettings"
'         Call MniAccountDefaultSettings_Click
      Case "MniActivityLog"
         Call MniActivityLog_Click
      Case "MniAddSkin"
         Call MniAddSkin_Click
      Case "MniAdminClosing"
         Call MniAdminClosing_Click
      Case "MniAdminClosingReport"
'         Call MniAdminClosingReport_Click
      Case "MniAdvances"
         Call MniAdvances_Click
      Case "MniAssignTasks"
         Call MniAssignTasks_Click
      Case "MniAutoBackup"
         Call MniAutoBackup_Click
      Case "MniBackup"
         Call MniBackup_Click
      Case "MniBalanceSheet"
'         Call MniBalanceSheet_Click
      Case "MniBankCashDepositReport"
'         Call MniBankCashDepositReport_Click
      Case "MniBankChequeDepositReport"
'         Call MniBankChequeDepositReport_Click
      Case "MniBankChequeIssuanceReport"
'         Call MniBankChequeIssuanceReport_Click
      Case "MniBankChequeReceiveReport"
'         Call MniBankChequeReceiveReport_Click
      Case "MniBankMachines"
         Call MniBankMachines_Click
      Case "MniBatchExpiryReport"
'         Call MniBatchExpiryReport_Click
      Case "MniBrands"
         Call MniBrands_Click
      Case "MniCashBook"
'         Call MniCashBook_Click
      Case "MniCashDeposite"
         Call MniCashDeposite_Click
      Case "MniCashFlow"
'         Call MniCashFlow_Click
      Case "MniCashPaymentVoucher"
         Call MniCashPaymentVoucher_Click
      Case "MniCashReceiveVoucher"
         Call MniCashReceiveVoucher_Click
      Case "MniChangePrice"
         Call MniChangePrice_Click
      Case "MniChartofAccounts"
         Call MniChartofAccounts_Click
      Case "MniChequeDeposit"
         Call MniChequeDeposit_Click
      Case "MniChequeIssuance"
         Call MniChequeIssuance_Click
      Case "MniChequeIssuanceReconcilation"
         Call MniChequeIssuanceReconcilation_Click
      Case "MniChequeReceive"
         Call MniChequeReceive_Click
      Case "MniChequeReceiveReconcilation"
         Call MniChequeReceiveReconcilation_Click
      Case "MniChangeCategory"
         Call MniChangeCategory_Click
      Case "MniCompany"
         Call MniCompany_Click
      Case "MniCompanyInformation"
         Call MniCompanyInformation_Click
      Case "MniCurrentStockExpiryValue"
'         Call MniCurrentStockExpiryValue_Click
      Case "MniCurrentStockWastage"
'         Call MniCurrentStockWastage_Click
      Case "MniCustomOrderBalance"
'         Call MniCustomOrderBalance_Click
      Case "MniCustomOrderBooking"
         Call MniCustomOrderBooking_Click
      Case "MniCustomOrderDelivery"
         Call MniCustomOrderDelivery_Click
      Case "MniCustomOrderPurchase"
         Call MniCustomOrderPurchase_Click
      Case "MniCustomOrderReturn"
         Call MniCustomOrderReturn_Click
      Case "MniCustomProductsAndMeasurements"
         Call MniCustomProductsAndMeasurements_Click
      Case "MniCustomerDemand"
         Call MniCustomerDemand_Click
      Case "MniCustomerDemandRegister"
'         Call MniCustomerDemandRegister_Click
      Case "MniCustomerList"
'         Call MniCustomerList_Click
      Case "MniCustomerSaleBills"
'         Call MniCustomerSaleBills_Click
      Case "MniCustomers"
         Call MniCustomers_Click
      Case "MniDailyActivity"
'         Call MniDailyActivity_Click
      Case "MniDateWiseProductionInOut"
'         Call MniDateWiseProductionInOut_Click
      Case "MniDateWiseProfit"
'         Call MniDateWiseProfit_Click
      Case "MniDateWiseProfitStoreWise"
'         Call MniDateWiseProfitStoreWise_Click
      Case "MniDateWiseSaleExpense"
'         Call MniDateWiseSaleExpense_Click
      Case "MniDeadProducts"
'         Call MniDeadProducts_Click
      Case "MniDefineStockLimit"
         Call MniDefineStockLimit_Click
      Case "MniDemandList"
'         Call MniDemandList_Click
      Case "MniDemandListStoreWise"
'         Call MniDemandListStoreWise_Click
      Case "MniDepartments"
         Call MniDepartments_Click
      Case "MniDesignations"
         Call MniDesignations_Click
      Case "MniDisputeInvoice"
         Call MniDisputeInvoice_Click
      Case "MniEmployeeAttendanceAll"
         Call MniEmployeeAttendanceAll_Click
      Case "MniEmployeeAttendanceIn"
         Call MniEmployeeAttendanceIn_Click
      Case "MniEmployeeAttendanceOut"
         Call MniEmployeeAttendanceOut_Click
      Case "MniEmployeeAttendanceReport"
'         Call MniEmployeeAttendanceReport_Click
      Case "MniEmployeeLeave"
         Call MniEmployeeLeave_Click
      Case "MniEmployeeList"
'         Call MniEmployeeList_Click
      Case "MniEmployees"
         Call MniEmployees_Click
      Case "MniExpSetting"
         Call MniExpSetting_Click
      Case "MniExpiryDamageClaimInvoice"
         Call MniExpiryDamageClaimInvoice_Click
      Case "MniExpiryDamageInvoice"
         Call MniExpiryDamageInvoice_Click
      Case "MniFinishedProduct"
         Call MniFinishedProduct_Click
      Case "MniFormulaInfo"
         Call MniFormulaInfo_Click
      Case "MniGatePassIn"
         Call MniGatePassIn_Click
      Case "MniGatePassOut"
         Call MniGatePassOut_Click
      Case "MniGroups"
         Call MniGroups_Click
      Case "MniHoliday"
         Call MniHoliday_Click
      Case "MniHotandColdMembers"
'         Call MniHotandColdMembers_Click
      Case "MniHotandColdProducts"
'         Call MniHotandColdProducts_Click
      Case "MniInstallmentCustomers"
         Call MniInstallmentCustomers_Click
      Case "MniJournalVoucher"
         Call MniJournalVoucher_Click
      Case "MniLedger"
'         Call MniLedger_Click
      Case "MniLedgerDetail"
'         Call MniLedgerDetail_Click
      Case "MniLiftInvoice"
         Call MniLiftInvoice_Click
      Case "MniLoans"
         Call MniLoans_Click
      Case "MniLockAccounts"
         Call MniLockAccounts_Click
      Case "MniManufacturedProduct"
'         Call MniManufacturedProduct_Click
      Case "MniManufacturedProducts"
         Call MniManufacturedProducts_Click
      Case "MniManufacturedReturn"
         Call MniManufacturedReturn_Click
      Case "MniMemberList"
'         Call MniMemberList_Click
      Case "MniMemberTypes"
         Call MniMemberTypes_Click
      Case "MniMembers"
         Call MniMembers_Click
      Case "MniMembersDiscount"
         Call MniMembersDiscount_Click
      Case "MniMeterReadingRegister"
'         Call MniMeterReadingRegister_Click
      Case "MniMeterReadings"
         Call MniMeterReadings_Click
      Case "MniMonthlyIncomeExpense"
         Call MniMonthlyIncomeExpense_Click
      Case "MniMultipleBarCodePrinting"
'         Call MniMultipleBarCodePrinting_Click
      Case "MniOpeningAccountsOrganization"
         Call MniOpeningAccountsOrganization_Click
      Case "MniOpeningProduct"
         Call MniOpeningProduct_Click
      Case "MniOpeningProductVerification"
         Call MniOpeningProductVerification_Click
      Case "MniOpeningStock"
         Call MniOpeningStock_Click
      Case "MniOrganization"
         Call MniOrganization_Click
      Case "MniPLSetting"
         Call MniPLSetting_Click
      Case "MniProfit"
'         Call MniProfit_Click
      Case "MniPacking"
         Call MniPacking_Click
      Case "MniPaymentInvoiceWise"
         Call MniPaymentInvoiceWise_Click
      Case "MniPaymentVender"
         Call MniPaymentVender_Click
      Case "MniPettyCash"
         Call MniPettyCash_Click
      Case "MniPettyCashVerification"
         Call MniPettyCashVerification_Click
      Case "MniProductAnalysisReport"
'         Call MniProductAnalysisReport_Click
      Case "MniProductAnalysisStoreValueReport"
'         Call MniProductAnalysisStoreValueReport_Click
      Case "MniProductDifference"
'         Call MniProductDifference_Click
      Case "MniProductLedger"
'         Call MniProductLedger_Click
      Case "MniProductList"
'         Call MniProductList_Click
      Case "MniProductNotIncludedList"
'         Call MniProductNotIncludedList_Click
      Case "MniProductOffer"
         Call MniProductOffer_Click
      Case "MniProductPriceList"
'         Call MniProductPriceList_Click
      Case "MniProductStockSummary"
'         Call MniProductStockSummary_Click
      Case "MniProductStockSummaryStoreWise"
'         Call MniProductStockSummaryStoreWise_Click
      Case "MniProductStockValue"
'         Call MniProductStockValue_Click
      Case "MniProductWiseSaleStoreWise"
'         Call MniProductWiseSaleStoreWise_Click
      Case "MniProductionIN"
         Call MniProductionIN_Click
      Case "MniProductionOut"
         Call MniProductionOut_Click
      Case "MniProductionRegister"
'         Call MniProductionRegister_Click
      Case "MniProducts"
         Call MniProducts_Click
      Case "MniProductsNotInPurchase"
'         Call MniProductsNotInPurchase_Click
      Case "MniProductsUsed"
'         Call MniProductsUsed_Click
      Case "MniProfitRegister"
'         Call MniProfitRegister_Click
      Case "MniPurchaseInvoice"
         Call MniPurchaseInvoice_Click
      Case "MniPurchaseOrder"
         Call MniPurchaseOrder_Click
      Case "MniPurchaseRegister"
'         Call MniPurchaseRegister_Click
      Case "MniPurchaseRegisterSerailWise"
'         Call MniPurchaseRegisterSerailWise_Click
      Case "MniPurchaseReturnInvoice"
         Call MniPurchaseReturnInvoice_Click
      Case "MniRecoveryCustomer"
         Call MniRecoveryCustomer_Click
      Case "MniRecoveryInvoiceWise"
         Call MniRecoveryInvoiceWise_Click
      Case "MniRecoverySheet"
'         Call MniRecoverySheet_Click
      Case "MniReferences"
         Call MniReferences_Click
      Case "MniReplacementInvoice"
         Call MniReplacementInvoice_Click
      Case "MniRestore"
         Call MniRestore_Click
      Case "MniSalary"
         Call MniSalary_Click
      Case "MniSalaryEmployeeWise"
'         Call MniSalaryEmployeeWise_Click
      Case "MniSaleCounters"
         Call MniSaleCounters_Click
      Case "MniSaleExpenseStoreWise"
'         Call MniSaleExpenseStoreWise_Click
      Case "MniSaleInvoice"
         Call MniSaleInvoice_Click
      Case "MniSaleInvoicePOS"
         Call MniSaleInvoicePOS_Click
      Case "MniSaleOrder"
         Call MniSaleOrder_Click
      Case "MniSaleOrderPOS"
         Call MniSaleOrderPOS_Click
      Case "MniSaleRegister"
'         Call MniSaleRegister_Click
      Case "MniSaleRegisterSerailWise"
'         Call MniSaleRegisterSerailWise_Click
      Case "MniSaleReturnInvoice"
         Call MniSaleReturnInvoice_Click
      Case "MniSaleReturnInvoicePOS"
         Call MniSaleReturnInvoicePOS_Click
      Case "MniSectors"
         Call MniSectors_Click
      Case "MniServiceInvoice"
         Call MniServiceInvoice_Click
      Case "MniServiceProducts"
         Call MniServiceProducts_Click
      Case "MniShifts"
         Call MniShifts_Click
      Case "MniSkinSelection"
         Call MniSkinSelection_Click
      Case "MniSleepingMembers"
'         Call MniSleepingMembers_Click
      Case "MniSoftwareDefaultSettings"
         Call MniSoftwareDefaultSettings_Click
      Case "MniStockAdjustment"
         Call MniStockAdjustment_Click
      Case "MniStockAdjustmentValueReport"
'         Call MniStockAdjustmentValueReport_Click
      Case "MniStockLimitStoreWise"
         Call MniStockLimitStoreWise_Click
'      Case "MniStockPendingTransfer"
'         Call MniStockPendingTransfer_Click
      Case "MniStockRegister"
'         Call MniStockRegister_Click
      Case "MniStoreWiseStockTransferDetail"
'         Call MniStoreWiseStockTransferDetail_Click
      Case "MniStockTransferInvoice"
         Call MniStockTransferInvoice_Click
      Case "MniStockValueRegister"
'         Call MniStockValueRegister_Click
      Case "MniStockWastageInvoice"
         Call MniStockWastageInvoice_Click
      Case "MniStoreList"
'         Call MniStoreList_Click
      Case "MniStores"
         Call MniStores_Click
      Case "MniSubGroup"
         Call MniSubGroup_Click
      Case "MniSubGroupList"
'         Call MniSubGroupList_Click
      Case "MniTables"
         Call MniTables_Click
      Case "MniTrialBalance"
'         Call MniTrialBalance_Click
      Case "MniPackageDealInfo"
         Call MniPackageDealInfo_Click
      Case "MniUnits"
         Call MniUnits_Click
      Case "MniUserDefaultSettings"
         Call MniUserDefaultSettings_Click
      Case "MniUsers"
         Call MniUsers_Click
      Case "MniVenderPurchaseBills"
'         Call MniVenderPurchaseBills_Click
      Case "MniVendors"
         Call MniVendors_Click
      Case "MniVendorList"
'         Call MniVendorList_Click
      Case "MniZones"
         Call MniZones_Click
   End Select
End Sub

Private Sub Timer2_Timer()
   On Error GoTo ErrorHandler
   Timer2.Enabled = FunSecurityCheck
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Sub Demo()
101:
     X = InputBoxDK("Enter your Password.", "Password Required")
If StrPtr(X) = 0 Then
  'Cancel pressed
   Exit Sub
ElseIf X = "" Then
   MsgBox "Please enter a password"
   GoTo 101:
Else
  'Ok pressed
  'Continue with your macro.
  'Password is stored in the variable "x"
End If
End Sub
