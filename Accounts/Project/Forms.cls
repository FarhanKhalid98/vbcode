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

Public Sub AccountsOpeningBalanceForm()
    DefAccountsOpeningBalance.Show
End Sub

Public Sub ChartOfAccountsForm()
    DefChartOfAccounts.Show
End Sub

Public Sub OrganizationalOpeningBalanceForm()
    FrmOpeningBalanceOrganizationWise.Show
End Sub

Public Sub LockAccountsForm()
    FrmLockAccounts.Show
End Sub

Public Sub CreditVoucherForm()
    Dim a As New FrmCreditVoucher
    a.Show
End Sub

Public Sub DebitVoucherForm()
   Dim a As New FrmDebitVoucher
   a.Show
End Sub

Public Sub JournalVoucherForm()
    Dim a As New FrmJournalVoucher
    a.Show
End Sub

Public Sub AdvancesForm()
    Dim a As New FrmAdvances
    a.Show
End Sub

Public Sub LoansForm()
    Dim a As New FrmLoans
    a.Show
End Sub

Public Sub PLSettingsForm()
   FrmPLSettingsOrganizationWise.Show
'    PLSettings.Show
End Sub
Public Sub PLSettingsWithoutOrgForm()
    PLSettings.Show
End Sub

Public Sub ExpenseSettingForm()
    ExpSettings.Show
End Sub

Public Sub AccountsDefaultSettingForm()
    FrmAccountsDefaultSetting.Show
End Sub
