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

Public Sub CashDepositForm()
    Dim a As New FrmBankCashDeposit
    a.Show
End Sub

Public Sub ChequeDepositForm()
    Dim a As New FrmBankChequeDeposit
    a.Show
End Sub

Public Sub ChequeIssuanceForm()
    Dim a As New FrmBankChequeIssuance
    a.Show
End Sub

Public Sub ChequeIssueReconciliationForm()
    Dim a As New FrmBankChequeIssueReconciliation
    a.Show
End Sub

Public Sub ChequeReceiveReconciliationForm()
    Dim a As New FrmBankChequeReceiveReconciliation
    a.Show
End Sub

Public Sub ChequeReceiveForm()
   Dim a As New FrmBankChequeReceive
   a.Show
End Sub
