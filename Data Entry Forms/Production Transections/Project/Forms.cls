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

Public Sub ManufacturedProductsForm()
    Dim a As New FrmManufacturedProducts
    a.Show
End Sub

Public Sub ManufacturedReturnForm()
    Dim a As New FrmManufacturedReturn
    a.Show
End Sub

Public Sub ProductionInForm()
    Dim a As New FrmProductionIn
    a.Show
End Sub

Public Sub ProductionOutForm()
    Dim a As New FrmProductionOut
    a.Show
End Sub

Public Sub GatePassInForm()
    Dim a As New FrmGatePassIn
    a.Show
End Sub

Public Sub GatePassOutForm()
    Dim a As New FrmGatePassOut
    a.Show
End Sub

