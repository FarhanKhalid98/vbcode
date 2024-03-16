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

Public Sub InitializeClass(objCN As ADODB.Connection)
    Declarations.SetConnection objCN
End Sub

Public Sub CompaniesForm()
    DefCompanies.Show vbModal
End Sub

Public Sub GroupsForm()
    DefGroups.Show vbModal
End Sub

Public Sub VendorsForm()
    DefVendors.Show vbModal
End Sub

Public Sub ProductsForm()
    DefProducts.Show vbModal
End Sub

Public Sub SubGroupsForm()
    DefSubGroups.Show vbModal
End Sub

Public Sub OpeningStockForm()
    FrmOpeningStock.Show vbModal
End Sub
