VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} CryRptPurchaseReturnInvoice 
   ClientHeight    =   9495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11310
   OleObjectBlob   =   "CryRptPurchaseReturnInvoice.dsx":0000
End
Attribute VB_Name = "CryRptPurchaseReturnInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DetailSection1_Format(ByVal pFormattingInfo As Object)
   RetailPrice1.Suppress = Not ObjRegistry.ShowRetailinPurchaseReturnPrint
End Sub

Private Sub PageHeaderSection1_Format(ByVal pFormattingInfo As Object)
   Text16.Suppress = Not ObjRegistry.ShowRetailinPurchaseReturnPrint
End Sub
