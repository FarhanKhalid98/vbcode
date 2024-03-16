VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} CrptSaleInvoice 
   ClientHeight    =   13215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14835
   OleObjectBlob   =   "CrptSaleInvoice.dsx":0000
End
Attribute VB_Name = "CrptSaleInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PageHeaderSection1_Format(ByVal pFormattingInfo As Object)
   Text10.Suppress = Not ObjRegistry.ShowCode
   Text32.Suppress = Not ObjRegistry.ShowSerial
End Sub

Private Sub ReportHeaderSection1_Format(ByVal pFormattingInfo As Object)
  If LogoPath = "" Then
      Address1.Left = 0
      Address1.Width = Address1.Width + Picture1.Width
      PhonNo1.Left = 0
      PhonNo1.Width = PhonNo1.Width + Picture1.Width
      CoompanyName1.Left = 0
      CoompanyName1.Width = CoompanyName1.Width + Picture1.Width
      Picture1.Suppress = True
   Else
      Picture1.SetOleLocation LogoPath
      Picture1.Suppress = False
   End If
End Sub

Private Sub DetailSection1_Format(ByVal pFormattingInfo As Object)
   code1.Suppress = Not ObjRegistry.ShowCode
   RecordNumber1.Suppress = Not ObjRegistry.ShowSerial
End Sub

