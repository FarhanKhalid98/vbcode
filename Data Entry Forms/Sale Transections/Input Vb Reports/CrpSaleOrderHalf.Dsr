VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} CrpSaleOrderHalf 
   ClientHeight    =   9495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   OleObjectBlob   =   "CrpSaleOrderHalf.dsx":0000
End
Attribute VB_Name = "CrpSaleOrderHalf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DetailSection1_Format(ByVal pFormattingInfo As Object)
   ProductID1.Suppress = Not ObjRegistry.ShowCode
   RecordNumber1.Suppress = Not ObjRegistry.ShowSerial
End Sub

Private Sub PageHeaderSection1_Format(ByVal pFormattingInfo As Object)
   If LogoPath = "" Then
      PrmAddress1.Left = 0
      PrmAddress1.Width = PrmAddress1.Width + Picture1.Width
      PrmPhone1.Left = 0
      PrmPhone1.Width = PrmPhone1.Width + Picture1.Width
      CompanyName1.Left = 0
      CompanyName1.Width = CompanyName1.Width + Picture1.Width
      Picture1.Suppress = True
   Else
      Picture1.SetOleLocation LogoPath
      Picture1.Suppress = False
   End If
End Sub

Private Sub PageHeaderSection2_Format(ByVal pFormattingInfo As Object)
   Text25.Suppress = Not ObjRegistry.ShowCode
   Text17.Suppress = Not ObjRegistry.ShowSerial
End Sub
