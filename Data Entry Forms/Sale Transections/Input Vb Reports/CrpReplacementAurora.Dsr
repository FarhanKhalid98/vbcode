VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} CrpReplacementAurora 
   ClientHeight    =   11355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15465
   OleObjectBlob   =   "CrpReplacementAurora.dsx":0000
End
Attribute VB_Name = "CrpReplacementAurora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ReportHeaderSection1_Format(ByVal pFormattingInfo As Object)
   If LogoPath = "" Then
      Picture1.Suppress = True
   Else
      Picture1.SetOleLocation LogoPath
      Picture1.Suppress = False
   End If
End Sub

