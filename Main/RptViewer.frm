VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form RptReportViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Previewing Report"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "RptViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RptViewer.frx":0ECA
   ScaleHeight     =   4.708
   ScaleMode       =   5  'Inch
   ScaleWidth      =   5.271
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   5325
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   6810
      lastProp        =   600
      _cx             =   12012
      _cy             =   9393
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "RptReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Report As New CRAXDDRT.Report
Option Explicit

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
    CRViewer1.Zoom 100
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Public Sub RemoveObjectColors()
    Dim X As Object
    Dim Y As CRAXDDRT.Section
    For Each Y In Report.Sections
        For Each X In Y.ReportObjects
            If Not UCase(X.Name) Like "NO*" Then
                If TypeOf X Is TextObject Or TypeOf X Is FieldObject Then
                    X.BackColor = &HFFFFFFFF
                ElseIf TypeOf X Is BoxObject Then
                    X.FillColor = vbWhite
                End If
            End If
        Next
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Report = Nothing
End Sub
