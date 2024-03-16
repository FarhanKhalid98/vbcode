VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptTrialBalance 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "RptTrialBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkDetailedTrial 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Detailed Trial Balance"
      Height          =   255
      Left            =   4455
      TabIndex        =   2
      Top             =   4590
      Value           =   1  'Checked
      Width           =   3285
   End
   Begin VB.CheckBox ChkExclude 
      BackColor       =   &H00B98A03&
      Caption         =   "Exclude Accounts Having Zero Balance."
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   7260
      Visible         =   0   'False
      Width           =   3285
   End
   Begin JeweledBut.JeweledButton CmdPreview 
      Height          =   420
      Left            =   3900
      TabIndex        =   3
      Top             =   5378
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Preview"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "RptTrialBalance.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   5250
      TabIndex        =   4
      Top             =   5378
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "RptTrialBalance.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdClose 
      Height          =   420
      Left            =   6600
      TabIndex        =   5
      Top             =   5385
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "RptTrialBalance.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   4455
      TabIndex        =   0
      Top             =   3885
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   48562179
      CurrentDate     =   38244
   End
   Begin MSComCtl2.DTPicker DtpTo 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6045
      TabIndex        =   1
      Top             =   3885
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   48562179
      CurrentDate     =   38244
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trial Balance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Top             =   180
      Width           =   1845
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      Height          =   225
      Left            =   4455
      TabIndex        =   8
      Top             =   3660
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   225
      Left            =   6045
      TabIndex        =   7
      Top             =   3660
      Width           =   1095
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "RptTrialBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Public ShowDetailed As Boolean
Dim vStrSQL As String

Private Sub Check1_Click()

End Sub

Private Sub ChkDetailedTrial_Click()
   ShowDetailed = Abs(ChkDetailedTrial.Value)
End Sub

Private Sub CmdClose_Click()
   Unload Me
End Sub

Private Sub CmdPreview_Click()
  On Error GoTo ErrorHandler
  If FunRefreshData = False Then Exit Sub
  If Rs.RecordCount = 0 Then
    MsgBox "No record found", vbInformation, "Information"
    Exit Sub
  Else
    Call SetCrystalReport
    RptReportViewer.Show vbModal, Me
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub CmdPrint_Click()
  On Error GoTo ErrorHandler
  If FunRefreshData = False Then Exit Sub
  If Rs.RecordCount = 0 Then
    MsgBox "No record found", vbInformation, "Information"
    Exit Sub
  Else
    Call SetCrystalReport
    RptReportViewer.Report.PrintOut
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event 9, 1, 1, 1
    KeyCode = 0
  End If
End Sub

Private Function FunRefreshData() As Boolean
  On Error GoTo ErrorHandler
  Dim vSQL As String, vwhere  As String
  ''If ChkExclude.Value = 1 Then
  vwhere = " and (AccountsBalances.Debit > 0 OR AccountsBalances.Credit > 0 or AccountsBalances.OpeningBal > 0 or AccountsBalances.Bal > 0)"
  ''End If
  CN.Execute "EXECUTE SPAccountsBalances '" & DtpFrom.Value & "','" & DtpTo.Value & "'"
  vSQL = "SELECT ChartOfAccounts.AccountNo, ChartOfAccounts.AccountName, AccountsBalances.OpeningDebit, AccountsBalances.OpeningCredit, " & _
    " AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, AccountsBalances.Debit, AccountsBalances.Credit, AccountsBalances.Bal, " & _
    " AccountsBalances.BalType FROM AccountsBalances INNER JOIN ChartOfAccounts ON AccountsBalances.AccountNo = ChartOfAccounts.AccountNo where 1=1 " & vwhere & " order by ChartOfAccounts.AccountNo"
  Set Rs = CN.Execute(vSQL)
  FunRefreshData = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
  FunRefreshData = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
   If ShowDetailed Then
     Set RptReportViewer.Report = New CrpTrialBalanceDetail
   Else
     Set RptReportViewer.Report = New CrpTrialBalanceSummary
   End If
   'RptReportViewer.Report.TxtCompanyName.SetText CN.Execute("select companyname from Project_Registry").Fields(0).Value
   RptReportViewer.Report.ReportTitle = "Trial Balance"
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue "From : " & Format(DtpFrom.Value, "dd/MM/yyyy") & ",   To : " & Format(DtpTo.Value, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ChkExclude.Value = 1, "Excluding", "Including") & " accounts having zero balance"
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(6).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.PaperOrientation = IIf(ShowDetailed, crLandscape, crPortrait)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Trial Balance"
   DtpFrom.Value = Date - 30
   DtpTo.Value = Date
   ShowDetailed = Abs(ChkDetailedTrial.Value)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
