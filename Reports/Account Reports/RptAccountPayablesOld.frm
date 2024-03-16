VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RptAccountPayablesOld 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "RptAccountPayablesOld.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1000
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton CmdPreview 
      Height          =   420
      Left            =   4028
      TabIndex        =   3
      Top             =   5438
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   741
      TX              =   "Preview"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "RptAccountPayablesOld.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   5378
      TabIndex        =   4
      Top             =   5438
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   741
      TX              =   "Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "RptAccountPayablesOld.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdClose 
      Height          =   420
      Left            =   6698
      TabIndex        =   5
      Top             =   5438
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   741
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "RptAccountPayablesOld.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   4028
      TabIndex        =   0
      Top             =   3368
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   572
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   48103427
      CurrentDate     =   38244
   End
   Begin MSComCtl2.DTPicker DtpTo 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5798
      TabIndex        =   1
      Top             =   3368
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   572
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   48103427
      CurrentDate     =   38244
   End
   Begin SITextBox.Txt TxtAmountLimit 
      Height          =   315
      Left            =   4050
      TabIndex        =   2
      Top             =   4275
      Width           =   1950
      _ExtentX        =   3450
      _ExtentY        =   550
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Payable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
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
      Width           =   2370
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Limit"
      Height          =   225
      Left            =   4050
      TabIndex        =   8
      Top             =   4005
      Width           =   1095
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   225
      Left            =   5798
      TabIndex        =   7
      Top             =   3143
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      Height          =   225
      Left            =   4028
      TabIndex        =   6
      Top             =   3143
      Width           =   1095
   End
End
Attribute VB_Name = "RptAccountPayablesOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset

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
   Dim vSQL As String
  CN.Execute "EXECUTE SPAccountsBalances '" & DtpFrom.Value & "','" & DtpTo.Value & "'"
  vSQL = "SELECT ChartOfAccounts.AccountNo, ChartOfAccounts.AccountName+ ' ' + isnull(p.phone1,'') + ' ' + isnull(p.phone2,'') + ' ' + isnull(p.Mobile,'') as AccountName, AccountsBalances.OpeningDebit,AccountsBalances.OpeningCredit, " & vbCrLf & _
        " AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, AccountsBalances.Debit, AccountsBalances.Credit, AccountsBalances.Bal," & vbCrLf & _
        " AccountsBalances.BalType, p.city  FROM AccountsBalances INNER JOIN ChartOfAccounts ON  AccountsBalances.AccountNo = ChartOfAccounts.AccountNo " & vbCrLf & _
        " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf & _
        " Where (Bal * case when baltype = 'Cr' then -1 else 1 end) " & IIf(Val(TxtAmountLimit.Text) = 0, " < 0 ", " between " & Val(TxtAmountLimit.Text) * -1 & " and -1 ") & vbCrLf & _
        " and AccountsBalances.accountno like '6%' and ChartOfAccounts.isdetailed =1 order by ChartOfAccounts.AccountNo"
  
   Set Rs = CN.Execute(vSQL)
   FunRefreshData = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
   FunRefreshData = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
   Set RptReportViewer.Report = New CrpAccountPayables
   'RptReportViewer.Report.TxtCompanyName.SetText CN.Execute("select companyname from Project_Registry").Fields(0).Value
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ReportTitle = "Accounts Payables"
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue "From : " & Format(DtpFrom.Value, "dd/MM/yyyy") & ",   To : " & Format(DtpTo.Value, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity) & IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.PaperOrientation = crPortrait
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
  ShowPicture Me
  AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
  SetWindowText Me.hWnd, "Accounts Payables"
  DtpFrom.Value = Date - 30
  DtpTo.Value = Date
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub
