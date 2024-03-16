VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmMonthlyIncomeExpense 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   15360
   ClipControls    =   0   'False
   Icon            =   "FrmMonthlyIncomeExpense.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtIncomeID 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      CausesValidation=   0   'False
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   5040
      TabIndex        =   0
      Top             =   4275
      Width           =   1020
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7695
      TabIndex        =   2
      Top             =   6690
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMonthlyIncomeExpense.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6390
      TabIndex        =   3
      Top             =   6690
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMonthlyIncomeExpense.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10305
      TabIndex        =   7
      Top             =   6690
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMonthlyIncomeExpense.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5085
      TabIndex        =   4
      Top             =   6690
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMonthlyIncomeExpense.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9000
      TabIndex        =   5
      Top             =   6690
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMonthlyIncomeExpense.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   3780
      TabIndex        =   6
      Top             =   6690
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMonthlyIncomeExpense.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtRemarks 
      Height          =   315
      Left            =   5055
      TabIndex        =   1
      Top             =   5265
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   5
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   6405
      TabIndex        =   14
      Top             =   4275
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "/"
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpTo 
      Height          =   315
      Left            =   7935
      TabIndex        =   13
      Top             =   4275
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "/"
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6405
      TabIndex        =   12
      Top             =   4020
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Income ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   11
      Top             =   4020
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7935
      TabIndex        =   10
      Top             =   4020
      Width           =   705
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5055
      TabIndex        =   9
      Top             =   5025
      Width           =   750
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   8
      Top             =   270
      Width           =   1995
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove this Row"
      End
   End
End
Attribute VB_Name = "FrmMonthlyIncomeExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim RsBodyExpenditure As New ADODB.Recordset
Dim RsBodyNet As New ADODB.Recordset
Dim vCounter As Integer
Dim Flag As Boolean
Dim sSql As String
Dim vStrSQL As String
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vStrComp As String, vCompanyName As String, vAddress As String, vemail As String
'----------------------------------

Private Function FunRefreshData() As Boolean
On Error GoTo ErrorHandler
  Dim vSQL As String
  Me.MousePointer = vbHourglass
'  CN.Execute ("EXEC ProdRptNetExpenditure '" & DtpIncomeFrom.Value & "','" & DtpIncome.Value & "'")
 
  vSQL = "Select NetName, NetIncome From NetBody B Inner Join NetHeader H on H.NetID = b.NetID where IncomeID = " & Val(TxtIncomeID.Text)
    Set RsBodyNet = CN.Execute(vSQL)
    
vSQL = "Select AccountName as ExpendName, sum(Amount) as ExpendIncome" & vbCrLf _
   + " From DebitVouchersBody b inner join DebitVouchers h on h.voucherno = b.voucherno" & vbCrLf _
   + " inner join ChartofAccounts c on c.AccountNo = b.AccountNo " & vbCrLf _
   + " where ( b.AccountNo like '5%' or b.AccountNo like '63%' ) and h.voucherdate between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "'" & vbCrLf _
   + " Group By AccountName " & vbCrLf _
   + " union all " & vbCrLf _
   + " Select 'Commision Paid To Bank (Bank Card)' , Sum((TotAmount- isnull(BillDisc,0))*commision/100)" & vbCrLf _
   + " from saleheader h inner join" & vbCrLf _
   + " (select b.billid, b.billdate, sum(amount) TotAmount from salebody b" & vbCrLf _
   + " group by b.billid, b.billdate)b" & vbCrLf _
   + " on h.billid = b.billid and h.billdate = b.billdate" & vbCrLf _
   + " where b.billdate between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "' and BankCard = 1"

 ' vSQL = "Select ExpendName, ExpendIncome from ExpenditureDetail B Inner Join ExpendHeader H on H.ExpendID = B.ExpendID where 1=1"
  Set RsBodyExpenditure = CN.Execute(vSQL)
  
  FunRefreshData = True
  Me.MousePointer = vbDefault
  Exit Function
ErrorHandler:
  Me.MousePointer = vbDefault
  FunRefreshData = False
  Call ShowErrorMessage
End Function

Private Sub SetCrystalReport1()
  On Error GoTo ErrorHandler
  Me.MousePointer = vbHourglass
'  Set RptReportViewer.Report = New CrptNetExpenditure
  RptReportViewer.Report.ReportTitle = ""
  'RptReportViewer.Report.Database.SetDataSource RsBodyExpenditure, 3, 1
 Dim vDevice As String, vDriver As String, vPort As String
   vStrSQL = "Select * from Registry"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         vDevice = IIf(IsNull(!DeviceName), "Abc", !DeviceName)
         vDriver = IIf(IsNull(!DriverName), "Xyz", !DriverName)
         vPort = IIf(IsNull(!Port), "LPT1", !Port)
         RptReportViewer.Report.SelectPrinter vDriver, vDevice, vPort
      End If
   End With
  RptReportViewer.Report.PaperOrientation = crPortrait
  Me.MousePointer = vbDefault
  Exit Sub
ErrorHandler:
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

Private Sub SetCrystalReport()
  On Error GoTo ErrorHandler
  Me.MousePointer = vbHourglass
'  Set RptReportViewer.Report = New CrptNetExpenditure
  'this code works through the RDC object model to identify a subreport object
  'in the main report
Dim crSecs As CRAXDRT.Sections
Dim crSec As CRAXDRT.Section
Dim crRepObjs As CRAXDRT.ReportObjects
Dim crSubRepObj As CRAXDRT.SubreportObject
Dim crSubReport As CRAXDRT.Report
Dim i As Integer
Dim X As Integer
Set crSecs = RptReportViewer.Report.Sections
For i = 1 To crSecs.Count
  Set crSec = crSecs.Item(i)
  Set crRepObjs = crSec.ReportObjects
    For X = 1 To crRepObjs.Count
      If crRepObjs.Item(X).Kind = crSubreportObject Then
         'If X = 1 And i = 4 Then
            Set crSubReport = RptReportViewer.Report.OpenSubreport(crRepObjs.Item(X).SubreportName)
            'the following code sets the subreport table to a different database
            crSubReport.Database.SetDataSource RsBodyExpenditure, 3, 1
            'set the value for a text object in the header of the subreport
            'CRReport.Subreport1_Text2.SetText "This is the subreport"
            'within this loop you can set other properties of the subreport and
            'the field objects and sections in it.
         'ElseIf X = 1 And i = 5 Then
         '   Set crSubReport = RptReportViewer.Report.OpenSubreport(crRepObjs.Item(X).SubreportName)
         '   crSubReport.Database.SetDataSource Rs1, 3, 1
         'End If
      End If
    Next
Next
  'RptReportViewer.Report.TxtCompanyName.SetText ObjSupernetRegistry.CompanyName
  RptReportViewer.Report.ReportTitle = ""
  'RptReportViewer.Report.ParameterFields(1).AddCurrentValue "" 'Format(DtpTo.Value, "dd/MM/yyyy")
  RptReportViewer.Report.Database.SetDataSource RsBodyNet, 3, 1
   Dim vDevice As String, vDriver As String, vPort As String
   vStrSQL = "Select * from Registry"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         vDevice = IIf(IsNull(!DeviceName), "Abc", !DeviceName)
         vDriver = IIf(IsNull(!DriverName), "Xyz", !DriverName)
         vPort = IIf(IsNull(!Port), "LPT1", !Port)
         RptReportViewer.Report.SelectPrinter vDriver, vDevice, vPort
      End If
   End With
  RptReportViewer.Report.PaperOrientation = crPortrait
  Me.MousePointer = vbDefault
  Exit Sub
ErrorHandler:
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

Private Sub SetReportParameterField()
On Error GoTo ErrorHandler
    With CN.Execute("Select CompanyName,Address,City,PhoneNo,email from Company")
      If .RecordCount > 0 Then
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(IsNull(!CompanyName), "", CStr(!CompanyName))
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(IsNull(!Address), "", !Address) & IIf(IsNull(!City), "", ", " & !City & ".")
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(IsNull(!PhoneNo), "", CStr(!PhoneNo))
      End If
    .Close
    End With
    RptReportViewer.Report.ParameterFields(4).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
    RptReportViewer.Report.ParameterFields(5).AddCurrentValue "Date From " & Format(DtpFrom.DateValue, "dd-MMM-yyyy") & " To " & Format(DtpTo.DateValue, "dd-MMM-yyyy")
    RptReportViewer.Report.ReportTitle = "General Report"
Exit Sub
ErrorHandler:
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   BtnPrint.Enabled = False
   
  If FunRefreshData = False Then Exit Sub
  If RsBodyNet.RecordCount = 0 And RsBodyExpenditure.RecordCount = 0 Then
    MsgBox "No record found", vbInformation, "Information"
    Exit Sub
  ElseIf RsBodyNet.RecordCount = 0 Then
    Call SetCrystalReport1
    Call SetReportParameterField
    RptReportViewer.Show vbModal, Me
  Else
    Call SetCrystalReport
    Call SetReportParameterField
    RptReportViewer.Show vbModal, Me
  End If
  BtnPrint.Enabled = True
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
    BtnPrint.Enabled = True
End Sub

Private Sub BtnClear_Click()
  On Error GoTo ErrorHandler
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
  End If
  If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
  CN.BeginTrans
  CN.Execute "Delete from NetBody WHere IncomeID = " & Val(TxtIncomeID.Text)
  CN.Execute "Delete from IncomeHeader WHere IncomeID = " & Val(TxtIncomeID.Text)
  CN.CommitTrans
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  If CN.Errors.Count > 0 Then CN.RollbackTrans
  Call ShowErrorMessage
End Sub

'Private Sub GetNetExpenditure()
'   On Error GoTo ErrorHandler
'   sSql = "Select * from IncomeHeader where IncomeID = " & Val(TxtIncomeID.Text) & " And IncomeDate = '" & DtpIncome.Value & "'"
'   With CN.Execute(sSql)
'      If Not .BOF Then
'          DtpIncome.Value = !IncomeDate
'          TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
'      End If
'      .Close
'   End With
''   Call PopulateDataToGrid
'   FormStatus = OpenMode
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub

Private Sub BtnOpen_Click()
'   SchMonthlyIncomeExpense.Show vbModal
'   If SchMonthlyIncomeExpense.ParaOutID <> Empty Then
'      TxtIncomeID.Text = SchMonthlyIncomeExpense.ParaOutID
'      DtpIncome.Value = SchMonthlyIncomeExpense.ParaOutDate
'      DtpIncomeFrom.Day = 1
'      DtpIncomeFrom.Month = DtpIncome.Month
'      DtpIncomeFrom.Year = DtpIncome.Year
'      GetNetExpenditure
'   End If
'  Exit Sub
'ErrorHandler:
'  Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
'   If vIsNewRecord Then
'      If CN.Execute("Select * from IncomeHeader where IncomeID = " & Val(TxtIncomeID.Text) & " And IncomeDate = '" & DtpFrom.DateValue & "'").RecordCount > 0 Then
'         MsgBox "This voucher already exists. A new voucher No. has been generated. Please try again", vbCritical, "Alert"
'         TxtIncomeID.Text = FunGetMaxID
'         Exit Sub
'      End If
'   End If
'   RsBody.Filter = 0
'   If RsBody.RecordCount = 0 Then
'       MsgBox "Please enter at least one entry to save", vbExclamation, "Alert"
'       If TxtProductID.Visible And TxtProductID.Enabled Then TxtProductID.SetFocus
'       Exit Sub
'   End If
   
   
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
   CN.Execute "EXECUTE SPAccountsBalances '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
   
  'vSQL = "SELECT ChartOfAccounts.AccountNo, ChartOfAccounts.AccountName+ ' ' + isnull(p.phone1,'') + ' ' + isnull(p.phone2,'') + ' ' + isnull(p.Mobile,'') as AccountName, AccountsBalances.OpeningDebit,AccountsBalances.OpeningCredit, " & vbCrLf & _
        " AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, AccountsBalances.Debit, AccountsBalances.Credit, AccountsBalances.Bal," & vbCrLf & _
        " AccountsBalances.BalType, p.city  FROM AccountsBalances INNER JOIN ChartOfAccounts ON  AccountsBalances.AccountNo = ChartOfAccounts.AccountNo " & vbCrLf & _
        " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf & _
        " Where (Bal * case when baltype = 'Cr' then -1 else 1 end) " & IIf(Val(TxtAmountLimit.Text) = 0, " < 0 ", " between " & Val(TxtAmountLimit.Text) * -1 & " and -1 ") & vbCrLf & _
        " and AccountsBalances.accountno like '6%' and ChartOfAccounts.isdetailed =1 order by ChartOfAccounts.AccountNo "

   CN.Execute ("ProdNet '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & TxtIncomeID.Text & "," & DateDiff("d", DtpFrom.DateValue, DtpTo.DateValue))
   CN.BeginTrans
   sSql = "Select * From IncomeHeader Where IncomeID =" & Val(TxtIncomeID.Text) & " And IncomeDate = '" & DtpFrom.DateValue & "'"
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, CN, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !IncomeID = Val(TxtIncomeID.Text)
         !IncomeDate = DtpFrom.DateValue
         'CN.Execute ("ProdRptNetExpenditure '" & DtpIncomeFrom.Value & "','" & DtpIncome.Value & "'")
      End If
      !Remarks = IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text)
      !UserNo = vUser
      .Update
      .Close
   End With
   CN.CommitTrans
   Call BtnPrint_Click
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Property Get FormStatus() As FormMode
  'Nothing
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
  'Based upon the value of vNewValue, we shall decide what controls to enable/disable
  On Error GoTo ErrorHandler
  vMode = vNewValue
  Select Case vNewValue
    Case Is = NewMode
      Call SubClearFields
'      Call PopulateDataToGrid
      BtnPrint.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtIncomeID.Text = 1 'FunGetMaxID
      If DtpFrom.Enabled And DtpFrom.Visible Then DtpFrom.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      BtnPrint.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      DtpFrom.SetFocus
      vIsNewRecord = False
    Case Is = changeMode
      BtnPrint.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub DtpFrom_Change()
    If DtpFrom.Visible = False Then Exit Sub
    If Me.ActiveControl.Name <> DtpFrom.Name Then Exit Sub
    If DtpFrom.Enabled And DtpFrom.Visible Then FormStatus = changeMode
End Sub

Private Sub DtpTo_Change()
    If DtpTo.Visible = False Then Exit Sub
    If Me.ActiveControl.Name <> DtpTo.Name Then Exit Sub
    If DtpTo.Enabled And DtpTo.Visible Then FormStatus = changeMode
End Sub

'Private Sub DtpIncome_Click()
'    If DtpIncome.Visible = False Then Exit Sub
'    If Me.ActiveControl.Name <> DtpIncome.Name Then Exit Sub
'    DtpIncome.Day = 1
'    DtpIncomeFrom.Day = 1
'    DtpIncomeFrom.Month = DtpIncome.Month
'    DtpIncomeFrom.Year = DtpIncome.Year
'    DtpIncome.Day = DateDiff("d", DtpIncome.Value, DateAdd("M", 1, DtpIncomeFrom.Value))
'    TxtIncomeID.Text = FunGetMaxID
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
  ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
'         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtProductID.SetFocus
'         Case TxtProductID.Name: If FunSelectProduct(ssFunctionKey, False) = True Then TxtOverQty.SetFocus 'Else TxtProductID.SetFocus
      End Select
'  ElseIf KeyCode = vbKeyEscape And (Me.ActiveControl.Name = TxtProductID.Name Or Me.ActiveControl.Name = TxtProductName.Name Or Me.ActiveControl.Name = TxtProductName.Name Or Me.ActiveControl.Name = TxtUnderQty.Name Or Me.ActiveControl.Name = Grid.Name) Then
'    Call ClearDetailArea
  ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSave.Enabled Then BtnSave_Click
            KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
         Case vbKeyO
            If BtnOpen.Enabled Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
      End Select
  ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
'         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtProductName.SetFocus
      End Select
  End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   If BtnSave.Enabled Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then FormStatus = changeMode
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "General Report"
  'DtpIncome.Enabled = ObjUserSecurity.TaskAllowance("ChangeDateInCreditVoucher") Or ObjUserSecurity.IsAdministrator
  ' SetWindowText Me.hWnd, "Cash Received Vouchers"
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
  On Error GoTo ErrorHandler
  FunGetMaxID = CN.Execute("Select isnull(max(IncomeID),0) from IncomeHeader -- Where IncomeDate = '" & DtpFrom.DateValue & "'").Fields(0) + 1
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub SubClearFields()
  On Error GoTo ErrorHandler
  Dim ctl As Control
  For Each ctl In Me.Controls
    If TypeOf ctl Is TextBox Or TypeOf ctl Is SITextBox.txt Then
      ctl.Text = ""
    ElseIf TypeOf ctl Is ComboBox Then
    
    End If
  Next
  'DtpIncome.Value = Date
  'DtpIncomeFrom.Value = Date
  'DtpIncomeFrom.Day = 1
  'MsgBox DateDiff("d", Date, DateAdd("M", 1, Date))
  'DtpIncome.Day = DateDiff("d", DtpIncome.Value, DateAdd("M", 1, DtpIncomeFrom.Value))
  
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Are you sure to close without save?", vbQuestion + vbApplicationModal + vbYesNo, "Alert") = vbNo Then
         Cancel = 1
      End If
   Else
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set RsReport = Nothing
    Set FrmMonthlyIncomeExpense = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

