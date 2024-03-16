VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RptLedger12 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RptLedger.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptToday 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBD0AB&
      Caption         =   "OptSummary"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3540
      TabIndex        =   3
      Top             =   3975
      Value           =   -1  'True
      Width           =   210
   End
   Begin VB.OptionButton OptAllDates 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBD0AB&
      Caption         =   "OptSummary"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3540
      TabIndex        =   0
      Top             =   2445
      Width           =   210
   End
   Begin VB.OptionButton OptFromToDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBD0AB&
      Caption         =   "OptSummary"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3540
      TabIndex        =   1
      Top             =   2955
      Width           =   210
   End
   Begin VB.OptionButton OptADay 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBD0AB&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3540
      TabIndex        =   2
      Top             =   3465
      Width           =   210
   End
   Begin JeweledBut.JeweledButton BtnAccount 
      Height          =   330
      Left            =   4710
      TabIndex        =   11
      Top             =   4965
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      TX              =   "..."
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
      MICON           =   "RptLedger.frx":7796
      BC              =   14737632
      FC              =   0
   End
   Begin VB.TextBox TxtAccountName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   5160
      TabIndex        =   10
      Top             =   4965
      Width           =   4275
   End
   Begin VB.TextBox TxtAccountNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3540
      TabIndex        =   4
      Top             =   4965
      Width           =   1140
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6750
      TabIndex        =   7
      Top             =   6390
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Close"
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
      MICON           =   "RptLedger.frx":77B2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3975
      TabIndex        =   5
      Top             =   6390
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Pre&view"
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
      MICON           =   "RptLedger.frx":77CE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5355
      TabIndex        =   6
      Top             =   6390
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Print"
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
      MICON           =   "RptLedger.frx":77EA
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpADay 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6045
      TabIndex        =   15
      Top             =   3465
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   45023235
      CurrentDate     =   38244
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6045
      TabIndex        =   17
      Top             =   2955
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   45023235
      CurrentDate     =   38244
   End
   Begin MSComCtl2.DTPicker DtpTo 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7620
      TabIndex        =   18
      Top             =   2955
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   45023235
      CurrentDate     =   38244
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   195
      Left            =   7620
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Date"
      Height          =   195
      Left            =   6045
      TabIndex        =   20
      Top             =   3285
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   195
      Left            =   6045
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date To Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3915
      TabIndex        =   16
      Top             =   2955
      Width           =   1785
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Today"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3915
      TabIndex        =   14
      Top             =   3975
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Dates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3915
      TabIndex        =   13
      Top             =   2445
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3915
      TabIndex        =   12
      Top             =   3465
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5160
      TabIndex        =   9
      Top             =   4725
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3540
      TabIndex        =   8
      Top             =   4725
      Width           =   1080
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
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
Attribute VB_Name = "RptLedger12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql As String, vDate As String
Dim RsReport As New ADODB.Recordset
Dim vStrComp As String, vCompanyName As String, vAddress As String, vemail As String

Private Sub BtnAccount_Click()
If FunSelectAccount(ssButton, False) = True Then
      TxtAccountNo.SetFocus
   Else
      TxtAccountNo.SetFocus
   End If
End Sub

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    '-- when Account No is written then it will check and all its related value will be write its appropriate places
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
        TxtAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    VStrSQL = "Select AccountNo,AccountName from chartofaccounts where AccountNo='" & SchAccounts.ParaOutAccountNo & "'"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtaccountName.Text = !AccountName
          FunSelectAccount = True
          .Close
          'If BtnSave.Visible = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectAccount = False
          .Close
          TxtAccountNo.Text = ""
          TxtaccountName.Text = ""
          'If BtnSave.Visible = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnClose_Click()
    Unload Me
End Sub

Private Sub BtnPreview_Click()
    If SetReport Then
        RptReportViewer.Caption = "Ledger Report"
        RptReportViewer.Show
    End If
End Sub

Private Sub BtnPrint_Click()
    If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub DtpADay_Change()
   vDate = " and VoucherDate = '" & DtpADay.Value & "'"
End Sub

Private Sub DtpFrom_Change()
   vDate = " and VoucherDate BETWEEN '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'"
End Sub

Private Sub DtpTo_Change()
   vDate = " and VoucherDate BETWEEN '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'"
End Sub

Private Sub Form_Load()
    'sSql = "select companyid,companyname from companies"
    'Rs.Open sSql, CN, adOpenStatic, adLockReadOnly
    'Call OptDetail_Click
    DtpADay.Value = Date
    DtpTo.Value = Date
    DtpFrom.Value = Date - 30
    OptAllDates.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Rs.Close
    'Set Rs = Nothing
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function SetReport() As Boolean
   On Error GoTo ErrorHandler
   SetReport = False
   If Trim(TxtAccountNo.Text) = "" Then
      MsgBox "Invalid account No", vbInformation + vbOKOnly, "Alert"
      TxtAccountNo.SetFocus
      Exit Function
   End If
   vStrComp = "Select CompanyName,Address,City,PhoneNo,email from Company"
    sSql = " Select av.*, accountname" & _
           " from vuallvouchers av inner join chartofaccounts ca on ca.accountno = av.accountno" & _
           " where av.AccountNo = '" & TxtAccountNo.Text & "'" & vDate & " ORDER BY VoucherDate"
    Me.MousePointer = vbHourglass
    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open sSql, CN, adOpenStatic, adLockReadOnly
    If RsReport.BOF Then
        MsgBox "No record exists.", vbInformation, Me.Caption
        Me.MousePointer = vbDefault
        Exit Function
    End If
    Set RptReportViewer.Report = New CrpLedger
    RptReportViewer.Report.Database.SetDataSource RsReport
    vStrComp = "Select CompanyName,Address,City,PhoneNo,email from Company"
    With CN.Execute(vStrComp)
      If .RecordCount > 0 Then
         vCompanyName = !CompanyName
         vAddress = !Address & IIf(IsNull(!City), "", ", " & !City) & IIf(IsNull(!PhoneNo), "", ". Phone # " & !PhoneNo) & IIf(IsNull(!email), "", vbCrLf & !email)
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue vAddress
      End If
   .Close
   End With
    SetReport = True
    Me.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub Label1_Click()
   OptFromToDate.Value = True
   Call OptFromToDate_Click
End Sub

Private Sub Label2_Click()
   OptADay.Value = True
   Call OptADay_Click
End Sub

Private Sub Label4_Click()
   OptAllDates.Value = True
   Call OptAllDates_Click
End Sub

Private Sub Label5_Click()
   OptToday.Value = True
   Call OptToday_Click
End Sub

Private Sub OptADay_Click()
   If Label7.Visible = True Then Label7.Visible = False
   If DtpFrom.Visible = True Then DtpFrom.Visible = False
   If DtpTo.Visible = True Then DtpTo.Visible = False
   Label8.Visible = True
   DtpADay.Visible = True
   vDate = " and VoucherDate = '" & DtpADay.Value & "'"
End Sub

Private Sub OptAllDates_Click()
   If Label7.Visible = True Then Label7.Visible = False
   If DtpFrom.Visible = True Then DtpFrom.Visible = False
   If DtpTo.Visible = True Then DtpTo.Visible = False
   If Label8.Visible = True Then Label8.Visible = False
   If DtpADay.Visible = True Then DtpADay.Visible = False
   vDate = ""
End Sub

Private Sub OptFromToDate_Click()
   If Label8.Visible = True Then Label8.Visible = False
   If DtpADay.Visible = True Then DtpADay.Visible = False
   Label7.Visible = True
   DtpFrom.Visible = True
   DtpTo.Visible = True
   vDate = " and VoucherDate BETWEEN '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'"
End Sub

Private Sub OptToday_Click()
   If Label7.Visible = True Then Label7.Visible = False
   If DtpFrom.Visible = True Then DtpFrom.Visible = False
   If DtpTo.Visible = True Then DtpTo.Visible = False
   If Label8.Visible = True Then Label8.Visible = False
   If DtpADay.Visible = True Then DtpADay.Visible = False
   vDate = " and VoucherDate = '" & Date & "'"
End Sub

Private Sub TxtAccountNo_Change()
   If ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   If TxtaccountName.Text <> "" Then
      TxtaccountName.Text = ""
   End If
End Sub

Private Sub TxtAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> "TxtAccountNo" Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtaccountName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
