VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptTrialBalanceNew 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptTrialBalanceNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbSortType 
      Height          =   315
      ItemData        =   "RptTrialBalanceNew.frx":0ECA
      Left            =   10271
      List            =   "RptTrialBalanceNew.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   6574
      Width           =   1275
   End
   Begin VB.ComboBox CmbSortName 
      Height          =   315
      ItemData        =   "RptTrialBalanceNew.frx":0ECE
      Left            =   8381
      List            =   "RptTrialBalanceNew.frx":0ED0
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   6574
      Width           =   1815
   End
   Begin VB.CheckBox ChkDetailedTrial 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Detailed Trial Balance"
      Height          =   255
      Left            =   4774
      TabIndex        =   7
      Top             =   6612
      Width           =   3285
   End
   Begin VB.CheckBox ChkOpening 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Include Opening"
      Height          =   255
      Left            =   4774
      TabIndex        =   8
      Top             =   6972
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.CheckBox ChkDetailed 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Account Detail"
      Height          =   255
      Left            =   4774
      TabIndex        =   2
      Top             =   4182
      Value           =   1  'Checked
      Width           =   3285
   End
   Begin VB.TextBox TxtAccountNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3814
      MaxLength       =   10
      TabIndex        =   1
      Top             =   3672
      Width           =   1020
   End
   Begin VB.TextBox TxtaccountName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5194
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3672
      Width           =   3825
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Account Level Limit"
      Height          =   1005
      Left            =   5554
      TabIndex        =   18
      Top             =   4632
      Width           =   1725
      Begin SITextBox.Txt TxtFrom 
         Height          =   315
         Left            =   270
         TabIndex        =   3
         Top             =   450
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         Appearance      =   0
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Masked          =   1
         IntegralPoint   =   1
      End
      Begin SITextBox.Txt TxtTo 
         Height          =   315
         Left            =   855
         TabIndex        =   4
         Top             =   450
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         Appearance      =   0
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Masked          =   1
         IntegralPoint   =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   270
         TabIndex        =   20
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   810
         TabIndex        =   19
         Top             =   225
         Width           =   240
      End
   End
   Begin VB.CheckBox ChkExclude 
      BackColor       =   &H00B98A03&
      Caption         =   "Exclude Accounts Having Zero Balance."
      Height          =   255
      Left            =   4774
      TabIndex        =   14
      Top             =   8524
      Visible         =   0   'False
      Width           =   3285
   End
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4841
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2967
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "..."
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
      MICON           =   "RptTrialBalanceNew.frx":0ED2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   3821
      TabIndex        =   0
      Top             =   2967
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtOrganizatonName 
      Height          =   315
      Left            =   5201
      TabIndex        =   13
      Tag             =   "nc"
      Top             =   2967
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JeweledBut.JeweledButton CmdPreview 
      Height          =   420
      Left            =   4470
      TabIndex        =   9
      Top             =   7590
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
      MICON           =   "RptTrialBalanceNew.frx":0EEE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   5786
      TabIndex        =   10
      Top             =   7587
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
      MICON           =   "RptTrialBalanceNew.frx":0F0A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdClose 
      Height          =   420
      Left            =   7136
      TabIndex        =   11
      Top             =   7587
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
      MICON           =   "RptTrialBalanceNew.frx":0F26
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   4781
      TabIndex        =   5
      Top             =   6162
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
      Left            =   6506
      TabIndex        =   6
      Top             =   6162
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
   Begin JeweledBut.JeweledButton BtnAccount 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   4841
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3672
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "RptTrialBalanceNew.frx":0F42
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Type"
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
      Left            =   10271
      TabIndex        =   30
      Top             =   6364
      Width           =   840
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Name"
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
      Left            =   8381
      TabIndex        =   29
      Top             =   6364
      Width           =   900
   End
   Begin VB.Label Label3 
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
      Height          =   225
      Left            =   6506
      TabIndex        =   26
      Top             =   5907
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      Height          =   225
      Left            =   4781
      TabIndex        =   25
      Top             =   5907
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3821
      TabIndex        =   24
      Top             =   3462
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5216
      TabIndex        =   23
      Top             =   3462
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
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
      Left            =   5201
      TabIndex        =   17
      Top             =   2742
      Width           =   1620
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
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
      Left            =   3821
      TabIndex        =   16
      Top             =   2742
      Width           =   1335
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
      Left            =   2700
      TabIndex        =   15
      Top             =   270
      Width           =   1845
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "RptTrialBalanceNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Public ShowDetailed As Boolean
Dim Application1 As New CRAXDRT.Application

Private Sub ChkDetailedTrial_Click()
   ShowDetailed = Abs(ChkDetailedTrial.Value)
End Sub

Private Sub BtnOrganization_Click()
   If FunSelectOrganizaton(ssButton, False) = True Then
      DtpFrom.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
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
  ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtOrganizationID.Name: If FunSelectOrganizaton(ssFunctionKey, True) = True Then TxtAccountNo.SetFocus
         Case TxtAccountNo.Name: If FunSelectAccount(ssFunctionKey, True) = True Then ChkDetailed.SetFocus
      End Select
  End If
End Sub

Private Function FunRefreshData() As Boolean
   On Error GoTo ErrorHandler
   Dim vSQL As String, vWhere  As String
   ''If ChkExclude.Value = 1 Then
   vWhere = " and (AccountsBalances.Debit > 0 OR AccountsBalances.Credit > 0 or AccountsBalances.OpeningBal > 0 or AccountsBalances.Bal > 0 )" & IIf(Trim(TxtOrganizationID.Text) = "", "", " And AccountsBalances.OrganizationID = " & TxtOrganizationID.Text) & IIf(TxtAccountNo.Text = "", "", " and ChartOfAccounts.AccountNo Like '" & TxtAccountNo.Text & "%'") & IIf(ChkDetailed.Value = 1, "", " and isDetailed = 0") & IIf(TxtFrom.Text = "", IIf(TxtTo.Text = "", "", " and AccountDepth Between " & Val(TxtFrom.Text) & " and " & Val(TxtTo.Text)), IIf(TxtTo.Text = "", " and AccountDepth >= " & Val(TxtFrom.Text), " and AccountDepth Between " & Val(TxtFrom.Text) & " and " & Val(TxtTo.Text)))
   ''End If
   vSQL = "EXECUTE SPAccountsBalancesNew '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & ChkOpening.Value
   CN.Execute vSQL
   'Calculate Average Cost
   CN.Execute "exec SPAverageCost '" & DtpTo.DateValue & "'"
   'Second Insert Closing Stock
   CN.Execute "EXECUTE SPClosingStockNew '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"

   vSQL = "SELECT AccountsBalances.OrganizationID, OrganizationName, cast(ChartOfAccounts.AccountNo as varchar(10)) as AccountNo, ChartOfAccounts.AccountName  +  isnull(' ('+p.city + ')','') as AccountName, AccountsBalances.OpeningDebit, AccountsBalances.OpeningCredit, " & vbCrLf & _
     " AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, AccountsBalances.Debit, AccountsBalances.Credit, AccountsBalances.Bal, " & vbCrLf & _
     " AccountsBalances.BalType, ChartOfAccounts.isDetailed, ChartOfAccounts.AccountDepth FROM AccountsBalances INNER JOIN ChartOfAccounts ON AccountsBalances.AccountNo = ChartOfAccounts.AccountNo  " & vbCrLf & _
     " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf & _
     " left Outer Join Organizations O On O.OrganizationID = AccountsBalances.OrganizationID where 1=1 " & vWhere & " order by " & CmbSortName.Text & " " & CmbSortType.Text
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
'     Set RptReportViewer.Report = New CrpTrialBalanceDetail
     Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\AccountReports\CrpTrialBalanceDetail.rpt")
   Else
'     Set RptReportViewer.Report = New CrpTrialBalanceSummary
     Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\AccountReports\CrpTrialBalanceSummary.rpt")
   End If
   'RptReportViewer.Report.TxtCompanyName.SetText CN.Execute("select companyname from Project_Registry").Fields(0).Value
   RptReportViewer.Report.ReportTitle = "Trial Balance"
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue "From : " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & ",   To : " & Format(DtpTo.DateValue, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ChkExclude.Value = 1, "Excluding", "Including") & " accounts having zero balance"
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(6).AddCurrentValue ObjRegistry.DevelopedBy
'   RptReportViewer.Report.ParameterFields(7).AddCurrentValue Trim(TxtOrganizationID.Text)
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.PaperOrientation = IIf(ShowDetailed, crLandscape, crPortrait)
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Trial Balance"
   ChkDetailedTrial_Click
   CmbSortName.Clear
   CmbSortName.AddItem "AccountNo"
   CmbSortName.AddItem "AccountName"
   CmbSortType.Clear
   CmbSortType.AddItem "Asc"
   CmbSortType.AddItem "Desc"
   
   CmbSortName.ListIndex = 0
  
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtOrganizationID_Change()
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizatonName.Text <> "" Then TxtOrganizatonName.Text = ""
End Sub

Private Sub TxtOrganizationID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtOrganizationID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectOrganizaton(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganizaton(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectOrganizaton(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganizaton = False: Exit Function
       TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    If TxtOrganizationID.Text = "" Then FunSelectOrganizaton = False: Exit Function
    vStrSQL = " Select * FROM Organizations where OrganizationID='" & TxtOrganizationID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizatonName.Text = !OrganizationName
          FunSelectOrganizaton = True
          .Close
          Exit Function
      Else
          FunSelectOrganizaton = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizatonName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInDetail = "False"
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
        TxtAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    If Trim(TxtAccountNo.Text) = "" Then Exit Function
    vStrSQL = " Select AccountNo, AccountName FROM ChartofAccounts where AccountNo= '" & TxtAccountNo.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtaccountName.Text = !AccountName
          FunSelectAccount = True
          Exit Function
      Else
          FunSelectAccount = False
          TxtAccountNo.Text = ""
          TxtaccountName.Text = ""
      End If
      .Close
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtAccountNo_Change()
   If TxtAccountNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   If TxtaccountName.Text <> "" Then TxtaccountName.Text = ""
End Sub

Private Sub TxtAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtAccountNo.Text) = "" Then Exit Sub
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

Private Sub BtnAccount_Click()
   On Error GoTo ErrorHandler
   If FunSelectAccount(ssButton, True) = True Then
      ChkDetailed.SetFocus
   Else
      TxtAccountNo.SetFocus
   End If
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

