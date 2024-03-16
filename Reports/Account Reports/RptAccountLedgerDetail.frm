VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptAccountLedgerDetail 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptAccountLedgerDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7421
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3536
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7421
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3086
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.TextBox TxtAccountNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4384
      MaxLength       =   10
      TabIndex        =   1
      Top             =   5749
      Width           =   1020
   End
   Begin VB.TextBox TxtaccountName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5764
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5749
      Width           =   3585
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   4871
      TabIndex        =   4
      Top             =   7624
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
      MICON           =   "RptAccountLedgerDetail.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   6221
      TabIndex        =   5
      Top             =   7624
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
      MICON           =   "RptAccountLedgerDetail.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7556
      TabIndex        =   6
      Top             =   7624
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
      MICON           =   "RptAccountLedgerDetail.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSearch 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5404
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5749
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
      MICON           =   "RptAccountLedgerDetail.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5374
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4924
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
      MICON           =   "RptAccountLedgerDetail.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   4354
      TabIndex        =   0
      Top             =   4924
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   3
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
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   5734
      TabIndex        =   8
      Tag             =   "nc"
      Top             =   4924
      Width           =   3585
      _ExtentX        =   6324
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   5336
      TabIndex        =   2
      Top             =   6769
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
      Left            =   7061
      TabIndex        =   3
      Top             =   6769
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
   Begin VB.Label Label4 
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
      Left            =   5336
      TabIndex        =   17
      Top             =   6544
      Width           =   885
   End
   Begin VB.Label Label1 
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
      Left            =   7091
      TabIndex        =   16
      Top             =   6544
      Width           =   705
   End
   Begin VB.Label LblOrganizationID 
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
      Left            =   4354
      TabIndex        =   15
      Top             =   4699
      Width           =   1335
   End
   Begin VB.Label LblOrganizationName 
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
      Left            =   5734
      TabIndex        =   14
      Top             =   4699
      Width           =   1620
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ledger Detail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   13
      Top             =   270
      UseMnemonic     =   0   'False
      Width           =   1755
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11610
      Top             =   30
      Width           =   330
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
      Left            =   4384
      TabIndex        =   12
      Top             =   5539
      Width           =   1020
   End
   Begin VB.Label Label2 
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
      Left            =   5779
      TabIndex        =   11
      Top             =   5539
      Width           =   1335
   End
End
Attribute VB_Name = "RptAccountLedgerDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim Application1 As New CRAXDRT.Application
Dim vStrComp As String, vCompanyName As String, vAddress As String, vEmail As String, vStrSQL

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
        TxtAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    If Trim(TxtAccountNo.Text) = "" Then Exit Function
    
   vStrSQL = " Select c.AccountNo, c.AccountName FROM ChartofAccounts c " & vbCrLf & _
         " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
         " Left Outer join Members m on c.AccountNo = cast(m.Prefix as varchar(2))  + cast(m.MemberID as varchar(10)) " & vbCrLf & _
         " where p.BarCode = '" & (TxtAccountNo.Text) & "' or m.BarCode = '" & (TxtAccountNo.Text) & "' or (c.AccountNo = '" & (TxtAccountNo.Text) & "' and c.isDetailed = 1 and c.isLocked = 0)"

    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtAccountNo.Text = !AccountNo
          TxtaccountName.Text = !AccountName
          With CN.Execute("Select Address from Parties where PartyID = '" & TxtAccountNo.Text & "'")
            If .RecordCount > 0 Then
               TxtAddress.Text = IIf(IsNull(!Address) = True, "", !Address)
            Else
               TxtAddress.Text = ""
            End If
          End With
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

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnGroup_Click()
If FunSelectOrganization(ssButton, False) = True Then
      TxtaccountName.SetFocus
   Else
     TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnPreview_Click()
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

Private Sub BtnPrint_Click()
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
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, True) = True Then TxtAccountNo.SetFocus
         Case TxtAccountNo.Name: If FunSelectAccount(ssFunctionKey, True) = True Then DtpFrom.SetFocus
      End Select
   End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub Label17_Click()

End Sub

Private Sub TxtAccountNo_Change()
   If TxtAccountNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   If TxtaccountName.Text <> "" Then TxtaccountName.Text = ""
End Sub

Private Sub TxtAccountNo_Validate(Cancel As Boolean)
   Dim vTemp As Boolean
   If Trim(TxtAccountNo.Text) = "" Then Exit Sub
   If Trim(TxtaccountName.Text) <> "" Then Exit Sub
   vTemp = FunSelectAccount(ssValidate, False)
   If vTemp = False Then
      Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSearch_Click()
   On Error GoTo ErrorHandler
   If FunSelectAccount(ssButton, True) = True Then
      DtpFrom.SetFocus
   Else
      TxtAccountNo.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunRefreshData() As Boolean
  On Error GoTo ErrorHandler
  Dim vSQL As String
  vSQL = "EXECUTE SPAccountsLedgerDetailOrg " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & ",'" & TxtAccountNo.Text & "', '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
  Set Rs = CN.Execute("EXECUTE SPAccountsLedgerDetailOrg " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & ",'" & TxtAccountNo.Text & "', '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'")
  FunRefreshData = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
  FunRefreshData = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
   'Set RptReportViewer.Report = New CrpAccountLedgerDetail
   Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\AccountReports\CrpAccountLedgerDetail.rpt")
  'RptReportViewer.Report.TxtCompanyName.SetText CN.Execute("select companyname from Project_Registry").Fields(0).Value
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ReportTitle = "Account Ledger"
   
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue "Account : " & TxtAccountNo.Text & "/" & TxtaccountName.Text & IIf(TxtAddress.Text = "", "", " (" & TxtAddress.Text & ")") & IIf(TxtCity.Text = "", "", vbCrLf & TxtCity.Text & ".")
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue "From Date " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To " & Format(DtpTo.DateValue, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue Trim(TxtOrganizationID.Text)
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(6).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity & ".") & IIf(ObjRegistry.CompanyPhoneNo = "", "", " Phone # " & ObjRegistry.CompanyPhoneNo)

   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   
   RptReportViewer.Report.PaperOrientation = crLandscape
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Ledger Detail"
   
   TxtOrganizationID.Text = IIf(ObjRegistry.OrganizationID = "Null", "", ObjRegistry.OrganizationID)
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible

   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtOrganizationID_Change()
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizationName.Text <> "" Then TxtOrganizationName.Text = ""
End Sub

Private Sub TxtOrganizationID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtOrganizationID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectOrganization(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganization(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectOrganization(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchOrganization.Show vbModal, Me
       If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganization = False: Exit Function
      TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
   End If
   If TxtOrganizationID.Text = "" Then FunSelectOrganization = False: Exit Function
   vStrSQL = " Select * FROM Organizations where OrganizationID='" & TxtOrganizationID.Text & "'"
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtOrganizationName.Text = !OrganizationName
         FunSelectOrganization = True
         .Close
         Exit Function
      Else
         FunSelectOrganization = False
         .Close
         TxtOrganizationID.Text = ""
         TxtOrganizationName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
