VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptCashBook 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptCashBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkIncludeBV 
      BackColor       =   &H00FF8080&
      Caption         =   "Include Banks"
      Height          =   255
      Left            =   9855
      TabIndex        =   18
      Top             =   5850
      Value           =   1  'Checked
      Width           =   1380
   End
   Begin VB.CheckBox ChkIncludeJV 
      BackColor       =   &H00FF8080&
      Caption         =   "Include JV"
      Height          =   255
      Left            =   8640
      TabIndex        =   17
      Top             =   5850
      Value           =   1  'Checked
      Width           =   1200
   End
   Begin VB.CheckBox ChkSummary 
      BackColor       =   &H00FF8080&
      Caption         =   "Summary"
      Height          =   255
      Left            =   9450
      TabIndex        =   16
      Top             =   6210
      Width           =   1785
   End
   Begin VB.CheckBox ChkOnlyCashReceivedVoucher 
      BackColor       =   &H00FF8080&
      Caption         =   "Only Cash Received Voucher"
      Height          =   255
      Left            =   6743
      TabIndex        =   15
      Top             =   6210
      Width           =   2640
   End
   Begin VB.CheckBox ChkOpening 
      BackColor       =   &H00FF8080&
      Caption         =   "Include Opening"
      Height          =   255
      Left            =   6743
      TabIndex        =   4
      Top             =   5864
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin JeweledBut.JeweledButton CmdPreview 
      Height          =   420
      Left            =   5715
      TabIndex        =   5
      Top             =   7035
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
      MICON           =   "RptCashBook.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7065
      TabIndex        =   6
      Top             =   7035
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
      MICON           =   "RptCashBook.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdClose 
      Height          =   420
      Left            =   8370
      TabIndex        =   7
      Top             =   7035
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
      MICON           =   "RptCashBook.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   312
      Left            =   6216
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4436
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
      MICON           =   "RptCashBook.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   312
      Left            =   5196
      TabIndex        =   1
      Top             =   4436
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
      Height          =   312
      Left            =   6576
      TabIndex        =   10
      Tag             =   "nc"
      Top             =   4436
      Width           =   3588
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
      Height          =   312
      Left            =   6156
      TabIndex        =   2
      Top             =   5228
      Width           =   1308
      _Version        =   65543
      _ExtentX        =   2307
      _ExtentY        =   550
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
      Height          =   312
      Left            =   7908
      TabIndex        =   3
      Top             =   5228
      Width           =   1308
      _Version        =   65543
      _ExtentX        =   2307
      _ExtentY        =   550
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6195
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3465
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
      MICON           =   "RptCashBook.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   5175
      TabIndex        =   0
      Top             =   3465
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
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   6555
      TabIndex        =   20
      Tag             =   "nc"
      Top             =   3465
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
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
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
      Left            =   6555
      TabIndex        =   22
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
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
      Left            =   5175
      TabIndex        =   21
      Top             =   3240
      Width           =   720
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
      Height          =   192
      Left            =   5196
      TabIndex        =   14
      Top             =   4208
      Width           =   1284
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
      Height          =   192
      Left            =   6576
      TabIndex        =   13
      Top             =   4208
      Width           =   1596
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
      Height          =   192
      Left            =   7920
      TabIndex        =   12
      Top             =   5000
      Width           =   708
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
      Height          =   192
      Left            =   6156
      TabIndex        =   11
      Top             =   5000
      Width           =   888
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Book"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   0
      Left            =   2700
      TabIndex        =   8
      Top             =   270
      Width           =   1464
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11640
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "RptCashBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim Application1 As New CRAXDRT.Application
Dim vStrComp As String, vCompanyName As String, vAddress As String, vEmail As String, vStrSQL
Dim vPhone As String

Private Sub BtnGroup_Click()
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
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then TxtOrganizationID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganizaton(ssFunctionKey, True) = True Then DtpFrom.SetFocus
      End Select
  End If
End Sub

Private Function FunRefreshData() As Boolean
  On Error GoTo ErrorHandler
  Dim vSQL As String
  If ObjRegistry.ShowMultiBranches = False Then
   vSQL = "EXECUTE SPCashBook '" & CN.Execute("Select dbo.DefaultValue('CashInHand')").Fields(0).Value & "','" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text) & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", TxtStoreID.Text) & "," & ChkOpening.Value
  Else
   vSQL = "EXECUTE SPCashBookStoreWise '" & CN.Execute("Select dbo.DefaultValue('CashInHand')").Fields(0).Value & "','" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text) & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", TxtStoreID.Text) & "," & ChkOpening.Value
  End If
  CN.Execute vSQL
  vSQL = "Select * from CashBook where 1=1 " & IIf(ChkOnlyCashReceivedVoucher.Value = 1, "and vouchertype = 'CRV' ", "") & IIf(ChkIncludeJV.Value = 0, "and vouchertype <> 'CJV' ", "") & IIf(ChkIncludeBV.Value = 0, "and vouchertype <> 'CBV' ", "") & "Order By VoucherDate, VoucherType, voucherno"
  Set Rs = CN.Execute(vSQL)
  Rs.Filter = "Debit > 0 OR Credit > 0"
  FunRefreshData = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
  FunRefreshData = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
'   If ChkOnlyCashReceivedVoucher.Value = 1 Then Else Set RptReportViewer.Report = New CrpCashBook
   If ChkOnlyCashReceivedVoucher.Value = 1 And ChkSummary.Value = 0 Then
      Set RptReportViewer.Report = New CrpCashBookGroupByAccountNo
   ElseIf ChkSummary.Value = 1 Then
      Set RptReportViewer.Report = New CrpCashBookSumGroupByAccountNo
   Else
      If ObjRegistry.ShowMultiBranches = False Then
         Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\AccountReports\CrpCashBook.rpt")
      Else
         Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\AccountReports\CrpCashBookStoreWise.rpt")
      End If
'      Set RptReportViewer.Report = New CrpCashBook
   End If
   
   
   
   RptReportViewer.Report.ReportTitle = "Cash Book"
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue " Date From :" & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To : " & Format(DtpTo.DateValue, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.PaperOrientation = crPortrait
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
  ShowPicture Me, 2
  AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
  SetWindowText Me.hWnd, "Cash Book"
  DtpFrom.DateValue = Date 'Format("01/" & Month(Date) & "/" & Year(Date), "dd/MM/yy")
  DtpTo.DateValue = Date
  
  If ObjRegistry.ShowMultiBranches = False Then
   LblStoreID.Visible = False
   LblStoreName.Visible = False
   TxtStoreID.Visible = False
   TxtStoreName.Visible = False
   BtnStore.Visible = False
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
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
Private Sub BtnStore_Click()
If FunSelectStore(ssButton, False) = True Then
      TxtOrganizationID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub


Private Sub TxtStoreID_Change()
If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtStoreID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectStore(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectStore(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectStore = False: Exit Function
        TxtStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    If Trim(TxtStoreID.Text) = "" Then Exit Function
'    If Len(TxtStoreID.Text) <= 3 Then
'      TxtStoreID.Text = Right("000" + CStr(Val(TxtStoreID.Text)), 3)
'    End If
    If TxtStoreID.Text = "" Then FunSelectStore = False: Exit Function
    vStrSQL = " Select StoreName FROM Stores where StoreID='" & TxtStoreID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtStoreName.Text = !StoreName
          FunSelectStore = True
          .Close
          Exit Function
      Else
          FunSelectStore = False
          .Close
          TxtStoreID.Text = ""
          TxtStoreName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function


