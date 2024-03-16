VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form RptCustomerSaleBills 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptCustomerSaleBills.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkReceivedBills 
      BackColor       =   &H00FF8080&
      Caption         =   "Received Bills"
      Height          =   255
      Left            =   6233
      TabIndex        =   4
      Top             =   5948
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox ChkUnReceivedBills 
      BackColor       =   &H00FF8080&
      Caption         =   "UnReceived Bills"
      Height          =   255
      Left            =   7793
      TabIndex        =   5
      Top             =   5948
      Value           =   1  'Checked
      Width           =   1560
   End
   Begin VB.OptionButton OptFromToDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBD0AB&
      Caption         =   "OptSummary"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4823
      TabIndex        =   1
      Top             =   4238
      Width           =   210
   End
   Begin VB.OptionButton OptAllDates 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBD0AB&
      Caption         =   "OptSummary"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4823
      TabIndex        =   0
      Top             =   3728
      Width           =   210
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8430
      TabIndex        =   8
      Top             =   6983
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Close"
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
      MICON           =   "RptCustomerSaleBills.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5655
      TabIndex        =   6
      Top             =   6983
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Pre&view"
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
      MICON           =   "RptCustomerSaleBills.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7035
      TabIndex        =   7
      Top             =   6983
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Print"
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
      MICON           =   "RptCustomerSaleBills.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   7478
      TabIndex        =   2
      Top             =   4260
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
      Left            =   9233
      TabIndex        =   3
      Top             =   4260
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
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   6945
      TabIndex        =   14
      Top             =   5175
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   50
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
   Begin JeweledBut.JeweledButton BtnCustomer 
      Height          =   330
      Left            =   6585
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5175
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
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
      MICON           =   "RptCustomerSaleBills.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   315
      Left            =   5415
      TabIndex        =   16
      Top             =   5175
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   10
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
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
      Left            =   5415
      TabIndex        =   18
      Top             =   4935
      Width           =   1050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
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
      Left            =   6945
      TabIndex        =   17
      Top             =   4935
      Width           =   1335
   End
   Begin VB.Label LblFromDate 
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
      Left            =   7478
      TabIndex        =   13
      Top             =   4035
      Width           =   885
   End
   Begin VB.Label LblToDate 
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
      Left            =   9248
      TabIndex        =   12
      Top             =   4035
      Width           =   705
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Sale Bills"
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
      Left            =   2700
      TabIndex        =   11
      Top             =   270
      Width           =   2580
   End
   Begin VB.Label LblAllDates 
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
      Left            =   5198
      TabIndex        =   10
      Top             =   3728
      Width           =   840
   End
   Begin VB.Label LblSeletedDates 
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
      Left            =   5198
      TabIndex        =   9
      Top             =   4238
      Width           =   1785
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
Attribute VB_Name = "RptCustomerSaleBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim sSql As String, vDate As String
Dim RsReport As New ADODB.Recordset

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      BtnPreview.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

Private Sub TxtCustomerID_Change()
   If TxtCustomerID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   If TxtCustomerName.Text <> "All Customers" Then
      TxtCustomerName.Text = "All Customers"
   End If
End Sub

Private Sub TxtCustomerID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerName.Text <> "All Customers" Then Exit Sub
   If Trim(TxtCustomerID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCustomer(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCustomer(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCustomer.Show vbModal, Me
        If SchCustomer.ParaOutCustomerID = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchCustomer.ParaOutCustomerID
    End If
    '---------------------------
    If Trim(TxtCustomerID.Text) = "" Then Exit Function
    VStrSQL = " Select * FROM Parties where PartyID = " & Val(TxtCustomerID.Text) & " AND PartyType <> 'C'"
    With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerName.Text = !PartyName
          FunSelectCustomer = True
          .Close
          Exit Function
      Else
          FunSelectCustomer = False
          .Close
          TxtCustomerID.Text = ""
          TxtCustomerName.Text = "All Customers"
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
        RptReportViewer.Caption = Me.Caption
        RptReportViewer.Show vbModal
    End If
End Sub

Private Sub BtnPrint_Click()
    If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub DtpFrom_Change()
   vDate = " and h.BillDate BETWEEN '" & DtpFrom.DateValue & "' AND '" & DtpTo.DateValue & "'"
End Sub

Private Sub DtpTo_Change()
   vDate = " and h.BillDate BETWEEN '" & DtpFrom.DateValue & "' AND '" & DtpTo.DateValue & "'"
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Customer Sale Bills"
   TxtCustomerName.Text = "All Customers"
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   OptFromToDate.Value = True
'   LblSeletedDates_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
   Set RsReport = Nothing
   Set RptCustomerSaleBills = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunRefreshData() As Boolean
  On Error GoTo ErrorHandler
  FunRefreshData = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
  FunRefreshData = False
End Function

Private Function SetReport() As Boolean
   On Error GoTo ErrorHandler
   SetReport = False
   Me.MousePointer = vbHourglass
   Dim vSQL As String, vParameter As String
   vParameter = IIf(ChkReceivedBills.Value = 0, IIf(ChkUnReceivedBills.Value = 1, " and (totalamount - IsNull(billdisc, 0) + isnull(OtherCharges,0)) - (isnull(CashReceived,0) + isnull(amount,0)+isnull(discount,0)) > 2", " and 1 = 2"), IIf(ChkUnReceivedBills.Value = 1, "", " and (totalamount - IsNull(billdisc, 0) + isnull(OtherCharges,0)) - (isnull(CashReceived,0) + isnull(amount,0)+isnull(discount,0)) <= 2"))
   vSQL = " select h.BillID, h.BillDate, totalamount - isnull(billdisc,0) + isnull(OtherCharges,0) as netamount, " & vbCrLf _
      + " isnull(CashReceived,0) + isnull(amount,0) + isnull(discount,0) as CashReceived, ca.AccountName as partyname, " & vbCrLf _
      + " case when LastReceivedDate is not null then LastReceivedDate when isnull(CashReceived,0) <> 0 then h.BillDate end as LastReceivedDate," & vbCrLf _
      + " (totalamount - IsNull(billdisc, 0) + isnull(OtherCharges,0)) - (isnull(CashReceived,0) + isnull(amount,0)+isnull(discount,0)) bal" & vbCrLf _
      + " from SaleHeader h left outer join " & vbCrLf _
      + " (select BillID, BillDate, max(BillDate) as LastReceivedDate, sum(amount) as amount, sum(discount) as Discount from RecoveryInvoice group by BillID, BillDate)i " & vbCrLf _
      + " on i.BillID = h.BillID and i.BillDate = h.BillDate" & vbCrLf _
      + " inner join (select BillID, BillDate from SaleBody Group By BillID, BillDate)b on h.BillID = b.BillID and h.BillDate = b.BillDate" & vbCrLf _
      + " left outer join ChartofAccounts ca on h.Customerid = ca.AccountNo" & vbCrLf _
      + " where 1=1 " & vParameter & vDate & IIf(TxtCustomerID.Text = "", "", " and h.Customerid = " & Val(TxtCustomerID.Text)) & vbCrLf _
      + " order by h.BillDate, h.BillID"
   Set RsReport = cn.Execute(vSQL)
   Set RptReportViewer.Report = New CrpCustomerSaleBills
   If RsReport.BOF Then
      MsgBox "No record exists.", vbInformation, Me.Caption
      Me.MousePointer = vbDefault
      Exit Function
   End If
   
   RptReportViewer.Report.Database.SetDataSource RsReport
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(TxtCustomerName.Text = "All Customers", "All Customers", "Customer Name : " & TxtCustomerName.Text)
   RptReportViewer.Report.ParameterFields(6).AddCurrentValue IIf(ChkReceivedBills.Value = 0, IIf(ChkUnReceivedBills.Value = 1, "Unpaid Bills", ""), IIf(ChkUnReceivedBills.Value = 1, "Paid and Unpaid Bills.", "Paid Bills."))
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   SetReport = True
   Me.MousePointer = vbDefault
   Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub LblAllDates_Click()
   OptAllDates.Value = True
   Call OptAllDates_Click
End Sub

Private Sub LblSeletedDates_Click()
   OptFromToDate.Value = True
   Call OptFromToDate_Click
End Sub

Private Sub OptAllDates_Click()
   If DtpFrom.Visible = True Then DtpFrom.Visible = False
   If LblFromDate.Visible = True Then LblFromDate.Visible = False
   If DtpTo.Visible = True Then DtpTo.Visible = False
   If LblToDate.Visible = True Then LblToDate.Visible = False
   vDate = ""
End Sub

Private Sub OptFromToDate_Click()
   LblFromDate.Visible = True
   LblToDate.Visible = True
   DtpFrom.Visible = True
   DtpTo.Visible = True
   vDate = " and h.BillDate BETWEEN '" & DtpFrom.DateValue & "' AND '" & DtpTo.DateValue & "'"
End Sub
