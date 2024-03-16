VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptPendingBillsToVendor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "RptPendingBillsToVendor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptFromToDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBD0AB&
      Caption         =   "OptSummary"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3060
      TabIndex        =   1
      Top             =   3180
      Width           =   210
   End
   Begin VB.OptionButton OptAllDates 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBD0AB&
      Caption         =   "OptSummary"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3060
      TabIndex        =   0
      Top             =   2670
      Width           =   210
   End
   Begin JeweledBut.JeweledButton BtnAccount 
      Height          =   330
      Left            =   4260
      TabIndex        =   9
      Top             =   4290
      Width           =   420
      _ExtentX        =   741
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
      MICON           =   "RptPendingBillsToVendor.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin VB.TextBox TxtAccountName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   4710
      TabIndex        =   8
      Top             =   4290
      Width           =   4275
   End
   Begin VB.TextBox TxtAccountNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3090
      TabIndex        =   2
      Top             =   4290
      Width           =   1140
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6585
      TabIndex        =   5
      Top             =   5385
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
      MICON           =   "RptPendingBillsToVendor.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3810
      TabIndex        =   3
      Top             =   5385
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
      MICON           =   "RptPendingBillsToVendor.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5190
      TabIndex        =   4
      Top             =   5385
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
      MICON           =   "RptPendingBillsToVendor.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5565
      TabIndex        =   10
      Top             =   3180
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   51183619
      CurrentDate     =   38244
   End
   Begin MSComCtl2.DTPicker DtpTo 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7140
      TabIndex        =   11
      Top             =   3180
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   51183619
      CurrentDate     =   38244
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Purchase Bills"
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
      Left            =   1980
      TabIndex        =   16
      Top             =   180
      Width           =   2910
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
      Left            =   3435
      TabIndex        =   15
      Top             =   2670
      Width           =   840
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
      Left            =   3435
      TabIndex        =   14
      Top             =   3180
      Width           =   1785
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   195
      Left            =   5565
      TabIndex        =   13
      Top             =   2985
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   195
      Left            =   7140
      TabIndex        =   12
      Top             =   2985
      Visible         =   0   'False
      Width           =   195
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
      Left            =   4710
      TabIndex        =   7
      Top             =   4050
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
      Left            =   3090
      TabIndex        =   6
      Top             =   4050
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
Attribute VB_Name = "RptPendingBillsToVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim sSql As String, vDate As String
Dim RsReport As New ADODB.Recordset

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
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
        TxtAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = "Select AccountNo,AccountName from chartofaccounts where AccountNo='" & SchAccounts.ParaOutAccountNo & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtAccountName.Text = !AccountName
          FunSelectAccount = True
          .Close
          'If BtnSave.Visible = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectAccount = False
          .Close
          TxtAccountNo.Text = ""
          TxtAccountName.Text = ""
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
        RptReportViewer.Caption = Me.Caption
        RptReportViewer.Show vbModal
    End If
End Sub

Private Sub BtnPrint_Click()
    If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub DtpFrom_Change()
   vDate = " and purchasedate BETWEEN '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'"
End Sub

Private Sub DtpTo_Change()
   vDate = " and purchasedate BETWEEN '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'"
End Sub

Private Sub Form_Load()
   ShowPicture Me
    SetWindowText Me.hWnd, "Pending  Bills to Vendor"
    Label4_Click
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
   Set RptPendingBillsToVendor = Nothing
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
   Dim vSQL As String
   vSQL = "select purid, purchasedate, totalamount-isnull(billdisc,0)+isnull(OtherCharges,0) as netamount, isnull(paidamount,0) as paidamount, totalamount-isnull(billdisc,0)-isnull(paidamount,0) bal, partyname from purchaseheader h " & _
         " inner join parties p on h.vendorid = p.partyid" & _
         " Where totalamount - IsNull(billdisc, 0) - IsNull(paidamount, 0) + isnull(OtherCharges,0) > 2 " & IIf(TxtAccountNo.Text = "", "", " and vendorID='" & TxtAccountNo.Text & "'") & vDate
   Set RsReport = CN.Execute(vSQL)
   If RsReport.BOF Then
      MsgBox "No record exists.", vbInformation, Me.Caption
      Me.MousePointer = vbDefault
      Exit Function
   End If
   Set RptReportViewer.Report = New CrpPendingBillstoVendor
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue " Account Name : " & TxtAccountName.Text
   With CN.Execute("Select CompanyName,Address,City,PhoneNo,email from Company")
      If .RecordCount > 0 Then
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(IsNull(!CompanyName), "", CStr(!CompanyName))
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(IsNull(!Address), "", !Address) & IIf(IsNull(!City), "", ", " & !City & ".")
         RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(IsNull(!PhoneNo), "", CStr(!PhoneNo))
      End If
      .Close
   End With
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
   RptReportViewer.Report.Database.SetDataSource RsReport
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

Private Sub Label4_Click()
   OptAllDates.Value = True
   Call OptAllDates_Click
End Sub

Private Sub OptAllDates_Click()
   If Label7.Visible = True Then Label7.Visible = False
   If DtpFrom.Visible = True Then DtpFrom.Visible = False
   If DtpTo.Visible = True Then DtpTo.Visible = False
   If Label9.Visible = True Then Label9.Visible = False
   vDate = ""
End Sub

Private Sub OptFromToDate_Click()
   Label9.Visible = True
   Label7.Visible = True
   DtpFrom.Visible = True
   DtpTo.Visible = True
   vDate = " and purchasedate BETWEEN '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'"
End Sub

Private Sub TxtAccountNo_Change()
   If ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   If TxtAccountName.Text <> "" Then
      TxtAccountName.Text = ""
   End If
End Sub

Private Sub TxtAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> "TxtAccountNo" Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtAccountName.Text <> "" Then Exit Sub
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
