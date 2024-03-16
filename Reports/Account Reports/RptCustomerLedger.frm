VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptCustomerLedger 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptCustomerLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtAccountNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5070
      MaxLength       =   10
      TabIndex        =   0
      Top             =   4275
      Width           =   1020
   End
   Begin VB.TextBox TxtaccountName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6465
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4275
      Width           =   3825
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   5310
      TabIndex        =   3
      Top             =   6645
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
      MICON           =   "RptCustomerLedger.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   6660
      TabIndex        =   4
      Top             =   6645
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
      MICON           =   "RptCustomerLedger.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7995
      TabIndex        =   5
      Top             =   6645
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
      MICON           =   "RptCustomerLedger.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5070
      TabIndex        =   1
      Top             =   4950
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   112197635
      CurrentDate     =   38244
   End
   Begin MSComCtl2.DTPicker DtpTo 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6840
      TabIndex        =   2
      Top             =   4950
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   92930051
      CurrentDate     =   38244
   End
   Begin JeweledBut.JeweledButton BtnSearch 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6090
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4275
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
      MICON           =   "RptCustomerLedger.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Ledger"
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
      TabIndex        =   12
      Top             =   270
      UseMnemonic     =   0   'False
      Width           =   2295
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
      Height          =   225
      Left            =   5070
      TabIndex        =   11
      Top             =   4065
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Name"
      Height          =   225
      Left            =   6465
      TabIndex        =   10
      Top             =   4065
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   225
      Left            =   6840
      TabIndex        =   7
      Top             =   4725
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      Height          =   225
      Left            =   5070
      TabIndex        =   6
      Top             =   4725
      Width           =   1095
   End
End
Attribute VB_Name = "RptCustomerLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSQL As String

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInWhereClause = " and AccountNo like '62%'"
        SchAccounts.ParaInDetail = ""
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
         Case TxtAccountNo.Name: If FunSelectAccount(ssFunctionKey, True) = True Then DtpFrom.SetFocus
      End Select
   End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
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
  Set Rs = CN.Execute("EXECUTE SPAccountsLedger '" & TxtAccountNo.Text & "', '" & DtpFrom.Value & "','" & DtpTo.Value & "'")
  FunRefreshData = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
  FunRefreshData = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
   Set RptReportViewer.Report = New CrpAccountLedger
   'RptReportViewer.Report.TxtCompanyName.SetText CN.Execute("select companyname from Project_Registry").Fields(0).Value
   RptReportViewer.Report.ReportTitle = "Account Ledger"
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue TxtAccountNo.Text & " - " & TxtaccountName.Text
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue "From : " & Format(DtpFrom.Value, "dd/MM/yyyy") & ",   To : " & Format(DtpTo.Value, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity) & IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
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
  SetWindowText Me.hWnd, "Ledger"
  DtpFrom.Value = Date - 30
  DtpTo.Value = Date
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub
