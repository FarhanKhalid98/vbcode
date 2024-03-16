VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptAccount 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkDetailed 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Account Detail"
      Height          =   255
      Left            =   6038
      TabIndex        =   1
      Top             =   4815
      Width           =   3285
   End
   Begin VB.TextBox TxtAccountNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5078
      MaxLength       =   10
      TabIndex        =   0
      Top             =   4065
      Width           =   1020
   End
   Begin VB.TextBox TxtaccountName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6458
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4065
      Width           =   3825
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   5700
      TabIndex        =   4
      Top             =   6855
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
      MICON           =   "RptAccount.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7050
      TabIndex        =   5
      Top             =   6855
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
      MICON           =   "RptAccount.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8385
      TabIndex        =   6
      Top             =   6855
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
      MICON           =   "RptAccount.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSearch 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6098
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4065
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
      MICON           =   "RptAccount.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   6165
      TabIndex        =   2
      Top             =   5820
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
      Left            =   7890
      TabIndex        =   3
      Top             =   5820
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
      Left            =   7920
      TabIndex        =   13
      Top             =   5595
      Width           =   705
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
      Left            =   6165
      TabIndex        =   12
      Top             =   5595
      Width           =   885
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Summary / Detail"
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
      UseMnemonic     =   0   'False
      Width           =   3405
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
      Left            =   5078
      TabIndex        =   10
      Top             =   3855
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
      Left            =   6473
      TabIndex        =   9
      Top             =   3855
      Width           =   1335
   End
End
Attribute VB_Name = "RptAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
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
  If ChkDetailed = 1 Then
    Set Rs = CN.Execute("EXECUTE ProdrptAccountDetail " & IIf(Trim(TxtAccountNo.Text) = "", "Null", "'" & TxtAccountNo.Text & "'") & ", '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'")
  Else
    Set Rs = CN.Execute("EXECUTE ProdrptAccountSummary " & IIf(Trim(TxtAccountNo.Text) = "", "Null", "'" & TxtAccountNo.Text & "'") & ", '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'")
  End If
    
  FunRefreshData = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
  FunRefreshData = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
   If ChkDetailed = 1 Then
     Set RptReportViewer.Report = New CrptAccountDetail
     RptReportViewer.Report.ReportTitle = "Account Detail"
   Else
     Set RptReportViewer.Report = New CrptAccountSummary
     RptReportViewer.Report.ReportTitle = "Account Summary"
   End If
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
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
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
