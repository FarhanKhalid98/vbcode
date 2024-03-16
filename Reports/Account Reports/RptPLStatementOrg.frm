VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptPLStatementOrg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptPLStatementOrg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5655
      TabIndex        =   19
      Top             =   5520
      Width           =   4050
      Begin VB.OptionButton OptMovingAvg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Moving Avg"
         Height          =   195
         Left            =   45
         TabIndex        =   22
         ToolTipText     =   "Simple Moving Average"
         Top             =   45
         Value           =   -1  'True
         Width           =   1350
      End
      Begin VB.OptionButton OptWeightedAvg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Weighted Avg"
         Height          =   195
         Left            =   1440
         TabIndex        =   21
         ToolTipText     =   "Weighted Mean"
         Top             =   45
         Width           =   1485
      End
      Begin VB.OptionButton OptLastPrice 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Last Price"
         Height          =   195
         Left            =   2970
         TabIndex        =   20
         Top             =   45
         Width           =   1035
      End
   End
   Begin VB.CheckBox ChkExclude 
      BackColor       =   &H00B98A03&
      Caption         =   "Exclude Accounts Having Zero Balance."
      Height          =   255
      Left            =   7099
      TabIndex        =   7
      Top             =   7410
      Visible         =   0   'False
      Width           =   3285
   End
   Begin JeweledBut.JeweledButton CmdPreview 
      Height          =   420
      Left            =   5475
      TabIndex        =   4
      Top             =   6255
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
      MICON           =   "RptPLStatementOrg.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   6825
      TabIndex        =   5
      Top             =   6255
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
      MICON           =   "RptPLStatementOrg.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdClose 
      Height          =   420
      Left            =   8175
      TabIndex        =   6
      Top             =   6255
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
      MICON           =   "RptPLStatementOrg.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3285
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
      MICON           =   "RptPLStatementOrg.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   4980
      TabIndex        =   0
      Top             =   3285
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
      Left            =   6360
      TabIndex        =   10
      Tag             =   "nc"
      Top             =   3285
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
      Left            =   5940
      TabIndex        =   2
      Top             =   4905
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
      Left            =   7665
      TabIndex        =   3
      Top             =   4905
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
   Begin JeweledBut.JeweledButton BtnSession 
      Height          =   330
      Left            =   6000
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4035
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
      MICON           =   "RptPLStatementOrg.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSessionID 
      Height          =   315
      Left            =   4995
      TabIndex        =   1
      Top             =   4035
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
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
   Begin SITextBox.Txt TxtSessionName 
      Height          =   315
      Left            =   6360
      TabIndex        =   16
      Top             =   4035
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
   Begin VB.Label Label47 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Session ID"
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
      Left            =   4995
      TabIndex        =   18
      Top             =   3825
      Width           =   930
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Session Name"
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
      Left            =   6360
      TabIndex        =   17
      Top             =   3825
      Width           =   1215
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
      Left            =   7695
      TabIndex        =   14
      Top             =   4680
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
      Left            =   5940
      TabIndex        =   13
      Top             =   4680
      Width           =   885
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
      Left            =   4980
      TabIndex        =   12
      Top             =   3060
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
      Left            =   6360
      TabIndex        =   11
      Top             =   3060
      Width           =   1620
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profit && Loss Statement"
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
      TabIndex        =   8
      Top             =   270
      Width           =   3105
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "RptPLStatementOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Public ShowDetailed As Boolean
Dim vStrSQL As String, vCompanyName As String, vAddress As String, vEmail As String

Private Sub BtnGroup_Click()
   If FunSelectOrganizaton(ssButton, False) = True Then
      TxtSessionID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub


Private Sub BtnSession_Click()
   If FunSelectSession(ssButton, False) = True Then
      DtpFrom.SetFocus
   Else
      TxtSessionID.SetFocus
   End If
End Sub

Private Function FunSelectSession(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSession.Show vbModal, Me
        If SchSession.ParaOutSessionID = "" Then FunSelectSession = False: Exit Function
        TxtSessionID.Text = SchSession.ParaOutSessionID
    End If
    '---------------------------
    If Trim(TxtSessionID.Text) = "" Then Exit Function
    If InStr(1, TxtSessionID.Text, ",") > 0 Then TxtSessionName.Text = "Selected Sessions": Exit Function
    vStrSQL = "Select * FROM Sessions s where SessionID=" & Val(TxtSessionID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSessionName.Text = !SessionName
          FunSelectSession = True
          .Close
          Exit Function
      Else
          FunSelectSession = False
          .Close
          TxtSessionID.Text = ""
          TxtSessionName.Text = "All Sessions"
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtSessionID_Change()
  On Error GoTo ErrorHandler
   If TxtSessionID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSessionID.Name Then Exit Sub
   If TxtSessionName.Text <> "" Then TxtSessionName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSessionID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSessionID.Name Then Exit Sub
   On Error GoTo ErrorHandler
'   If TxtSessionName.Text <> "All Sessions" Then Exit Sub
   If Trim(TxtSessionID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSession(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSession(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
         Case TxtOrganizationID.Name: If FunSelectOrganizaton(ssFunctionKey, True) = True Then DtpFrom.SetFocus
      End Select
  End If
End Sub

Private Function FunRefreshData() As Boolean
    On Error GoTo ErrorHandler
    Dim vSQL As String
    CN.Execute "EXECUTE SPAccountsBalancesNew '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
    If OptMovingAvg.Value = True Then
        vSQL = "EXECUTE SPPLStatementValueOrgMoving '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text) & "," & IIf(Trim(TxtSessionID.Text) = "", "Null", TxtSessionID.Text)
        CN.Execute vSQL
        vSQL = "Select * from tempPLStatement where Bal <> 0 Order by  OrganizationID, sr"
    ElseIf OptLastPrice.Value = True Then
         If ObjRegistry.PLSamePR = True Then
            CN.Execute "exec SPProductPurchase '" & DtpTo.DateValue & "'"
            vSQL = "EXECUTE SPPLStatementSaleCostOrgLast '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text) & "," & IIf(Trim(TxtSessionID.Text) = "", "Null", TxtSessionID.Text)
         Else
            vSQL = "EXECUTE SPPLStatementValueOrgLast '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text) & "," & IIf(Trim(TxtSessionID.Text) = "", "Null", TxtSessionID.Text)
         End If
         CN.Execute vSQL
         vSQL = "Select * from tempPLStatement where Bal <> 0 Order by  OrganizationID, sr"
    ElseIf OptWeightedAvg.Value = True Then
         If ObjRegistry.PLSamePR = True Then
            CN.Execute "exec SPAverageCost '" & DtpTo.DateValue & "'"
            vSQL = "EXECUTE SPPLStatementSaleCostOrg '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text) & "," & IIf(Trim(TxtSessionID.Text) = "", "Null", TxtSessionID.Text)
         Else
            CN.Execute "exec SPAverageCostNew '" & DtpTo.DateValue & "'"
            vSQL = "EXECUTE SPPLStatementValueOrg '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text) & "," & IIf(Trim(TxtSessionID.Text) = "", "Null", TxtSessionID.Text)
         End If
        CN.Execute vSQL
        vSQL = "Select * from tempPLStatement where Bal <> 0 Order by  OrganizationID, sr"
    End If
    Set Rs = CN.Execute(vSQL)
    FunRefreshData = True
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
    FunRefreshData = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
   Set RptReportViewer.Report = New CrpPLStatementValurOrg
  'RptReportViewer.Report.TxtCompanyName.SetText ObjSupernetRegistry.CompanyName
   RptReportViewer.Report.ReportTitle = "Profit & Loss Statement"
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity) & IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue "From : " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & ",   To : " & Format(DtpTo.DateValue, "dd/MM/yyyy")
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
   SetWindowText Me.hWnd, "Profit & Loss Statement"
   DtpFrom.DateValue = Format("01/" & Month(Date) & "/" & Year(Date), "dd/MM/yy")
   DtpTo.DateValue = Date
   OptLastPrice.Visible = ObjRegistry.ShowLastPriceOption
   OptWeightedAvg.Visible = ObjRegistry.ShowWeightedAvgOption
   OptMovingAvg.Visible = ObjRegistry.ShowMovingAvgOption
   If ObjRegistry.ShowWeightedAvgOption Then
      OptWeightedAvg.Value = True
   ElseIf ObjRegistry.ShowMovingAvgOption Then
      OptMovingAvg.Value = True
   Else
      OptLastPrice.Value = True
   End If
      
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

