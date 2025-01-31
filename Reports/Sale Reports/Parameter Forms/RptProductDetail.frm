VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptCustomerWiseSaleDetail 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "RptProductDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6780
      TabIndex        =   5
      Top             =   6210
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
      MICON           =   "RptProductDetail.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3960
      TabIndex        =   3
      Top             =   6210
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
      MICON           =   "RptProductDetail.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5340
      TabIndex        =   4
      Top             =   6210
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
      MICON           =   "RptProductDetail.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Top             =   3315
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   11
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
      IntegralPoint   =   10
   End
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   3315
      Width           =   4350
      _ExtentX        =   7673
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
      Left            =   4440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3300
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
      MICON           =   "RptProductDetail.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   4455
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
      Left            =   6075
      TabIndex        =   2
      Top             =   4455
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Detail (Customer Wise)"
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
      Left            =   1980
      TabIndex        =   12
      Top             =   180
      Width           =   3915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   4230
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   195
      Left            =   6090
      TabIndex        =   10
      Top             =   4230
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   9
      Top             =   3090
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4830
      TabIndex        =   8
      Top             =   3090
      Width           =   1335
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
Attribute VB_Name = "RptCustomerWiseSaleDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim sSql As String

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      DtpFrom.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
      RptReportViewer.Caption = Me.Caption
      RptReportViewer.Show
   End If
End Sub

Private Sub BtnPrint_Click()
   If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
         Case vbKeyV
            If BtnPreview.Enabled Then BtnPreview_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCustomerID.Name: If FunSelectCustomer(ssFunctionKey, True) = True Then DtpFrom.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_Load()
    ShowPicture Me
    SetWindowText Me.hwnd, "Sale Detail (Customer Wise)"
    DtpTo.DateValue = Date
    DtpFrom.DateValue = Date - 30
    TxtCustomerName.Text = "All Customers"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
   Set RptCustomerWiseSaleDetail = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function SetReport() As Boolean
    On Error GoTo ErrorHandler
    SetReport = False
    Me.MousePointer = vbHourglass
    Dim RsReport As New ADODB.Recordset
    Set RsReport = CN.Execute("EXEC ProdRptCustomerWiseSaleDetail '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "','" & IIf(TxtCustomerName.Text = "All Customers", "", TxtCustomerID.Text) & "'")
    Set RptReportViewer.Report = New CrptCustomerWiseSaleDetail
    If RsReport.BOF Then
        MsgBox "No record exists.", vbInformation, Me.Caption
        Me.MousePointer = vbDefault
        Exit Function
    End If
    RptReportViewer.Report.Database.SetDataSource RsReport
    With CN.Execute("Select CompanyName,Address,City,PhoneNo,email from Company")
      If .RecordCount > 0 Then
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(IsNull(!CompanyName), "", CStr(!CompanyName))
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(IsNull(!Address), "", !Address) & IIf(IsNull(!City), "", ", " & !City & ".")
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(IsNull(!PhoneNo), "", CStr(!PhoneNo))
      End If
    .Close
    End With
    RptReportViewer.Report.ParameterFields(4).AddCurrentValue "" 'IIf(Trim(TxtCustomerID.Text) = "", " All Customers ", "Customer : " & TxtCustomerName.Text & " - " & TxtCustomerID.Text)
    RptReportViewer.Report.ParameterFields(5).AddCurrentValue " Date From :" & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To : " & Format(DtpTo.DateValue, "dd/MM/yyyy")
    RptReportViewer.Report.ParameterFields(6).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    RptReportViewer.Report.PaperOrientation = crPortrait
    SetReport = True
    Me.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

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
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCustomer.Show vbModal, Me
        If SchCustomer.ParaOutCustomerID = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchCustomer.ParaOutCustomerID
    End If
    '---------------------------
    If Trim(TxtCustomerID.Text) = "" Then Exit Function
    vStrSQL = " Select * FROM Parties where PartyID = '" & TxtCustomerID.Text & "' AND PartyType = 'C'"
    With CN.Execute(vStrSQL)
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
