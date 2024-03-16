VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmReceivingInvoicePrint 
   BorderStyle     =   0  'None
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   Icon            =   "FrmReceivingInvoicePrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Left            =   13725
      TabIndex        =   7
      Top             =   990
      Visible         =   0   'False
      Width           =   4200
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4200
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Tag             =   "NC"
         Text            =   "FrmReceivingInvoicePrint.frx":0ECA
         Top             =   300
         Width           =   3975
      End
      Begin VB.Label LblClose 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3915
         TabIndex        =   9
         Top             =   90
         Width           =   135
      End
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   4590
      TabIndex        =   1
      Top             =   2235
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   9
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
      Mandatory       =   1
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpBillDate 
      Height          =   315
      Left            =   5235
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   2235
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
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   6540
      TabIndex        =   3
      Top             =   2235
      Width           =   750
      _ExtentX        =   1323
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
   End
   Begin SITextBox.Txt TxtEmployeeName 
      Height          =   315
      Left            =   7650
      TabIndex        =   5
      Top             =   2235
      Width           =   1530
      _ExtentX        =   2699
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
   Begin JeweledBut.JeweledButton BtnEmployee 
      Height          =   330
      Left            =   7290
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2235
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
      MICON           =   "FrmReceivingInvoicePrint.frx":1041
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7980
      TabIndex        =   14
      Top             =   4800
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
      MICON           =   "FrmReceivingInvoicePrint.frx":105D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6660
      TabIndex        =   15
      Top             =   4800
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
      MICON           =   "FrmReceivingInvoicePrint.frx":1079
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   9180
      TabIndex        =   6
      Tag             =   "D"
      Top             =   2235
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Alignment       =   1
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
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   9180
      TabIndex        =   16
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   6540
      TabIndex        =   13
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   7650
      TabIndex        =   12
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " Bill ID"
      Height          =   195
      Left            =   4590
      TabIndex        =   11
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   5235
      TabIndex        =   10
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving Invoice Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   0
      Top             =   270
      Width           =   3975
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   13320
      Top             =   1463
      Width           =   345
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
      Begin VB.Menu MniCostPrice 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmReceivingInvoicePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim vMode As FormMode
Dim vCounter As Integer
Dim vDate As Date, vHDiff As Integer, vSystemDate As Boolean
'Dim RsDetail As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vMaxBinID As Integer
Dim vIsNewRecord As Boolean
Dim Flag As Boolean, vIsEdit As Boolean
Dim vSave As Boolean
Dim vBm As Variant
Dim UniCode As Variant
Dim DateFlag As Boolean, vAutoEnterBeforeQty As Boolean
'Dim vSystemDate As Date
Dim sSQL As String, vWhere As String
Dim vStrSQL As String, vAutoPrintSaleOrder As Boolean, vPrintKitchenInoices As Boolean
Dim vQtyLoose As Double, vTotalAmount As Double
Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String, vTotDisc As Double
Dim vCompanyAddress, vCompanyCity, vCompanyPhoneNo, vAddSpace, vStatement   As String
Dim i As Integer, vCashDrawer As Boolean, vLaserInvoice As Boolean, vPrintHeader  As Boolean, vNoofPrints As Byte, vX As Integer, vY As Integer
'----------------------------------
Private Sub BtnClose_Click()
    Unload Me
End Sub

Private Sub BtnPrint_Click()
    vStrSQL = "Select * from Employees"
     
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
   
   If RsReport.RecordCount = 0 Then Exit Sub
   RptReportViewer.Report.ReportTitle = "Receiving Invoice Print"
   Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpReceivingInvoicePrintAurora.rpt")

    RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
    RptReportViewer.Report.ParameterFields(2).AddCurrentValue vCompanyAddress & vCompanyCity
    RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy
    RptReportViewer.Report.ParameterFields(4).AddCurrentValue vCompanyPhoneNo
    RptReportViewer.Report.ParameterFields(5).AddCurrentValue vAddSpace
    RptReportViewer.Report.ParameterFields(6).AddCurrentValue IIf(TxtBillID.Text = "", "", TxtBillID.Text)
    RptReportViewer.Report.ParameterFields(7).AddCurrentValue DtpBillDate.DateValue
    RptReportViewer.Report.ParameterFields(8).AddCurrentValue IIf(TxtEmployeeName.Text = "", "", TxtEmployeeName.Text)
    RptReportViewer.Report.ParameterFields(9).AddCurrentValue IIf(TxtAmount.Text = "", "", TxtAmount.Text)
    RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
'    RptReportViewer.Report.PrintOut False
    RptReportViewer.Show
    Exit Sub
ErrorHandler:
   Call ShowErrorMessage

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   
   If KeyCode = vbKeyEscape Then
        FraHelp.Visible = False
   ElseIf KeyCode = vbKeyReturn Then
        keybd_event 9, 1, 1, 1
        KeyCode = 0
   End If
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   vCompanyName = ObjRegistry.CompanyName
   vCompanyAddress = ObjRegistry.CompanyAddress
   vCompanyCity = IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   vCompanyPhoneNo = IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)

    Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub TxtEmployeeID_Change()
   If ActiveControl.Name <> TxtEmployeeID.Name Then Exit Sub
   If TxtEmployeeName.Text <> "" Then TxtEmployeeName.Text = ""
End Sub

Private Sub TxtEmployeeID_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If TxtEmployeeName.Text <> "" Then Exit Sub
    If TxtEmployeeID.Text = "" Then Exit Sub
    Dim vTemp As Boolean
    vTemp = Not FunSelectEmployee(ssValidate, True)
    If vTemp = True Then
        vTemp = Not FunSelectEmployee(ssButton, False)
    End If
    Cancel = vTemp
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub
Private Function FunSelectEmployee(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchEmployee.Show vbModal, Me
        If SchEmployee.ParaOutEmployeeID = "" Then FunSelectEmployee = False: Exit Function
        TxtEmployeeID.Text = SchEmployee.ParaOutEmployeeID
    End If
    '---------------------------
    If Trim(TxtEmployeeID.Text) = "" Then Exit Function
    sSQL = "Select *" & vbCrLf _
            + " from Employees" & vbCrLf _
            + " where isLockEmployee = 0 and EmpID = " & Val(TxtEmployeeID.Text)
    With cn.Execute(sSQL)
      If .RecordCount > 0 Then
        TxtEmployeeName.Text = !empname
        FunSelectEmployee = True
        .Close
        Exit Function
      Else
        FunSelectEmployee = False
        .Close
        TxtEmployeeID.Text = ""
        Exit Function
      End If
    End With
Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function


