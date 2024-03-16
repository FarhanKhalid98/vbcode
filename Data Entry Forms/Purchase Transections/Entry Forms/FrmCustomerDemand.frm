VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmCustomerDemand 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   15360
   ClipControls    =   0   'False
   Icon            =   "FrmCustomerDemand.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtDemandID 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      CausesValidation=   0   'False
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   4566
      TabIndex        =   0
      Top             =   2903
      Width           =   1020
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7695
      TabIndex        =   2
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6390
      TabIndex        =   3
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10305
      TabIndex        =   7
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5085
      TabIndex        =   4
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9000
      TabIndex        =   6
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   3780
      TabIndex        =   5
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtRemarks 
      Height          =   315
      Left            =   5048
      TabIndex        =   1
      Top             =   6934
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   100
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
   Begin SITextBox.Txt TxtVenderID 
      Height          =   315
      Left            =   2070
      TabIndex        =   12
      Top             =   4650
      Width           =   930
      _ExtentX        =   1640
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtVenderName 
      Height          =   315
      Left            =   3360
      TabIndex        =   13
      Top             =   4650
      Width           =   3645
      _ExtentX        =   6429
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
   Begin SITextBox.Txt TxtAddress 
      Height          =   315
      Left            =   7005
      TabIndex        =   14
      Top             =   4650
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   100
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
   Begin SITextBox.Txt TxtCity 
      Height          =   315
      Left            =   11535
      TabIndex        =   15
      Top             =   4650
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin JeweledBut.JeweledButton BtnVender 
      Height          =   330
      Left            =   3000
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4650
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
      MICON           =   "FrmCustomerDemand.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtPaidAmount 
      Height          =   315
      Left            =   8355
      TabIndex        =   23
      Top             =   5865
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Alignment       =   1
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
      IntegralPoint   =   9
   End
   Begin SITextBox.Txt TxtTotalPayable 
      Height          =   315
      Left            =   7050
      TabIndex        =   24
      Top             =   5865
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
      Masked          =   2
   End
   Begin SITextBox.Txt TxtPreviousPayable 
      Height          =   315
      Left            =   5790
      TabIndex        =   25
      Top             =   5865
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
      Masked          =   2
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpDemandDate 
      Height          =   315
      Left            =   5784
      TabIndex        =   29
      Top             =   2903
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
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
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   7509
      TabIndex        =   30
      Tag             =   "NC"
      Top             =   2903
      Width           =   945
      _ExtentX        =   1667
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
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   8814
      TabIndex        =   31
      Tag             =   "NC"
      Top             =   2903
      Width           =   1980
      _ExtentX        =   3493
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
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8454
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2903
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
      MICON           =   "FrmCustomerDemand.frx":0F8E
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Amount"
      Height          =   195
      Left            =   8355
      TabIndex        =   28
      Top             =   5640
      Width           =   900
   End
   Begin VB.Label LblTtlPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payable"
      Height          =   195
      Left            =   7095
      TabIndex        =   27
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Payable"
      Height          =   195
      Left            =   5745
      TabIndex        =   26
      Top             =   5640
      Width           =   1260
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   7509
      TabIndex        =   22
      Top             =   2685
      Width           =   1095
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   8814
      TabIndex        =   21
      Top             =   2685
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender ID"
      Height          =   195
      Left            =   2055
      TabIndex        =   20
      Top             =   4440
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Name"
      Height          =   195
      Left            =   3360
      TabIndex        =   19
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   7005
      TabIndex        =   18
      Top             =   4440
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   11535
      TabIndex        =   17
      Top             =   4470
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GRN Date"
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
      Left            =   5784
      TabIndex        =   11
      Top             =   2685
      Width           =   885
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   5048
      TabIndex        =   10
      Top             =   6694
      Width           =   750
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goods Report Notes (GRN)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   9
      Top             =   270
      Width           =   3510
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GRN ID"
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
      Left            =   4569
      TabIndex        =   8
      Top             =   2670
      Width           =   675
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove this Row"
      End
   End
End
Attribute VB_Name = "FrmCustomerDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vCounter As Integer
Dim Flag As Boolean
Dim sSql As String
Dim vStrSQL As String
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vStrComp As String, vCompanyName As String, vAddress As String, vemail As String
'----------------------------------

Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
'      TxtBillNo.SetFocus
   Else
      TxtVenderID.SetFocus
   End If
End Sub

Private Sub TxtVenderID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtVenderName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVender(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVender(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   BtnPrint.Enabled = False
   
   vStrSQL = " Select D.*, StoreName, GroupName, SubGroupName, CompanyName, PartyName from CustomerDemands D Left Outer Join Stores st on st.storeid = d.storeid" & vbCrLf _
            + " Left outer Join Groups g on D.GroupID = g.GroupID " & vbCrLf _
            + " Left outer Join Companies c on D.CompanyID = c.CompanyID " & vbCrLf _
            + " Left outer Join SubGroups s on D.SubGroupID = s.SubGroupID " & vbCrLf _
            + " Left outer Join Parties p on d.VendorID = p.PartyID" & vbCrLf _
            + " where D.DemandID = " & Val(TxtDemandID.Text)

   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
  
'   Set RptReportViewer.Report = New CrptCustomerDemand
   
   RptReportViewer.Report.ReportTitle = "Customer Demand"
   
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   With cn.Execute("Select CompanyName,Address,City,PhoneNo,email from Company")
      If .RecordCount > 0 Then
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(IsNull(!CompanyName), "", CStr(!CompanyName))
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(IsNull(!Address), "", !Address) & IIf(IsNull(!City), "", ", " & !City & ".")
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(IsNull(!PhoneNo), "", CStr(!PhoneNo))
      End If
    .Close
    End With
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue cn.Execute("Select Name from Manufacturer").Fields(0).Value
   Dim vDevice As String, vDriver As String, vPort As String
   vStrSQL = "Select * from Registry"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         vDevice = IIf(IsNull(!DeviceName), "Abc", !DeviceName)
         vDriver = IIf(IsNull(!DriverName), "Xyz", !DriverName)
         vPort = IIf(IsNull(!Port), "LPT1", !Port)
         RptReportViewer.Report.SelectPrinter vDriver, vDevice, vPort
      End If
   End With
   'RptReportViewer.Report.PaperOrientation = crPortrait
   RptReportViewer.Show
   'RptReportViewer.Report.PrintOut False
   BtnPrint.Enabled = True
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
    BtnPrint.Enabled = True
End Sub

Private Sub BtnClear_Click()
  On Error GoTo ErrorHandler
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
  End If
  If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
  cn.BeginTrans
  cn.Execute "Delete from CustomerDemands Where DemandID = " & Val(TxtDemandID.Text)
  cn.CommitTrans
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  If cn.Errors.Count > 0 Then cn.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Sub GetCustomerDemand()
   On Error GoTo ErrorHandler
   sSql = "Select D.*, StoreName, GroupName, SubGroupName, CompanyName, PartyName from CustomerDemands D Left outer Join Stores st on st.storeid = D.storeid Left outer Join Groups g on D.GroupID = g.GroupID Left outer Join SubGroups s on D.SubGroupID = s.SubGroupID Left outer Join Companies c on D.CompanyID = c.CompanyID Left outer Join Parties p on d.VendorID = p.PartyID where DemandID = " & Val(TxtDemandID.Text)
   With cn.Execute(sSql)
      If Not .BOF Then
          DtpDemandDate.DateValue = !DemandDate
'          TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
'          TxtStoreName.Text = IIf(IsNull(!StoreName), "", !StoreName)
'          TxtGroupID.Text = IIf(IsNull(!GroupID), "", !GroupID)
'          TxtGroupName.Text = IIf(IsNull(!GroupName), "", !GroupName)
'          TxtCompanyID.Text = IIf(IsNull(!companyid), "", !companyid)
'          TxtCompanyName.Text = IIf(IsNull(!CompanyName), "", !CompanyName)
'          TxtSubGroupID.Text = IIf(IsNull(!SubGroupID), "", !SubGroupID)
'          TxtSubGroupName.Text = IIf(IsNull(!SubGroupName), "", !SubGroupName)
'          TxtVendorID.Text = IIf(IsNull(!VendorID), "", !VendorID)
'          TxtVendorName.Text = IIf(IsNull(!PartyName), "", !PartyName)
'          TxtProductName.Text = !ProductName
'          TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
      End If
      .Close
   End With
'   Call PopulateDataToGrid
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
'   SchCustomerDemands.Show vbModal
'   If SchCustomerDemands.ParaOutDemandID <> "" Then
'      TxtDemandID.Text = SchCustomerDemands.ParaOutDemandID
'      GetCustomerDemand
'   End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If vIsNewRecord Then
      If cn.Execute("Select * from CustomerDemands where DemandID = " & Val(TxtDemandID.Text)).RecordCount > 0 Then
         MsgBox "This voucher already exists. A new voucher No. has been generated. Please try again", vbCritical, "Alert"
         TxtDemandID.Text = FunGetMaxID
         Exit Sub
      End If
   End If
'   RsBody.Filter = 0
'   If RsBody.RecordCount = 0 Then
'       MsgBox "Please enter at least one entry to save", vbExclamation, "Alert"
'       If TxtProductID.Visible And TxtProductID.Enabled Then TxtProductID.SetFocus
'       Exit Sub
'   End If
   
   
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
  
   cn.BeginTrans
   sSql = "Select * From CustomerDemands Where DemandID =" & Val(TxtDemandID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, cn, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !DemandID = Val(TxtDemandID.Text)
      End If
'      !DemandDate = DtpDemandDate.DateValue
'      !StoreID = IIf(Trim(TxtStoreID.Text) = "", Null, Val(TxtStoreID.Text))
'      !GroupID = IIf(Trim(TxtGroupID.Text) = "", Null, (TxtGroupID.Text))
'      !companyid = IIf(Trim(TxtCompanyID.Text) = "", Null, Val(TxtCompanyID.Text))
'      !SubGroupID = IIf(Trim(TxtSubGroupID.Text) = "", Null, Val(TxtSubGroupID.Text))
'      !VendorID = IIf(Trim(TxtVendorID.Text) = "", Null, Val(TxtVendorID.Text))
'      !ProductName = Trim(TxtProductName.Text)
'      !Remarks = IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text)
      !UserNo = vUser
      .Update
      .Close
   End With
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Property Get FormStatus() As FormMode
  'Nothing
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
  'Based upon the value of vNewValue, we shall decide what controls to enable/disable
  On Error GoTo ErrorHandler
  vMode = vNewValue
  Select Case vNewValue
    Case Is = NewMode
      Call SubClearFields
'      Call PopulateDataToGrid
      BtnPrint.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtDemandID.Text = FunGetMaxID
      If DtpDemandDate.Enabled And DtpDemandDate.Visible Then DtpDemandDate.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      BtnPrint.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      DtpDemandDate.SetFocus
      vIsNewRecord = False
    Case Is = changeMode
      BtnPrint.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
'  ElseIf KeyCode = vbKeyEscape And (Me.ActiveControl.Name = TxtProductID.Name Or Me.ActiveControl.Name = TxtProductName.Name Or Me.ActiveControl.Name = TxtProductName.Name Or Me.ActiveControl.Name = TxtUnderQty.Name Or Me.ActiveControl.Name = Grid.Name) Then
'    Call ClearDetailArea
  ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSave.Enabled Then BtnSave_Click
            KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
         Case vbKeyO
            If BtnOpen.Enabled Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
      End Select
  ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
'         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtProductName.SetFocus
'         Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus
'         Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtSubGroupID.SetFocus
'         Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then TxtVendorID.SetFocus
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, True) = True Then TxtRemarks.SetFocus
      End Select
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If BtnSave.Enabled Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then FormStatus = changeMode
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Customer Demand"
   DtpDemandDate.DateValue = Date
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
  On Error GoTo ErrorHandler
  FunGetMaxID = cn.Execute("Select isnull(max(DemandID),0) from CustomerDemands").Fields(0) + 1
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub SubClearFields()
  On Error GoTo ErrorHandler
  Dim ctl As Control
  For Each ctl In Me.Controls
    If TypeOf ctl Is TextBox Or TypeOf ctl Is SITextBox.txt Then
      ctl.Text = ""
    ElseIf TypeOf ctl Is ComboBox Then
    
    End If
  Next
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Are you sure to close without save?", vbQuestion + vbApplicationModal + vbYesNo, "Alert") = vbNo Then
         Cancel = 1
      End If
   Else
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set RsReport = Nothing
    Set FrmCustomerDemand = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function FunSelectVender(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
'        SchAccounts.CmbFilter = "Vendors"
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectVender = False: Exit Function
        TxtVenderID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select c.AccountNo, c.AccountName as AccountName, Address, City" & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " where c.AccountNo = '" & (TxtVenderID.Text) & "' and (c.AccountNo like '6%') and isDetailed = 1 and isLocked = 0"
    
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtVenderName.Text = !AccountName
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtPreviousPayable.Text = cn.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtVenderID.Text & "','" & DtpDemandDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
          vStrSQL = " Select isnull(Sum(TotalAmount - isnull(BillDisc,0) + isnull(OtherCharges,0)),0) as Amount " & vbCrLf _
                  + " FROM PurchaseHeader h INNER JOIN (Select PurId, PurchaseDate, Sum(amount) TTLValue FROM PurchaseBody Group By PurId, PurchaseDate)B " & vbCrLf _
                  + " ON h.PurId = B.PurId and h.PurchaseDate = B.PurchaseDate " & vbCrLf _
                  + " where VendorID = '" & (TxtVenderID.Text) & "' and h.PurchaseDate = '" & DtpDemandDate.DateValue & "' and h.PurID >= " & Val(TxtDemandID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
          TxtPreviousPayable.Text = TxtPreviousPayable.Text - cn.Execute(vStrSQL).Fields(0).Value
          lblPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Previous Receivable", "Previous Payable")
          TxtPreviousPayable.Text = Abs(TxtPreviousPayable.Text)
          FunSelectVender = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = changeMode
          Exit Function
      Else
          FunSelectVender = False
          .Close
          TxtVenderID.Text = ""
          TxtVenderName.Text = ""
          TxtAddress.Text = ""
          TxtCity.Text = ""
          TxtPreviousPayable.Text = ""
          lblPayable.Caption = "Previous Payable"
          LblTtlPayable.Caption = "Total Payable"
          If BtnSave.Enabled = False Then FormStatus = changeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function


