VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmGRN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   15360
   ClipControls    =   0   'False
   Icon            =   "GRN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1170
      TabIndex        =   37
      Top             =   1215
      Width           =   1245
   End
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   1170
      TabIndex        =   35
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   1575
      Width           =   2115
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "GRN.frx":0ECA
      Left            =   1125
      List            =   "GRN.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Tag             =   "1"
      Top             =   2025
      Width           =   3276
   End
   Begin VB.TextBox TxtGRNID 
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
      TabIndex        =   14
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
      MICON           =   "GRN.frx":0ECE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6390
      TabIndex        =   17
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
      MICON           =   "GRN.frx":0EEA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10305
      TabIndex        =   19
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
      MICON           =   "GRN.frx":0F06
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5085
      TabIndex        =   15
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
      MICON           =   "GRN.frx":0F22
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9000
      TabIndex        =   18
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
      MICON           =   "GRN.frx":0F3E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   3780
      TabIndex        =   16
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
      MICON           =   "GRN.frx":0F5A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtRemarks 
      Height          =   315
      Left            =   5048
      TabIndex        =   13
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
      Left            =   2190
      TabIndex        =   5
      Top             =   4410
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
      Left            =   3480
      TabIndex        =   7
      Top             =   4410
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
      Left            =   7125
      TabIndex        =   8
      Top             =   4410
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
      Left            =   11655
      TabIndex        =   9
      Top             =   4410
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
      Left            =   3120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4410
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
      MICON           =   "GRN.frx":0F76
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotalPayable 
      Height          =   315
      Left            =   9053
      TabIndex        =   12
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpGRNDate 
      Height          =   315
      Left            =   5784
      TabIndex        =   1
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
      TabIndex        =   2
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      MICON           =   "GRN.frx":0F92
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtPreviousPayable 
      Height          =   315
      Left            =   7050
      TabIndex        =   11
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
   Begin SITextBox.Txt TxtGRNAmount 
      Height          =   315
      Left            =   5048
      TabIndex        =   10
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
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Print Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   36
      Top             =   1665
      Width           =   840
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printer"
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
      Left            =   180
      TabIndex        =   33
      Top             =   2070
      Width           =   570
   End
   Begin VB.Label LblNetAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "GRN Amount"
      Height          =   195
      Left            =   5048
      TabIndex        =   32
      Top             =   5640
      Width           =   945
   End
   Begin VB.Label LblTtlPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payable"
      Height          =   195
      Left            =   9053
      TabIndex        =   31
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Payable"
      Height          =   195
      Left            =   7050
      TabIndex        =   30
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
      TabIndex        =   29
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
      TabIndex        =   28
      Top             =   2685
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender ID"
      Height          =   195
      Left            =   2175
      TabIndex        =   27
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Name"
      Height          =   195
      Left            =   3480
      TabIndex        =   26
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   7125
      TabIndex        =   25
      Top             =   4200
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   11655
      TabIndex        =   24
      Top             =   4230
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GRN Date"
      Height          =   195
      Left            =   5790
      TabIndex        =   23
      Top             =   2685
      Width           =   750
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   5055
      TabIndex        =   22
      Top             =   6690
      Width           =   630
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goods Received Notes (GRN)"
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
      TabIndex        =   21
      Top             =   270
      Width           =   3870
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
      Height          =   195
      Left            =   4575
      TabIndex        =   20
      Top             =   2670
      Width           =   570
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove this Row"
      End
   End
End
Attribute VB_Name = "FrmGRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vCounter As Integer
Dim Flag As Boolean
Dim ssql As String
Dim vStrSQL As String
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vStrComp As String, vCompanyName As String, vAddress As String, vemail As String
Dim vPrinter() As String
'----------------------------------

Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
      TxtRemarks.SetFocus
   Else
      TxtVenderID.SetFocus
   End If
End Sub

Private Sub TxtGRNAmount_Change()
On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtGRNAmount.Name Then Exit Sub
   Call SubCalculateFooter
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
   
   vStrSQL = " Select G.*, OrganizationName, PartyName, Address, City from GRN G Left Outer Join Organizations O on O.Organizationid = G.Organizationid" & vbCrLf _
            + " Left outer Join Parties p on G.VendorID = p.PartyID" & vbCrLf _
            + " where g.GRNID = " & Val(TxtGRNID.Text)

   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
  
  If cmbPrintType.Text = "Half Page" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptGRNInvoiceHalf.rpt")
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.x
      RptReportViewer.Report.RightMargin = 225
   ElseIf cmbPrintType.Text = "Thermal" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptGRNInvoiceAurora.rpt")
      RptReportViewer.Report.TopMargin = 0
      RptReportViewer.Report.LeftMargin = 0
      RptReportViewer.Report.RightMargin = 0
   Else
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptGRNInvoice.rpt")
   End If
'   Set RptReportViewer.Report = New CrptCustomerGRN
   
   
   RptReportViewer.Report.ReportTitle = "Goods Report Notes"
   
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
'   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
''   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
''   RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
   
   vPrinter = Split(CmbPrinters.Text, ",")
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   
   If ChkIsPreview.Value = 1 Then
      RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False
   End If
   
   cn.Execute ("Insert Into UserActivities values ('GRN Invoice'" & "," & TxtGRNID.Text & ",'" & DtpGRNDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
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
  cn.Execute "Delete from GRN Where GRNID = " & Val(TxtGRNID.Text)
  cn.CommitTrans
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  If cn.Errors.Count > 0 Then cn.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Sub GETGRN()
   On Error GoTo ErrorHandler
   ssql = "Select G.*, OrganizationName, PartyName, Address, City from GRN G Left outer Join Organizations O on o.OrganizationID= G.OrganizationID Left outer Join Parties p on G.VendorID = p.PartyID where GRNID = " & Val(TxtGRNID.Text)
   With cn.Execute(ssql)
      If Not .BOF Then
          DtpGRNDate.DateValue = !GRNDate
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          TxtVenderID.Text = IIf(IsNull(!vendorID), "", !vendorID)
          TxtVenderName.Text = IIf(IsNull(!PartyName), "", !PartyName)
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtGRNAmount.Text = !GRNAmount
          TxtPreviousPayable.Text = IIf(IsNull(!PreviousAmount), "", !PreviousAmount)
          lblPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Previous Receivable", "Previous Payable")
          LblTtlPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Total Receivable", "Total Payable")
          TxtPreviousPayable.Text = Abs(Val(TxtPreviousPayable.Text))
          TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
          Call SubCalculateFooter
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
   SchGRN.Show vbModal
   If SchGRN.ParaOutGRNID <> "" Then
      TxtGRNID.Text = SchGRN.ParaOutGRNID
      GETGRN
   End If
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
      If cn.Execute("Select * from GRN where GRNID = " & Val(TxtGRNID.Text)).RecordCount > 0 Then
         MsgBox "This voucher already exists. A new voucher No. has been generated. Please try again", vbCritical, "Alert"
         TxtGRNID.Text = FunGetMaxID
         Exit Sub
      End If
   End If
'   RsBody.Filter = 0
   If Val(TxtGRNAmount.Text) = 0 Then
      MsgBox "GRN Amount cannot be Zero", vbInformation, "Error"
      TxtGRNAmount.SetFocus
      Exit Sub
   End If
   
   
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
  
   cn.BeginTrans
   ssql = "Select * From GRN Where GRNID =" & Val(TxtGRNID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, cn, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !GRNID = Val(TxtGRNID.Text)
      End If
      !GRNDate = DtpGRNDate.DateValue
      !vendorID = TxtVenderID.Text
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !GRNAmount = Round(Val(TxtGRNAmount.Text))
      !PreviousAmount = IIf(lblPayable.Caption = "Previous Receivable", Val(TxtPreviousPayable.Text), Val(TxtPreviousPayable.Text) * -1)
      !Remarks = IIf(TxtRemarks.Text = "", Null, TxtRemarks.Text)
      !UserNo = vUser
      .Update
      .Close
   End With
   cn.CommitTrans
   
   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'GRN' and LocalComputerName = '" & LocalComputerName & "'"
   If cn.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview ) Values ('" & LocalComputerName & "', 'GRN','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & " Where FormType = 'GRN' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   cn.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   
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
      TxtGRNID.Text = FunGetMaxID
      If DtpGRNDate.Enabled And DtpGRNDate.Visible Then DtpGRNDate.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      BtnPrint.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      DtpGRNDate.SetFocus
      vIsNewRecord = False
    Case Is = ChangeMode
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
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then If TxtVenderID.Enabled Then TxtVenderID.SetFocus Else TxtOrganizationID.SetFocus
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, True) = True Then TxtRemarks.SetFocus
      End Select
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If BtnSave.Enabled Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then FormStatus = ChangeMode
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Goods Received Notes"
   
   DtpGRNDate.DateValue = Date
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   cmbPrintType.Clear
   cmbPrintType.AddItem "Full Page"
   cmbPrintType.AddItem "Half Page"
   cmbPrintType.AddItem "Thermal"
   cmbPrintType.ListIndex = 0
   
   CmbPrinters.Clear
   CmbPrinters.AddItem "Default,winspool,LPT1"
   Dim p
   For Each p In Printers
      CmbPrinters.AddItem p.DeviceName & "," & p.DriverName & "," & p.Port
   Next p
   CmbPrinters.ListIndex = 0
   
   
   '''''''''''''''' Form Default Setting  ''''''''''''''''''''''
   ssql = "select * from FormDefaultSetting Where FormType = 'GRN' and LocalComputerName = '" & LocalComputerName & "'"
   With cn.Execute(ssql)
     If .RecordCount > 0 Then
        cmbPrintType.Text = !Size
        ChkIsPreview.Value = Abs(!IsPreview)
        If Not IsNull(!DeviceName) Then
            CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
        Else
            CmbPrinters.ListIndex = 0
        End If
     End If
     .Close
   End With
   ''''''''''''''''''''''''''''''''''''''''''''''
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
  On Error GoTo ErrorHandler
  If ObjRegistry.AllowContinuousBillNo = True Then
      FunGetMaxID = cn.Execute("Select isnull(max(GRNID),0)+1 from GRN").Fields(0)
   ElseIf ObjRegistry.AllowMonthlyBillNo = True Then
      FunGetMaxID = cn.Execute("Select isnull(max(GRNID),0)+1 from GRN where Month(GRNDATE) = '" & Month(DtpGRNDate.DateValue) & "' and  year(GRNDATE) ='" & Year(DtpGRNDate.DateValue) & "'").Fields(0)
   ElseIf ObjRegistry.AllowDailyBillNo = True Then
      FunGetMaxID = cn.Execute("Select isnull(max(GRNID),0)+1 from GRN where GRNDATE = '" & DtpGRNDate.DateValue & "'").Fields(0)
   Else
      FunGetMaxID = cn.Execute("Select isnull(max(GRNID),0)+1 from GRN").Fields(0)
   End If
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
    Set FrmGRN = Nothing
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
          TxtPreviousPayable.Text = cn.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtVenderID.Text & "','" & DtpGRNDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
          vStrSQL = " Select isnull(Sum(TotalAmount - isnull(BillDisc,0) + isnull(OtherCharges,0)),0) as Amount " & vbCrLf _
                  + " FROM PurchaseHeader h INNER JOIN (Select PurId, PurchaseDate, Sum(amount) TTLValue FROM PurchaseBody Group By PurId, PurchaseDate)B " & vbCrLf _
                  + " ON h.PurId = B.PurId and h.PurchaseDate = B.PurchaseDate " & vbCrLf _
                  + " where VendorID = '" & (TxtVenderID.Text) & "' and h.PurchaseDate = '" & DtpGRNDate.DateValue & "' and h.PurID >= " & Val(TxtGRNID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
          TxtPreviousPayable.Text = TxtPreviousPayable.Text - cn.Execute(vStrSQL).Fields(0).Value
          lblPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Previous Receivable", "Previous Payable")
          TxtPreviousPayable.Text = Abs(TxtPreviousPayable.Text)
          FunSelectVender = True
          .Close
          Call SubCalculateFooter
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectVender = False
          .Close
          TxtVenderID.Text = ""
          TxtVenderName.Text = ""
          TxtAddress.Text = ""
          TxtCity.Text = ""
          TxtPreviousPayable.Text = ""
          TxtTotalPayable.Text = ""
          TxtGRNAmount.Text = ""
          lblPayable.Caption = "Previous Payable"
          LblTtlPayable.Caption = "Total Payable"
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function


Private Sub TxtOrganizationID_Change()
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizationName.Text <> "" Then TxtOrganizationName.Text = ""
End Sub

Private Sub TxtOrganizationID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtOrganizationID.Text) = "" Then Exit Sub
   If TxtOrganizationName.Text <> "" Then Exit Sub
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

Private Sub BtnOrganization_Click()
   If FunSelectOrganization(ssButton, False) = True Then
      TxtVenderID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub
Private Function FunSelectOrganization(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganization = False: Exit Function
        TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Organizations where OrganizationID=" & Val(TxtOrganizationID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizationName.Text = !OrganizationName
          FunSelectOrganization = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectOrganization = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizationName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub SubCalculateFooter()
   On Error GoTo ErrorHandler
   TxtTotalPayable.Text = Abs(Val(TxtGRNAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousPayable.Text, Val(TxtPreviousPayable.Text) * -1)))
   LblTtlPayable.Caption = IIf(Val(TxtGRNAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousPayable.Text, Val(TxtPreviousPayable.Text) * -1)) < 0, "Total Receivable", "Total Payable")
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
