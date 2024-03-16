VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmBankChqReceive 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11985
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   Icon            =   "FrmBankChqReceive.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmBankChqReceive.frx":0ECA
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   StartUpPosition =   2  'CenterScreen
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   3450
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   35
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
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   9240
      TabIndex        =   6
      Top             =   2460
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
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
   End
   Begin SITextBox.Txt TxtRCVName 
      Height          =   315
      Left            =   3000
      TabIndex        =   15
      Top             =   2460
      Width           =   3015
      _ExtentX        =   5318
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
   Begin JeweledBut.JeweledButton btnAccount 
      Height          =   315
      Left            =   2625
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2460
      Width           =   375
      _ExtentX        =   661
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
      MICON           =   "FrmBankChqReceive.frx":783B
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnClose 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8655
      TabIndex        =   13
      Top             =   8040
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
      MICON           =   "FrmBankChqReceive.frx":7857
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6015
      TabIndex        =   8
      Top             =   8040
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
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
      MICON           =   "FrmBankChqReceive.frx":7873
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4695
      TabIndex        =   9
      Top             =   8040
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
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
      MICON           =   "FrmBankChqReceive.frx":788F
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3375
      TabIndex        =   10
      Top             =   8040
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
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
      MICON           =   "FrmBankChqReceive.frx":78AB
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtchequeNo 
      Height          =   315
      Left            =   6015
      TabIndex        =   4
      Top             =   2460
      Width           =   1830
      _ExtentX        =   3228
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtRCVID 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   2460
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   8
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
   Begin JeweledBut.JeweledButton btndelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7335
      TabIndex        =   12
      Top             =   8040
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
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
      MICON           =   "FrmBankChqReceive.frx":78C7
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtVochID 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      Locked          =   -1  'True
      MaxLength       =   8
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
   Begin JeweledBut.JeweledButton btnPrint 
      Height          =   420
      Left            =   2055
      TabIndex        =   11
      Top             =   8040
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
      MICON           =   "FrmBankChqReceive.frx":78E3
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      CausesValidation=   0   'False
      Height          =   4665
      Left            =   1200
      TabIndex        =   7
      Top             =   2775
      Width           =   9735
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   5
      stylesets.count =   1
      stylesets(0).Name=   "style"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "FrmBankChqReceive.frx":78FF
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   2
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ActiveRowStyleSet=   "style"
      Columns.Count   =   5
      Columns(0).Width=   3200
      Columns(0).Caption=   "Receiving ID"
      Columns(0).Name =   "ReceivingID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5292
      Columns(1).Caption=   "Receiving Name"
      Columns(1).Name =   "ReceiveBy"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Cheque No"
      Columns(2).Name =   "ChequeNo"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2461
      Columns(3).Caption=   "Cheque Date"
      Columns(3).Name =   "ChequeDate"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   7
      Columns(3).NumberFormat=   "dd/MM/yyyy"
      Columns(3).FieldLen=   256
      Columns(4).Width=   2461
      Columns(4).Caption=   "Amount"
      Columns(4).Name =   "Amount"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   17171
      _ExtentY        =   8229
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSCalendarWidgets_A.SSDateCombo dtpVochDate 
      Height          =   315
      Left            =   2145
      TabIndex        =   1
      Top             =   1320
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
   Begin SSCalendarWidgets_A.SSDateCombo dtpChequedate 
      Height          =   315
      Left            =   7845
      TabIndex        =   5
      Top             =   2460
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Left            =   3450
      TabIndex        =   24
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   255
      Left            =   9240
      TabIndex        =   23
      Top             =   2220
      Width           =   855
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Cheque Receive"
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
      Left            =   1920
      TabIndex        =   22
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Date"
      Height          =   195
      Left            =   7845
      TabIndex        =   21
      Top             =   2220
      Width           =   945
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving Name"
      Height          =   195
      Left            =   3000
      TabIndex        =   20
      Top             =   2250
      Width           =   1185
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving ID"
      Height          =   195
      Left            =   1200
      TabIndex        =   19
      Top             =   2250
      Width           =   930
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   11610
      Top             =   60
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Date"
      Height          =   195
      Left            =   2145
      TabIndex        =   18
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher ID"
      Height          =   195
      Left            =   1200
      TabIndex        =   17
      Top             =   1080
      Width           =   810
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No"
      Height          =   195
      Left            =   6015
      TabIndex        =   16
      Top             =   2220
      Width           =   810
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmBankChqReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSql As String
Dim vStrSQL As String
Dim vIsNewRecord As Boolean
Dim vIsNewRow As Boolean
Dim vCounter As Integer
Dim Flag As Boolean
Dim vMode As FormMode
Dim RsReport As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset

Private Sub btnClear_Click()
   FormStatus = NewMode
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub btndelete_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   CN.BeginTrans
   Grid.Redraw = False
   Grid.RemoveAll
   CN.Execute "Delete from BankChqRCVBody where VoucherID = " & Val(TxtVoucherID.Text)
   CN.Execute "Delete from BankChqRCVHeader where VoucherID = " & Val(TxtVoucherID.Text)
   CN.CommitTrans
   Grid.Redraw = True
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   vStrSQL = "Select H.*, B.* from BankChqRCVHeader H inner join BankChqRCVBody B on H.VoucherID = B.VoucherID Where H.VoucherID = " & TxtVoucherID.Text
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set RptReportViewer.Report = New CRptBankChqReceive
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
'        RptReportViewer.Report.ParameterFields(1).AddCurrentValue "XYZ Limited"
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue objRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(objRegistry.CompanyAddress = "", "", objRegistry.CompanyAddress) & IIf(objRegistry.CompanyCity = "", "", ", " & objRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(objRegistry.CompanyPhoneNo = "", ".", " Phone # " & objRegistry.CompanyPhoneNo)
   '   VstrSql = "Select (Select IsNull(Sum(DiscVal),0) from PurchaseBody where PurID =" & txtPurchaseID.Text & ") + (Select IsNull(Sum(Discount),0) from PurchaseHeader where PurID = " & txtPurchaseID.Text & ")"
'   RptReportViewer.Report.ParameterFields(4).AddCurrentValue CStr(CN.Execute(VstrSql).Fields.Item(0).Value)
'   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
'   RptReportViewer.Report.PaperOrientation = crPortrait
   RptReportViewer.Show
   'RptReportViewer.Report.PrintOut False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If Not (UCase(ActiveControl.Name) Like UCase("txt*")) Then Exit Sub
 If btnSave.Enabled = False Then FormStatus = changemode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If btnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
         Set frmObj = Nothing
      Next
         Set RsBody = Nothing
         Set FrmBankChqReceive = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   SetWindowText Me.hwnd, "Bank Cheque Receive"
   FormStatus = NewMode
End Sub

Private Property Get FormStatus() As FormMode
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
  On Error GoTo ErrorHandler
   vMode = vNewValue
   Select Case vNewValue
   Case Is = NewMode
      Call SubClearFields
      If RsBody.State = adStateOpen Then RsBody.Close
      btnOpen.Enabled = True
      btndelete.Enabled = False
      btnSave.Enabled = False
      btnClear.Enabled = True
      btnPrint.Enabled = False
      TxtVoucherID.Text = FunGetMaxID
      PopulateDataToGrid
      If dtpVoucherDate.Enabled And dtpVoucherDate.Visible Then dtpVoucherDate.SetFocus
      vIsNewRecord = True
      vIsNewRow = True
   Case Is = OpenMode
      btnOpen.Enabled = True
      btndelete.Enabled = True
      btnClear.Enabled = True
      btnSave.Enabled = False
      btnPrint.Enabled = True
      vIsNewRecord = False
      vIsNewRow = True
   Case Is = changemode
      btnOpen.Enabled = False
      btndelete.Enabled = False
      btnSave.Enabled = True
      btnPrint.Enabled = False
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         ctl.Text = ""
      ElseIf TypeOf ctl Is SITextBox.txt Then
         ctl.Text = ""
      End If
   Next
   dtpVoucherDate.DateValue = Date
   dtpChequedate.DateValue = Date
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ChequeNo").Text = " "
   Grid.Update
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = CN.Execute("Select isnull(max(VoucherID),0) from BankChqRCVHeader").Fields(0) + 1
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "select * from BankChqRCVBody where VoucherID = ' " & Val(TxtVoucherID.Text) & " ' ", CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = "Select B.ReceivingID, ReceivingName, B.ActChequeNo, B.ActChequeDate, B.ActAmount From BankChqRCVBody B  where B.VoucherID =" & Val(TxtVoucherID.Text)
      With CN.Execute(sSql)
         If .RecordCount > 0 Then
            Grid.Redraw = False
            Grid.MoveFirst
            Grid.RemoveAll
            Grid.AllowAddNew = True
            While Not .EOF
               Grid.AddNew
               Grid.Columns("ReceivingID").Text = !ReceivingID
               Grid.Columns("ReceiveBy").Text = IIf(IsNull(!ReceivingName), "", !ReceivingName)
               Grid.Columns("ChequeNo").Text = IIf(IsNull(!ActChequeNo), "", !ActChequeNo)
               Grid.Columns("ChequeDate").Text = (!ActChequeDate)
               Grid.Columns("amount").Value = Val(!ActAmount)
               .MoveNext
            Wend
         End If
         .Close
      End With
      Grid.AddNew
      Grid.Columns("ChequeNo").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
      If KeyCode = vbKeyReturn Then
         If ActiveControl.Name = Grid.Name Then
            Grid_DblClick
         Else
            keybd_event 9, 1, 1, 1
            KeyCode = 0
         End If
      ElseIf KeyCode = vbKeyEscape Then
         Call SubClearDetailArea: TxtRCVID.SetFocus
      ElseIf KeyCode = vbKeyF1 Then
         Select Case ActiveControl.Name
            'Case TxtBankActID.Name: If FunSelectAccount(ssFunctionKey, False) = True Then TxtRCVID.SetFocus
            Case TxtRCVID.Name: If FunSelectPayee(ssFunctionKey, False) = True Then TxtChequeNo.SetFocus
         End Select
      ElseIf Shift = vbCtrlMask Then
         Select Case KeyCode
            Case vbKeyS
               If btnSave.Enabled = True Then btnSave_Click
               KeyCode = 0
            Case vbKeyW
               If btnClear.Enabled = True Then btnClear_Click
               KeyCode = 0
            Case vbKeyQ
               If btnClose.Enabled = True Then BtnClose_Click
               KeyCode = 0
            Case vbKeyO
               If btnOpen.Enabled = True Then BtnOpen_Click
               KeyCode = 0
            Case vbKeyP
               If btnPrint.Enabled = True Then BtnPrint_Click
               KeyCode = 0
            Case vbKeyR
               If btndelete.Enabled = True Then btndelete_Click
               KeyCode = 0
            Case vbKeyDelete
               MniRemoveRow_Click
               KeyCode = 0
         End Select
      ElseIf ActiveControl.Name = TxtRCVID.Name Then
         If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtRCVID.Name Then
         KeyCode = 0
      End If
   End If
   Exit Sub
ErrorHandler:
     Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   Call Grid_LostFocus
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("ChequeNo").Text) = "" Then
      TxtRCVID.Text = ""
      TxtRCVID.Enabled = True
      btnAccount.Enabled = True
      vIsNewRow = True
      If TxtRCVID.Enabled Then TxtRCVID.SetFocus
   Else
      btnAccount.Enabled = False
      TxtRCVID.Enabled = False
      vIsNewRow = False
      If TxtChequeNo.Enabled Then TxtChequeNo.SetFocus
   End If
End Sub

Private Sub Grid_GotFocus()
   Flag = True
   TxtRCVID.Enabled = False
   btnAccount.Enabled = False
End Sub
Private Sub SubClearDetailArea()
   TxtChequeNo.Text = ""
   dtpChequedate.DateValue = Date
   TxtAmount.Text = ""
   TxtRCVName.Text = ""
   TxtRCVID.Text = ""
End Sub

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchBankAct.Show vbModal, Me
      If SchBankAct.ParaOutID = "" Then FunSelectAccount = False: Exit Function
      TxtBankActID.Text = SchBankAct.ParaOutID
   End If
   Dim vStrSQL As String
   vStrSQL = "select * from ChartofAccounts where AccountNo = '" & Val(TxtBankActID.Text) & "'"
   With CN.Execute(vStrSQL)
         If .RecordCount > 0 Then
            TxtBankActName.Text = !AccountName
            .Close
            FunSelectAccount = True
            If btnSave.Enabled = False Then FormStatus = changemode
            Exit Function
         Else
            FunSelectAccount = False
            .Close
            TxtBankActName.Text = ""
            If btnSave.Enabled = False Then FormStatus = changemode
         End If
      End With
      Exit Function
ErrorHandler:
      Call ShowErrorMessage
End Function

Private Function FunSelectPayee(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchAccounts.ParaInWhereClause = "" '" and c.accountno like '6%'"
      SchAccounts.Show vbModal, Me
      If SchAccounts.ParaOutAccountNo = "" Then FunSelectPayee = False: Exit Function
      TxtRCVID.Text = SchAccounts.ParaOutAccountNo
   End If
   Dim vStrSQL As String
   If Trim(TxtRCVID.Text) = "" Then FunSelectPayee = False: Exit Function
   vStrSQL = "select * from ChartofAccounts where AccountNo = '" & (TxtRCVID.Text) & "'"
   With CN.Execute(vStrSQL)
         If .RecordCount > 0 Then
            TxtRCVName.Text = !AccountName
            .Close
            FunSelectPayee = True
            If btnSave.Enabled = False Then FormStatus = changemode
            Exit Function
         Else
            FunSelectPayee = False
            .Close
            TxtRCVName.Text = ""
            If btnSave.Enabled = False Then FormStatus = changemode
         End If
      End With
      Exit Function
ErrorHandler:
      Call ShowErrorMessage
End Function

Private Sub TxtRCVID_Change()
   If ActiveControl.Name <> TxtRCVID.Name Then Exit Sub
   If TxtRCVName.Text <> "" Then
      TxtRCVName.Text = ""
      TxtRCVID.Text = ""
   End If
End Sub

Private Sub TxtRCVID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtRCVID_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtRCVID.Name Then Exit Sub
   If TxtRCVName.Text <> "" Then Exit Sub
   If Trim(TxtRCVID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectPayee(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectPayee(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtBankActID_Change()
If ActiveControl.Name <> TxtBankActID.Name Then Exit Sub
   If TxtBankActName.Text <> "" Then
      TxtBankActID.Text = ""
      TxtBankActName.Text = ""
    End If
    
End Sub

Private Sub TxtBankActID_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtBankActID.Name Then Exit Sub
   If TxtBankActName.Text <> "" Then Exit Sub
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

Private Sub TxtAmount_LostFocus()
   Select Case ActiveControl.Name
   Case TxtRCVID.Name, TxtChequeNo.Name, dtpChequedate.Name
      Exit Sub
   End Select
   Call GetDataFromTextBoxesToGrid
End Sub

Private Sub GetDataFromTextBoxesToGrid()
   On Error GoTo ErrorHandler
   If Trim(TxtChequeNo.Text) = "" Then
      MsgBox " Please Specify ChequeNo ", vbInformation + vbOKOnly, "Error"
      If TxtChequeNo.Enabled = True Then TxtChequeNo.SetFocus
      Exit Sub
   End If
   If TxtAmount.Text = "" Then
      MsgBox " Please Specify Amount ", vbInformation + vbOKOnly, "Error"
      If TxtAmount.Enabled = True Then TxtAmount.SetFocus
      Exit Sub
   End If
   If CN.Execute("Select ActChequeNo from BankChqRCVBody where ActChequeNo = '" & TxtChequeNo.Text & "' and VoucherID <> " & TxtVoucherID.Text).EOF = False Then
      MsgBox "Cheque No. '" & TxtChequeNo.Text & "'  Already Exists in DataBase ", vbInformation + vbOKOnly, "Error"
      If TxtChequeNo.Enabled = True Then TxtChequeNo.SetFocus
      Exit Sub
   End If
   RsBody.Filter = "ActChequeNo = '" & TxtChequeNo.Text & "'"
   If vIsNewRow = True Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ChequeNo").Text = TxtChequeNo.Text
         RsBody!VoucherID = TxtVoucherID.Text
      Else
         MsgBox "Cheque No. '" & TxtChequeNo.Text & "'  Already Exists in DataBase ", vbInformation + vbOKOnly, "Alert"
         RsBody.Filter = 0
         Call SubClearDetailArea
         TxtChequeNo.SetFocus
         Exit Sub
      End If
   Else
      If RsBody.RecordCount <> 0 Then
         If TxtChequeNo.Text <> Grid.Columns("ChequeNO").Text Then
            MsgBox "Cheque No. '" & TxtChequeNo.Text & "'  Already Exists", vbInformation + vbOKOnly, "Error"
            If TxtChequeNo.Enabled = True Then TxtChequeNo.SetFocus
            Exit Sub
         End If
      End If
      RsBody.Filter = "ActChequeNo = '" & Grid.Columns("ChequeNO").Text & "'"
   End If
   With Grid
      .Columns("ReceivingID").Text = TxtRCVID.Text
      .Columns("Amount").Text = Val(TxtAmount.Text)
      .Columns("ChequeNO").Text = TxtChequeNo.Text
      .Columns("ChequeDate").Text = dtpChequedate.DateValue
      .Columns("ReceiveBy").Text = TxtRCVName.Text
      RsBody!ReceivingID = TxtRCVID.Text
      RsBody!ActChequeNo = TxtChequeNo.Text
      RsBody!ActChequeDate = dtpChequedate.DateValue
      RsBody!ActAmount = Val(TxtAmount.Text)
      RsBody!ReceivingName = IIf((TxtRCVName.Text = ""), Null, TxtRCVName.Text)
      .MoveLast
      If Trim(.Columns("ChequeNo").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("ChequeNo").Text = " "
         .AllowAddNew = False
      End If
   End With
   vIsNewRow = True
   Call SubClearDetailArea
   If TxtRCVID.Enabled = True Then TxtRCVID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Grid.Columns("ChequeNo").Text = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub
Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then GetDatabackFromGridToTextBoxes
End Sub
Private Sub MniRemoveRow_Click()
On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ReceivingID").Text) = "" Then Exit Sub
   RsBody.Filter = "ReceivingID = " & Grid.Columns("ReceivingID").Text
   RsBody.Delete
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   RsBody.Filter = 0
   GetDatabackFromGridToTextBoxes
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   FormStatus = changemode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDatabackFromGridToTextBoxes()
   On Error GoTo ErrorHandler
   With Grid
      If Grid.Rows > 0 Then
         'TxtRCVID.Text = .Columns("AcPayeeID").Text
         'TxtRCVName.Text = .Columns("ACPayeeName").Text
         TxtChequeNo.Text = .Columns("ChequeNO").Text
         dtpChequedate.DateValue = IIf(.Columns("ChequeDate").Value = Empty, Date, .Columns("ChequeDate").Value)
         TxtRCVName.Text = .Columns("ReceiveBy").Text
         TxtAmount.Text = .Columns("amount").Text
         TxtRCVID.Text = .Columns("ReceivingID").Text
      End If
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub btnSave_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
'   If TxtBankActID.Text = "" Then
'      MsgBox " Please Specify Bank Account ", vbInformation + vbOKOnly, "Error"
'      If TxtBankActID.Enabled = True Then TxtBankActID.SetFocus
'    Exit Sub
'    End If
'    If TxtRCVID.Text = "" Then
'      MsgBox " Please Specify Account PayeeID ", vbInformation + vbOKOnly, "Error"
'      If TxtRCVID.Enabled = True Then TxtRCVID.SetFocus
'    Exit Sub
'    End If
   If vIsNewRecord Then
      If CN.Execute("Select * from BankChqRCVHeader where VoucherID=" & Val(TxtVoucherID.Text)).RecordCount > 0 Then
         MsgBox "This Voucher ID already exists. A new Voucher ID. has been generated. Please try again", vbCritical, "Alert"
         TxtVoucherID.Text = FunGetMaxID
         Exit Sub
      End If
   End If
   Dim Rs As New ADODB.Recordset
   sSql = "select * from BankChqRCVHeader where VoucherID = " & Val(TxtVoucherID.Text)
   Rs.Open sSql, CN, adOpenStatic, adLockPessimistic
   With Rs
      If .BOF Then
         .AddNew
         Rs!VoucherID = Val(TxtVoucherID.Text)
      End If
      !VoucherDate = dtpVoucherDate.DateValue
      !Description = IIf((TxtDescription.Text = ""), Null, TxtDescription.Text)
      !UserNo = vUser
      .Update
      .Close
   End With
   RsBody.Filter = 0
   RsBody.MoveFirst
   For vCounter = 1 To RsBody.RecordCount
      RsBody!VoucherID = TxtVoucherID.Text
      RsBody.Update
      RsBody.MoveNext
   Next vCounter
   RsBody.UpdateBatch
   RsBody.MoveFirst
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchChqReceive.Show vbModal, Me
   If SchChqReceive.ParaOutVoucherNo <> 0 Then
      TxtVoucherID.Text = SchChqReceive.ParaOutVoucherNo
      GetCompeleteInfo
   End If
   dtpVoucherDate.SetFocus
End Sub

Private Sub GetCompeleteInfo()
   On Error GoTo ErrorHandler
   sSql = "select H.VoucherID, H.VoucherDate, H.Description from BankChqRCVHeader H Left Join BankChqRCVBody B on H.VoucherID = B.VoucherID   where H.VoucherID = " & Val(TxtVoucherID.Text)
   With CN.Execute(sSql)
      If Not .BOF Then
         TxtVoucherID.Text = !VoucherID
         dtpVoucherDate.DateValue = !VoucherDate
         TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
      End If
      .Close
   End With
   PopulateDataToGrid
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub btnAccount_Click()
   If FunSelectPayee(ssButton, False) = True Then
      TxtChequeNo.SetFocus
   Else
      TxtRCVID.SetFocus
   End If
End Sub
