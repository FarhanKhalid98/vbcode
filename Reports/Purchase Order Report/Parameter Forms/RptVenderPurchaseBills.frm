VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptVenderPurchaseBills 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptVenderPurchaseBills.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkPaidBills 
      BackColor       =   &H00FF8080&
      Caption         =   "Paid Bills"
      Height          =   255
      Left            =   6233
      TabIndex        =   5
      Top             =   5948
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox ChkUnPaidBills 
      BackColor       =   &H00FF8080&
      Caption         =   "UnPaid Bills"
      Height          =   255
      Left            =   7793
      TabIndex        =   6
      Top             =   5948
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.OptionButton OptFromToDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBD0AB&
      Caption         =   "OptSummary"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4740
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
      Left            =   4740
      TabIndex        =   0
      Top             =   3728
      Width           =   210
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8265
      TabIndex        =   9
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
      MICON           =   "RptVenderPurchaseBills.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5490
      TabIndex        =   7
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
      MICON           =   "RptVenderPurchaseBills.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6870
      TabIndex        =   8
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
      MICON           =   "RptVenderPurchaseBills.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7245
      TabIndex        =   2
      Top             =   4238
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   133627907
      CurrentDate     =   38244
   End
   Begin MSComCtl2.DTPicker DtpTo 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8820
      TabIndex        =   3
      Top             =   4238
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   133627907
      CurrentDate     =   38244
   End
   Begin SITextBox.Txt TxtVendorID 
      Height          =   315
      Left            =   4665
      TabIndex        =   4
      Top             =   5123
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
   Begin SITextBox.Txt TxtVendorName 
      Height          =   315
      Left            =   6345
      TabIndex        =   15
      Top             =   5123
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
   Begin JeweledBut.JeweledButton BtnVendor 
      Height          =   330
      Left            =   5985
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5108
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
      MICON           =   "RptVenderPurchaseBills.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
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
      Left            =   6375
      TabIndex        =   18
      Top             =   4898
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor ID"
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
      Left            =   4665
      TabIndex        =   17
      Top             =   4898
      Width           =   870
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
      Left            =   2700
      TabIndex        =   14
      Top             =   270
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
      Left            =   5115
      TabIndex        =   13
      Top             =   3728
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
      Left            =   5115
      TabIndex        =   12
      Top             =   4238
      Width           =   1785
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   195
      Left            =   7245
      TabIndex        =   11
      Top             =   4043
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   195
      Left            =   8820
      TabIndex        =   10
      Top             =   4043
      Visible         =   0   'False
      Width           =   195
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
Attribute VB_Name = "RptVenderPurchaseBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim sSql As String, vDate As String
Dim RsReport As New ADODB.Recordset

Private Sub BtnVendor_Click()
   If FunSelectVendor(ssButton, False) = True Then
      BtnPreview.SetFocus
   Else
      TxtVendorID.SetFocus
   End If
End Sub

Private Sub TxtVendorID_Change()
   If TxtVendorID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtVendorID.Name Then Exit Sub
   If TxtVendorName.Text <> "All Venders" Then
      TxtVendorName.Text = "All Venders"
   End If
End Sub

Private Sub TxtVendorID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtVendorID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtVendorName.Text <> "All Venders" Then Exit Sub
   If Trim(TxtVendorID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVendor(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVendor(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectVendor(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchVendor.Show vbModal, Me
        If SchVendor.ParaOutVendorID = "" Then FunSelectVendor = False: Exit Function
        TxtVendorID.Text = SchVendor.ParaOutVendorID
    End If
    '---------------------------
    If Trim(TxtVendorID.Text) = "" Then Exit Function
    vStrSQL = " Select * FROM Parties where PartyID = '" & TxtVendorID.Text & "' AND PartyType <> 'C'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtVendorName.Text = !PartyName
          FunSelectVendor = True
          .Close
          Exit Function
      Else
          FunSelectVendor = False
          .Close
          TxtVendorID.Text = ""
          TxtVendorName.Text = "All Venders"
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
   vDate = " and h.purchasedate BETWEEN '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'"
End Sub

Private Sub DtpTo_Change()
   vDate = " and h.purchasedate BETWEEN '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'"
End Sub

Private Sub Form_Load()
   ShowPicture Me, 2
   TxtVendorName.Text = "All Venders"
   SetWindowText Me.hWnd, "Vender Purchase Bills"
   DtpFrom.Value = Date - 30
   DtpTo.Value = Date
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
   Set RptVenderPurchaseBills = Nothing
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
   vParameter = IIf(ChkPaidBills.Value = 0, IIf(ChkUnPaidBills.Value = 1, " and (totalamount - IsNull(billdisc, 0) + isnull(OtherCharges,0)) - (isnull(paidamount,0) + isnull(amount,0)+isnull(discount,0)) > 2", " and 1 = 2"), IIf(ChkUnPaidBills.Value = 1, "", " and (totalamount - IsNull(billdisc, 0) + isnull(OtherCharges,0)) - (isnull(paidamount,0) + isnull(amount,0)+isnull(discount,0)) <= 2"))
   vSQL = " select h.purid, h.purchasedate, totalamount - isnull(billdisc,0) + isnull(OtherCharges,0) as netamount, " & vbCrLf _
      + " isnull(paidamount,0) + isnull(amount,0) + isnull(discount,0) as PaidAmount, partyname, " & vbCrLf _
      + " case when LastPaidDate is not null then LastPaidDate when isnull(paidamount,0) <> 0 then h.purchasedate end as LastPaidDate," & vbCrLf _
      + " (totalamount - IsNull(billdisc, 0) + isnull(OtherCharges,0)) - (isnull(paidamount,0) + isnull(amount,0)+isnull(discount,0)) bal" & vbCrLf _
      + " from purchaseheader h left outer join " & vbCrLf _
      + " (select purid, purchasedate, max(purchasedate) as LastPaidDate, sum(amount) as amount, sum(discount) as Discount from PaymentInvoice group by purid, purchasedate)i " & vbCrLf _
      + " on i.purid = h.purid and i.purchasedate = h.purchasedate" & vbCrLf _
      + " inner join (select PurID, PurchaseDate from PurchaseBody Group By PurID, PurchaseDate)b on h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate" & vbCrLf _
      + " inner join parties p on h.vendorid = p.partyid" & vbCrLf _
      + " where 1=1 " & vParameter & vDate & IIf(TxtVendorID.Text = "", "", " and h.vendorid = '" & TxtVendorID.Text & "'") & vbCrLf _
      + " order by h.Purchasedate, h.Purid"
   Set RsReport = CN.Execute(vSQL)
   If RsReport.BOF Then
      MsgBox "No record exists.", vbInformation, Me.Caption
      Me.MousePointer = vbDefault
      Exit Function
   End If
   Set RptReportViewer.Report = New CrpVenderPurchaseBills
   
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(TxtVendorName.Text = "All Venders", "All Venders", "Vender Name : " & TxtVendorName.Text)
   RptReportViewer.Report.ParameterFields(6).AddCurrentValue IIf(ChkPaidBills.Value = 0, IIf(ChkUnPaidBills.Value = 1, "Unpaid Bills", ""), IIf(ChkUnPaidBills.Value = 1, "Paid and Unpaid Bills.", "Paid Bills."))
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
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
   vDate = " and h.purchasedate BETWEEN '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'"
End Sub
