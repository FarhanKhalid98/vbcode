VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptBatchPrint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptBatchPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   2460
      TabIndex        =   23
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   5955
      Width           =   2115
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "RptBatchPrint.frx":0ECA
      Left            =   1305
      List            =   "RptBatchPrint.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "1"
      Top             =   6315
      Width           =   3276
   End
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2445
      TabIndex        =   21
      Top             =   5445
      Width           =   1245
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7755
      TabIndex        =   8
      Top             =   7905
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
      MICON           =   "RptBatchPrint.frx":0ECE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   6315
      TabIndex        =   6
      Top             =   7905
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
      MICON           =   "RptBatchPrint.frx":0EEA
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   6146
      TabIndex        =   4
      Top             =   6038
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
      Left            =   7901
      TabIndex        =   5
      Top             =   6038
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
   Begin JeweledBut.JeweledButton BtnOrganizaton 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6214
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4043
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
      MICON           =   "RptBatchPrint.frx":0F06
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   5194
      TabIndex        =   0
      Top             =   4043
      Width           =   1020
      _ExtentX        =   1799
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtOrganizatonName 
      Height          =   315
      Left            =   6574
      TabIndex        =   12
      Tag             =   "nc"
      Top             =   4043
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
   Begin JeweledBut.JeweledButton BtnSector 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6221
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4718
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
      MICON           =   "RptBatchPrint.frx":0F22
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSectorID 
      Height          =   315
      Left            =   5201
      TabIndex        =   1
      Top             =   4718
      Width           =   1020
      _ExtentX        =   1799
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtSectorName 
      Height          =   315
      Left            =   6581
      TabIndex        =   16
      Tag             =   "nc"
      Top             =   4718
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
   Begin SITextBox.Txt TxtBillIDFrom 
      Height          =   315
      Left            =   6896
      TabIndex        =   2
      Top             =   5243
      Width           =   1230
      _ExtentX        =   2170
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
   End
   Begin SITextBox.Txt TxtBillIIDTo 
      Height          =   315
      Left            =   8936
      TabIndex        =   3
      Top             =   5243
      Width           =   1230
      _ExtentX        =   2170
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
      Left            =   2460
      TabIndex        =   25
      Top             =   5655
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
      Left            =   630
      TabIndex        =   24
      Top             =   6360
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   8291
      TabIndex        =   20
      Top             =   5288
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Range From"
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
      Left            =   5201
      TabIndex        =   19
      Top             =   5288
      Width           =   1470
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector ID"
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
      Left            =   5201
      TabIndex        =   18
      Top             =   4523
      Width           =   825
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector Name"
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
      Left            =   6581
      TabIndex        =   17
      Top             =   4523
      Width           =   1110
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
      Left            =   6574
      TabIndex        =   14
      Top             =   3848
      Width           =   1620
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
      Left            =   5194
      TabIndex        =   13
      Top             =   3848
      Width           =   1335
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale invoice Batch Print"
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
      Left            =   1980
      TabIndex        =   10
      Top             =   180
      Width           =   3120
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
      Left            =   6146
      TabIndex        =   9
      Top             =   5813
      Width           =   885
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
      Left            =   7916
      TabIndex        =   7
      Top             =   5813
      Width           =   705
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
Attribute VB_Name = "RptBatchPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim isUrdu As Boolean
Dim ssql, vStrSQL As String
Dim RsReport As New ADODB.Recordset
Dim Application1 As New CRAXDRT.Application
Dim vBillID As Integer
Dim vBillDate As String

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnOrganizaton_Click()
   If FunSelectOrganization(ssButton, False) = True Then
      If TxtSectorID.Visible Then TxtSectorID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
      RptReportViewer.Caption = Me.Caption
      RptReportViewer.Show
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   If ObjRegistry.AllowUrduProduct = True Then
'      If MsgBox("Do you want to print this invoice in Urdu", vbQuestion + vbYesNo, "Alert") = vbYes Then
'         isUrdu = True
'      Else
'         isUrdu = False
'      End If
   End If
   ssql = ""
   ssql = ssql + "  select h.BillID, h.BillDate, o.OrganizationID, ReportName, SaleReportName" & vbCrLf _
               + " from SaleHeader h " & vbCrLf _
               + " left outer join Organizations o on o.OrganizationID = h.OrganizationID" & vbCrLf _
               + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
               + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
               + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
               + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
               + " where h.BillDate between '" & DtpFrom.DateValue & "' and  '" & DtpTo.DateValue & "'"
   
   If Val(TxtBillIDFrom.Text) <> 0 And Val(TxtBillIIDTo.Text) <> 0 Then
   ssql = ssql + " and  h.BillID between " & Val(TxtBillIDFrom.Text) & " and " & Val(TxtBillIIDTo.Text)
   End If
   If Val(TxtOrganizationID.Text) <> 0 Then
   ssql = ssql + " and h.organizationid = " & TxtOrganizationID.Text
   End If
   If Val(TxtSectorID.Text) <> 0 Then
   ssql = ssql + " and pr.sectorid = " & TxtSectorID.Text
   End If
   ssql = ssql + " Order by H.SID"
   
   
   With cn.Execute(ssql)
   While Not .EOF
   
      vStrSQL = "Select h.BillID, h.BillDate, h.StoreID, UserName, ExpiryInvoice, BillTIme as EntryTime, EntryDate, h.PromiseDate, h.OrganizationID, OrganizationName, Customerid, cast(H.CustomerID as varchar(10))  + ' - ' + isnull(Pr.PartyName,AccountName)  as Customer_Name_ID," & vbCrLf _
                + " pr.address, LicenceNo, SectorName, ZoneName, H.StoreID, StoreName, BiltyNo, VehicleNo, h.Description, h.Remarks," & vbCrLf _
                + " Isnull(H.BillDiscPer, 0) BillDiscPer, Isnull(H.BillDisc,0) BillDisc, isnull(OtherCharges,0) as OtherCharges,  isnull(h.ServiceCharges,0) as ServiceCharges," & vbCrLf _
                + " TotalAmount,  isnull(TotalExpense,0) as TotalExpense,  CompanyName, " & IIf(ObjRegistry.AllowUrduProduct = False, "GroupName", "GroupName1") & "  as GroupName, SubGroupName, BrandName, SeasonName, b.ProductID as Code, " & IIf(ObjRegistry.AllowUrduProduct = False, "p.ProductName", "isnull(p.ProductName1,p.ProductName)") & "  as ProductName, dbo.FunSaleBodySerial(b.BillID,b.BillDate, b.ProductId) Serial," & vbCrLf _
                + " dbo.FunSaleBodyOffer(b.BillID,b.BillDate,b.ProductId) ProductOffer, isnull(QtyPack,0)QtyPack, isnull(b.Multiplier,0)Multiplier, isnull(b.GrossQty,0)GrossQty, isnull(b.GrossUnit,0)GrossUnit, Qty," & vbCrLf _
                + " P.RetailPrice, P.PurPrice, Bonus, b.DiscPc, b.DiscPer, DiscVal, Offer, Cast(b.Tradeoffer1 as varchar(5)) + ' + ' + cast(b.tradeoffer2 as varchar(5)) TradeOffer_12, tradevalue, Extraschemevalue, b.ExtraSchemePer," & vbCrLf _
                + " b.SaleTaxPer, SaleTaxval, AdvTaxVal, AdvTaxPer, ExtraTaxVal, ExtraTaxPer, h.CNIC, h.MobileNo,  b.SC, h.Empid, empname, price, Amount, previousAmount, CashReceived, isnull(BatchNo,'') as BatchNo, BillNo," & vbCrLf _
                + " Abbreviation + '/' + cast(b.Multiplier as varchar(10)) as packing, isnull(P.ListPrice,0) as ListPrice," & vbCrLf _
                + " isnull( pr.Phone1  + ', ','') + isnull( pr.Phone2 + ', ','')  + isnull( pr.mobile + ', ','') +  isnull( pr.mobile2 + ', ','') as Moblie, packingname, pr.city, " & vbCrLf _
                + " isnull(pr.Address,'') + isnull(' - ' + pr.City,'') + isnull(',' + pr.Phone1,'') + isnull(',' + pr.Phone2, '') + isnull(',' + pr.Mobile, '') as AddressFull, isLastPrice " & vbCrLf _
                + " from SaleBody b inner join products p on b.productid = p.productid" & vbCrLf _
                + " inner join SaleHeader h on H.SID = B.SID" & vbCrLf _
                + " inner join users ur on ur.UserNo = h.UserNo Left Outer jOin companies cmp on cmp.companyid = p.companyid " & vbCrLf _
                + " Left Outer jOin Groups g on g.Groupid = p.Groupid Left Outer jOin SubGroups sg on sg.subGroupid = p.subGroupid" & vbCrLf _
                + " Left Outer jOin Brands bd on bd.brandid = p.brandid" & vbCrLf _
                + " Left Outer jOin Seasons se on se.Seasonid = p.Seasonid" & vbCrLf _
                + " LEFT OUTER JOIN packings pak on pak.packingid = b.packingid" & vbCrLf _
                + " left outer join Organizations o on o.OrganizationID = h.OrganizationID" & vbCrLf _
                + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
                + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
                + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
                + " left outer join Sectors Sec on Sec.SectorID = Pr.SectorID" & vbCrLf _
                + " left outer join Zones Z on Z.ZoneID = Sec.ZoneID" & vbCrLf _
                + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
                + " where h.BillDate = '" & !BillDate & "' and h.BillID = " & !BillID & IIf(ObjRegistry.AllowOrderByCodeinInvoices, " Order By Code", " Order By SerialNo")
         
      If ObjRegistry.LaserPrintofSaleInvoice = True Then
'         vStrSQL = "Select UserName, h.billid, h.BillDate, isnull(h.BillTime,0) as BillTime, h.Description, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, isnull(h.STax,0) as STax, isnull(h.cashReceived,0) as CashReceived, " & IIf(isUrdu = True, "p.ProductName1", "p.ProductName") & " as ProductName, unitname, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + Qty as Qty, round(cast(b.price as numeric(9,2))/isnull(multiplier,1),3) as price, b.amount, b.DiscVal, InvoiceNo" & vbCrLf _
               + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else h.CustomerID + ' - ' + AccountName End as Customer, isnull(pr.Address,'') + isnull(' (' + pr.City + ')','') as Address, Cash, Credit, BankCard, b.ProductID, PreviousAmount, isnull(OtherCharges,0) as OtherCharges, h.Empid, e.empname, dbo.FunSaleBodySerial(b.BillID,b.BillDate, b.ProductId) Serial, h.TableID, isnull(TableName,'') as TableName, null as DeliveryDate, isnull(h.isPrinted,0) as isPrinted," & IIf(ObjRegistry.AllowUrduProduct = False, " isnull(Remarks,'')", " isnull(RemarksUrdu,'')") & " as Remarks, pr.Phone1 " & vbCrLf _
               + " from saleHeader h inner join salebody b on h.billid = b.billid and h.BillDate = b.BillDate" & vbCrLf _
               + " inner join products p on p.productid = b.productid" & vbCrLf _
               + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
               + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
               + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
               + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
               + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
               + " left outer join Tables t on t.TableID = h.TableID " & vbCrLf _
               + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
               + " where h.BillDate = '" & !BillDate & "' and h.BillID = " & !BillID & IIf(ObjRegistry.AllowOrderByCodeinInvoices, "Order By Code", "Order By SerialNo")
      
'      vStrSQL = " select UserName, h.billid, h.BillDate, isnull(h.BillTime,0) as BillTime, h.Description, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.cashReceived,0) as CashReceived, p.ProductName /*case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.billid, h.BillDate) end */ ProductName, unitname, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + Qty as Qty, b.price/isnull(multiplier,1) as price, b.amount, b.DiscVal, InvoiceNo" & vbCrLf _
            + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else h.CustomerID + ' - ' + AccountName End as Customer, isnull(pr.Address,'') as Address, Cash, Credit, BankCard, b.ProductID, PreviousAmount, isnull(OtherCharges,0) as OtherCharges,  h.Empid, e.empname, dbo.FunSaleBodySerial(b.BillID,b.BillDate, b.ProductId) Serial " & vbCrLf _
            + " from saleHeader h inner join salebody b on h.billid = b.billid and h.BillDate = b.BillDate" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
            + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
            + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
            + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
            + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate ='" & DtpBillDate.DateValue & "' Order By SerialNo"
      End If


    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
   
    RptReportViewer.Report.SelectPrinter "abc", "xyz", "ghi"
   
'   Set RptReportViewer.Report = New CrpSaleInvoiceHalf1
   
   If ObjRegistry.LaserPrintofSaleInvoice = True Then
'      Set RptReportViewer.Report = New CrpSaleInvoiceHalf1
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleInvoiceHalf1.rpt")
      RptReportViewer.Report.PaperSize = crPaperA4
      RptReportViewer.Report.PaperOrientation = crLandscape
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.x
      RptReportViewer.Report.RightMargin = 225
   Else
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\" & IIf(IsNull(!SaleReportName), "CrptSaleInvoice", !SaleReportName) & ".rpt")
      
'      Set RptReportViewer.Report = New CrptSaleInvoice
      RptReportViewer.Report.PaperOrientation = crPortrait
   End If
   
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ReportTitle = "Sale Invoice"
   If ObjRegistry.LaserPrintofSaleInvoice = True Then
'      RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy  'CN.Execute("Select Name from Manufacturer").Fields(0).Value
'      RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
'      RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.AddSpace = True, ".", "")
'      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(ObjRegistry.CashReceived)
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
      RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CStr(ObjRegistry.Statement)
      RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(ObjRegistry.Statement)
      RptReportViewer.Report.ParameterFields(8).AddCurrentValue ""
'      RptReportViewer.Report.ParameterFields(9).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
   Else
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
      RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CStr(ObjRegistry.Statement)
   End If
   If ObjRegistry.PrintHeadersSaleInvoice = True Then
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   Else
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ""
   End If
   If ObjRegistry.PreviewSaleInoice Then
      RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False, CInt(IIf(IsNull(ObjRegistry.NoofPrints) Or Val(ObjRegistry.NoofPrints) = 0, 1, ObjRegistry.NoofPrints))
   End If
'   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
      
      .MoveNext
      Wend
      .Close
   End With
      
  
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage

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
'         Case vbKeyV
'            If BtnPreview.Enabled Then BtnPreview_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, True) = True Then TxtSectorID.SetFocus
         Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, True) = True Then TxtBillIDFrom.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Sale Invoice Batch Print"
   DtpTo.DateValue = Date
   DtpFrom.DateValue = Date
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
   Set RptBatchPrint = Nothing
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
   Set RsReport = cn.Execute("EXEC ProdRptDateWiseSaleExpense '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'")
'   Set RptReportViewer.Report = New CrptDateWiseSaleExpense
   If RsReport.BOF Then
      MsgBox "No record exists.", vbInformation, Me.Caption
      Me.MousePointer = vbDefault
      Exit Function
   End If
   RptReportViewer.Report.ReportTitle = "Date Wise Sale Expense"
   RptReportViewer.Report.Database.SetDataSource RsReport
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue " Date Range : From " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To " & Format(DtpTo.DateValue, "dd/MM/yyyy")
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port

   RptReportViewer.Report.PaperOrientation = crPortrait
   SetReport = True
   Me.MousePointer = vbDefault
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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
   vTemp = Not FunSelectOrganization(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganization(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectOrganization(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganization = False: Exit Function
       TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    If TxtOrganizationID.Text = "" Then FunSelectOrganization = False: Exit Function
    vStrSQL = " Select * FROM Organizations where OrganizationID='" & TxtOrganizationID.Text & "'"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizatonName.Text = !OrganizationName
          FunSelectOrganization = True
          .Close
          Exit Function
      Else
          FunSelectOrganization = False
          .Close
         TxtOrganizationID.Text = ""
          TxtOrganizatonName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtSectorID_Change()
   If TxtSectorID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   If TxtSectorName.Text <> "" Then TxtSectorName.Text = ""
End Sub

Private Sub TxtSectorID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSectorID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSector(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSector(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSector(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSector.Show vbModal, Me
        If SchSector.ParaOutSectorID = "" Then FunSelectSector = False: Exit Function
        TxtSectorID.Text = SchSector.ParaOutSectorID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Sectors where SectorID=" & Val(TxtSectorID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSectorName.Text = !SectorName
          FunSelectSector = True
          .Close
          Exit Function
             FunSelectSector = True
   Else
          FunSelectSector = False
          .Close
          TxtSectorID.Text = ""
          TxtSectorName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

