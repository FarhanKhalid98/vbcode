VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form SchReplacement 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   796
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCustomerName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5168
      TabIndex        =   9
      Top             =   2723
      Width           =   1980
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7148
      TabIndex        =   8
      Top             =   2723
      Width           =   2040
   End
   Begin VB.TextBox TxtManualBillNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   9188
      TabIndex        =   7
      Top             =   2723
      Width           =   1260
   End
   Begin VB.TextBox TxtReplaceID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1793
      TabIndex        =   1
      Top             =   2723
      Width           =   975
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6660
      Left            =   1800
      TabIndex        =   0
      Top             =   3060
      Width           =   8970
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   16579021
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "SchReplacement.frx":0000
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
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   14
      Columns(0).Width=   1720
      Columns(0).Caption=   "Replace ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2196
      Columns(1).Caption=   "Replace Date"
      Columns(1).Name =   "Date"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).NumberFormat=   "dd/MM/yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   3493
      Columns(2).Caption=   "Customer Name"
      Columns(2).Name =   "CustomerName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 4"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "Store Name"
      Columns(3).Name =   "StoreName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 8"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "Manual Bill No"
      Columns(4).Name =   "ManualBillNo"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 10"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1667
      Columns(5).Caption=   "Return Amt"
      Columns(5).Name =   "ReturnAmount"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1693
      Columns(6).Caption=   "Sale Amt"
      Columns(6).Name =   "SaleAmount"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 5"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   2937
      Columns(7).Caption=   "Bill Type"
      Columns(7).Name =   "BillType"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 6"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1879
      Columns(8).Caption=   "CO"
      Columns(8).Name =   "CO"
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 5"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "Tag"
      Columns(9).Name =   "Tag"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1085
      Columns(10).Caption=   "Closed"
      Columns(10).Name=   "Closed"
      Columns(10).DataField=   "Column 7"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(10).Style=   2
      Columns(11).Width=   3200
      Columns(11).Visible=   0   'False
      Columns(11).Caption=   "SSID"
      Columns(11).Name=   "SSID"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "RSID"
      Columns(12).Name=   "RSID"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "SID"
      Columns(13).Name=   "SID"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   15822
      _ExtentY        =   11748
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   6406
      TabIndex        =   2
      Top             =   9893
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Select"
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
      MICON           =   "SchReplacement.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7726
      TabIndex        =   3
      Top             =   9893
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
      MICON           =   "SchReplacement.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   330
      Left            =   2768
      TabIndex        =   13
      Top             =   2723
      Width           =   1200
      _Version        =   65543
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpToDate 
      Height          =   330
      Left            =   3968
      TabIndex        =   14
      Top             =   2723
      Width           =   1200
      _Version        =   65543
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.Label LblCustomerName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
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
      Left            =   5168
      TabIndex        =   12
      Top             =   2498
      Width           =   1335
   End
   Begin VB.Label LblTag 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
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
      Left            =   7148
      TabIndex        =   11
      Top             =   2498
      Width           =   345
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill No"
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
      Left            =   9173
      TabIndex        =   10
      Top             =   2498
      Width           =   1245
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   270
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Replace Date"
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
      Left            =   2768
      TabIndex        =   5
      Top             =   2498
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   13328
      Top             =   1628
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Replace ID"
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
      Left            =   1793
      TabIndex        =   4
      Top             =   2498
      Width           =   975
   End
End
Attribute VB_Name = "SchReplacement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSQL As String, vPartyName As String, vTag As String, vManualBillNo As String
Dim vOrder As String, vDirection As String, vCol As Byte, vSearchInPreviousState As Boolean
Public ParaOutSID As String
Public ParaOutSSID As String
Public ParaOutRSID As String
Public ParaOutReplaceID As String
Public ParaOutReplaceDate As String
Public ParaOutBillID As String
Public ParaOutBillDate As String
Public ParaOutReturnID As String
Public ParaOutReturnDate As String
Public ParaInReplaceDate As String

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Set Rs = New ADODB.Recordset
   vStrSQL = " select r.*, Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName + isnull(' (' + City + ')','') End as CustomerName, UserName, " & vbCrLf _
         + " case when cash = 1 then 'Cash' " & vbCrLf _
         + " when credit = 1 and isnull(s.CashReceived,0) > 0 and isnull(BankAmount,0) > 0 then 'Credit + Cash + Bankd card' When credit = 1 and isnull(s.CashReceived,0) > 0 and isnull(BankAmount,0) = 0 then 'Credit + Cash'  When credit = 1 and isnull(s.CashReceived,0) = 0 and isnull(BankAmount,0) > 0 then 'Credit + BankCard'  When credit = 1 and isnull(s.CashReceived,0) = 0 and isnull(BankAmount,0) = 0 then 'Credit'  " & vbCrLf _
         + " when BankCard = 1 and isnull(s.CashReceived,0) = 0 then 'Bank Card' when BankCard = 1 and isnull(s.CashReceived,0) > 0 then 'Bank Card + Cash' end as BillType, InvType, " & vbCrLf _
         + " r.isPosted, StoreName, r.Tag, isnull(ManualBillNo,'')ManualBillNo" & vbCrLf _
         + " from ReplacementHeader r " & vbCrLf _
         + " inner join saleheader s on r.SSID = s.Sid" & vbCrLf _
         + " inner join chartofaccounts c on c.AccountNo = s.customerid " & vbCrLf _
         + " left outer join Parties pt on pt.PartyID = s.customerid " & vbCrLf _
         + " inner join users u on u.userno = r.userno " & vbCrLf _
         + " inner join Stores st on s.StoreID = st.StoreID " & vbCrLf _
         + " WHERE r.ReplaceDate Between '" & DtpFromDate.DateValue & "' and '" & DtpToDate.DateValue & "'" & IIf(ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False, " and r.isPosted = 0 and r.userno=" & ObjUserSecurity.UserNo, "") & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID) & vPartyName & vTag & vManualBillNo & vOrder & vDirection
   Rs.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("SID").DataField = "SID"
   Grid.Columns("SSID").DataField = "SSID"
   Grid.Columns("RSID").DataField = "RSID"
   Grid.Columns("ID").DataField = "ReplaceID"
   Grid.Columns("Date").DataField = "ReplaceDate"
   Grid.Columns("CustomerName").DataField = "CustomerName"
   Grid.Columns("StoreName").DataField = "StoreName"
   Grid.Columns("SaleAmount").DataField = "SaleAmount"
   Grid.Columns("ReturnAmount").DataField = "ReturnAmount"
   Grid.Columns("CO").DataField = "UserName"
   Grid.Columns("BillType").DataField = "BillType"
   Grid.Columns("Closed").DataField = "isPosted"
   Grid.Columns("Tag").DataField = "Tag"
   Grid.Columns("ManualBillNo").DataField = "ManualBillNo"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Me.ParaOutReplaceID = -1
   Me.ParaOutReplaceDate = ""
   Me.ParaOutBillID = ""
   Me.ParaOutSID = ""
   Me.ParaOutSSID = ""
   Me.ParaOutRSID = ""
   Me.ParaOutBillDate = ""
   Me.ParaOutReturnID = ""
   Me.ParaOutReturnDate = ""
   Unload Me
End Sub

Private Sub BtnSelect_Click()
   On Error GoTo ErrorHandler
   If Grid.Rows = 0 Then Exit Sub
   Me.ParaOutReplaceID = Rs!ReplaceId
   Me.ParaOutReplaceDate = Rs!ReplaceDate
   Me.ParaOutSID = Rs!SID
   Me.ParaOutSSID = Rs!SSID
   Me.ParaOutRSID = Rs!RSID
   Me.ParaOutBillID = Rs!BillID
   Me.ParaOutBillDate = Rs!BillDate
   Me.ParaOutReturnID = Rs!ReturnID
   Me.ParaOutReturnDate = Rs!ReturnDate
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpFromDate_Change()
   If DtpFromDate.IsDateValid = False Then Exit Sub
   If DtpToDate.Visible = False Then DtpToDate.DateValue = DtpFromDate.DateValue
   Call LoadGrid
End Sub

Private Sub DtpToDate_Change()
   If DtpToDate.IsDateValid = False Then Exit Sub
   Call LoadGrid
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"
   
   Dim vDateDiff As Byte
   
   vDateDiff = ObjRegistry.SearchDateDifference
   vSearchInPreviousState = ObjRegistry.ProductSearchOpenInPreviousState
   
   LblTag.Visible = ObjRegistry.Tag
   TxtTag.Visible = ObjRegistry.Tag
   Grid.Columns("Tag").Visible = ObjRegistry.Tag
         
   Grid.Columns("ManualBillNo").Visible = ObjRegistry.ManualBillNoVisible
   LblManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   TxtManualBillNo.Visible = ObjRegistry.ManualBillNoVisible

   Grid.Columns("StoreName").Visible = ObjRegistry.StoreVisible
   If ObjUserSecurity.IsAdministrator = False Then
      Grid.Columns("SaleAmount").Visible = Not ObjRegistry.HideSaleAmount
      Grid.Columns("ReturnAmount").Visible = Not ObjRegistry.HideSaleAmount
   End If
   
   Dim vWidth As Long, i As Integer
   vWidth = 0
   For i = 0 To Grid.Cols - 1
      If Grid.Columns(i).Visible = True Then
         vWidth = vWidth + Grid.Columns(i).Width
      End If
   Next i
   Grid.Width = vWidth + 18
   
   DtpToDate.DateValue = Me.ParaInReplaceDate
   DtpFromDate.DateValue = DtpToDate.DateValue - vDateDiff
   
   If vDateDiff = 0 Then
      DtpToDate.Visible = False
      LblCustomerName.Left = LblCustomerName.Left - DtpToDate.Width
      TxtCustomerName.Left = TxtCustomerName.Left - DtpToDate.Width
      LblTag.Left = LblTag.Left - DtpToDate.Width
      TxtTag.Left = TxtTag.Left - DtpToDate.Width
      LblManualBillNo.Left = LblManualBillNo.Left - DtpToDate.Width
      TxtManualBillNo.Left = TxtManualBillNo.Left - DtpToDate.Width
   End If

   Me.ParaOutReplaceID = -1
   Me.ParaOutReplaceDate = ""
   
   vOrder = " Order By ReplaceID"
   vDirection = " Desc"
   If vSearchInPreviousState = False Then
      vTag = ""
      vPartyName = ""
      vManualBillNo = ""
   End If
   Call LoadGrid
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtReplaceID.Name, DtpFromDate.Name, DtpToDate.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
   Set Rs = Nothing
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then BtnSelect_Click
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
   vOrder = " order by " & Grid.Columns(ColIndex).DataField
   If vCol = ColIndex Then
      vDirection = IIf(vDirection = " Asc", " Desc", " Asc")
   Else
      vDirection = " Asc"
   End If
   vCol = ColIndex
   LoadGrid
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9
      TxtReplaceID.Text = Chr(KeyAscii): TxtReplaceID.SelStart = Len(TxtReplaceID.Text): TxtReplaceID.SetFocus
   End Select
End Sub

Private Sub Image1_Click()
   Unload Me
End Sub

Private Sub TxtReplaceID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtReplaceID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "SaleID = " & TxtReplaceID.Text, , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCustomerName_Change()
   On Error GoTo ErrorHandler
   vPartyName = " and (Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName + isnull(' (' + City + ')','') End) like '%" & TxtCustomerName.Text & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTag_Change()
   On Error GoTo ErrorHandler
   vTag = " and r.Tag like '%" & TxtTag.Text & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtManualBillNo_Change()
   On Error GoTo ErrorHandler
   vManualBillNo = " and ManualBillNo like '%" & (TxtManualBillNo.Text) & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

