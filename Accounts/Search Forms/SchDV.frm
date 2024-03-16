VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form SchDV 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtStoreName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4800
      TabIndex        =   16
      Top             =   1425
      Width           =   1725
   End
   Begin VB.TextBox TxtVoucherNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   225
      TabIndex        =   13
      Top             =   1425
      Width           =   1020
   End
   Begin VB.TextBox TxtTotalAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2610
      TabIndex        =   12
      Top             =   9135
      Width           =   3630
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6525
      TabIndex        =   2
      Top             =   1425
      Width           =   2040
   End
   Begin VB.TextBox TxtAccountName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8865
      TabIndex        =   3
      Top             =   1425
      Width           =   3255
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   6375
      TabIndex        =   5
      Top             =   9675
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Select"
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
      MICON           =   "SchDV.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7695
      TabIndex        =   6
      Top             =   9675
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Cancel"
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
      MICON           =   "SchDV.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnFind 
      Default         =   -1  'True
      Height          =   420
      Left            =   12135
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1335
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Refresh"
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
      MICON           =   "SchDV.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   1305
      TabIndex        =   0
      Top             =   1425
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
      Left            =   2610
      TabIndex        =   1
      Top             =   1425
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridDetail 
      Height          =   7365
      Left            =   8865
      TabIndex        =   9
      Top             =   1755
      Width           =   5655
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
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
      stylesets(0).Picture=   "SchDV.frx":0054
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
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   3
      Columns(0).Width=   3519
      Columns(0).Caption=   "Account Name"
      Columns(0).Name =   "AccountName"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3519
      Columns(1).Caption=   "Narration"
      Columns(1).Name =   "Narration"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 2"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2434
      Columns(2).Caption=   "Amount"
      Columns(2).Name =   "Amount"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 1"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   9975
      _ExtentY        =   12991
      _StockProps     =   79
      BackColor       =   15724527
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   7380
      Left            =   225
      TabIndex        =   14
      Top             =   1755
      Width           =   8595
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
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
      stylesets(0).Picture=   "SchDV.frx":0070
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
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   9
      Columns(0).Width=   1746
      Columns(0).Caption=   "Voucher #"
      Columns(0).Name =   "ID"
      Columns(0).Alignment=   2
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1958
      Columns(1).Caption=   "Date"
      Columns(1).Name =   "Date"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).NumberFormat=   "dd/MM/yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   1931
      Columns(2).Caption=   "Total Amount"
      Columns(2).Name =   "Amount"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1773
      Columns(3).Caption=   "User Name"
      Columns(3).Name =   "UserName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1588
      Columns(4).Caption=   "Entry Time"
      Columns(4).Name =   "EntryTime"
      Columns(4).DataField=   "Column 7"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3043
      Columns(5).Caption=   "StoreName"
      Columns(5).Name =   "StoreName"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 6"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3519
      Columns(6).Caption=   "Tag"
      Columns(6).Name =   "Tag"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 4"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "StoreID"
      Columns(7).Name =   "StoreID"
      Columns(7).DataField=   "Column 6"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "SID"
      Columns(8).Name =   "SID"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   15161
      _ExtentY        =   13017
      _StockProps     =   79
      BackColor       =   15724527
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
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
      Left            =   4815
      TabIndex        =   17
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No"
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
      Left            =   225
      TabIndex        =   15
      Top             =   1200
      Width           =   1020
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
      Left            =   6540
      TabIndex        =   11
      Top             =   1200
      Width           =   345
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Index           =   0
      Left            =   3000
      TabIndex        =   10
      Top             =   270
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name"
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
      Left            =   8865
      TabIndex        =   8
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   13950
      Top             =   945
      Width           =   330
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "-----  Voucher Date Range ------"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   1200
      Width           =   2670
   End
End
Attribute VB_Name = "SchDV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Public ParaOutVoucherNo As Long
Public ParaOutStoreID As Long
Public ParaOutSID As Long
Dim vOrder As String, vDirection As String, vCol As Byte, vSQL As String

Private Sub LoadData()
   On Error GoTo ErrorHandler
   Set Rs = New ADODB.Recordset
   vSQL = " Select H.SID, h.VoucherNo ID, Voucherdate as Date, Sum(Amount) as Amount, Substring(CONVERT(varchar(20),isnull(ServerEntry,0)),13,7) as ServerEntry, UserName, Tag, StoreName, H.StoreID" & vbCrLf _
      + " from Debitvouchers h inner join Debitvouchersbody b on h.voucherno = b.voucherno and h.storeid = b.storeid" & vbCrLf _
      + " inner join chartofaccounts c on c.AccountNo = b.AccountNo" & vbCrLf _
      + " inner join Users u on u.UserNo = h.UserNo" & vbCrLf _
      + " inner join Stores s on s.StoreID = h.StoreID" & vbCrLf _
      + " Where Voucherdate between '" & DtpFrom.DateValue & "' AND '" & DtpTo.DateValue & "'" & vbCrLf _
      + IIf(Trim(TxtAccountName.Text) = "", "", " and accountname like '%" & TxtAccountName.Text & "%'") & vbCrLf _
      + IIf(Trim(TxtStoreName.Text) = "", "", " and StoreName like '%" & TxtStoreName.Text & "%'") & vbCrLf _
      + IIf(Trim(TxtTag.Text) = "", "", " and Tag like '%" & TxtTag.Text & "%'") & vbCrLf _
      + IIf(ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False, " and h.UserNo=" & ObjUserSecurity.UserNo, "") & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID) & vbCrLf _
      + " Group by h.SID, h.VoucherNo, Voucherdate, UserName, Tag, StoreName, H.StoreID, ServerEntry" & vOrder & vDirection
    
   Rs.Open vSQL, cn
   Set Grid.DataSource = Rs
   Grid.Columns("SID").DataField = "SID"
   Grid.Columns("ID").DataField = "ID"
   Grid.Columns("Date").DataField = "Date"
   Grid.Columns("Amount").DataField = "Amount"
   Grid.Columns("EntryTime").DataField = "ServerEntry"
   Grid.Columns("StoreID").DataField = "StoreID"
   Grid.Columns("StoreName").DataField = "StoreName"
   Grid.Columns("UserName").DataField = "UserName"
   Grid.Columns("Tag").DataField = "Tag"
   
   
    vSQL = " Select isnull(Sum(Amount),0) as Amount " & vbCrLf _
      + " from Debitvouchers h inner join Debitvouchersbody b on h.voucherno = b.voucherno and h.storeid = b.storeid" & vbCrLf _
      + " inner join chartofaccounts c on c.AccountNo = b.AccountNo" & vbCrLf _
      + " inner join Users u on u.UserNo = h.UserNo" & vbCrLf _
      + " Where Voucherdate between '" & DtpFrom.DateValue & "' AND '" & DtpTo.DateValue & "'" & vbCrLf _
      + IIf(Trim(TxtAccountName.Text) = "", "", " and accountname like '%" & TxtAccountName.Text & "%'") & vbCrLf _
      + IIf(Trim(TxtTag.Text) = "", "", " and Tag like '%" & TxtTag.Text & "%'") & vbCrLf _
      + IIf(ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False, " and h.UserNo=" & ObjUserSecurity.UserNo, "")
   
   With cn.Execute(vSQL)
      If .RecordCount > 0 Then
         TxtTotalAmount.Text = .Fields(0)
      End If
   End With
      
   LoadDetail
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadDetail()
  On Error GoTo ErrorHandler
    vSQL = " Select c.AccountName + isnull(' (' + p.Address + ')','')  as AccountName, b.Narration, Amount from Debitvouchers h" & vbCrLf _
      + " inner join Debitvouchersbody b on h.voucherno = b.voucherno and h.storeid = b.storeid" & vbCrLf _
      + " inner join chartofaccounts c on c.accountno = b.accountno" & vbCrLf _
      + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
      + " Where h.voucherno = " & Val(Grid.Columns("ID").Text) & " and h.StoreID = " & Val(Grid.Columns("StoreID").Text)
    Set GridDetail.DataSource = cn.Execute(vSQL)
    GridDetail.Columns("AccountName").DataField = "AccountName"
    GridDetail.Columns("Narration").DataField = "Narration"
    GridDetail.Columns("Amount").DataField = "Amount"
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnFind_Click()
    LoadData
End Sub

Private Sub BtnClose_Click()
  Me.ParaOutVoucherNo = 0
  Me.ParaOutStoreID = 0
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutVoucherNo = Rs!ID
  Me.ParaOutStoreID = Rs!StoreID
  Me.ParaOutSID = Rs!SID
  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, DtpFrom.Name, DtpTo.Name, TxtTag.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"
   
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   Me.ParaOutVoucherNo = 0
   vOrder = " Order by ID"
   vDirection = " Desc"
   LoadData
   LblTag.Visible = ObjRegistry.Tag
   TxtTag.Visible = ObjRegistry.Tag
   Grid.Columns("Tag").Visible = ObjRegistry.Tag
   
   TxtTotalAmount.Visible = ObjRegistry.ShowGrandTotalinSearch
   
   If TxtTag.Visible = False Then
      Dim vWidth As Long, i As Integer
      vWidth = 0
      For i = 0 To Grid.Cols - 1
         If Grid.Columns(i).Visible = True Then
            vWidth = vWidth + Grid.Columns(i).Width
         End If
      Next i
   End If
   Grid.Width = vWidth + 18
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
   LoadData
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   LoadDetail
End Sub

Private Sub GridDetail_DblClick()
   If GridDetail.Rows > 0 Then BtnSelect_Click
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtVoucherNo_Change()
  On Error GoTo ErrorHandler
  If Trim(TxtVoucherNo.Text) = "" Then Grid.MoveFirst: Exit Sub
  Rs.Find "ID=" & TxtVoucherNo.Text, , adSearchForward, 1
  If Rs.EOF Then Grid.MoveLast
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub
