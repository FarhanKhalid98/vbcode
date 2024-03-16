VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form SchService 
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
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5644
      TabIndex        =   10
      Top             =   2723
      Width           =   2040
   End
   Begin VB.TextBox TxtCustomerName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3709
      TabIndex        =   3
      Top             =   2723
      Width           =   1935
   End
   Begin VB.TextBox TxtBillID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1812
      TabIndex        =   1
      Top             =   2723
      Width           =   690
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6660
      Left            =   1812
      TabIndex        =   0
      Top             =   3053
      Width           =   11745
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
      stylesets(0).Picture=   "SchService.frx":0000
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
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
      Columns.Count   =   12
      Columns(0).Width=   1191
      Columns(0).Caption=   "Bill ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2143
      Columns(1).Caption=   "Bill Date"
      Columns(1).Name =   "Date"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).NumberFormat=   "dd/MM/yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   3413
      Columns(2).Caption=   "Customer Name"
      Columns(2).Name =   "CustomerName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 4"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2752
      Columns(3).Caption=   "Store Name"
      Columns(3).Name =   "StoreName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 11"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1270
      Columns(4).Caption=   "Bill Time"
      Columns(4).Name =   "BillTime"
      Columns(4).DataField=   "Column 8"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1296
      Columns(5).Caption=   "Ttl Items"
      Columns(5).Name =   "TotalItems"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1667
      Columns(6).Caption=   "Ttl Amount"
      Columns(6).Name =   "Amount"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 5"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1296
      Columns(7).Caption=   "Bill Type"
      Columns(7).Name =   "BillType"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 6"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1773
      Columns(8).Caption=   "CO"
      Columns(8).Name =   "CO"
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 5"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   4683
      Columns(9).Caption=   "Tag"
      Columns(9).Name =   "Tag"
      Columns(9).CaptionAlignment=   0
      Columns(9).DataField=   "Column 10"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1111
      Columns(10).Caption=   "Closed"
      Columns(10).Name=   "Closed"
      Columns(10).DataField=   "Column 7"
      Columns(10).DataType=   11
      Columns(10).FieldLen=   256
      Columns(10).Style=   2
      Columns(11).Width=   1508
      Columns(11).Caption=   "Replaced"
      Columns(11).Name=   "Replaced"
      Columns(11).DataField=   "Column 9"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(11).Style=   2
      TabNavigation   =   1
      _ExtentX        =   20717
      _ExtentY        =   11747
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
      Left            =   6387
      TabIndex        =   4
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
      MICON           =   "SchService.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7707
      TabIndex        =   5
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
      MICON           =   "SchService.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpDate 
      Height          =   330
      Left            =   2509
      TabIndex        =   2
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
      Left            =   5644
      TabIndex        =   11
      Top             =   2498
      Width           =   345
   End
   Begin VB.Label Label3 
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
      Left            =   3709
      TabIndex        =   9
      Top             =   2498
      Width           =   1335
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
      TabIndex        =   8
      Top             =   270
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      Left            =   2509
      TabIndex        =   7
      Top             =   2498
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   13309
      Top             =   1628
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill ID"
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
      Left            =   1812
      TabIndex        =   6
      Top             =   2498
      Width           =   525
   End
End
Attribute VB_Name = "SchService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim VStrSQL As String, vPartyName As String, vTag As String
Dim vOrder As String, vDirection As String, vCol As Byte
Public ParaOutBillID As String
Public ParaOutBillDate As String
Public ParaInBillDate As String

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   If DtpDate.IsDateValid = False Then Exit Sub
   Set Rs = New ADODB.Recordset
   VStrSQL = "SELECT h.BillID as SaleID, h.BillDate as SaleDate, Substring(CONVERT(varchar(20),isnull(BillTime,0)),13,7) as BillTime, case when credit = 1 then 'Credit' when cash = 1 then 'Cash' when BankCard = 1 then 'Bank Card' end as BillType " & vbCrLf _
         + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName End as CustomerName, TotalAmount-isnull(billdisc,0) as TotalAmount, TotalItems, UserName, isPosted, StoreName, Tag" & vbCrLf _
         + " FROM ServiceHeader h INNER JOIN" & vbCrLf _
         + " (SELECT BillID,BillDate, sum(qty) as TotalItems, sum(amount) Amount FROM ServiceBody GROUP BY BillID, BillDate) b" & vbCrLf _
         + " ON h.billID = b.billID and h.BillDate = b.BillDate" & vbCrLf _
         + " left outer JOIN chartofaccounts c ON h.CustomerID = c.AccountNo " & vbCrLf _
         + " INNER JOIN users u ON h.userno = u.userno " & vbCrLf _
         + " INNER JOIN Stores s ON s.StoreID = h.StoreID " & vbCrLf _
         + " WHERE h.BillDate ='" & DtpDate.DateValue & "'" & IIf(ObjUserSecurity.IsAdministrator = False, " and isPosted = 0 and h.userno=" & ObjUserSecurity.UserNo, "") & vPartyName & vTag & vOrder & vDirection
   Rs.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "SaleID"
   Grid.Columns("Date").DataField = "SaleDate"
   Grid.Columns("BillTime").DataField = "BillTime"
   Grid.Columns("CustomerName").DataField = "CustomerName"
   Grid.Columns("TotalItems").DataField = "TotalItems"
   Grid.Columns("Amount").DataField = "TotalAmount"
   Grid.Columns("CO").DataField = "username"
   Grid.Columns("BillType").DataField = "BillType"
   Grid.Columns("Closed").DataField = "isPosted"
   Grid.Columns("StoreName").DataField = "StoreName"
   Grid.Columns("Tag").DataField = "Tag"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Me.ParaOutBillID = -1
   Me.ParaOutBillDate = ""
   Unload Me
End Sub

Private Sub BtnSelect_Click()
   On Error GoTo ErrorHandler
   If Grid.Rows = 0 Then Exit Sub
   Me.ParaOutBillID = Rs!SaleID
   Me.ParaOutBillDate = Rs!SaleDate
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpDate_Change()
   Call LoadGrid
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"
   DtpDate.DateValue = Me.ParaInBillDate
   Me.ParaOutBillID = -1
   Me.ParaOutBillDate = ""
   vOrder = " order by SaleID"
   vDirection = " desc"
   vPartyName = ""
   vTag = ""
   Call LoadGrid
   LblTag.Visible = ObjRegistry.Tag
   TxtTag.Visible = ObjRegistry.Tag
   Grid.Columns("Tag").Visible = ObjRegistry.Tag
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
   If ObjUserSecurity.IsAdministrator = False Then
      Grid.Columns("Amount").Visible = Not ObjRegistry.HideSaleAmount
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtBillID.Name, DtpDate.Name
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
      TxtBillID.Text = Chr(KeyAscii): TxtBillID.SelStart = Len(TxtBillID.Text): TxtBillID.SetFocus
   End Select
End Sub

Private Sub Image1_Click()
   Unload Me
End Sub

Private Sub TxtBillID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtBillID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "SaleID = " & TxtBillID.Text, , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCustomerName_Change()
   On Error GoTo ErrorHandler
   vPartyName = " and (Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName End) like '%" & TxtCustomerName.Text & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTag_Change()
   On Error GoTo ErrorHandler
   vTag = " and tag Like '%" & TxtTag.Text & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
