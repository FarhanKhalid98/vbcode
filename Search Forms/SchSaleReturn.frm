VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form SchSaleReturn 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   11940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   796
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTableName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   9308
      TabIndex        =   5
      Top             =   2723
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox TxtCustomerName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5168
      TabIndex        =   2
      Top             =   2723
      Width           =   1935
   End
   Begin VB.TextBox TxtTTLAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7103
      TabIndex        =   3
      Top             =   2723
      Width           =   945
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   10568
      TabIndex        =   6
      Top             =   2723
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.TextBox TxtManualBillNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8048
      TabIndex        =   4
      Top             =   2723
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox TxtReturnID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1793
      TabIndex        =   1
      Top             =   2723
      Width           =   885
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   6406
      TabIndex        =   7
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
      MICON           =   "SchSaleReturn.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7726
      TabIndex        =   8
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
      MICON           =   "SchSaleReturn.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6660
      Left            =   1793
      TabIndex        =   0
      Top             =   3053
      Width           =   11820
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
      stylesets(0).Picture=   "SchSaleReturn.frx":0038
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
      Columns.Count   =   13
      Columns(0).Width=   1561
      Columns(0).Caption=   "Return ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2196
      Columns(1).Caption=   "Return Date"
      Columns(1).Name =   "Date"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).NumberFormat=   "dd/MM/yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   3228
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
      Columns(3).DataField=   "Column 7"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "Table Name"
      Columns(4).Name =   "TableName"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 11"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1508
      Columns(5).Caption=   "TTL Items"
      Columns(5).Name =   "TotalItems"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1879
      Columns(6).Caption=   "TTL Amount"
      Columns(6).Name =   "Amount"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 5"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   2037
      Columns(7).Caption=   "CO"
      Columns(7).Name =   "CO"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 5"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "Manual Bill No"
      Columns(8).Name =   "ManualBillNo"
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 10"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "Tag"
      Columns(9).Name =   "Tag"
      Columns(9).DataField=   "Column 8"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1138
      Columns(10).Caption=   "Closed"
      Columns(10).Name=   "Closed"
      Columns(10).DataField=   "Column 6"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(10).Style=   2
      Columns(11).Width=   1508
      Columns(11).Caption=   "Replaced"
      Columns(11).Name=   "Replaced"
      Columns(11).DataField=   "Column 9"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(11).Style=   2
      Columns(12).Width=   3200
      Columns(12).Caption=   "SID"
      Columns(12).Name=   "SID"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20849
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   330
      Left            =   2678
      TabIndex        =   17
      Top             =   2723
      Width           =   1245
      _Version        =   65543
      _ExtentX        =   2196
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
      Left            =   3923
      TabIndex        =   18
      Top             =   2723
      Width           =   1245
      _Version        =   65543
      _ExtentX        =   2196
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
   Begin VB.Label LblTableName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name"
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
      Left            =   9308
      TabIndex        =   16
      Top             =   2498
      Visible         =   0   'False
      Width           =   1035
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
      TabIndex        =   15
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
      Left            =   10568
      TabIndex        =   14
      Top             =   2498
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label LblTTLAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ttl Amount"
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
      Left            =   7103
      TabIndex        =   13
      Top             =   2498
      Width           =   930
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill #"
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
      Left            =   8093
      TabIndex        =   12
      Top             =   2498
      Visible         =   0   'False
      Width           =   1125
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
      TabIndex        =   11
      Top             =   270
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
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
      Left            =   2693
      TabIndex        =   10
      Top             =   2498
      Width           =   1050
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
      Caption         =   "Return ID"
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
      TabIndex        =   9
      Top             =   2498
      Width           =   840
   End
End
Attribute VB_Name = "SchSaleReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSQL As String, vPartyName As String, vTag As String, vManualBillNo As String, vTtlAmount As String, vTableName As String
Dim vOrder As String, vDirection As String, vCol As Byte, vSearchInPreviousState As Boolean
Public ParaOutSID As Double
Public ParaOutReturnID As Integer
Public ParaOutReturnDate As String
Public ParaInReturnDate As String

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Set Rs = New ADODB.Recordset
   vStrSQL = " SELECT H.SID, h.ReturnID as ID, h.ReturnDate as Date, TableName" & vbCrLf _
         + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName + isnull(' (' + City + ')','') End as PartyName, round(b.TotalAmount,0)-isnull(BillDisc,0)+isnull(ServiceCharges,0)+isnull(STax,0)+isnull(OtherCharges,0) as TotalAmount,TotalItems, UserName, isposted, isReplace, StoreName, Tag, isnull(ManualBillNo,'')ManualBillNo" & vbCrLf _
         + " FROM SaleReturnHeader h INNER JOIN" & vbCrLf _
         + " (SELECT SID, sum(qty) as TotalItems, sum(amount) totalamount FROM SaleReturnBody GROUP BY SID) b" & vbCrLf _
         + " ON H.SID = B.SID" & vbCrLf _
         + " left outer JOIN ChartofAccounts c ON h.CustomerID = c.AccountNo " & vbCrLf _
         + " left outer JOIN Parties pt ON pt.PartyID = c.AccountNo " & vbCrLf _
         + " left outer JOIN Tables tb ON tb.TableID = h.TableID " & vbCrLf _
         + " INNER JOIN users u ON h.userno = u.userno " & vbCrLf _
         + " INNER JOIN Stores s ON s.StoreID = h.StoreID " & vbCrLf _
         + " WHERE h.ReturnDate Between '" & DtpFromDate.DateValue & "' and '" & DtpToDate.DateValue & "'" & IIf(ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False, " and isPosted = 0 and h.userno = " & ObjUserSecurity.UserNo, "") & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID) & vPartyName & vTableName & vTag & vManualBillNo & vOrder & vDirection
         
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("SID").DataField = "SID"
   Grid.Columns("ID").DataField = "ID"
   Grid.Columns("Date").DataField = "Date"
   Grid.Columns("CustomerName").DataField = "PartyName"
   Grid.Columns("TableName").DataField = "TableName"
   Grid.Columns("TotalItems").DataField = "TotalItems"
   Grid.Columns("Amount").DataField = "TotalAmount"
   Grid.Columns("CO").DataField = "username"
   Grid.Columns("Closed").DataField = "isPosted"
   Grid.Columns("Replaced").DataField = "isReplace"
   Grid.Columns("StoreName").DataField = "StoreName"
   Grid.Columns("Tag").DataField = "Tag"
   Grid.Columns("ManualBillNo").DataField = "ManualBillNo"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Me.ParaOutSID = -1
   Me.ParaOutReturnID = -1
   Me.ParaOutReturnDate = ""
   Unload Me
End Sub

Private Sub BtnSelect_Click()
   On Error GoTo ErrorHandler
   If Grid.rows = 0 Then Exit Sub
     If Abs(Rs!isReplace) = 1 Then
      Me.ParaOutSID = -1
      Me.ParaOutReturnID = -1
      Me.ParaOutReturnDate = ""
   Else
      Me.ParaOutSID = Rs!SID
      Me.ParaOutReturnID = Rs!ID
      Me.ParaOutReturnDate = Rs!Date
   End If
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
   
   LblManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   TxtManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   Grid.Columns("ManualBillNo").Visible = ObjRegistry.ManualBillNoVisible
   
   LblTableName.Visible = ObjRegistry.TableVisible
   TxtTableName.Visible = ObjRegistry.TableVisible
   Grid.Columns("TableName").Visible = ObjRegistry.TableVisible
   
   If ObjUserSecurity.IsAdministrator = False Then
      Grid.Columns("Amount").Visible = Not ObjRegistry.HideSaleAmount
   End If
   
   Dim vWidth As Long, i As Integer
   vWidth = 0
   For i = 0 To Grid.Cols - 1
      If Grid.Columns(i).Visible = True Then
         vWidth = vWidth + Grid.Columns(i).Width
      End If
   Next i
   Grid.Width = vWidth + 18
   
   DtpToDate.DateValue = ParaInReturnDate
   DtpFromDate.DateValue = DtpToDate.DateValue - vDateDiff
      
   If vDateDiff = 0 Then
      DtpToDate.Visible = False
      LblCustomerName.Left = LblCustomerName.Left - DtpToDate.Width
      TxtCustomerName.Left = TxtCustomerName.Left - DtpToDate.Width
      LblTtlAmount.Left = LblTtlAmount.Left - DtpToDate.Width
      TxtTTLAmount.Left = TxtTTLAmount.Left - DtpToDate.Width
      LblTableName.Left = LblTableName.Left - DtpToDate.Width
      TxtTableName.Left = TxtTableName.Left - DtpToDate.Width
      LblTag.Left = LblTag.Left - DtpToDate.Width
      TxtTag.Left = TxtTag.Left - DtpToDate.Width
      LblManualBillNo.Left = LblManualBillNo.Left - DtpToDate.Width
      TxtManualBillNo.Left = TxtManualBillNo.Left - DtpToDate.Width
   End If
   Me.ParaOutSID = -1
   Me.ParaOutReturnID = -1
   Me.ParaOutReturnDate = ""
   
   vOrder = " Order by Date Desc, ID"
   vDirection = " Desc"
   If vSearchInPreviousState = False Then
      vTag = ""
      vTtlAmount = ""
      vPartyName = ""
      vManualBillNo = ""
      vTableName = ""
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
      Case Grid.Name, TxtReturnID.Name, DtpFromDate.Name, DtpToDate.Name
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
  If Grid.rows > 0 Then BtnSelect_Click
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
      TxtReturnID.Text = Chr(KeyAscii): TxtReturnID.SelStart = Len(TxtReturnID.Text): TxtReturnID.SetFocus
   End Select
End Sub

Private Sub Image1_Click()
   Unload Me
End Sub

Private Sub TxtReturnID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtReturnID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "ID = " & TxtReturnID.Text, , adSearchForward, 1
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

Private Sub TxtTableName_Change()
   On Error GoTo ErrorHandler
   vTableName = " and TableName Like '%" & TxtTableName.Text & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTag_Change()
   On Error GoTo ErrorHandler
   vTag = " and Tag Like '%" & TxtTag.Text & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTTLAmount_Change()
   On Error GoTo ErrorHandler
   vTtlAmount = IIf(Val(TxtTTLAmount.Text) = 0, "", " and TotalAmount = " & Val(TxtTTLAmount.Text))
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
