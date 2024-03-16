VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form SchItemCode 
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
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2010
      Width           =   1890
   End
   Begin VB.TextBox TxtPurchase 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   9510
      TabIndex        =   10
      Top             =   2685
      Width           =   1485
   End
   Begin VB.TextBox TxtRetailPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8250
      TabIndex        =   1
      Top             =   2685
      Width           =   1260
   End
   Begin VB.TextBox TxtProductName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3345
      TabIndex        =   0
      Top             =   2685
      Width           =   4905
   End
   Begin VB.TextBox TxtProductID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1515
      TabIndex        =   5
      Top             =   2685
      Width           =   1830
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6195
      Left            =   1508
      TabIndex        =   2
      Top             =   3015
      Width           =   12480
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
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
      stylesets(0).Picture=   "SchItemCode.frx":0000
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
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   10
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ProductID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 1"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Item Code"
      Columns(1).Name =   "ItemCode"
      Columns(1).DataField=   "Column 7"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   8652
      Columns(2).Caption=   "Name"
      Columns(2).Name =   "ProductName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2196
      Columns(3).Caption=   "Retail Price"
      Columns(3).Name =   "Price"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 6"
      Columns(3).DataType=   5
      Columns(3).FieldLen=   256
      Columns(4).Width=   2196
      Columns(4).Caption=   "Disc Price"
      Columns(4).Name =   "DiscPrice"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 7"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "WS Price"
      Columns(5).Name =   "WholeSalePrice"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3995
      Columns(6).Caption=   "Barcode"
      Columns(6).Name =   "Barcode"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "Purchase Price"
      Columns(7).Name =   "PurchasePrice"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 5"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "ColourID"
      Columns(8).Name =   "ColourID"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "SizeID"
      Columns(9).Name =   "SizeID"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   22013
      _ExtentY        =   10927
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
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7740
      TabIndex        =   4
      Top             =   9848
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
      MICON           =   "SchItemCode.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   6435
      TabIndex        =   3
      Top             =   9848
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
      MICON           =   "SchItemCode.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1620
      TabIndex        =   13
      Top             =   1785
      Width           =   1320
   End
   Begin VB.Label lblPurchase 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Price"
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
      Left            =   9510
      TabIndex        =   11
      Top             =   2415
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Retail Price"
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
      Left            =   8250
      TabIndex        =   9
      Top             =   2415
      Width           =   1005
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   3000
      TabIndex        =   8
      Top             =   270
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   1515
      TabIndex        =   7
      Top             =   2460
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   14580
      Top             =   180
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   3345
      TabIndex        =   6
      Top             =   2430
      Width           =   495
   End
End
Attribute VB_Name = "SchItemCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim VStrSQL As String
Dim vWords() As String
Dim Flag As Boolean
Public ParaOutID, ParaOutItemCode, ParaOutColourID, ParaOutSizeID As String, vCompanyIndex As Integer
Public ParaInWhere As String, ParaInPurchase As Boolean, ParaInWholeSale As Boolean
Dim vItemCode, vProductName As String, vPubName As String, vRetailPrice As String, vPurPrice As String, vFlag As Boolean
Dim vOrder As String, vDirection As String, vCol As Byte

Private Sub BtnClose_Click()
  Me.ParaOutID = ""
  Me.ParaOutItemCode = ""
  Me.ParaOutColourID = ""
  Me.ParaOutSizeID = ""
  Unload Me
  'Me.Hide
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutID = Rs!Productid
  Me.ParaOutItemCode = Rs!ItemCode
  Me.ParaOutColourID = Rs!ColourID
  Me.ParaOutSizeID = Rs!SizeID
  Unload Me
  'Me.Hide
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtProductID.Name, TxtProductName.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub CmbCompany_Click()
   On Error GoTo ErrorHandler
   If CmbCompany.Visible = False Then Exit Sub
   vCompanyIndex = CmbCompany.ListIndex
   Call LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"
   
   With CN.Execute("Select * FROM Companies order by CompanyName")
      CmbCompany.AddItem "All Companies"
      CmbCompany.ItemData(CmbCompany.NewIndex) = 0
      Do Until .EOF
         CmbCompany.AddItem !CompanyName
         CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
         .MoveNext
      Loop
   End With
   
   If CmbCompany.ListCount > 0 Then
      If vCompanyIndex > 0 Then
         CmbCompany.ListIndex = vCompanyIndex
      Else
         CmbCompany.ListIndex = 0
      End If
   End If
   
'   Call LoadGrid
   
  
   vOrder = " order by productname"
'   Me.ParaOutID = ""
   vFlag = ObjRegistry.ProductSearchOpenInPreviousState
   Grid.Columns("PurchasePrice").Visible = Me.ParaInPurchase
   If Me.ParaInPurchase = True Then
      Grid.Columns("PurchasePrice").Width = 90
      lblPurchase.Visible = True
      TxtPurchase.Visible = True
   End If
   Grid.Columns("WholeSalePrice").Visible = Me.ParaInWholeSale
   If Me.ParaInWholeSale = True Then
      Grid.Columns("WholeSalePrice").Width = 90
      lblPurchase.Visible = True
      TxtPurchase.Visible = True
      lblPurchase.Caption = "WS Price"
   End If
   Grid.Columns("Barcode").Visible = ObjRegistry.ShowBarcodeProductSearch
   Dim vWidth As Long, i As Integer
   vWidth = 0
   For i = 0 To Grid.Cols - 1
      If Grid.Columns(i).Visible = True Then
         vWidth = vWidth + Grid.Columns(i).Width
      End If
   Next i
   Grid.Width = vWidth + 18

   Me.ParaInPurchase = False
   If vFlag = False Then
      vProductName = ""
      vRetailPrice = ""
      vPurPrice = ""
   End If
   Call LoadGrid
'   If Me.ParaOutID <> "" Then Rs.Find "ProductID ='" & Right("00000" + CStr(Val(Me.ParaOutID)), 5) & "'", , adSearchForward, 1
'   If Me.ParaOutItemCode <> "" Then Rs.Find "ItemCode ='" & Me.ParaOutItemCode, adSearchForward, 1
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
   Set Rs = Nothing
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then Call BtnSelect_Click
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
   vOrder = " ORDER By " & Grid.Columns(ColIndex).DataField
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
   Case Asc("a") To Asc("z"), Asc("A") To Asc("Z")
      TxtProductName.Text = Chr(KeyAscii): TxtProductName.SelStart = Len(TxtProductName.Text): TxtProductName.SetFocus:
   Case vbKey0 To vbKey9
      TxtProductID.Text = Chr(KeyAscii): TxtProductID.SelStart = Len(TxtProductID.Text): TxtProductID.SetFocus
   End Select
End Sub

Private Sub Image1_Click()
   Unload Me
   'Me.Hide
End Sub

Private Sub TxtProductID_Change()
  On Error GoTo ErrorHandler
   vWords = Split(TxtProductID.Text, " ")
   vItemCode = ""
   For i = 0 To UBound(vWords)
       vItemCode = vItemCode & " and ItemCode like '" & Replace(vWords(i), "'", "''") & "%'"
   Next
'   vProductName = " and Productname like '%" & Replace(TxtProductName.Text, "'", "''") & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtProductID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

'Private Sub TxtCode_Change()
'On Error GoTo ErrorHandler
'   If Trim(TxtCode.Text) = "" Then Grid.MoveFirst: Exit Sub
'   Rs.Find "Code like '" & TxtCode.Text & "%'", , adSearchForward, 1
'   If Rs.EOF Then Grid.MoveLast
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Set Rs = New ADODB.Recordset
   If vOrder = " ORDER By Productid" Then vOrder = " ORDER By P.Productid"
   'VStrSQL = "SELECT p.productid, Code, productname, retailprice,retailprice-discpc as DiscPrice FROM products p inner join (select * from ProductBarcodes where len(code)=7)b on p.productid = b.productid where productname like '%" & TxtProductName.Text & "%'" & vOrder & vDirection
   If ObjRegistry.ShowBarcodeProductSearch = True Then
      VStrSQL = "SELECT P.Productid, ProductName, RetailPrice, retailprice-isnull(discpc,0) as DiscPrice, PurPrice, WSPrice, '' as Stock, code, ItemCode, ColourID, SizeID" & vbCrLf _
            + " FROM products p eft outer join productcolours PC on PC.ProductID = P.ProductID left outer join ProductSizes PS on PS.ProductID = P.ProductID left outer join publishers pub on Pub.Pubid = p.PubID left outer join (Select max(code) Code, productid from productbarcodes group by productid) b on b.productid = p.productid" & vbCrLf _
            + " where 1=1 and itemcode is not null " & IIf(CmbCompany.ListIndex > 0, " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & vProductName & vPubName & vRetailPrice & vPurPrice & ParaInWhere & vOrder & vDirection
   Else
      VStrSQL = "SELECT P.Productid, ProductName, RetailPrice, retailprice-isnull(discpc,0) as DiscPrice, PurPrice, WSPrice, '' as Stock, '' as Code, ItemCode, ColourID, SizeID" & vbCrLf _
            + " FROM products p left outer join productcolours PC on PC.ProductID = P.ProductID left outer join ProductSizes PS on PS.ProductID = P.ProductID" & vbCrLf _
            + " where 1=1 and itemcode is not null " & IIf(CmbCompany.ListIndex > 0, " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & vItemCode & vProductName & vPubName & vRetailPrice & vPurPrice & ParaInWhere & vOrder & vDirection
   End If
   'VStrSQL = "SELECT * FROM SchProduct where productname like '%" & TxtProductName.Text & "%'" & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.CursorLocation = adUseClient
   Rs.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("ItemCode").DataField = "ItemCode"
   Grid.Columns("ProductID").DataField = "ProductID"
   Grid.Columns("ProductName").DataField = "ProductName"
   Grid.Columns("Price").DataField = "RetailPrice"
   Grid.Columns("DiscPrice").DataField = "DiscPrice"
   Grid.Columns("PurchasePrice").DataField = "PurPrice"
   Grid.Columns("WholeSalePrice").DataField = "WSPrice"
   Grid.Columns("ColourID").DataField = "ColourID"
   Grid.Columns("SizeID").DataField = "SizeID"
   Grid.Columns("Barcode").DataField = "Code"

   Flag = True
   'Grid.DataMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadBarGrid()
   On Error GoTo ErrorHandler
   'VStrSQL = "SELECT p.productid, Code, productname, retailprice,retailprice-discpc as DiscPrice FROM products p inner join (select * from ProductBarcodes where len(code)=7)b on p.productid = b.productid where productname like '%" & TxtProductName.Text & "%'" & vOrder & vDirection
   VStrSQL = "SELECT p.Productid, ProductName, RetailPrice, retailprice-isnull(discpc,0) as DiscPrice, '' as Stock " & vbCrLf _
            + " FROM products p inner join ProductBarcodes b on p.ProductID = b.ProductID " & vbCrLf _
            + " where Code like '" & TxtProductID.Text & "%'" & vProductName & vPurPrice & vRetailPrice & ParaInWhere & vOrder & vDirection
   'VStrSQL = "SELECT * FROM SchProduct where productname like '%" & TxtProductName.Text & "%'" & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.CursorLocation = adUseClient
   Rs.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   'Grid.DataMode
   Flag = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtProductName_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtProductName_Change()
   On Error GoTo ErrorHandler
   vWords = Split(TxtProductName.Text, " ")
   vProductName = ""
   For i = 0 To UBound(vWords)
       vProductName = vProductName & " and Productname like '%" & Replace(vWords(i), "'", "''") & "%'"
   Next
'   vProductName = " and Productname like '%" & Replace(TxtProductName.Text, "'", "''") & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtRetailPrice_Change()
   On Error GoTo ErrorHandler
   vRetailPrice = IIf(Val(TxtRetailPrice.Text) = 0, "", " and RetailPrice = " & Val(TxtRetailPrice.Text))
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtPurchase_Change()
   On Error GoTo ErrorHandler
   If Grid.Columns("WholeSalePrice").Visible Then
      vPurPrice = IIf(Val(TxtPurchase.Text) = 0, "", " and WSPrice = " & Val(TxtPurchase.Text))
   Else
      vPurPrice = IIf(Val(TxtPurchase.Text) = 0, "", " and PurPrice = " & Val(TxtPurchase.Text))
   End If
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
