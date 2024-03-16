VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form SchProduct 
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
   Begin VB.TextBox TxtPurchase2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6225
      TabIndex        =   20
      Top             =   2685
      Width           =   1485
   End
   Begin VB.TextBox TxtPubName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8985
      TabIndex        =   16
      Top             =   2685
      Width           =   3420
   End
   Begin VB.ComboBox CmbPublisher 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2025
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2010
      Width           =   1890
   End
   Begin VB.TextBox TxtPurchase 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6225
      TabIndex        =   10
      Top             =   2370
      Width           =   1485
   End
   Begin VB.TextBox TxtRetailPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7725
      TabIndex        =   1
      Top             =   2685
      Width           =   1260
   End
   Begin VB.TextBox TxtProductName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1320
      TabIndex        =   0
      Top             =   2685
      Width           =   4905
   End
   Begin VB.TextBox TxtProductID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   255
      TabIndex        =   5
      Top             =   2685
      Width           =   1065
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6195
      Left            =   255
      TabIndex        =   2
      Top             =   3015
      Width           =   14625
      ScrollBars      =   3
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
      stylesets(0).Picture=   "SchProduct.frx":0000
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
      Columns(0).Width=   1879
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ProductID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 1"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   8281
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "ProductName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 2"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Purchase Price"
      Columns(2).Name =   "PurchasePrice"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 5"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "List Price"
      Columns(3).Name =   "ListPrice"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 7"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "WS Price"
      Columns(4).Name =   "WholeSalePrice"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 5"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1773
      Columns(5).Caption=   "Retail Price"
      Columns(5).Name =   "Price"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 6"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3149
      Columns(6).Caption=   "Company Name"
      Columns(6).Name =   "CompanyName"
      Columns(6).DataField=   "Column 9"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3863
      Columns(7).Caption=   "Barcode"
      Columns(7).Name =   "Barcode"
      Columns(7).DataField=   "Column 6"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "ProductDesc"
      Columns(8).Name =   "ProductDesc"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1614
      Columns(9).Caption=   "Disc Price"
      Columns(9).Name =   "DiscPrice"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 7"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   25797
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
      MICON           =   "SchProduct.frx":001C
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
      MICON           =   "SchProduct.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblRack 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rack"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   11520
      TabIndex        =   21
      Top             =   1890
      Width           =   645
   End
   Begin VB.Label LblStockCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   11340
      TabIndex        =   19
      Top             =   990
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LblStock 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   11250
      TabIndex        =   18
      Top             =   1290
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Writer / Pub"
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
      Left            =   9000
      TabIndex        =   17
      Top             =   2475
      Width           =   1065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Writer / Pub "
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
      Left            =   2160
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   1125
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
      Left            =   270
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
      Left            =   6225
      TabIndex        =   11
      Top             =   2145
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
      Left            =   7725
      TabIndex        =   9
      Top             =   2460
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
      Left            =   255
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
      Caption         =   "Product Name"
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
      Left            =   1335
      TabIndex        =   6
      Top             =   2475
      Width           =   1215
   End
End
Attribute VB_Name = "SchProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSQL As String
Dim vWords() As String
Dim Flag As Boolean
Public ParaOutID As String, vCompanyIndex As Integer
Public ParaInWhere As String, ParaInPurchase As Boolean, ParaInWholeSale, ParainShowStock  As Boolean
Dim vProductName As String, vPubName As String, vRetailPrice As String, vPurPrice As String, vFlag As Boolean
Dim vOrder As String, vDirection As String, vCol As Byte
Dim vQtyLoose As Double
Dim vRoundFigure As Integer


Private Sub BtnClose_Click()
  Me.ParaOutID = ""
  Unload Me
  'Me.Hide
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.rows = 0 Then Exit Sub
  Me.ParaOutID = Rs!Productid
  Unload Me
  'Me.Hide
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub CmbPublisher_Click()
   On Error GoTo ErrorHandler
   If CmbPublisher.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbPublisher.Name Then Exit Sub
   Call LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FindComboIndex(ByVal cmb As ComboBox, ByVal result As String) As Boolean
   On Error GoTo ErrorHandler
   Dim i As Integer
   For i = 0 To cmb.ListCount - 1
      If result = cmb.ItemData(i) Then
         cmb.ListIndex = i
         FindComboIndex = True
         Exit Function
      End If
   Next i
   FindComboIndex = False
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectCompany() As Boolean
   On Error GoTo ErrorHandler
   SchCompany.Show vbModal, Me
   If SchCompany.ParaOutCompanyID = "" Then FunSelectCompany = False: Exit Function
   FunSelectCompany = FindComboIndex(CmbCompany, SchCompany.ParaOutCompanyID)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyEscape Then
      Call BtnClose_Click
   ElseIf KeyCode = vbKeyDown Then
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & IIf(ObjRegistry.StoreID = "", 1, ObjRegistry.StoreID) & " and Productid = " & Val(Grid.Columns("ProductID").Text)
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(Grid.Columns("ProductID").Text) & "," & IIf(ObjRegistry.StoreID = "", 1, ObjRegistry.StoreID) & ",0,0,0,0,0,0,'" & Date + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & Val(Grid.Columns("ProductID").Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(Grid.Columns("ProductID").Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Caption = LblStock.Caption & " " & " Total Qty: " & vQtyLoose
         LblStock.Visible = ParainShowStock
         LblStockCaption.Visible = ParainShowStock
   ElseIf KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtProductID.Name, TxtProductName.Name
         Call BtnSelect_Click
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      'Select Case ActiveControl.Name
      '   Case CmbCompany.Name:
         If FunSelectCompany() = True Then CmbCompany.SetFocus
      'End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
   
   
   
   vRoundFigure = Val(ObjRegistry.RoundfigureInSearchForm)
   CmbPublisher.Clear
   With cn.Execute("Select * FROM Publishers order by PubName")
      CmbPublisher.AddItem "All Publisher"
      CmbPublisher.ItemData(CmbPublisher.NewIndex) = Asc(Left("000", 1)) & Asc(Mid("000", 2, 1)) & Asc(Mid("000", 3, 1))
      Do Until .EOF
         CmbPublisher.AddItem !PubName
         CmbPublisher.ItemData(CmbPublisher.NewIndex) = Asc(Left(!PubID, 1)) & Asc(Mid(!PubID, 2, 1)) & Asc(Mid(!PubID, 3, 1))
         .MoveNext
      Loop
   End With
   CmbPublisher.ListIndex = 0
   
   With cn.Execute("Select * FROM Companies order by CompanyName")
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
      TxtPurchase2.Visible = True
   End If
   
   Grid.Columns("WholeSalePrice").Visible = Me.ParaInWholeSale
   If Me.ParaInWholeSale = True Then
      Grid.Columns("WholeSalePrice").Width = 90
      lblPurchase.Visible = True
      TxtPurchase.Visible = True
      TxtPurchase2.Visible = True
      lblPurchase.Caption = "WS Price"
   End If
   
   Grid.Columns("ListPrice").Visible = ObjRegistry.isShowListPrice
   If ObjRegistry.isShowListPrice Then
      Grid.Columns("PurchasePrice").Visible = True
      Grid.Columns("PurchasePrice").Width = 90
      Grid.Columns("ListPrice").Width = 90
      Grid.Columns("WholeSalePrice").Visible = True
      Grid.Columns("WholeSalePrice").Width = 90
   Else
      Grid.Columns("DiscPrice").Visible = True
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
   LblRack.Caption = ""
   Call LoadGrid
   If Me.ParaOutID <> "" Then Rs.Find "ProductID = " & Val(Me.ParaOutID), , adSearchForward, 1
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
  If Grid.rows > 0 Then Call BtnSelect_Click
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

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & IIf(ObjRegistry.StoreID = "", 1, ObjRegistry.StoreID) & " and Productid = " & Val(Grid.Columns("ProductID").Text)
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(Grid.Columns("ProductID").Text) & "," & IIf(ObjRegistry.StoreID = "", 1, ObjRegistry.StoreID) & ",0,0,0,0,0,0,'" & Date + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & Val(Grid.Columns("ProductID").Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(Grid.Columns("ProductID").Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Caption = LblStock.Caption & " " & " Total Qty: " & vQtyLoose
         LblStock.Visible = ParainShowStock
         LblStockCaption.Visible = ParainShowStock
                  
         vStrSQL = " SELECT isnull(RackName,'') " & vbCrLf _
           + " from Products p " & vbCrLf _
           + " left outer join Racks Rk on Rk.RackID = p.RackID " & vbCrLf _
           + " where productid = " & Val(Grid.Columns("ProductID").Text)
           With cn.Execute(vStrSQL)
            If .EOF = False Then LblRack.Caption = .Fields(0).Value
           End With
           
'         ''' latest Comment
End Sub

Private Sub Image1_Click()
   Unload Me
   'Me.Hide
End Sub

Private Sub TxtProductID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtProductID.Text) = "" Then Grid.MoveFirst: If Flag = False Then LoadGrid: Exit Sub
   If Len(TxtProductID.Text) < 6 Then
      If Flag = False Then LoadGrid
      Rs.Find "ProductID = " & Val(TxtProductID.Text), , adSearchForward, 1
      If Rs.EOF Then Grid.MoveLast
   Else
      LoadBarGrid
   End If
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
      vStrSQL = "SELECT P.Productid, ProductName, RackName, CompanyName, Desc1, Cast(isnull(ListPrice,0) as decimal(9," & vRoundFigure & ")) as ListPrice, cast(RetailPrice as decimal(9," & vRoundFigure & ")) RetailPrice, Cast(retailprice-isnull(discpc,0) as decimal(9," & vRoundFigure & ")) as DiscPrice, " & IIf(ObjRegistry.ShowDiscPurPrice = True, "cast(PurPrice-isnull(purDiscPC*isnull(multiplier,1),0) as decimal(9," & vRoundFigure & ")) as PurPrice ", " PurPrice") & ", Cast(WSPrice as decimal(9," & vRoundFigure & ")) WSPrice, '' as Stock, code " & vbCrLf _
            + " FROM products p left outer join publishers pub on Pub.Pubid = p.PubID left outer join (Select max(code) Code, productid from productbarcodes group by productid) b on b.productid = p.productid" & vbCrLf _
            + " left outer join Companies C on C.CompanyID = p.CompanyID " & vbCrLf & _
            " left outer join Racks Rk on Rk.RackID = p.RackID " & vbCrLf & _
            " left outer join productpacking pp on p.productid = pp.productid and p.PurchasePackingID = pp.packingid" & vbCrLf & _
            " where 1=1 " & IIf(CmbCompany.ListIndex > 0, " and C.CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & IIf(CmbPublisher.ListIndex > 0, " and P.pubid ='" & GetPubID(CmbPublisher) & "'", "") & vProductName & vPubName & vRetailPrice & vPurPrice & ParaInWhere & vOrder & vDirection
   Else
      vStrSQL = "SELECT P.Productid, ProductName, RackName, CompanyName, Desc1, cast(isnull(ListPrice,0) as decimal(9," & vRoundFigure & "))  as ListPrice, cast(RetailPrice as decimal(9," & vRoundFigure & ")) RetailPrice, cast(retailprice-isnull(discpc,0) as decimal(9," & vRoundFigure & ")) as DiscPrice, " & IIf(ObjRegistry.ShowDiscPurPrice = True, "cast(PurPrice-isnull(purDiscPC*isnull(multiplier,1),0) as decimal(9," & vRoundFigure & ")) as PurPrice ", " PurPrice") & ", cast(WSPrice as decimal(9" & vRoundFigure & ")) WSPrice, '' as Stock, '' as Code" & vbCrLf _
            + " FROM products p  " & vbCrLf _
            + " left outer join Companies C on C.CompanyID = p.CompanyID " & vbCrLf & _
            " left outer join Racks Rk on Rk.RackID = p.RackID " & vbCrLf & _
            " left outer join productpacking pp on p.productid = pp.productid and p.PurchasePackingID = pp.packingid" & vbCrLf & _
            " where 1=1 " & IIf(CmbCompany.ListIndex > 0, " and C.CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & IIf(CmbPublisher.ListIndex > 0, " and P.pubid ='" & GetPubID(CmbPublisher) & "'", "") & vProductName & vPubName & vRetailPrice & vPurPrice & ParaInWhere & vOrder & vDirection
   End If
   

   'VStrSQL = "SELECT * FROM SchProduct where productname like '%" & TxtProductName.Text & "%'" & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.CursorLocation = adUseClient
   Rs.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("ProductID").DataField = "ProductID"
   Grid.Columns("ProductName").DataField = "ProductName"
   Grid.Columns("CompanyName").DataField = "CompanyName"
   Grid.Columns("ProductDesc").DataField = "Desc1"
   Grid.Columns("Price").DataField = "RetailPrice"
   Grid.Columns("DiscPrice").DataField = "DiscPrice"
   Grid.Columns("PurchasePrice").DataField = "PurPrice"
   Grid.Columns("WholeSalePrice").DataField = "WSPrice"
   Grid.Columns("ListPrice").DataField = "ListPrice"
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
   vStrSQL = "SELECT p.Productid, ProductName, ListPrice, RetailPrice, retailprice-isnull(discpc,0) as DiscPrice, '' as Stock " & vbCrLf _
            + " FROM products p inner join ProductBarcodes b on p.ProductID = b.ProductID " & vbCrLf _
            + " where Code like '" & TxtProductID.Text & "%'" & vProductName & vPurPrice & vRetailPrice & ParaInWhere & vOrder & vDirection
   'VStrSQL = "SELECT * FROM SchProduct where productname like '%" & TxtProductName.Text & "%'" & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.CursorLocation = adUseClient
   Rs.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
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

Private Sub TxtPubName_Change()
vWords = Split(TxtPubName.Text, " ")
   vPubName = ""
   For i = 0 To UBound(vWords)
       vPubName = vPubName & " and Pubname like '%" & Replace(vWords(i), "'", "''") & "%'"
   Next
'   vProductName = " and Productname like '%" & Replace(TxtProductName.Text, "'", "''") & "%'"
   LoadGrid
End Sub

Private Sub TxtPurchase2_Change()
On Error GoTo ErrorHandler
   If Grid.Columns("WholeSalePrice").Visible Then
      vPurPrice = IIf(Val(TxtPurchase2.Text) = 0, "", " and WSPrice between " & Val(TxtPurchase.Text) & " And " & Val(TxtPurchase2.Text))
   Else
      vPurPrice = IIf(Val(TxtPurchase2.Text) = 0, "", " and PurPrice between " & Val(TxtPurchase.Text) & " And " & Val(TxtPurchase2.Text))
   End If
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
      vPurPrice = IIf(Val(TxtPurchase.Text) = 0, "", " and WSPrice between " & Val(TxtPurchase.Text) & " And " & Val(TxtPurchase2.Text))
   Else
      vPurPrice = IIf(Val(TxtPurchase.Text) = 0, "", " and PurPrice between " & Val(TxtPurchase.Text) & " And " & Val(TxtPurchase2.Text))
   End If
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function GetPubID(cmb As ComboBox) As String
    On Error GoTo ErrorHandler
    If cmb.ListIndex < 0 Then Exit Function
    GetPubID = Chr(Left(cmb.ItemData(cmb.ListIndex), 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 3, 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 5, 2))
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function
