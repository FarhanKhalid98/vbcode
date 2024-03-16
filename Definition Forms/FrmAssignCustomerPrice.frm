VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmAssignCustomerPrice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmAssignCustomerPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   6105
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1350
      Width           =   1665
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1350
      Width           =   1890
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   1230
      TabIndex        =   4
      Top             =   1320
      Width           =   1755
   End
   Begin VB.TextBox TxtProductID 
      Height          =   345
      Left            =   285
      TabIndex        =   3
      Top             =   1320
      Width           =   930
   End
   Begin VB.ComboBox CmbSortBy 
      Height          =   315
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1350
      Width           =   1170
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5850
      Left            =   0
      TabIndex        =   8
      Top             =   1755
      Width           =   11985
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   15
      stylesets.count =   2
      stylesets(0).Name=   "SelectedCol"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   12713983
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
      stylesets(0).Picture=   "FrmAssignCustomerPrice.frx":0ECA
      stylesets(1).Name=   "SelectedRow"
      stylesets(1).ForeColor=   16777215
      stylesets(1).BackColor=   8388608
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "FrmAssignCustomerPrice.frx":0EE6
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   2
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   106
      Columns.Count   =   15
      Columns(0).Width=   1138
      Columns(0).Caption=   "P ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   5715
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1614
      Columns(2).Caption=   "Packing"
      Columns(2).Name =   "Packing"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   767
      Columns(3).Caption=   "Mul"
      Columns(3).Name =   "Multiplier"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   5
      Columns(3).NumberFormat=   "########.##"
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1402
      Columns(4).Caption=   "Pur Price"
      Columns(4).Name =   "PurPrice"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   5
      Columns(4).NumberFormat=   "########.##"
      Columns(4).FieldLen=   256
      Columns(5).Width=   1402
      Columns(5).Caption=   "W Price"
      Columns(5).Name =   "WSPrice"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1402
      Columns(6).Caption=   "R Price"
      Columns(6).Name =   "RetailPrice"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1111
      Columns(7).Caption=   "Mrgn%"
      Columns(7).Name =   "Margin"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1270
      Columns(8).Caption=   "Disc/PC"
      Columns(8).Name =   "DiscPC"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1111
      Columns(9).Caption=   "Disc %"
      Columns(9).Name =   "DiscPer"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1217
      Columns(10).Caption=   "Min Lt."
      Columns(10).Name=   "MinStockLimit"
      Columns(10).Alignment=   1
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   5
      Columns(10).FieldLen=   256
      Columns(11).Width=   1217
      Columns(11).Caption=   "Max Lt."
      Columns(11).Name=   "MaxStockLimit"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   820
      Columns(12).Caption=   "Lock"
      Columns(12).Name=   "Lock"
      Columns(12).CaptionAlignment=   2
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(12).Style=   2
      Columns(13).Width=   1164
      Columns(13).Caption=   "NoCost"
      Columns(13).Name=   "NoCost"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(13).Style=   2
      Columns(14).Width=   820
      Columns(14).Caption=   "Raw"
      Columns(14).Name=   "Raw"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(14).Style=   2
      TabNavigation   =   1
      _ExtentX        =   21140
      _ExtentY        =   10319
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
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   3030
      TabIndex        =   5
      Top             =   1350
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      TX              =   "Filter"
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
      MICON           =   "FrmAssignCustomerPrice.frx":0F02
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnApply 
      Height          =   315
      Left            =   10770
      TabIndex        =   7
      Top             =   1350
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      TX              =   "Apply All"
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
      MICON           =   "FrmAssignCustomerPrice.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   9960
      TabIndex        =   6
      Top             =   1350
      Width           =   795
      _ExtentX        =   1402
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   7
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   4740
      TabIndex        =   16
      Top             =   8055
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
      MICON           =   "FrmAssignCustomerPrice.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   6060
      TabIndex        =   17
      Top             =   8055
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
      MICON           =   "FrmAssignCustomerPrice.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7380
      TabIndex        =   18
      Top             =   8055
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
      MICON           =   "FrmAssignCustomerPrice.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnMemberDisc 
      Height          =   420
      Left            =   3345
      TabIndex        =   19
      Top             =   8055
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   741
      TX              =   "Member Disc."
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
      MICON           =   "FrmAssignCustomerPrice.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
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
      Left            =   6105
      TabIndex        =   15
      Top             =   1125
      Width           =   1065
   End
   Begin VB.Label Label1 
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
      Left            =   4200
      TabIndex        =   14
      Top             =   1125
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1230
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
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
      TabIndex        =   12
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort BY"
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
      Left            =   7800
      TabIndex        =   11
      Top             =   1125
      Width           =   660
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Per"
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
      Left            =   9960
      TabIndex        =   10
      Top             =   1125
      Width           =   735
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allocate Product Price"
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
      Left            =   1890
      TabIndex        =   9
      Top             =   180
      Width           =   2910
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmAssignCustomerPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As New ADODB.Recordset
Public vSuppressUpdateEvent As Boolean
Dim ssql As String

Private Sub BtnApply_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
       Grid.Columns("DiscPer").Value = Val(TxtDiscPer.Text)
       Grid.Columns("DiscPc").Value = Round((Val(Grid.Columns("RetailPrice").Value) * Val(TxtDiscPer.Text) / 100), 2)
       Grid.MoveNext
   Next i
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnFilter_Click()
   On Error GoTo ErrorHandler
   'If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
Abc:
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  Me.MousePointer = vbHourglass
  ssql = "Select distinct p.ProductID, ProductName, PurPrice, WSPrice, RetailPrice, DiscPC, DiscPer, MinStockLimit, MaxStockLimit, isLocked, IsNoCostProduct, IsRawProduct FROM Products p left outer join ProductBarCodes b on p.productid = b.productid where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = '" & Right("00000" + CStr(Val(TxtProductID.Text)), 5) & "' or Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'")
  Rs.Open ssql, CN, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  ssql = "SELECT distinct p.ProductID, ProductName, PurPrice, WSPrice, RetailPrice, DiscPC, DiscPer, MinStockLimit, MaxStockLimit, isLocked, IsNoCostProduct, IsRawProduct FROM Products p left outer join ProductBarCodes b on p.productid = b.productid where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = '" & Right("00000" + CStr(Val(TxtProductID.Text)), 5) & "' or Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & " Order by p.ProductID --ProductName"
  With CN.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("PurPrice").Value = !PurPrice
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("WSPrice").Value = !WSPrice
        If (IsNull(!RetailPrice) Or !RetailPrice = 0) Then
           Grid.Columns("Margin").Value = 0
        Else
           Grid.Columns("Margin").Value = Round((IIf(IsNull(!RetailPrice), 0, !RetailPrice) - IIf(IsNull(!PurPrice), 0, !PurPrice)) * 100 / IIf(IsNull(!RetailPrice) Or !RetailPrice = 0, 1, !RetailPrice), 2)
        End If
        Grid.Columns("DiscPC").Value = !DiscPC
        Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
        Grid.Columns("MinStockLimit").Value = IIf(IsNull(!MinStockLimit), 0, !MinStockLimit)
        Grid.Columns("MaxStockLimit").Value = IIf(IsNull(!MaxStockLimit), 0, !MaxStockLimit)
        Grid.Columns("Lock").Value = Abs(!IsLocked)
        Grid.Columns("NoCost").Value = Abs(!IsNoCostProduct)
        Grid.Columns("Raw").Value = Abs(!IsRawProduct)
        Grid.Update
        .MoveNext
      Loop
  End With
  vSuppressUpdateEvent = False
  Grid.Redraw = True
  Grid.MoveFirst
  Me.MousePointer = vbDefault
  Exit Sub
ErrorHandler:
  If Err.Number = 91 Then GoTo Abc
  Grid.Redraw = True
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

Private Sub CmbCompany_Click()
   On Error GoTo ErrorHandler
   If CmbCompany.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'----------------------------------
Private Sub CmbGroup_Click()
   On Error GoTo ErrorHandler
   If CmbGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbGroup.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ReQuery()
   On Error GoTo ErrorHandler
   Dim vBm As Variant
   Dim i As Integer
    
   Rs.ReQuery
   Me.MousePointer = vbHourglass
   Grid.Redraw = False
   vSuppressUpdateEvent = True
   vBm = Grid.Bookmark
   Grid.MoveFirst
   For i = 0 To Grid.Rows - 1
      Rs!PurPrice = Grid.Columns("PurPrice").CellValue(Grid.GetBookmark(i))
      Rs!RetailPrice = Grid.Columns("RetailPrice").CellValue(Grid.GetBookmark(i))
      Rs!WSPrice = Grid.Columns("WSPrice").CellValue(Grid.GetBookmark(i))
      Rs!DiscPC = Grid.Columns("DiscPC").CellValue(Grid.GetBookmark(i))
      Rs!DiscPer = Grid.Columns("DiscPer").CellValue(Grid.GetBookmark(i))
      Rs!MinStockLimit = Grid.Columns("MinStockLimit").CellValue(Grid.GetBookmark(i))
      Rs!MaxStockLimit = Grid.Columns("MaxStockLimit").CellValue(Grid.GetBookmark(i))
      Rs!IsLocked = Val(Grid.Columns("Lock").CellValue(Grid.GetBookmark(i)))
      Rs!IsNoCostProduct = Val(Grid.Columns("NoCost").CellValue(Grid.GetBookmark(i)))
      Rs!IsRawProduct = Val(Grid.Columns("Raw").CellValue(Grid.GetBookmark(i)))
      Rs.Update
   Next i
   Grid.Bookmark = vBm
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Rs.UpdateBatch
   MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnMemberDisc_Click()
   FrmMembersDiscount.Show vbModal
End Sub

Private Sub PopulateGrid()
  On Error GoTo ErrorHandler
   If Rs.State = adStateOpen Then
     Rs.CancelBatch
     Rs.Close
   End If
   'CmbCompany.ListIndex = 0
   Me.MousePointer = vbHourglass
   Rs.Open "Select * FROM Products where 1=1 " & IIf(CmbGroup.ListIndex > 0, " and groupid ='" & GetGroupID(CmbGroup) & "'", "") & IIf(CmbCompany.ListIndex > 0, " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), ""), CN, adOpenStatic, adLockBatchOptimistic
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
  'sSql = " SELECT p.*, /*isnull(PackingName,'') as PackingName, isnull(Multiplier,'') as Multiplier,*/isnull(StockLimit,0) as StockLimit from Products p" & vbCrLf _
       + " /*left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId*/ where 1=1 and groupid = '" & GetGroupID(CmbGroup) & "'"
 
   ssql = " SELECT * from Products where 1=1 " & IIf(CmbGroup.ListIndex > 0, " and groupid ='" & GetGroupID(CmbGroup) & "'", "") & IIf(CmbCompany.ListIndex > 0, " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & " Order by " & CmbSortBy.Text
 
   With CN.Execute(ssql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("ID").Text = !Productid
         Grid.Columns("Name").Text = !ProductName
         'Grid.Columns("Packing").Text = !PackingName
         'Grid.Columns("Multiplier").Text = !Multiplier
         Grid.Columns("PurPrice").Value = !PurPrice
         Grid.Columns("RetailPrice").Value = !RetailPrice
         Grid.Columns("WSPrice").Value = !WSPrice
         If (IsNull(!RetailPrice) Or !RetailPrice = 0) Then
            Grid.Columns("Margin").Value = 0
         Else
            Grid.Columns("Margin").Value = Round((IIf(IsNull(!RetailPrice), 0, !RetailPrice) - IIf(IsNull(!PurPrice), 0, !PurPrice)) * 100 / IIf(IsNull(!RetailPrice) Or !RetailPrice = 0, 1, !RetailPrice), 2)
         End If
         Grid.Columns("DiscPC").Value = !DiscPC
         Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
         Grid.Columns("MinStockLimit").Value = IIf(IsNull(!MinStockLimit), 0, !MinStockLimit)
         Grid.Columns("MaxStockLimit").Value = IIf(IsNull(!MaxStockLimit), 0, !MaxStockLimit)
         Grid.Columns("Lock").Value = !IsLocked
         Grid.Columns("NoCost").Value = (!IsNoCostProduct)
         Grid.Columns("Raw").Value = (!IsRawProduct)
         Grid.Update
         .MoveNext
      Loop
   End With
   vSuppressUpdateEvent = False
   Grid.Redraw = True
   Grid.MoveFirst
   'If Grid.Visible Then Grid.SetFocus
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
  Call CmbGroup_Click
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   Grid.Update
   Rs.MoveFirst
   While Not Rs.EOF
      If Rs.EditMode <> adEditNone Then
         Call ActivityLog("Change Price", eEdit, , , Rs!Productid)
      End If
      Rs.MoveNext
   Wend
   Rs.UpdateBatch
   MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function GetGroupID(cmb As ComboBox) As String
    On Error GoTo ErrorHandler
    If cmb.ListIndex < 0 Then Exit Function
    GetGroupID = Chr(Left(cmb.ItemData(cmb.ListIndex), 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 3, 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 5, 2))
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub CmbSortBy_Click()
   On Error GoTo ErrorHandler
   PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Change Price"
   CmbSortBy.Clear
   CmbSortBy.AddItem "ProductID"
   CmbSortBy.AddItem "ProductName"
   CmbGroup.Clear
   CmbCompany.Clear
   With CN.Execute("Select * FROM Groups order by GroupName")
      CmbGroup.AddItem "All Groups"
      CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left("000", 1)) & Asc(Mid("000", 2, 1)) & Asc(Mid("000", 3, 1))
      Do Until .EOF
         CmbGroup.AddItem !GroupName
         CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
   End With
   With CN.Execute("Select * FROM Companies order by CompanyName")
      CmbCompany.AddItem "All Companies"
      CmbCompany.ItemData(CmbCompany.NewIndex) = 0
      Do Until .EOF
         CmbCompany.AddItem !CompanyName
         CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
         .MoveNext
      Loop
   End With
   Grid.Columns("Name").Locked = Not ObjUserSecurity.IsAdministrator
   If CmbCompany.ListCount > 0 Then CmbCompany.ListIndex = 1 Else CmbCompany.ListIndex = 0
   CmbGroup.ListIndex = 0
   CmbSortBy.ListIndex = 1
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name <> Grid.Name Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
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
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  'If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   Rs.Find "ProductID= '" & Grid.Columns("ID").Text & "'", , adSearchForward, 1
   If Rs.EOF Then MsgBox "Cannot Locate Record for updation. Please Try again", vbCritical, "Error": Cancel = True: Exit Sub
   Rs!ProductName = Grid.Columns("Name").Text
   Rs!PurPrice = Val(Grid.Columns("PurPrice").Value)
   Rs!WSPrice = Val(Grid.Columns("WSPrice").Value)
   Rs!RetailPrice = Val(Grid.Columns("RetailPrice").Value)
   Rs!DiscPC = Val(Grid.Columns("DiscPC").Value)
   Rs!DiscPer = Val(Grid.Columns("DiscPer").Value)
   Rs!MinStockLimit = Val(Grid.Columns("MinStockLimit").Value)
   Rs!MaxStockLimit = Val(Grid.Columns("MaxStockLimit").Value)
   Rs!IsLocked = Val(Grid.Columns("Lock").Value)
   Rs!IsNoCostProduct = Val(Grid.Columns("NoCost").Value)
   Rs!IsRawProduct = Val(Grid.Columns("Raw").Value)
   Rs.Update
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_Change()
   On Error GoTo ErrorHandler
   If Grid.Col = 4 Then 'Pur Price
       Grid.Columns(7).Value = Round((Val(Grid.Columns(6).Value) - Val(Grid.Columns(4).Value)) * 100 / IIf(Val(Grid.Columns(6).Value) = 0, 1, Val(Grid.Columns(6).Value)), 2)
   End If
   If Grid.Col = 6 Then 'Retail Price
       Grid.Columns(7).Value = Round((Val(Grid.Columns(6).Value) - Val(Grid.Columns(4).Value)) * 100 / IIf(Val(Grid.Columns(6).Value) = 0, 1, Val(Grid.Columns(6).Value)), 2)
       Grid.Columns(8).Value = Round((Val(Grid.Columns(6).Value) * Val(Grid.Columns(9).Value) / 100), 2)
   End If
   If Grid.Col = 9 Then 'DiscPer
       Grid.Columns(8).Value = Round((Val(Grid.Columns(6).Value) * Val(Grid.Columns(9).Value) / 100), 2)
   End If
   If Grid.Col = 8 Then 'DiscPc
       Grid.Columns(9).Value = Round((Val(Grid.Columns(8).Value) * 100) / Val(Grid.Columns(6).Value), 2)
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_GotFocus()
   Grid.Row = 0
   Grid.Col = 0
   SendKeys "{Right}"
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event vbKeyRight, 1, 1, 1
      KeyCode = 0
   End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub
