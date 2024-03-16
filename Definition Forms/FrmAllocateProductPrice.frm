VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmAllocateProductPrice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmAllocateProductPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtCustomerID 
      Height          =   345
      Left            =   9735
      TabIndex        =   16
      Top             =   1530
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   8625
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2475
      Width           =   1665
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2475
      Width           =   1890
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   3750
      TabIndex        =   4
      Top             =   2445
      Width           =   1755
   End
   Begin VB.TextBox TxtProductID 
      Height          =   345
      Left            =   2805
      TabIndex        =   3
      Top             =   2445
      Width           =   930
   End
   Begin VB.ComboBox CmbSortBy 
      Height          =   315
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2475
      Width           =   1170
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5850
      Left            =   1935
      TabIndex        =   6
      Top             =   2880
      Width           =   11490
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   10
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
      stylesets(0).Picture=   "FrmAllocateProductPrice.frx":0ECA
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
      stylesets(1).Picture=   "FrmAllocateProductPrice.frx":0EE6
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
      Columns.Count   =   10
      Columns(0).Width=   1588
      Columns(0).Caption=   "P ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   7223
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2275
      Columns(2).Caption=   "Packing"
      Columns(2).Name =   "Packing"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1111
      Columns(3).Caption=   "Mul"
      Columns(3).Name =   "Multiplier"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   5
      Columns(3).NumberFormat=   "########.##"
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "Pur Price"
      Columns(4).Name =   "PurPrice"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   5
      Columns(4).NumberFormat=   "########.##"
      Columns(4).FieldLen=   256
      Columns(5).Width=   1905
      Columns(5).Caption=   "Trade Price"
      Columns(5).Name =   "WSPrice"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "Disc/PC"
      Columns(6).Name =   "DiscPC"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   2143
      Columns(7).Caption=   "Allocated Price"
      Columns(7).Name =   "AllocatedPrice"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1482
      Columns(8).Caption=   "Disc %"
      Columns(8).Name =   "DiscPer"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1455
      Columns(9).Caption=   "Allocate"
      Columns(9).Name =   "Checked"
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(9).Style=   2
      TabNavigation   =   1
      _ExtentX        =   20267
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
      Left            =   5550
      TabIndex        =   5
      Top             =   2475
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
      MICON           =   "FrmAllocateProductPrice.frx":0F02
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   5723
      TabIndex        =   13
      Top             =   9180
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
      MICON           =   "FrmAllocateProductPrice.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7043
      TabIndex        =   14
      Top             =   9180
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
      MICON           =   "FrmAllocateProductPrice.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8363
      TabIndex        =   15
      Top             =   9180
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
      MICON           =   "FrmAllocateProductPrice.frx":0F56
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
      Left            =   8625
      TabIndex        =   12
      Top             =   2250
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
      Left            =   6720
      TabIndex        =   11
      Top             =   2250
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
      Left            =   3750
      TabIndex        =   10
      Top             =   2205
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
      Left            =   2790
      TabIndex        =   9
      Top             =   2205
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
      Left            =   10320
      TabIndex        =   8
      Top             =   2250
      Width           =   660
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
      Left            =   2700
      TabIndex        =   7
      Top             =   270
      Width           =   2910
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmAllocateProductPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As New ADODB.Recordset
Public vSuppressUpdateEvent As Boolean
Public ParaInCustomerID As String
Dim ssql As String

Private Sub BtnFilter_Click()
   On Error GoTo ErrorHandler
   'If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
Abc:
  TxtCustomerID.Text = ParaInCustomerID
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  Me.MousePointer = vbHourglass
  ssql = "Select * FROM CustomerProductPrice where CustomerID = '" & ParaInCustomerID & "' and " & IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text) & " or Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'")
  Rs.Open ssql, cn, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  ssql = "SELECT p.ProductID, ProductName, isnull(c.Price,WSPrice) as Price, WSPrice, isnull(c.DiscPer,0) as DiscPer, case when c.price is null then 0 else 1 end as Checked FROM Products p left outer join CustomerProductPrice c on p.ProductID = c.ProductID where CustomerID = '621' and " & IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text) & " or Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & " Order by p.ProductID --ProductName"
  With cn.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("AllocatedPrice").Value = !Price
        Grid.Columns("WSPrice").Value = !WSPrice
        Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
        Grid.Columns("Checked").Value = !Checked
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

Private Sub PopulateGrid()
  On Error GoTo ErrorHandler
   If Rs.State = adStateOpen Then
     Rs.CancelBatch
     Rs.Close
   End If
   'CmbCompany.ListIndex = 0
   Me.MousePointer = vbHourglass
   Rs.Open "Select * FROM CustomerProductPrice where CustomerID = " & Val(ParaInCustomerID), cn, adOpenStatic, adLockBatchOptimistic
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
  'sSql = " SELECT p.*, /*isnull(PackingName,'') as PackingName, isnull(Multiplier,'') as Multiplier,*/isnull(StockLimit,0) as StockLimit from Products p" & vbCrLf _
       + " /*left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId*/ where 1=1 and groupid = '" & GetGroupID(CmbGroup) & "'"
 
'   ssql = " SELECT * from Products where 1=1 " & IIf(CmbGroup.ListIndex > 0, " and groupid ='" & GetGroupID(CmbGroup) & "'", "") & IIf(CmbCompany.ListIndex > 0, " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & " Order by " & CmbSortBy.Text
   ssql = "SELECT p.ProductID, ProductName, isnull(c.Price,WSPrice) as Price, WSPrice, isnull(c.DiscPer,0) as DiscPer, case when c.price is null then 0 else 1 end as Checked FROM Products p left outer join (select * from CustomerProductPrice where CustomerID = '" & ParaInCustomerID & "' )c on p.ProductID = c.ProductID where 1=1 " & IIf(CmbGroup.ListIndex > 0, " and groupid ='" & GetGroupID(CmbGroup) & "'", "") & IIf(CmbCompany.ListIndex > 0, " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & " Order by " & CmbSortBy.Text

   With cn.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("AllocatedPrice").Value = !Price
        Grid.Columns("WSPrice").Value = !WSPrice
        Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
        Grid.Columns("Checked").Value = !Checked
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
   Rs.Filter = ""
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
   SetWindowText Me.hWnd, "Allocate Product Price"
   CmbSortBy.Clear
   CmbSortBy.AddItem "ProductID"
   CmbSortBy.AddItem "ProductName"
   CmbGroup.Clear
   CmbCompany.Clear
   With cn.Execute("Select * FROM Groups order by GroupName")
      CmbGroup.AddItem "All Groups"
      CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left("000", 1)) & Asc(Mid("000", 2, 1)) & Asc(Mid("000", 3, 1))
      Do Until .EOF
         CmbGroup.AddItem !GroupName
         CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
   End With
   With cn.Execute("Select * FROM Companies order by CompanyName")
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  'If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
  'MsgBox "Updated"
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
'   Rs.Find "ProductID= '" & Grid.Columns("ID").Text & "'", , adSearchForward, 1
'   If Rs.EOF Then MsgBox "Cannot Locate Record for updation. Please Try again", vbCritical, "Error": Cancel = True: Exit Sub
'   Rs!ProductName = Grid.Columns("Name").Text
'   Rs!PurPrice = Val(Grid.Columns("PurPrice").Value)
'   Rs!WSPrice = Val(Grid.Columns("WSPrice").Value)
'   Rs!RetailPrice = Val(Grid.Columns("RetailPrice").Value)
'   Rs!DiscPC = Val(Grid.Columns("DiscPC").Value)
'   Rs!DiscPer = Val(Grid.Columns("DiscPer").Value)
'   Rs!MinStockLimit = Val(Grid.Columns("MinStockLimit").Value)
'   Rs!MaxStockLimit = Val(Grid.Columns("MaxStockLimit").Value)
'   Rs!IsLocked = Val(Grid.Columns("Lock").Value)
'   Rs!IsNoCostProduct = Val(Grid.Columns("NoCost").Value)
'   Rs!IsRawProduct = Val(Grid.Columns("Raw").Value)
'   Rs.Update
   Rs.Filter = "ProductID = " & Val(Grid.Columns("ID").Text) & " and CustomerID = " & ParaInCustomerID
   If Rs.RecordCount = 0 And Abs(Grid.Columns("Checked").Value) = 1 Then
      Rs.AddNew
      Rs!CustomerID = Val(ParaInCustomerID)
      Rs!Productid = Val(Grid.Columns("ID").Text)
      Rs!Price = Val(Grid.Columns("AllocatedPrice").Text)
      Rs!DiscPer = Val(Grid.Columns("DiscPer").Value)
   ElseIf Rs.RecordCount = 1 And Abs(Grid.Columns("Checked").Value) = 0 Then
      Rs.Delete
   ElseIf Rs.RecordCount = 1 Then
      Rs!Price = Val(Grid.Columns("AllocatedPrice").Text)
      Rs!DiscPer = Val(Grid.Columns("DiscPer").Text)
      Rs.Update
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_Change()
   On Error GoTo ErrorHandler
   If Grid.Col = 9 Then
      If Abs(Grid.Columns("Checked").Value) = 1 Then
         Dim vZoneID As String, vTotal As Long, vZoneName As String, vSQL As String
         vZoneID = cn.Execute("Select dbo.FunGetZoneID(" & Val(ParaInCustomerID) & ")").Fields(0).Value
         vSQL = " Select count(*) as Total" & vbCrLf _
            + " From Products p inner join CustomerProductPrice cp on p.productid = cp.productid" & vbCrLf _
            + " inner join Parties pt on pt.partyid = cp.customerid" & vbCrLf _
            + " inner join sectors s on s.sectorid = pt.sectorid" & vbCrLf _
            + " inner join Zones z on z.zoneid = s.zoneid" & vbCrLf _
            + " where s.ZoneID = " & vZoneID & " and cp.ProductID = " & Val(Grid.Columns("ID").Text)
         
         vTotal = cn.Execute(vSQL).Fields(0).Value
         vSQL = "Select ZoneName from Zones where ZoneID = " & vZoneID
         vZoneName = cn.Execute(vSQL).Fields(0).Value
         If vTotal <> 0 Then
            If MsgBox("This Product is allocated to " & vTotal & " Customer(s) in " & vZoneName & " Zone ." & vbCrLf & "Do You want To Allocate more Customers in this Zone?", vbYesNo, "Allocate Product") = vbNo Then
               Grid.Columns("Checked").Value = 0
            End If
         End If
      End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub Grid_Change()
'   On Error GoTo ErrorHandler
'   If Grid.Col = 4 Then 'Pur Price
'       Grid.Columns(7).Value = Round((Val(Grid.Columns(6).Value) - Val(Grid.Columns(4).Value)) * 100 / IIf(Val(Grid.Columns(6).Value) = 0, 1, Val(Grid.Columns(6).Value)), 2)
'   End If
'   If Grid.Col = 6 Then 'Retail Price
'       Grid.Columns(7).Value = Round((Val(Grid.Columns(6).Value) - Val(Grid.Columns(4).Value)) * 100 / IIf(Val(Grid.Columns(6).Value) = 0, 1, Val(Grid.Columns(6).Value)), 2)
'       Grid.Columns(8).Value = Round((Val(Grid.Columns(6).Value) * Val(Grid.Columns(9).Value) / 100), 2)
'   End If
'   If Grid.Col = 9 Then 'DiscPer
'       Grid.Columns(8).Value = Round((Val(Grid.Columns(6).Value) * Val(Grid.Columns(9).Value) / 100), 2)
'   End If
'   If Grid.Col = 8 Then 'DiscPc
'       Grid.Columns(9).Value = Round((Val(Grid.Columns(8).Value) * 100) / Val(Grid.Columns(6).Value), 2)
'   End If
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub

Private Sub Grid_GotFocus()
   On Error GoTo ErrorHandler
   Grid.Row = 0
   Grid.Col = 0
   SendKeys "{Right}"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
