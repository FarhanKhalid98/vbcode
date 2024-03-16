VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmMembersDiscount 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmMembersDiscount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbSubDepartment 
      Height          =   315
      Left            =   8010
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1305
      Width           =   1890
   End
   Begin VB.ComboBox cmbDepartment 
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1305
      Width           =   1890
   End
   Begin VB.ComboBox CmbSubGroup 
      Height          =   315
      Left            =   9728
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2018
      Width           =   1755
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   6173
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2018
      Width           =   1710
   End
   Begin VB.TextBox TxtDiscPer 
      Height          =   345
      Left            =   11528
      TabIndex        =   15
      Top             =   2018
      Width           =   705
   End
   Begin VB.TextBox TxtProductID 
      Height          =   345
      Left            =   1973
      TabIndex        =   12
      Top             =   1988
      Width           =   975
   End
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   4778
      TabIndex        =   11
      Top             =   1973
      Width           =   900
      _ExtentX        =   1588
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
      MICON           =   "FrmMembersDiscount.frx":0ECA
      BC              =   12632256
      FC              =   0
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   2963
      TabIndex        =   9
      Top             =   1988
      Width           =   1755
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   7913
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2018
      Width           =   1755
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5850
      Left            =   2700
      TabIndex        =   3
      Top             =   2610
      Width           =   10680
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
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
      stylesets(0).Picture=   "FrmMembersDiscount.frx":0EE6
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
      stylesets(1).Picture=   "FrmMembersDiscount.frx":0F02
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
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
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   106
      Columns.Count   =   6
      Columns(0).Width=   1852
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   8308
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2117
      Columns(2).Caption=   "Cost"
      Columns(2).Name =   "Cost"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   5
      Columns(2).NumberFormat=   "########.##"
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1852
      Columns(3).Caption=   "Retail Price"
      Columns(3).Name =   "RetailPrice"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   5
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   2090
      Columns(4).Caption=   "Profit Margin %"
      Columns(4).Name =   "ProfitMargin"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   5
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1561
      Columns(5).Caption=   "Disc %"
      Columns(5).Name =   "DiscPer"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   5
      Columns(5).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   18838
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
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6233
      TabIndex        =   4
      Top             =   8963
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
      MICON           =   "FrmMembersDiscount.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7538
      TabIndex        =   5
      Top             =   8963
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
      MICON           =   "FrmMembersDiscount.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10163
      TabIndex        =   6
      Top             =   8963
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
      MICON           =   "FrmMembersDiscount.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnApply 
      Height          =   315
      Left            =   12338
      TabIndex        =   14
      Top             =   2018
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      TX              =   "Apply"
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
      MICON           =   "FrmMembersDiscount.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label LblSubDepartment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Department"
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
      Left            =   8010
      TabIndex        =   22
      Top             =   1035
      Width           =   1380
   End
   Begin VB.Label LblDepartment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Left            =   6120
      TabIndex        =   21
      Top             =   1035
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubGroup Name"
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
      Left            =   9728
      TabIndex        =   18
      Top             =   1793
      Width           =   1395
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
      Left            =   6173
      TabIndex        =   17
      Top             =   1793
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc %"
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
      Left            =   11528
      TabIndex        =   16
      Top             =   1793
      Width           =   585
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
      Left            =   1973
      TabIndex        =   13
      Top             =   1748
      Width           =   930
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
      Left            =   2963
      TabIndex        =   10
      Top             =   1748
      Width           =   1215
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Members Discount"
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
      TabIndex        =   8
      Top             =   270
      Width           =   2460
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
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
      Left            =   7913
      TabIndex        =   7
      Top             =   1793
      Width           =   1065
   End
End
Attribute VB_Name = "FrmMembersDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim ssql As String, i As Integer

Private Sub BtnApply_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
      Grid.Columns("DiscPer").Value = Val(TxtDiscPer.Text)
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
  ssql = "Select * FROM MembersDiscount"
  Rs.Open ssql, cn, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  ssql = " SELECT p.ProductID, ProductName, Cost, RetailPrice, isnull(d.DiscPer,0) as DiscPer," & vbCrLf _
      + " cast(case when RetailPrice <> 0 then ((RetailPrice-Cost)/RetailPrice) * 100 else Cost end as numeric(7,3)) as ProfitMargin" & vbCrLf _
      + " FROM Products p inner join CurrentStock cs on cs.ProductID = p.ProductID" & vbCrLf _
      + " left outer join MembersDiscount d on p.productid = d.productid" & vbCrLf _
      + " where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text)) & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & " Order by p.ProductID --ProductName"
  
  With cn.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("Cost").Value = !Cost
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("ProfitMargin").Value = !ProfitMargin
        Grid.Columns("DiscPer").Value = !DiscPer
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

Private Sub CmbDepartment_Click()
   On Error GoTo ErrorHandler
   If cmbDepartment.Visible = False Then Exit Sub
   If ActiveControl.Name <> cmbDepartment.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbGroup_Click()
   On Error GoTo ErrorHandler
   If CmbGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbGroup.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub cmbSubDepartment_Click()
   On Error GoTo ErrorHandler
   If cmbSubDepartment.Visible = False Then Exit Sub
   If ActiveControl.Name <> cmbSubDepartment.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
   Call CmbGroup_Click
   'Call BtnFilter_Click
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   Grid.Update
'   Rs.MoveFirst
'   While Not Rs.EOF
'      If Rs.EditMode <> adEditNone Then
'         Call ActivityLog("Change Price", eEdit, , , Rs!ProductID)
'      End If
'      Rs.MoveNext
'   Wend
   Rs.Filter = ""
   Rs.UpdateBatch
   MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function GetGroupID(cmb As ComboBox) As String
    On Error GoTo ErrorHandler
    If cmb.ListIndex <= 0 Then Exit Function
    GetGroupID = Chr(Left(cmb.ItemData(cmb.ListIndex), 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 3, 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 5, 2))
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub PopulateGrid()
   On Error GoTo ErrorHandler
     If Rs.State = adStateOpen Then
     Rs.CancelBatch
     Rs.Close
   End If
   Me.MousePointer = vbHourglass
   
   Dim vStrSearch
   If ObjRegistry.isShowDepartment Then vStrSearch = IIf(cmbDepartment.ListIndex = 0, "", " and DepartmentID =" & cmbDepartment.ItemData(cmbDepartment.ListIndex))
   If ObjRegistry.isShowSubDepartment Then vStrSearch = vStrSearch & IIf(cmbSubDepartment.ListIndex = 0, "", " and SubDepartmentID =" & cmbSubDepartment.ItemData(cmbSubDepartment.ListIndex)) & vbCrLf _

   ssql = "Select * FROM MembersDiscount"
   Rs.Open ssql, cn, adOpenStatic, adLockBatchOptimistic
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
  
   ssql = " SELECT p.ProductID, ProductName, Cost, RetailPrice, isnull(d.DiscPer,0) as DiscPer," & vbCrLf _
      + " cast(case when RetailPrice <> 0 then ((RetailPrice-Cost)/RetailPrice) * 100 else Cost end as numeric(7,3)) as ProfitMargin" & vbCrLf _
      + " FROM Products p inner join CurrentStock cs on cs.ProductID = p.ProductID" & vbCrLf _
      + " left outer join MembersDiscount d on p.productid = d.productid" & vbCrLf _
      + " where 1=1 " & vStrSearch & IIf(CmbCompany.ListIndex = 0, "", " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & IIf(CmbGroup.ListIndex = 0, "", " and GroupID ='" & GetGroupID(CmbGroup) & "'") & IIf(CmbSubGroup.ListIndex = 0, "", " and SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & " Order by ProductName"
  
   With cn.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("Cost").Value = !Cost
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("ProfitMargin").Value = !ProfitMargin
        Grid.Columns("DiscPer").Value = !DiscPer
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
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub CmbSubGroup_Click()
   On Error GoTo ErrorHandler
   If CmbSubGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbSubGroup.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Members Discount"
   
   CmbCompany.Clear
   With cn.Execute("Select * FROM Companies Order By CompanyName")
      CmbCompany.AddItem "All Companies"
      CmbCompany.ItemData(CmbCompany.NewIndex) = 0
      Do Until .EOF
         CmbCompany.AddItem !CompanyName
         CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
         .MoveNext
      Loop
   End With
   
   CmbGroup.Clear
   With cn.Execute("Select * FROM Groups Order By GroupName")
      CmbGroup.AddItem "All Groups"
      CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left("000", 1)) & Asc(Mid("000", 2, 1)) & Asc(Mid("000", 3, 1))
      Do Until .EOF
         CmbGroup.AddItem !GroupName
         CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
   End With
   
     
   CmbSubGroup.Clear
   With cn.Execute("Select * FROM SubGroups Order By SubGroupName")
      CmbSubGroup.AddItem "All SubGroups"
      CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = 0
      Do Until .EOF
         CmbSubGroup.AddItem !SubGroupName
         CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = !SubGroupID
         .MoveNext
      Loop
   End With
      
   cmbDepartment.Visible = ObjRegistry.isShowDepartment
   LblDepartment.Visible = ObjRegistry.isShowDepartment
   cmbSubDepartment.Visible = ObjRegistry.isShowSubDepartment
   LblSubDepartment.Visible = ObjRegistry.isShowSubDepartment
   
   If ObjRegistry.isShowDepartment Then
      cmbDepartment.Clear
      With cn.Execute("Select * FROM Departments Order By Department")
         cmbDepartment.AddItem "All Departments"
         cmbDepartment.ItemData(cmbDepartment.NewIndex) = 0
         Do Until .EOF
            cmbDepartment.AddItem !Department
            cmbDepartment.ItemData(cmbDepartment.NewIndex) = !DepartmentID
            .MoveNext
         Loop
      End With
      cmbDepartment.ListIndex = 0
   End If
   If ObjRegistry.isShowSubDepartment Then
      cmbSubDepartment.Clear
      With cn.Execute("Select * FROM SubDepartments Order By SubDepartmentName")
         cmbSubDepartment.AddItem "All SubDepartments"
         cmbSubDepartment.ItemData(cmbSubDepartment.NewIndex) = 0
         Do Until .EOF
            cmbSubDepartment.AddItem !SubDepartmentName
            cmbSubDepartment.ItemData(cmbSubDepartment.NewIndex) = !SubDepartmentID
            .MoveNext
         Loop
      End With
      cmbSubDepartment.ListIndex = 0
   End If
   
   If ObjRegistry.isShowSubDepartment Or ObjRegistry.isShowDepartment Then
      CmbCompany.ListIndex = 0
   Else
      If CmbCompany.ListCount > 0 Then CmbCompany.ListIndex = 1 Else CmbCompany.ListIndex = 0
   End If
   CmbGroup.ListIndex = 0
   CmbSubGroup.ListIndex = 0
   PopulateGrid
'  Grid.Columns("Name").Locked = Not ObjUserSecurity.IsAdministrator
'  If CmbCompany.ListCount > 0 Then CmbCompany.ListIndex = 0
'  CmbGroup.ListIndex = 0
   
   'Call BtnFilter_Click
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
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   'If Grid.Visible = False Then Exit Sub
   'If ActiveControl.Name <> Grid.Name Then Exit Sub
   If Val(Grid.Columns("DiscPer").Value) = 0 Then Grid.Columns("DiscPer").Value = 0
   Rs.Filter = " ProductID = " & Val(Grid.Columns("ID").Text)
   If Rs.RecordCount = 0 And Val(Grid.Columns("DiscPer").Value) > 0 Then
      Rs.AddNew
      Rs!Productid = Grid.Columns("ID").Text
      Rs!DiscPer = Val(Grid.Columns("DiscPer").Value)
      Rs!isChanged = 0
   ElseIf Rs.RecordCount = 1 And Val(Grid.Columns("DiscPer").Value) = 0 Then
      Rs.Delete
   ElseIf Rs.RecordCount = 1 Then
      Rs!DiscPer = Val(Grid.Columns("DiscPer").Value)
      Rs!isChanged = 1
      Rs.Update
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_Change()
   If BtnSave.Enabled = False Then BtnSave.Enabled = True
End Sub

Private Sub Grid_GotFocus()
   Grid.Row = 0
   Grid.Col = 0
'   SendKeys "{Right}"
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
