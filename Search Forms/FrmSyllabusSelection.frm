VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmSyllabusSelection 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmSyllabusSelection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbSchool 
      Height          =   2520
      ItemData        =   "FrmSyllabusSelection.frx":0ECA
      Left            =   3435
      List            =   "FrmSyllabusSelection.frx":0ECC
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Text            =   "CmbGroup"
      Top             =   2295
      Width           =   4995
   End
   Begin VB.ComboBox CmbClass 
      Height          =   2520
      Left            =   8505
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "CmbCompany"
      Top             =   2295
      Width           =   3060
   End
   Begin VB.TextBox TxtSyllabusName 
      Enabled         =   0   'False
      Height          =   345
      Left            =   6488
      TabIndex        =   2
      Top             =   1545
      Width           =   2385
   End
   Begin VB.TextBox TxtSyllabusID 
      Enabled         =   0   'False
      Height          =   345
      Left            =   5505
      TabIndex        =   1
      Top             =   1545
      Width           =   930
   End
   Begin VB.ComboBox CmbSortBy 
      Height          =   315
      Left            =   11895
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   1170
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4095
      Left            =   3645
      TabIndex        =   5
      Top             =   4905
      Width           =   7845
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
      stylesets(0).Picture=   "FrmSyllabusSelection.frx":0ECE
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
      stylesets(1).Picture=   "FrmSyllabusSelection.frx":0EEA
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
      Columns.Count   =   6
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
      Columns(2).Width=   1667
      Columns(2).Caption=   "Price"
      Columns(2).Name =   "Price"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1402
      Columns(3).Caption=   "Qty"
      Columns(3).Name =   "Qty"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1640
      Columns(4).Caption=   "Amount"
      Columns(4).Name =   "Amount"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1164
      Columns(5).Caption=   "Show"
      Columns(5).Name =   "Lock"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Style=   2
      TabNavigation   =   1
      _ExtentX        =   13838
      _ExtentY        =   7223
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
      Left            =   13095
      TabIndex        =   9
      Top             =   1485
      Visible         =   0   'False
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
      MICON           =   "FrmSyllabusSelection.frx":0F06
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6255
      TabIndex        =   15
      Top             =   10035
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
      MICON           =   "FrmSyllabusSelection.frx":0F22
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7575
      TabIndex        =   16
      Top             =   10035
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
      MICON           =   "FrmSyllabusSelection.frx":0F3E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8895
      TabIndex        =   17
      Top             =   10035
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
      MICON           =   "FrmSyllabusSelection.frx":0F5A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnAll 
      Height          =   315
      Left            =   9675
      TabIndex        =   8
      Top             =   9045
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      TX              =   "UnCheck All"
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
      MICON           =   "FrmSyllabusSelection.frx":0F76
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnQty 
      Height          =   315
      Left            =   4860
      TabIndex        =   7
      Top             =   9090
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
      MICON           =   "FrmSyllabusSelection.frx":0F92
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   4050
      TabIndex        =   6
      Top             =   9090
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      Left            =   3645
      TabIndex        =   18
      Top             =   9135
      Width           =   300
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Name"
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
      Left            =   3435
      TabIndex        =   14
      Top             =   2070
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class Name"
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
      Left            =   8505
      TabIndex        =   13
      Top             =   2070
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Syllabust Name"
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
      Left            =   5505
      TabIndex        =   12
      Top             =   1305
      Width           =   1320
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
      Left            =   11895
      TabIndex        =   11
      Top             =   1305
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Syllabus Selection"
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
      TabIndex        =   10
      Top             =   270
      Width           =   2430
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmSyllabusSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As New ADODB.Recordset
Public vSuppressUpdateEvent As Boolean
Public ParaOutID As String
Dim vMargin As String
Dim ssql As String

Private Sub BtnAll_Click()
   On Error GoTo ErrorHandler
   Dim vShow As Boolean
   If BtnAll.Caption = "UnCheck All" Then
      BtnAll.Caption = "Check All"
      vShow = False
   Else
      BtnAll.Caption = "UnCheck All"
      vShow = True
   End If
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
       Grid.Columns("Lock").Value = Abs(vShow)
       Grid.MoveNext
   Next i
   Grid.Redraw = True
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub



Private Sub BtnFilter_Click()
   On Error GoTo ErrorHandler
   'If ActiveControl.Name <> CmbClass.Name Then Exit Sub
Abc:
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  Me.MousePointer = vbHourglass
'  ssql = "Select distinct p.ProductID, ProductName, Price, Qty, Qty, DiscPC, DiscPer, MinStockLimit, MaxStockLimit, isLocked, IsNoCostProduct, IsRawProduct FROM Products p inner join Syllabus b on p.productid = b.productid where 1=1 " & IIf(TxtSyllabusID.Text = "", "", " and p.ProductID = '" & Right("00000" + CStr(Val(TxtSyllabusID.Text)), 5) & "' or Code = '" & Val(TxtSyllabusID.Text) & "'") & IIf(Trim(TxtSyllabusName.Text) = "", "", " and ProductName like '%" & TxtSyllabusName.Text & "%'")
'  Rs.Open ssql, CN, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  
  
   ssql = " Select H.syllabusID, H.SyllabusName, b.code, b.productid, Productname, QtyLoose, Price, Amount, isShow  from syllabusheader H " & vbCrLf _
       + " inner join SyllabusBody b on b.syllabusID = h.syllabusID inner join Products P on P.ProductID = b.ProductID Where 1=1 and H.SyllabusID = " & TxtSyllabusID.Text
 
   With CN.Execute(ssql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("ID").Text = !Productid
         Grid.Columns("Name").Text = !ProductName
         Grid.Columns("Price").Value = !Price
         Grid.Columns("Qty").Value = !QtyLoose
         Grid.Columns("Amount").Value = !Amount
'
         Grid.Columns("Lock").Value = 1
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

Private Sub BtnQty_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
      Grid.Columns("Qty").Value = Val(TxtQty.Text)
      Grid.Columns("Amount").Value = Grid.Columns("Qty").Value * Grid.Columns("Price").Value
      Grid.MoveNext
   Next i
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbClass_Click()
   On Error GoTo ErrorHandler
   If CmbClass.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbClass.Name Then Exit Sub
   PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'----------------------------------
Private Sub CmbSchool_Click()
   On Error GoTo ErrorHandler
   If CmbSchool.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbSchool.Name Then Exit Sub
   GetClasses
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
      Rs!Price = Grid.Columns("Price").CellValue(Grid.GetBookmark(i))
      Rs!Qty = Grid.Columns("Qty").CellValue(Grid.GetBookmark(i))
      Rs!Qty = Grid.Columns("Qty").CellValue(Grid.GetBookmark(i))
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
'   FrmMembersDiscount.Show vbModal
End Sub


Private Sub BtnClear_Click()
  If CmbSchool.ListCount = 0 Then Exit Sub
  Call GetClasses
  Call PopulateGrid
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   Grid.Update
   Me.ParaOutID = Trim(TxtSyllabusID.Text)
   If TxtSyllabusID.Text = "" Then
      If Rs.State = adStateOpen Then
         Rs.CancelBatch
         Rs.Close
      End If
      Unload Me
      Exit Sub
   End If
   Rs.MoveFirst
   While Not Rs.EOF
'      If Rs.EditMode <> adEditNone Then
'         Call ActivityLog("Change Price", eEdit, , , Rs!Productid)
'      End If
      Rs.MoveNext
   Wend
   Rs.UpdateBatch
   
   Unload Me
'   MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
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
'   PopulateGrid
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
   CmbSchool.Clear
   CmbClass.Clear
   With CN.Execute("Select SchoolID, SchoolName FROM Schools  ")
'      CmbSchool.AddItem "All Groups"
'      CmbSchool.ItemData(CmbSchool.NewIndex) = Asc(Left("000", 1)) & Asc(Mid("000", 2, 1)) & Asc(Mid("000", 3, 1))
      Do Until .EOF
         CmbSchool.AddItem !SchoolName
         CmbSchool.ItemData(CmbSchool.NewIndex) = !Schoolid
'         CmbSchool.ItemData(CmbSchool.NewIndex) = Asc(Left(!Schoolid, 1)) & Asc(Mid(!Schoolid, 2, 1)) & Asc(Mid(!Schoolid, 3, 1))
         .MoveNext
      Loop
   End With
   If CmbSchool.ListCount > 0 Then CmbSchool.ListIndex = 0 Else Exit Sub
   CmbSortBy.ListIndex = 1
   GetClasses
   Call PopulateGrid
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
  If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   Rs.Find "ProductID= '" & Grid.Columns("ID").Text & "'", , adSearchForward, 1
   If Rs.EOF Then MsgBox "Cannot Locate Record for updation. Please Try again", vbCritical, "Error": Cancel = True: Exit Sub
   Rs!isShow = Val(Grid.Columns("Lock").Value)
   Rs!QtyLoose = Val(Grid.Columns("Qty").Value)
   Rs!Amount = Val(Grid.Columns("Amount").Value)
   Rs.Update
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_Change()
   On Error GoTo ErrorHandler
   If Grid.Col = 4 Then Exit Sub
'   If Grid.Col = 4 Then 'Pur Price
'       Grid.Columns(7).Value = Round((Val(Grid.Columns(6).Value) - Val(Grid.Columns(4).Value)) * 100 / IIf(Val(Grid.Columns(6).Value) = 0, 1, Val(Grid.Columns(6).Value)), 2)
'   End If
'   If Grid.Col = 5 Then 'T Price
'       Grid.Columns(7).Value = Round((Val(Grid.Columns(5).Value) - Val(Grid.Columns(4).Value)) * 100 / IIf(Val(Grid.Columns(5).Value) = 0, 1, Val(Grid.Columns(5).Value)), 2)
'       Grid.Columns(8).Value = Round((Val(Grid.Columns(5).Value) * Val(Grid.Columns(9).Value) / 100), 2)
'       If Val(Grid.Columns(5).Value) = 0 Then Grid.Columns(7).Value = 0
'   End If
'   If Grid.Col = 6 Then 'Retail Price
'       Grid.Columns(7).Value = Round((Val(Grid.Columns(6).Value) - (Val(Grid.Columns(4).Value / IIf(Val(Grid.Columns(3).Value) = 0, 1, Val(Grid.Columns(3).Value))))) * 100 / IIf(Val(Grid.Columns(6).Value) = 0, 1, Val(Grid.Columns(6).Value)), 2)
'       Grid.Columns(8).Value = Round((Val(Grid.Columns(6).Value) * Val(Grid.Columns(9).Value) / 100), 2)
'   End If
'   If Grid.Col = 9 Then 'DiscPer
'       Grid.Columns(8).Value = Round((Val(Grid.Columns(6).Value) * Val(Grid.Columns(9).Value) / 100), 2)
'   End If
'   If Grid.Col = 8 Then 'DiscPc
'       Grid.Columns(9).Value = Round((Val(Grid.Columns(8).Value) * 100) / Val(Grid.Columns(6).Value), 2)
'   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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


Private Sub PopulateGrid()
   On Error GoTo ErrorHandler
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
   'CmbClass.ListIndex = 0
   TxtSyllabusID.Text = ""
   TxtSyllabusName.Text = ""
   Me.MousePointer = vbHourglass
   If CmbClass.ListIndex >= 0 Then
      ssql = " Select H.syllabusID, H.SyllabusName, b.code, b.productid, Productname, QtyLoose, Price, Amount, isShow  from syllabusheader H " & vbCrLf _
      + "inner join Schools S on S.SchoolID = H.SchoolID Inner join Classes C on C.ClassID = H.ClassID  " & vbCrLf _
       + " inner join SyllabusBody b on b.syllabusID = h.syllabusID inner join Products P on P.ProductID = b.ProductID Where S.SchoolID = " & CmbSchool.ItemData(CmbSchool.ListIndex) & " and C.ClassID =" & CmbClass.ItemData(CmbClass.ListIndex)
   Else
      ssql = " Select H.syllabusID, H.SyllabusName, b.code, b.productid, Productname, QtyLoose, Price, Amount, isShow  from syllabusheader H " & vbCrLf _
      + "inner join Schools S on S.SchoolID = H.SchoolID Inner join Classes C on C.ClassID = H.ClassID  " & vbCrLf _
       + " inner join SyllabusBody b on b.syllabusID = h.syllabusID inner join Products P on P.ProductID = b.ProductID Where S.SchoolID = " & CmbSchool.ItemData(CmbSchool.ListIndex)
   End If
   
   With CN.Execute(ssql)
      Do Until .EOF
         TxtSyllabusID.Text = !syllabusid
         TxtSyllabusName = !SyllabusName
         Grid.AddNew
         Grid.Columns("ID").Text = !Productid
         Grid.Columns("Name").Text = !ProductName
         Grid.Columns("Price").Value = !Price
         Grid.Columns("Qty").Value = !QtyLoose
         Grid.Columns("Amount").Value = !Amount
         Grid.Columns("Lock").Value = !isShow
         Grid.Update
         .MoveNext
      Loop
   End With
   If Rs.State = adStateOpen Then
     Rs.CancelBatch
     Rs.Close
   End If
   Rs.Open "Select * FROM SyllabusBody where 1=1 and SyllabusID = " & Val(TxtSyllabusID.Text), CN, adOpenStatic, adLockBatchOptimistic
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


Private Sub GetClasses()
   On Error GoTo ErrorHandler
      CmbClass.Clear
      ssql = "select C.* from syllabusHeader H inner join Schools S on S.SchoolID = H.SchoolID Inner join Classes C on C.ClassID = H.ClassID  Where S.SchoolID = " & CmbSchool.ItemData(CmbSchool.ListIndex)
      With CN.Execute(ssql)
'      CmbClass.AddItem "All Companies"
'      CmbClass.ItemData(CmbClass.NewIndex) = 0
      Do Until .EOF
         CmbClass.AddItem !ClassName
         CmbClass.ItemData(CmbClass.NewIndex) = !ClassID
         .MoveNext
      Loop
   End With
   Grid.Columns("Name").Locked = Not ObjUserSecurity.IsAdministrator
   If CmbClass.ListCount > 0 Then CmbClass.ListIndex = 0 'Else CmbClass.ListIndex = 0
    Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

