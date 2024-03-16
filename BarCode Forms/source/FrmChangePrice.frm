VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmChangePrice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmChangePrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrmHistory 
      Height          =   1275
      Left            =   135
      TabIndex        =   11
      Top             =   7650
      Visible         =   0   'False
      Width           =   3915
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridHistory 
         Height          =   1050
         Left            =   45
         TabIndex        =   12
         Top             =   135
         Width           =   3780
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   4
         stylesets.count =   3
         stylesets(0).Name=   "SelectedCol"
         stylesets(0).ForeColor=   0
         stylesets(0).BackColor=   12713983
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "FrmChangePrice.frx":0ECA
         stylesets(1).Name=   "Select"
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
         stylesets(1).Picture=   "FrmChangePrice.frx":0EE6
         stylesets(2).Name=   "SelectedRow"
         stylesets(2).ForeColor=   16777215
         stylesets(2).BackColor=   8388608
         stylesets(2).HasFont=   -1  'True
         BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(2).Picture=   "FrmChangePrice.frx":0F02
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
         Columns.Count   =   4
         Columns(0).Width=   1588
         Columns(0).Caption=   "PurPrice"
         Columns(0).Name =   "PurPrice"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1746
         Columns(1).Caption=   "Retail Price"
         Columns(1).Name =   "RetailPrice"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1244
         Columns(2).Caption=   "Margin"
         Columns(2).Name =   "Margin"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1614
         Columns(3).Caption=   "Margin Per"
         Columns(3).Name =   "MarginPer"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   6667
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "History"
         BackColor       =   15724527
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1395
      Width           =   2835
   End
   Begin VB.TextBox TxtProductName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1815
      TabIndex        =   1
      Top             =   1425
      Width           =   3255
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   8610
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1395
      Width           =   2835
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5850
      Left            =   90
      TabIndex        =   2
      Top             =   1755
      Width           =   11865
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
      stylesets(0).Picture=   "FrmChangePrice.frx":0F1E
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
      stylesets(1).Picture=   "FrmChangePrice.frx":0F3A
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
      Columns.Count   =   10
      Columns(0).Width=   1852
      Columns(0).Caption=   "Product ID"
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
      Columns(2).Width=   1852
      Columns(2).Caption=   "Packing"
      Columns(2).Name =   "Packing"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1349
      Columns(3).Caption=   "Multiplier"
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
      Columns(5).Width=   1376
      Columns(5).Caption=   "List Price"
      Columns(5).Name =   "ListPrice"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1667
      Columns(6).Caption=   "Retail Price"
      Columns(6).Name =   "RetailPrice"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1482
      Columns(7).Caption=   "Retail/PC"
      Columns(7).Name =   "RetailPerPC"
      Columns(7).Alignment=   1
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1508
      Columns(8).Caption=   "Disc/PC"
      Columns(8).Name =   "DiscPC"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1693
      Columns(9).Caption=   "Stock Limit"
      Columns(9).Name =   "StockLimit"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   5
      Columns(9).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20929
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
      Left            =   4335
      TabIndex        =   3
      Top             =   8115
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
      MICON           =   "FrmChangePrice.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   5640
      TabIndex        =   4
      Top             =   8115
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
      MICON           =   "FrmChangePrice.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8265
      TabIndex        =   5
      Top             =   8115
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
      MICON           =   "FrmChangePrice.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5775
      TabIndex        =   10
      Top             =   1170
      Width           =   1125
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Price"
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
      Height          =   360
      Index           =   0
      Left            =   1890
      TabIndex        =   8
      Top             =   180
      Width           =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1815
      TabIndex        =   7
      Top             =   1200
      Width           =   1020
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8625
      TabIndex        =   6
      Top             =   1170
      Width           =   900
   End
End
Attribute VB_Name = "FrmChangePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As New ADODB.Recordset
Public vSuppressUpdateEvent As Boolean
Dim sSql As String

Private Sub CmbCompany_Click()
  On Error GoTo ErrorHandler
  If CmbCompany.Visible = False Then Exit Sub
  If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
Abc:
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  CmbGroup.ListIndex = -1
  Me.MousePointer = vbHourglass
  Rs.Open "Select * FROM Products where 1=1 and Companyid = " & CmbCompany.ItemData(CmbCompany.ListIndex), CN, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  sSql = " SELECT p.*, isnull(PackingName,'') as PackingName, isnull(Multiplier,'') as Multiplier, isnull(StockLimit,0) as StockLimit from Products p" & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 and Companyid = " & CmbCompany.ItemData(CmbCompany.ListIndex) & " order by ProductName"
  With CN.Execute(sSql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("PurPrice").Value = !PurPrice
        Grid.Columns("ListPrice").Value = !ListPrice
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("RetailPerPc").Value = IIf(IsNull(!RetailPerPC), 0, !RetailPerPC)
        Grid.Columns("DiscPC").Value = !DiscPC
        Grid.Columns("StockLimit").Value = IIf(IsNull(!StockLimit), 0, !StockLimit)
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

'----------------------------------
Private Sub CmbGroup_Click()
  On Error GoTo ErrorHandler
  If ActiveControl.Name <> CmbGroup.Name Then Exit Sub
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  Me.MousePointer = vbHourglass
  CmbCompany.ListIndex = -1
  Rs.Open "Select * FROM Products where 1=1 and groupid ='" & GetGroupID(CmbGroup) & "'", CN, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
   sSql = " SELECT p.*, isnull(PackingName,'') as PackingName, isnull(Multiplier,'') as Multiplier, isnull(StockLimit,0) as StockLimit from Products p" & vbCrLf _
        + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
        + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 and GroupID = '" & GetGroupID(CmbGroup) & "'"
   With CN.Execute(sSql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("Packing").Text = !Packingname
        Grid.Columns("Multiplier").Text = !Multiplier
        Grid.Columns("PurPrice").Value = !PurPrice
        Grid.Columns("ListPrice").Value = !ListPrice
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("RetailPerPc").Value = !RetailPerPC
        Grid.Columns("DiscPC").Value = !DiscPC
        Grid.Columns("StockLimit").Value = !StockLimit
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
  FrmHistory.Visible = False
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   Grid.Update
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

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hwnd, "Change Price"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   With CN.Execute("Select * FROM Groups order by GroupName")
      Do Until .EOF
         CmbGroup.AddItem !GroupName
         CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
   End With
   With CN.Execute("Select * FROM Companies order by CompanyName")
      Do Until .EOF
          CmbCompany.AddItem !CompanyName
          CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
          .MoveNext
      Loop
   End With
   If CmbCompany.ListCount > 0 Then CmbCompany.ListIndex = 0
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
   End If
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  'If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
  On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   Rs.Find "ProductID = '" & Grid.Columns("ID").Text & "'", , adSearchForward, 1
   If Rs.EOF Then MsgBox "Cannot Locate Record for updation. Please Try again", vbCritical, "Error": Cancel = True: Exit Sub
   Rs!ProductName = Grid.Columns("Name").Text
   Rs!PurPrice = Grid.Columns("PurPrice").Value
   Rs!ListPrice = Grid.Columns("ListPrice").Value
   Rs!RetailPrice = Grid.Columns("RetailPrice").Value
   Rs!RetailPerPC = IIf(Grid.Columns("RetailPerPc").Value = "", Null, Grid.Columns("RetailPerPc").Value)
   Rs!DiscPC = Grid.Columns("DiscPC").Value
   Rs!StockLimit = Grid.Columns("StockLimit").Value
   Rs.Update
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Grid_GotFocus()
    On Error GoTo ErrorHandler
    Grid.Row = 0
    Grid.Col = 0
    'SendKeys "{Right}"
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

Private Sub PopulateDataToHistoryGrid()
    Dim PSQL As String
    PSQL = "select top 1 isnull(b.Price,PurPrice) as Price from Products p left outer join Purchasebody b on p.ProductID = b.ProductID" & vbCrLf & _
    " left outer join PurchaseHeader h  on h.PurchaseID = b.PurchaseID " & vbCrLf & _
    " where p.productid = '" & Grid.Columns("ID").Text & "' and isnull(PurchaseDate,'01-01-2010')  < Getdate() order by PurchaseDate Desc"
    
    sSql = "select top 1 isnull(b.Price,RetailPrice) as Price from Products p left outer join Salebody b on p.ProductID = b.ProductID" & vbCrLf & _
    " left outer join SaleHeader h  on h.SaleID = b.SaleID " & vbCrLf & _
    " where p.productid = '" & Grid.Columns("ID").Text & "' and isnull(SaleDate ,'01-01-2010') < Getdate() order by SaleDate Desc"
    
    GridHistory.Redraw = False
    GridHistory.MoveFirst
    GridHistory.RemoveAll
    GridHistory.AllowAddNew = True
    GridHistory.AddNew
    With CN.Execute(PSQL)
        If .RecordCount > 0 Then GridHistory.Columns("PurPrice").Value = CN.Execute(PSQL).Fields(0).Value
        .Close
    End With
    With CN.Execute(sSql)
        If .RecordCount > 0 Then GridHistory.Columns("RetailPrice").Value = CN.Execute(sSql).Fields(0).Value
        .Close
    End With
    GridHistory.Columns("Margin").Value = Val(GridHistory.Columns("RetailPrice").Value) - Val(GridHistory.Columns("PurPrice").Value)
    If Val(GridHistory.Columns("RetailPrice").Value) <> 0 Then
        GridHistory.Columns("MarginPer").Value = Round((Val(GridHistory.Columns("RetailPrice").Value) - Val(GridHistory.Columns("PurPrice").Value)) / GridHistory.Columns("RetailPrice").Value * 100, 2)
    End If
    GridHistory.Redraw = True
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    PopulateDataToHistoryGrid
    FrmHistory.Visible = True
    FrmHistory.ZOrder 0
    GridHistory.Visible = True
    GridHistory.ZOrder 0
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtProductName_Change()
  On Error GoTo ErrorHandler
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  'Me.MousePointer = vbHourglass
 
  Rs.Open "Select * FROM Products where 1=1 and groupid = '" & GetGroupID(CmbGroup) & "'", CN, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  sSql = " SELECT p.*, isnull(PackingName,'') as PackingName, isnull(Multiplier,'') as Multiplier, isnull(StockLimit,0) as StockLimit from Products p" & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 and groupid = '" & GetGroupID(CmbGroup) & "'" & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'")
       ' & IIf(CmbGroup.ListIndex > 0, " and groupid ='" & GetGroupID(CmbGroup) & "'", "")
   With CN.Execute(sSql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("Packing").Text = !Packingname
        Grid.Columns("Multiplier").Text = !Multiplier
        Grid.Columns("ListPrice").Value = !ListPrice
        Grid.Columns("PurPrice").Value = !PurPrice
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("DiscPC").Value = !DiscPC
        Grid.Columns("StockLimit").Value = !StockLimit
        Grid.Update
        .MoveNext
      Loop
  End With
   vSuppressUpdateEvent = False
  Grid.Redraw = True
  Grid.MoveFirst
  'Grid.SetFocus
  'Grid.Col = 2
  'Me.MousePointer = vbDefault
  Exit Sub
ErrorHandler:
  Grid.Redraw = True
  'Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub
