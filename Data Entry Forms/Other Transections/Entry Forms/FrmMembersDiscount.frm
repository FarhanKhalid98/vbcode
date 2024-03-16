VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmMembersDiscount 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmMembersDiscount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtProductID 
      Height          =   345
      Left            =   795
      TabIndex        =   12
      Top             =   1320
      Width           =   1155
   End
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   3825
      TabIndex        =   11
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
      MICON           =   "FrmMembersDiscount.frx":0ECA
      BC              =   12632256
      FC              =   0
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   1965
      TabIndex        =   9
      Top             =   1320
      Width           =   1755
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   5355
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1395
      Width           =   2835
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   8295
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1395
      Width           =   2835
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5850
      Left            =   810
      TabIndex        =   1
      Top             =   1755
      Width           =   10680
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
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
      Columns(1).Width=   8361
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
      Columns(2).DataField=   "Column 4"
      Columns(2).DataType=   5
      Columns(2).NumberFormat=   "########.##"
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1852
      Columns(3).Caption=   "Retail Price"
      Columns(3).Name =   "RetailPrice"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 5"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1640
      Columns(4).Caption=   "Margin %"
      Columns(4).Name =   "Margin"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 6"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1958
      Columns(5).Caption=   "Disc %"
      Columns(5).Name =   "DiscPer"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 7"
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
      Left            =   4335
      TabIndex        =   2
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
      MICON           =   "FrmMembersDiscount.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   5640
      TabIndex        =   3
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
      MICON           =   "FrmMembersDiscount.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8265
      TabIndex        =   4
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
      MICON           =   "FrmMembersDiscount.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   795
      TabIndex        =   13
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1965
      TabIndex        =   10
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5370
      TabIndex        =   8
      Top             =   1170
      Width           =   1125
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
      Left            =   1890
      TabIndex        =   6
      Top             =   180
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8310
      TabIndex        =   5
      Top             =   1170
      Width           =   900
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
Dim sSql As String

Private Sub BtnFilter_Click()
   On Error GoTo ErrorHandler
   'If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
Abc:
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  Me.MousePointer = vbHourglass
  sSql = "Select * FROM Members "
  Rs.Open sSql, CN, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  sSql = " SELECT p.ProductID, ProductName, Cost, RetailPrice, isnull(d.DiscPer,0) as DiscPer," & vbCrLf _
      + " cast(case when RetailPrice <> 0 then ((RetailPrice-Cost)/RetailPrice) * 100 else Cost end as numeric(7,3)) as ProfitMargin" & vbCrLf _
      + " FROM Products p inner join CurrentStock cs on cs.ProductID = p.ProductID" & vbCrLf _
      + " left outer join MembersDiscount d on p.productid = d.productid" & vbCrLf _
      + " where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = '" & Right("00000" + CStr(Val(TxtProductID.Text)), 5) & "' or Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & " Order by p.ProductID --ProductName"
  
  With CN.Execute(sSql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !ProductID
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("PurPrice").Value = !Cost
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("Margin").Value = !Margin
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

'Private Sub CmbCompany_Click()
'   On Error GoTo ErrorHandler
'   If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
'Abc:
'   If Rs.State = adStateOpen Then
'     Rs.CancelBatch
'     Rs.Close
'   End If
'   If CmbGroup.ListIndex <> -1 Then CmbGroup.ListIndex = -1
'   Me.MousePointer = vbHourglass
'   sSql = "Select * FROM Members "
'   Rs.Open sSql, CN, adOpenStatic, adLockBatchOptimistic
'   Grid.Redraw = False
'   Grid.CancelUpdate
'   Grid.RemoveAll
'   vSuppressUpdateEvent = True
'   sSql = "SELECT * from Products where 1=1 " & IIf(CmbCompany.ListIndex > -1, " and Companyid = " & CmbCompany.ItemData(CmbCompany.ListIndex), "") & " order by ProductName"
'  With CN.Execute(sSql)
'      Do Until .EOF
'        Grid.AddNew
'        Grid.Columns("ID").Text = !ProductID
'        Grid.Columns("Name").Text = !ProductName
'        Grid.Columns("PurPrice").Value = !Cost
'        Grid.Columns("RetailPrice").Value = !RetailPrice
'        Grid.Columns("Margin").Value = !Margin
'        Grid.Columns("DiscPer").Value = !DiscPer
'        Grid.Update
'        .MoveNext
'      Loop
'  End With
'  vSuppressUpdateEvent = False
'  Grid.Redraw = True
'  Grid.MoveFirst
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrorHandler:
'  If Err.Number = 91 Then GoTo Abc
'  Grid.Redraw = True
'  Me.MousePointer = vbDefault
'  Call ShowErrorMessage
'End Sub

'----------------------------------
'Private Sub CmbGroup_Click()
'  On Error GoTo ErrorHandler
'  If ActiveControl.Name <> CmbGroup.Name Then Exit Sub
'  If Rs.State = adStateOpen Then
'    Rs.CancelBatch
'    Rs.Close
'  End If
'  CmbCompany.ListIndex = 0
'  Me.MousePointer = vbHourglass
'  Rs.Open "Select * FROM Products where 1=1 " & IIf(CmbGroup.ListIndex > 1, " and groupid ='" & GetGroupID(CmbGroup) & "'", ""), CN, adOpenStatic, adLockBatchOptimistic
'  Grid.Redraw = False
'  Grid.CancelUpdate
'  Grid.RemoveAll
'  vSuppressUpdateEvent = True
'  'sSql = " SELECT p.*, /*isnull(PackingName,'') as PackingName, isnull(Multiplier,'') as Multiplier,*/isnull(StockLimit,0) as StockLimit from Products p" & vbCrLf _
'       + " /*left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
'       + " left outer join Packings pa on pa.PackingID = pp.PackingId*/ where 1=1 and groupid = '" & GetGroupID(CmbGroup) & "'"
'
'  sSql = " SELECT * from Products where 1=1 " & IIf(CmbGroup.ListIndex > 1, " and groupid ='" & GetGroupID(CmbGroup) & "'", "") & " Order by ProductName"
'
'  With CN.Execute(sSql)
'      Do Until .EOF
'        Grid.AddNew
'        Grid.Columns("ID").Text = !ProductID
'        Grid.Columns("Name").Text = !ProductName
'        Grid.Columns("PurPrice").Value = !Cost
'        Grid.Columns("RetailPrice").Value = !RetailPrice
'        Grid.Columns("Margin").Value = !Margin
'        Grid.Columns("DiscPer").Value = !DiscPer
'        Grid.Update
'        .MoveNext
'      Loop
'  End With
'  vSuppressUpdateEvent = False
'  Grid.Redraw = True
'  Grid.MoveFirst
'  'If Grid.Visible Then Grid.SetFocus
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrorHandler:
'  Grid.Redraw = True
'  Me.MousePointer = vbDefault
'  Call ShowErrorMessage
'End Sub

Private Sub BtnClear_Click()
'  Call CmbGroup_Click
   Call BtnFilter_Click
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
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Members Discount"
'   CmbGroup.Clear
'   CmbCompany.Clear
'   With CN.Execute("Select * FROM Groups order by GroupName")
'      CmbGroup.AddItem "All Groups"
'      Do Until .EOF
'         CmbGroup.AddItem !GroupName
'         CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
'         .MoveNext
'      Loop
'   End With
'   With CN.Execute("Select * FROM Companies order by CompanyName")
'      CmbCompany.AddItem "All Companies"
'      Do Until .EOF
'         CmbCompany.AddItem !CompanyName
'         CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
'         .MoveNext
'      Loop
'   End With
'   Grid.Columns("Name").Locked = Not ObjUserSecurity.IsAdministrator
'  If CmbCompany.ListCount > 0 Then CmbCompany.ListIndex = 0
   Call BtnFilter_Click
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
   If Grid.Visible = False Then Exit Sub
   If ActiveControl.Name <> Grid.Name Then Exit Sub
   If Val(Grid.Columns("DiscPer").Value) = 0 Then Grid.Columns("DiscPer").Value = 0
   Rs.Filter = " ProductID = '" & Grid.Columns("ID").Text & "'"
   If Rs.RecordCount = 0 And Val(Grid.Columns("DiscPer").Value) > 0 Then
      Rs.AddNew
      Rs!ProductID = Grid.Columns("ID").Text
      Rs!DiscPer = Val(Grid.Columns("DiscPer").Value)
   ElseIf Rs.RecordCount = 1 And Val(Grid.Columns("DiscPer").Value) = 0 Then
      Rs.Delete
   ElseIf Rs.RecordCount = 1 Then
      Rs!DiscPer = Val(Grid.Columns("DiscPer").Value)
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
