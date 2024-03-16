VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmAllocateProduct 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmAllocateProduct.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3645
      TabIndex        =   4
      Top             =   6720
      Width           =   195
   End
   Begin VB.CheckBox ChkAll 
      Height          =   225
      Left            =   11385
      TabIndex        =   18
      Top             =   1710
      Width           =   195
   End
   Begin VB.ComboBox CmbSubGroup 
      Height          =   315
      Left            =   7605
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1215
      Width           =   1755
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   4050
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1215
      Width           =   1710
   End
   Begin VB.TextBox TxtProductID 
      Height          =   345
      Left            =   210
      TabIndex        =   14
      Top             =   1230
      Width           =   975
   End
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   3015
      TabIndex        =   13
      Top             =   1215
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
      MICON           =   "FrmAllocateProduct.frx":0ECA
      BC              =   12632256
      FC              =   0
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   1200
      TabIndex        =   11
      Top             =   1230
      Width           =   1755
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   5790
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1215
      Width           =   1755
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4860
      Left            =   195
      TabIndex        =   3
      Top             =   1665
      Width           =   11625
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   7
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
      stylesets(0).Picture=   "FrmAllocateProduct.frx":0EE6
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
      stylesets(1).Picture=   "FrmAllocateProduct.frx":0F02
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
      Columns.Count   =   7
      Columns(0).Width=   1667
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   5265
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2646
      Columns(2).Caption=   "Company"
      Columns(2).Name =   "Company"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).NumberFormat=   "########.##"
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   2646
      Columns(3).Caption=   "Group"
      Columns(3).Name =   "Group"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   2646
      Columns(4).Caption=   "Sub Group"
      Columns(4).Name =   "SubGroup"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   3334
      Columns(5).Caption=   "Customer Name"
      Columns(5).Name =   "CustomerName"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1296
      Columns(6).Caption=   "Select"
      Columns(6).Name =   "Select"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   0
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   11
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      TabNavigation   =   1
      _ExtentX        =   20505
      _ExtentY        =   8572
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
      Left            =   4043
      TabIndex        =   6
      Top             =   7620
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
      MICON           =   "FrmAllocateProduct.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   5363
      TabIndex        =   7
      Top             =   7620
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
      MICON           =   "FrmAllocateProduct.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6683
      TabIndex        =   8
      Top             =   7620
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
      MICON           =   "FrmAllocateProduct.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnCustomer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5580
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   6720
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "..."
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
      MICON           =   "FrmAllocateProduct.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   315
      Left            =   4980
      TabIndex        =   5
      Top             =   6720
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   5940
      TabIndex        =   20
      Tag             =   "nc"
      Top             =   6720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
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
      Left            =   3915
      TabIndex        =   21
      Top             =   6750
      Width           =   795
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
      Left            =   7605
      TabIndex        =   17
      Top             =   990
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
      Left            =   4050
      TabIndex        =   16
      Top             =   990
      Width           =   1320
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
      Left            =   210
      TabIndex        =   15
      Top             =   990
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
      Left            =   1200
      TabIndex        =   12
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allocate Product"
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
      TabIndex        =   10
      Top             =   180
      Width           =   2160
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
      Left            =   5790
      TabIndex        =   9
      Top             =   990
      Width           =   1065
   End
End
Attribute VB_Name = "FrmAllocateProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim ssql As String, i As Integer
Public ParaInZoneID As String

Private Sub CmbCompany_Click()
   On Error GoTo ErrorHandler
   If CmbCompany.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
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

Private Sub CmbSubGroup_Click()
   On Error GoTo ErrorHandler
   If CmbSubGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbSubGroup.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkAll_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
      Grid.Columns("Select").Value = ChkAll.Value
      Grid.MoveNext
   Next i
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub BtnFilter_Click()
  On Error GoTo ErrorHandler
  Me.MousePointer = vbHourglass
   Dim vWords() As String
   Dim vProductName As String
   vWords = Split(TxtProductName.Text, " ")
   vProductName = ""
   For i = 0 To UBound(vWords)
       vProductName = vProductName & " and Productname like '%" & Replace(vWords(i), "'", "''") & "%'"
   Next
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
     
     ssql = " SELECT z.ProductID, ProductName, GroupName, CompanyName, SubGroupName, PartyName" & vbCrLf _
      + " FROM (select ProductID, PartyID from ZoneAllotment where zoneid = " & ParaInZoneID & ")z inner join Products p on p.ProductID = z.ProductID" & vbCrLf _
      + " Left Outer join Parties pr on z.PartyID = pr.PartyID" & vbCrLf _
      + " Left Outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
      + " Left Outer join SubGroups s on s.SubGroupID = p.SubGroupID " & vbCrLf _
      + " Left Outer join Companies c on c.CompanyID = p.CompanyID " & vbCrLf _
      + " where p.isLocked = 0 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = '" & Right("00000" + CStr(Val(TxtProductID.Text)), 5) & "'") & vProductName & " Order by p.ProductID --ProductName"

   
  With CN.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("Group").Value = !GroupName
        Grid.Columns("SubGroup").Value = !SubGroupName
        Grid.Columns("Company").Value = !CompanyName
        Grid.Columns("CustomerName").Value = !PartyName
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

Private Sub BtnClear_Click()
   ChkCustomer.Value = 0
   Call PopulateGrid
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   Dim vSQL As String
   Grid.Update
   If ChkCustomer.Value = 1 Then
      ssql = "Update ZoneAllotment set "
      If ChkCustomer.Value = 1 Then
         vSQL = vSQL & " PartyID = " & IIf(Trim(TxtCustomerID.Text) = "", "Null", "'" & TxtCustomerID.Text & "'") & IIf(vSQL = "", "", ",")
      End If

      ssql = ssql & vSQL
      Grid.MoveFirst
      Grid.Redraw = False
      For i = 0 To Grid.Rows - 1
         If Grid.Columns("Select").Value = True Then
            CN.Execute ssql & " Where ProductID = '" & Grid.Columns("ID").Value & "' and ZoneID = " & ParaInZoneID
         End If
         Grid.MoveNext
      Next i
      Grid.Redraw = True
   
      MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
   Else
      MsgBox "Your must select at least one Product.", vbOKOnly + vbInformation, "Information"
   End If
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub ChkCustomer_Click()
   TxtCustomerID.Enabled = ChkCustomer.Value
   BtnCustomer.Enabled = ChkCustomer.Value
   TxtCustomerName.Enabled = ChkCustomer.Value
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
   Me.MousePointer = vbHourglass
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
  
   ssql = " SELECT z.ProductID, ProductName, GroupName, CompanyName, SubGroupName, PartyName" & vbCrLf _
      + " FROM (select ProductID, PartyID from ZoneAllotment where zoneid = " & ParaInZoneID & ")z inner join Products p on p.ProductID = z.ProductID" & vbCrLf _
      + " Left Outer join Parties pr on z.PartyID = pr.PartyID" & vbCrLf _
      + " Left Outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
      + " Left Outer join SubGroups s on s.SubGroupID = p.SubGroupID " & vbCrLf _
      + " Left Outer join Companies c on c.CompanyID = p.CompanyID " & vbCrLf _
      + " where p.isLocked = 0 " & IIf(CmbCompany.ListIndex = 0, "", " and p.CompanyID = " & CmbCompany.ItemData(CmbCompany.ListIndex)) & vbCrLf _
      + IIf(CmbGroup.ListIndex = 0, "", " and p.GroupID ='" & GetGroupID(CmbGroup) & "'") & vbCrLf _
      + IIf(CmbSubGroup.ListIndex = 0, "", " and p.SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & vbCrLf _
      + " Order by ProductName"
  
   With CN.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("Group").Value = !GroupName
        Grid.Columns("SubGroup").Value = !SubGroupName
        Grid.Columns("Company").Value = !CompanyName
        Grid.Columns("CustomerName").Value = !PartyName
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

Private Sub UpdateProduct()
   On Error GoTo ErrorHandler
   Dim vSQL As String
   vSQL = " insert into ZoneAllotment(ProductID, ZoneID, PartyID)" & vbCrLf _
      + " Select p.ProductID, " & ParaInZoneID & ", Null" & vbCrLf _
      + " from (select ProductID from ZoneAllotment where ZoneID = " & ParaInZoneID & ")z right outer join Products p on p.ProductID = z.ProductID " & vbCrLf _
      + " where z.ProductID is null"
   CN.Execute vSQL
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Allocate Product"
   
   UpdateProduct
   
   ChkAll.ZOrder 0
   'ChkCompany_Click
   
   CmbCompany.Clear
   With CN.Execute("Select * FROM Companies Order By CompanyName")
      CmbCompany.AddItem "All Companies"
      CmbCompany.ItemData(CmbCompany.NewIndex) = 0
      Do Until .EOF
         CmbCompany.AddItem !CompanyName
         CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
         .MoveNext
      Loop
   End With
   
   CmbGroup.Clear
   With CN.Execute("Select * FROM Groups Order By GroupName")
      CmbGroup.AddItem "All Groups"
      CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left("000", 1)) & Asc(Mid("000", 2, 1)) & Asc(Mid("000", 3, 1))
      Do Until .EOF
         CmbGroup.AddItem !GroupName
         CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
   End With
   
   CmbSubGroup.Clear
   With CN.Execute("Select * FROM SubGroups Order By SubGroupName")
      CmbSubGroup.AddItem "All SubGroups"
      CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = 0
      Do Until .EOF
         CmbSubGroup.AddItem !SubGroupName
         CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = !SubGroupID
         .MoveNext
      Loop
   End With

   If CmbCompany.ListCount > 0 Then CmbCompany.ListIndex = 1 Else CmbCompany.ListIndex = 0
   CmbGroup.ListIndex = 0
   CmbCompany.ListIndex = 0
   CmbSubGroup.ListIndex = 0
   PopulateGrid
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
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCustomerID.Name: If FunSelectCustomer(ssFunctionKey, True) = True Then If BtnSave.Enabled Then BtnSave.SetFocus
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

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCustomer.Show vbModal, Me
        If SchCustomer.ParaOutCustomerID = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchCustomer.ParaOutCustomerID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Parties where PartyID=" & Val(TxtCustomerID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerName.Text = !PartyName
          FunSelectCustomer = True
          .Close
          Exit Function
      Else
          FunSelectCustomer = False
          .Close
          TxtCustomerID.Text = ""
          TxtCustomerName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
     If BtnSave.Enabled Then BtnSave.SetFocus
    Else
     If TxtCustomerID.Enabled Then TxtCustomerID.SetFocus
   End If
End Sub

Private Sub TxtCustomerID_Change()
   If TxtCustomerID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   If TxtCustomerName.Text <> "" Then TxtCustomerName.Text = ""
End Sub

Private Sub TxtCustomerID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCustomer(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCustomer(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

