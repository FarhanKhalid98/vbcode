VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form SchAccounts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtAdress 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6420
      MaxLength       =   30
      TabIndex        =   1
      Top             =   3345
      Width           =   5475
   End
   Begin VB.ComboBox CmbFilter 
      Height          =   315
      ItemData        =   "SchAccounts.frx":0000
      Left            =   11925
      List            =   "SchAccounts.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3345
      Width           =   2715
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3263
      MaxLength       =   30
      TabIndex        =   0
      Top             =   3345
      Width           =   3150
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Default         =   -1  'True
      Height          =   420
      Left            =   6383
      TabIndex        =   3
      Top             =   9855
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
      MICON           =   "SchAccounts.frx":0004
      BC              =   16777215
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7688
      TabIndex        =   4
      Top             =   9855
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Cancel"
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
      MICON           =   "SchAccounts.frx":0020
      BC              =   16777215
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6000
      Left            =   518
      TabIndex        =   9
      Top             =   3660
      Width           =   14325
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   8
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   16760767
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
      stylesets(0).Picture=   "SchAccounts.frx":003C
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      Columns.Count   =   8
      Columns(0).Width=   2302
      Columns(0).Caption=   "A/C No."
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5503
      Columns(1).Caption=   "Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4868
      Columns(2).Caption=   "Address"
      Columns(2).Name =   "Address"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2461
      Columns(3).Caption=   "City"
      Columns(3).Name =   "City"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2566
      Columns(4).Caption=   "Sector"
      Columns(4).Name =   "SectorName"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3016
      Columns(5).Caption=   "ContactNo"
      Columns(5).Name =   "ContactNo"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2566
      Columns(6).Caption=   "ParentAccount"
      Columns(6).Name =   "ParentAccount"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1455
      Columns(7).Caption=   "Selection"
      Columns(7).Name =   "Selection"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   11
      Columns(7).FieldLen=   256
      Columns(7).Style=   2
      Columns(7).HasForeColor=   -1  'True
      TabNavigation   =   1
      _ExtentX        =   25268
      _ExtentY        =   10583
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   6420
      TabIndex        =   8
      Top             =   3075
      Width           =   690
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Left            =   3308
      TabIndex        =   7
      Top             =   3075
      Width           =   495
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parent Accounts"
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
      Left            =   11925
      TabIndex        =   6
      Top             =   3075
      Width           =   1425
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
      TabIndex        =   5
      Top             =   270
      Width           =   1245
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   13208
      Top             =   1635
      Width           =   330
   End
End
Attribute VB_Name = "SchAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Public ParaOutAccountNo As String, vSelection As String
Public ParaOutAccountName As String
Public ParaInDetail As String
Public ParaInWhereClause As String
Public ParaInAllowListSelection As Boolean 'Whether allow to filter from combo or not.
Dim vOrder As String, vDirection As String, vCol As Byte, vSQL As String
Dim vSuppressUpdateEvent As String
Dim vWhere As String


Private Sub cmbfilter_click()
  On Error GoTo ErrorHandler
  If CmbFilter.Visible = False Then Exit Sub
  If ActiveControl.Name <> CmbFilter.Name Then Exit Sub
'    LoadData1
     LoadData2
  
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub
Private Sub LoadData1()
  Set Rs = cn.Execute(vSQL)
  Set Grid.DataSource = Rs
  Grid.Columns("ID").DataField = "AccountNo"
  Grid.Columns("Name").DataField = "AccountName"
  Grid.Columns("Address").DataField = "Address"
  Grid.Columns("City").DataField = "City"
  Grid.Columns("ParentAccount").DataField = "ParentAccount"
  Grid.Columns("Selection").DataField = "Selection"
End Sub

Private Sub LoadData2()
 On Error GoTo ErrorHandler
  If CmbFilter.ListIndex > 0 Then
    If CmbFilter.Text = "Customers" Then
    vWhere = " And (c.AccountNo like '" & CmbFilter.ItemData(CmbFilter.ListIndex) & "%' or c.accountno like '64%')"
    Else
    vWhere = " And c.AccountNo like '" & CmbFilter.ItemData(CmbFilter.ListIndex) & "%'"
    End If
   Else
   vWhere = ""
  End If
  vWhere = vWhere & IIf(Me.ParaInDetail = "False", " and c.IsDetailed=0", " and c.IsDetailed=1")
  vSQL = " Select c.AccountNo as AccountNo, c.AccountName as AccountName, isnull(p.Address,'') Address, Isnull(p.City,'') City, Isnull(SectorName, '') SectorName, isnull(ca.AccountName,'') as ParentAccount,  isnull(mobile,'') + ' ' + isnull(mobile2,'') + isnull(Phone1,'') + ' ' + isnull(Phone2,'')  ContactNo, 0 As Selection" & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " left outer join Sectors s on p.Sectorid = s.Sectorid  " & vbCrLf _
         + " left outer join ChartofAccounts ca on c.ParentAccountNo = ca.AccountNo  " & vbCrLf _
         + " where 1=1 and (isLockParty = 0 or isLockParty is null) " & Me.ParaInWhereClause & vWhere & IIf(TxtFilter.Text = "", "", " and (c.AccountName + isnull(p.Address,'') +  isnull(p.City,'') + isnull(p.Phone1,'') + isnull(p.Phone2,'') + isnull(p.mobile,'') + isnull(p.mobile2,'')  ) like '%" & Replace(TxtFilter.Text, "'", "''") & "%'") & IIf(TxtAdress.Text = "", "", " and address like '" & TxtAdress.Text & "%'") & vOrder
    
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vSQL, cn, adOpenStatic, adLockReadOnly
   'Set Grid.DataSource = Rs
   'Grid.Columns("ID").DataField = "SectorId"
   'Grid.Columns("Name").DataField = "SectorName"
   Grid.Redraw = False
   vSuppressUpdateEvent = True
   Grid.CancelUpdate
   Grid.RemoveAll
   Do Until Rs.EOF
     Grid.AddNew
     Grid.Columns("ID").Text = Rs!AccountNo
     Grid.Columns("Name").Text = Rs!AccountName
     Grid.Columns("Address").Text = Rs!Address
     Grid.Columns("ContactNo").Text = Rs!ContactNo
     Grid.Columns("City").Text = Rs!City
     Grid.Columns("SectorName").Text = Rs!SectorName
     Grid.Columns("ParentAccount").Text = Rs!ParentAccount
     Grid.Columns("Selection").Value = 0
     Grid.Update
     Rs.MoveNext
   Loop
   vSuppressUpdateEvent = False
   Grid.Redraw = True
   Grid.MoveFirst
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Me.ParaOutAccountNo = ""
  Me.ParaOutAccountName = ""
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtFilter.Name, CmbFilter.Name
         Call BtnSelect_Click
      End Select
   ElseIf KeyCode = vbKeyDown Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtFilter.Name, TxtAdress.Name, CmbFilter.Name
         Grid.SetFocus
      End Select
   End If
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   If Grid.Columns("Selection").Value = True Then
      vSelection = vSelection & Grid.Columns("ID").Text & ","
   Else
      vSelection = Replace(vSelection, "'" & Grid.Columns("ID").Text & "',", "")
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
  If Grid.rows > 0 Then BtnSelect_Click
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
   If Grid.rows = 0 Then Exit Sub
   Call Grid_BeforeUpdate(0)
   If vSelection <> "" Then
      vSelection = Left(vSelection, Len(vSelection) - 1)
   Else
      vSelection = Grid.Columns("ID").Text
   End If
   Me.ParaOutAccountNo = vSelection
   Me.ParaOutAccountName = Grid.Columns("Name").Text
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   Set Rs = New ADODB.Recordset
   CmbFilter.AddItem "-- ALL PARENT ACCOUNTS --", 0
   With cn.Execute("Select AccountNo, AccountName from ChartofAccounts Where isDetailed = 0 order by AccountName")
      Do Until .EOF
         CmbFilter.AddItem !AccountName
         CmbFilter.ItemData(CmbFilter.NewIndex) = !AccountNo
         .MoveNext
      Loop
   End With
   vOrder = " Order By AccountName"
   If CmbFilter.ListCount > 0 Then CmbFilter.ListIndex = 0
   CmbFilter.Enabled = Me.ParaInAllowListSelection
   vSelection = ""
   LoadData2
   Me.ParaOutAccountNo = ""
   Me.ParaOutAccountName = ""
   Me.ParaInDetail = IIf(Me.ParaInDetail = "", "True", "False")
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
    Select Case Grid.Columns(ColIndex).DataField
     Case "Column 0"
        vOrder = " order by C.AccountNo"
    Case "Column 1"
        vOrder = " order by C.AccountName"
    Case "Column 2"
        vOrder = " order by Address"
    Case "Column 3"
        vOrder = " order by City"
    Case "Column 4"
        vOrder = " order by ca.AccountName"
        
End Select
    
    
'   vOrder = " order by " & Grid.Columns(ColIndex).DataField
   If vCol = ColIndex Then
      vDirection = IIf(vDirection = " Asc", " Desc", " Asc")
   Else
      vDirection = " Asc"
   End If
   vCol = ColIndex
'   LoadGrid
    LoadData2
'   Call cmbfilter_click
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case Asc("a") To Asc("z"), Asc("A") To Asc("Z"), vbKey0 To vbKey9
      TxtFilter.Text = Chr(KeyAscii): TxtFilter.SelStart = Len(TxtFilter.Text): TxtFilter.SetFocus
   End Select
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtAdress_Change()
 On Error GoTo ErrorHandler
  LoadData2
'  cmbfilter_click
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub TxtFilter_Change()
  On Error GoTo ErrorHandler
'  If Trim(TxtFilter.Text) = "" Then Grid.MoveFirst: Exit Sub
'  Rs.Find "AccountName like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
'  If Rs.EOF Then Grid.MoveLast
LoadData2
  cmbfilter_click
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub
