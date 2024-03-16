VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SchProduct 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SchProduct1.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   525
      Left            =   8490
      Top             =   975
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   926
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SchProduct1.frx":6971
      Height          =   4230
      Left            =   1575
      TabIndex        =   6
      Top             =   2010
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   7461
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   3765
      TabIndex        =   0
      Top             =   1380
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   556
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   1935
      TabIndex        =   1
      Top             =   1380
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   556
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6015
      TabIndex        =   3
      Top             =   8205
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
      MICON           =   "SchProduct1.frx":6986
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4710
      TabIndex        =   2
      Top             =   8205
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
      MICON           =   "SchProduct1.frx":69A2
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      Height          =   195
      Left            =   1950
      TabIndex        =   5
      Top             =   1170
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3750
      TabIndex        =   4
      Top             =   1140
      Width           =   1020
   End
End
Attribute VB_Name = "SchProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSql As String
Public ParaOutID As String
Dim vOrder As String, vDirection As String, vCol As Byte

Private Sub BtnClose_Click()
  Me.ParaOutID = ""
  Unload Me
  'Me.Hide
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutID = Rs!Productid
  Unload Me
  'Me.Hide
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_Activate()
'   Desktop.Caption = "Search"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case DataGrid1.Name, TxtProductID.Name, TxtProductName.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    'Set Rs = New ADODB.Recordset
    'vStrSql = "SELECT p.productid,p.code,p.productname,p.retailprice,p.purprice,v.qty FROM Products p inner join vuCurrentStock v on v.productid=p.productid" 'where p.productname like '%" & TxtProductName.Text & "%'" & vOrder & vDirection
    'If Rs.State = adStateOpen Then Rs.Close
    'Rs.Open vStrSql, CN, adOpenStatic, adLockReadOnly
    'LoadData
    Set DataGrid1.DataSource = Nothing
    Adodc1.RecordSource = ""
    Adodc1.ConnectionString = ""
    Adodc1.ConnectionString = CN.ConnectionString
    Adodc1.RecordSource = "SELECT * FROM Products"
    Adodc1.Refresh
    'txtSqlStatement = Adodc1.RecordSource
    If Adodc1.Recordset.Fields.Count = 0 Then
        DataGrid1.ClearFields
    Else
        Set DataGrid1.DataSource = Adodc1.Recordset
        DataGrid1.ClearFields
        DataGrid1.ReBind
    End If
    'Adodc1.ConnectionString = CN.ConnectionString
    'Adodc1.RecordSource = vStrSql
    'Adodc1.Refresh
    'Set DataGrid1.DataSource = Adodc1.Recordset
    'DataGrid1.ReBind
    'vOrder = " order by p.productname "
    'Me.ParaOutID = ""
    Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then Call BtnSelect_Click
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
      vOrder = " order by " & Grid.Columns(ColIndex).Name
      If vCol = ColIndex Then
         vDirection = IIf(vDirection = " Asc", " Desc", " Asc")
      Else
         vDirection = " Asc"
      End If
   vCol = ColIndex
   'Rs.Sort = Grid.Columns(ColIndex).Name
   LoadData
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case Asc("a") To Asc("z"), Asc("A") To Asc("Z")
      TxtProductName.Text = Chr(KeyAscii): TxtProductName.SetFocus
   Case vbKey0 To vbKey9
      TxtProductID.Text = Chr(KeyAscii): TxtProductID.SetFocus
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim obj
   For Each obj In Me
      Set obj = Nothing
   Next
End Sub

Private Sub Image1_Click()
   Unload Me
   'Me.Hide
End Sub


Private Sub TxtProductID_Change()
   On Error GoTo ErrorHandler
   'If Trim(TxtProductID.Text) = "" Then 'Grid.MoveFirst: Exit Sub
   'Rs.Find "ProductID like '" & TxtProductID.Text & "%'", , adSearchForward, 1
   'If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtProductID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then DataGrid1.SetFocus
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

Private Sub TxtProductName_Change()
   On Error GoTo ErrorHandler
    vStrSql = "select * from products where productname like '%" & TxtProductName.Text & "%'"
  ' If Trim(TxtProductName.Text) = "" Then Grid.MoveFirst: Exit Sub
  ' Rs.Find "ProductName like '" & TxtProductName.Text & "%'", , adSearchForward, 1
  ' If Rs.EOF Then Grid.MoveLast
  LoadData
  Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadData()
   'Grid.DataMode
   'Rs.Filter = "productname like '%" & TxtProductName.Text & "%'"
  vStrSql = "SELECT p.productid,p.code,p.productname,p.retailprice,p.purprice,v.qty FROM Products p inner join vuCurrentStock v on v.productid=p.productid where p.productname like '%" & TxtProductName.Text & "%'"
  'Adodc1.RecordSource = vStrSql
  'Adodc1.Refresh
  'Set DataGrid1.DataSource = Adodc1.Recordset
  'Adodc1.Recordset.Filter = "productname like '" & TxtProductName.Text & "'"
  'DataGrid1.ClearFields
  'DataGrid1.ReBind
   Me.Refresh
   Adodc1.RecordSource = vStrSql
   Adodc1.Refresh
   Set DataGrid1.DataSource = Adodc1.Recordset
End Sub

Private Sub TxtProductName_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then DataGrid1.SetFocus
End Sub
