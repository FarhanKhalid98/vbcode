VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form SchCompany 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCompanyName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5239
      TabIndex        =   2
      Top             =   3323
      Width           =   5325
   End
   Begin VB.TextBox TxtCompanyID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2742
      TabIndex        =   1
      Top             =   3323
      Width           =   2490
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6000
      Left            =   2745
      TabIndex        =   0
      Top             =   3660
      Width           =   9915
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   3
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   12566463
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
      stylesets(0).Picture=   "SchCompany.frx":0000
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
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
      Columns.Count   =   3
      Columns(0).Width=   4392
      Columns(0).Caption=   "Company ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   9419
      Columns(1).Caption=   "Company Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Selection"
      Columns(2).Name =   "Selection"
      Columns(2).Alignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   11
      Columns(2).FieldLen=   256
      Columns(2).Style=   2
      TabNavigation   =   1
      _ExtentX        =   17489
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
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6769
      TabIndex        =   4
      Top             =   9863
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
      MICON           =   "SchCompany.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   5464
      TabIndex        =   3
      Top             =   9863
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
      MICON           =   "SchCompany.frx":0038
      BC              =   14737632
      FC              =   0
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
      TabIndex        =   7
      Top             =   270
      Width           =   1245
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   12379
      Top             =   1658
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Height          =   195
      Left            =   5239
      TabIndex        =   6
      Top             =   3098
      Width           =   1125
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Company ID"
      Height          =   195
      Left            =   2742
      TabIndex        =   5
      Top             =   3098
      Width           =   870
   End
End
Attribute VB_Name = "SchCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSQL As String
Public ParaOutCompanyID As String, vSelection As String
Dim vOrder As String, vDirection As String, vCol As Byte

Private Sub BtnClose_Click()
  Me.ParaOutCompanyID = ""
  Unload Me
End Sub

Private Sub BtnSelect_Click()
   On Error GoTo ErrorHandler
   If Grid.Rows = 0 Then Exit Sub
   'Call Grid_BeforeUpdate(0)
   If vSelection <> "" Then
      vSelection = Left(vSelection, Len(vSelection) - 1)
   Else
      vSelection = Grid.Columns("ID").Text
   End If
   Me.ParaOutCompanyID = vSelection
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtCompanyID.Name, TxtCompanyName.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    ShowPicture Me, 2
    AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
    SetWindowText Me.hWnd, "Search"
    Set Rs = New ADODB.Recordset
'    VStrSQL = "SELECT * FROM Parties WHERE PartyType='V'"
'    Rs.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
'    Set Grid.DataSource = Rs
'    Grid.Columns("ID").DataField = "CompanyID"
'    Grid.Columns("Name").DataField = "CompanyName"
'    Grid.Columns("Address").DataField = "Address"
'    Grid.Columns("City").DataField = "City"
    vSelection = ""
    Me.ParaOutCompanyID = ""
    vOrder = " Order By CompanyName"
    LoadData
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

Private Sub Grid_Change()
   On Error GoTo ErrorHandler
'   If vSuppressUpdateEvent Then Exit Sub
   If Grid.Columns("Selection").Value = True Then
      vSelection = vSelection & Grid.Columns("ID").Text & ","
   Else
      vSelection = Replace(vSelection, Grid.Columns("ID").Text & ",", "")
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then Call BtnSelect_Click
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
Select Case ColIndex
   Case 0
      vOrder = " order by CompanyID"
   Case 1
      vOrder = " order by CompanyName"
   End Select
   If vCol = ColIndex Then
      vDirection = IIf(vDirection = " Asc", " Desc", " Asc")
   Else
      vDirection = " Asc"
   End If
   vCol = ColIndex
   LoadData
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case Asc("a") To Asc("z"), Asc("A") To Asc("Z")
      TxtCompanyName.Text = Chr(KeyAscii): TxtCompanyName.SelStart = Len(TxtCompanyName.Text): TxtCompanyName.SetFocus
   Case vbKey0 To vbKey9
      TxtCompanyID.Text = Chr(KeyAscii): TxtCompanyID.SelStart = Len(TxtCompanyID.Text): TxtCompanyID.SetFocus
   End Select
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtCompanyID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtCompanyID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "CompanyID =" & TxtCompanyID.Text, , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCompanyID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtCompanyName_Change()
   On Error GoTo ErrorHandler
'   If Trim(TxtCompanyName.Text) = "" Then Grid.MoveFirst: Exit Sub
'   Rs.Find "CompanyName like '" & TxtCompanyName.Text & "%'", , adSearchForward, 1
'   If Rs.EOF Then Grid.MoveLast
   LoadData
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadData()
   On Error GoTo ErrorHandler
   vStrSQL = "SELECT * FROM Companies where 1=1 " & IIf(Trim(TxtCompanyName.Text) = "", "", " and CompanyName like '%" & TxtCompanyName.Text & "%'") & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
   'Set Grid.DataSource = Rs
   'Grid.Columns("ID").DataField = "SectorId"
   'Grid.Columns("Name").DataField = "SectorName"
   Grid.Redraw = False
'   vSuppressUpdateEvent = True
   Grid.CancelUpdate
   Grid.RemoveAll
   Do Until Rs.EOF
     Grid.AddNew
     Grid.Columns("ID").Text = Rs!companyid
     Grid.Columns("Name").Text = Rs!CompanyName
     Grid.Columns("Selection").Value = 0
     Grid.Update
     Rs.MoveNext
   Loop
'   vSuppressUpdateEvent = False
   Grid.Redraw = True
   Grid.MoveFirst
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCompanyName_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub
