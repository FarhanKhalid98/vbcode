VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form SchBrand 
   BorderStyle     =   0  'None
   ClientHeight    =   11940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15450
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   796
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtBrandName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5239
      TabIndex        =   2
      Top             =   3323
      Width           =   5325
   End
   Begin VB.TextBox TxtBrandID 
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
      Left            =   2749
      TabIndex        =   0
      Top             =   3653
      Width           =   8025
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
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
      stylesets(0).Picture=   "SchBrand.frx":0000
      AllowUpdate     =   0   'False
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
      Columns.Count   =   2
      Columns(0).Width=   4392
      Columns(0).Caption=   "Brand ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   9419
      Columns(1).Caption=   "Brand Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   14155
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
      MICON           =   "SchBrand.frx":001C
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
      MICON           =   "SchBrand.frx":0038
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
      Caption         =   "Brand Name"
      Height          =   195
      Left            =   5239
      TabIndex        =   6
      Top             =   3098
      Width           =   885
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Brand ID"
      Height          =   195
      Left            =   2749
      TabIndex        =   5
      Top             =   3098
      Width           =   630
   End
End
Attribute VB_Name = "SchBrand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSQL As String
Public ParaOutBrandID As String
Dim vOrder As String, vDirection As String, vCol As Byte

Private Sub BtnClose_Click()
  Me.ParaOutBrandID = ""
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutBrandID = Rs!BrandID
  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtBrandID.Name, TxtBrandName.Name
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
'    Grid.Columns("ID").DataField = "BrandID"
'    Grid.Columns("Name").DataField = "BrandName"
'    Grid.Columns("Address").DataField = "Address"
'    Grid.Columns("City").DataField = "City"
    Me.ParaOutBrandID = ""
    vOrder = " Order By BrandName"
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

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then Call BtnSelect_Click
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
Select Case ColIndex
   Case 0
      vOrder = " order by BrandID"
   Case 1
      vOrder = " order by BrandName"
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
      TxtBrandName.Text = Chr(KeyAscii): TxtBrandName.SelStart = Len(TxtBrandName.Text): TxtBrandName.SetFocus
   Case vbKey0 To vbKey9
      TxtBrandID.Text = Chr(KeyAscii): TxtBrandID.SelStart = Len(TxtBrandID.Text): TxtBrandID.SetFocus
   End Select
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtBrandID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtBrandID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "BrandID = " & TxtBrandID.Text, , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtBrandID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtBrandName_Change()
   On Error GoTo ErrorHandler
'   If Trim(TxtBrandName.Text) = "" Then Grid.MoveFirst: Exit Sub
'   Rs.Find "BrandName like '" & TxtBrandName.Text & "%'", , adSearchForward, 1
'   If Rs.EOF Then Grid.MoveLast
   LoadData
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadData()
   On Error GoTo ErrorHandler
   vStrSQL = "SELECT * FROM Brands where 1=1 " & IIf(Trim(TxtBrandName.Text) = "", "", " and BrandName like '%" & TxtBrandName.Text & "%'") & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "BrandID"
   Grid.Columns("Name").DataField = "BrandName"
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtBrandName_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub
