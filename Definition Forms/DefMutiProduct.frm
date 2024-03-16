VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form DefMultiProduct 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6015
      TabIndex        =   1
      Top             =   8250
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
      MICON           =   "DefMutiProduct.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4710
      TabIndex        =   0
      Top             =   8250
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
      MICON           =   "DefMutiProduct.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multi Product"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   180
      Width           =   2310
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "DefMultiProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim VStrSQL As String
Public ParaOutBrandID As String
Dim vOrder As String, vDirection As String, vCol As Byte

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
'  If Grid.Rows = 0 Then Exit Sub
'  Me.ParaOutBrandID = Rs!BrandID
'  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    ShowPicture Me, 2
    AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
    SetWindowText Me.hWnd, "Multi "
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
   VStrSQL = "SELECT * FROM Brands where 1=1 " & IIf(Trim(TxtBrandName.Text) = "", "", " and BrandName like '%" & TxtBrandName.Text & "%'") & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
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
