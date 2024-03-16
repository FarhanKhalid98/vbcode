VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JWLDBUTN.OCX"
Begin VB.Form SchParty 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SchParty.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPartyName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1590
      TabIndex        =   2
      Top             =   1710
      Width           =   2550
   End
   Begin VB.TextBox TxtPartyID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   390
      TabIndex        =   1
      Top             =   1710
      Width           =   1185
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6000
      Left            =   390
      TabIndex        =   0
      Top             =   2070
      Width           =   11055
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   -2147483634
      stylesets(0).BackColor=   -2147483635
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
      stylesets(0).Picture=   "SchParty.frx":E7FA
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
      RowHeight       =   609
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   4
      Columns(0).Width=   2064
      Columns(0).Caption=   "Party ID"
      Columns(0).Name =   "ID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4524
      Columns(1).Caption=   "Party Name"
      Columns(1).Name =   "Name"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   8573
      Columns(2).Caption=   "Address"
      Columns(2).Name =   "Address"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3784
      Columns(3).Caption=   "City"
      Columns(3).Name =   "City"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   19500
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
      Left            =   5520
      TabIndex        =   4
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
      MICON           =   "SchParty.frx":E816
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4215
      TabIndex        =   3
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
      MICON           =   "SchParty.frx":E832
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   1485
      Width           =   825
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Party ID"
      Height          =   195
      Left            =   390
      TabIndex        =   5
      Top             =   1485
      Width           =   570
   End
End
Attribute VB_Name = "SchParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSql As String
Public ParaOutPartyID As String
Dim vOrder As String, vDirection As String, vCol As Byte

Private Sub BtnClose_Click()
  Me.ParaOutPartyID = ""
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutPartyID = Rs!PartyID
  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtPartyID.Name, TxtPartyName.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
'    VStrSQL = "SELECT * FROM Parties WHERE PartyType='V'"
'    Rs.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
'    Set Grid.DataSource = Rs
'    Grid.Columns("ID").DataField = "PartyID"
'    Grid.Columns("Name").DataField = "PartyName"
'    Grid.Columns("Address").DataField = "Address"
'    Grid.Columns("City").DataField = "City"
    Me.ParaOutPartyID = ""
    LoadData
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
      vOrder = " order by PartyId"
   Case 1
      vOrder = " order by PartyName"
   Case 2
      vOrder = " order by Address"
   Case 3
      vOrder = " order by city"
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
      TxtPartyName.Text = Chr(KeyAscii): TxtPartyName.SelStart = Len(TxtPartyName.Text): TxtPartyName.SetFocus
   Case vbKey0 To vbKey9
      TxtPartyID.Text = Chr(KeyAscii): TxtPartyID.SelStart = Len(TxtPartyID.Text): TxtPartyID.SetFocus
   End Select
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtPartyID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtPartyID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "PartyID like '" & TxtPartyID.Text & "%'", , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtPartyName_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtPartyName.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "PartyName like '" & TxtPartyName.Text & "%'", , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
   
End Sub

Private Sub LoadData()
   vStrSql = "SELECT * FROM Parties WHERE PartyType='V'" & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSql, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "PartyID"
   Grid.Columns("Name").DataField = "PartyName"
   Grid.Columns("Address").DataField = "Address"
   Grid.Columns("City").DataField = "City"
End Sub
