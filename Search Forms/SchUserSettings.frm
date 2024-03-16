VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form SchUserSettings 
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
   Begin VB.TextBox TxtStoreName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5558
      TabIndex        =   4
      Top             =   2550
      Visible         =   0   'False
      Width           =   3585
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5985
      Left            =   1763
      TabIndex        =   0
      Top             =   3150
      Width           =   11775
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
      stylesets(0).Picture=   "SchUserSettings.frx":0000
      AllowUpdate     =   0   'False
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
      BackColorEven   =   15724527
      BackColorOdd    =   16777215
      RowHeight       =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   6
      Columns(0).Width=   1879
      Columns(0).Caption=   "Store ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4313
      Columns(1).Caption=   "Store Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2275
      Columns(2).Caption=   "Organization ID"
      Columns(2).Name =   "OrganizationID"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 3"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   5398
      Columns(3).Caption=   "Organization Name"
      Columns(3).Name =   "OrganizationName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 4"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "User Name"
      Columns(4).Name =   "UserName"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 2"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "No Of Prints"
      Columns(5).Name =   "NoOfPrints"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20770
      _ExtentY        =   10557
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
      Left            =   7658
      TabIndex        =   2
      Top             =   9495
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
      MICON           =   "SchUserSettings.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   6353
      TabIndex        =   1
      Top             =   9495
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
      MICON           =   "SchUserSettings.frx":0038
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
      TabIndex        =   5
      Top             =   270
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
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
      Left            =   5558
      TabIndex        =   3
      Top             =   2325
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   13268
      Top             =   1605
      Width           =   330
   End
End
Attribute VB_Name = "SchUserSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vStrSQL As String
Public ParaOutUserNo As Byte
Dim vOrder As String, vDirection As String, vCol As Byte

Private Sub BtnClose_Click()
  Me.ParaOutUserNo = 0
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutUserNo = Rs!UserNo
  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtStoreName.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hwnd, "Search"
   LoadGrid
   Me.ParaOutUserNo = 0
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadGrid()
   vStrSQL = " Select h.UserNo as UserNo, h.storeid as StoreID, StoreName, h.OrganizationID as OrganizationID, OrganizationName, UserName, isnull(NoOfPrints,0) as NoOfPrints" & _
         " FROM UserRegistry h left outer join stores s on h.storeid  = s.storeid left outer join Organizations o on h.OrganizationID = O.OrganizationID left outer join users u on u.userno = h.userno " & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "StoreID"
   Grid.Columns("Name").DataField = "StoreName"
   Grid.Columns("OrganizationID").DataField = "OrganizationID"
   Grid.Columns("OrganizationName").DataField = "OrganizationName"
   Grid.Columns("UserName").DataField = "UserName"
   Grid.Columns("NoOfPrints").DataField = "NoOfPrints"
   
End Sub

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then Call BtnSelect_Click
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
   vOrder = " order by " & Grid.Columns(ColIndex).DataField
   If vCol = ColIndex Then
      vDirection = IIf(vDirection = " Asc", " Desc", " Asc")
   Else
      vDirection = " Asc"
   End If
   vCol = ColIndex
   LoadGrid
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case Asc("a") To Asc("z"), Asc("A") To Asc("Z"), vbKey0 To vbKey9
      TxtStoreName.Text = Chr(KeyAscii): TxtStoreName.SelStart = Len(TxtStoreName.Text): TxtStoreName.SetFocus
   End Select
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtStoreName_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtStoreName.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "StoreName like '" & Replace(TxtStoreName.Text, "'", "''") & "%'", , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtStoreName_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub
