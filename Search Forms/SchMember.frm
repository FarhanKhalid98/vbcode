VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form SchMember 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtBarcode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5370
      TabIndex        =   6
      Top             =   2948
      Width           =   1785
   End
   Begin VB.TextBox TxtMemberName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7155
      TabIndex        =   7
      Top             =   2948
      Width           =   4170
   End
   Begin VB.TextBox TxtMemberID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4118
      TabIndex        =   5
      Top             =   2948
      Width           =   1245
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5985
      Left            =   4118
      TabIndex        =   0
      Top             =   3278
      Width           =   7185
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
      stylesets(0).Picture=   "SchMember.frx":0000
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
      Columns.Count   =   3
      Columns(0).Width=   2196
      Columns(0).Caption=   "Member ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3149
      Columns(1).Caption=   "Barcode"
      Columns(1).Name =   "Barcode"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 2"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   6800
      Columns(2).Caption=   "Member Name"
      Columns(2).Name =   "Name"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 1"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   12674
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
      Left            =   7718
      TabIndex        =   2
      Top             =   9488
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
      MICON           =   "SchMember.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   6413
      TabIndex        =   1
      Top             =   9488
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
      MICON           =   "SchMember.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode"
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
      Left            =   5400
      TabIndex        =   9
      Top             =   2745
      Width           =   720
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
      TabIndex        =   8
      Top             =   270
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
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
      Left            =   7170
      TabIndex        =   4
      Top             =   2730
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
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
      Left            =   4118
      TabIndex        =   3
      Top             =   2723
      Width           =   930
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   10613
      Top             =   2003
      Width           =   330
   End
End
Attribute VB_Name = "SchMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim VStrSQL As String
Dim vMemberName As String, vBarcode As String
Public ParaOutMemberID As String
Dim vOrder As String, vDirection As String, vCol As Byte

Private Sub BtnClose_Click()
  Me.ParaOutMemberID = ""
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutMemberID = Rs!MemberID
  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtMemberName.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"
   vMemberName = ""
   vBarcode = ""
   LoadGrid
   Me.ParaOutMemberID = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadGrid()
   VStrSQL = " Select *" & _
         " FROM Members where 1=1 and IsLockMember = 0 " & vMemberName & vBarcode & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "MemberID"
   Grid.Columns("Barcode").DataField = "Barcode"
   Grid.Columns("Name").DataField = "MemberName"
End Sub

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then Call BtnSelect_Click
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
   vOrder = " Order by " & Grid.Columns(ColIndex).DataField
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
   Case Asc("a") To Asc("z"), Asc("A") To Asc("Z")
      TxtMemberName.Text = Chr(KeyAscii): TxtMemberName.SelStart = Len(TxtMemberName.Text): TxtMemberName.SetFocus
   Case vbKey0 To vbKey9
      TxtMemberID.Text = Chr(KeyAscii): TxtMemberID.SelStart = Len(TxtMemberID.Text): TxtMemberID.SetFocus
   End Select
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtMemberID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtMemberID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "MemberID = " & TxtMemberID.Text, , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtMemberID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtBarCode_Change()
   On Error GoTo ErrorHandler
   vBarcode = IIf(TxtBarCode.Text = "", "", " and Barcode like '" & Replace(TxtBarCode.Text, "'", "''") & "%'")
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtMemberName_Change()
   On Error GoTo ErrorHandler
   vMemberName = IIf(TxtMemberName.Text = "", "", " and MemberName like '%" & Replace(TxtMemberName.Text, "'", "''") & "%'")
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtMemberName_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub
