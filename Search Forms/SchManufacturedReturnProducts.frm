VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form SchManufacturedReturnProducts 
   AutoRedraw      =   -1  'True
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
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1073
      TabIndex        =   1
      Top             =   1125
      Width           =   1365
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6660
      Left            =   1073
      TabIndex        =   0
      Top             =   1455
      Width           =   9855
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
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
      stylesets(0).Picture=   "SchManufacturedReturnProducts.frx":0000
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
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   6
      Columns(0).Width=   2381
      Columns(0).Caption=   "Manufactured ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2858
      Columns(1).Caption=   "Manufactured  Date"
      Columns(1).Name =   "Date"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).NumberFormat=   "dd/MM/yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   4577
      Columns(2).Caption=   "Store Name"
      Columns(2).Name =   "StoreName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 4"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1667
      Columns(3).Caption=   "Total Items"
      Columns(3).Name =   "TotalItems"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 5"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2434
      Columns(4).Caption=   "Amount"
      Columns(4).Name =   "Amount"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 5"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2990
      Columns(5).Caption=   "CO"
      Columns(5).Name =   "CO"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   17383
      _ExtentY        =   11747
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
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4703
      TabIndex        =   4
      Top             =   8310
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
      MICON           =   "SchManufacturedReturnProducts.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   6023
      TabIndex        =   5
      Top             =   8310
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
      MICON           =   "SchManufacturedReturnProducts.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   330
      Left            =   2940
      TabIndex        =   2
      Top             =   1110
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "/"
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpToDate 
      Height          =   330
      Left            =   4335
      TabIndex        =   3
      Top             =   1110
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "/"
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin VB.Label Label1 
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
      Left            =   1920
      TabIndex        =   8
      Top             =   180
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "------------ Manufactured Date --------------"
      Height          =   195
      Left            =   2970
      TabIndex        =   7
      Top             =   870
      Width           =   2640
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured ID"
      Height          =   195
      Left            =   1073
      TabIndex        =   6
      Top             =   915
      Width           =   1200
   End
End
Attribute VB_Name = "SchManufacturedReturnProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSQL As String
Dim vOrder As String, vDirection As String, vCol As Byte
Public ParaOutManufacturedID As Long

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Set Rs = New ADODB.Recordset
   vStrSQL = "SELECT h.ManufacturedID as ID, h.ManufacturedDate As Date" & vbCrLf _
         + " , StoreName, TotalAmount,TotalItems, UserName" & vbCrLf _
         + " FROM ManufacturedProductsHeader h INNER JOIN" & vbCrLf _
         + " (SELECT ManufacturedID, sum(qtyloose) as TotalItems, sum(amount) as TotalAmount FROM ManufacturedProductsUsed GROUP BY ManufacturedID) b" & vbCrLf _
         + " ON h.ManufacturedID = b.ManufacturedID " & vbCrLf _
         + " INNER JOIN Stores s  ON s.StoreID = h.StoreID " & vbCrLf _
         + " INNER JOIN users u ON h.userno = u.userno " & vbCrLf _
         + " WHERE h.ManufacturedDate between '" & DtpFromDate.DateValue & "' and '" & DtpToDate.DateValue & "'" & IIf(ObjUserSecurity.IsAdministrator = False, " and h.userno=" & ObjUserSecurity.UserNo, "") & vOrder & vDirection
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Me.ParaOutManufacturedID = 0
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutManufacturedID = Rs!ID
  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub DtpFromDate_Change()
   Call LoadGrid
End Sub

Private Sub DtpToDate_Change()
   Call LoadGrid
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me
   DtpFromDate.DateValue = Date - 30
   DtpToDate.DateValue = Date
   SetWindowText Me.hWnd, "Search"
   Me.ParaOutManufacturedID = 0
   vOrder = " order by ID"
   vDirection = " desc"
   Grid.Columns("ID").DataField = "ID"
   Grid.Columns("Date").DataField = "Date"
   Grid.Columns("StoreName").DataField = "StoreName"
   Grid.Columns("TotalItems").DataField = "TotalItems"
   Grid.Columns("Amount").DataField = "TotalAmount"
   Grid.Columns("CO").DataField = "username"
   Call LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtID.Name, DtpFromDate.Name, DtpToDate.Name
         Call BtnSelect_Click
      End Select
   End If
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
  If Grid.Rows > 0 Then BtnSelect_Click
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

'Private Sub Grid_KeyPress(KeyAscii As Integer)
'   Select Case KeyAscii
'   Case vbKey0 To vbKey9
'      TxtID.Text = Chr(KeyAscii): TxtID.SelStart = Len(TxtID.Text): TxtID.SetFocus
'   End Select
'End Sub

Private Sub Image1_Click()
   Unload Me
End Sub

Private Sub TxtID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "ID = " & TxtID.Text, , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
