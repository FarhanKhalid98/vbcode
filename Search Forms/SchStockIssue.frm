VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SchStockIssue 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F8E8D6&
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
   Begin VB.TextBox TxtSalemanName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3353
      TabIndex        =   8
      Top             =   3698
      Width           =   4740
   End
   Begin VB.TextBox TxtIssueID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2251
      TabIndex        =   1
      Top             =   3698
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8296
      TabIndex        =   2
      Top             =   3698
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   123011075
      CurrentDate     =   38244
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4845
      Left            =   2243
      TabIndex        =   0
      Top             =   4028
      Width           =   9855
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
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
      stylesets(0).Picture=   "SchStockIssue.frx":0000
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
      Columns.Count   =   5
      Columns(0).Width=   1931
      Columns(0).Caption=   "Issue ID"
      Columns(0).Name =   "ID"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2434
      Columns(1).Caption=   "Issue Date"
      Columns(1).Name =   "Date"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   6138
      Columns(2).Caption=   "Saleman Name"
      Columns(2).Name =   "SalemanName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 4"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2805
      Columns(3).Caption=   "Total Items"
      Columns(3).Name =   "TotalItems"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 5"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3545
      Columns(4).Caption=   "Amount"
      Columns(4).Name =   "Amount"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 5"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   17383
      _ExtentY        =   8546
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
      Left            =   5866
      TabIndex        =   4
      Top             =   9098
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
      MICON           =   "SchStockIssue.frx":001C
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7186
      TabIndex        =   5
      Top             =   9098
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
      MICON           =   "SchStockIssue.frx":0038
      BC              =   12632256
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpTo 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   9871
      TabIndex        =   3
      Top             =   3698
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   123011075
      CurrentDate     =   38244
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3000
      TabIndex        =   10
      Top             =   270
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   12788
      Top             =   2003
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Saleman Name"
      Height          =   195
      Left            =   3458
      TabIndex        =   9
      Top             =   3458
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue ID"
      Height          =   195
      Left            =   2243
      TabIndex        =   7
      Top             =   3458
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-------------  Purchase Date Range -------------"
      Height          =   195
      Left            =   8341
      TabIndex        =   6
      Top             =   3458
      Width           =   2895
   End
End
Attribute VB_Name = "SchStockIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vOrder As String, vDirection As String, vCol As Byte
Public ParaOutIssueID As Long
Public ParaOutIssueDate As Date

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Set Rs = New ADODB.Recordset
'   Rs.Open "SELECT StockIssueHeader.IssueID, convert(varchar(10),IssueDate,3) as IssueDate, SalemanName as SalemanName" & vbCrLf _
'         + ",(TotalAmount+((TotalAmount*PerSalesTax)/100))-TotalDiscount as TotalAmount,TotalItems" & vbCrLf _
'         + " FROM StockIssueHeader INNER JOIN" & vbCrLf _
'         + " (SELECT IssueID, count(IssueID) as TotalItems FROM StockIssueBody GROUP BY IssueID) StockIssueBody" & vbCrLf _
'         + " ON StockIssueHeader.IssueID = StockIssueBody.IssueID" & vbCrLf _
'         + " INNER JOIN Parties ON StockIssueHeader.PartyID = Parties.PartyID " & vbCrLf _
'         + " WHERE IssueDate Between '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'" & vOrder & vDirection, CN, adOpenStatic, adLockReadOnly
         
   Rs.Open "SELECT StockIssueHeader.IssueID, convert(varchar(10),IssueDate,3) as IssueDate, SalemanName as SalemanName" & vbCrLf _
         + ",NetAmount,TotalItems" & vbCrLf _
         + " FROM StockIssueHeader INNER JOIN" & vbCrLf _
         + " (SELECT IssueID, count(IssueID) as TotalItems FROM StockIssueBody GROUP BY IssueID) StockIssueBody" & vbCrLf _
         + " ON StockIssueHeader.IssueID = StockIssueBody.IssueID" & vbCrLf _
         + " INNER JOIN Salesman ON StockIssueHeader.SalemanID = Salesman.SalemanID " & vbCrLf _
         + " WHERE IssueDate Between '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'" & vOrder & vDirection, CN, adOpenStatic, adLockReadOnly
   
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "IssueID"
   Grid.Columns("Date").DataField = "IssueDate"
   Grid.Columns("SalemanName").DataField = "SalemanName"
   Grid.Columns("TotalItems").DataField = "TotalItems"
   Grid.Columns("Amount").DataField = "NetAmount"
 Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Me.ParaOutIssueID = 0
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutIssueID = Rs!issueId
  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub DtpFrom_Change()
   Call LoadGrid
End Sub

Private Sub DtpTo_Change()
   Call LoadGrid
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case ActiveControl.Name
   Case TxtIssueID.Name
      Call NonNumeric(KeyAscii, ActiveControl, True)
   End Select
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"
   DtpFrom.Value = Date - 30
   DtpTo.Value = Date
   Me.ParaOutIssueID = 0
   Call LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtIssueID.Name, DtpFrom.Name, DtpTo.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then BtnSelect_Click
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
Select Case ColIndex
   Case 0
      vOrder = " order by StockIssueHeader.IssueID"
   Case 1
      vOrder = " order by StockIssueHeader.IssueDate"
   Case 2
      vOrder = " order by SalemanName"
   Case 3
      vOrder = " order by TotalItems"
   Case 4
      vOrder = " order by NetAmount"
End Select
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
   Case vbKey0 To vbKey9
      TxtIssueID.Text = Chr(KeyAscii): TxtIssueID.SelStart = Len(TxtIssueID.Text): TxtIssueID.SetFocus
   Case Asc("a") To Asc("z"), Asc("A") To Asc("Z")
      TxtSalemanName.Text = Chr(KeyAscii): TxtSalemanName.SelStart = Len(TxtSalemanName.Text): TxtSalemanName.SetFocus
   End Select
End Sub

Private Sub Image1_Click()
   Unload Me
End Sub

Private Sub TxtIssueID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtIssueID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "IssueID = " & TxtIssueID.Text, , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub



Private Sub TxtSalemanName_Change()
On Error GoTo ErrorHandler
   If Trim(TxtSalemanName.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "SalemanName like '" & TxtSalemanName.Text & "%'", , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage

End Sub
