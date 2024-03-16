VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SchStockExpiry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F8E8D6&
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
   Picture         =   "SchStockExpiry.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Begin VB.TextBox TxtExpiryID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2693
      TabIndex        =   1
      Top             =   1710
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5955
      TabIndex        =   2
      Top             =   1710
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   48431107
      CurrentDate     =   38244
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4845
      Left            =   2685
      TabIndex        =   0
      Top             =   2130
      Width           =   6630
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
      stylesets(0).Picture=   "SchStockExpiry.frx":E0B8
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
      Columns.Count   =   4
      Columns(0).Width=   1931
      Columns(0).Caption=   "Expiry ID"
      Columns(0).Name =   "ID"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2434
      Columns(1).Caption=   "Expiry Date"
      Columns(1).Name =   "Date"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2805
      Columns(2).Caption=   "Total Items"
      Columns(2).Name =   "TotalItems"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 5"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3969
      Columns(3).Caption=   "Amount"
      Columns(3).Name =   "Amount"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 5"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   11695
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
      Left            =   4703
      TabIndex        =   4
      Top             =   7110
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
      MICON           =   "SchStockExpiry.frx":E0D4
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   6023
      TabIndex        =   5
      Top             =   7110
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
      MICON           =   "SchStockExpiry.frx":E0F0
      BC              =   12632256
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpTo 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7530
      TabIndex        =   3
      Top             =   1710
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   48431107
      CurrentDate     =   38244
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
      Caption         =   "Expiry ID"
      Height          =   195
      Left            =   2685
      TabIndex        =   7
      Top             =   1470
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-------------  Expiry Date Range -------------"
      Height          =   195
      Left            =   6090
      TabIndex        =   6
      Top             =   1470
      Width           =   2640
   End
End
Attribute VB_Name = "SchStockExpiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vOrder As String, vDirection As String, vCol As Byte
Public ParaOutExpiryID As Long

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Set Rs = New ADODB.Recordset
'   Rs.Open "SELECT StockExpiryHeader.ExpiryID, convert(varchar(10),ExpiryDate,3) as ExpiryDate, SalemanName as SalemanName" & vbCrLf _
'         + ",(TotalAmount+((TotalAmount*PerSalesTax)/100))-TotalDiscount as TotalAmount,TotalItems" & vbCrLf _
'         + " FROM StockExpiryHeader INNER JOIN" & vbCrLf _
'         + " (SELECT ExpiryID, count(ExpiryID) as TotalItems FROM StockExpiryBody GROUP BY ExpiryID) StockExpiryBody" & vbCrLf _
'         + " ON StockExpiryHeader.ExpiryID = StockExpiryBody.ExpiryID" & vbCrLf _
'         + " INNER JOIN Parties ON StockExpiryHeader.PartyID = Parties.PartyID " & vbCrLf _
'         + " WHERE ExpiryDate Between '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'" & vOrder & vDirection, CN, adOpenStatic, adLockReadOnly
         
   Rs.Open "SELECT StockExpiryHeader.ExpiryID, convert(varchar(10),ExpiryDate,3) as ExpiryDate" & vbCrLf _
         + ",NetAmount,TotalItems" & vbCrLf _
         + " FROM StockExpiryHeader INNER JOIN" & vbCrLf _
         + " (SELECT ExpiryID, count(ExpiryID) as TotalItems FROM StockExpiryBody GROUP BY ExpiryID) StockExpiryBody" & vbCrLf _
         + " ON StockExpiryHeader.ExpiryID = StockExpiryBody.ExpiryID" & vbCrLf _
         + " WHERE ExpiryDate Between '" & DtpFrom.Value & "' AND '" & DtpTo.Value & "'" & vOrder & vDirection, CN, adOpenStatic, adLockReadOnly
   
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "ExpiryID"
   Grid.Columns("Date").DataField = "ExpiryDate"
   Grid.Columns("TotalItems").DataField = "TotalItems"
   Grid.Columns("Amount").DataField = "NetAmount"
 Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Me.ParaOutExpiryID = 0
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutExpiryID = Rs!ExpiryID
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
   Case TxtExpiryID.Name
      Call NonNumeric(KeyAscii, ActiveControl, True)
   End Select
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
  DtpFrom.Value = Date - 30
  DtpTo.Value = Date
  Me.ParaOutExpiryID = 0
  Call LoadGrid
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtExpiryID.Name, DtpFrom.Name, DtpTo.Name
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
      vOrder = " order by StockExpiryHeader.ExpiryID"
   Case 1
      vOrder = " order by StockExpiryHeader.ExpiryDate"
   Case 2
      vOrder = " order by TotalItems"
   Case 3
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
      TxtExpiryID.Text = Chr(KeyAscii): TxtExpiryID.SelStart = Len(TxtExpiryID.Text): TxtExpiryID.SetFocus
   End Select
End Sub

Private Sub Image1_Click()
   Unload Me
End Sub

Private Sub TxtExpiryID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtExpiryID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "ExpiryID = " & TxtExpiryID.Text, , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub



