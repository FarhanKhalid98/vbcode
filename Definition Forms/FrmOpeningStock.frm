VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmOpeningStock 
   AutoRedraw      =   -1  'True
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
   Picture         =   "FrmOpeningStock.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   1575
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   7
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6195
      Left            =   1440
      TabIndex        =   8
      Top             =   1890
      Width           =   9120
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   5
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
      stylesets(0).Picture=   "FrmOpeningStock.frx":77CC
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedRow"
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
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   5
      Columns(0).Width=   2408
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   6826
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2143
      Columns(2).Caption=   "Opening Qty"
      Columns(2).Name =   "Qty"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   5
      Columns(2).NumberFormat=   "########.##"
      Columns(2).FieldLen=   256
      Columns(3).Width=   1931
      Columns(3).Caption=   "Pur Price"
      Columns(3).Name =   "PurPrice"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   5
      Columns(3).FieldLen=   256
      Columns(4).Width=   2249
      Columns(4).Caption=   "Amount"
      Columns(4).Name =   "Amount"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   5
      Columns(4).FieldLen=   256
      _ExtentX        =   16087
      _ExtentY        =   10927
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
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   4043
      TabIndex        =   2
      Top             =   8310
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
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
      MICON           =   "FrmOpeningStock.frx":77E8
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   5363
      TabIndex        =   3
      Top             =   8310
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Reset"
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
      MICON           =   "FrmOpeningStock.frx":7804
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6683
      TabIndex        =   4
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
      MICON           =   "FrmOpeningStock.frx":7820
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   2445
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
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
      MICON           =   "FrmOpeningStock.frx":783C
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2805
      TabIndex        =   10
      Top             =   1575
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   6675
      TabIndex        =   1
      Top             =   1575
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DecimalPoint    =   3
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   7890
      TabIndex        =   11
      Top             =   1575
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   8985
      TabIndex        =   12
      Top             =   1575
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Price"
      Height          =   195
      Left            =   7875
      TabIndex        =   14
      Top             =   1365
      Width           =   645
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   9015
      TabIndex        =   13
      Top             =   1350
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   6690
      TabIndex        =   9
      Top             =   1350
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   2835
      TabIndex        =   7
      Top             =   1365
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      Height          =   195
      Left            =   1425
      TabIndex        =   6
      Top             =   1365
      Width           =   765
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   60
      Width           =   345
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmOpeningStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vCounter As Integer
Dim vIsNewRecord As Boolean
Dim RsBody As New ADODB.Recordset
Dim Flag As Boolean
Dim sSql As String
Dim vStrSQL As String
Dim vIsNewRow As Boolean

Private Sub BtnClear_Click()
 On Error GoTo ErrorHandler
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtQty.SetFocus
   Else
      TxtProductID.SetFocus
   End If
End Sub

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   '-- when Product ID is written then it will check and all its related value will be write its appropriate places
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      'SchProduct.ParaInSaleID = Val(TxtSaleID.Text)
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
   vStrSQL = "SELECT ProductName, PurPrice from Products where ProductID='" & TxtProductID.Text & "'"
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductName.Text = !ProductName
         TxtPurPrice.Text = !PurPrice
         FunSelectProduct = True
         .Close
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         TxtProductID.Text = ""
         TxtProductName.Text = ""
         TxtPurPrice.Text = ""
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
          Set frmObj = Nothing
      Next
      Set RsBody = Nothing
      Set FrmOpeningStock = Nothing
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub TxtProductID_Change()
   If ActiveControl.Name <> TxtProductID.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
      TxtPurPrice.Text = ""
   End If
End Sub

Private Sub TxtProductID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtProductId_Validate(Cancel As Boolean)
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtProductID.Text) = "" Then Exit Sub
   vTemp = FunSelectProduct(ssValidate, False)
   If vTemp = False Then
      vTemp = FunSelectProduct(ssButton, False)
      Cancel = False
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  'If VIsPosted And ObjUserSecurity.IsAdministrator = False Then
  '  MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
  '  Exit Sub
  'End If
'  Header Validation
   RsBody.Filter = ""
   If Grid.Rows = 1 Then
      MsgBox "Enter atleast one product to save", vbExclamation, "Alert"
      TxtProductID.SetFocus
      Exit Sub
   End If
   
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
   CN.BeginTrans
   '-------------------------------------------------------------------------
   RsBody.UpdateBatch
   CN.CommitTrans
   'CN.Execute "exec SpcurrentStock"
   Grid.Redraw = True
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
  SetWindowText Me.hWnd, "Opening Stock"
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
      'If Me.ActiveControl.Name = Grid.Name And Grid.AddItemRowIndex(Grid.Bookmark) = Grid.Rows - 1 Then
      '   Grid.Update
      'End If
    End If
    If KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtProductID.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtQty.SetFocus
      End Select
    ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtProductID.Name Then
         KeyCode = 0
         BtnSave.SetFocus
      End If
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyS
                If BtnSave.Enabled Then BtnSave_Click
                KeyCode = 0
            Case vbKeyQ
                If BtnClose.Enabled Then BtnClose_Click
                KeyCode = 0
        End Select
    End If
    Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Or UCase(Me.ActiveControl.Name) Like "DTP*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Property Get FormStatus() As FormMode
  'Nothing
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
   'Based upon the value of vNewValue, we shall decide what controls to enable/disable
   On Error GoTo ErrorHandler
   vMode = vNewValue
   Select Case vNewValue
   Case Is = NewMode
      Call SubClearFields
      If RsBody.State = adStateOpen Then RsBody.Close
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      PopulateDataToGrid
      TxtProductID.Enabled = True
      vIsNewRecord = True
      vIsNewRow = True
   Case Is = OpenMode
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      vIsNewRecord = False
      vIsNewRow = True
   Case Is = ChangeMode
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         ctl.Text = ""
      ElseIf TypeOf ctl Is ComboBox Then
      End If
   Next
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ID").Text = " "
   Grid.Update
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from OpeningStock order by productid", CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      '================================================
      sSql = "select pr.Productname, os.* from OpeningStock os join Products pr on os.Productid=pr.Productid order by os.productid"
      With CN.Execute(sSql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ID").Text = !Productid
            Grid.Columns("Name").Text = !ProductName
            Grid.Columns("Qty").Value = !Qty
            Grid.Columns("PurPrice").Value = !PurPrice
            Grid.Columns("Amount").Value = !Amount
            .MoveNext
         Wend
         .Close
         Grid.Row = 0
      End With
      Grid.AddNew
      Grid.Columns("id").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
   Grid.FirstRow = 0
End Sub

Private Sub GetDataFromTexBoxesToGrid()
On Error GoTo ErrorHandler
   
   If Trim(TxtProductID.Text) = "" Then
      'MsgBox "Enter Group ID.", vbExclamation, "Alert"
      If TxtProductID.Enabled = True Then TxtProductID.SetFocus
      Exit Sub
   End If
   
   If Trim(TxtQty.Text) = "" Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      If TxtQty.Enabled = True Then TxtQty.SetFocus
      Exit Sub
   End If
      
   '-------------------------------------------------------------------

   RsBody.Filter = "ProductID='" & TxtProductID.Text & "'"
   If vIsNewRow Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ID").Text = TxtProductID.Text
         RsBody!Productid = TxtProductID.Text
      Else
         MsgBox "The record already exist"
         SubClearDetailArea
         If TxtProductID.Enabled Then TxtProductID.SetFocus
         Exit Sub
      End If
   End If
   Grid.Redraw = False
   With Grid
      .Columns("Name").Text = TxtProductName.Text
      .Columns("Qty").Text = TxtQty.Text
      .Columns("PurPrice").Text = TxtPurPrice.Text
      .Columns("Amount").Text = TxtAmount.Text
      RsBody!Qty = Val(TxtQty.Text)
      RsBody!PurPrice = Val(TxtPurPrice.Text)
      RsBody!Amount = Val(TxtAmount.Text)
      .MoveLast
      If Trim(.Columns("ID").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("id").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearDetailArea
   TxtProductID.SetFocus
   vIsNewRow = True
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtPurPrice_Change()
   Call SubCalculate
End Sub

Private Sub TxtQty_Change()
   Call SubCalculate
End Sub

Private Sub TxtQty_LostFocus()
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub SubCalculate()
   TxtAmount.Text = Val(TxtQty.Text) * Val(TxtPurPrice.Text)
End Sub

Private Sub SubClearDetailArea()
   TxtProductID.Enabled = True
   TxtProductID.Text = ""
   TxtProductName.Text = ""
   TxtQty.Text = ""
   TxtPurPrice.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ID").Text
      TxtProductName.Text = .Columns("Name").Text
      TxtQty.Text = .Columns("Qty").Text
      TxtPurPrice.Text = .Columns("PurPrice").Text
      TxtAmount.Text = .Columns("Amount").Text
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   Call Grid_LostFocus
End Sub

Private Sub Grid_GotFocus()
   Flag = True
   TxtProductID.Enabled = False
  End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("ID").Text) = "" Then
      TxtProductID.Text = ""
      TxtProductID.Enabled = True
      TxtProductID.SetFocus
      vIsNewRow = True
   Else
      TxtProductID.Enabled = False
      If TxtQty.Enabled = True And TxtQty.Visible Then TxtQty.SetFocus
      vIsNewRow = False
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(Grid.Columns("ID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ID").Text) = "" Then Exit Sub
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   RsBody.Filter = "ProductID='" & Grid.Columns("ID").Text & "'"
   If RsBody.RecordCount > 0 Then RsBody.Delete
   RsBody.Filter = ""
   Grid.DeleteSelected
   Grid.SelBookmarks.RemoveAll
   Grid.Refresh
   GetDataBackFromGridToTexBoxes
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
