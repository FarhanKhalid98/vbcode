VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmBankChequeIssueReconciliation 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11985
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   Icon            =   "frmBankChqIssuReconciliation.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmBankChqIssuReconciliation.frx":0ECA
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton btnClose 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7950
      TabIndex        =   7
      Top             =   7470
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmBankChqIssuReconciliation.frx":783B
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6630
      TabIndex        =   3
      Top             =   7470
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmBankChqIssuReconciliation.frx":7857
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5310
      TabIndex        =   4
      Top             =   7470
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmBankChqIssuReconciliation.frx":7873
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2670
      TabIndex        =   6
      Top             =   7470
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmBankChqIssuReconciliation.frx":788F
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btndelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3990
      TabIndex        =   5
      Top             =   7470
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmBankChqIssuReconciliation.frx":78AB
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      CausesValidation=   0   'False
      Height          =   4710
      Left            =   1095
      TabIndex        =   2
      Top             =   2250
      Width           =   9795
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   7
      stylesets.count =   3
      stylesets(0).Name=   "Style1"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   13817275
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmBankChqIssuReconciliation.frx":78C7
      stylesets(1).Name=   "style"
      stylesets(1).ForeColor=   16777215
      stylesets(1).BackColor=   8388608
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "frmBankChqIssuReconciliation.frx":78E3
      stylesets(2).Name=   "StyleCol"
      stylesets(2).BackColor=   14548991
      stylesets(2).Picture=   "frmBankChqIssuReconciliation.frx":78FF
      MultiLine       =   0   'False
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
      SelectTypeRow   =   0
      SelectByCell    =   -1  'True
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ActiveRowStyleSet=   "Style1"
      Columns.Count   =   7
      Columns(0).Width=   3200
      Columns(0).Caption=   "Cheque No"
      Columns(0).Name =   "ChequeNo"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   2461
      Columns(1).Caption=   "Cheque Date"
      Columns(1).Name =   "ChequeDate"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd/MM/yyyy"
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2461
      Columns(2).Caption=   "Amount"
      Columns(2).Name =   "Amount"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   3942
      Columns(3).Caption=   "Party Name"
      Columns(3).Name =   "PartyName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1588
      Columns(4).Caption=   "Reconcile"
      Columns(4).Name =   "Reconcile"
      Columns(4).Alignment=   2
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   11
      Columns(4).FieldLen=   256
      Columns(4).Style=   2
      Columns(4).HasBackColor=   -1  'True
      Columns(4).BackColor=   16777215
      Columns(4).Nullable=   0
      Columns(5).Width=   1535
      Columns(5).Caption=   "Bounce"
      Columns(5).Name =   "Bounce"
      Columns(5).Alignment=   2
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Style=   2
      Columns(5).HasBackColor=   -1  'True
      Columns(5).BackColor=   16777215
      Columns(5).Nullable=   0
      Columns(6).Width=   1588
      Columns(6).Caption=   "Return"
      Columns(6).Name =   "Return"
      Columns(6).Alignment=   2
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   11
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      Columns(6).HasBackColor=   -1  'True
      Columns(6).BackColor=   16777215
      Columns(6).Nullable=   0
      TabNavigation   =   1
      _ExtentX        =   17277
      _ExtentY        =   8308
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSCalendarWidgets_A.SSDateCombo dtpFromDate 
      Height          =   315
      Left            =   2205
      TabIndex        =   0
      Top             =   1620
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
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
      Height          =   315
      Left            =   4455
      TabIndex        =   1
      Top             =   1620
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   195
      Left            =   3780
      TabIndex        =   10
      Top             =   1665
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      Height          =   195
      Left            =   1380
      TabIndex        =   9
      Top             =   1665
      Width           =   735
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Cheque Issuance Reconciliation "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1845
      TabIndex        =   8
      Top             =   135
      Width           =   4845
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   11610
      Top             =   60
      Width           =   345
   End
End
Attribute VB_Name = "FrmBankChequeIssueReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSql As String
Dim vCounter As Integer
Dim vMode As FormMode
Dim RsBody As New ADODB.Recordset

Private Sub btnClear_Click()
   FormStatus = NewMode
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub dtpFromDate_Change()
   Call PopulateDataToGrid
End Sub

Private Sub DtpToDate_Change()
   Call PopulateDataToGrid
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If Not (UCase(ActiveControl.Name) Like UCase("txt*")) Then Exit Sub
   If btnSave.Enabled = False Then FormStatus = changemode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If btnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
         Set frmObj = Nothing
      Next
         Set RsBody = Nothing
         Set FrmBankChequeIssueReconciliation = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_Click()
   On Error GoTo ErrorHandler
   With Grid
      If .Col = 4 Or .Col = 5 Or .Col = 6 Then
         .Columns("Reconcile").Value = 0
         .Columns("Bounce").Value = 0
         .Columns("Return").Value = 0
         .ActiveCell.Value = -1
         FormStatus = changemode
      End If
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Call Grid_Click
End Sub

Private Sub Grid_LostFocus()
   Grid.ActiveCell.StyleSet = ""
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   SetWindowText Me.hwnd, "Bank Cheque Issuance Reconciliation"
   FormStatus = NewMode
End Sub

Private Property Get FormStatus() As FormMode
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
  On Error GoTo ErrorHandler
   vMode = vNewValue
   Select Case vNewValue
   Case Is = NewMode
      Call SubClearFields
      If RsBody.State = adStateOpen Then RsBody.Close
      btnOpen.Enabled = True
      btndelete.Enabled = False
      btnSave.Enabled = False
      btnClear.Enabled = True
      PopulateDataToGrid
   Case Is = OpenMode
      btnOpen.Enabled = True
      btndelete.Enabled = True
      btnClear.Enabled = True
      btnSave.Enabled = False
   Case Is = changemode
      btnOpen.Enabled = False
      btndelete.Enabled = False
      btnSave.Enabled = True
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
      ElseIf TypeOf ctl Is SITextBox.txt Then
         ctl.Text = ""
      End If
   Next
   DtpToDate.DateValue = Date
   dtpFromDate.DateValue = Date
   Grid.CancelUpdate
   Grid.RemoveAll
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "select * from BankChequeIssueBody where ActChequeDate Between ' " & dtpFromDate.DateValue & " ' and  '" & DtpToDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = "Select H.ACPayeeName, B.* From BankChequeIssueBody B  Inner Join BankChequeIssueHeader H on B.VoucherID = H.VoucherID Where ActChequeDate Between  '" & dtpFromDate.DateValue & "' And '" & DtpToDate.DateValue & "'"
      With CN.Execute(sSql)
         If .RecordCount > 0 Then
            Grid.Redraw = False
            Grid.MoveFirst
            Grid.RemoveAll
            Grid.AllowAddNew = True
            While Not .EOF
               Grid.AddNew
               Grid.Columns("ChequeNo").Text = IIf(IsNull(!ActChequeNo), "", !ActChequeNo)
               Grid.Columns("ChequeDate").Text = (!ActChequeDate)
               Grid.Columns("amount").Value = Val(!ActAmount)
               Grid.Columns("PartyName").Text = IIf(IsNull(!ACPayeeName), "", !ACPayeeName)
               Grid.Columns("Reconcile").Value = IIf(IsNull(!Reconcile), 0, !Reconcile)
               Grid.Columns("Bounce").Value = IIf(IsNull(!Bounce), 0, !Bounce)
               Grid.Columns("Return").Value = IIf(IsNull(!ReturnChq), 0, !ReturnChq)
               .MoveNext
            Wend
         End If
         .Close
      End With
      Grid.AllowAddNew = False
      Grid.Redraw = True
      Grid.MoveFirst
   End If
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
      If KeyCode = vbKeyReturn Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      ElseIf Shift = vbCtrlMask Then
         Select Case KeyCode
            Case vbKeyS
               If btnSave.Enabled = True Then btnSave_Click
               KeyCode = 0
            Case vbKeyW
               If btnClear.Enabled = True Then btnClear_Click
               KeyCode = 0
            Case vbKeyQ
               If BtnClose.Enabled = True Then BtnClose_Click
               KeyCode = 0
'            Case vbKeyO
'               If btnOpen.Enabled = True Then BtnOpen_Click
'               KeyCode = 0
'            Case vbKeyP
'               If btnPrint.Enabled = True Then BtnPrint_Click
'               KeyCode = 0
'            Case vbKeyR
'               If btndelete.Enabled = True Then btndelete_Click
'               KeyCode = 0
         End Select
      End If
   Exit Sub
ErrorHandler:
     Call ShowErrorMessage
End Sub

Private Sub Grid_GotFocus()
   Grid.ActiveCell.StyleSet = "StyleCol"
   Grid.Col = 4
End Sub

Private Sub btnSave_Click()
   On Error GoTo ErrorHandler
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
'      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
   RsBody.Filter = 0
   RsBody.MoveFirst
   Grid.MoveFirst
   For vCounter = 1 To RsBody.RecordCount
      RsBody!Reconcile = Grid.Columns("Reconcile").Value
      RsBody!Bounce = Grid.Columns("Bounce").Value
      RsBody!ReturnChq = Grid.Columns("Return").Value
      RsBody.Update
      Grid.MoveNext
      RsBody.MoveNext
   Next vCounter
   RsBody.UpdateBatch
   RsBody.MoveFirst
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
