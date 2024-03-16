VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmPurchasePendingInvoice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmPurchasePendingInvoiceOld.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   11205
      TabIndex        =   4
      Top             =   855
      Visible         =   0   'False
      Width           =   4200
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Tag             =   "NC"
         Text            =   "FrmPurchasePendingInvoiceOld.frx":0ECA
         Top             =   360
         Width           =   3930
      End
      Begin VB.Label LblClose 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3915
         TabIndex        =   6
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   4635
      TabIndex        =   0
      Top             =   6900
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
      MICON           =   "FrmPurchasePendingInvoiceOld.frx":0F55
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   2760
      TabIndex        =   1
      Top             =   6900
      Visible         =   0   'False
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
      MICON           =   "FrmPurchasePendingInvoiceOld.frx":0F71
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6090
      TabIndex        =   2
      Top             =   6900
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
      MICON           =   "FrmPurchasePendingInvoiceOld.frx":0F8D
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   2685
      Left            =   3008
      TabIndex        =   8
      Top             =   2880
      Width           =   5985
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   4
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
      stylesets(0).ForeColor=   -2147483634
      stylesets(0).BackColor=   -2147483635
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
      stylesets(0).Picture=   "FrmPurchasePendingInvoiceOld.frx":0FA9
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowColumnMoving=   2
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   4
      Columns(0).Width=   1693
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ProductID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   4180
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "ProductName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1640
      Columns(2).Caption=   "Qty"
      Columns(2).Name =   "Qty"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   2
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   2117
      Columns(3).Caption=   "Received Qty"
      Columns(3).Name =   "ReceivedQty"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   2
      Columns(3).FieldLen=   256
      _ExtentX        =   10557
      _ExtentY        =   4736
      _StockProps     =   79
      BackColor       =   15724527
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
   Begin SITextBox.Txt TxtPurchaseID 
      Height          =   315
      Left            =   3000
      TabIndex        =   9
      Top             =   1515
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      Mandatory       =   1
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpPurchaseDate 
      Height          =   315
      Left            =   4290
      TabIndex        =   10
      Top             =   1515
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
      AllowEdit       =   0   'False
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   7560
      TabIndex        =   13
      Tag             =   "NC"
      Top             =   1515
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   10
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   7560
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase ID"
      Height          =   195
      Left            =   3015
      TabIndex        =   12
      Top             =   1320
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Date"
      Height          =   195
      Left            =   4305
      TabIndex        =   11
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label LblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11205
      TabIndex        =   7
      Top             =   585
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Pending Invoice"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   180
      Width           =   3420
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11610
      Top             =   45
      Width           =   375
   End
End
Attribute VB_Name = "FrmPurchasePendingInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim RsDiff As New ADODB.Recordset
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.
Dim sSql As String
Dim vCounter As Integer

Private Sub BtnClear_Click()
  FormStatus = SelectionMode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
             If BtnSave.Enabled Then BtnSave_Click
             KeyCode = 0
'         Case vbKeyW
'             If BtnClear.Enabled Then BtnClear_Click
'             KeyCode = 0
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
'      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler

   CN.BeginTrans
    CN.Execute ("Insert into PurchaseHeader Select * from PurchasePendingHeader where Purid =" & TxtPurchaseID.Text & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "'")
    Grid.MoveFirst
       Set Rs = New ADODB.Recordset
       Rs.Open "Select * FROM PurchaseBody where Purid = " & TxtPurchaseID.Text & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "'", CN, adOpenDynamic, adLockOptimistic
       Set RsDiff = New ADODB.Recordset
       RsDiff.Open "Select * FROM DifferenceInvoice where InvoiceType = 'PI' And InvoiceId = " & TxtPurchaseID.Text & " And InvoiceDate = '" & DtpPurchaseDate.DateValue & "'", CN, adOpenDynamic, adLockOptimistic
        For vCounter = 1 To Grid.Rows
            With CN.Execute("Select * from PurchasePendingBody where Purid=" & TxtPurchaseID.Text & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "' And ProductID ='" & Grid.Columns("ProductID").Text & "'")
            'sSql = "Insert into PurchaseBody (PurID, PurchaseDate, ProductID, QtyLoose, Price, DiscPC, Amount, Code, DiscPer, DiscVal, PackingID, QtyPack, Multiplier, Bonus) Values (" & !PurID & ",'" & !PurchaseDate & "','" & !ProductID & "'," & !QtyLoose & "," & !Price & "," & !DiscPC & "," & !Amount & "," & !Code & "," & !DiscPer & "," & !DiscVal & "," & IIf(IsNull(!PackingID), Null, !PackingID) & "," & IIf(IsNull(!QtyPack), Null, !QtyPack) & "," & IIf(IsNull(!Multiplier), Null, !Multiplier) & "," & !Bonus & ") "
            Rs.AddNew
            Rs!PurID = !PurID
            Rs!PurchaseDate = !PurchaseDate
            Rs!ProductID = !ProductID
            Rs!QtyLoose = !QtyLoose
            Rs!Price = !Price
            Rs!DiscPC = !DiscPC
            Rs!Amount = !Amount
            Rs!Code = !Code
            Rs!DiscPer = !DiscPer
            Rs!DiscVal = !DiscVal
            Rs!PackingID = !PackingID
            Rs!QtyPack = !QtyPack
            Rs!Multiplier = !Multiplier
            Rs!Bonus = !Bonus
            .Close
            End With
            Rs.Update
            If Grid.Columns("Qty").Value <> Grid.Columns("ReceivedQty").Value Then
                RsDiff.AddNew
                RsDiff!InvoiceID = TxtPurchaseID.Text
                RsDiff!InvoiceDate = DtpPurchaseDate.DateValue
                RsDiff!InvoiceType = "PI"
                RsDiff!StoreID = TxtStoreID.Text
                RsDiff!ProductID = Grid.Columns("ProductID").Text
                If Grid.Columns("Qty").Value > Grid.Columns("ReceivedQty").Value Then
                    RsDiff!QtyMore = Grid.Columns("Qty").Value - Grid.Columns("ReceivedQty").Value
                Else
                    RsDiff!QtyLess = Grid.Columns("Qty").Value - Grid.Columns("ReceivedQty").Value
                End If
                RsDiff.Update
            End If
        Grid.MoveNext
        Next
    
    
    CN.Execute ("Delete PurchasePendingBody where Purid=" & TxtPurchaseID.Text & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "'")

   CN.CommitTrans
   MsgBox "Data Saved Successfully", vbInformation, "Alert"
   'If vIsNewRecord = False Then Call ActivityLog("Companies", eEdit, TxtID.Text)
'   Set Rs = New ADODB.Recordset
'   Rs.Open " Select * FROM Companies where CompanyID = '" & TxtID.Text & "'", CN, adOpenDynamic, adLockOptimistic
'   If vIsNewRecord Then
'      Rs.AddNew
'      Rs!companyid = TxtID.Text
'   End If
'   Rs!CompanyName = TxtName.Text
'   Rs.Update
   'If vIsNewRecord = True Then Call ActivityLog("Companies", eAdd, TxtID.Text)
'   FormStatus = NewMode
   BtnSave.Enabled = False
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
   CN.RollbackTrans
End Sub

Private Sub LblClose_Click()
   FraHelp.Visible = False
End Sub

Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblHelp.ForeColor = &H800000
   FraHelp.ZOrder 0
   FraHelp.Visible = True
End Sub

Private Sub LblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If LblHelp.FontUnderline = True Then Exit Sub
   LblHelp.FontUnderline = True
End Sub

Private Sub LblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblHelp.ForeColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Purchase Pending Invoice"
   HelpLocation Me
   Call GetPurchasePending
 '  FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
         BtnSave.Enabled = False
         BtnClear.Enabled = True
'         If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
         vIsNewRecord = True
     Case Is = OpenMode
        
         BtnClear.Enabled = True
'         Grid.Enabled = False
         vIsNewRecord = False
     Case Is = ChangeMode
         BtnSave.Enabled = True
     Case Is = SelectionMode
'         Grid.Enabled = True
        
         BtnSave.Enabled = False
         BtnClear.Enabled = False
'         TxtName.Enabled = False
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
          Set frmObj = Nothing
      Next
      Set Rs = Nothing
      Set FrmPurchasePendingInvoice = Nothing
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub GetPurchasePending()
   On Error GoTo ErrorHandler
   sSql = "select h.* FROM PurchasePendingHeader H where h.PurID = " & vPurchaseID & " And h.PurchaseDate = '" & vPurchaseDate & "'"
   With CN.Execute(sSql)
      If Not .BOF Then
          TxtPurchaseID.Text = !PurID
          DtpPurchaseDate.Date = !PurchaseDate
          TxtStoreID.Text = !StoreID
      End If
      .Close
   End With
   Call PopulateDataToGrid
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   sSql = "select p.productname, b.* from PurchasePendingBody b join products p on p.productid = b.productid where Purid = " & vPurchaseID & " And PurchaseDate = '" & vPurchaseDate & "'"
   With CN.Execute(sSql)
      Grid.Redraw = False
      Grid.MoveFirst
      Grid.RemoveAll
      Grid.AllowAddNew = True
      While Not .EOF
         Grid.AddNew
         Grid.Columns("ProductID").Text = !ProductID
         Grid.Columns("ProductName").Text = !ProductName
         Grid.Columns("Qty").Value = IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack) + !QtyLoose
         Grid.Columns("ReceivedQty").Value = IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack) + !QtyLoose
         .MoveNext
      Wend
      .Close
   End With
   Grid.MoveFirst
'      Grid.AddNew
'      Grid.Columns("ProductID").Text = " "
   Grid.AllowAddNew = False
   Grid.Redraw = True
End Sub

