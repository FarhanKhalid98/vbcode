VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptBankChequeReceive 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptBankChequeReceive.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptReconcile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reconcile"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5340
      TabIndex        =   5
      Top             =   6038
      Width           =   1185
   End
   Begin VB.OptionButton OptBounce 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bounce"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6510
      TabIndex        =   6
      Top             =   6038
      Width           =   1140
   End
   Begin VB.OptionButton OptReturn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Return"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7665
      TabIndex        =   7
      Top             =   6038
      Width           =   1185
   End
   Begin VB.OptionButton OptAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8835
      TabIndex        =   8
      Top             =   6038
      Value           =   -1  'True
      Width           =   1185
   End
   Begin SITextBox.Txt TxtChequeNo 
      Height          =   315
      Left            =   5340
      TabIndex        =   3
      Top             =   4598
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   16
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
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton BtnChequeNo 
      Height          =   330
      Left            =   6960
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4598
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
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
      MICON           =   "RptBankChequeReceive.frx":0ECA
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   315
      Left            =   7320
      TabIndex        =   1
      Top             =   3893
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
      Left            =   8625
      TabIndex        =   2
      Top             =   3893
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
   Begin SITextBox.Txt TxtVoucherID 
      Height          =   315
      Left            =   5340
      TabIndex        =   0
      Top             =   3893
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   10
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
      IntegralPoint   =   15
      Mandatory       =   1
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   195
         Left            =   1995
         TabIndex        =   23
         Top             =   0
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   195
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   30
      End
   End
   Begin JeweledBut.JeweledButton BtnVoucher 
      Height          =   330
      Left            =   6960
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3893
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
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
      MICON           =   "RptBankChequeReceive.frx":0EE6
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8430
      TabIndex        =   11
      Top             =   7028
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Close"
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
      MICON           =   "RptBankChequeReceive.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   5655
      TabIndex        =   9
      Top             =   7028
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Pre&view"
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
      MICON           =   "RptBankChequeReceive.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7035
      TabIndex        =   10
      Top             =   7028
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Print"
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
      MICON           =   "RptBankChequeReceive.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnReceive 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6960
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5333
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
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
      MICON           =   "RptBankChequeReceive.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtReceivingID 
      Height          =   315
      Left            =   5340
      TabIndex        =   4
      Top             =   5333
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   15
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
      IntegralPoint   =   3
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtReceivingName 
      Height          =   315
      Left            =   7320
      TabIndex        =   25
      Tag             =   "nc"
      Top             =   5333
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Left            =   8625
      TabIndex        =   14
      Top             =   3683
      Width           =   705
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      Left            =   7320
      TabIndex        =   15
      Top             =   3683
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving Name"
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
      Left            =   7305
      TabIndex        =   16
      Top             =   5123
      Width           =   1410
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No"
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
      Left            =   5340
      TabIndex        =   17
      Top             =   4403
      Width           =   960
   End
   Begin VB.Label LblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Cheque Receive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2700
      TabIndex        =   18
      Top             =   270
      Width           =   3510
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving ID"
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
      Left            =   5340
      TabIndex        =   19
      Top             =   5123
      Width           =   1125
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   -1410
      TabIndex        =   21
      Top             =   7635
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher ID"
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
      Left            =   5340
      TabIndex        =   20
      Top             =   3683
      Width           =   975
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove this Row"
      End
   End
End
Attribute VB_Name = "RptBankChequeReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql As String
Dim VStrSQL As String

Private Function FunSelectPayee(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchActReceiving.Show vbModal, Me
      If SchActReceiving.ParaOutID = "" Then FunSelectPayee = False: Exit Function
      TxtReceivingID.Text = SchActReceiving.ParaOutID
   End If
   Dim VStrSQL As String
   VStrSQL = "select * from Parties where PartyID = '" & Val(TxtReceivingID.Text) & "'"
   With CN.Execute(VStrSQL)
         If .RecordCount > 0 Then
            TxtReceivingName.Text = !PartyName
            .Close
            FunSelectPayee = True
            Exit Function
         Else
            FunSelectPayee = False
            .Close
            TxtReceivingID.Text = ""
            TxtReceivingName.Text = ""
         End If
      End With
      Exit Function
ErrorHandler:
      Call ShowErrorMessage
End Function

Private Function FunSelectChequeNo(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchReceivingCheque.Show vbModal, Me
      If SchReceivingCheque.ParaOutID = "" Then FunSelectChequeNo = False: Exit Function
      TxtChequeNo.Text = SchReceivingCheque.ParaOutID
   End If
   Dim VStrSQL As String
   VStrSQL = "select ActChequeNo from BankChqRcvBody Where ActChequeNo =  '" & Val(TxtChequeNo.Text) & "'"
   With CN.Execute(VStrSQL)
         If .RecordCount > 0 Then
            .Close
            FunSelectChequeNo = True
            Exit Function
         Else
            FunSelectChequeNo = False
            .Close
'            TxtActPayeeName.Text = ""
'            TxtChequeNo.Text = ""
'            TxtVendorAddress.Text = ""
'            TxtVendorCity.Text = ""
         End If
      End With
      Exit Function
ErrorHandler:
      Call ShowErrorMessage
End Function

Private Function FunSelectVoucher(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchChqReceive.Show vbModal, Me
      If SchChqReceive.ParaOutVoucherNo = Null Then FunSelectVoucher = False: Exit Function
      TxtVoucherID.Text = SchChqReceive.ParaOutVoucherNo
   End If
    '---------------------------
    If Trim(TxtVoucherID.Text) = "" Then Exit Function
'    If Len(TxtVoucherID.Text) <= 5 Then
'      TxtVoucherID.Text = Right("00000" + CStr(Val(TxtVoucherID.Text)), 5)
'    End If
    If TxtVoucherID.Text = "" Then FunSelectVoucher = False: Exit Function
    VStrSQL = " SELECT VoucherID, VoucherDate From BankChequeReceiveHeader Where VoucherID = " & TxtVoucherID.Text
  
   With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtVoucherID.Text = !VoucherID
         DtpFromDate.Date = !VoucherDate
         DtpToDate.Date = !VoucherDate
         FunSelectVoucher = True
         .Close
         Exit Function
      Else
         FunSelectVoucher = False
         .Close
         TxtVoucherID.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnChequeNo_Click()
   If FunSelectChequeNo(ssButton, True) = True Then
      TxtReceivingID.SetFocus
   Else
      TxtChequeNo.SetFocus
   End If
End Sub

Private Sub BtnVoucher_Click()
   If FunSelectVoucher(ssButton, True) = True Then
      DtpFromDate.SetFocus
   Else
      TxtVoucherID.SetFocus
   End If
End Sub

Private Sub BtnReceive_Click()
   If FunSelectPayee(ssButton, False) = True Then
      BtnPreview.SetFocus
   Else
      TxtReceivingID.SetFocus
   End If
End Sub

Private Sub TxtChequeNo_Change()
   If TxtChequeNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtChequeNo.Name Then Exit Sub
End Sub

Private Sub TxtChequeNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtChequeNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtChequeNo.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectChequeNo(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectChequeNo(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtVoucherID_Change()
   If TxtVoucherID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtVoucherID.Name Then Exit Sub
End Sub

Private Sub TxtVoucherID_Validate(Cancel As Boolean)
  If Me.ActiveControl.Name <> TxtVoucherID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtVoucherID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVoucher(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVoucher(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtReceivingID_change()
   If TxtReceivingID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtReceivingID.Name Then Exit Sub
   If TxtReceivingName.Text <> "" Then TxtReceivingName.Text = ""
End Sub

Private Sub TxtReceivingID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtReceivingID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtReceivingID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectPayee(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectPayee(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub TxtReceivingName_Change()
'   If TxtReceivingName.Visible = False Then Exit Sub
'   If ActiveControl.Name <> TxtReceivingName.Name Then Exit Sub
'End Sub
'
'Private Sub TxtReceivingName_Validate(Cancel As Boolean)
'If Me.ActiveControl.Name <> TxtReceivingName.Name Then Exit Sub
'   On Error GoTo ErrorHandler
'   If TxtReceivingName.Text = "" Then Exit Sub
'   Dim vTemp As Boolean
'   vTemp = Not FunSelectDepositBy(ssValidate, True)
'   If vTemp = True Then
'      vTemp = Not FunSelectDepositBy(ssButton, False)
'   End If
'   Cancel = vTemp
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
       RptReportViewer.Caption = "Cheque Receive"
       RptReportViewer.Show vbModal
   End If
End Sub

Private Sub BtnPrint_Click()
    If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
         Case vbKeyV
            If BtnPreview.Enabled Then BtnPreview_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtVoucherID.Name: If FunSelectVoucher(ssFunctionKey, True) = True Then DtpFromDate.SetFocus
         Case TxtChequeNo.Name: If FunSelectChequeNo(ssFunctionKey, True) = True Then TxtReceivingID.SetFocus
         Case TxtReceivingID.Name: If FunSelectPayee(ssFunctionKey, True) = True Then BtnPreview.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Bank Cheque Receive"
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
   Set RptBankChequeReceive = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function SetReport() As Boolean
    On Error GoTo ErrorHandler
    Dim RsReport As New ADODB.Recordset
    SetReport = False
'    VStrSQL = "Select OS.*, (Isnull(OS.QtyPack,0) * Isnull(OS.Multiplier,0)) + Isnull(OS.QtyLoose,0) NetQty, P.ProductName, PK.PackingName, S.StoreName, C.CompanyName, G.GroupName, SG.SubGroupName" & vbCrLf _
'    + "from OpeningStock OS" & vbCrLf _
'    + "Left Outer Join  Products P on P.ProductID = OS.ProductID" & vbCrLf _
'    + "Left Outer Join   Packings PK on OS.PackingID = PK.PackingID" & vbCrLf _
'    + "Left Outer Join   Stores S on S.VoucherID = OS.VoucherID" & vbCrLf _
'    + "Left Outer Join Companies C on P.comPanyID = C.comPanyID" & vbCrLf _
'    + "Left Outer Join Groups G on P.GroupID = G.GroupID" & vbCrLf _
'    + "Left Outer Join SubGroups SG on P.SubGroupID = SG.SubGroupID Where 1=1 " & IIf(Trim(TxtVoucherID.Text) = "", "", " And S.VoucherID = " & TxtVoucherID.Text) & vbCrLf _
'    + "" & IIf(Trim(txtBankID.Text) = "", "", " And C.CompanyID = " & txtBankID.Text) & vbCrLf _
'    + "" & IIf(Trim(TxtReceivingID.Text) = "", "", " And G.GroupID = " & TxtReceivingID.Text) & vbCrLf _
'    + "" & IIf(Trim(TxtReceivingName.Text) = "", "", " And SG.SubGroupID = " & TxtReceivingName.Text) & vbCrLf _
'    + "" & IIf(Trim(TxtChequeNo.Text) = "", "", " And P.ProductID = " & TxtChequeNo.Text)
    Me.MousePointer = vbHourglass
'    If RsReport.State = adStateOpen Then RsReport.Close
'   RsReport.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
   CrptBankChequeReceiveParameter.DiscardSavedData
   Set RsReport = CN.Execute("ProdRptBankChequeReceive '" & DtpFromDate.Date & "','" & DtpToDate.Date & "','" & TxtVoucherID.Text & "','" & TxtReceivingID.Text & "','" & TxtChequeNo.Text & "'," & Abs(OptReconcile.Value) & "," & Abs(OptBounce.Value) & "," & Abs(OptReturn.Value))
   Set RptReportViewer.Report = New CrptBankChequeReceiveParameter
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
    
   If RsReport.BOF Then
       MsgBox "No record exists.", vbInformation, Me.Caption
       Me.MousePointer = vbDefault
       Exit Function
   End If
   'RptReportViewer.Report.Database.SetDataSource RsReport
    
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port

   'RptReportViewer.Report.PaperOrientation = crLandscape
   SetReport = True
   Me.MousePointer = vbDefault
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
   Me.MousePointer = vbDefault
End Function
