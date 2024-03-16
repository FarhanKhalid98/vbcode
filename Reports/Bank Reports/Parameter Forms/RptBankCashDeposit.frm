VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptBankCashDeposit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptBankCashDeposit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8423
      TabIndex        =   8
      Top             =   6904
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
      MICON           =   "RptBankCashDeposit.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   5648
      TabIndex        =   6
      Top             =   6904
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
      MICON           =   "RptBankCashDeposit.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7028
      TabIndex        =   7
      Top             =   6904
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
      MICON           =   "RptBankCashDeposit.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtVoucherID 
      Height          =   315
      Left            =   5273
      TabIndex        =   0
      Top             =   4001
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   195
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   195
         Left            =   1995
         TabIndex        =   20
         Top             =   0
         Width           =   1020
      End
   End
   Begin JeweledBut.JeweledButton BtnVoucher 
      Height          =   330
      Left            =   6893
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3986
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
      MICON           =   "RptBankCashDeposit.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSlip 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6893
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5801
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "RptBankCashDeposit.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSlipNo 
      Height          =   315
      Left            =   5273
      TabIndex        =   4
      Top             =   5801
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
   Begin JeweledBut.JeweledButton BtnSubGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9728
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5801
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "RptBankCashDeposit.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnBank 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6893
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4961
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "RptBankCashDeposit.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt txtBankID 
      Height          =   315
      Left            =   5273
      TabIndex        =   3
      Top             =   4961
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
      IntegralPoint   =   3
      Mandatory       =   1
   End
   Begin SITextBox.Txt txtBankName 
      Height          =   315
      Left            =   7253
      TabIndex        =   9
      Tag             =   "nc"
      Top             =   4961
      Width           =   2730
      _ExtentX        =   4815
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   315
      Left            =   7253
      TabIndex        =   1
      Top             =   4001
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
      Left            =   8558
      TabIndex        =   2
      Top             =   4001
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
   Begin SITextBox.Txt TxtDepositBy 
      Height          =   315
      Left            =   7253
      TabIndex        =   5
      Top             =   5801
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   40
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   3
      Mandatory       =   1
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit By"
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
      Left            =   7208
      TabIndex        =   24
      Top             =   5606
      Width           =   930
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
      Left            =   8558
      TabIndex        =   23
      Top             =   3806
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
      Left            =   7253
      TabIndex        =   22
      Top             =   3806
      Width           =   885
   End
   Begin VB.Label LblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Cash Deposit"
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
      TabIndex        =   19
      Top             =   270
      Width           =   3015
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slip No"
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
      Left            =   5273
      TabIndex        =   18
      Top             =   5606
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C  Name"
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
      Left            =   7208
      TabIndex        =   15
      Top             =   4751
      Width           =   1440
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C ID"
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
      Left            =   5273
      TabIndex        =   14
      Top             =   4751
      Width           =   1095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   -1410
      TabIndex        =   12
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
      Left            =   5273
      TabIndex        =   11
      Top             =   3806
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
Attribute VB_Name = "RptBankCashDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql As String
Dim VStrSQL As String

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchBankAct.Show vbModal, Me
      If SchBankAct.ParaOutID = "" Then FunSelectAccount = False: Exit Function
      txtBankID.Text = SchBankAct.ParaOutID
   End If
   Dim VStrSQL As String
   VStrSQL = "select * from ChartofAccounts where AccountNo =  '" & Val(txtBankID.Text) & "'"
   With CN.Execute(VStrSQL)
         If .RecordCount > 0 Then
            TxtBankName.Text = !AccountName
            .Close
            FunSelectAccount = True
            Exit Function
         Else
            FunSelectAccount = False
            .Close
            TxtBankName.Text = ""
           End If
      End With
      Exit Function
ErrorHandler:
      Call ShowErrorMessage
End Function

Private Function FunSelectSlip(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSlip.Show vbModal, Me
        If SchSlip.ParaOutID = "" Then FunSelectSlip = False: Exit Function
        TxtSlipNo.Text = SchSlip.ParaOutID
    End If
    '---------------------------
    If Trim(TxtSlipNo.Text) = "" Then Exit Function
    
      TxtSlipNo.Text = TxtSlipNo.Text
    
    If TxtSlipNo.Text = "" Then FunSelectSlip = False: Exit Function
    VStrSQL = " Select SlipNo FROM BankCashDepositBody where SlipNo='" & TxtSlipNo.Text & "'"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         
          FunSelectSlip = True
          .Close
          Exit Function
      Else
          FunSelectSlip = False
          .Close
          TxtSlipNo.Text = ""
'          TxtGroupName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectDepositBy(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchDepositBy.Show vbModal, Me
        If SchDepositBy.ParaOutDepositBy = "" Then FunSelectDepositBy = False: Exit Function
        TxtDepositBy.Text = SchDepositBy.ParaOutDepositBy
    End If
    '---------------------------
    VStrSQL = " Select DepositBy FROM BankCashDepositBody where DepositBy = '" & TxtDepositBy.Text & "'"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          'TxtSubGroupName.Text = !SubGroupName
          FunSelectDepositBy = True
          .Close
          Exit Function
      Else
          FunSelectDepositBy = False
          .Close
          TxtDepositBy.Text = ""
          'TxtSubGroupName.Text = ""
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
      SchCashDeposit.Show vbModal, Me
      If SchCashDeposit.ParaOutVoucherNo = Null Then FunSelectVoucher = False: Exit Function
      TxtVoucherID.Text = SchCashDeposit.ParaOutVoucherNo
   End If
    '---------------------------
    If Trim(TxtVoucherID.Text) = "" Then Exit Function
'    If Len(TxtVoucherID.Text) <= 5 Then
'      TxtVoucherID.Text = Right("00000" + CStr(Val(TxtVoucherID.Text)), 5)
'    End If
    If TxtVoucherID.Text = "" Then FunSelectVoucher = False: Exit Function
    VStrSQL = " SELECT VoucherID, VoucherDate From BankCashDepositHeader Where VoucherID = " & TxtVoucherID.Text
  
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
        ' MsgBox "Invalid VoucherID ID.", vbOKOnly, "Alert"
         TxtVoucherID.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function




Private Sub BtnVoucher_Click()
   If FunSelectVoucher(ssButton, True) = True Then
      DtpFromDate.SetFocus
   Else
      TxtVoucherID.SetFocus
   End If
End Sub

Private Sub BtnSlip_Click()
   If FunSelectSlip(ssButton, False) = True Then
      TxtDepositBy.SetFocus
   Else
      TxtSlipNo.SetFocus
   End If
End Sub

Private Sub btnBank_Click()
   If FunSelectAccount(ssButton, False) = True Then
      TxtSlipNo.SetFocus
   Else
      txtBankID.SetFocus
   End If
End Sub

Private Sub BtnSubGroup_Click()
   If FunSelectDepositBy(ssButton, False) = True Then
      BtnPreview.SetFocus
   Else
      TxtDepositBy.SetFocus
   End If
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

Private Sub txtBankID_Change()
   If txtBankID.Visible = False Then Exit Sub
   If ActiveControl.Name <> txtBankID.Name Then Exit Sub
   If TxtBankName.Text <> "" Then TxtBankName.Text = ""
End Sub

Private Sub txtBankID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> txtBankID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If txtBankID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSlipNo_change()
   If TxtSlipNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSlipNo.Name Then Exit Sub
   
End Sub

Private Sub TxtSlipNo_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtSlipNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtSlipNo.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSlip(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSlip(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDepositBy_Change()
   If TxtDepositBy.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtDepositBy.Name Then Exit Sub
  
End Sub

Private Sub TxtDepositBy_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtDepositBy.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtDepositBy.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectDepositBy(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectDepositBy(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
       RptReportViewer.Caption = "Cash Deposit"
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
         Case txtBankID.Name: If FunSelectAccount(ssFunctionKey, True) = True Then TxtSlipNo.SetFocus
         Case TxtSlipNo.Name: If FunSelectSlip(ssFunctionKey, True) = True Then TxtDepositBy.SetFocus
         Case TxtDepositBy.Name: If FunSelectDepositBy(ssFunctionKey, True) = True Then BtnPreview.SetFocus
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
   SetWindowText Me.hWnd, "Bank Cash Deposit"
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
   Set RptBankCashDeposit = Nothing
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
   Me.MousePointer = vbHourglass
   CrptBankCashDepositParameter.DiscardSavedData
   Set RsReport = CN.Execute("ProdRptBankCashDeposit '" & DtpFromDate.Date & "','" & DtpToDate.Date & "','" & TxtVoucherID.Text & "','" & txtBankID.Text & "','" & TxtSlipNo.Text & "','" & TxtDepositBy.Text & "'")
   Set RptReportViewer.Report = New CrptBankCashDepositParameter
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   If RsReport.BOF Then
      MsgBox "No record exists.", vbInformation, Me.Caption
      Me.MousePointer = vbDefault
      Exit Function
   End If
   'RptReportViewer.Report.Database.SetDataSource RsReport
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
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
