VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmChequeIssue 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmChequeIssue.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7343
      TabIndex        =   12
      Top             =   7170
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
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
      MICON           =   "FrmChequeIssue.frx":6971
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6023
      TabIndex        =   9
      Top             =   7170
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
      MICON           =   "FrmChequeIssue.frx":698D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3383
      TabIndex        =   11
      Top             =   7170
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
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
      MICON           =   "FrmChequeIssue.frx":69A9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8663
      TabIndex        =   13
      Top             =   7170
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
      MICON           =   "FrmChequeIssue.frx":69C5
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4703
      TabIndex        =   10
      Top             =   7170
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
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
      MICON           =   "FrmChequeIssue.frx":69E1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2063
      TabIndex        =   14
      Top             =   7170
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
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
      MICON           =   "FrmChequeIssue.frx":69FD
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   6885
      TabIndex        =   5
      Top             =   3330
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   7
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpVoucherDate 
      Height          =   315
      Left            =   4943
      TabIndex        =   1
      Top             =   1635
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
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
   Begin JeweledBut.JeweledButton BtnBankAccount 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4305
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2490
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmChequeIssue.frx":6A19
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBankAccountName 
      Height          =   315
      Left            =   4665
      TabIndex        =   24
      Top             =   2490
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   50
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
   Begin SITextBox.Txt TxtChequeNo 
      Height          =   315
      Left            =   4905
      TabIndex        =   4
      Top             =   3330
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   15
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
      IntegralPoint   =   15
   End
   Begin JeweledBut.JeweledButton BtnAccount 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4305
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4185
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmChequeIssue.frx":6A35
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtAccountName 
      Height          =   315
      Left            =   4665
      TabIndex        =   30
      Top             =   4185
      Width           =   3825
      _ExtentX        =   6747
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
   End
   Begin SITextBox.Txt TxtVoucherNo 
      Height          =   315
      Left            =   3285
      TabIndex        =   0
      Top             =   1650
      Width           =   1260
      _ExtentX        =   2223
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
   End
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   3285
      TabIndex        =   8
      Top             =   5865
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   50
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
   Begin SITextBox.Txt TxtReceivedBy 
      Height          =   315
      Left            =   3285
      TabIndex        =   7
      Top             =   5025
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtBankAccountNo 
      Height          =   315
      Left            =   3285
      TabIndex        =   2
      Top             =   2490
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpChequeDate 
      Height          =   315
      Left            =   3285
      TabIndex        =   3
      Top             =   3330
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
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
   Begin SITextBox.Txt TxtAccountNo 
      Height          =   315
      Left            =   3285
      TabIndex        =   6
      Top             =   4185
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name"
      Height          =   195
      Left            =   4680
      TabIndex        =   29
      Top             =   3975
      Width           =   1065
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay To"
      Height          =   195
      Left            =   3285
      TabIndex        =   28
      Top             =   3960
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Date"
      Height          =   195
      Left            =   3285
      TabIndex        =   26
      Top             =   3120
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No"
      Height          =   195
      Left            =   4905
      TabIndex        =   25
      Top             =   3120
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account Name"
      Height          =   195
      Left            =   4665
      TabIndex        =   23
      Top             =   2265
      Width           =   1485
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/c No."
      Height          =   195
      Left            =   3285
      TabIndex        =   22
      Top             =   2265
      Width           =   990
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Issue"
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
      Left            =   2070
      TabIndex        =   20
      Top             =   135
      Width           =   3300
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No"
      Height          =   195
      Left            =   3285
      TabIndex        =   19
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   3285
      TabIndex        =   18
      Top             =   5640
      Width           =   795
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Received By"
      Height          =   195
      Left            =   3285
      TabIndex        =   17
      Top             =   4800
      Width           =   915
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   6885
      TabIndex        =   16
      Top             =   3120
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Date"
      Height          =   195
      Left            =   4943
      TabIndex        =   15
      Top             =   1410
      Width           =   990
   End
End
Attribute VB_Name = "FrmChequeIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsReport As New ADODB.Recordset
Dim vStrSQL As String
Dim vIsNewRecord As Boolean
Dim vMode As FormMode

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
        TxtAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = "Select * FROM ChartofAccounts where AccountNo='" & TxtAccountNo.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtAccountName.Text = !AccountName
          FunSelectAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectAccount = False
          .Close
          TxtAccountNo.Text = ""
          TxtAccountName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectBankAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectBankAccount = False: Exit Function
        TxtBankAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = "Select * FROM ChartofAccounts where AccountNo='" & TxtBankAccountNo.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtBankAccountName.Text = !AccountName
          FunSelectBankAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectBankAccount = False
          .Close
          TxtBankAccountNo.Text = ""
          TxtBankAccountName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = CN.Execute("Select isnull(max(VoucherNo),0)+1 from BankCheques").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunValidateHeaderInfo() As Boolean
   On Error GoTo ErrorHandler
   FunValidateHeaderInfo = False
   If vIsNewRecord = True Then
   Else
      If ObjUserSecurity.IsEdit = False Then
         MsgBox "You are not authorized to modify a record", vbCritical, "Error"
         Exit Function
      End If
   End If
   If Trim(TxtBankAccountNo.Text) = "" Then
       MsgBox "Please enter the Bank AccountNo", vbExclamation + vbApplicationModal + vbOKOnly, "Alert"
       If TxtBankAccountNo.Enabled And TxtBankAccountNo.Visible Then TxtBankAccountNo.SetFocus
       Exit Function
   End If
   If Trim(TxtChequeNo.Text) = "" Then
       MsgBox "Please enter the Cheque No", vbExclamation + vbApplicationModal + vbOKOnly, "Alert"
       If TxtChequeNo.Enabled And TxtChequeNo.Visible Then TxtChequeNo.SetFocus
       Exit Function
   End If
   If Val(TxtAmount.Text) = 0 Then
       MsgBox "Please enter the Amount", vbExclamation + vbApplicationModal + vbOKOnly, "Alert"
       If TxtAmount.Enabled And TxtAmount.Visible Then TxtAmount.SetFocus
       Exit Function
   End If
   If Trim(TxtAccountNo.Text) = "" Then
       MsgBox "Please enter the Pay To Account", vbExclamation + vbApplicationModal + vbOKOnly, "Alert"
       If TxtAccountNo.Enabled And TxtAccountNo.Visible Then TxtAccountNo.SetFocus
       Exit Function
   End If
   If Trim(TxtReceivedBy.Text) = "" Then
       MsgBox "Please enter the Received By", vbExclamation + vbApplicationModal + vbOKOnly, "Alert"
       If TxtReceivedBy.Enabled And TxtReceivedBy.Visible Then TxtReceivedBy.SetFocus
       Exit Function
   End If
   FunValidateHeaderInfo = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub GetOpen()
   On Error GoTo ErrorHandler
   vStrSQL = " SELECT * from BankCheques" & vbCrLf _
            + " WHERE VoucherNo = " & Val(TxtVoucherNo.Text)
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         DtpVoucherDate.DateValue = !VoucherDate
         TxtBankAccountNo.Text = !BankAccountNo
         DtpChequeDate.DateValue = !ChequeDate
         TxtChequeNo.Text = !ChequeNo
         TxtAmount.Text = !Amount
         TxtAccountNo.Text = !AccountNo
         TxtReceivedBy.Text = IIf(IsNull(!ReceivedBy), "", !ReceivedBy)
         TxtDescription.Text = IIf(IsNull(!Narration), "", !Narration)
      End If
      .Close
   End With
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

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
   vIsNewRecord = True
   Set ctl = Nothing
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
      Call SubClearFields
      BtnPrint.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtVoucherNo.Text = FunGetMaxID
      DtpVoucherDate.DateValue = Date
      If DtpVoucherDate.Enabled And DtpVoucherDate.Visible Then DtpVoucherDate.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      BtnPrint.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      DtpVoucherDate.SetFocus
      vIsNewRecord = False
    Case Is = ChangeMode
      BtnDelete.Enabled = False
      BtnPrint.Enabled = False
      BtnSave.Enabled = True
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub BtnAccount_Click()
   If FunSelectAccount(ssButton, False) = True Then
      TxtReceivedBy.SetFocus
   Else
      TxtAccountNo.SetFocus
   End If
End Sub

Private Sub BtnBankAccount_Click()
   If FunSelectBankAccount(ssButton, False) = True Then
      DtpChequeDate.SetFocus
   Else
      TxtBankAccountNo.SetFocus
   End If
End Sub

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

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   If ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Are you sure to remove this VocuherNo?", vbApplicationModal + vbYesNo + vbQuestion, "Alert") = vbNo Then Exit Sub
   CN.BeginTrans
      CN.Execute "Delete From BankCheques where VoucherNo=" & TxtVoucherNo.Text
   CN.CommitTrans
   FormStatus = NewMode
   TxtVoucherNo.SetFocus
Exit Sub
ErrorHandler:
   CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchChequeIssue.Show vbModal, Me
   If ParaOutID <> "" Then
      TxtVoucherNo.Text = Val(ParaOutID)
      GetOpen
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub BtnPrint_Click()
On Error GoTo ErrorHandler
'    Set RsReport = CN.Execute("exec SPEasyLoad " & TxtTransectionID.Text)
'    Set RptReportViewer.ReportToDisplay = New CrpEasyLoad
'    RptReportViewer.ReportToDisplay.Database.SetDataSource RsReport, 3, 1
'    RptReportViewer.ReportToDisplay.ParameterFields(1).AddCurrentValue CN.Execute("Select companyName from Company").Fields(0).Value
'    RptReportViewer.ReportToDisplay.ParameterFields(2).AddCurrentValue CN.Execute("Select dbo.FunGetAddress(" & ObjUserSecurity.UserNo & ")").Fields(0).Value
'    RptReportViewer.ReportToDisplay.SelectPrinter "RASDD.DLL", "CBM1000 Partial Cut", "Com1"
'    RptReportViewer.ReportToDisplay.TopMargin = 0
'    RptReportViewer.ReportToDisplay.PaperOrientation = crPortrait
'    'RptReportViewer.Show vbModal
'    RptReportViewer.ReportToDisplay.PrintOut False
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If FunValidateHeaderInfo = False Then Exit Sub
   Dim Rs As New ADODB.Recordset
   CN.BeginTrans
   '--- Saving Informations ---
      If Rs.State = adStateOpen Then Rs.Close
      Rs.Open "Select * From BankCheques Where VoucherNo=" & TxtVoucherNo.Text, CN, adOpenStatic, adLockPessimistic
      If Rs.RecordCount = 0 Then
          Rs.AddNew
          Rs!VoucherNo = TxtVoucherNo.Text
      End If
      Rs!VoucherDate = DtpVoucherDate.DateValue
      Rs!BankAccountNo = TxtBankAccountNo.Text
      Rs!ChequeDate = DtpChequeDate.DateValue
      Rs!ChequeNo = TxtChequeNo.Text
      Rs!Amount = TxtAmount.Text
      Rs!AccountNo = TxtAccountNo.Text
      Rs!ReceivedBy = IIf(Trim(TxtReceivedBy.Text) = "", Null, TxtReceivedBy.Text)
      Rs!Narration = IIf(Trim(TxtDescription.Text) = "", Null, TxtDescription.Text)
      Rs!UserNo = ObjUserSecurity.UserNo
      Rs.Update
   CN.CommitTrans
   If MsgBox("Are you want to print this Transection?", vbApplicationModal + vbYesNo + vbQuestion, "Alert") = vbYes Then
      BtnPrint_Click
   End If
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   CN.RollbackTrans
   Call ShowErrorMessage
   If Rs.State = adStateOpen Then
      If Rs.BOF = False And Rs.EOF = False Then
         If Rs.EditMode = adEditAdd Or Rs.EditMode = adEditInProgress Then
            Rs.CancelUpdate
         End If
      End If
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSave.Enabled Then BtnSave_Click
            KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
         Case vbKeyO
            If BtnOpen.Enabled Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtBankAccountNo.Name: If FunSelectBankAccount(ssFunctionKey, True) = True Then DtpChequeDate.SetFocus Else TxtBankAccountNo.SetFocus
         Case TxtAccountNo.Name: If FunSelectAccount(ssFunctionKey, False) = True Then TxtReceivedBy.SetFocus Else TxtAccountNo.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    FormStatus = NewMode
    Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Are you sure to close without save?", vbQuestion + vbApplicationModal + vbYesNo, "Alert") = vbNo Then
         Cancel = 1
      End If
   Else
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set FrmChequeIssue = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtAccountNo_Change()
   If ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   If Trim(TxtAccountName.Text) <> "" Then TxtAccountName.Text = ""
End Sub

Private Sub TxtAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtAccountName.Text) <> "" Then Exit Sub
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

Private Sub TxtBankAccountNo_Change()
   If ActiveControl.Name <> TxtBankAccountNo.Name Then Exit Sub
   If Trim(TxtBankAccountName.Text) <> "" Then TxtBankAccountName.Text = ""
End Sub

Private Sub TxtBankAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtBankAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtBankAccountName.Text) <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBankAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBankAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
