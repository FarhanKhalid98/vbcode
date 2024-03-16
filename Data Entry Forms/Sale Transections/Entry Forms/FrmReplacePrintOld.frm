VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmReplacePrintOld 
   BorderStyle     =   0  'None
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCredit 
      BorderStyle     =   0  'None
      Height          =   1665
      Left            =   810
      TabIndex        =   13
      Top             =   3510
      Width           =   6105
      Begin VB.TextBox TxtNetAmountCredit 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   2025
      End
      Begin SITextBox.Txt TxtCustomerID 
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   930
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Appearance      =   0
         MaxLength       =   11
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
         IntegralPoint   =   10
         Mandatory       =   1
      End
      Begin SITextBox.Txt TxtCustomerName 
         Height          =   315
         Left            =   1755
         TabIndex        =   16
         Top             =   930
         Width           =   4350
         _ExtentX        =   7673
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
         Masked          =   5
      End
      Begin JeweledBut.JeweledButton BtnCustomer 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   1395
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   930
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
         MICON           =   "FrmReplacePrintOld.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   195
         Left            =   1755
         TabIndex        =   18
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Difference"
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
         Left            =   1020
         TabIndex        =   15
         Top             =   210
         Width           =   900
      End
   End
   Begin VB.Frame FrameCash 
      BorderStyle     =   0  'None
      Height          =   1950
      Left            =   1260
      TabIndex        =   8
      Top             =   900
      Width           =   4425
      Begin VB.TextBox TxtNetAmount 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   2025
      End
      Begin SITextBox.Txt TxtCashCustomer 
         Height          =   315
         Left            =   30
         TabIndex        =   2
         Top             =   360
         Width           =   4350
         _ExtentX        =   7673
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
         Masked          =   5
      End
      Begin SITextBox.Txt TxtCashPaid 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   1365
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
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
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   30
         TabIndex        =   20
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Difference"
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
         Left            =   990
         TabIndex        =   12
         Top             =   990
         Width           =   900
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Paid"
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
         Left            =   990
         TabIndex        =   11
         Top             =   1395
         Width           =   870
      End
   End
   Begin VB.CheckBox ChkPrint 
      Caption         =   "&Print"
      Height          =   285
      Left            =   780
      TabIndex        =   10
      Top             =   270
      Value           =   1  'Checked
      Width           =   705
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   2205
      TabIndex        =   7
      Top             =   90
      Width           =   2175
      Begin VB.OptionButton OptCredit 
         Caption         =   "&Credit"
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton OptCash 
         Caption         =   "&Cash"
         Height          =   285
         Left            =   210
         TabIndex        =   0
         Top             =   240
         Width           =   765
      End
   End
   Begin JeweledBut.JeweledButton BtnCancel 
      Height          =   420
      Left            =   3630
      TabIndex        =   6
      Top             =   3045
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Cancel"
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
      MICON           =   "FrmReplacePrintOld.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOk 
      Height          =   420
      Left            =   2325
      TabIndex        =   5
      Top             =   3045
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "OK"
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
      MICON           =   "FrmReplacePrintOld.frx":0038
      BC              =   14737632
      FC              =   0
   End
End
Attribute VB_Name = "FrmReplacePrintOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ParaOutSelection As Boolean
Public ParaInPrint As Boolean

Public Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is SITextBox.txt Then
         ctl.Text = ""
      ElseIf TypeOf ctl Is TextBox Then
         ctl.Text = ""
      End If
   Next
   OptCash.Value = True
   TxtCashCustomer.Text = "Counter Sale"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   FunValidation = False
   If OptCredit.Value = True Then
      If Trim(TxtCustomerID.Text) = "" Then
         MsgBox "Please specify a Customer ID", vbExclamation, "Alert"
         TxtCustomerID.SetFocus
         Exit Function
      End If
   End If
   If OptCash.Value = True Then
      If Val(TxtCashPaid.Text) <> Val(TxtNetAmount.Text) Then
         MsgBox LblCaption.Caption & " not equal to Net Amount", vbExclamation, "Alert"
         TxtCashPaid.SetFocus
         Exit Function
      End If
   End If
   FunValidation = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
  If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select * FROM ChartofAccounts where AccountNo = '" & (TxtCustomerID.Text) & "' and (AccountNo like '6%' or AccountNo like '5%' or AccountNo Like '3%') and isDetailed = 1 and isLocked = 0"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerName.Text = !AccountName
          FunSelectCustomer = True
          .Close
          Exit Function
      Else
          FunSelectCustomer = False
          .Close
          TxtCustomerID.Text = ""
          TxtCustomerName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnCancel_Click()
   On Error GoTo ErrorHandler
   ParaOutSelection = False
   Me.Hide
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      BtnOk.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

Private Sub BtnOk_Click()
   On Error GoTo ErrorHandler
   If FunValidation = False Then Exit Sub
   ParaOutSelection = True
   Me.Hide
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkPrint_Click()
   If ChkPrint.Value = 1 Then
      ParaInPrint = True
   Else
      ParaInPrint = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
   If Shift = vbCtrlMask And KeyCode = vbKeyReturn Then
      BtnOk_Click
   ElseIf KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyEscape Then
      BtnCancel_Click
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnOk.Enabled Then BtnOk_Click
            KeyCode = 0
         Case vbKeyW
            If BtnCancel.Enabled Then BtnCancel_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCustomerID.Name: If FunSelectCustomer(ssFunctionKey, False) = True Then BtnOk.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   FrameCash.Top = 900
   FrameCredit.Top = 1080
   ChkPrint.Value = Abs(ParaInPrint)
   OptCash.Value = True
   Call OptCash_Click
   TxtCashCustomer.Text = "Counter Sale"
End Sub

Private Sub OptCash_Click()
   FrameCash.Visible = True
   FrameCredit.Visible = False
End Sub

Private Sub OptCredit_Click()
   FrameCash.Visible = False
   FrameCredit.Visible = True
End Sub

Private Sub TxtNetAmount_Change()
   TxtNetAmountCredit.Text = TxtNetAmount.Text
   TxtCashPaid.Text = TxtNetAmount.Text
End Sub

Private Sub TxtCustomerID_Change()
   If TxtCustomerID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   If TxtCustomerName.Text <> "" Then
      TxtCustomerName.Text = ""
   End If
End Sub

Private Sub TxtCustomerID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerName.Text <> "" Then Exit Sub
   If Trim(TxtCustomerID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCustomer(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCustomer(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
