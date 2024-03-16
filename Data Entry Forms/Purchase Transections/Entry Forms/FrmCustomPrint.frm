VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmCustomPrint 
   BorderStyle     =   0  'None
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkPrint 
      Caption         =   "&Print"
      Height          =   285
      Left            =   675
      TabIndex        =   26
      Top             =   270
      Value           =   1  'Checked
      Width           =   705
   End
   Begin VB.Frame FrameBank 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   510
      TabIndex        =   15
      Top             =   3480
      Width           =   6195
      Begin VB.TextBox TxtCommision 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4725
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   855
         Visible         =   0   'False
         Width           =   900
      End
      Begin SITextBox.Txt TxtBankMachineID 
         Height          =   315
         Left            =   60
         TabIndex        =   17
         Top             =   1425
         Width           =   1335
         _ExtentX        =   2355
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
         Mandatory       =   1
      End
      Begin SITextBox.Txt TxtBankMachineName 
         Height          =   315
         Left            =   1755
         TabIndex        =   18
         Top             =   1425
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   556
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Masked          =   5
      End
      Begin JeweledBut.JeweledButton BtnBankMachine 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   1395
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1425
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
         MICON           =   "FrmCustomPrint.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtInvoiceNo 
         Height          =   315
         Left            =   1560
         TabIndex        =   20
         Top             =   840
         Width           =   2025
         _ExtentX        =   3572
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
      End
      Begin SITextBox.Txt TxtBankVender 
         Height          =   315
         Left            =   675
         TabIndex        =   21
         Top             =   285
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   556
         Appearance      =   0
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Masked          =   5
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Machine ID"
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   1215
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Machine Name"
         Height          =   195
         Left            =   1755
         TabIndex        =   24
         Top             =   1215
         Width           =   1500
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   870
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   675
         TabIndex        =   22
         Top             =   45
         Width           =   1665
      End
   End
   Begin VB.Frame FrameCash 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1290
      TabIndex        =   12
      Top             =   750
      Width           =   4425
      Begin SITextBox.Txt TxtCashVender 
         Height          =   315
         Left            =   30
         TabIndex        =   13
         Top             =   360
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   556
         Appearance      =   0
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Masked          =   5
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   30
         TabIndex        =   14
         Top             =   120
         Width           =   1305
      End
   End
   Begin VB.Frame FrameCredit 
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   720
      TabIndex        =   7
      Top             =   780
      Width           =   6105
      Begin SITextBox.Txt TxtVenderID 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
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
         Mandatory       =   1
      End
      Begin SITextBox.Txt TxtVenderName 
         Height          =   315
         Left            =   1755
         TabIndex        =   8
         Top             =   330
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   556
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Masked          =   5
      End
      Begin JeweledBut.JeweledButton BtnVender 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   1395
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   330
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
         MICON           =   "FrmCustomPrint.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Vender ID"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   195
         Left            =   1755
         TabIndex        =   10
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   1680
      TabIndex        =   6
      Top             =   90
      Width           =   3525
      Begin VB.OptionButton OptCredit 
         Caption         =   "&Credit"
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton OptBankCard 
         Caption         =   "&Bank Card"
         Height          =   285
         Left            =   2100
         TabIndex        =   2
         Top             =   240
         Width           =   1125
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
      Left            =   3510
      TabIndex        =   5
      Top             =   2910
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Cancel"
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
      MICON           =   "FrmCustomPrint.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOk 
      Height          =   420
      Left            =   2205
      TabIndex        =   4
      Top             =   2910
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "OK"
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
      MICON           =   "FrmCustomPrint.frx":0054
      BC              =   14737632
      FC              =   0
   End
End
Attribute VB_Name = "FrmCustomPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ParaOutSelection As Boolean
Public ParaInChoice As String

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
   TxtCashVender.Text = "Counter Sale"
   TxtBankVender.Text = "Counter Sale"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   FunValidation = False
   If OptBankCard.Value = True Then
      If Trim(TxtBankMachineID.Text) = "" Then
         MsgBox "Please specify a Bank Machine ID", vbExclamation, "Alert"
         TxtBankMachineID.SetFocus
         Exit Function
      End If
   End If
   If OptCredit.Value = True Then
      If Trim(TxtVenderID.Text) = "" Then
         MsgBox "Please specify a Vender ID", vbExclamation, "Alert"
         TxtVenderID.SetFocus
         Exit Function
      End If
   End If
   FunValidation = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectBankMachine(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchBankMachine.Show vbModal, Me
        If SchBankMachine.ParaOutBankMachineID = "" Then FunSelectBankMachine = False: Exit Function
        TxtBankMachineID.Text = SchBankMachine.ParaOutBankMachineID
    End If
    '---------------------------
    vStrSQL = " Select * FROM BankMachines where BankMachineID=" & Val(TxtBankMachineID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtBankMachineName.Text = !BankMachineName
          TxtCommision.Text = !Commision
          FunSelectBankMachine = True
          .Close
          Exit Function
      Else
          FunSelectBankMachine = False
          .Close
          TxtBankMachineID.Text = ""
          TxtBankMachineName.Text = ""
          TxtCommision.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectVender(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInWhereClause = " and AccountNo like '6%' and isDetailed = 1 and isLocked = 0"
        SchAccounts.ParaInDetail = ""
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectVender = False: Exit Function
        TxtVenderID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = "Select * FROM ChartofAccounts where AccountNo = " & Val(TxtVenderID.Text) & " and AccountNo like '6%' and isDetailed = 1 and isLocked = 0"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtVenderName.Text = !PartyName
          FunSelectVender = True
          .Close
          Exit Function
      Else
          FunSelectVender = False
          .Close
          TxtVenderID.Text = ""
          TxtVenderName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnBankMachine_Click()
   If FunSelectBankMachine(ssButton, False) = True Then
      BtnOk.SetFocus
   Else
      TxtBankMachineID.SetFocus
   End If
End Sub

Private Sub BtnCancel_Click()
   On Error GoTo ErrorHandler
   ParaOutSelection = False
   Me.Hide
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
      BtnOk.SetFocus
   Else
      TxtVenderID.SetFocus
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
         Case TxtBankMachineID.Name: If FunSelectBankMachine(ssFunctionKey, True) = True Then BtnOk.SetFocus
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, False) = True Then BtnOk.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   FrameBank.Top = 765
   FrameCash.Top = 765
   FrameCredit.Top = 765
   If ParaInChoice = "Cash" Or ParaInChoice = "" Then
      OptCash.Value = True
      Call OptCash_Click
   ElseIf ParaInChoice = "Credit" Then
      OptCredit.Value = True
      Call OptCredit_Click
   ElseIf ParaInChoice = "Bank" Then
      OptBankCard.Value = True
      Call OptBankCard_Click
   End If
   TxtCashVender.Text = "Counter Sale"
   TxtBankVender.Text = "Counter Sale"
End Sub

Private Sub OptBankCard_Click()
   FrameCash.Visible = False
   FrameCredit.Visible = False
   FrameBank.Visible = True
   TxtBankVender.Text = IIf(TxtBankVender.Text = "", "Counter Sale", TxtBankVender.Text)
   If Trim(TxtBankMachineID.Text) <> "" Then Exit Sub
   With cn.Execute("select * from registry")
      If .RecordCount > 0 Then
         If Not IsNull(!BankMachineID) Then
            TxtBankMachineID.Text = IIf(IsNull(!BankMachineID), "", !BankMachineID)
            FunSelectBankMachine ssValidate, True
         End If
      End If
      .Close
   End With
End Sub

Private Sub OptCash_Click()
   FrameCash.Visible = True
   FrameCredit.Visible = False
   FrameBank.Visible = False
   TxtCashVender.Text = IIf(TxtCashVender.Text = "", "Counter Sale", TxtCashVender.Text)
End Sub

Private Sub OptCredit_Click()
   FrameCash.Visible = False
   FrameCredit.Visible = True
   FrameBank.Visible = False
End Sub

Private Sub TxtVenderID_Change()
   If TxtVenderID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   If TxtVenderName.Text <> "" Then TxtVenderName.Text = ""
End Sub

Private Sub TxtVenderID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtVenderName.Text <> "" Then Exit Sub
   If Trim(TxtVenderID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVender(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVender(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtBankMachineID_Change()
   If TxtBankMachineID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtBankMachineID.Name Then Exit Sub
   If TxtBankMachineName.Text <> "" Then
      TxtBankMachineName.Text = ""
      TxtCommision.Text = ""
   End If
End Sub

Private Sub TxtBankMachineID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtBankMachineID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtBankMachineName.Text <> "" Then Exit Sub
   If Trim(TxtBankMachineID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBankMachine(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBankMachine(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
