VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmRegistrySettings 
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
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkHidePurchaseAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hide Amount in Purchase Transections for Standard User"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   11
      Top             =   6300
      Width           =   4470
   End
   Begin VB.CheckBox ChkHideSaleAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hide Amount in Previous Sale Transections For Standard User"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   10
      Top             =   5775
      Width           =   4740
   End
   Begin VB.CheckBox ChkMemberVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Member in Sale Transections"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   9
      Top             =   5235
      Width           =   2850
   End
   Begin VB.CheckBox ChkEmployeeVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Employee in Sale Transections"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   8
      Top             =   4710
      Width           =   2985
   End
   Begin VB.CheckBox ChkChangePrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Allow Retail Price is Changable in Sale Invoice By Administrator"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   7
      Top             =   4185
      Width           =   4875
   End
   Begin VB.CheckBox ChkCostVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Cost in Sale Invoice"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   6
      Top             =   3645
      Width           =   2355
   End
   Begin VB.CheckBox ChkCashReceived 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Amount is equal to Cash Received By Default"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   5
      Top             =   3120
      Width           =   4065
   End
   Begin VB.CheckBox ChkNegativeSale 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Allow Negative Sales"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   3
      Top             =   2595
      Width           =   1905
   End
   Begin VB.CheckBox ChkAddSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Extra Space at the end of Sale Invoice"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   2
      Top             =   2055
      Width           =   3795
   End
   Begin VB.CheckBox ChkStoreVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Stores in Invoices"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   0
      Top             =   1530
      Width           =   2085
   End
   Begin VB.TextBox TxtStatement 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   810
      MaxLength       =   100
      TabIndex        =   12
      Top             =   7230
      Width           =   10515
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   4688
      TabIndex        =   13
      Top             =   7995
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
      MICON           =   "FrmRegistrySettings.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6038
      TabIndex        =   14
      Top             =   7995
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
      MICON           =   "FrmRegistrySettings.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   7245
      TabIndex        =   1
      Tag             =   "NC"
      Top             =   3375
      Width           =   795
      _ExtentX        =   1402
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
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   8400
      TabIndex        =   17
      Tag             =   "NC"
      Top             =   3375
      Width           =   2415
      _ExtentX        =   4260
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8040
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3375
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
      MICON           =   "FrmRegistrySettings.frx":0038
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBankMachineID 
      Height          =   315
      Left            =   7245
      TabIndex        =   4
      Top             =   4035
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   8400
      TabIndex        =   20
      Top             =   4035
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   8040
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4035
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
      MICON           =   "FrmRegistrySettings.frx":0054
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Machine Name"
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
      Left            =   9000
      TabIndex        =   25
      Top             =   3825
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
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
      Left            =   8415
      TabIndex        =   24
      Top             =   3150
      Width           =   1005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   465
      X2              =   735
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2400
      Left            =   6975
      Top             =   2295
      Width           =   4065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Box Setting"
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
      Left            =   7020
      TabIndex        =   23
      Top             =   2475
      Width           =   3960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Machine ID"
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
      Left            =   7245
      TabIndex        =   22
      Top             =   3825
      Width           =   1485
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
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
      Left            =   7245
      TabIndex        =   19
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Settings"
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
      Left            =   1980
      TabIndex        =   16
      Top             =   195
      Width           =   1890
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Footer Statement"
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
      Left            =   810
      TabIndex        =   15
      Top             =   7005
      Width           =   1785
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   11625
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmRegistrySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkStoreVisible_Click()
   TxtStoreID.Enabled = ChkStoreVisible.Value = 1
   TxtStoreName.Enabled = TxtStoreID.Enabled
   BtnStore.Enabled = TxtStoreID.Enabled
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
'         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtEmployeeID.SetFocus
'         Case TxtbankID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then If TxtMemberID.Visible = True Then If TxtCode.Enabled Then TxtCode.SetFocus
      End Select
   End If
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  If FunValidation = False Then Exit Sub
  CN.Execute ("UPDATE Registry Set StoreVisible = " & ChkStoreVisible.Value _
  & ", StoreID = " & IIf(Trim(TxtStoreID.Text) = "", "Null", Val(TxtStoreID.Text)) _
  & ", BankMachineID = " & IIf(Trim(TxtBankMachineID.Text) = "", "Null", Val(TxtBankMachineID.Text)) _
  & ", AddSpace = " & ChkAddSpace.Value _
  & ", NegativeSale = " & ChkNegativeSale.Value _
  & ", CashReceived = " & ChkCashReceived.Value _
  & ", CostVisible = " & ChkCostVisible.Value _
  & ", ChangePrice = " & ChkChangePrice.Value _
  & ", EmpVisible = " & ChkEmployeeVisible.Value _
  & ", MemberVisible = " & ChkMemberVisible.Value _
  & ", HideSaleAmount = " & ChkHideSaleAmount.Value _
  & ", HidePurchaseAmount = " & ChkHidePurchaseAmount.Value _
  & ", Statement = '" & TxtStatement.Text & "'")
  MsgBox "Your Registry Settings has been Changed successfully", vbInformation, "Information"
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
'  If Trim(TxtName.Text) = "" Then
'    MsgBox "Please specify a Company Name", vbExclamation, "Alert"
'    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
'    Exit Function
'  End If
  'All Ok, now validation is success
  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Registry Settings"
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   With CN.Execute("Select r.*, StoreName, BankMachineName FROM Registry r left outer join Stores s on r.storeid = s.storeid left outer join BankMachines b on b.BankMachineID = r.BankMachineID")
      ChkStoreVisible.Value = Abs(!StoreVisible)
      TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
      TxtStoreName.Text = IIf(IsNull(!StoreName), "", !StoreName)
      TxtBankMachineID.Text = IIf(IsNull(!BankMachineID), "", !BankMachineID)
      TxtBankMachineName.Text = IIf(IsNull(!BankMachineName), "", !BankMachineName)
      ChkAddSpace.Value = Abs(!AddSpace)
      ChkNegativeSale.Value = Abs(!NegativeSale)
      ChkCashReceived.Value = Abs(!CashReceived)
      ChkCostVisible.Value = Abs(!CostVisible)
      ChkChangePrice.Value = Abs(!ChangePrice)
      ChkEmployeeVisible.Value = Abs(!EmpVisible)
      ChkMemberVisible.Value = Abs(!MemberVisible)
      ChkHideSaleAmount.Value = Abs(!HideSaleAmount)
      ChkHidePurchaseAmount.Value = Abs(!HidePurchaseAmount)
      TxtStatement.Text = !Statement
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub
