VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form DefUsers 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtStoreID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9270
      MaxLength       =   10
      TabIndex        =   31
      Top             =   1995
      Width           =   675
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EFC09E&
      Caption         =   "Show Price Wtih Short Key"
      Height          =   2280
      Left            =   11115
      TabIndex        =   26
      Top             =   5745
      Width           =   3375
      Begin VB.CheckBox ChkWeightedPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "F6:- Weighted Price"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   30
         Top             =   1035
         Width           =   2445
      End
      Begin VB.CheckBox ChkWSPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "F7:- Whole Sale Price"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   29
         Top             =   1350
         Width           =   2805
      End
      Begin VB.CheckBox ChkShowPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "F4:- Procuct Price"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   28
         Top             =   360
         Width           =   2805
      End
      Begin VB.CheckBox ChkLastPurchasePrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "F5:- Last Purchase Price"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   27
         Top             =   720
         Width           =   2385
      End
   End
   Begin VB.CheckBox ChkReadOnly 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Read Only"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10305
      TabIndex        =   24
      Top             =   8145
      Width           =   1320
   End
   Begin VB.CheckBox ChkLockUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Lock User"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10305
      TabIndex        =   9
      Top             =   8400
      Width           =   1320
   End
   Begin VB.Frame FraFacility 
      BackColor       =   &H00EFC09E&
      Caption         =   "Facility"
      Height          =   2280
      Left            =   7530
      TabIndex        =   21
      Top             =   5745
      Width           =   3375
      Begin VB.CheckBox ChkEditClosingInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Edit Closing Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   36
         Top             =   1980
         Width           =   2385
      End
      Begin VB.CheckBox ChkEditDefination 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Edit Defination"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   25
         Top             =   720
         Width           =   2385
      End
      Begin VB.CheckBox ChkChangePriceSaleInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Change Retail Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   1665
         Width           =   2805
      End
      Begin VB.CheckBox ChkChangeRetail 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Change Retail Sale POS"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   1350
         Width           =   2445
      End
      Begin VB.CheckBox ChkDelete 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Remove"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   1035
         Width           =   1320
      End
      Begin VB.CheckBox ChkEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Edit"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   6
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EFC09E&
      Caption         =   "Type of User"
      Height          =   1155
      Left            =   9255
      TabIndex        =   20
      Top             =   4500
      Width           =   3375
      Begin VB.OptionButton OptManager 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Manager"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   4
         ToolTipText     =   "Standard users are restricted to perform only those tasks which are explicitly assigned to them by the System administrator."
         Top             =   570
         Width           =   1320
      End
      Begin VB.OptionButton OptAdmin 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Administrator"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   5
         ToolTipText     =   $"DefUsers.frx":0000
         Top             =   825
         Width           =   1530
      End
      Begin VB.OptionButton OptNormal 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Standard User"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   3
         ToolTipText     =   "Standard users are restricted to perform only those tasks which are explicitly assigned to them by the System administrator."
         Top             =   315
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.TextBox TxtConfirmPassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   9285
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3930
      Width           =   3360
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   9285
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3345
      Width           =   3360
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4320
      Left            =   2580
      TabIndex        =   13
      Top             =   2730
      Width           =   4425
      ScrollBars      =   2
      _Version        =   196616
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
      stylesets(0).Picture=   "DefUsers.frx":0089
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
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "User No"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6720
      Columns(1).Caption=   "User Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   7805
      _ExtentY        =   7620
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
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9285
      MaxLength       =   30
      TabIndex        =   0
      Top             =   2715
      Width           =   3360
   End
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   2685
      TabIndex        =   14
      Top             =   9015
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "New"
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
      MICON           =   "DefUsers.frx":00A5
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   4005
      TabIndex        =   15
      Top             =   9015
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Edit"
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
      MICON           =   "DefUsers.frx":00C1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   5310
      TabIndex        =   16
      Top             =   9015
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
      MICON           =   "DefUsers.frx":00DD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8895
      TabIndex        =   10
      Top             =   9030
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
      MICON           =   "DefUsers.frx":00F9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   10215
      TabIndex        =   11
      Top             =   9030
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
      MICON           =   "DefUsers.frx":0115
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   11505
      TabIndex        =   12
      Top             =   9060
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
      MICON           =   "DefUsers.frx":0131
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   10305
      TabIndex        =   32
      Top             =   1995
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Height          =   315
      Left            =   9945
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   1995
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "DefUsers.frx":014D
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   9270
      TabIndex        =   35
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   10305
      TabIndex        =   34
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
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
      Index           =   0
      Left            =   2700
      TabIndex        =   22
      Top             =   270
      Width           =   1035
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Type password:"
      Height          =   225
      Left            =   9285
      TabIndex        =   19
      Top             =   3705
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   225
      Left            =   9285
      TabIndex        =   18
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   225
      Left            =   9285
      TabIndex        =   17
      Top             =   2490
      Width           =   1335
   End
End
Attribute VB_Name = "DefUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.
Public ParaInUserNo As Integer

Private Sub TxtStoreID_Change()
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If Trim(TxtStoreID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectStore(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectStore(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtName.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub
Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectStore = False: Exit Function
        TxtStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Stores where islock = 0 and StoreID=" & Val(TxtStoreID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtStoreName.Text = !StoreName
          FunSelectStore = True
          .Close
          Exit Function
      Else
          FunSelectStore = False
          .Close
          TxtStoreID.Text = ""
          TxtStoreName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function


Private Sub ChkDelete_Click()
   If ActiveControl.Name <> ChkDelete.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkDelete_GotFocus()
   ChkDelete.ForeColor = vbRed
End Sub

Private Sub ChkDelete_LostFocus()
   ChkDelete.ForeColor = vbBlack
End Sub

Private Sub ChkChangeRetail_Click()
   If ActiveControl.Name <> ChkChangeRetail.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkChangeRetail_GotFocus()
   ChkChangeRetail.ForeColor = vbRed
End Sub

Private Sub ChkChangeRetail_LostFocus()
   ChkChangeRetail.ForeColor = vbBlack
End Sub
Private Sub ChkChangePriceSaleInvoice_Click()
   If ActiveControl.Name <> ChkChangePriceSaleInvoice.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkChangePriceSaleInvoice_GotFocus()
   ChkChangePriceSaleInvoice.ForeColor = vbRed
End Sub

Private Sub ChkChangePriceSaleInvoice_LostFocus()
   ChkChangePriceSaleInvoice.ForeColor = vbBlack
End Sub

Private Sub ChkEdit_Click()
   If ActiveControl.Name <> ChkEdit.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkEdit_GotFocus()
   ChkEdit.ForeColor = vbRed
End Sub

Private Sub ChkEdit_LostFocus()
   ChkEdit.ForeColor = vbBlack
End Sub

Private Sub ChkEditDefination_Click()
   If ActiveControl.Name <> ChkEditDefination.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkEditDefination_GotFocus()
   ChkEditDefination.ForeColor = vbRed
End Sub

Private Sub ChkEditDefination_LostFocus()
   ChkEditDefination.ForeColor = vbBlack
End Sub

Private Sub BtnClear_Click()
  FormStatus = SelectionMode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
    End If
    If Shift = vbCtrlMask Then
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
            Case vbKeyN
                If BtnNew.Enabled Then BtnNew_Click
                KeyCode = 0
            Case vbKeyO
                If BtnOpen.Enabled Then BtnOpen_Click
                KeyCode = 0
            Case vbKeyR
                If BtnDelete.Enabled Then BtnDelete_Click
                KeyCode = 0
        End Select
    End If
    If UCase(Me.ActiveControl.Name) Like UCase("txt*") Then FormStatus = ChangeMode
    If UCase(Me.ActiveControl.Name) Like UCase("opt*") Then FormStatus = ChangeMode
    If UCase(Me.ActiveControl.Name) Like UCase("chk*") Then FormStatus = ChangeMode
    Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  Dim vTbl As String
  If Rs.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
    vTbl = Common.ChildDataExists("Users", "UserNo=" & Rs!UserNo, "")
    If vTbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vTbl, vbCritical, "Error"
      Exit Sub
    End If
    CN.Execute "Exec ProdActivityLog 'Users'," & Me.ParaInUserNo & ",3," & Rs!UserNo
    Rs.Delete
    If Rs.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
    Rs.MoveNext
    Grid.MoveNext
    If Rs.EOF Then Rs.MoveLast
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnNew_Click()
  FormStatus = NewMode
End Sub

Private Sub BtnOpen_Click()
  On Error GoTo ErrorHandler
  If Rs.RecordCount > 0 Then
    If Rs.BOF = False And Rs.EOF = False Then
      FormStatus = OpenMode
    End If
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  If FunValidation = False Then Exit Sub
  If vIsNewRecord = False Then CN.Execute "Exec ProdActivityLog 'Users'," & Me.ParaInUserNo & ",2," & Rs!UserNo
  If vIsNewRecord Then
    Rs.AddNew
    Rs!UserNo = FunGetMaxID()
    Rs!isChanged = 0
  Else
    Rs!isChanged = 1
  End If
  Rs!UserName = TxtName.Text
  Rs!StoreID = IIf(Trim(TxtStoreID.Text) = "", Null, TxtStoreID.Text)
  Rs!password = EncryptStr(IIf(txtPassword.Text = "", "empty", txtPassword.Text), True)
  Rs!IsAdministrator = OptAdmin.Value
  Rs!IsManager = OptManager.Value
  Rs!IsEdit = IIf(OptAdmin.Value = True, 1, ChkEdit.Value)
  Rs!IsEditDefination = IIf(OptAdmin.Value = True, 1, ChkEditDefination.Value)
  Rs!IsDelete = IIf(OptAdmin.Value = True, 1, ChkDelete.Value)
  Rs!IsChangeRetail = ChkChangeRetail.Value
  Rs!ChangePriceSaleInvoice = ChkChangePriceSaleInvoice.Value
  Rs!ShowPrice = ChkShowPrice.Value
  Rs!LastPurchasePrice = ChkLastPurchasePrice.Value
  Rs!WeightedPrice = ChkWeightedPrice.Value
  Rs!WSPrice = ChkWSPrice.Value
  Rs!IsReadOnly = ChkReadOnly.Value
  Rs!IsEditClosingInvoice = ChkEditClosingInvoice.Value
  Rs!IsLock = ChkLockUser.Value
  Rs.Update
  If vIsNewRecord = True Then CN.Execute "Exec ProdActivityLog 'Users'," & Me.ParaInUserNo & ",1," & Rs!UserNo
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Integer
  On Error GoTo ErrorHandler
  FunGetMaxID = CN.Execute("Select isnull(max(userno),0)+1 from users").Fields(0)
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
  If Trim(TxtName.Text) = "" Then
    MsgBox "Please specify a user name", vbExclamation, "Alert"
    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
    Exit Function
  End If
'  If Trim(txtPassword.Text) = "" Then
'    MsgBox "Please specify a password for the user", vbExclamation, "Alert"
'    If txtPassword.Enabled And txtPassword.Visible Then txtPassword.SetFocus
'    Exit Function
'  End If
  If StrComp(txtPassword.Text, TxtConfirmPassword.Text, vbBinaryCompare) <> 0 Then
    MsgBox "Your both passwords don't match. Please try again", vbExclamation, "Alert"
    TxtConfirmPassword.SetFocus
    Exit Function
  End If
  If vIsNewRecord Then
    If CN.Execute("Select * from users where username = '" & TxtName.Text & "'").RecordCount > 0 Then
      MsgBox "This user name already exists for the system. Please provide a unique user name and try again", vbExclamation, "Alert"
      TxtName.SetFocus
      Exit Function
    End If
  End If
  'All Ok, now validation is success
  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hwnd, "Users"
   Set Rs = New ADODB.Recordset
   Rs.Open "Select * FROM users where userno<>1", CN, adOpenDynamic, adLockOptimistic
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "userno"
   Grid.Columns("Name").DataField = "username"
   FormStatus = NewMode
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
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtName.Enabled = False
      txtPassword.Enabled = False
      TxtConfirmPassword.Enabled = False
      TxtName.Text = ""
      txtPassword.Text = ""
      TxtStoreID.Text = ""
      TxtStoreName.Text = ""
      TxtConfirmPassword.Text = ""
      TxtName.Enabled = True
      txtPassword.Enabled = True
      TxtConfirmPassword.Enabled = True
      Frame1.Enabled = True
      FraFacility.Enabled = True
      OptNormal.Value = True
      ChkDelete.Value = 0
      ChkEdit.Value = 0
      ChkEditDefination.Value = 0
      ChkLockUser.Value = 0
      ChkChangeRetail.Value = 0
      ChkChangePriceSaleInvoice.Value = 0
      Grid.Enabled = False
      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnClear.Enabled = True
      Grid.Enabled = False
      TxtName.Enabled = True
      txtPassword.Enabled = True
      TxtConfirmPassword.Enabled = True
      Frame1.Enabled = True
      FraFacility.Enabled = True
      If OptAdmin = True Then FraFacility.Enabled = False
      If OptNormal = True Or OptManager = True Then FraFacility.Enabled = True
      TxtName.SetFocus
      vIsNewRecord = False
    Case Is = ChangeMode
      BtnSave.Enabled = True
    Case Is = SelectionMode
      Grid.Enabled = True
      BtnNew.Enabled = True
      If mvarIsAdministrator = True Then
         BtnOpen.Enabled = True
         BtnDelete.Enabled = True
      Else
         BtnOpen.Enabled = False
         BtnDelete.Enabled = False
      End If
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtName.Enabled = False
      txtPassword.Enabled = False
      TxtConfirmPassword.Enabled = False
      Frame1.Enabled = False
      FraFacility.Enabled = False
      Call Grid_RowColChange(0, 0)
      Grid.SetFocus
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Grid_Click()
On Error GoTo ErrorHandler
    If Grid.Rows > 0 Then Call Grid_RowColChange(0, 0)
    Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo ErrorHandler
  If Rs.RecordCount > 0 And Grid.Enabled Then
    TxtName.Text = Grid.Columns("Name").Text
    TxtStoreID.Text = IIf(IsNull(Rs!StoreID), "", Rs!StoreID)
      If Trim(TxtStoreID.Text) <> "" Then
         TxtStoreName.Text = CN.Execute("Select StoreName from Stores Where StoreID = '" & TxtStoreID.Text & "'").Fields(0)
      Else
         TxtStoreName.Text = ""
      End If
    txtPassword.Text = UnEncryptStr(Rs!password, True)
    TxtConfirmPassword.Text = UnEncryptStr(Rs!password, True)
    OptAdmin.Value = Rs!IsAdministrator
    OptManager.Value = IIf(Rs!IsAdministrator = True, False, Rs!IsManager)
    OptNormal.Value = IIf(Rs!IsAdministrator = True, False, Not Rs!IsManager)
    ChkEdit.Value = Abs(Rs!IsEdit)
    ChkEditDefination.Value = Abs(Rs!IsEditDefination)
    ChkDelete.Value = Abs(Rs!IsDelete)
    ChkChangeRetail.Value = Abs(Rs!IsChangeRetail)
    ChkChangePriceSaleInvoice.Value = Abs(Rs!ChangePriceSaleInvoice)
    ChkShowPrice.Value = Abs(Rs!ShowPrice)
    ChkLastPurchasePrice.Value = Abs(Rs!LastPurchasePrice)
    ChkWeightedPrice.Value = Abs(Rs!WeightedPrice)
    ChkWSPrice.Value = Abs(Rs!WSPrice)
    ChkReadOnly.Value = Abs(Rs!IsReadOnly)
    ChkLockUser.Value = Abs(Rs!IsLock)
    ChkEditClosingInvoice.Value = Abs(Rs!IsEditClosingInvoice)
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub OptAdmin_Click()
   If ActiveControl.Name <> OptAdmin.Name Then Exit Sub
   If FraFacility.Enabled = True Then FraFacility.Enabled = False
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub OptAdmin_GotFocus()
   OptAdmin.ForeColor = vbRed
End Sub

Private Sub OptAdmin_LostFocus()
   OptAdmin.ForeColor = vbBlack
End Sub

Private Sub OptManager_Click()
   If ActiveControl.Name <> OptManager.Name Then Exit Sub
   If FraFacility.Enabled = False Then FraFacility.Enabled = True
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub OptManager_GotFocus()
   OptManager.ForeColor = vbRed
End Sub

Private Sub OptManager_LostFocus()
   OptManager.ForeColor = vbBlack
End Sub

Private Sub OptNormal_Click()
   If ActiveControl.Name <> OptNormal.Name Then Exit Sub
   If FraFacility.Enabled = False Then FraFacility.Enabled = True
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub OptNormal_GotFocus()
   OptNormal.ForeColor = vbRed
End Sub

Private Sub OptNormal_LostFocus()
   OptNormal.ForeColor = vbBlack
End Sub

