VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form DefCreateAccounts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "DefCreateAccounts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
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
      Height          =   2355
      Left            =   11430
      TabIndex        =   14
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
         Height          =   1950
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Tag             =   "NC"
         Text            =   "DefCreateAccounts.frx":0ECA
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
         TabIndex        =   16
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtNarration 
      Appearance      =   0  'Flat
      Height          =   600
      Left            =   4403
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   5588
      Width           =   6555
   End
   Begin VB.TextBox TxtAccountType 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   4403
      MaxLength       =   30
      TabIndex        =   5
      Top             =   4268
      Width           =   4485
   End
   Begin VB.TextBox TxtPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   4403
      MaxLength       =   10
      TabIndex        =   4
      Top             =   3518
      Width           =   1005
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5423
      MaxLength       =   10
      TabIndex        =   0
      Top             =   3518
      Width           =   1200
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4403
      MaxLength       =   30
      TabIndex        =   1
      Top             =   4958
      Width           =   4485
   End
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   5723
      TabIndex        =   9
      Top             =   7418
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "New"
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
      MICON           =   "DefCreateAccounts.frx":0F31
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   4418
      TabIndex        =   8
      Top             =   7418
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
      MICON           =   "DefCreateAccounts.frx":0F4D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7043
      TabIndex        =   10
      Top             =   7418
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
      MICON           =   "DefCreateAccounts.frx":0F69
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   9083
      TabIndex        =   11
      Top             =   7418
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
      MICON           =   "DefCreateAccounts.frx":0F85
      BC              =   14737632
      FC              =   0
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
      Left            =   11385
      TabIndex        =   17
      Top             =   540
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create Accounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   13
      Top             =   270
      Width           =   2325
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11640
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Narration:"
      Height          =   225
      Left            =   4403
      TabIndex        =   12
      Top             =   5363
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
      Height          =   225
      Left            =   4403
      TabIndex        =   6
      Top             =   4043
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account No:"
      Height          =   225
      Left            =   4403
      TabIndex        =   3
      Top             =   3293
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name:"
      Height          =   225
      Left            =   4403
      TabIndex        =   2
      Top             =   4718
      Width           =   1335
   End
End
Attribute VB_Name = "DefCreateAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vMode As FormMode
Public ParaInIsNew As Boolean     'Whether new record or not.
Public ParaInAccountNo As String  'For opening/modification only
Public ParaInParentAccountNo  As String 'Where this account will be created.
Public ParaInIsGroup As Boolean 'Whether this is a group account or Trans Account
Public ParaInParentAccountName As String  'Will be displayed in Account Type Field
Public ParaInIsLocked As Boolean 'Whether Save button will be enabled
Public ParaOutUpdateSuccess As Boolean  'Will refresh the Grid in Caller form

Private Sub BtnClear_Click()
  FormStatus = SelectionMode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Case vbKeyW
                If BtnClear.Enabled Then BtnClear_Click
                KeyCode = 0
            Case vbKeyQ
                If BtnClose.Enabled Then BtnClose_Click
                KeyCode = 0
            Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
               KeyCode = 0
            Case vbKeyN
                If BtnNew.Enabled Then BtnNew_Click
                KeyCode = 0
        End Select
    End If
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

'Private Sub BtnDelete_Click()
'  On Error GoTo ErrorHandler
'  If Rs.RecordCount > 0 Then
'    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
'    Rs.Delete
'    If Rs.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
'    Rs.MoveNext
'    If Rs.EOF Then Rs.MoveLast
'  End If
'  Exit Sub
'ErrorHandler:
'  Call ShowErrorMessage
'End Sub

Private Sub BtnNew_Click()
  FormStatus = NewMode
End Sub

'Private Sub BtnOpen_Click()
'  On Error GoTo ErrorHandler
'  If Rs.RecordCount > 0 Then
'    If Rs.BOF = False And Rs.EOF = False Then
'      FormStatus = OpenMode
'    End If
'  End If
'  Exit Sub
'ErrorHandler:
'  Call ShowErrorMessage
'End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  If FunValidation = False Then Exit Sub
  If ParaInIsNew = False Then Call ActivityLog("Chart of Accounts", eEdit, , , TxtPrefix.Text & TxtID.Text)
  If ParaInIsNew Then
    Rs.AddNew
    Rs!AccountNo = TxtPrefix.Text & TxtID.Text
    Rs!AccountType = Me.ParaInParentAccountName 'CN.Execute("Select AccountName From ChartOfAccounts where Accountno = '" & Me.ParaInParentAccountNo & "'").Fields(0)
    Rs!UserNo = vUser
    Rs!AccountDepth = cn.Execute("Select AccountDepth+1 From ChartOfAccounts where Accountno = " & Me.ParaInParentAccountNo).Fields(0)
    Rs!ParentAccountNo = Me.ParaInParentAccountNo
    Rs!isdetailed = Not (Me.ParaInIsGroup)
    Rs!openingdebit = 0
    Rs!openingCredit = 0
    Rs!IsLocked = 0
    Rs!iseditable = 1
    Rs!balflag = 0
    Rs!plflag = 0
    Rs!ExpFlag = 0
  End If
  Rs!AccountName = TxtName.Text
  Rs!Narration = TxtNarration.Text
  Rs.Update
  If ParaInIsNew = True Then Call ActivityLog("Chart of Accounts", eAdd, , , TxtPrefix.Text & TxtID.Text)
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
  Dim vCounter As Integer
  If Me.ParaInIsNew Then
    If Trim(TxtID.Text) = "" Then
      MsgBox "Please specify a Account No.", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
    If cn.Execute("Select * From ChartofAccounts where Accountno = " & Val(TxtPrefix.Text & TxtID.Text)).RecordCount > 0 Then
        MsgBox "This account No. already exists. Please Provide another account No.", vbExclamation, "Alert"
        TxtID.SetFocus
        Exit Function
    End If
    For vCounter = 1 To Len(Trim(TxtID.Text))
      If Asc(UCase(Mid(TxtID.Text, vCounter, 1))) < 48 Or Asc(UCase(Mid(TxtID.Text, vCounter, 1))) > 57 Then
        MsgBox "The Account No. must contain Numeric characters only.", vbExclamation, "Alert"
        If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
        
      End If
    Next
  End If
  If Trim(TxtName.Text) = "" Then
    MsgBox "Please specify the Account name", vbExclamation, "Alert"
    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
    Exit Function
  End If
  'All Ok, now validation is success
  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
  If UCase(Me.ActiveControl.Name) Like "TXT*" Then FormStatus = ChangeMode
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
  SetWindowText Me.hWnd, "Create Accounts"
  ShowPicture Me, 2
  AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
  HelpLocation Me
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * FROM ChartOfAccounts Where AccountNo = " & Val(Me.ParaInAccountNo), cn, adOpenDynamic, adLockOptimistic
    If Rs.RecordCount > 0 And Me.ParaInIsNew = False Then
      TxtPrefix.Text = IIf(IsNull(Rs!ParentAccountNo), "", Rs!ParentAccountNo)
      'TxtID.Text = Replace(Rs!AccountNo, TxtPrefix.Text, "") ' old
      ' new edit by farhan on 25-07-2007
      TxtID.Text = Mid(Rs!AccountNo, Len(TxtPrefix.Text) + 1, Len(Rs!AccountNo))
      TxtAccountType.Text = Me.ParaInParentAccountName
      TxtName.Text = Rs!AccountName
      TxtNarration.Text = IIf(IsNull(Rs!Narration), "", Rs!Narration)
      FormStatus = OpenMode
    Else
      TxtPrefix.Text = Me.ParaInParentAccountNo
      TxtAccountType.Text = Me.ParaInParentAccountName
      FormStatus = NewMode
    End If
    Me.ParaOutUpdateSuccess = False
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
      TxtNarration.Text = ""
      Me.ParaInIsNew = True
      Me.ParaInIsLocked = False
      BtnNew.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtName.Enabled = False
      TxtID.Enabled = False
      TxtName.Text = ""
      TxtID.Text = ""
      'TxtID.Text = CN.Execute("Select substring(cast(isnull(max(cast(AccountNo as int)+1),0) as Varchar(10))," & Len(Me.ParaInParentAccountNo) + 1 & ",10) From ChartOfAccounts where ParentAccountNo ='" & Me.ParaInParentAccountNo & "'").Fields(0)
      With cn.Execute("Select substring(cast(isnull(max(AccountNo+1),0) as Varchar(10))," & Len(Me.ParaInParentAccountNo) + 1 & ",10) as ID From ChartOfAccounts where ParentAccountNo ='" & Me.ParaInParentAccountNo & "'")
         If .RecordCount = 0 Then
            TxtID.Text = 1
         Else
            TxtID.Text = !ID
         End If
      End With
      TxtName.Enabled = True
      TxtID.Enabled = True
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      ParaInIsNew = True
    Case Is = OpenMode
      BtnNew.Enabled = False
      BtnClear.Enabled = False
      TxtName.Enabled = True
      TxtID.Enabled = False
      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
      
    Case Is = ChangeMode
      If Me.ParaInIsLocked = False Then BtnSave.Enabled = True
      
    Case Is = SelectionMode
'      BtnNew.Enabled = True
'      BtnSave.Enabled = False
'      BtnClear.Enabled = False
'      TxtName.Enabled = False
'      TxtID.Enabled = False
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Form_Unload(Cancel As Integer)
  Me.ParaInIsGroup = False
  Me.ParaInIsLocked = False
  Me.ParaInAccountNo = ""
  Me.ParaInParentAccountName = ""
  Me.ParaInParentAccountNo = ""
  Me.ParaInIsNew = False
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

