VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form DefUsers 
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
   Picture         =   "DefUsers.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtConfirmPassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   6600
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4785
      Width           =   3360
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   6600
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4200
      Width           =   3360
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4320
      Left            =   1335
      TabIndex        =   3
      Top             =   2535
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
      stylesets(0).Picture=   "DefUsers.frx":6971
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
      Columns(0).Width=   1852
      Columns(0).Caption=   "User No."
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4948
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
      Left            =   6600
      MaxLength       =   30
      TabIndex        =   0
      Top             =   3570
      Width           =   3360
   End
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   1020
      TabIndex        =   7
      Top             =   7800
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
      MICON           =   "DefUsers.frx":698D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   2340
      TabIndex        =   8
      Top             =   7800
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Change"
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
      MICON           =   "DefUsers.frx":69A9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   3660
      TabIndex        =   9
      Top             =   7800
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
      MICON           =   "DefUsers.frx":69C5
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6465
      TabIndex        =   10
      Top             =   7800
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
      MICON           =   "DefUsers.frx":69E1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7785
      TabIndex        =   11
      Top             =   7800
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
      MICON           =   "DefUsers.frx":69FD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   9105
      TabIndex        =   12
      Top             =   7800
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
      MICON           =   "DefUsers.frx":6A19
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Type password:"
      Height          =   225
      Left            =   6600
      TabIndex        =   6
      Top             =   4560
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   225
      Left            =   6600
      TabIndex        =   5
      Top             =   3975
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   225
      Left            =   6600
      TabIndex        =   4
      Top             =   3345
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
  If vIsNewRecord Then
    Rs.AddNew
    Rs!UserNo = FunGetMaxID()
  End If
  Rs!UserName = TxtName.Text
  Rs!password = TxtPassword.Text
  Rs!IsAdministrator = 1
  Rs.Update
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
  If Trim(TxtPassword.Text) = "" Then
    MsgBox "Please specify a password for the user", vbExclamation, "Alert"
    If TxtPassword.Enabled And TxtPassword.Visible Then TxtPassword.SetFocus
    Exit Function
  End If
  If StrComp(TxtPassword.Text, TxtConfirmPassword.Text, vbTextCompare) <> 0 Then
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
      TxtPassword.Enabled = False
      TxtConfirmPassword.Enabled = False
      TxtName.Text = ""
      TxtPassword.Text = ""
      TxtConfirmPassword.Text = ""
      
      TxtName.Enabled = True
      TxtPassword.Enabled = True
      TxtConfirmPassword.Enabled = True
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
      TxtPassword.Enabled = True
      TxtConfirmPassword.Enabled = True
      TxtName.SetFocus
      vIsNewRecord = False
    Case Is = ChangeMode
      BtnSave.Enabled = True
    Case Is = SelectionMode
      Grid.Enabled = True
      BtnNew.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnSave.Enabled = False
      BtnClear.Enabled = False
      TxtName.Enabled = False
      TxtPassword.Enabled = False
      TxtConfirmPassword.Enabled = False
      Grid.SetFocus
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo ErrorHandler
  If Rs.RecordCount > 0 And Grid.Enabled Then
    TxtName.Text = Grid.Columns("Name").Text
    TxtPassword.Text = Rs!password
    TxtConfirmPassword.Text = Rs!password
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtName_Change()
  If TxtName.Enabled = True Then FormStatus = ChangeMode
End Sub
Private Sub TxtPassword_Change()
  If TxtPassword.Enabled = True Then FormStatus = ChangeMode
End Sub
Private Sub TxtconfirmPassword_Change()
  If TxtConfirmPassword.Enabled = True Then FormStatus = ChangeMode
End Sub

