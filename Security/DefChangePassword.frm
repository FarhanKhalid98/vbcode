VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form DefChangePassword 
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
   Begin VB.TextBox TxtNewPassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   6000
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5153
      Width           =   3360
   End
   Begin VB.TextBox TxtConfirmPassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   6000
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5843
      Width           =   3360
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   6000
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4478
      Width           =   3360
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3788
      Width           =   3360
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6300
      TabIndex        =   4
      Top             =   7148
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
      MICON           =   "DefChangePassword.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7650
      TabIndex        =   5
      Top             =   7148
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
      MICON           =   "DefChangePassword.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
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
      TabIndex        =   10
      Top             =   270
      Width           =   3195
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      Height          =   225
      Left            =   6000
      TabIndex        =   9
      Top             =   4928
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Type new password:"
      Height          =   225
      Left            =   6000
      TabIndex        =   8
      Top             =   5618
      Width           =   1980
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
      Height          =   225
      Left            =   6000
      TabIndex        =   7
      Top             =   4253
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   225
      Left            =   6000
      TabIndex        =   6
      Top             =   3563
      Width           =   1335
   End
End
Attribute VB_Name = "DefChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ParaInUserNo As Integer

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event 9, 1, 1, 1
    KeyCode = 0
  End If
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  If FunValidation = False Then Exit Sub
  CN.Execute "Exec ProdActivityLog 'Change Password'," & Me.ParaInUserNo & ",2," & Me.ParaInUserNo
  CN.Execute ("UPDATE Users Set Password = '" & Replace(EncryptStr(IIf(TxtNewPassword.Text = "", "empty", TxtNewPassword.Text), True), "'", "''") & "' Where UserNo = " & Me.ParaInUserNo)
  MsgBox "Your password has been changed successfully", vbInformation, "Information"
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
'  If Trim(TxtName.Text) = "" Then
'    MsgBox "Please specify a user name", vbExclamation, "Alert"
'    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
'    Exit Function
'  End If
'  If Trim(TxtNewPassword.Text) = "" Then
'    MsgBox "Please specify the New Password.", vbExclamation, "Alert"
'    If TxtNewPassword.Enabled And TxtNewPassword.Visible Then TxtNewPassword.SetFocus
'    Exit Function
'  End If
  If StrComp(TxtNewPassword.Text, TxtConfirmPassword.Text, vbBinaryCompare) <> 0 Then
    MsgBox "Your both new passwords don't match. Please try again", vbExclamation, "Alert"
    TxtConfirmPassword.SetFocus
    Exit Function
  End If
  If StrComp(EncryptStr(IIf(txtPassword.Text = "", "empty", txtPassword.Text), True), CN.Execute("Select password from users where userno = " & Me.ParaInUserNo & "").Fields(0), vbBinaryCompare) <> 0 Then
    MsgBox "Incorrect Old Password. Please try again.", vbExclamation, "Alert"
    txtPassword.SetFocus
    Exit Function
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
   SetWindowText Me.hwnd, "Change Password"
   With CN.Execute("Select * FROM users Where UserNo = " & Me.ParaInUserNo)
      If .RecordCount = 0 Then
         MsgBox "This user don't exists in the system.", vbCritical, "Error"
         Exit Sub
      End If
      TxtName.Text = !UserName
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub
