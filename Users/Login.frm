VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5280
   ClientLeft      =   2805
   ClientTop       =   3060
   ClientWidth     =   7515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Login.frx":0000
   ScaleHeight     =   3119.599
   ScaleMode       =   0  'User
   ScaleWidth      =   7056.178
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton CmdLogin 
      Default         =   -1  'True
      Height          =   420
      Left            =   3480
      TabIndex        =   4
      Top             =   3240
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Login"
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
      MICON           =   "Login.frx":80FE
      BC              =   14737632
      FC              =   0
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   3630
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1845
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3630
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2580
      Width           =   2325
   End
   Begin JeweledBut.JeweledButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   4800
      TabIndex        =   5
      Top             =   3240
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
      MICON           =   "Login.frx":811A
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Image ImgExit 
      Height          =   225
      Left            =   7275
      Top             =   15
      Width           =   180
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Index           =   0
      Left            =   3630
      TabIndex        =   0
      Top             =   1620
      Width           =   795
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Index           =   1
      Left            =   3630
      TabIndex        =   2
      Top             =   2370
      Width           =   690
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ParaOutUserNo As Byte

Private Sub CmdCancel_Click()
   ParaOutUserNo = 0
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   End If
End Sub

Private Sub CmdLogin_Click()
  On Error GoTo ErrorHandler
  If Trim(txtUserName.Text) = "" Then
    MsgBox "Please provide a user name", vbExclamation, "Alert"
    txtUserName.SetFocus
    Exit Sub
  'ElseIf Trim(txtPassword.Text) = "" Then
  '  MsgBox "Please provide a password", vbExclamation, "Alert"
  '  txtPassword.SetFocus
  '  Exit Sub
  End If
  'Now if the username and password are provided, we shall verify the authenticity
  With CN.Execute("SElect * FROM Users Where Username='" & txtUserName.Text & "'")
    If .RecordCount = 0 Then
      MsgBox "This user does not exists. Please provide a valid user and try again", vbExclamation, "Alert"
      txtUserName.SelStart = 0
      txtUserName.SelLength = Len(txtUserName.Text)
      txtUserName.SetFocus
      Exit Sub
    Else
      If StrComp(TxtPassword.Text, !password, vbTextCompare) = 0 Then
        ParaOutUserNo = !UserNo
        Unload Me
      Else
        ParaOutUserNo = 0
        MsgBox "Incorrect Password. Please provide the correct password and try again", vbExclamation, "Alert"
        TxtPassword.SelStart = 0
        TxtPassword.SelLength = Len(TxtPassword.Text)
        TxtPassword.SetFocus
        Exit Sub
      End If
    End If
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   ParaOutUserNo = 0
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub
