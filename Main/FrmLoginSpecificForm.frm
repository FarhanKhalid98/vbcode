VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmLoginSpecificForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   5280
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   7515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmLoginSpecificForm.frx":0000
   ScaleHeight     =   3119.599
   ScaleMode       =   0  'User
   ScaleWidth      =   7056.179
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbSession 
      Height          =   315
      Left            =   3630
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2925
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   3630
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "Admin"
      Top             =   1440
      Width           =   2325
   End
   Begin JeweledBut.JeweledButton BtnLogin 
      Height          =   420
      Left            =   3480
      TabIndex        =   4
      Top             =   3645
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
      MICON           =   "FrmLoginSpecificForm.frx":80FE
      BC              =   14737632
      FC              =   0
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3630
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Tag             =   "Admin"
      Text            =   "P@ssport0309"
      Top             =   2175
      Width           =   2325
   End
   Begin JeweledBut.JeweledButton BtnCancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   4800
      TabIndex        =   5
      Top             =   3645
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
      MICON           =   "FrmLoginSpecificForm.frx":811A
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblSession 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Session"
      Height          =   195
      Left            =   3630
      TabIndex        =   7
      Top             =   2700
      Width           =   555
   End
   Begin VB.Image ImgExit 
      Height          =   210
      Left            =   7260
      Top             =   15
      Width           =   210
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Index           =   0
      Left            =   3630
      TabIndex        =   0
      Top             =   1215
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
      Top             =   1965
      Width           =   690
   End
End
Attribute VB_Name = "FrmLoginSpecificForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ParaOutUserNo, ParaOutSessionID As Byte
Dim vMobileNo() As String, vMobile As String, sSql As String


Private Sub BtnCancel_Click()
   ParaOutUserNo = 0
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = txtPassword.Name Then
         Call BtnLogin_Click
      Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   End If
End Sub

Private Sub BtnLogin_Click()
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
  If Trim(txtPassword.Text) = "" Then
    MsgBox "Please provide a Password", vbExclamation, "Alert"
    txtPassword.SetFocus
    Exit Sub
  End If
  'Now if the username and password are provided, we shall verify the authenticity
  With CN.Execute("SElect * FROM Users Where (islock = 0 or islock is null) and Username='" & txtUserName.Text & "'")
    If .RecordCount = 0 Then
      MsgBox "This user does not exists. Please provide a valid user and try again", vbExclamation, "Alert"
      txtUserName.SelStart = 0
      txtUserName.SelLength = Len(txtUserName.Text)
      txtUserName.SetFocus
      Exit Sub
    Else
      If StrComp(txtPassword.Text, IIf(EncryptStr(!password, False) = "empty", "", EncryptStr(!password, False)), vbTextCompare) = 0 Then
        ParaOutUserNo = !UserNo
        CN.Execute "Exec ProdActivityLog 'Login'," & Me.ParaOutUserNo & ",1," & Me.ParaOutUserNo
        If CmbSession.Text <> "" Then Me.ParaOutSessionID = CmbSession.ItemData(CmbSession.ListIndex) Else Me.ParaOutSessionID = 0
        '/******* Mobile SMS *************/
         If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnLogin Then
            vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
            For i = 0 To UBound(vMobileNo)
               vMobile = "+92" + Right(vMobileNo(i), 10)
               If Len(vMobile) = 13 Then
                  sSql = txtUserName & " LogIn at " & Now
                  sSql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & sSql & "','')"
                  CN.Execute sSql
               End If
            Next
         End If
        Unload Me
      Else
        ParaOutUserNo = 0
        MsgBox "Incorrect Password. Please provide the correct password and try again", vbExclamation, "Alert"
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
        txtPassword.SetFocus
        Exit Sub
      End If
    End If
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   SetWindowText Me.hwnd, "Login"
   ParaOutUserNo = 0
   
   CmbSession.Clear
   CmbSession.AddItem ""
   With CN.Execute("select * from Sessions")
      While Not .EOF
         CmbSession.AddItem !SessionName
         CmbSession.ItemData(CmbSession.NewIndex) = !SessionID
         .MoveNext
      Wend
      .Close
   End With
   
   With CN.Execute("select * from Sessions Where SessionID = (Select value from sysindexs Where Registrykey = 'SessionID') ")
      If Not .EOF Then
         CmbSession.Text = !SessionName
      End If
      .Close
   End With
      
   If ObjRegistry.ShowSession = False Then
      LblSession.Visible = False
      CmbSession.Visible = False
   End If
   
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

