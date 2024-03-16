VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmCompanyInformation 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtShortName 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2925
      MaxLength       =   23
      TabIndex        =   5
      Top             =   8505
      Width           =   7410
   End
   Begin VB.TextBox TxtEMail 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2925
      MaxLength       =   50
      TabIndex        =   4
      Top             =   7650
      Width           =   7410
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2925
      MaxLength       =   100
      TabIndex        =   0
      Top             =   3900
      Width           =   7410
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2925
      MaxLength       =   100
      TabIndex        =   1
      Top             =   4830
      Width           =   7410
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2925
      MaxLength       =   50
      TabIndex        =   2
      Top             =   5760
      Width           =   7410
   End
   Begin VB.TextBox TxtPhone 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2925
      MaxLength       =   100
      TabIndex        =   3
      Top             =   6705
      Width           =   7410
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   5235
      TabIndex        =   6
      Top             =   9630
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
      MICON           =   "FrmCompanyInformation.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6585
      TabIndex        =   7
      Top             =   9630
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
      MICON           =   "FrmCompanyInformation.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Short Name For Barcode Print"
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
      Left            =   2925
      TabIndex        =   14
      Top             =   8280
      Width           =   3375
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Information"
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
      Left            =   2700
      TabIndex        =   13
      Top             =   270
      Width           =   2970
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company E.Mail"
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
      Left            =   2925
      TabIndex        =   12
      Top             =   7425
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Phone"
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
      Left            =   2925
      TabIndex        =   11
      Top             =   6480
      Width           =   1380
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
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
      Left            =   2925
      TabIndex        =   10
      Top             =   3675
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Address"
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
      Left            =   2925
      TabIndex        =   9
      Top             =   4620
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company City"
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
      Left            =   2925
      TabIndex        =   8
      Top             =   5550
      Width           =   1155
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   12165
      Top             =   1860
      Width           =   330
   End
End
Attribute VB_Name = "FrmCompanyInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtName.Text & "' where RegistryKey = 'CompanyName'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtAddress.Text & "' where RegistryKey = 'CompanyAddress'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtCity.Text & "' where RegistryKey = 'CompanyCity'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtPhone.Text & "' where RegistryKey = 'CompanyPhoneNo'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtEMail.Text & "' where RegistryKey = 'CompanyEmail'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtShortName.Text & "' where RegistryKey = 'CompanyShortName'")
   MsgBox "Your Company Information has been changed successfully", vbInformation, "Information"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
  If Trim(TxtName.Text) = "" Then
    MsgBox "Please specify a Company Name", vbExclamation, "Alert"
    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
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
   SetWindowText Me.hwnd, "Company Information"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   
   TxtName.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'CompanyName'").Fields(0).Value
   TxtAddress.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'CompanyAddress'").Fields(0).Value
   TxtCity.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'CompanyCity'").Fields(0).Value
   TxtPhone.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'CompanyPhoneNo'").Fields(0).Value
   TxtEMail.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'CompanyEmail'").Fields(0).Value
   TxtShortName.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'CompanyShortName'").Fields(0).Value
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub
