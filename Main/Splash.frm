VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form Splash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar LblProgress 
      Height          =   255
      Left            =   555
      TabIndex        =   1
      Top             =   4455
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
      Scrolling       =   1
   End
   Begin VB.Label LblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading... Please wait!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   585
      TabIndex        =   0
      Top             =   4200
      Width           =   1650
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vCounter As Integer

Private Sub Form_Load()
    'LblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
    SetWindowText Me.hwnd, "Splash"
    vCounter = 1
End Sub
'
'Private Sub Timer1_Timer()
'   Select Case vCounter
'   Case 1:
'     Splash.LblProgress.Value = 25
'     Splash.LblStatus.Caption = "Connection established with the Database..."
'     DoEvents
'   Case 2:
'     Splash.LblProgress.Value = 35
'     Splash.LblStatus.Caption = "Initializing the Accounts Forms..."
'     DoEvents
'   Case 3:
'     Splash.LblProgress.Value = 45
'     Splash.LblStatus.Caption = "Initializing the Definition Forms..."
'     DoEvents
'   Case 4:
'     Splash.LblProgress.Value = 55
'     Splash.LblStatus.Caption = "Initializing Purchase Forms..."
'     DoEvents
'   Case 5:
'     Splash.LblProgress.Value = 65
'     Splash.LblStatus.Caption = "Initializing Sales Forms..."
'     DoEvents
'   Case 6:
'     Splash.LblProgress.Value = 75
'     Splash.LblStatus.Caption = "Initializing Installments Forms..."
'     DoEvents
'   Case 7:
'     Splash.LblProgress.Value = 85
'     Splash.LblStatus.Caption = "Initializing Sale Reports..."
'     DoEvents
'   Case 8:
'     Splash.LblProgress.Value = 95
'     Splash.LblStatus.Caption = "Initializing Daily working Reports..."
'     DoEvents
'   Case 9:
'     Splash.LblProgress.Value = 100
'     Splash.LblStatus.Caption = "Loading Main User Interface..."
'     Desktop.Show
'     Unload Splash
'   End Select
'   vCounter = vCounter + 1
'End Sub
