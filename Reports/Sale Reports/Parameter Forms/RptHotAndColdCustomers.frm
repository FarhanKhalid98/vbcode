VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form RptHotAndColdCustomers 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptHotAndColdCustomers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtZoneName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6578
      TabIndex        =   21
      Top             =   5501
      Width           =   3585
   End
   Begin VB.TextBox TxtZoneID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5198
      TabIndex        =   5
      Top             =   5501
      Width           =   1020
   End
   Begin VB.TextBox TxtSectorID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5198
      TabIndex        =   6
      Top             =   6206
      Width           =   1020
   End
   Begin VB.TextBox TxtSectorName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6578
      TabIndex        =   20
      Top             =   6206
      Width           =   3585
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EFC09E&
      Caption         =   "Direction"
      Height          =   705
      Left            =   6075
      TabIndex        =   18
      Top             =   2809
      Width           =   1665
      Begin VB.OptionButton OptCold 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Cold"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   855
         TabIndex        =   1
         Top             =   360
         Width           =   675
      End
      Begin VB.OptionButton OptHot 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Hot"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EFC09E&
      Caption         =   "Criteria"
      Height          =   1110
      Left            =   2385
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   1755
      Begin VB.OptionButton OptQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Unit Sold"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   315
         TabIndex        =   12
         ToolTipText     =   "Standard users are restricted to perform only those tasks which are explicitly assigned to them by the System administrator."
         Top             =   360
         Width           =   960
      End
      Begin VB.OptionButton OptAmt 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Sale Value"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   315
         TabIndex        =   13
         ToolTipText     =   $"RptHotAndColdCustomers.frx":0ECA
         Top             =   720
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8610
      TabIndex        =   11
      Top             =   8104
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Close"
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
      MICON           =   "RptHotAndColdCustomers.frx":0F53
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   5835
      TabIndex        =   9
      Top             =   8104
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Pre&view"
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
      MICON           =   "RptHotAndColdCustomers.frx":0F6F
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   7215
      TabIndex        =   10
      Top             =   8104
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Print"
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
      MICON           =   "RptHotAndColdCustomers.frx":0F8B
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   6375
      TabIndex        =   7
      Top             =   7301
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "/"
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpTo 
      Height          =   315
      Left            =   8130
      TabIndex        =   8
      Top             =   7301
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "/"
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SITextBox.Txt TxtNoOfProduct 
      Height          =   315
      Left            =   8055
      TabIndex        =   2
      Top             =   3049
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      Alignment       =   2
      Appearance      =   0
      MaxLength       =   4
      Text            =   "20"
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
   End
   Begin JeweledBut.JeweledButton BtnZone 
      Height          =   330
      Left            =   6225
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5501
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
      MICON           =   "RptHotAndColdCustomers.frx":0FA7
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSector 
      Height          =   330
      Left            =   6225
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6206
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
      MICON           =   "RptHotAndColdCustomers.frx":0FC3
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOrganizaton 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6210
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4151
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "RptHotAndColdCustomers.frx":0FDF
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   5190
      TabIndex        =   3
      Top             =   4151
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtOrganizatonName 
      Height          =   315
      Left            =   6570
      TabIndex        =   29
      Tag             =   "nc"
      Top             =   4151
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin JeweledBut.JeweledButton BtnCompany 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6225
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4811
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "RptHotAndColdCustomers.frx":0FFB
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   5205
      TabIndex        =   4
      Top             =   4811
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   3
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   6585
      TabIndex        =   33
      Tag             =   "nc"
      Top             =   4811
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company ID"
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
      Left            =   5190
      TabIndex        =   35
      Top             =   4586
      Width           =   1035
   End
   Begin VB.Label Label3 
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
      Left            =   6585
      TabIndex        =   34
      Top             =   4586
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
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
      Left            =   6570
      TabIndex        =   31
      Top             =   3956
      Width           =   1620
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
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
      Left            =   5190
      TabIndex        =   30
      Top             =   3956
      Width           =   1335
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone Name"
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
      Left            =   6585
      TabIndex        =   27
      Top             =   5291
      Width           =   990
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone ID"
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
      Left            =   5205
      TabIndex        =   26
      Top             =   5291
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector ID"
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
      Left            =   5205
      TabIndex        =   25
      Top             =   5996
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector Name"
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
      Left            =   6585
      TabIndex        =   24
      Top             =   5996
      Width           =   1110
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "No Of Customers"
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
      Left            =   7845
      TabIndex        =   19
      Top             =   2831
      Width           =   1440
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hot And Cold Customers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   16
      Top             =   270
      Width           =   4290
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Left            =   8145
      TabIndex        =   15
      Top             =   7076
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      Left            =   6375
      TabIndex        =   14
      Top             =   7076
      Width           =   885
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove this Row"
      End
   End
End
Attribute VB_Name = "RptHotAndColdCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql As String

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
       RptReportViewer.Caption = Me.Caption
       RptReportViewer.Show vbModal
   End If
End Sub

Private Sub BtnPrint_Click()
   If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
         Case vbKeyV
            If BtnPreview.Enabled Then BtnPreview_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus Else TxtOrganizationID.SetFocus
         Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus Else TxtZoneID.SetFocus
         Case TxtZoneID.Name: If FunSelectZone(ssFunctionKey, True) = True Then TxtSectorID.SetFocus Else TxtZoneID.SetFocus
         Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, True) = True Then DtpFrom.SetFocus Else TxtSectorID.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Hot And Cold Customers"
   DtpTo.DateValue = Date
   DtpFrom.DateValue = Date - 30
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
   Set RptHotAndColdCustomers = Nothing
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function SetReport() As Boolean
   On Error GoTo ErrorHandler
   SetReport = False
   Me.MousePointer = vbHourglass
   Dim RsReport As New ADODB.Recordset
   
   Set RsReport = cn.Execute("EXEC ProdRptHotAndColdCustomers " & "'" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text) & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", TxtCompanyID.Text) & "," & IIf(Trim(TxtZoneID.Text) = "", "Null", TxtZoneID.Text) & "," & IIf(Trim(TxtSectorID.Text) = "", "Null", TxtSectorID.Text) & "," & IIf(OptAmt.Value = True, "SaleAmt", "Visits") & "," & IIf(OptHot.Value = True, 1, 0) & ",'" & Val(TxtNoOfProduct.Text) & "'")

   'Set RsReport = CN.Execute("EXEC ProdRptHotAndColdCustomers " & "'" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtCompanyID.Text) = "", "''", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "''", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "''", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtBrandID.Text) = "", "''", "'" & TxtBrandID.Text & "'") & "," & IIf(OptAmt.Value = True, "'Value '", "' Qty '") & "," & IIf(OptHot.Value = True, 1, 0) & ",'" & Val(TxtNoOfProduct.Text) & "'")
   'Set RsReport = CN.Execute("EXEC ProdRptProductWiseProfit null,null,null,null,'2007-07-27', '2008-09-27'")
   Set RptReportViewer.Report = New CrptHotAndColdCustomers
   'Set RptReportViewer.Report = New CRptCurrentExpiryValue
   If RsReport.BOF Then
      MsgBox "No record exists.", vbInformation, Me.Caption
      Me.MousePointer = vbDefault
      Exit Function
   End If
   RptReportViewer.Report.ReportTitle = "Hot and Cold Customers"
   RptReportViewer.Report.Database.SetDataSource RsReport
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
'    RptReportViewer.Report.PaperOrientation = crLandscape
   SetReport = True
   Me.MousePointer = vbDefault
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnZone_Click()
   If FunSelectZone(ssButton, False) = True Then
      TxtSectorID.SetFocus
   Else
      TxtZoneID.SetFocus
   End If
End Sub

Private Sub TxtZoneID_Change()
   If TxtZoneID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtZoneID.Name Then Exit Sub
   If TxtZoneName.Text <> "" Then TxtZoneName.Text = ""
End Sub

Private Sub TxtZoneID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtZoneID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtZoneID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectZone(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectZone(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectZone(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchZone.Show vbModal, Me
      If SchZone.ParaOutZoneID = "" Then FunSelectZone = False: Exit Function
      TxtZoneID.Text = SchZone.ParaOutZoneID
   End If
    '---------------------------
   VStrSQL = " Select * FROM Zones where ZoneID=" & Val(TxtZoneID.Text)
   With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtZoneName.Text = !ZoneName
         FunSelectZone = True
         .Close
         Exit Function
      Else
         FunSelectZone = False
         .Close
         TxtZoneID.Text = ""
         TxtZoneName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnSector_Click()
   If FunSelectSector(ssButton, False) = True Then
      DtpFrom.SetFocus
   Else
      TxtSectorID.SetFocus
   End If
End Sub

Private Sub TxtSectorID_Change()
   If TxtSectorID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   If TxtSectorName.Text <> "" Then TxtSectorName.Text = ""
End Sub

Private Sub TxtSectorID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSectorID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSector(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSector(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSector(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchSector.Show vbModal, Me
      If SchSector.ParaOutSectorID = "" Then FunSelectSector = False: Exit Function
      TxtSectorID.Text = SchSector.ParaOutSectorID
   End If
    '---------------------------
   VStrSQL = " Select * FROM Sectors where SectorID=" & Val(TxtSectorID.Text)
   With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtSectorName.Text = !SectorName
         FunSelectSector = True
         .Close
         Exit Function
      Else
         FunSelectSector = False
         .Close
         TxtSectorID.Text = ""
         TxtSectorName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnOrganizaton_Click()
   If FunSelectOrganization(ssButton, False) = True Then
      TxtCompanyID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub TxtOrganizationID_Change()
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizatonName.Text <> "" Then TxtOrganizatonName.Text = ""
End Sub

Private Sub TxtOrganizationID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtOrganizationID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectOrganization(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganization(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectOrganization(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganization = False: Exit Function
       TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    If TxtOrganizationID.Text = "" Then FunSelectOrganization = False: Exit Function
    VStrSQL = " Select * FROM Organizations where OrganizationID='" & TxtOrganizationID.Text & "'"
    With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizatonName.Text = !OrganizationName
          FunSelectOrganization = True
          .Close
          Exit Function
      Else
          FunSelectOrganization = False
          .Close
         TxtOrganizationID.Text = ""
          TxtOrganizatonName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnCompany_Click()
   If FunSelectCompany(ssButton, False) = True Then
      TxtZoneID.SetFocus
   Else
      TxtCompanyID.SetFocus
   End If
End Sub

Private Sub TxtCompanyID_Change()
   If TxtCompanyID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCompanyID.Name Then Exit Sub
   If TxtCompanyName.Text <> "" Then TxtCompanyName.Text = ""
End Sub

Private Sub TxtCompanyID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtCompanyID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCompanyID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCompany(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCompany(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectCompany(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCompany.Show vbModal, Me
        If SchCompany.ParaOutCompanyID = "" Then FunSelectCompany = False: Exit Function
        TxtCompanyID.Text = SchCompany.ParaOutCompanyID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Companies where CompanyID=" & Val(TxtCompanyID.Text)
    With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtCompanyName.Text = !CompanyName
          FunSelectCompany = True
          .Close
          Exit Function
      Else
          FunSelectCompany = False
          .Close
          TxtCompanyID.Text = ""
          TxtCompanyName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

