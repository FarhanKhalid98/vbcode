VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptLedgerNew 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   Icon            =   "RptLedgerNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkZeroBalance 
      BackColor       =   &H00FF8080&
      Caption         =   "Include Zero Balance"
      Height          =   255
      Left            =   8730
      TabIndex        =   37
      Top             =   7245
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.CheckBox ChkGroupTypeWise 
      BackColor       =   &H00FF8080&
      Caption         =   "Group Type Wise"
      Height          =   255
      Left            =   6780
      TabIndex        =   36
      Top             =   7560
      Width           =   1875
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5235
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2325
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.CheckBox ChkOpening 
      BackColor       =   &H00FF8080&
      Caption         =   "Include Opening"
      Height          =   255
      Left            =   6780
      TabIndex        =   22
      Top             =   7230
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.TextBox TxtAccountNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5235
      MaxLength       =   10
      TabIndex        =   0
      Top             =   5940
      Width           =   1020
   End
   Begin VB.TextBox TxtaccountName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6615
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5940
      Width           =   3585
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   5760
      TabIndex        =   3
      Top             =   7950
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Preview"
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
      MICON           =   "RptLedgerNew.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7080
      TabIndex        =   4
      Top             =   7950
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
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
      MICON           =   "RptLedgerNew.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8415
      TabIndex        =   5
      Top             =   7950
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
      MICON           =   "RptLedgerNew.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSearch 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6255
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5940
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "RptLedgerNew.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6210
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3225
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
      MICON           =   "RptLedgerNew.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   5190
      TabIndex        =   7
      Top             =   3225
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   2
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
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   6570
      TabIndex        =   8
      Tag             =   "nc"
      Top             =   3225
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   6195
      TabIndex        =   1
      Top             =   6645
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
      Left            =   7920
      TabIndex        =   2
      Top             =   6645
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
   Begin JeweledBut.JeweledButton BtnZone 
      Height          =   330
      Left            =   6210
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3825
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
      MICON           =   "RptLedgerNew.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtzoneID 
      Height          =   315
      Left            =   5190
      TabIndex        =   9
      Top             =   3825
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
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
   Begin SITextBox.Txt TxtZoneName 
      Height          =   315
      Left            =   6570
      TabIndex        =   25
      Top             =   3825
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
   Begin JeweledBut.JeweledButton BtnSector 
      Height          =   330
      Left            =   6210
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4470
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
      MICON           =   "RptLedgerNew.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSectorID 
      Height          =   315
      Left            =   5190
      TabIndex        =   11
      Top             =   4470
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
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
   Begin SITextBox.Txt TxtSectorName 
      Height          =   315
      Left            =   6570
      TabIndex        =   27
      Top             =   4470
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
   Begin JeweledBut.JeweledButton BtnSession 
      Height          =   330
      Left            =   6195
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5205
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
      MICON           =   "RptLedgerNew.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSessionID 
      Height          =   315
      Left            =   5190
      TabIndex        =   13
      Top             =   5205
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
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
   Begin SITextBox.Txt TxtSessionName 
      Height          =   315
      Left            =   6555
      TabIndex        =   33
      Top             =   5205
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Session Name"
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
      Left            =   6555
      TabIndex        =   35
      Top             =   4995
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Session ID"
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
      TabIndex        =   34
      Top             =   4995
      Width           =   930
   End
   Begin VB.Label Label7 
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
      Left            =   5190
      TabIndex        =   31
      Top             =   4260
      Width           =   825
   End
   Begin VB.Label Label6 
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
      Left            =   6570
      TabIndex        =   30
      Top             =   4260
      Width           =   1110
   End
   Begin VB.Label Label3 
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
      Left            =   5175
      TabIndex        =   29
      Top             =   3615
      Width           =   705
   End
   Begin VB.Label Label5 
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
      Left            =   6555
      TabIndex        =   28
      Top             =   3615
      Width           =   990
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
      Left            =   6195
      TabIndex        =   20
      Top             =   6420
      Width           =   885
   End
   Begin VB.Label Label1 
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
      Left            =   7950
      TabIndex        =   19
      Top             =   6420
      Width           =   705
   End
   Begin VB.Label LblOrganizationID 
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
      TabIndex        =   18
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label LblOrganizationName 
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
      TabIndex        =   17
      Top             =   3000
      Width           =   1620
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ledger"
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
      TabIndex        =   16
      Top             =   270
      UseMnemonic     =   0   'False
      Width           =   1005
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11093
      Top             =   2581
      Width           =   330
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5235
      TabIndex        =   15
      Top             =   5730
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6615
      TabIndex        =   14
      Top             =   5730
      Width           =   1335
   End
End
Attribute VB_Name = "RptLedgerNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Application1 As New CRAXDRT.Application
Dim vStrComp As String, vCompanyName As String, vAddress As String, vEmail As String
Dim vSQL As String, vStrSQL As String

Private Sub BtnZone_Click()
   If FunSelectZone(ssButton, False) = True Then
      TxtSectorID.SetFocus
   Else
      TxtzoneID.SetFocus
   End If
End Sub

Private Function FunSelectZone(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchZone.Show vbModal, Me
        If SchZone.ParaOutZoneID = "" Then FunSelectZone = False: Exit Function
        TxtzoneID.Text = SchZone.ParaOutZoneID
    End If
    '---------------------------
    If Trim(TxtzoneID.Text) = "" Then Exit Function
    If InStr(1, TxtzoneID.Text, ",") > 0 Then TxtZoneName.Text = "Selected Zones": Exit Function
    vStrSQL = " Select * FROM Zones where ZoneID=" & Val(TxtzoneID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtZoneName.Text = !ZoneName
          FunSelectZone = True
          .Close
          Exit Function
      Else
          FunSelectZone = False
          .Close
          TxtzoneID.Text = ""
          TxtZoneName.Text = "All Zones"
          End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtZoneID_Change()
   If ActiveControl.Name <> TxtzoneID.Name Then Exit Sub
   If TxtZoneName.Text <> "All Zones" Then TxtZoneName.Text = "All Zones"
End Sub

Private Sub TxtZoneID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtzoneID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtZoneName.Text <> "All Zones" Then Exit Sub
   If Trim(TxtzoneID.Text) = "" Then Exit Sub
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

Private Sub BtnSector_Click()
   If FunSelectSector(ssButton, False) = True Then
      TxtSessionID.SetFocus
   Else
      TxtSectorID.SetFocus
   End If
End Sub

Private Function FunSelectSector(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSector.Show vbModal, Me
        If SchSector.ParaOutSectorID = "" Then FunSelectSector = False: Exit Function
        TxtSectorID.Text = SchSector.ParaOutSectorID
    End If
    '---------------------------
    If Trim(TxtSectorID.Text) = "" Then Exit Function
    If InStr(1, TxtSectorID.Text, ",") > 0 Then TxtSectorName.Text = "Selected Sectors": Exit Function
    vStrSQL = "Select * FROM Sectors s where SectorID=" & Val(TxtSectorID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSectorName.Text = !SectorName
          FunSelectSector = True
          .Close
          Exit Function
      Else
          FunSelectSector = False
          .Close
          TxtSectorID.Text = ""
          TxtSectorName.Text = "All Sectors"
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtSectorID_Change()
   If ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   If TxtSectorName.Text <> "All Sectors" Then TxtSectorName.Text = "All Sectors"
End Sub

Private Sub TxtSectorID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSectorName.Text <> "All Sectors" Then Exit Sub
   If Trim(TxtSectorID.Text) = "" Then Exit Sub
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

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and c.isDetailed = 1 and c.isLocked = 0"
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
        TxtAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    If Trim(TxtAccountNo.Text) = "" Then Exit Function
    
   vStrSQL = " Select c.AccountNo, c.AccountName FROM ChartofAccounts c " & vbCrLf & _
     " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
     " Left Outer join Members m on c.AccountNo = cast(m.Prefix as varchar(2))  + cast(m.MemberID as varchar(10)) " & vbCrLf & _
     " where p.BarCode = '" & (TxtAccountNo.Text) & "' or m.BarCode = '" & (TxtAccountNo.Text) & "' or (c.AccountNo = '" & (TxtAccountNo.Text) & "' and c.isDetailed = 1 and c.isLocked = 0)"

    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtAccountNo.Text = !AccountNo
          TxtaccountName.Text = !AccountName
          With CN.Execute("Select Address, City from Parties where PartyID = '" & TxtAccountNo.Text & "'")
            If .RecordCount > 0 Then
               TxtAddress.Text = IIf(IsNull(!Address) = True, "", !Address)
               TxtCity.Text = IIf(IsNull(!City) = True, "", !City)
            Else
               TxtAddress.Text = ""
               TxtCity.Text = ""
            End If
          End With
          FunSelectAccount = True
          Exit Function
      Else
          FunSelectAccount = False
'          TxtAccountNo.Text = ""
          TxtaccountName.Text = ""
          TxtAddress.Text = ""
          TxtCity.Text = ""
      End If
      .Close
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnOrganization_Click()
   If FunSelectOrganization(ssButton, False) = True Then
      TxtzoneID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnPreview_Click()
  On Error GoTo ErrorHandler
  If FunRefreshData = False Then Exit Sub
  If Rs.RecordCount = 0 Then
    MsgBox "No record found", vbInformation, "Information"
    Exit Sub
  Else
    Call SetCrystalReport
    RptReportViewer.Show vbModal, Me
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
  On Error GoTo ErrorHandler
  If FunRefreshData = False Then Exit Sub
  If Rs.RecordCount = 0 Then
    MsgBox "No record found", vbInformation, "Information"
    Exit Sub
  Else
    Call SetCrystalReport
    RptReportViewer.Report.PrintOut
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, True) = True Then TxtzoneID.SetFocus
         Case TxtzoneID.Name: If FunSelectZone(ssFunctionKey, True) = True Then TxtSectorID.SetFocus
         Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, True) = True Then TxtAccountNo.SetFocus
         Case TxtAccountNo.Name: If FunSelectAccount(ssFunctionKey, True) = True Then DtpFrom.SetFocus
      End Select
   End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtAccountNo_Change()
   If TxtAccountNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   If TxtaccountName.Text <> "" Then TxtaccountName.Text = ""
End Sub

Private Sub TxtAccountNo_Validate(Cancel As Boolean)
   Dim vTemp As Boolean
   If Trim(TxtAccountNo.Text) = "" Then Exit Sub
   vTemp = Not FunSelectAccount(ssValidate, True)
'    If vTemp = True Then
'        vTemp = Not FunSelectAccount(ssButton, False)
'    End If
    Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSearch_Click()
   On Error GoTo ErrorHandler
   If FunSelectAccount(ssButton, True) = True Then
      DtpFrom.SetFocus
   Else
      TxtAccountNo.SetFocus
   End If
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function FunRefreshData() As Boolean
   On Error GoTo ErrorHandler
   Dim vSQL As String, i As Integer
   Me.MousePointer = vbHourglass
   
   
    vSQL = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[#AccountsLedger]') and OBJECTPROPERTY(id, N'IsTable') = 1)" & vbCrLf & _
     "drop Table [dbo].[#AccountsLedger]"
      CN.Execute vSQL
'     CN.Execute "drop Table [dbo].[#AccountsLedger]"
      
    
    vSQL = " CREATE TABLE [dbo].[#AccountsLedger] (" & vbCrLf & _
      " [organizationID] [tinyint] NULL ," & vbCrLf & _
      " [AccountNo] [varchar] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & vbCrLf & _
      " [VoucherType] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & vbCrLf & _
      " [VoucherNo] [int] NULL ," & vbCrLf & _
      " [StrVoucherNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & vbCrLf & _
      " [VoucherDate] [smalldatetime] NULL ," & vbCrLf & _
      " [Debit] [numeric](12, 2) NULL ," & vbCrLf & _
      " [Credit] [numeric](12, 2) NULL ," & vbCrLf & _
      " [Naration] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & vbCrLf & _
      " [EntryTime] [datetime] NULL ," & vbCrLf & _
      " [SessionID] [smallint] NULL" & vbCrLf & _
      ") ON [PRIMARY]"

   CN.Execute vSQL
     
   'CN.Execute "Delete From #AccountsLedger"
   vSQL = "Select c.* from ChartofAccounts c " & vbCrLf & _
         " left outer join Parties p on p.PartyID = c.AccountNo " & vbCrLf & _
         " left outer join Sectors s on s.SectorID = p.SectorID " & vbCrLf & _
         " left outer join Zones t on t.ZoneID = s.ZoneID " & vbCrLf & _
         " Where 1=1 " & vbCrLf & _
        IIf(Trim(TxtAccountNo.Text) = "", "", " and c.AccountNo = '" & TxtAccountNo.Text & "'") & vbCrLf & _
        IIf(Trim(TxtSectorID.Text) = "", "", " and p.sectorid in (" & TxtSectorID.Text & ")") & vbCrLf & _
        IIf(Trim(TxtzoneID.Text) = "", "", " and t.Zoneid in (" & TxtzoneID.Text & ")")

   With CN.Execute(vSQL)
      While Not .EOF
         vSQL = "EXECUTE SPAccountsLedgerNew " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & ",'" & !AccountNo & "', '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & ChkOpening.Value
         'vSQL = "EXECUTE SPAccountsLedger '" & !AccountNo & "', '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
         CN.Execute vSQL
         .MoveNext
      Wend
   End With
   
   CN.Execute "Delete From #AccountsLedger Where Debit = 0 And Credit = 0"
   
   If ChkZeroBalance.Value = 0 Then
      CN.Execute "Delete From #AccountsLedger Where AccountNo in (select accountno from #AccountsLedger Group By Accountno having sum(debit) = sum(credit))"
   End If
'   CN.Execute "Delete From AccountsLedger"
'   CN.Execute "Insert into Accountsledger Select * from #AccountsLedger"
   
   vSQL = "SELECT a.organizationID, OrganizationName, ca.AccountNo,  str(ca.AccountNo) + '-' + ca.AccountName + ' ' + isnull(P.Address,'') + ' '  + isnull(p.phone1,'') + ' ' + isnull(p.phone2,'') + ' ' + isnull(p.Mobile,'') + isnull(' '+p.Mobile2,'') + ' ' + isnull(P.Description,'') as AccountName," & vbCrLf & _
        " s.ZoneID, ZoneName, p.SectorID, SectorName, a.vouchertype, a.voucherno, a.strvoucherno,voucherdate, a.debit, a.credit, a.naration as narration " & vbCrLf & _
        " FROM #AccountsLedger a INNER JOIN ChartOfAccounts ca ON a.AccountNo = ca.AccountNo " & vbCrLf & _
        " left outer JOIN Organizations o ON  o.organizationID = a.organizationID " & vbCrLf & _
        " left outer JOIN Parties p ON  p.PartyID = ca.AccountNo " & vbCrLf & _
        " left outer join sectors s on p.sectorid = s.sectorid " & vbCrLf & _
        " left outer join Zones t on t.Zoneid = s.Zoneid " & vbCrLf & _
        " Where 1=1 " & vbCrLf & _
        IIf(Trim(TxtSectorID.Text) = "", "", " and p.sectorid in (" & TxtSectorID.Text & ")") & vbCrLf & _
        IIf(Trim(TxtzoneID.Text) = "", "", " and t.Zoneid in (" & TxtzoneID.Text & ")") & vbCrLf & _
        " order by EntryTime"
  
   Set Rs = CN.Execute(vSQL)
   vSQL = " Drop TABLE [dbo].[#AccountsLedger] "
   CN.Execute vSQL
   Me.MousePointer = vbDefault
'   Set Rs = CN.Execute("EXECUTE SPAccountsLedgerNew " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & ",'" & TxtAccountNo.Text & "', '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & ChkOpening.Value)
   FunRefreshData = True
   Exit Function
ErrorHandler:
'   If Err.Number = -2147217900 Then
'      Resume Next
'   End If
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
   FunRefreshData = False
End Function

Private Function FunRefreshData1() As Boolean
   On Error GoTo ErrorHandler
   Dim i As Integer
   Me.MousePointer = vbHourglass
   'CN.Execute "Delete From AccountsLedger"
   
'   vSQL = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[#AccountsLedger]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'         "delete from [dbo].[#AccountsLedger]"
'   CN.Execute vSQL
   
'   vSQL = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[#AccountsLedger]') and OBJECTPROPERTY(id, N'IsTable') = 1)" & vbCrLf & _
'      "drop Table [#AccountsLedger]"
'   vSQL = "drop Table [#AccountsLedger]"
'   CN.Execute vSQL
   
   'CN.Execute "Delete From #AccountsLedger"
   'CN.Execute "Delete From #AccountsLedger1"
   
   vSQL = " CREATE TABLE [dbo].[#AccountsLedger] (" & vbCrLf & _
      " [organizationID] [tinyint] NULL ," & vbCrLf & _
      " [AccountNo] [varchar] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & vbCrLf & _
      " [VoucherType] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & vbCrLf & _
      " [VoucherNo] [int] NULL ," & vbCrLf & _
      " [StrVoucherNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & vbCrLf & _
      " [VoucherDate] [smalldatetime] NULL ," & vbCrLf & _
      " [Debit] [numeric](12, 2) NULL ," & vbCrLf & _
      " [Credit] [numeric](12, 2) NULL ," & vbCrLf & _
      " [Naration] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & vbCrLf & _
      " [EntryTime] [datetime] NULL ," & vbCrLf & _
      " [SessionID] [smallint] NULL" & vbCrLf & _
      ") ON [PRIMARY]"

   CN.Execute vSQL
   vSQL = "Select c.* from ChartofAccounts c " & vbCrLf & _
         " left outer join Parties p on p.PartyID = c.AccountNo " & vbCrLf & _
         " left outer join Sectors s on s.SectorID = p.SectorID " & vbCrLf & _
         " left outer join Zones t on t.ZoneID = s.ZoneID " & vbCrLf & _
         " Where 1=1 " & vbCrLf & _
        IIf(Trim(TxtAccountNo.Text) = "", "", " and c.AccountNo = '" & TxtAccountNo.Text & "'") & vbCrLf & _
        IIf(Trim(TxtSectorID.Text) = "", "", " and p.sectorid in (" & TxtSectorID.Text & ")") & vbCrLf & _
        IIf(Trim(TxtzoneID.Text) = "", "", " and t.Zoneid in (" & TxtzoneID.Text & ")")

   With CN.Execute(vSQL)
      While Not .EOF
         vSQL = "EXECUTE SPAccountsLedgerNew " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & ",'" & !AccountNo & "', '" & DtpTo.DateValue - 150 & "','" & DtpTo.DateValue & "'," & ChkOpening.Value
         'vSQL = "EXECUTE SPAccountsLedger '" & !AccountNo & "', '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
         CN.Execute vSQL
         .MoveNext
      Wend
   End With
   
   'CN.Execute "Delete From #AccountsLedger1"
   
   'CN.Execute "insert into #AccountsLedger select * From #AccountsLedger"
   
   CN.Execute "Delete From #AccountsLedger Where Debit = 0 And Credit = 0"
   
   If ChkZeroBalance.Value = 0 Then
      CN.Execute "Delete From #AccountsLedger Where AccountNo in (select accountno from #AccountsLedger Group By Accountno having sum(debit) = sum(credit))"
   End If
   
   vSQL = "SELECT a.organizationID, OrganizationName, ca.AccountNo,  str(ca.AccountNo) + '-' + ca.AccountName + ' ' + isnull(P.Address,'') + ' '  + isnull(p.phone1,'') + ' ' + isnull(p.phone2,'') + ' ' + isnull(p.Mobile,'') + isnull(' '+p.Mobile2,'') + ' ' + isnull(P.Description,'') as AccountName," & vbCrLf & _
        " s.ZoneID, ZoneName, p.SectorID, SectorName, a.vouchertype, a.voucherno, a.strvoucherno,voucherdate, a.debit, a.credit, a.naration as narration " & vbCrLf & _
        " FROM #AccountsLedger a INNER JOIN ChartOfAccounts ca ON a.AccountNo = ca.AccountNo " & vbCrLf & _
        " left outer JOIN Organizations o ON  o.organizationID = a.organizationID " & vbCrLf & _
        " left outer JOIN Parties p ON  p.PartyID = ca.AccountNo " & vbCrLf & _
        " left outer join sectors s on p.sectorid = s.sectorid " & vbCrLf & _
        " left outer join Zones t on t.Zoneid = s.Zoneid " & vbCrLf & _
        " Where 1=1 " & vbCrLf & _
        IIf(Trim(TxtSectorID.Text) = "", "", " and p.sectorid in (" & TxtSectorID.Text & ")") & vbCrLf & _
        IIf(Trim(TxtzoneID.Text) = "", "", " and t.Zoneid in (" & TxtzoneID.Text & ")") & vbCrLf & _
        " order by VoucherDate"
  
   Set Rs1 = CN.Execute(vSQL)
   vSQL = " Drop TABLE [dbo].[#AccountsLedger]"
   CN.Execute vSQL
   Me.MousePointer = vbDefault
   'Set Rs = CN.Execute("EXECUTE SPAccountsLedgerNew " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & ",'" & TxtAccountNo.Text & "', '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & ChkOpening.Value)
   FunRefreshData1 = True
   Exit Function
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
   FunRefreshData1 = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
   Call FunRefreshData1
   If ChkGroupTypeWise.Value = 1 Then
      Set RptReportViewer.Report = New CrpAccountLedgerGroupTypeWise
   Else
'      Set RptReportViewer.Report = New CrpAccountLedger
       Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\AccountReports\CrpAccountLedger.rpt")
   End If
   
   'this code works through the RDC object model to identify a subreport object
   'in the main report
   Dim crSecs As CRAXDRT.Sections
   Dim crSec As CRAXDRT.Section
   Dim crRepObjs As CRAXDRT.ReportObjects
   Dim crSubRepObj As CRAXDRT.SubreportObject
   Dim crSubReport As CRAXDRT.Report
   Dim i As Integer
   Dim x As Integer

   Set crSecs = RptReportViewer.Report.Sections
   For i = 1 To crSecs.Count
     Set crSec = crSecs.Item(i)
     Set crRepObjs = crSec.ReportObjects
       For x = 1 To crRepObjs.Count
         If crRepObjs.Item(x).Kind = crSubreportObject Then
            If x = 6 And i = 9 Then
               Set crSubReport = RptReportViewer.Report.OpenSubreport(crRepObjs.Item(x).SubreportName)
               'the following code sets the subreport table to a different database
               crSubReport.Database.SetDataSource Rs1, 3, 1
'               crSubReport.Database.SetDataSource RsReport2, 3, 1
               'set the value for a text object in the header of the subreport
               'CRReport.Subreport1_Text2.SetText "This is the subreport"
               'within this loop you can set other properties of the subreport and
               'the field objects and sections in it.
            End If
         End If
       Next
   Next
   

  'RptReportViewer.Report.TxtCompanyName.SetText CN.Execute("select companyname from Project_Registry").Fields(0).Value
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ReportTitle = "Ledger Report"
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue "Account : " & TxtAccountNo.Text & "/" & TxtaccountName.Text & IIf(TxtAddress.Text = "", "", " (" & TxtAddress.Text & ")") & IIf(TxtCity.Text = "", "", vbCrLf & TxtCity.Text & ".")
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue "From Date " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To " & Format(DtpTo.DateValue, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue Trim(TxtOrganizationID.Text)
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(6).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity & ".") & IIf(ObjRegistry.CompanyPhoneNo = "", "", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.PaperOrientation = crPortrait
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Ledger"
   
   TxtOrganizationID.Text = IIf(ObjRegistry.OrganizationID = "Null", "", ObjRegistry.OrganizationID)
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
  
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtOrganizationID_Change()
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizationName.Text <> "" Then TxtOrganizationName.Text = ""
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
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganization = False: Exit Function
       TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    If TxtOrganizationID.Text = "" Then FunSelectOrganization = False: Exit Function
    vStrSQL = " Select * FROM Organizations where OrganizationID='" & TxtOrganizationID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizationName.Text = !OrganizationName
          FunSelectOrganization = True
          .Close
          Exit Function
      Else
          FunSelectOrganization = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizationName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnSession_Click()
   If FunSelectSession(ssButton, False) = True Then
      TxtAccountNo.SetFocus
   Else
      TxtSessionID.SetFocus
   End If
End Sub

Private Function FunSelectSession(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSession.Show vbModal, Me
        If SchSession.ParaOutSessionID = "" Then FunSelectSession = False: Exit Function
        TxtSessionID.Text = SchSession.ParaOutSessionID
    End If
    '---------------------------
    If Trim(TxtSessionID.Text) = "" Then Exit Function
    If InStr(1, TxtSessionID.Text, ",") > 0 Then TxtSessionName.Text = "Selected Sessions": Exit Function
    vStrSQL = "Select * FROM Sessions s where SessionID=" & Val(TxtSessionID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSessionName.Text = !SessionName
          FunSelectSession = True
          .Close
          Exit Function
      Else
          FunSelectSession = False
          .Close
          TxtSessionID.Text = ""
          TxtSessionName.Text = "All Sessions"
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtSessionID_Change()
   If ActiveControl.Name <> TxtSessionID.Name Then Exit Sub
   If TxtSessionName.Text <> "All Sessions" Then TxtSessionName.Text = "All Sessions"
End Sub

Private Sub TxtSessionID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSessionID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSessionName.Text <> "All Sessions" Then Exit Sub
   If Trim(TxtSessionID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSession(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSession(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

