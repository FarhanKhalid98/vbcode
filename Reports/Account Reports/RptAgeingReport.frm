VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptAgeingReport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptAgeingReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Left            =   3765
      TabIndex        =   47
      Top             =   3585
      Width           =   3585
   End
   Begin VB.TextBox TxtZoneID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2385
      TabIndex        =   3
      Top             =   3585
      Width           =   1020
   End
   Begin VB.TextBox TxtSectorID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2385
      TabIndex        =   4
      Top             =   4380
      Width           =   1020
   End
   Begin VB.TextBox TxtSectorName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3765
      TabIndex        =   46
      Top             =   4380
      Width           =   3585
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8685
      TabIndex        =   42
      Top             =   2655
      Width           =   3495
      Begin VB.OptionButton RdoPaid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Paid"
         Height          =   255
         Left            =   180
         TabIndex        =   45
         Top             =   10
         Width           =   885
      End
      Begin VB.OptionButton RdoUnPaid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Not Paid"
         Height          =   255
         Left            =   1125
         TabIndex        =   44
         Top             =   10
         Width           =   1140
      End
      Begin VB.OptionButton RdoBoth 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Both"
         Height          =   255
         Left            =   2385
         TabIndex        =   43
         Top             =   10
         Value           =   -1  'True
         Width           =   930
      End
   End
   Begin VB.CheckBox ChkOpening 
      BackColor       =   &H00FF8080&
      Caption         =   "Include Opening"
      Height          =   255
      Left            =   9465
      TabIndex        =   40
      Top             =   3210
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.ComboBox CmbVoucherType 
      Height          =   315
      Left            =   10680
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   3615
      Width           =   780
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   10605
      TabIndex        =   36
      Top             =   2220
      Width           =   2250
      Begin VB.OptionButton RdoDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detail"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   10
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton RdoSummary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary"
         Height          =   255
         Left            =   945
         TabIndex        =   37
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   8220
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   2220
      Width           =   2355
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8100
      TabIndex        =   11
      Top             =   6660
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
      MICON           =   "RptAgeingReport.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5265
      TabIndex        =   9
      Top             =   6660
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
      MICON           =   "RptAgeingReport.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6705
      TabIndex        =   10
      Top             =   6660
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
      MICON           =   "RptAgeingReport.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3525
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   8745
      Visible         =   0   'False
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
      MICON           =   "RptAgeingReport.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   8865
      TabIndex        =   7
      Top             =   4260
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
      Left            =   10620
      TabIndex        =   8
      Top             =   4260
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
   Begin JeweledBut.JeweledButton BtnUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   10290
      Visible         =   0   'False
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
      MICON           =   "RptAgeingReport.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtUserName 
      Height          =   315
      Left            =   3840
      TabIndex        =   19
      Tag             =   "nc"
      Top             =   10290
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   2505
      TabIndex        =   1
      Top             =   8745
      Visible         =   0   'False
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
   Begin JeweledBut.JeweledButton BtnCustomer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3345
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2865
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
      MICON           =   "RptAgeingReport.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   3705
      TabIndex        =   15
      Tag             =   "nc"
      Top             =   2865
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
   Begin JeweledBut.JeweledButton BtnEmployee 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3390
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5160
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
      MICON           =   "RptAgeingReport.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmpName 
      Height          =   315
      Left            =   3750
      TabIndex        =   17
      Tag             =   "nc"
      Top             =   5160
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
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   3885
      TabIndex        =   13
      Tag             =   "nc"
      Top             =   8745
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtUserNo 
      Height          =   315
      Left            =   2460
      TabIndex        =   6
      Top             =   10290
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   315
      Left            =   2340
      TabIndex        =   2
      Top             =   2880
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   11
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
   Begin SITextBox.Txt TxtEmpID 
      Height          =   315
      Left            =   2370
      TabIndex        =   5
      Top             =   5160
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   5
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
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2100
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
      MICON           =   "RptAgeingReport.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   2340
      TabIndex        =   0
      Top             =   2100
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
      Left            =   3720
      TabIndex        =   32
      Tag             =   "nc"
      Top             =   2100
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
   Begin JeweledBut.JeweledButton BtnZone 
      Height          =   330
      Left            =   3420
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   3585
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
      MICON           =   "RptAgeingReport.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSector 
      Height          =   330
      Left            =   3420
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   4380
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
      MICON           =   "RptAgeingReport.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label10 
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
      Left            =   3780
      TabIndex        =   53
      Top             =   3375
      Width           =   990
   End
   Begin VB.Label Label9 
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
      Left            =   2400
      TabIndex        =   52
      Top             =   3375
      Width           =   705
   End
   Begin VB.Label Label4 
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
      Left            =   2400
      TabIndex        =   51
      Top             =   4170
      Width           =   825
   End
   Begin VB.Label Label3 
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
      Left            =   3780
      TabIndex        =   50
      Top             =   4170
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VoucherType"
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
      Left            =   9450
      TabIndex        =   41
      Top             =   3645
      Width           =   1140
   End
   Begin VB.Label Label21 
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
      Left            =   3720
      TabIndex        =   34
      Top             =   1875
      Width           =   1590
   End
   Begin VB.Label Label20 
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
      Left            =   2340
      TabIndex        =   33
      Top             =   1875
      Width           =   1290
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   868
      X2              =   536
      Y1              =   321
      Y2              =   321
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   865
      X2              =   865
      Y1              =   136
      Y2              =   321
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   536
      X2              =   865
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   538
      X2              =   538
      Y1              =   136
      Y2              =   321
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
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
      Left            =   3900
      TabIndex        =   30
      Top             =   8550
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
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
      Left            =   2370
      TabIndex        =   29
      Top             =   4950
      Width           =   1080
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
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
      Left            =   3765
      TabIndex        =   28
      Top             =   4950
      Width           =   1365
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
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
      Left            =   2325
      TabIndex        =   27
      Top             =   2685
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
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
      Left            =   3720
      TabIndex        =   26
      Top             =   2685
      Width           =   1335
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
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
      Left            =   2550
      TabIndex        =   25
      Top             =   8505
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   3855
      TabIndex        =   24
      Top             =   10095
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Left            =   2505
      TabIndex        =   23
      Top             =   10050
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label6 
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
      Left            =   10635
      TabIndex        =   22
      Top             =   4035
      Width           =   705
   End
   Begin VB.Label Label5 
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
      Left            =   8865
      TabIndex        =   21
      Top             =   4035
      Width           =   885
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ageing Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2700
      TabIndex        =   20
      Top             =   270
      Width           =   1650
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
Attribute VB_Name = "RptAgeingReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim RsCredit As New ADODB.Recordset
Dim RsDebit As New ADODB.Recordset
Dim vSQL As String
Dim vCounter As Integer
Dim vPaidRdo As String
Dim vOverAmount, vAccountNo, vRecoveryAmount, vDifferenceAmount As Double

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
      If RdoDetail.Value = True Then
         RptReportViewer.Caption = "Detail (" & CmbGroup.Text & ")"
      Else
         RptReportViewer.Caption = "Summary (" & CmbGroup.Text & ")"
      End If
      RptReportViewer.Show vbModal
   End If
End Sub

Private Sub BtnPrint_Click()
   If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub BtnZone_Click()
   If FunSelectZone(ssButton, False) = True Then
      TxtSectorID.SetFocus
   Else
      TxtZoneID.SetFocus
   End If

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
        Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, True) = True Then TxtStoreID.SetFocus
        Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then TxtCustomerID.SetFocus
        Case TxtCustomerID.Name: If FunSelectCustomer(ssFunctionKey, True) = True Then TxtEmpID.SetFocus
        Case TxtZoneID.Name: If FunSelectZone(ssFunctionKey, True) = True Then TxtSectorID.SetFocus
        Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, True) = True Then TxtEmpID.SetFocus
        Case TxtEmpID.Name: If FunSelectEmployee(ssFunctionKey, True) = True Then BtnPreview.SetFocus
        Case TxtEmpID.Name: If FunSelectEmployee(ssFunctionKey, True) = True Then TxtUserNo.SetFocus
        Case TxtUserNo.Name: If FunSelectUser(ssFunctionKey, True) = True Then BtnPreview.SetFocus
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
   SetWindowText Me.hWnd, "Ageing Report"
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   
   CmbVoucherType.AddItem ("All")
   CmbVoucherType.AddItem ("SI")
   
   CmbVoucherType.ListIndex = 0
   
   CmbGroup.AddItem ("Sale Ageing Report")
   CmbGroup.AddItem ("Purchase Ageing Report")
   CmbGroup.AddItem ("Ageing Report")
   
      
   CmbGroup.ListIndex = 0
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
   'Set RsReport = Nothing
   Set RptRecoveryRegister = Nothing
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
    
   If RdoDetail.Value = True Then
      Select Case CmbGroup.Text
         Case "Ageing Report"
            Set RptReportViewer.Report = New CrpAgeingReport
            vSQL = "EXEC ProdRptAgeing'" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(Trim(TxtCustomerID.Text) = "", "Null", "'" & TxtCustomerID.Text & "'") & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", "'" & TxtStoreID.Text & "'") & "," & IIf(Trim(TxtEmpID.Text) = "", "Null", "'" & TxtEmpID.Text & "'") & "," & IIf(Trim(TxtUserNo.Text) = "", "Null", "'" & TxtUserNo.Text & "'")
            Set RsReport = CN.Execute(vSQL)
            RptReportViewer.Report.PaperOrientation = crPortrait
         Case "Sale Ageing Report"
            AgeingCalculate
            Set RptReportViewer.Report = New CrpSaleAgeingReport
'            vSQL = "EXEC ProdRptSaleAgeing'" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtCustomerID.Text) = "", "Null", "'" & TxtCustomerID.Text & "'") & "," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(CmbVoucherType.ListIndex = 0, "Null", "'" & CmbVoucherType.Text & "'") & "," & IIf(RdoBoth.Value = True, "Null", IIf(RdoPaid.Value = True, 1, 0))
            vSQL = "EXEC ProdRptSaleAgeing'" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtCustomerID.Text) = "", "Null", "'" & TxtCustomerID.Text & "'") & "," & vbCrLf _
               & IIf(Trim(TxtZoneID.Text) = "", "Null", TxtZoneID.Text) & "," & vbCrLf _
               & IIf(Trim(TxtSectorID.Text) = "", "Null", TxtSectorID.Text) & "," & vbCrLf _
               & IIf(Trim(TxtEmpID.Text) = "", "Null", TxtEmpID.Text) & "," & vbCrLf _
               & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(CmbVoucherType.ListIndex = 0, "Null", "'" & CmbVoucherType.Text & "'") & "," & IIf(RdoBoth.Value = True, "Null", IIf(RdoPaid.Value = True, 1, 0))
            Set RsReport = CN.Execute(vSQL)
            RptReportViewer.Report.PaperOrientation = crLandscape
         Case "Purchase Ageing Report"
            AgeingCalculate
            Set RptReportViewer.Report = New CrpPurchaseAgeingReport
'            vSQL = "EXEC ProdRptPurchaseAgeing'" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtCustomerID.Text) = "", "Null", "'" & TxtCustomerID.Text & "'") & "," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(CmbVoucherType.ListIndex = 0, "Null", "'" & CmbVoucherType.Text & "'") & "," & IIf(RdoBoth.Value = True, "Null", IIf(RdoPaid.Value = True, 1, 0))
            vSQL = "EXEC ProdRptPurchaseAgeing'" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "',Null," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(CmbVoucherType.ListIndex = 0, "Null", "'" & CmbVoucherType.Text & "'") & "," & IIf(RdoBoth.Value = True, "Null", IIf(RdoPaid.Value = True, 1, 0))
            Set RsReport = CN.Execute(vSQL)
            RptReportViewer.Report.PaperOrientation = crLandscape
      End Select
   Else
      Select Case CmbGroup.Text
         Case "Ageing Report"
            Set RptReportViewer.Report = New CrpAgeingReport
         Case "Sale Ageing Report"
            Set RptReportViewer.Report = New CrpSaleAgeingReport
      End Select
   End If
   
   If RsReport.BOF Then
      MsgBox "No record exists.", vbInformation, Me.Caption
      Me.MousePointer = vbDefault
      Exit Function
   End If
   
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport

'   RptReportViewer.Report.ReportTitle = "Ageing Report"
   
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", ".", ", E.Mail : " & ObjRegistry.CompanyEMail)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue " Date From :" & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To : " & Format(DtpTo.DateValue, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
'   RptReportViewer.Report.PaperOrientation = crPortrait
      
'    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    'RptReportViewer.Report.PaperOrientation = crLandscape
   SetReport = True
   Me.MousePointer = vbDefault
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnUser_Click()
   If FunSelectUser(ssButton, False) = True Then
      BtnPreview.SetFocus
   Else
      TxtUserNo.SetFocus
   End If
End Sub

Private Function FunSelectUser(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchUser.Show vbModal, Me
        If SchUser.ParaOutUserNo = "" Then FunSelectUser = False: Exit Function
        TxtUserNo.Text = SchUser.ParaOutUserNo
    End If
    '---------------------------
    vStrSQL = " Select * FROM Users where UserNo=" & Val(TxtUserNo.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtUserName.Text = !UserName
          FunSelectUser = True
          .Close
          Exit Function
      Else
          FunSelectUser = False
          .Close
          TxtUserNo.Text = ""
          TxtUserName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function


Private Sub TxtUserNo_Change()
   If TxtUserNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtUserNo.Name Then Exit Sub
   If TxtUserName.Text <> "" Then TxtUserName.Text = ""
End Sub

Private Sub TxtUserNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtUserNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtUserNo.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectUser(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectUser(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub BtnEmployee_Click()
      If FunSelectEmployee(ssButton, False) = True Then
      BtnPreview.SetFocus
   Else
      TxtEmpID.SetFocus
   End If
End Sub

Private Function FunSelectEmployee(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchEmployee.Show vbModal, Me
        If SchEmployee.ParaOutEmployeeID = "" Then FunSelectEmployee = False: Exit Function
        TxtEmpID.Text = SchEmployee.ParaOutEmployeeID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Employees where EmpID=" & Val(TxtEmpID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtEmpName.Text = !EmpName
          FunSelectEmployee = True
          .Close
          Exit Function
      Else
          FunSelectEmployee = False
          .Close
          TxtEmpID.Text = ""
          TxtEmpName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtEmpID_Change()
   If TxtEmpID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtEmpID.Name Then Exit Sub
   If TxtEmpName.Text <> "" Then TxtEmpName.Text = ""
End Sub

Private Sub TxtEmpID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtEmpID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtEmpID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectEmployee(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectEmployee(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      TxtSectorID.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInDetail = ""
        SchAccounts.CmbFilter = "Customers"
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo like '3%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select c.* FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " where BarCode = '" & (TxtCustomerID.Text) & "' or (c.AccountNo = '" & (TxtCustomerID.Text) & "' and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo like '3%') and c.isDetailed = 1 and c.isLocked = 0)"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerID.Text = !AccountNo
          TxtCustomerName.Text = !AccountName
          FunSelectCustomer = True
          .Close
          Exit Function
      Else
          FunSelectCustomer = False
          .Close
          TxtCustomerID.Text = ""
          TxtCustomerName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtCustomerID_Change()
   If TxtCustomerID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   If TxtCustomerName.Text <> "" Then TxtCustomerName.Text = ""
End Sub

Private Sub TxtCustomerID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCustomer(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCustomer(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtCustomerID.SetFocus
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
    If Trim(TxtStoreID.Text) = "" Then Exit Function
    If TxtStoreID.Text = "" Then FunSelectStore = False: Exit Function
    vStrSQL = " Select StoreName FROM Stores where StoreID='" & TxtStoreID.Text & "'"
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

Private Sub TxtStoreID_Change()
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
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

Private Sub BtnOrganization_Click()
   If FunSelectOrganization(ssButton, True) = True Then
      TxtCustomerID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Function FunSelectOrganization(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganization = False: Exit Function
        TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Organizations where OrganizationID = " & Val(TxtOrganizationID.Text)
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
Private Sub AgeingCalculate()
On Error GoTo ErrorHandler
   
   CN.Execute "Delete Ageing"
   
   '''' EXECUTE Account Ledger
'   CN.Execute "Delete accountsledger"
'   vSQL = "EXECUTE SPAccountsLedgerNew " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & ",'" & TxtCustomerID.Text & "', '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & ChkOpening.Value
'   CN.Execute vSQL


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

'   CN.Execute "Delete accountsledger"
   If CmbGroup.Text = "Sale Ageing Report" Then
       vSQL = "Select c.* from ChartofAccounts c " & vbCrLf & _
         " left outer join Parties p on p.PartyID = c.AccountNo " & vbCrLf & _
         " left outer join Sectors s on s.SectorID = p.SectorID " & vbCrLf & _
         " left outer join Zones t on t.ZoneID = s.ZoneID " & vbCrLf & _
         " Where 1=1 And c.AccountNo like '62%' and c.AccountNo <> '621' " & vbCrLf & _
        IIf(Trim(TxtCustomerID.Text) = "", "", " and c.AccountNo = '" & TxtCustomerID.Text & "'") & vbCrLf & _
        IIf(Trim(TxtSectorID.Text) = "", "", " and p.sectorid in (" & TxtSectorID.Text & ")") & vbCrLf & _
        IIf(Trim(TxtZoneID.Text) = "", "", " and t.Zoneid in (" & TxtZoneID.Text & ")")

'      vSQL = "Select c.* from ChartofAccounts c  Where c.AccountNo like '62%' and c.AccountNo <> '621' and c.AccountNo <> '62'" & vbCrLf & _
'         IIf(Trim(TxtCustomerID.Text) = "", "", " and c.AccountNo = '" & TxtCustomerID.Text & "'")
   Else
      vSQL = "Select c.* from ChartofAccounts c " & vbCrLf & _
         " left outer join Parties p on p.PartyID = c.AccountNo " & vbCrLf & _
         " left outer join Sectors s on s.SectorID = p.SectorID " & vbCrLf & _
         " left outer join Zones t on t.ZoneID = s.ZoneID " & vbCrLf & _
         " Where 1=1 And c.AccountNo like '61%' and c.AccountNo <> '61' " & vbCrLf & _
        IIf(Trim(TxtCustomerID.Text) = "", "", " and c.AccountNo = '" & TxtCustomerID.Text & "'") & vbCrLf & _
        IIf(Trim(TxtSectorID.Text) = "", "", " and p.sectorid in (" & TxtSectorID.Text & ")") & vbCrLf & _
        IIf(Trim(TxtZoneID.Text) = "", "", " and t.Zoneid in (" & TxtZoneID.Text & ")")
'      vSQL = "Select c.* from ChartofAccounts c  Where c.AccountNo like '61%' and c.AccountNo <> '61'" & vbCrLf & _
'         IIf(Trim(TxtCustomerID.Text) = "", "", " and c.AccountNo = '" & TxtCustomerID.Text & "'")
   End If
   
      With CN.Execute(vSQL)
         While Not .EOF
            vSQL = "EXECUTE SPAccountsLedgerNew " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & ",'" & !AccountNo & "', '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & ChkOpening.Value
            CN.Execute vSQL
            vAccountNo = !AccountNo
            If CmbGroup.Text = "Sale Ageing Report" Then
               Call SaleAgeing
            Else
               Call PurchaseAgeing
            End If
            
            .MoveNext
            CN.Execute "Delete #accountsledger"
         Wend
      End With
   ''''''''''''''''''''''''''''''''''
     
'   If CmbGroup.Text = "Sale Ageing Report" Then
'      Call SaleAgeing
'   Else
'      Call PurchaseAgeing
'   End If
   vSQL = " Drop TABLE [dbo].[#AccountsLedger] "
   CN.Execute vSQL

   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SaleAgeing()
   On Error GoTo ErrorHandler
   
   '''' Get Credit
   If RsCredit.State = adStateOpen Then RsCredit.Close
   vSQL = "Select * from #accountsledger where Credit <> 0 And AccountNo = '" & vAccountNo & "' Order by VoucherDate Asc"
   RsCredit.Open vSQL, CN, adOpenStatic, adLockBatchOptimistic
   If Not RsCredit.EOF Then vOverAmount = RsCredit!Credit Else vOverAmount = 0
   ''''''''''''''''''''''''''''''''''
   
   '''' Get Debit
   If RsDebit.State = adStateOpen Then RsDebit.Close
   vSQL = "Select * from #accountsledger where Debit <> 0 And AccountNo = '" & vAccountNo & "' Order by VoucherDate Asc"
   RsDebit.Open vSQL, CN, adOpenStatic, adLockBatchOptimistic
   ''''''''''''''''''''''''''''''''''
   
'''' Checking Debit / Credit / OverAmount
   While Not RsDebit.EOF
      Do While Not RsCredit.EOF
         If vOverAmount > RsDebit!Debit Then
            vOverAmount = vOverAmount - RsDebit!Debit
            SaleAgeingInsert
            Exit Do
         Else
            RsCredit.MoveNext
            If Not RsCredit.EOF Then vOverAmount = vOverAmount + RsCredit!Credit
         End If
      Loop
      If RsCredit.EOF Then SaleAgeingInsert
      RsDebit.MoveNext
   Wend
   ''''''''''''''''''''''''''''''''''
   
   '''' Add Remainig credit Amount in OverAmount
   While Not RsCredit.EOF
      RsCredit.MoveNext
      If Not RsCredit.EOF Then
         vOverAmount = vOverAmount + RsCredit!Credit
         vSQL = "Insert into Ageing (PaymentVoucherNo, PaymentDate,  OrganizationID, PaymentType, LastPayment, AccountNo, OverAmount) values (" & RsCredit!VoucherNo & ",'" & RsCredit!VoucherDate & "'," & IIf(IsNull(RsCredit!OrganizationID), "Null", RsCredit!OrganizationID) & ",'" & RsCredit!vouchertype & "'," & RsCredit!Credit & ",'" & vAccountNo & "'," & vOverAmount & ")"
         CN.Execute (vSQL)
      End If
   Wend
   ''''''''''''''''''''''''''''''''''
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub SaleAgeingInsert()
   If Not RsCredit.EOF Then
   
      vSQL = "Insert into Ageing ( " & vbCrLf & _
      "VoucherNo, VoucherDate, VoucherType, PaymentVoucherNo, PaymentDate,  OrganizationID, " & vbCrLf & _
      "PaymentType, AccountNo, Debit, RecoveryAmount, LastPayment, OverAmount, Paid " & vbCrLf & _
      ") " & vbCrLf & _
      "Values( " & vbCrLf & _
      RsDebit!VoucherNo & ",'" & RsDebit!VoucherDate & "','" & RsDebit!vouchertype & "'," & RsCredit!VoucherNo & ",'" & RsCredit!VoucherDate & "'," & IIf(IsNull(RsDebit!OrganizationID), "Null", RsDebit!OrganizationID) & vbCrLf & _
      ",'" & RsCredit!vouchertype & "','" & RsDebit!AccountNo & "'," & RsDebit!Debit & "," & vOverAmount + RsDebit!Debit & "," & RsCredit!Credit & "," & vOverAmount & ",1 " & vbCrLf & _
      ") "
      
   Else
   
      vOverAmount = vOverAmount - RsDebit!Debit
      vSQL = "Insert into Ageing ( " & vbCrLf & _
      "VoucherNo, VoucherDate, VoucherType, PaymentVoucherNo, PaymentDate,  OrganizationID, " & vbCrLf & _
      "PaymentType, AccountNo, Debit, RecoveryAmount, LastPayment, OverAmount, Paid " & vbCrLf & _
      ") " & vbCrLf & _
      "Values( " & vbCrLf & _
      RsDebit!VoucherNo & ",'" & RsDebit!VoucherDate & "','" & RsDebit!vouchertype & "',Null, Null," & IIf(IsNull(RsDebit!OrganizationID), "Null", RsDebit!OrganizationID) & vbCrLf & _
      ",Null,'" & RsDebit!AccountNo & "'," & RsDebit!Debit & "," & vOverAmount + RsDebit!Debit & "," & 0 & "," & vOverAmount & ",0 " & vbCrLf & _
      ") "
      
   End If
   CN.Execute (vSQL)
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PurchaseAgeing()
   On Error GoTo ErrorHandler
   
   
   '''' Get Credit
   If RsCredit.State = adStateOpen Then RsCredit.Close
   vSQL = "Select * from #accountsledger where Credit <> 0 And AccountNo = '" & vAccountNo & "' Order by VoucherDate Asc"
   RsCredit.Open vSQL, CN, adOpenStatic, adLockBatchOptimistic
   ''''''''''''''''''''''''''''''''''
   
   '''' Get Debit
   If RsDebit.State = adStateOpen Then RsDebit.Close
   vSQL = "Select * from #accountsledger where Debit <> 0 And AccountNo = '" & vAccountNo & "' Order by VoucherDate Asc"
   RsDebit.Open vSQL, CN, adOpenStatic, adLockBatchOptimistic
   If Not RsDebit.EOF Then vOverAmount = RsDebit!Debit Else vOverAmount = 0
   
'''' Checking Debit / Credit / OverAmount
   While Not RsCredit.EOF
      Do While Not RsDebit.EOF
         If vOverAmount > RsCredit!Credit Then
            vOverAmount = vOverAmount - RsCredit!Credit
            PurchaseAgeingInsert
            Exit Do
         Else
            RsDebit.MoveNext
            If Not RsDebit.EOF Then vOverAmount = vOverAmount + RsDebit!Debit
         End If
      Loop
      If RsDebit.EOF Then PurchaseAgeingInsert
      RsCredit.MoveNext
   Wend
   ''''''''''''''''''''''''''''''''''
   
   '''' Add Remainig credit Amount in OverAmount
   While Not RsDebit.EOF
      RsDebit.MoveNext
      If Not RsDebit.EOF Then
         vOverAmount = vOverAmount + RsDebit!Debit
         vSQL = "Insert into Ageing (PaymentVoucherNo, PaymentDate,  OrganizationID, PaymentType, LastPayment, AccountNo, OverAmount) values (" & RsDebit!VoucherNo & ",'" & RsDebit!VoucherDate & "'," & IIf(IsNull(RsDebit!OrganizationID), "Null", RsDebit!OrganizationID) & ",'" & RsDebit!vouchertype & "'," & RsDebit!Debit & ",'" & vAccountNo & "'," & vOverAmount & ")"
         CN.Execute (vSQL)
      End If
   Wend
   ''''''''''''''''''''''''''''''''''
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub PurchaseAgeingInsert()
   If Not RsDebit.EOF Then
   
      vSQL = "Insert into Ageing ( " & vbCrLf & _
      "VoucherNo, VoucherDate, VoucherType, PaymentVoucherNo, PaymentDate,  OrganizationID, " & vbCrLf & _
      "PaymentType, AccountNo, Purchase, RecoveryAmount, LastPayment, OverAmount, Paid " & vbCrLf & _
      ") " & vbCrLf & _
      "Values( " & vbCrLf & _
      RsCredit!VoucherNo & ",'" & RsCredit!VoucherDate & "','" & RsCredit!vouchertype & "'," & RsCredit!VoucherNo & ",'" & RsCredit!VoucherDate & "'," & IIf(IsNull(RsCredit!OrganizationID), "Null", RsCredit!OrganizationID) & vbCrLf & _
      ",'" & RsCredit!vouchertype & "','" & RsCredit!AccountNo & "'," & RsCredit!Credit & "," & vOverAmount + RsCredit!Credit & "," & RsCredit!Credit & "," & vOverAmount & ",1 " & vbCrLf & _
      ") "
      
   Else
   
      vOverAmount = vOverAmount - RsCredit!Credit
      vSQL = "Insert into Ageing ( " & vbCrLf & _
      "VoucherNo, VoucherDate, VoucherType, PaymentVoucherNo, PaymentDate,  OrganizationID, " & vbCrLf & _
      "PaymentType, AccountNo, Purchase, RecoveryAmount, LastPayment, OverAmount, Paid " & vbCrLf & _
      ") " & vbCrLf & _
      "Values( " & vbCrLf & _
      RsCredit!VoucherNo & ",'" & RsCredit!VoucherDate & "','" & RsCredit!vouchertype & "',Null, Null," & IIf(IsNull(RsCredit!OrganizationID), "Null", RsCredit!OrganizationID) & vbCrLf & _
      ",Null,'" & RsCredit!AccountNo & "'," & RsCredit!Credit & "," & vOverAmount + RsCredit!Credit & "," & 0 & "," & vOverAmount & ", " & IIf(RsCredit!Credit = (vOverAmount + RsCredit!Credit), 1, 0) & vbCrLf & _
      ") "
      
   End If
   CN.Execute (vSQL)
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
    '-- when Company ID is written then it will check and all its related value will be write its appropriate places
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchZone.Show vbModal, Me
        If SchZone.ParaOutZoneID = "" Then FunSelectZone = False: Exit Function
        TxtZoneID.Text = SchZone.ParaOutZoneID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Zones where ZoneID=" & Val(TxtZoneID.Text)
    With CN.Execute(vStrSQL)
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
      TxtEmpID.SetFocus
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
    '-- when Company ID is written then it will check and all its related value will be write its appropriate places
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSector.Show vbModal, Me
        If SchSector.ParaOutSectorID = "" Then FunSelectSector = False: Exit Function
        TxtSectorID.Text = SchSector.ParaOutSectorID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Sectors where SectorID=" & Val(TxtSectorID.Text)
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
          TxtSectorName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function


