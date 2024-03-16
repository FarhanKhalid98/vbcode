VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptRoceverySheet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptRoceverySheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   4410
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   4995
      Width           =   2355
   End
   Begin VB.CheckBox ChkIncludeZeroBalance 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Include Zero Balance"
      Height          =   255
      Left            =   6255
      TabIndex        =   47
      Top             =   6255
      Width           =   2430
   End
   Begin VB.CheckBox ChkEmployee 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Employee"
      Height          =   255
      Left            =   8595
      TabIndex        =   46
      Top             =   4185
      Value           =   1  'Checked
      Width           =   1710
   End
   Begin VB.ComboBox CmbFilter 
      Height          =   315
      ItemData        =   "RptRoceverySheet.frx":0ECA
      Left            =   6120
      List            =   "RptRoceverySheet.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   6705
      Width           =   2715
   End
   Begin VB.CheckBox ChkOrganization 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Organization"
      Height          =   255
      Left            =   8550
      TabIndex        =   35
      Top             =   1695
      Value           =   1  'Checked
      Width           =   1710
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2415
      TabIndex        =   34
      Top             =   7290
      Width           =   6690
      Begin VB.OptionButton RdoNoRecovery 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Recovery  From Date To Date"
         Height          =   255
         Left            =   45
         TabIndex        =   7
         Top             =   45
         Width           =   2730
      End
      Begin VB.OptionButton RdoRecovery 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Recovery  From Date To Date"
         Height          =   255
         Left            =   2790
         TabIndex        =   8
         Top             =   45
         Width           =   2445
      End
      Begin VB.OptionButton RdoBoth 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All Accounts"
         Height          =   255
         Left            =   5355
         TabIndex        =   9
         Top             =   45
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CheckBox ChkSector 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Sector"
      Height          =   255
      Left            =   8580
      TabIndex        =   33
      Top             =   3225
      Value           =   1  'Checked
      Width           =   1710
   End
   Begin VB.CheckBox ChkZone 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Zone"
      Height          =   255
      Left            =   8580
      TabIndex        =   32
      Top             =   2385
      Value           =   1  'Checked
      Width           =   1710
   End
   Begin VB.TextBox TxtZoneName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4905
      TabIndex        =   22
      Top             =   2385
      Width           =   3585
   End
   Begin VB.TextBox TxtZoneID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3525
      MaxLength       =   2
      TabIndex        =   1
      Top             =   2385
      Width           =   1020
   End
   Begin VB.TextBox TxtSectorID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3525
      MaxLength       =   3
      TabIndex        =   2
      Top             =   3225
      Width           =   1020
   End
   Begin VB.TextBox TxtSectorName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4905
      TabIndex        =   21
      Top             =   3225
      Width           =   3585
   End
   Begin VB.CheckBox ChkExclude 
      BackColor       =   &H00B98A03&
      Caption         =   "Exclude Accounts Having Zero Balance."
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   8700
      Visible         =   0   'False
      Width           =   3285
   End
   Begin JeweledBut.JeweledButton CmdPreview 
      Height          =   420
      Left            =   3780
      TabIndex        =   10
      Top             =   7905
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
      MICON           =   "RptRoceverySheet.frx":0ECE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   5145
      TabIndex        =   11
      Top             =   7905
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
      MICON           =   "RptRoceverySheet.frx":0EEA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdClose 
      Height          =   420
      Left            =   6480
      TabIndex        =   12
      Top             =   7905
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
      MICON           =   "RptRoceverySheet.frx":0F06
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4545
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1650
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
      MICON           =   "RptRoceverySheet.frx":0F22
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   3525
      TabIndex        =   0
      Top             =   1650
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
   Begin SITextBox.Txt TxtOrganizatonName 
      Height          =   315
      Left            =   4905
      TabIndex        =   16
      Tag             =   "nc"
      Top             =   1650
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
      Left            =   4260
      TabIndex        =   3
      Top             =   5640
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
      Left            =   5985
      TabIndex        =   4
      Top             =   5640
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
      Left            =   4545
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2385
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
      MICON           =   "RptRoceverySheet.frx":0F3E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSector 
      Height          =   330
      Left            =   4545
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3225
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
      MICON           =   "RptRoceverySheet.frx":0F5A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtFrom 
      Height          =   315
      Left            =   3105
      TabIndex        =   5
      Top             =   6705
      Width           =   1410
      _ExtentX        =   2487
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
   Begin SITextBox.Txt TxtTo 
      Height          =   315
      Left            =   4605
      TabIndex        =   6
      Top             =   6705
      Width           =   1410
      _ExtentX        =   2487
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
   Begin JeweledBut.JeweledButton BtnWeeklyRecovery 
      Height          =   540
      Left            =   10860
      TabIndex        =   38
      Top             =   2228
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      TX              =   "Weekly Recovery"
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
      MICON           =   "RptRoceverySheet.frx":0F76
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDiscount 
      Height          =   540
      Left            =   10860
      TabIndex        =   39
      Top             =   2783
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      TX              =   "Discount Report"
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
      MICON           =   "RptRoceverySheet.frx":0F92
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnAgeingReport 
      Height          =   540
      Left            =   10860
      TabIndex        =   40
      Top             =   3368
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      TX              =   "Ageing Report"
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
      MICON           =   "RptRoceverySheet.frx":0FAE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnEmployee 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4530
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4170
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
      MICON           =   "RptRoceverySheet.frx":0FCA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmpName 
      Height          =   315
      Left            =   4890
      TabIndex        =   42
      Tag             =   "nc"
      Top             =   4170
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
   Begin SITextBox.Txt TxtEmpID 
      Height          =   315
      Left            =   3510
      TabIndex        =   43
      Top             =   4170
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Type"
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
      Left            =   4455
      TabIndex        =   49
      Top             =   4725
      Width           =   1065
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
      Left            =   4905
      TabIndex        =   45
      Top             =   3960
      Width           =   1365
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
      Left            =   3510
      TabIndex        =   44
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parent Accounts"
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
      Left            =   6105
      TabIndex        =   37
      Top             =   6480
      Width           =   1425
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   4605
      TabIndex        =   31
      Top             =   6480
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   3120
      TabIndex        =   30
      Top             =   6480
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-------------- Amount Limit -------------"
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
      Left            =   3120
      TabIndex        =   29
      Top             =   6255
      Width           =   2835
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
      Left            =   4905
      TabIndex        =   28
      Top             =   2175
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
      Left            =   3525
      TabIndex        =   27
      Top             =   2175
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
      Left            =   3525
      TabIndex        =   26
      Top             =   3015
      Width           =   825
   End
   Begin VB.Label Label5 
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
      Left            =   4905
      TabIndex        =   25
      Top             =   3015
      Width           =   1110
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
      Left            =   4905
      TabIndex        =   20
      Top             =   1425
      Width           =   1590
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
      Left            =   3525
      TabIndex        =   19
      Top             =   1425
      Width           =   1290
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
      Left            =   6015
      TabIndex        =   18
      Top             =   5415
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
      Left            =   4260
      TabIndex        =   17
      Top             =   5415
      Width           =   885
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recovery Sheet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   14
      Top             =   270
      Width           =   2115
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "RptRoceverySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSQL As String
Dim Application1 As New CRAXDRT.Application

Private Sub BtnAgeingReport_Click()
   RptAgeingReport.Show
End Sub

Private Sub BtnDiscount_Click()
   RptDisc.Show
End Sub

Private Sub BtnEmployee_Click()
   If FunSelectEmployee(ssButton, False) = True Then
      DtpFrom.SetFocus
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

Private Sub ChkEmployee_Click()
   TxtEmpID.Enabled = ChkEmployee.Value
   BtnEmployee.Enabled = ChkEmployee.Value
   vEmployee = Not BtnEmployee.Enabled
End Sub

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


Private Sub BtnOrganization_Click()
   If FunSelectOrganizaton(ssButton, False) = True Then
      DtpFrom.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnWeeklyRecovery_Click()
   RptWeeklyRecovery.Show
End Sub

Private Sub BtnZone_Click()
   If FunSelectZone(ssButton, False) = True Then
      TxtSectorID.SetFocus
   Else
      TxtZoneID.SetFocus
   End If
End Sub

Private Sub ChkOrganization_Click()
   TxtOrganizationID.Enabled = ChkOrganization.Value
   BtnOrganization.Enabled = ChkOrganization.Value
   vOrganization = Not BtnOrganization.Enabled
End Sub

Private Sub ChkSector_Click()
   TxtSectorID.Enabled = ChkSector.Value
   BtnSector.Enabled = ChkSector.Value
   vSector = Not BtnSector.Enabled
End Sub

Private Sub ChkZone_Click()
   TxtZoneID.Enabled = ChkZone.Value
   BtnZone.Enabled = ChkZone.Value
   vZone = Not BtnZone.Enabled
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

Private Sub CmdClose_Click()
   Unload Me
End Sub

Private Sub CmdPreview_Click()
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

Private Sub CmdPrint_Click()
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
         Case TxtOrganizationID.Name: If FunSelectOrganizaton(ssFunctionKey, True) = True Then DtpFrom.SetFocus
         Case TxtZoneID.Name: If FunSelectZone(ssFunctionKey, True) = True Then TxtSectorID.SetFocus
         Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, True) = True Then DtpFrom.SetFocus
      End Select
  End If
End Sub

Private Function FunRefreshData() As Boolean
   On Error GoTo ErrorHandler
   Dim vSQL As String, vWhere  As String
  
'    vWhere = " And c.AccountNo like '" & CmbFilter.ItemData(CmbFilter.ListIndex) & "%'"

   If ChkOrganization.Value = 1 Then
      If RdoBoth.Value = True Then
         vWhere = ""
      ElseIf RdoNoRecovery.Value = True Then
         vWhere = " and not (dbo.FunLastRecoveryDateOrganization('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo, AccountsBalances.OrganizationID) between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "')"
      ElseIf RdoRecovery.Value = True Then
         vWhere = " and dbo.FunLastRecoveryDateOrganization('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo, AccountsBalances.OrganizationID) between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "'"
      End If
      
      CN.Execute "EXECUTE SPAccountsBalancesNew '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
         If ChkEmployee.Value = 1 Then
            vSQL = "SELECT AccountsBalances.OrganizationID, OrganizationName, ChartOfAccounts.AccountNo, " & vbCrLf _
            + " ChartOfAccounts.AccountName + isnull(' '+p.Address,'') + isnull(' '+p.phone1 ,'') + isnull(' '+p.phone2,'') + isnull(' '+p.Mobile,'') + isnull(' '+p.Mobile2,'') + isnull(' '+p.ContactPerson,'') as AccountName, " & vbCrLf _
            + " AccountsBalances.OpeningDebit,AccountsBalances.OpeningCredit,  AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, " & vbCrLf _
            + "  dbo.FunLastRecoveryAmountOrganization('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo, AccountsBalances.OrganizationID) as LastReceived," & vbCrLf _
            + "  dbo.FunLastRecoveryDateOrganization('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo, AccountsBalances.OrganizationID) as LastRecoveryDate," & vbCrLf _
            + " AccountsBalances.Debit , AccountsBalances.Credit, Bal, AccountsBalances.BalType, p.city, sec.sectorid, SectorName, emp.empid, empname, Z.ZoneID, ZoneName, p.Description as Remarks" & vbCrLf _
            + " From AccountsBalances " & vbCrLf _
            + " INNER JOIN ChartOfAccounts ON  AccountsBalances.AccountNo = ChartOfAccounts.AccountNo " & vbCrLf _
            + " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf _
            + " left outer join  " & vbCrLf _
            + " (Select Customerid, empid from Recoveryheader h inner join RecoveryCustomer C on h.RecoveryID = c.RecoveryID Where recoveryDate between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "' group by customerid,empid ) R on AccountsBalances.AccountNo = R.Customerid  " & vbCrLf _
            + " left outer JOIN Employees Emp ON Emp.EmpID = R.Empid  " & vbCrLf _
            + " left outer join sectors sec on sec.sectorid = p.sectorid " & vbCrLf _
            + " left outer join zones z on z.zoneid = sec.zoneid " & vbCrLf _
            + " left Outer Join Organizations O On O.OrganizationID = AccountsBalances.OrganizationID " & vbCrLf _
            + IIf(ChkIncludeZeroBalance.Value = 0, " Where (Bal * case when baltype = 'Cr' then -1 else 1 end) " & IIf(Val(TxtTo.Text) = 0, " > " & Val(TxtFrom.Text), " between " & IIf(Val(TxtFrom.Text) = 0, 1, Val(TxtFrom.Text)) & " and " & Val(TxtTo.Text)), "where 1=1") & vbCrLf _
            + " And isdetailed=1 and accountsbalances.accountno like '" & IIf(CmbFilter.ListIndex > 0, CmbFilter.ItemData(CmbFilter.ListIndex), 6) & "%' " & IIf(Trim(TxtOrganizationID.Text) = "", "", " And AccountsBalances.OrganizationID = " & TxtOrganizationID.Text) & vbCrLf _
            & IIf(Trim(TxtZoneID.Text) = "", "", " And z.ZoneID = " & TxtZoneID.Text) & vbCrLf _
            & IIf(Trim(TxtSectorID.Text) = "", "", " And p.SectorID = " & TxtSectorID.Text) & vbCrLf _
            & IIf(Trim(TxtEmpID.Text) = "", "", " And Emp.EmpID = " & TxtEmpID.Text) & vbCrLf & _
            vWhere & vbCrLf & _
            " Order By ChartOfAccounts.AccountName"
          Else
            vSQL = "SELECT AccountsBalances.OrganizationID, OrganizationName, ChartOfAccounts.AccountNo, " & vbCrLf _
            + " ChartOfAccounts.AccountName + isnull(' '+p.Address,'') + isnull(' '+p.phone1 ,'') + isnull(' '+p.phone2,'') + isnull(' '+p.Mobile,'') + isnull(' '+p.Mobile2,'') + isnull(' '+p.ContactPerson,'') as AccountName, " & vbCrLf _
            + " AccountsBalances.OpeningDebit,AccountsBalances.OpeningCredit,  AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, " & vbCrLf _
            + "  dbo.FunLastRecoveryAmountOrganization('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo, AccountsBalances.OrganizationID) as LastReceived," & vbCrLf _
            + "  dbo.FunLastRecoveryDateOrganization('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo, AccountsBalances.OrganizationID) as LastRecoveryDate," & vbCrLf _
            + " AccountsBalances.Debit , AccountsBalances.Credit, Bal, AccountsBalances.BalType, p.city, sec.sectorid, SectorName,  Z.ZoneID, ZoneName, p.Description as Remarks" & vbCrLf _
            + " From AccountsBalances " & vbCrLf _
            + " INNER JOIN ChartOfAccounts ON  AccountsBalances.AccountNo = ChartOfAccounts.AccountNo " & vbCrLf _
            + " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf _
            + " left outer join sectors sec on sec.sectorid = p.sectorid " & vbCrLf _
            + " left outer join zones z on z.zoneid = sec.zoneid " & vbCrLf _
            + " left Outer Join Organizations O On O.OrganizationID = AccountsBalances.OrganizationID " & vbCrLf _
            + IIf(ChkIncludeZeroBalance.Value = 0, " Where (Bal * case when baltype = 'Cr' then -1 else 1 end) " & IIf(Val(TxtTo.Text) = 0, " > " & Val(TxtFrom.Text), " between " & IIf(Val(TxtFrom.Text) = 0, 1, Val(TxtFrom.Text)) & " and " & Val(TxtTo.Text)), "where 1=1") & vbCrLf _
            + " And isdetailed=1 and accountsbalances.accountno like '" & IIf(CmbFilter.ListIndex > 0, CmbFilter.ItemData(CmbFilter.ListIndex), 6) & "%' " & IIf(Trim(TxtOrganizationID.Text) = "", "", " And AccountsBalances.OrganizationID = " & TxtOrganizationID.Text) & vbCrLf _
            & IIf(Trim(TxtZoneID.Text) = "", "", " And z.ZoneID = " & TxtZoneID.Text) & vbCrLf _
            & IIf(Trim(TxtSectorID.Text) = "", "", " And p.SectorID = " & TxtSectorID.Text) & vbCrLf _
            & IIf(Trim(TxtEmpID.Text) = "", "", " And Emp.EmpID = " & TxtEmpID.Text) & vbCrLf & _
            vWhere & vbCrLf & _
            " Order By ChartOfAccounts.AccountName"
           End If
   Else
      If RdoBoth.Value = True Then
         vWhere = ""
      ElseIf RdoNoRecovery.Value = True Then
         vWhere = " and not (dbo.FunLastRecoveryDate('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo) between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "')"
      ElseIf RdoRecovery.Value = True Then
         vWhere = " and dbo.FunLastRecoveryDate('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo) between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "'"
      End If
      
      CN.Execute "EXECUTE SPAccountsBalances '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
'      vSQL = " SELECT null OrganizationID, null OrganizationName, ChartOfAccounts.AccountNo, " & vbCrLf _
'           + " ChartOfAccounts.AccountName + isnull(' '+p.Address,'') + isnull(' '+p.phone1 ,'') + isnull(' '+p.phone2,'') + isnull(' '+p.Mobile,'') + isnull(' '+p.Mobile2,'') + isnull(' '+p.ContactPerson,'') as AccountName, " & vbCrLf _
'           + " AccountsBalances.OpeningDebit,AccountsBalances.OpeningCredit,  AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, " & vbCrLf _
'           + " dbo.FunLastRecoveryAmount('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo) as LastReceived," & vbCrLf _
'           + " dbo.FunLastRecoveryDate('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo) as LastRecoveryDate," & vbCrLf _
'           + " AccountsBalances.Debit , AccountsBalances.Credit, Bal, AccountsBalances.BalType, p.city, sec.sectorid, SectorName, Z.ZoneID, emp.empid, empname, ZoneName, p.Description as Remarks" & vbCrLf _
'           + " From AccountsBalances " & vbCrLf _
'           + " INNER JOIN ChartOfAccounts ON  AccountsBalances.AccountNo = ChartOfAccounts.AccountNo " & vbCrLf _
'           + " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf _
'           + " left outer JOIN Employees Emp ON  Emp.EmpID = ChartOfAccounts.AccountNo " & vbCrLf _
'           + " left outer join sectors sec on sec.sectorid = p.sectorid " & vbCrLf _
'           + " left outer join zones z on z.zoneid = sec.zoneid " & vbCrLf _
'           + " Where (Bal * case when baltype = 'Cr' then -1 else 1 end) " & IIf(Val(TxtTo.Text) = 0, " > " & Val(TxtFrom.Text), " between " & IIf(Val(TxtFrom.Text) = 0, 1, Val(TxtFrom.Text)) & " and " & Val(TxtTo.Text)) & _
'           " And isdetailed=1 and accountsbalances.accountno like '" & IIf(CmbFilter.ListIndex > 0, CmbFilter.ItemData(CmbFilter.ListIndex), 6) & "%' " & vbCrLf _
'           & IIf(Trim(TxtZoneID.Text) = "", "", " And z.ZoneID = " & TxtZoneID.Text) & vbCrLf _
'           & IIf(Trim(TxtSectorID.Text) = "", "", " And p.SectorID = " & TxtSectorID.Text) & vbCrLf _
'           & IIf(Trim(TxtEmpID.Text) = "", "", " And Emp.EmpID = " & TxtEmpID.Text) & vbCrLf & _
'             vWhere & vbCrLf & _
'           " Order By ChartOfAccounts.AccountName".
      If ChkEmployee.Value = 1 Then
               vSQL = " SELECT null OrganizationID, null OrganizationName, ChartOfAccounts.AccountNo, " & vbCrLf _
                  + " ChartOfAccounts.AccountName + isnull(' '+p.Address,'') + isnull(' '+p.phone1 ,'') + isnull(' '+p.phone2,'') + isnull(' '+p.Mobile,'') + isnull(' '+p.Mobile2,'') + isnull(' '+p.ContactPerson,'') as AccountName, " & vbCrLf _
                  + " AccountsBalances.OpeningDebit,AccountsBalances.OpeningCredit,  AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, " & vbCrLf _
                  + " dbo.FunLastRecoveryAmount('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo) as LastReceived," & vbCrLf _
                  + " dbo.FunLastRecoveryDate('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo) as LastRecoveryDate," & vbCrLf _
                  + " AccountsBalances.Debit , AccountsBalances.Credit, Bal, AccountsBalances.BalType, p.city, sec.sectorid, SectorName, Z.ZoneID, emp.empid, empname, ZoneName, p.Description as Remarks" & vbCrLf _
                  + " From AccountsBalances " & vbCrLf _
                  + " INNER JOIN ChartOfAccounts ON  AccountsBalances.AccountNo = ChartOfAccounts.AccountNo " & vbCrLf _
                  + " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf _
                  + " left outer join  " & vbCrLf _
                  + " (Select Customerid, empid from Recoveryheader h inner join RecoveryCustomer C on h.RecoveryID = c.RecoveryID Where recoveryDate between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "' group by customerid,empid ) R on AccountsBalances.AccountNo = R.Customerid  " & vbCrLf _
                  + " left outer JOIN Employees Emp ON Emp.EmpID = R.Empid  " & vbCrLf _
                  + " left outer join sectors sec on sec.sectorid = p.sectorid " & vbCrLf _
                  + " left outer join zones z on z.zoneid = sec.zoneid " & vbCrLf _
                  + IIf(ChkIncludeZeroBalance.Value = 0, " Where (Bal * case when baltype = 'Cr' then -1 else 1 end) " & IIf(Val(TxtTo.Text) = 0, " > " & Val(TxtFrom.Text), " between " & IIf(Val(TxtFrom.Text) = 0, 1, Val(TxtFrom.Text)) & " and " & Val(TxtTo.Text)), "where 1=1") & vbCrLf _
                  + " And isdetailed=1 and accountsbalances.accountno like '" & IIf(CmbFilter.ListIndex > 0, CmbFilter.ItemData(CmbFilter.ListIndex), 6) & "%' " & IIf(Trim(TxtOrganizationID.Text) = "", "", " And AccountsBalances.OrganizationID = " & TxtOrganizationID.Text) & vbCrLf _
                  & IIf(Trim(TxtZoneID.Text) = "", "", " And z.ZoneID = " & TxtZoneID.Text) & vbCrLf _
                  & IIf(Trim(TxtSectorID.Text) = "", "", " And p.SectorID = " & TxtSectorID.Text) & vbCrLf _
                  & IIf(Trim(TxtEmpID.Text) = "", "", " And Emp.EmpID = " & TxtEmpID.Text) & vbCrLf & _
                  vWhere & vbCrLf & _
                  " Order By ChartOfAccounts.AccountName"
      Else
          vSQL = " SELECT null OrganizationID, null OrganizationName, ChartOfAccounts.AccountNo, " & vbCrLf _
                  + " ChartOfAccounts.AccountName + isnull(' '+p.Address,'') + isnull(' '+p.phone1 ,'') + isnull(' '+p.phone2,'') + isnull(' '+p.Mobile,'') + isnull(' '+p.Mobile2,'') + isnull(' '+p.ContactPerson,'') as AccountName, " & vbCrLf _
                  + " AccountsBalances.OpeningDebit,AccountsBalances.OpeningCredit,  AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, " & vbCrLf _
                  + " dbo.FunLastRecoveryAmount('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo) as LastReceived," & vbCrLf _
                  + " dbo.FunLastRecoveryDate('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo) as LastRecoveryDate," & vbCrLf _
                  + " AccountsBalances.Debit , AccountsBalances.Credit, Bal, AccountsBalances.BalType, p.city, sec.sectorid, SectorName, Z.ZoneID, ZoneName, p.Description as Remarks" & vbCrLf _
                  + " From AccountsBalances " & vbCrLf _
                  + " INNER JOIN ChartOfAccounts ON  AccountsBalances.AccountNo = ChartOfAccounts.AccountNo " & vbCrLf _
                  + " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf _
                  + " left outer join sectors sec on sec.sectorid = p.sectorid " & vbCrLf _
                  + " left outer join zones z on z.zoneid = sec.zoneid " & vbCrLf _
                  + IIf(ChkIncludeZeroBalance.Value = 0, " Where (Bal * case when baltype = 'Cr' then -1 else 1 end) " & IIf(Val(TxtTo.Text) = 0, " > " & Val(TxtFrom.Text), " between " & IIf(Val(TxtFrom.Text) = 0, 1, Val(TxtFrom.Text)) & " and " & Val(TxtTo.Text)), "where 1=1") & vbCrLf _
                  + " And isdetailed=1 and accountsbalances.accountno like '" & IIf(CmbFilter.ListIndex > 0, CmbFilter.ItemData(CmbFilter.ListIndex), 6) & "%' " & IIf(Trim(TxtOrganizationID.Text) = "", "", " And AccountsBalances.OrganizationID = " & TxtOrganizationID.Text) & vbCrLf _
                  & IIf(Trim(TxtZoneID.Text) = "", "", " And z.ZoneID = " & TxtZoneID.Text) & vbCrLf _
                  & IIf(Trim(TxtSectorID.Text) = "", "", " And p.SectorID = " & TxtSectorID.Text) & vbCrLf _
                  & IIf(Trim(TxtEmpID.Text) = "", "", " And Emp.EmpID = " & TxtEmpID.Text) & vbCrLf & _
                  vWhere & vbCrLf & _
                  " Order By ChartOfAccounts.AccountName"
   End If
   
   End If
 
 'dbo.FunLastRecoveryDate('" & DtpTo.DateValue & "', ChartOfAccounts.AccountNo, AccountsBalances.OrganizationID)
  
  Set Rs = CN.Execute(vSQL)
  FunRefreshData = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
  FunRefreshData = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
   Select Case CmbGroup.Text
         Case "Recovery Sheet"
'            Set RptReportViewer.Report = New CrpRecoverySheet
            Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\AccountReports\CrpRecoverySheet.rpt")
            RptReportViewer.Report.ReportTitle = "Recovery Sheet"
         Case "Recovery Sheet Employee Wise"
'            Set RptReportViewer.Report = New CrpRecoverySheetEmployeeWise
            Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\AccountReports\CrpRecoverySheetEmployeeWise.rpt")
            RptReportViewer.Report.ReportTitle = "Recovery Sheet Employee Wise"
         End Select
      
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue "From : " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & ",   To : " & Format(DtpTo.DateValue, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity) & IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
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
   SetWindowText Me.hWnd, "Recovery Sheet"
   
   CmbGroup.AddItem ("Recovery Sheet")
   CmbGroup.AddItem ("Recovery Sheet Employee Wise")
   CmbGroup.ListIndex = 0
   
   CmbFilter.AddItem "-- ALL PARENT ACCOUNTS --", 0
   With CN.Execute("Select AccountNo, AccountName from ChartofAccounts Where isDetailed = 0 and AccountDepth = 1 and AccountNo in ('61','62','63') order by AccountName")
      Do Until .EOF
         CmbFilter.AddItem !AccountName
         CmbFilter.ItemData(CmbFilter.NewIndex) = !AccountNo
         .MoveNext
      Loop
   End With
   CmbFilter.ListIndex = 0
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   vZone = Not BtnZone.Enabled
   vSector = Not BtnSector.Enabled
   vEmployee = Not BtnEmployee.Enabled
   vOrganization = Not BtnOrganization.Enabled
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
   vTemp = Not FunSelectOrganizaton(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganizaton(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectOrganizaton(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganizaton = False: Exit Function
       TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    If TxtOrganizationID.Text = "" Then FunSelectOrganizaton = False: Exit Function
    vStrSQL = " Select * FROM Organizations where OrganizationID='" & TxtOrganizationID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizatonName.Text = !OrganizationName
          FunSelectOrganizaton = True
          .Close
          Exit Function
      Else
          FunSelectOrganizaton = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizatonName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

