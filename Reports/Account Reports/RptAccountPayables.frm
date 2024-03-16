VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptAccountPayables 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptAccountPayables.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptVendorsLastPayment 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Only Vendors With Last Payment"
      Height          =   255
      Left            =   9585
      TabIndex        =   57
      Top             =   5790
      Width           =   2625
   End
   Begin VB.CheckBox ChkOpening 
      BackColor       =   &H00FF8080&
      Caption         =   "Include Opening"
      Height          =   255
      Left            =   4950
      TabIndex        =   54
      Top             =   6975
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.CheckBox ChkAllEmp 
      BackColor       =   &H00FFC0C0&
      Caption         =   "All Employees"
      Height          =   255
      Left            =   10080
      TabIndex        =   47
      Top             =   2925
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.OptionButton OptEmployees 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Only Employees"
      Height          =   255
      Left            =   4758
      TabIndex        =   46
      Top             =   5790
      Width           =   1545
   End
   Begin VB.OptionButton OptCustomers 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Only Customers"
      Height          =   255
      Left            =   6366
      TabIndex        =   45
      Top             =   5790
      Width           =   1545
   End
   Begin VB.OptionButton OptVendors 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Only Vendors"
      Height          =   255
      Left            =   7974
      TabIndex        =   44
      Top             =   5790
      Width           =   1545
   End
   Begin VB.OptionButton OptParties 
      BackColor       =   &H00FFFFFF&
      Caption         =   "All Parties"
      Height          =   255
      Left            =   3150
      TabIndex        =   43
      Top             =   5790
      Value           =   -1  'True
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6420
      TabIndex        =   32
      Top             =   7425
      Width           =   1965
      Begin VB.OptionButton RdoSummary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary"
         Height          =   255
         Left            =   855
         TabIndex        =   8
         Top             =   45
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton RdoDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detail"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   45
         Width           =   840
      End
   End
   Begin VB.CheckBox ChkZone 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Zone"
      Height          =   255
      Left            =   10095
      TabIndex        =   31
      Top             =   3615
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.CheckBox ChkSector 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Sector"
      Height          =   255
      Left            =   10095
      TabIndex        =   30
      Top             =   4410
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.TextBox TxtSectorName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6368
      TabIndex        =   22
      Top             =   4395
      Width           =   3585
   End
   Begin VB.TextBox TxtSectorID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4988
      TabIndex        =   2
      Top             =   4395
      Width           =   1020
   End
   Begin VB.TextBox TxtZoneID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4988
      TabIndex        =   1
      Top             =   3600
      Width           =   1020
   End
   Begin VB.TextBox TxtZoneName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6368
      TabIndex        =   21
      Top             =   3600
      Width           =   3585
   End
   Begin VB.CheckBox ChkExclude 
      BackColor       =   &H00B98A03&
      Caption         =   "Exclude Accounts Having Zero Balance."
      Height          =   255
      Left            =   7125
      TabIndex        =   12
      Top             =   6975
      Value           =   1  'Checked
      Width           =   3285
   End
   Begin JeweledBut.JeweledButton CmdPreview 
      Height          =   420
      Left            =   5655
      TabIndex        =   9
      Top             =   9045
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
      MICON           =   "RptAccountPayables.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   6990
      TabIndex        =   10
      Top             =   9030
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
      MICON           =   "RptAccountPayables.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdClose 
      Height          =   420
      Left            =   8340
      TabIndex        =   11
      Top             =   9030
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
      MICON           =   "RptAccountPayables.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6023
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
      MICON           =   "RptAccountPayables.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   5003
      TabIndex        =   0
      Top             =   2865
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
      Left            =   6383
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
   Begin SITextBox.Txt TxtFrom 
      Height          =   315
      Left            =   6105
      TabIndex        =   5
      Top             =   8520
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   5925
      TabIndex        =   3
      Top             =   6465
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
      Left            =   7650
      TabIndex        =   4
      Top             =   6465
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
      Left            =   6023
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3600
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
      MICON           =   "RptAccountPayables.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSector 
      Height          =   330
      Left            =   6023
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4395
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
      MICON           =   "RptAccountPayables.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtTo 
      Height          =   315
      Left            =   7605
      TabIndex        =   6
      Top             =   8520
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
   Begin JeweledBut.JeweledButton BtnCustomerType 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6008
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   2085
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
      MICON           =   "RptAccountPayables.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerTypeID 
      Height          =   315
      Left            =   4988
      TabIndex        =   34
      Top             =   2085
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
      IntegralPoint   =   10
   End
   Begin SITextBox.Txt TxtCustomerType 
      Height          =   315
      Left            =   6368
      TabIndex        =   35
      Tag             =   "NC"
      Top             =   2085
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
   Begin JeweledBut.JeweledButton BtnVender 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6038
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5205
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
      MICON           =   "RptAccountPayables.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtVenderName 
      Height          =   315
      Left            =   6398
      TabIndex        =   39
      Tag             =   "nc"
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
   Begin SITextBox.Txt TxtVenderID 
      Height          =   315
      Left            =   5018
      TabIndex        =   40
      Top             =   5205
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
      IntegralPoint   =   10
   End
   Begin JeweledBut.JeweledButton BtnEmployee 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6038
      TabIndex        =   48
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1395
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
      MICON           =   "RptAccountPayables.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmpName 
      Height          =   315
      Left            =   6398
      TabIndex        =   49
      Tag             =   "nc"
      Top             =   1395
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
   Begin SITextBox.Txt TxtEmpID 
      Height          =   315
      Left            =   5018
      TabIndex        =   50
      Top             =   1395
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtRemarks 
      Height          =   315
      Left            =   6150
      TabIndex        =   55
      Top             =   9840
      Visible         =   0   'False
      Width           =   3285
      _ExtentX        =   5794
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
   Begin VB.Label LblTakeTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "It will take Some time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   10035
      TabIndex        =   58
      Top             =   5535
      Visible         =   0   'False
      Width           =   2190
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
      Left            =   6120
      TabIndex        =   56
      Top             =   8295
      Width           =   420
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   6150
      TabIndex        =   53
      Top             =   9615
      Visible         =   0   'False
      Width           =   750
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
      Left            =   5018
      TabIndex        =   52
      Top             =   1185
      Visible         =   0   'False
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
      Left            =   6413
      TabIndex        =   51
      Top             =   1185
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vender ID"
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
      Left            =   5018
      TabIndex        =   42
      Top             =   5010
      Width           =   870
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Name"
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
      Left            =   6413
      TabIndex        =   41
      Top             =   5010
      Width           =   1155
   End
   Begin VB.Label LblCustomerType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Type"
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
      Left            =   6368
      TabIndex        =   37
      Top             =   1875
      Width           =   1275
   End
   Begin VB.Label LblCustomerTypeID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Type ID"
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
      Left            =   4988
      TabIndex        =   36
      Top             =   1875
      Width           =   1530
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
      Left            =   7605
      TabIndex        =   29
      Top             =   8295
      Width           =   240
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
      Left            =   6383
      TabIndex        =   28
      Top             =   4185
      Width           =   1110
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
      Left            =   5003
      TabIndex        =   27
      Top             =   4185
      Width           =   825
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
      Left            =   5003
      TabIndex        =   26
      Top             =   3390
      Width           =   705
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
      Left            =   6383
      TabIndex        =   25
      Top             =   3390
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
      Index           =   0
      Left            =   5925
      TabIndex        =   20
      Top             =   6240
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
      Left            =   7680
      TabIndex        =   19
      Top             =   6240
      Width           =   705
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
      Left            =   6120
      TabIndex        =   18
      Top             =   8070
      Width           =   2835
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
      Left            =   5003
      TabIndex        =   17
      Top             =   2640
      Width           =   1290
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
      Left            =   6383
      TabIndex        =   16
      Top             =   2640
      Width           =   1590
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Payable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   0
      Left            =   2700
      TabIndex        =   13
      Top             =   270
      Width           =   2244
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "RptAccountPayables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim Application1 As New CRAXDRT.Application
Dim vStrSQL As String

Private Sub BtnGroup_Click()
   If FunSelectOrganizaton(ssButton, False) = True Then
      TxtZoneID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub ChkSector_Click()
   TxtSectorID.Enabled = ChkSector.Value
   BtnSector.Enabled = ChkSector.Value
   vSector = Not BtnSector.Enabled
End Sub

Private Sub BtnZone_Click()
   If FunSelectZone(ssButton, False) = True Then
      TxtSectorID.SetFocus
   Else
      TxtZoneID.SetFocus
   End If
End Sub

Private Sub ChkZone_Click()
   TxtZoneID.Enabled = ChkZone.Value
   BtnZone.Enabled = ChkZone.Value
   vZone = Not BtnZone.Enabled
End Sub



Private Sub OptCustomers_Click()
   LblTakeTime.Visible = False
End Sub

Private Sub OptEmployees_Click()
   LblTakeTime.Visible = False
End Sub

Private Sub OptParties_Click()
      LblTakeTime.Visible = False
End Sub

Private Sub OptVendors_Click()
   LblTakeTime.Visible = False
End Sub

Private Sub OptVendorsLastPayment_Click()
      LblTakeTime.Visible = True
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
      TxtVenderID.SetFocus
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
         Case TxtCustomerTypeID.Name: If FunSelectCustomerType(ssFunctionKey, False) = True Then If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganizaton(ssFunctionKey, True) = True Then TxtZoneID.SetFocus
         Case TxtZoneID.Name: If FunSelectZone(ssFunctionKey, True) = True Then TxtSectorID.SetFocus
         Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, True) = True Then If TxtVenderID.Enabled Then TxtVenderID.SetFocus
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, True) = True Then DtpFrom.SetFocus
      End Select
  End If
End Sub

Private Function FunRefreshData() As Boolean
  On Error GoTo ErrorHandler
  Dim vSQL As String, vWhere  As String
  Me.MousePointer = vbHourglass
  If OptVendorsLastPayment.Value = True Then
      
      CN.Execute "Delete From AccountsLedger"
        
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
   
   vSQL = "Select c.* from ChartofAccounts c  " & vbCrLf & _
      " left outer JOIN Parties p ON  p.PartyID = c.AccountNo " & vbCrLf & _
        " left outer Join Sectors S on S.sectorId = p.sectorID" & vbCrLf & _
        " left outer join Zones Z on Z.ZoneID = S.ZoneID " & vbCrLf & _
        " Where c.AccountNo Like '61%'" & vbCrLf & _
        IIf(Trim(TxtZoneID.Text) = "", "", " And z.ZoneID in ( " & TxtZoneID.Text & ")") & vbCrLf & _
        IIf(Trim(TxtSectorID.Text) = "", "", " And p.SectorID in ( " & TxtSectorID.Text & ")")
      With CN.Execute(vSQL)
         While Not .EOF
            vSQL = "EXECUTE SPAccountsLedgerNew " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & ",'" & !AccountNo & "', '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & ChkOpening.Value
            CN.Execute vSQL
            .MoveNext
         Wend
      End With
      CN.Execute "Insert into Accountsledger Select * from #AccountsLedger"
      CN.Execute "Delete Accountsledger where debit = 0"
      vSQL = "DELETE t1 " & vbCrLf & _
            "FROM Accountsledger t1, Accountsledger t2 " & vbCrLf & _
            "Where t1.AccountNo = t2.AccountNo  " & vbCrLf & _
            "AND t1.Entrytime < t2.Entrytime "
      CN.Execute (vSQL)
      vSQL = " Drop TABLE [dbo].[#AccountsLedger] "
   CN.Execute vSQL
  End If
  
     
   CN.Execute "EXECUTE SPAccountsBalancesNew '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & ChkOpening.Value
   
'   vSQL = "SELECT AccountsBalances.OrganizationID, OrganizationName, ChartOfAccounts.AccountNo, ChartOfAccounts.AccountName+ ' ' + isnull(p.phone1,'') + ' ' + isnull(p.phone2,'') + ' ' + isnull(p.Mobile,'') as AccountName, AccountsBalances.OpeningDebit,AccountsBalances.OpeningCredit, " & vbCrLf & _
        " AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, AccountsBalances.Debit, AccountsBalances.Credit, AccountsBalances.Bal," & vbCrLf & _
        " AccountsBalances.BalType, p.city, Z.ZoneID, ZoneName, p.SectorID, SectorName  FROM AccountsBalances INNER JOIN ChartOfAccounts ON  AccountsBalances.AccountNo = ChartOfAccounts.AccountNo " & vbCrLf & _
        " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf & _
        " left outer Join Sectors S on S.sectorId = p.sectorID" & vbCrLf & _
        " left outer join Zones Z on Z.ZoneID = S.ZoneID" & vbCrLf & _
        " left Outer Join Organizations O On O.OrganizationID = AccountsBalances.OrganizationID " & vbCrLf & _
        " Where (Bal * case when baltype = 'Cr' then -1 else 1 end) " & IIf(Val(TxtTo.Text) = 0, " < 0 ", " between " & Val(TxtTo.Text) * -1 & " and " & IIf(Val(TxtFrom.Text) = 0, 1, Val(TxtFrom.Text)) * -1) & vbCrLf & _
        " and AccountsBalances.accountno <> '621' and AccountsBalances.accountno like '6" & IIf(ChkAllEmp.Value = 1, "3%'", "%'") & " and ChartOfAccounts.isdetailed =1 " & IIf(Trim(TxtOrganizationID.Text) = "", "", " And AccountsBalances.OrganizationID = " & TxtOrganizationID.Text) & vbCrLf _
        & IIf(Trim(TxtVenderID.Text) = "", "", " And p.PartyID ='" & TxtVenderID.Text & "'") & vbCrLf _
        & IIf(Trim(TxtZoneID.Text) = "", "", " And z.ZoneID in ( " & TxtZoneID.Text & ")") & vbCrLf _
        & IIf(Trim(TxtSectorID.Text) = "", "", " And p.SectorID in ( " & TxtSectorID.Text & ")") & vbCrLf & _
        " order by ChartOfAccounts.AccountNo"
    
    If OptVendorsLastPayment.Value = True Then
    
      vSQL = "SELECT AccountsBalances.OrganizationID, OrganizationName, ChartOfAccounts.AccountNo, ChartOfAccounts.AccountName+ ' ' + isnull(p.phone1,'') + ' ' + isnull(p.phone2,'') + ' ' + isnull(p.Mobile,'') + isnull(' '+p.Mobile2,'') as AccountName, AccountsBalances.OpeningDebit,AccountsBalances.OpeningCredit, " & vbCrLf & _
        " AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, AccountsBalances.Debit, AccountsBalances.Credit, AccountsBalances.Bal, Al.VoucherDate, Al.Debit as LastPayment," & vbCrLf & _
        " AccountsBalances.BalType, p.city, Z.ZoneID, ZoneName, p.SectorID, SectorName  FROM AccountsBalances INNER JOIN ChartOfAccounts ON  AccountsBalances.AccountNo = ChartOfAccounts.AccountNo " & vbCrLf & _
        " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf & _
        " left outer join AccountsLedger AL ON AL.AccountNo  =  ChartOfAccounts.AccountNo  " & vbCrLf & _
        " left outer Join Sectors S on S.sectorId = p.sectorID" & vbCrLf & _
        " left outer join Zones Z on Z.ZoneID = S.ZoneID" & vbCrLf & _
        " left Outer Join Organizations O On O.OrganizationID = AccountsBalances.OrganizationID " & vbCrLf & _
        " Where accountsbalances.accountno <> '621' and (Bal * case when baltype = 'Cr' then -1 else 1 end) " & IIf(Val(TxtTo.Text) = 0, " < 0 ", " between " & Val(TxtTo.Text) * -1 & " and " & IIf(Val(TxtFrom.Text) = 0, 1, Val(TxtFrom.Text)) * -1) & vbCrLf & _
        " and isdetailed=1 and accountsbalances.accountno like " & IIf(OptParties.Value = True, "'6%'", "") & IIf(OptEmployees.Value = True, "'63%'", "") & IIf(OptCustomers.Value = True, "'62%'", "") & IIf(OptVendorsLastPayment.Value = True, "'61%'", "") & IIf(Trim(TxtOrganizationID.Text) = "", "", " And AccountsBalances.OrganizationID = " & TxtOrganizationID.Text) & vbCrLf _
        + IIf(Trim(TxtZoneID.Text) = "", "", " And z.ZoneID in ( " & TxtZoneID.Text & ")") & vbCrLf _
        + IIf(Trim(TxtSectorID.Text) = "", "", " And p.SectorID in ( " & TxtSectorID.Text & ")") & vbCrLf _
        + IIf(Trim(TxtEmpID.Text) = "", "", " And p.EmpID = '" & TxtEmpID.Text & "'") & vbCrLf _
        & IIf(Trim(TxtVenderID.Text) = "", "", " And p.PartyID ='" & TxtVenderID.Text & "'") & vbCrLf _
        + IIf(Trim(TxtCustomerTypeID.Text) = "", "", " And p.CustomerTypeID = '" & TxtCustomerTypeID.Text & "'") & vbCrLf _
        + IIf(Trim(TxtRemarks.Text) = "", "", " And p.Remarks = '" & TxtRemarks.Text & "'") & vbCrLf _
        + " order by ChartOfAccounts.AccountName"
    Else
    
      vSQL = "SELECT AccountsBalances.OrganizationID, OrganizationName, ChartOfAccounts.AccountNo, ChartOfAccounts.AccountName+ ' ' + isnull(p.phone1,'') + ' ' + isnull(p.phone2,'') + ' ' + isnull(p.Mobile,'') + isnull(' '+p.Mobile2,'') as AccountName, AccountsBalances.OpeningDebit,AccountsBalances.OpeningCredit, " & vbCrLf & _
        " AccountsBalances.OpeningBal, AccountsBalances.OpeningBalType, AccountsBalances.Debit, AccountsBalances.Credit, AccountsBalances.Bal," & vbCrLf & _
        " AccountsBalances.BalType, p.city, Z.ZoneID, ZoneName, p.SectorID, SectorName  FROM AccountsBalances INNER JOIN ChartOfAccounts ON  AccountsBalances.AccountNo = ChartOfAccounts.AccountNo " & vbCrLf & _
        " left outer JOIN Parties p ON  p.PartyID = ChartOfAccounts.AccountNo " & vbCrLf & _
        " left outer Join Sectors S on S.sectorId = p.sectorID" & vbCrLf & _
        " left outer join Zones Z on Z.ZoneID = S.ZoneID" & vbCrLf & _
        " left Outer Join Organizations O On O.OrganizationID = AccountsBalances.OrganizationID " & vbCrLf & _
        " Where accountsbalances.accountno <> '621'and (Bal * case when baltype = 'Cr' then -1 else 1 end) " & IIf(Val(TxtTo.Text) = 0, IIf(Trim(TxtFrom.Text) = "", " < 0  ", " <= " & Val(TxtFrom.Text) * -1), " between " & Val(TxtTo.Text) * -1 & " and " & IIf(Val(TxtFrom.Text) = 0, 1, Val(TxtFrom.Text)) * -1) & vbCrLf & _
        " and isdetailed=1 and accountsbalances.accountno like " & IIf(OptParties.Value = True, "'6%'", "") & IIf(OptEmployees.Value = True, "'63%'", "") & IIf(OptCustomers.Value = True, "'62%'", "") & IIf(OptVendors.Value = True, "'61%'", "") & IIf(Trim(TxtOrganizationID.Text) = "", "", " And AccountsBalances.OrganizationID = " & TxtOrganizationID.Text) & vbCrLf _
        + IIf(Trim(TxtZoneID.Text) = "", "", " And z.ZoneID in ( " & TxtZoneID.Text & ")") & vbCrLf _
        + IIf(Trim(TxtSectorID.Text) = "", "", " And p.SectorID in ( " & TxtSectorID.Text & ")") & vbCrLf _
        + IIf(Trim(TxtEmpID.Text) = "", "", " And p.EmpID = '" & TxtEmpID.Text & "'") & vbCrLf _
        & IIf(Trim(TxtVenderID.Text) = "", "", " And p.PartyID ='" & TxtVenderID.Text & "'") & vbCrLf _
        + IIf(Trim(TxtCustomerTypeID.Text) = "", "", " And p.CustomerTypeID = '" & TxtCustomerTypeID.Text & "'") & vbCrLf _
        + IIf(Trim(TxtRemarks.Text) = "", "", " And p.Remarks = '" & TxtRemarks.Text & "'") & vbCrLf _
        + " order by ChartOfAccounts.AccountName"
   
   End If
  Set Rs = CN.Execute(vSQL)
  Me.MousePointer = vbDefault
  FunRefreshData = True
  Exit Function
ErrorHandler:
   Me.MousePointer = vbDefault
  Call ShowErrorMessage
  FunRefreshData = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
   If RdoDetail.Value = True Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\AccountReports\CrpAccountBalancesDetail.rpt")
      'Set RptReportViewer.Report = New CrpAccountBalancesDetail
   ElseIf RdoSummary.Value = True Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & IIf(OptVendorsLastPayment.Value = True, "\reports\AccountReports\CrpAccountBalancesSummaryWithLastPayment.rpt", "\reports\AccountReports\CrpAccountBalancesSummary.rpt"))
      'Set RptReportViewer.Report = New CrpAccountBalancesSummary
   End If
   RptReportViewer.Report.ReportTitle = "Account Payable"
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity) & IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue "Date From " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To " & Format(DtpTo.DateValue, "dd/MM/yyyy")
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
   SetWindowText Me.hWnd, "Account Payable"
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   
'   TxtCustomerTypeID.Text = ObjRegistry.CustomerTypeVisible
'   FunSelectCustomerType ssValidate, True
'   TxtCustomerTypeID.Visible = ObjRegistry.CustomerTypeVisible
'   BtnCustomerType.Visible = ObjRegistry.CustomerTypeVisible
'   TxtCustomerType.Visible = ObjRegistry.CustomerTypeVisible
'   LblCustomerTypeID.Visible = ObjRegistry.CustomerTypeVisible
'   LblCustomerType.Visible = ObjRegistry.CustomerTypeVisible
   
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
Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
      DtpFrom.SetFocus
   Else
      TxtVenderID.SetFocus
   End If
End Sub

Private Function FunSelectVender(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInDetail = ""
        SchAccounts.CmbFilter = "Vendors"
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectVender = False: Exit Function
        TxtVenderID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select c.* FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " where BarCode = '" & (TxtVenderID.Text) & "' or (c.AccountNo = '" & (TxtVenderID.Text) & "' and (c.AccountNo like '6%') and c.isDetailed = 1 and c.isLocked = 0)"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtVenderID.Text = !AccountNo
          TxtVenderName.Text = !AccountName
          FunSelectVender = True
          .Close
          Exit Function
      Else
          FunSelectVender = False
          .Close
          TxtVenderID.Text = ""
          TxtVenderName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtVenderID_Change()
   If TxtVenderID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   If TxtVenderName.Text <> "" Then TxtVenderName.Text = ""
End Sub

Private Sub TxtVenderID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtVenderID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVender(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVender(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub BtnCustomerType_Click()
   If FunSelectCustomerType(ssButton, False) = True Then
      If TxtSectorID.Enabled Then TxtSectorID.SetFocus
   Else
      If TxtCustomerTypeID.Enabled Then TxtCustomerTypeID.SetFocus
   End If
End Sub
Private Sub TxtCustomerTypeID_Change()
   If TxtCustomerTypeID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomerTypeID.Name Then Exit Sub
   If TxtCustomerTypeID.Text = "" Then TxtCustomerType.Text = ""
End Sub

Private Sub TxtCustomerTypeID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtCustomerTypeID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerTypeID.Text = "" Then Exit Sub
   If TxtCustomerType.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCustomerType(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCustomerType(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Function FunSelectCustomerType(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCustomerType.Show vbModal, Me
        If SchCustomerType.ParaOutID = "" Then FunSelectCustomerType = False: Exit Function
        TxtCustomerTypeID.Text = SchCustomerType.ParaOutID
    End If
    '---------------------------
    vStrSQL = " Select * FROM CustomerTypes where CustomerTypeID = '" & TxtCustomerTypeID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerType.Text = !CustomerType
          FunSelectCustomerType = True
          .Close
          Exit Function
      Else
          FunSelectCustomerType = False
          .Close
          TxtCustomerTypeID.Text = ""
          TxtCustomerType.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

