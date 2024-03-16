VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptSaleRegister 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptSaleRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8415
      TabIndex        =   152
      Top             =   9765
      Visible         =   0   'False
      Width           =   4050
      Begin VB.OptionButton OptMovingAvg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Moving Avg"
         Height          =   195
         Left            =   45
         TabIndex        =   155
         ToolTipText     =   "Simple Moving Average"
         Top             =   45
         Width           =   1350
      End
      Begin VB.OptionButton OptWeightedAvg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Weighted Avg"
         Height          =   195
         Left            =   1440
         TabIndex        =   154
         ToolTipText     =   "Weighted Mean"
         Top             =   45
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton OptLastPrice 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Last Price"
         Height          =   195
         Left            =   2970
         TabIndex        =   153
         Top             =   45
         Width           =   1035
      End
   End
   Begin MSComCtl2.DTPicker DTPTimeFrom 
      Height          =   345
      Left            =   8460
      TabIndex        =   149
      Top             =   8550
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "hh:mm:ss"
      Format          =   137035778
      UpDown          =   -1  'True
      CurrentDate     =   36494
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      Left            =   11220
      Style           =   2  'Dropdown List
      TabIndex        =   101
      Top             =   7635
      Width           =   1950
   End
   Begin VB.ComboBox CmbSortType 
      Height          =   315
      ItemData        =   "RptSaleRegister.frx":0ECA
      Left            =   10680
      List            =   "RptSaleRegister.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   96
      Top             =   9345
      Width           =   1275
   End
   Begin VB.ComboBox CmbSortName 
      Height          =   315
      ItemData        =   "RptSaleRegister.frx":0ECE
      Left            =   8655
      List            =   "RptSaleRegister.frx":0ED0
      Style           =   2  'Dropdown List
      TabIndex        =   93
      Top             =   9345
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7860
      TabIndex        =   88
      Top             =   6945
      Width           =   5265
      Begin VB.OptionButton RdoAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All"
         Height          =   255
         Left            =   4185
         TabIndex        =   28
         Top             =   10
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton RdoCredit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Credit"
         Height          =   255
         Left            =   1245
         TabIndex        =   26
         Top             =   10
         Width           =   825
      End
      Begin VB.OptionButton RdoCash 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cash"
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   10
         Width           =   1050
      End
      Begin VB.OptionButton RdoBankCard 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bank Card"
         Height          =   255
         Left            =   2670
         TabIndex        =   27
         Top             =   10
         Width           =   1050
      End
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   7830
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   7290
      Width           =   3390
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11265
      TabIndex        =   60
      Top             =   7290
      Width           =   1860
      Begin VB.OptionButton RdoSummary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary"
         Height          =   255
         Left            =   900
         TabIndex        =   31
         Top             =   10
         Width           =   960
      End
      Begin VB.OptionButton RdoDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detail"
         Height          =   255
         Left            =   75
         TabIndex        =   30
         Top             =   10
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7860
      TabIndex        =   59
      Top             =   6600
      Width           =   5265
      Begin VB.OptionButton RdoNet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Net Sale"
         Height          =   300
         Left            =   4125
         TabIndex        =   24
         Top             =   15
         Width           =   915
      End
      Begin VB.OptionButton RdoInv 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sale"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   15
         Width           =   630
      End
      Begin VB.OptionButton RdoReturn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sale Return"
         Height          =   255
         Left            =   1215
         TabIndex        =   22
         Top             =   15
         Width           =   1140
      End
      Begin VB.OptionButton RdoBoth 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Both"
         Height          =   255
         Left            =   2700
         TabIndex        =   23
         Top             =   15
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11400
      TabIndex        =   36
      Top             =   10245
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
      MICON           =   "RptSaleRegister.frx":0ED2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8640
      TabIndex        =   34
      Top             =   10215
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
      MICON           =   "RptSaleRegister.frx":0EEE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10005
      TabIndex        =   35
      Top             =   10245
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
      MICON           =   "RptSaleRegister.frx":0F0A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   5790
      TabIndex        =   61
      Top             =   420
      Visible         =   0   'False
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   16
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
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1605
      TabIndex        =   11
      Top             =   8325
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   15
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   2625
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   8325
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
      MICON           =   "RptSaleRegister.frx":0F26
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2985
      TabIndex        =   37
      Top             =   8325
      Width           =   3585
      _ExtentX        =   6324
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
   Begin JeweledBut.JeweledButton BtnOrganizaton 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8925
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3690
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
      MICON           =   "RptSaleRegister.frx":0F42
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   7905
      TabIndex        =   18
      Top             =   3690
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtOrganizatonName 
      Height          =   315
      Left            =   9285
      TabIndex        =   39
      Tag             =   "nc"
      Top             =   3690
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   40
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2025
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
      MICON           =   "RptSaleRegister.frx":0F5E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   2985
      TabIndex        =   41
      Tag             =   "nc"
      Top             =   2025
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
      Height          =   345
      Left            =   8460
      TabIndex        =   32
      Top             =   8145
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   609
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
      Left            =   10350
      TabIndex        =   33
      Top             =   8145
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
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   54
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   6435
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
      MICON           =   "RptSaleRegister.frx":0F7A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   1605
      TabIndex        =   8
      Top             =   6435
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   2985
      TabIndex        =   55
      Tag             =   "nc"
      Top             =   6435
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
   Begin JeweledBut.JeweledButton BtnSubGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   56
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7065
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
      MICON           =   "RptSaleRegister.frx":0F96
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   1605
      TabIndex        =   9
      Top             =   7065
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   2985
      TabIndex        =   57
      Tag             =   "nc"
      Top             =   7065
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
      Left            =   2625
      TabIndex        =   52
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5805
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
      MICON           =   "RptSaleRegister.frx":0FB2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   1605
      TabIndex        =   7
      Top             =   5805
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   2985
      TabIndex        =   53
      Tag             =   "nc"
      Top             =   5805
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
   Begin JeweledBut.JeweledButton BtnUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   50
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5175
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
      MICON           =   "RptSaleRegister.frx":0FCE
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtUserNo 
      Height          =   315
      Left            =   1605
      TabIndex        =   6
      Top             =   5175
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
   Begin SITextBox.Txt TxtUserName 
      Height          =   315
      Left            =   2985
      TabIndex        =   51
      Tag             =   "nc"
      Top             =   5175
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
      Left            =   1605
      TabIndex        =   1
      Top             =   2025
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin JeweledBut.JeweledButton BtnCustomer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   46
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3915
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
      MICON           =   "RptSaleRegister.frx":0FEA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPartyID 
      Height          =   315
      Left            =   1605
      TabIndex        =   4
      Top             =   3915
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
      IntegralPoint   =   10
   End
   Begin JeweledBut.JeweledButton BtnEmpName 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   48
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4545
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
      MICON           =   "RptSaleRegister.frx":1006
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmpID 
      Height          =   315
      Left            =   1605
      TabIndex        =   5
      Top             =   4545
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtPartyName 
      Height          =   315
      Left            =   2985
      TabIndex        =   47
      Tag             =   "nc"
      Top             =   3915
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
   Begin SITextBox.Txt TxtEmpName 
      Height          =   315
      Left            =   2985
      TabIndex        =   49
      Tag             =   "nc"
      Top             =   4545
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
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   44
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3285
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
      MICON           =   "RptSaleRegister.frx":1022
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSectorID 
      Height          =   315
      Left            =   1605
      TabIndex        =   3
      Top             =   3285
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtSectorName 
      Height          =   315
      Left            =   2985
      TabIndex        =   45
      Tag             =   "nc"
      Top             =   3285
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
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2655
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
      MICON           =   "RptSaleRegister.frx":103E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtZoneID 
      Height          =   315
      Left            =   1605
      TabIndex        =   2
      Top             =   2655
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtZoneName 
      Height          =   315
      Left            =   2985
      TabIndex        =   43
      Tag             =   "nc"
      Top             =   2655
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
   Begin JeweledBut.JeweledButton BtnMember 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8910
      TabIndex        =   89
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1575
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
      MICON           =   "RptSaleRegister.frx":105A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtMemberName 
      Height          =   315
      Left            =   9270
      TabIndex        =   90
      Tag             =   "nc"
      Top             =   1575
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
   Begin SITextBox.Txt TxtMemberID 
      Height          =   315
      Left            =   7890
      TabIndex        =   15
      Top             =   1575
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin JeweledBut.JeweledButton BtnBrand 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   97
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7695
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
      MICON           =   "RptSaleRegister.frx":1076
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBrandID 
      Height          =   315
      Left            =   1605
      TabIndex        =   10
      Top             =   7695
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   15
   End
   Begin SITextBox.Txt TxtBrandName 
      Height          =   315
      Left            =   2985
      TabIndex        =   98
      Tag             =   "nc"
      Top             =   7695
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
   Begin JeweledBut.JeweledButton BtnDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8910
      TabIndex        =   103
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2220
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
      MICON           =   "RptSaleRegister.frx":1092
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDepartmentID 
      Height          =   315
      Left            =   7890
      TabIndex        =   16
      Top             =   2220
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtDepartmentName 
      Height          =   315
      Left            =   9270
      TabIndex        =   104
      Tag             =   "nc"
      Top             =   2220
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
      Left            =   2610
      TabIndex        =   107
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1350
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
      MICON           =   "RptSaleRegister.frx":10AE
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtVenderName 
      Height          =   315
      Left            =   2970
      TabIndex        =   108
      Tag             =   "nc"
      Top             =   1350
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
      Left            =   1590
      TabIndex        =   0
      Top             =   1350
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
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
      IntegralPoint   =   6
   End
   Begin JeweledBut.JeweledButton BtnCustomerType 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8955
      TabIndex        =   111
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   6075
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
      MICON           =   "RptSaleRegister.frx":10CA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerTypeID 
      Height          =   315
      Left            =   7935
      TabIndex        =   21
      Top             =   6075
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
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
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtCustomerType 
      Height          =   315
      Left            =   9315
      TabIndex        =   112
      Tag             =   "NC"
      Top             =   6075
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
   Begin JeweledBut.JeweledButton BtnSubDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8895
      TabIndex        =   115
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2955
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
      MICON           =   "RptSaleRegister.frx":10E6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubDepartmentID 
      Height          =   315
      Left            =   7875
      TabIndex        =   17
      Top             =   2955
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtSubDepartmentName 
      Height          =   315
      Left            =   9255
      TabIndex        =   116
      Tag             =   "nc"
      Top             =   2955
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
   Begin JeweledBut.JeweledButton TxtSeason 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8940
      TabIndex        =   119
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4485
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
      MICON           =   "RptSaleRegister.frx":1102
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSeasonID 
      Height          =   315
      Left            =   7920
      TabIndex        =   19
      Top             =   4485
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtSeasonName 
      Height          =   315
      Left            =   9300
      TabIndex        =   120
      Tag             =   "nc"
      Top             =   4485
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
   Begin SITextBox.Txt TxtItemCode 
      Height          =   315
      Left            =   1620
      TabIndex        =   12
      Top             =   9030
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   15
   End
   Begin JeweledBut.JeweledButton BtnItemCode 
      Height          =   330
      Left            =   2640
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   9045
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
      MICON           =   "RptSaleRegister.frx":111E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtItemCodeName 
      Height          =   315
      Left            =   3000
      TabIndex        =   124
      Top             =   9030
      Width           =   3585
      _ExtentX        =   6324
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
   Begin JeweledBut.JeweledButton BtnItemDesc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2640
      TabIndex        =   127
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   10440
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
      MICON           =   "RptSaleRegister.frx":113A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtItemDescID 
      Height          =   315
      Left            =   1620
      TabIndex        =   14
      Top             =   10470
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtItemDescName 
      Height          =   315
      Left            =   3000
      TabIndex        =   128
      Tag             =   "nc"
      Top             =   10470
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
   Begin JeweledBut.JeweledButton BtnDescription 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2640
      TabIndex        =   131
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   9810
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
      MICON           =   "RptSaleRegister.frx":1156
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDescriptionID 
      Height          =   315
      Left            =   1620
      TabIndex        =   13
      Top             =   9795
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtDescriptionName 
      Height          =   315
      Left            =   3000
      TabIndex        =   132
      Tag             =   "nc"
      Top             =   9795
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
      Left            =   8970
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   5250
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
      MICON           =   "RptSaleRegister.frx":1172
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSessionID 
      Height          =   315
      Left            =   7965
      TabIndex        =   136
      Top             =   5250
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtSessionName 
      Height          =   315
      Left            =   9330
      TabIndex        =   137
      Top             =   5250
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
   Begin SITextBox.Txt TxtBillIDTo 
      Height          =   315
      Left            =   13680
      TabIndex        =   150
      Top             =   1920
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
   Begin SITextBox.Txt TxtManualBillIDTo 
      Height          =   315
      Left            =   13680
      TabIndex        =   142
      Top             =   3360
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
   Begin SITextBox.Txt TxtManualBillIDFrom 
      Height          =   315
      Left            =   13680
      TabIndex        =   145
      Top             =   3000
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
   Begin SITextBox.Txt TxtBillIDFrom 
      Height          =   315
      Left            =   13680
      TabIndex        =   148
      Top             =   1560
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
   Begin MSComCtl2.DTPicker DTPTimeTo 
      Height          =   315
      Left            =   10350
      TabIndex        =   151
      Top             =   8505
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "hh:mm:ss"
      Format          =   40239106
      UpDown          =   -1  'True
      CurrentDate     =   0.499988425925926
   End
   Begin JeweledBut.JeweledButton BtnUpdateStock 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   11745
      TabIndex        =   156
      Top             =   8145
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      TX              =   "Update Stock"
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
      MICON           =   "RptSaleRegister.frx":118E
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill ID"
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
      Left            =   13928
      TabIndex        =   147
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label53 
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
      Left            =   13080
      TabIndex        =   146
      Top             =   3060
      Width           =   420
   End
   Begin VB.Label Label52 
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
      Left            =   13200
      TabIndex        =   144
      Top             =   3420
      Width           =   240
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manual ID"
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
      Left            =   13755
      TabIndex        =   143
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label Label49 
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
      Left            =   13080
      TabIndex        =   141
      Top             =   1620
      Width           =   420
   End
   Begin VB.Label Label50 
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
      Left            =   13200
      TabIndex        =   140
      Top             =   1980
      Width           =   240
   End
   Begin VB.Label Label47 
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
      Left            =   7965
      TabIndex        =   139
      Top             =   5040
      Width           =   930
   End
   Begin VB.Label Label46 
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
      Left            =   9330
      TabIndex        =   138
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description ID"
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
      Left            =   1620
      TabIndex        =   134
      Top             =   9585
      Width           =   1230
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description Name"
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
      Left            =   3000
      TabIndex        =   133
      Top             =   9585
      Width           =   1515
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Desc."
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
      Left            =   1620
      TabIndex        =   130
      Top             =   10260
      Width           =   930
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Desc. Name"
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
      Left            =   3015
      TabIndex        =   129
      Top             =   10260
      Width           =   1470
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code Name"
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
      Left            =   3000
      TabIndex        =   126
      Top             =   8820
      Width           =   990
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      Left            =   1620
      TabIndex        =   125
      Top             =   8820
      Width           =   870
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Season ID"
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
      Left            =   7920
      TabIndex        =   122
      Top             =   4275
      Width           =   900
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Season Name"
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
      Left            =   9300
      TabIndex        =   121
      Top             =   4275
      Width           =   1185
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Dept. ID"
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
      Left            =   7875
      TabIndex        =   118
      Top             =   2745
      Width           =   1125
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Dept. Name"
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
      Left            =   9255
      TabIndex        =   117
      Top             =   2745
      Width           =   1410
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
      Left            =   7935
      TabIndex        =   114
      Top             =   5745
      Width           =   1530
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
      Left            =   9315
      TabIndex        =   113
      Top             =   5745
      Width           =   1275
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
      Left            =   2985
      TabIndex        =   110
      Top             =   1155
      Width           =   1155
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
      Left            =   1590
      TabIndex        =   109
      Top             =   1155
      Width           =   870
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
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
      Left            =   9270
      TabIndex        =   106
      Top             =   2010
      Width           =   1530
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department ID"
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
      Left            =   7890
      TabIndex        =   105
      Top             =   2010
      Width           =   1245
   End
   Begin VB.Label LblType 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   10725
      TabIndex        =   102
      Top             =   7680
      Width           =   435
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand ID"
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
      Left            =   1605
      TabIndex        =   100
      Top             =   7500
      Width           =   765
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Name"
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
      Left            =   2985
      TabIndex        =   99
      Top             =   7500
      Width           =   1050
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Type"
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
      Left            =   10680
      TabIndex        =   95
      Top             =   9120
      Width           =   840
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Name"
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
      Left            =   8655
      TabIndex        =   94
      Top             =   9120
      Width           =   900
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
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
      Left            =   7890
      TabIndex        =   92
      Top             =   1365
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
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
      Left            =   9270
      TabIndex        =   91
      Top             =   1365
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   884
      X2              =   884
      Y1              =   428
      Y2              =   598
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   516
      X2              =   516
      Y1              =   432
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   884
      X2              =   518
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   516
      X2              =   884
      Y1              =   430
      Y2              =   430
   End
   Begin VB.Label Label26 
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
      Left            =   1605
      TabIndex        =   87
      Top             =   2445
      Width           =   705
   End
   Begin VB.Label Label25 
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
      Left            =   2985
      TabIndex        =   86
      Top             =   2445
      Width           =   990
   End
   Begin VB.Label Label24 
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
      Left            =   2985
      TabIndex        =   85
      Top             =   3090
      Width           =   1110
   End
   Begin VB.Label Label22 
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
      Left            =   1605
      TabIndex        =   84
      Top             =   3090
      Width           =   825
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
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
      Left            =   2985
      TabIndex        =   83
      Top             =   4305
      Width           =   915
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
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
      Left            =   1605
      TabIndex        =   82
      Top             =   4305
      Width           =   630
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
      Left            =   1605
      TabIndex        =   81
      Top             =   3690
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
      Left            =   2985
      TabIndex        =   80
      Top             =   3690
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
      Left            =   1605
      TabIndex        =   79
      Top             =   1815
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
      Left            =   2985
      TabIndex        =   78
      Top             =   4965
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
      Left            =   1605
      TabIndex        =   77
      Top             =   4965
      Width           =   660
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group ID"
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
      Left            =   1605
      TabIndex        =   76
      Top             =   6210
      Width           =   780
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
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
      Left            =   2985
      TabIndex        =   75
      Top             =   6210
      Width           =   1065
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group ID"
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
      Left            =   1605
      TabIndex        =   74
      Top             =   6840
      Width           =   1170
   End
   Begin VB.Label Label12 
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
      Left            =   1605
      TabIndex        =   73
      Top             =   5595
      Width           =   1035
   End
   Begin VB.Label Label11 
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
      Left            =   2985
      TabIndex        =   72
      Top             =   5595
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group Name"
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
      Left            =   2985
      TabIndex        =   71
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
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
      Left            =   1605
      TabIndex        =   70
      Top             =   8115
      Width           =   930
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
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
      Left            =   2985
      TabIndex        =   69
      Top             =   8115
      Width           =   1215
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
      Left            =   7905
      TabIndex        =   68
      Top             =   3495
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Left            =   9285
      TabIndex        =   67
      Top             =   3495
      Width           =   1620
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
      Left            =   2985
      TabIndex        =   66
      Top             =   1815
      Width           =   1005
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
      Left            =   10350
      TabIndex        =   65
      Top             =   7965
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
      Left            =   8460
      TabIndex        =   64
      Top             =   7965
      Width           =   885
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Register"
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
      TabIndex        =   63
      Top             =   270
      Width           =   1515
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   5790
      TabIndex        =   62
      Top             =   225
      Visible         =   0   'False
      Width           =   720
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
Attribute VB_Name = "RptSaleRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql, vSelectedParameter As String

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
'        SchAccounts.CmbFilter = "Customers"
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectCustomer = False: Exit Function
        TxtPartyID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    VStrSQL = "Select c.AccountNo, c.AccountName as AccountName, Address, City, p.Description, isnull(p.isWholeSale,1) as isWholeSale" & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " where c.AccountNo in  ( " & (TxtPartyID.Text) & " ) and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and isDetailed = 1 and isLocked = 0"
    
    VStrSQL = VStrSQL + " union all Select EmpID, EmpName as AccountName, Address, City, '', 1 as isWholeSale" & vbCrLf _
         + " from Employees" & vbCrLf _
         + " where EmpID in ( " & (TxtPartyID.Text) & " ) and isLockEmployee = 0"
        
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtPartyName.Text = !AccountName
          FunSelectCustomer = True
          .Close
          Exit Function
      Else
          FunSelectCustomer = False
          .Close
          TxtPartyID.Text = ""
          TxtPartyName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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
    With CN.Execute(VStrSQL)
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

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCode.Text) = "" Then Exit Function
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    VStrSQL = " SELECT p.productid, code, ProductName" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = " & Val(TxtCode.Text) & " or code = '" & TxtCode.Text & "'"
  
   With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !ProductID
         TxtProductName.Text = !ProductName
         FunSelectProduct = True
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtProductID.Text = ""
         TxtCode.Text = ""
         TxtProductName.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnDescription_Click()
If FunSelectDescription(ssButton, False) = True Then
     TxtItemDescID.SetFocus
   Else
      TxtDescriptionID.SetFocus
   End If
End Sub

Private Sub BtnItemCode_Click()
   If FunSelectItemCode(ssButton, True) = True Then
      TxtItemCode.SetFocus
   Else
      TxtDescriptionID.SetFocus
   End If
End Sub

Private Sub BtnItemDesc_Click()
If FunSelectItemDesc(ssButton, False) = True Then
     TxtMemberID.SetFocus
   Else
      TxtItemDescID.SetFocus
   End If
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtItemCode.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnOrganizaton_Click()
   If FunSelectOrganization(ssButton, False) = True Then
      TxtSeasonID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      TxtEmpID.SetFocus
   Else
      TxtPartyID.SetFocus
   End If
End Sub

Private Sub BtnEmpName_Click()
   If FunSelectEmployee(ssButton, False) = True Then
      TxtUserNo.SetFocus
   Else
      TxtEmpID.SetFocus
   End If
End Sub

Private Sub BtnSector_Click()
  If FunSelectSector(ssButton, False) = True Then
      TxtPartyID.SetFocus
   Else
      TxtSectorID.SetFocus
   End If
End Sub

Private Sub BtnSession_Click()
   If FunSelectSession(ssButton, False) = True Then
      If TxtCustomerTypeID.Visible Then TxtCustomerTypeID.SetFocus Else RdoInv.SetFocus
   Else
      TxtSessionID.SetFocus
   End If
End Sub

Private Function FunSelectSession(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSession.Show vbModal, Me
        If SchSession.ParaOutSessionID = "" Then FunSelectSession = False: Exit Function
        TxtSessionID.Text = SchSession.ParaOutSessionID
    End If
    '---------------------------
    If Trim(TxtSessionID.Text) = "" Then Exit Function
    If InStr(1, TxtSessionID.Text, ",") > 0 Then TxtSessionName.Text = "Selected Sessions": Exit Function
    VStrSQL = "Select * FROM Sessions s where SessionID=" & Val(TxtSessionID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtSessionName.Text = !SessionName
          FunSelectSession = True
          .Close
          Exit Function
      Else
          FunSelectSession = False
          .Close
          TxtSessionID.Text = ""
          TxtSessionName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnUpdateStock_Click()
On Error GoTo ErrorHandler
    Me.MousePointer = vbHourglass
    CN.Execute ("ProdUpdatCurrentStockStore")
    CN.Execute ("ProdUpdatCurrentStock")
    Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage

End Sub

Private Sub TxtSessionID_Change()
  On Error GoTo ErrorHandler
   If TxtSessionID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSessionID.Name Then Exit Sub
   If TxtSessionName.Text <> "" Then TxtSessionName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSessionID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSessionID.Name Then Exit Sub
   On Error GoTo ErrorHandler
'   If TxtSessionName.Text <> "All Sessions" Then Exit Sub
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


Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtZoneID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub BtnSubDepartment_Click()
   If FunSelectSubDepartment(ssButton, False) = True Then
     TxtOrganizationID.SetFocus
   Else
      TxtSubDepartmentID.SetFocus
   End If

End Sub

Private Sub BtnUser_Click()
   If FunSelectUser(ssButton, False) = True Then
      TxtCompanyID.SetFocus
   Else
      TxtUserNo.SetFocus
   End If
End Sub

Private Sub BtnZone_Click()
   If FunSelectZone(ssButton, False) = True Then
      TxtSectorID.SetFocus
   Else
      TxtZoneID.SetFocus
   End If
End Sub



Private Sub TxtItemCode_Change()
   If TxtItemCode.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtItemCode.Name Then Exit Sub
   If TxtItemCodeName.Text <> "" Then TxtItemCodeName.Text = ""
End Sub

Private Sub TxtItemCode_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtItemCode.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtItemCode.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectItemCode(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectItemCode(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub TxtProductName_Change()
   If ActiveControl.Name <> TxtProductName.Name Then Exit Sub
   If TxtProductID.Text <> "" Then TxtProductID.Text = ""
End Sub

Private Sub TxtSeason_Click()
   If FunSelectSeason(ssButton, False) = True Then
     TxtSessionID.SetFocus
   Else
      TxtSeasonID.SetFocus
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
    VStrSQL = " Select * FROM Sectors where SectorID in ( " & TxtSectorID.Text & " )"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtSectorName.Text = !SectorName
          FunSelectSector = True
          .Close
          Exit Function
             FunSelectSector = True
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

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
'      TxtCode.Text = ""
      TxtProductID.Text = ""
      TxtProductName.Text = ""
   End If
End Sub

Private Sub TxtCode_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = Not FunSelectProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
    With CN.Execute(VStrSQL)
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

Private Sub BtnCompany_Click()
   If FunSelectCompany(ssButton, False) = True Then
      TxtGroupID.SetFocus
   Else
      TxtCompanyID.SetFocus
   End If
End Sub

Private Sub TxtGroupID_Change()
   If TxtGroupID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   If TxtGroupName.Text <> "" Then TxtGroupName.Text = ""
End Sub

Private Sub TxtGroupID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtGroupID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectGroup(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectGroup(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchGroup.Show vbModal, Me
        If SchGroup.ParaOutGroupID = "" Then FunSelectGroup = False: Exit Function
        TxtGroupID.Text = SchGroup.ParaOutGroupID
    End If
    '---------------------------
    If Trim(TxtGroupID.Text) = "" Then Exit Function
    If Len(TxtGroupID.Text) <= 3 Then
      TxtGroupID.Text = Right("000" + CStr(Val(TxtGroupID.Text)), 3)
    End If
    If TxtGroupID.Text = "" Then FunSelectGroup = False: Exit Function
    VStrSQL = " Select * FROM Groups where GroupID = '" & TxtGroupID.Text & "'"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtGroupName.Text = !GroupName
          FunSelectGroup = True
          .Close
          Exit Function
      Else
          FunSelectGroup = False
          .Close
          TxtGroupID.Text = ""
          TxtGroupName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnGroup_Click()
   If FunSelectGroup(ssButton, False) = True Then
      TxtSubGroupID.SetFocus
   Else
      TxtGroupID.SetFocus
   End If
End Sub

Private Sub TxtPartyID_Change()
   If TxtPartyID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtPartyID.Name Then Exit Sub
   If TxtPartyName.Text <> "" Then TxtPartyName.Text = ""
End Sub

Private Sub TxtPartyID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtPartyID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtPartyID.Text = "" Then Exit Sub
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

Private Function FunSelectEmployee(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
       SchEmployee.Show vbModal, Me
       If SchEmployee.ParaOutEmployeeID = "" Then FunSelectEmployee = False: Exit Function
       TxtEmpID.Text = SchEmployee.ParaOutEmployeeID
   End If
   '---------------------------
   VStrSQL = " Select * FROM Employees where EmpID=" & Val(TxtEmpID.Text)
   With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtEmpName.Text = !EmpName
         FunSelectEmployee = True
         .Close
         Exit Function
         FunSelectEmployee = True
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

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
    If SetReport Then
      If RdoDetail.Value = True Then
          If RdoInv.Value = True Then
              RptReportViewer.Caption = "Sale Detail (" & CmbGroup.Text & ")"
          ElseIf RdoReturn.Value = True Then
              RptReportViewer.Caption = "Sale Return Detail (" & CmbGroup.Text & ")"
          Else
              RptReportViewer.Caption = "Sale & Sale Return Detail (" & CmbGroup.Text & ")"
          End If
      Else
          If RdoInv.Value = True Then
              RptReportViewer.Caption = "Sale Summary (" & CmbGroup.Text & ")"
          ElseIf RdoReturn.Value = True Then
              RptReportViewer.Caption = "Sale Return Summary (" & CmbGroup.Text & ")"
          Else
              RptReportViewer.Caption = "Sale & Sale Return Summary (" & CmbGroup.Text & ")"
          End If
      End If
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
        Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, True) = True Then TxtStoreID.SetFocus
        Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then TxtZoneID.SetFocus
        Case TxtZoneID.Name: If FunSelectZone(ssFunctionKey, True) = True Then TxtSectorID.SetFocus
        Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, True) = True Then TxtPartyID.SetFocus
        Case TxtPartyID.Name: If FunSelectCustomer(ssFunctionKey, True) = True Then TxtEmpID.SetFocus
        Case TxtEmpID.Name: If FunSelectEmployee(ssFunctionKey, True) = True Then TxtUserNo.SetFocus
        Case TxtUserNo.Name: If FunSelectUser(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus
        Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtGroupID.SetFocus
        Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtSubGroupID.SetFocus
        Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then TxtBrandID.SetFocus
        Case TxtBrandID.Name: If FunSelectBrand(ssFunctionKey, True) = True Then TxtCode.SetFocus
        Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then If TxtMemberID.Visible Then TxtMemberID.SetFocus Else TxtCode.SetFocus
        Case TxtMemberID.Name: If FunSelectMember(ssFunctionKey, True) = True Then TxtDepartmentID.SetFocus
        Case TxtDepartmentID.Name: If FunSelectDepartment(ssFunctionKey, True) = True Then TxtOrganizationID.SetFocus
        Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, True) = True Then If TxtCustomerTypeID.Visible Then TxtCustomerTypeID.SetFocus Else RdoInv.SetFocus
        Case TxtCustomerTypeID.Name: If FunSelectCustomerType(ssFunctionKey, True) = True Then RdoInv.SetFocus
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
   SetWindowText Me.hWnd, "Sale Register"
   
   DTPTimeFrom.Value = "00:00:00"
   DTPTimeTo.Value = "23:59:59"
   DTPTimeFrom.Visible = ObjRegistry.TimeWiseReport
   DTPTimeTo.Visible = ObjRegistry.TimeWiseReport
   
'   OptLastPrice.Visible = ObjRegistry.ShowLastPriceOption
'   OptWeightedAvg.Visible = ObjRegistry.ShowWeightedAvgOption
'   OptMovingAvg.Visible = ObjRegistry.ShowMovingAvgOption
'   If ObjRegistry.ShowWeightedAvgOption Then
'      OptWeightedAvg.Value = True
'   ElseIf ObjRegistry.ShowMovingAvgOption Then
'      OptMovingAvg.Value = True
'   Else
'      OptLastPrice.Value = True
'   End If
   
   CmbGroup.Clear
   CmbGroup.AddItem ("Brand Wise")
   CmbGroup.AddItem ("Company Wise")
   CmbGroup.AddItem ("Customer Wise")
   CmbGroup.AddItem ("Customer Invoice Wise")
   CmbGroup.AddItem ("Date Wise")
   CmbGroup.AddItem ("Department Wise")
   CmbGroup.AddItem ("SubDepartment Wise")
   CmbGroup.AddItem ("Employee Wise")
   CmbGroup.AddItem ("Group Wise")
   CmbGroup.AddItem ("Group Customer Wise")
   CmbGroup.AddItem ("Invoice Wise")
   CmbGroup.AddItem ("Invoice Disc Wise")
   CmbGroup.AddItem ("Invoice Wise Trade Offer")
   CmbGroup.AddItem ("Invoice Wise Tax")
   CmbGroup.AddItem ("Product Wise Gate Pass")
'   CmbGroup.AddItem ("Product Wise Gate Pass Multiple Columns")
   CmbGroup.AddItem ("Member Wise")
   CmbGroup.AddItem ("Month Wise")
   CmbGroup.AddItem ("Month Wise Customer")
   CmbGroup.AddItem ("Month Wise Product")
   CmbGroup.AddItem ("Organization Wise")
   CmbGroup.AddItem ("Product Wise")
   CmbGroup.AddItem ("Product Stock Wise")
   CmbGroup.AddItem ("Product Drugs Wise")
   CmbGroup.AddItem ("Product Wise All Fields")
   CmbGroup.AddItem ("Product Wise Specific Fields")
   CmbGroup.AddItem ("Product Wise Loading")
   CmbGroup.AddItem ("Sector Wise")
   CmbGroup.AddItem ("Store Wise")
   CmbGroup.AddItem ("SubGroup Wise")
   CmbGroup.AddItem ("Type Wise")
   CmbGroup.AddItem ("User Wise")
   CmbGroup.AddItem ("Vendor Wise")
   CmbGroup.AddItem ("Zone Wise")
   
   LblType.Visible = ObjRegistry.InvType
   CmbType.Visible = ObjRegistry.InvType
   
   CmbType.Clear
   CmbType.AddItem ""
   With CN.Execute("select * from InvTypes")
      If .RecordCount > 0 Then
         While Not .EOF
            CmbType.AddItem ![InvType]
            .MoveNext
         Wend
      End If
   End With
   
   CmbSortName.Clear
   CmbSortName.AddItem "ProductID"
   CmbSortName.AddItem "ProductName"
   CmbSortType.Clear
   CmbSortType.AddItem "Ascending"
   CmbSortType.AddItem "Descending"
   CmbSortName.ListIndex = 0
   CmbSortType.ListIndex = 0
   
   TxtCustomerTypeID.Visible = ObjRegistry.CustomerTypeVisible
   BtnCustomerType.Visible = ObjRegistry.CustomerTypeVisible
   TxtCustomerType.Visible = ObjRegistry.CustomerTypeVisible
   LblCustomerTypeID.Visible = ObjRegistry.CustomerTypeVisible
   LblCustomerType.Visible = ObjRegistry.CustomerTypeVisible

   TxtSessionID.Text = vSessionID
   FunSelectSession ssValidate, True
   'CmbGroup.AddItem ("Sale Detail (All Wise)")
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
   Set RptSaleRegister = Nothing
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
        
   If OptMovingAvg.Value = True Then
      CN.Execute "exec SPProductPurchase '" & DtpTo.DateValue & "'"
   ElseIf OptLastPrice.Value = True Then
      CN.Execute "exec SPProductPurchase '" & DtpTo.DateValue & "'"
   ElseIf OptWeightedAvg.Value = True Then
      CN.Execute "exec SPAverageCost '" & DtpTo.DateValue & "'"
   End If
'   cn.Execute "exec SPProductPurchase '" & DtpTo.DateValue & "'"
'   cn.Execute "exec SPAverageCost '" & DtpTo.DateValue & "'"

   sSql = "EXEC ProdRptSaleRegister '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "','" & Format(DTPTimeFrom.Value, "hh:mm:ss") & "','" & Format(DTPTimeTo.Value, "hh:mm:ss") & "'," & IIf(Trim(TxtBillIDFrom.Text) = "", "Null", TxtBillIDFrom.Text) & "," & IIf(Trim(TxtBillIDTo.Text) = "", "Null", TxtBillIDTo.Text) & "," & IIf(Trim(TxtManualBillIDFrom.Text) = "", "Null", "'" & TxtManualBillIDFrom.Text & "'") & "," & IIf(Trim(TxtManualBillIDTo.Text) = "", "Null", "'" & TxtManualBillIDTo.Text & "'") & "," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & vbCrLf _
                              & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", "'" & TxtStoreID.Text & "'") & "," & IIf(Trim(TxtZoneID.Text) = "", "Null", "'" & TxtZoneID.Text & "'") & vbCrLf _
                              & "," & IIf(Trim(TxtSectorID.Text) = "", "Null", "'" & TxtSectorID.Text & "'") & "," & IIf(Trim(TxtPartyID.Text) = "", "'", "'" & TxtPartyID.Text) & "'," & IIf(Trim(TxtCustomerTypeID.Text) = "", "Null", TxtCustomerTypeID.Text) & "," & IIf(Trim(TxtEmpID.Text) = "", "Null", "'" & TxtEmpID.Text & "'") & "," & IIf(Trim(TxtUserNo.Text) = "", "Null", "'" & TxtUserNo.Text & "'") & vbCrLf _
                              & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtBrandID.Text) = "", "Null", TxtBrandID.Text) & "," & IIf(Trim(TxtDepartmentID.Text) = "", "Null", TxtDepartmentID.Text) & "," & IIf(Trim(TxtProductID.Text) = "", "Null", Val(TxtProductID.Text)) & vbCrLf _
                              & "," & IIf(Trim(TxtItemCode.Text) = "", "Null", "'" & TxtItemCode.Text & "'") & "," & IIf(Trim(TxtSubDepartmentID.Text) = "", "Null", TxtSubDepartmentID.Text) & "," & IIf(Trim(TxtDescriptionID.Text) = "", "Null", TxtDescriptionID.Text) & "," & IIf(Trim(TxtItemDescID.Text) = "", "Null", TxtItemDescID.Text) & "," & IIf(Trim(TxtSeasonID.Text) = "", "Null", TxtSeasonID.Text) & vbCrLf _
                              & "," & IIf(RdoBoth.Value = True Or RdoNet.Value = True, "Null", IIf(RdoInv.Value = True, 0, 1)) & "," & IIf(RdoAll.Value = True, "Null", IIf(RdoBankCard.Value = True, "'BankCard'", IIf(RdoCredit.Value = True, "'Credit'", IIf(RdoCash.Value = True, "'Cash'", 0)))) & "," & IIf(Trim(TxtProductID.Text) <> "", "Null", "'%" & TxtProductName.Text & "%'") & "," & IIf(Trim(TxtProductID.Text) = Trim(TxtCode.Text), "Null", "'" & TxtCode.Text & "'") & "," & IIf(Trim(CmbType.Text) = "", "null", "'" & CmbType.Text & "'") & vbCrLf _
                              & "," & IIf(Trim(TxtMemberID.Text) = "", "Null", "'" & TxtMemberID.Text & "'") & "," & IIf(Trim(TxtVenderID.Text) = "", "Null", "'" & TxtVenderID.Text & "'") & "," & IIf(Trim(TxtSessionID.Text) = "", "Null", TxtSessionID.Text) & "," & "'" & CmbSortName.Text & " " & CmbSortType.Text & "'"
'                              IIf(RdoBoth.Value = True Or RdoNet.Value = True, "Null", IIf(RdoInv.Value = True, 0, 1))
   
'   Set RsReport = cn.Execute(sSql)
   
   If RdoDetail.Value = True Then
        Select Case CmbGroup.Text
            Case "Organization Wise"
                Set RptReportViewer.Report = New CrptSaleDetailOrgWise
            Case "Store Wise"
                Set RptReportViewer.Report = New CrptSaleDetailStoreWise
            Case "Zone Wise"
                Set RptReportViewer.Report = New CrptSaleDetailZoneWise
            Case "Sector Wise"
                Set RptReportViewer.Report = New CrptSaleDetailSectorWise
            Case "Customer Wise"
                Set RptReportViewer.Report = New CrptSaleDetailCustomerWise
            Case "Employee Wise"
                Set RptReportViewer.Report = New CrptSaleDetailEmpWise
            Case "User Wise"
                Set RptReportViewer.Report = New CrptSaleDetailUserWise
            Case "Company Wise"
                Set RptReportViewer.Report = New CrptSaleDetailCompanyWise
            Case "Group Wise"
                Set RptReportViewer.Report = New CrptSaleDetailGroupWise
            Case "Customer Invoice Wise"
               Set RptReportViewer.Report = New CrptSaleDetailCustomerInvoiceWise
            Case "Group Customer Wise"
                Set RptReportViewer.Report = New CrptSaleDetailGroupCustomer
            Case "SubGroup Wise"
                Set RptReportViewer.Report = New CrptSaleDetailSubGroupWise
            Case "Product Wise"
                Set RptReportViewer.Report = New CrptSaleDetailProductWise
            Case "Product Stock Wise"
                Set RptReportViewer.Report = New CrptSaleDetailProductStockWise
            Case "Product Drugs Wise"
                Set RptReportViewer.Report = New CrptSaleDetailProductDrugsWise
            Case "Product Wise Loading"
                Set RptReportViewer.Report = New CrptSaleDetailProductWiseLoading
            Case "Product Wise Gate Pass"
                Set RptReportViewer.Report = New CrptSaleSummaryProductWiseGatePass
            Case "Product Wise Gate Pass Multiple Columns"
                Set RptReportViewer.Report = New CrptSaleSummaryProductWiseGatePassMC
            Case "Product Wise All Fields"
                Set RptReportViewer.Report = New CrptSaleDetailProductWiseAllFields
            Case "Product Wise Specific Fields"
                Set RptReportViewer.Report = New CrptSaleDetailProductWiseSpecificFields
            Case "Date Wise"
                Set RptReportViewer.Report = New CrptSaleDetailDateWise
            Case "Member Wise"
                Set RptReportViewer.Report = New CrptSaleDetailMemberWise
            Case "Month Wise Customer"
                Set RptReportViewer.Report = New CrptSaleDetailMonthWise
            Case "Month Wise Product"
                Set RptReportViewer.Report = New CrptSaleDetailMonthWise
            Case "Month Wise"
                Set RptReportViewer.Report = New CrptSaleDetailMonthWise
            Case "Invoice Wise"
                Set RptReportViewer.Report = New CrptSaleDetailInvoiceWise
            Case "Invoice Disc Wise"
                Set RptReportViewer.Report = New CrptSaleDetailInvoiceDiscWise
            Case "Invoice Wise Trade Offer"
                Set RptReportViewer.Report = New CrptSaleDetailInvoiceWiseTradeOffer
            Case "Invoice Wise Tax"
                Set RptReportViewer.Report = New CrptSaleDetailInvoiceWiseTax
            Case "Brand Wise"
                Set RptReportViewer.Report = New CrptSaleDetailBrandWise
            Case "Department Wise"
                Set RptReportViewer.Report = New CrptSaleDetailDepartmentWise
            Case "SubDepartment Wise"
                Set RptReportViewer.Report = New CrptSaleDetailSubDepartmentWise
            Case "Vendor Wise"
                Set RptReportViewer.Report = New CrptSaleDetailVendorWise
            Case "Type Wise"
                Set RptReportViewer.Report = New CrptSaleDetailTypeWise
            Case "Sale (All Wise)"
                Set RptReportViewer.Report = New CrptSaleDetailAllWise
        End Select
    
   Else
        Select Case CmbGroup.Text
            Case "Organization Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryOrgWise
            Case "Store Wise"
                 Set RptReportViewer.Report = New CrptSaleSummaryStoreWise
'                  MsgBox "Report is not working.", vbInformation, Me.Caption
'                  Me.MousePointer = vbDefault
'                  Exit Function
            Case "Zone Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryZoneWise
            Case "Sector Wise"
                Set RptReportViewer.Report = New CrptSaleSummarySectorWise
            Case "Customer Wise"
'                Set RptReportViewer.Report = New CrptSaleSummaryCustomerWise
                Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\SaleReports\CrptSaleSummaryCustomerWise.rpt")
            Case "Customer Invoice Wise"
'               Set RptReportViewer.Report = New CrptSaleSummaryCustomerInvoiceWise
                 Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\SaleReports\CrptSaleSummaryCustomerInvoiceWise.rpt")
            Case "Employee Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryEmpWise
            Case "User Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryUserWise
            Case "Company Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryCompanyWise
            Case "Group Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryGroupWise
            Case "Group Customer Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryGroupCustomerWise
            Case "SubGroup Wise"
                Set RptReportViewer.Report = New CrptSaleSummarySubGroupWise
            Case "Product Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryProductWise
            Case "Product Stock Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryProductStockWise
            Case "Product Wise Loading"
                Set RptReportViewer.Report = New CrptSaleDetailProductWiseLoading
            Case "Product Wise Gate Pass"
                Set RptReportViewer.Report = New CrptSaleSummaryProductWiseGatePass
            Case "Product Wise Gate Pass Multiple Columns"
                Set RptReportViewer.Report = New CrptSaleSummaryProductWiseGatePassMC
            Case "Date Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryDateWise
            Case "Member Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryMemberWise
            Case "Month Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryMonthWise
            Case "Month Wise Product"
                Set RptReportViewer.Report = New CrptSaleSummaryMonthWiseProduct
            Case "Month Wise Customer"
                Set RptReportViewer.Report = New CrptSaleSummaryMonthWiseCustomer
            Case "Invoice Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryInvoiceWise
            Case "Invoice Disc Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryInvoiceDiscWise
            Case "Brand Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryBrandWise
            Case "Department Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryDepartmentWise
            Case "SubDepartment Wise"
                Set RptReportViewer.Report = New CrptSaleSummarySubDepartmentWise
            Case "Vendor Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryVendorWise
            Case "Type Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryTypeWise
        End Select
    End If
           
    Set RsReport = CN.Execute(sSql)
    
    If RsReport.BOF Then
        MsgBox "No record exists.", vbInformation, Me.Caption
        Me.MousePointer = vbDefault
        Exit Function
    End If
    RptReportViewer.Report.Database.SetDataSource RsReport
    
    If RdoDetail.Value = True Then
        If RdoInv.Value = True Then
            RptReportViewer.Report.ReportTitle = "Sale Detail (" & CmbGroup.Text & ")"
        ElseIf RdoReturn.Value = True Then
            RptReportViewer.Report.ReportTitle = "Sale Return Detail (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Report.ReportTitle = "Sale & Sale Return Detail (" & CmbGroup.Text & ")"
        End If
    Else
        If RdoInv.Value = True Then
            RptReportViewer.Report.ReportTitle = "Sale Summary (" & CmbGroup.Text & ")"
        ElseIf RdoReturn.Value = True Then
            RptReportViewer.Report.ReportTitle = "Sale Return Summary (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Report.ReportTitle = "Sale & Sale Return Summary (" & CmbGroup.Text & ")"
        End If
        
    End If
   
    
    If RdoNet.Value Then
      RptReportViewer.Report.DeleteGroup 1
    End If
    Call SubSelectedParameterFields
    RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName & IIf(ObjRegistry.CompanyCity = "", "", " - " & ObjRegistry.CompanyCity)
    RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress)
    RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", "Phone # " & ObjRegistry.CompanyPhoneNo)
    RptReportViewer.Report.ParameterFields(5).AddCurrentValue "Date From " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To " & Format(DtpTo.DateValue, "dd/MM/yyyy")
    RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
    RptReportViewer.Report.ParameterFields(6).AddCurrentValue vSelectedParameter
    RptReportViewer.Report.ParameterFields(7).AddCurrentValue IIf(Val(TxtBillIDTo.Text) = 0, "", "Bill No. " & Val(TxtBillIDFrom.Text) & " - " & Val(TxtBillIDTo.Text))
    RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
    
'    If ObjRegistry.IsPortrait = False Then RptReportViewer.Report.PaperOrientation = crLandscape
    
    If CmbGroup.Text = "Month Wise Product" Or CmbGroup.Text = "Product Wise Loading" Or CmbGroup.Text = "Product Wise Gate Pass Multiple Columns" Or CmbGroup.Text = "Product Wise Gate Pass" Then
        RptReportViewer.Report.PaperOrientation = crPortrait
    ElseIf CmbGroup.Text = "Product Wise All Fields" Then
        RptReportViewer.Report.PaperSize = crPaperA3
    Else
        RptReportViewer.Report.PaperOrientation = crLandscape
    End If
    
    'RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    'RptReportViewer.Report.PaperOrientation = crLandscape
    SetReport = True
    Me.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
    Me.MousePointer = vbDefault
End Function

Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectStore = False: Exit Function
        TxtStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    If Trim(TxtStoreID.Text) = "" Then Exit Function
    If Len(TxtStoreID.Text) <= 3 Then
      TxtStoreID.Text = Right("000" + CStr(Val(TxtStoreID.Text)), 3)
    End If
    If TxtStoreID.Text = "" Then FunSelectStore = False: Exit Function
    VStrSQL = " Select StoreName FROM Stores where StoreID='" & TxtStoreID.Text & "'"
    With CN.Execute(VStrSQL)
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

Private Sub TxtSubGroupID_Change()
   If TxtSubGroupID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSubGroupID.Name Then Exit Sub
   If TxtSubGroupName.Text <> "" Then TxtSubGroupName.Text = ""
End Sub

Private Sub TxtSubGroupID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtSubGroupID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSubGroupID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSubGroup(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSubGroup(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSubGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSubGroup.Show vbModal, Me
        If SchSubGroup.ParaOutSubGroupID = "" Then FunSelectSubGroup = False: Exit Function
        TxtSubGroupID.Text = SchSubGroup.ParaOutSubGroupID
    End If
    '---------------------------
    VStrSQL = " Select * FROM SubGroups where SubGroupID = " & Val(TxtSubGroupID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtSubGroupName.Text = !SubGroupName
          FunSelectSubGroup = True
          .Close
          Exit Function
      Else
          FunSelectSubGroup = False
          .Close
          TxtSubGroupID.Text = ""
          TxtSubGroupName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnSubGroup_Click()
   If FunSelectSubGroup(ssButton, False) = True Then
      TxtBrandID.SetFocus
   Else
      TxtSubGroupID.SetFocus
   End If
End Sub

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

Private Function FunSelectUser(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchUser.Show vbModal, Me
        If SchUser.ParaOutUserNo = "" Then FunSelectUser = False: Exit Function
        TxtUserNo.Text = SchUser.ParaOutUserNo
    End If
    '---------------------------
    VStrSQL = " Select * FROM Users where UserNo=" & Val(TxtUserNo.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtUserName.Text = !UserName
          FunSelectUser = True
          .Close
          Exit Function
             FunSelectUser = True
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
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtZoneName.Text = !ZoneName
         FunSelectZone = True
         .Close
         Exit Function
         FunSelectZone = True
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

Private Sub BtnBrand_Click()
   If FunSelectBrand(ssButton, False) = True Then
      TxtCode.SetFocus
   Else
      TxtBrandID.SetFocus
   End If
End Sub

Private Function FunSelectBrand(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchBrand.Show vbModal, Me
        If SchBrand.ParaOutBrandID = "" Then FunSelectBrand = False: Exit Function
        TxtBrandID.Text = SchBrand.ParaOutBrandID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Brands where BrandID=" & Val(TxtBrandID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtBrandName.Text = !BrandName
          FunSelectBrand = True
          .Close
          Exit Function
      Else
          FunSelectBrand = False
          .Close
          TxtBrandID.Text = ""
          TxtBrandName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtBrandID_Change()
   If TxtBrandID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtBrandID.Name Then Exit Sub
   If TxtBrandName.Text <> "" Then TxtBrandName.Text = ""
End Sub

Private Sub TxtBrandID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtBrandID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtBrandID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBrand(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBrand(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnMember_Click()
   If FunSelectMember(ssButton, False) = True Then
      TxtDepartmentID.SetFocus
   Else
      TxtMemberID.SetFocus
   End If
End Sub

Private Function FunSelectMember(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchMember.Show vbModal, Me
        If SchMember.ParaOutMemberID = "" Then FunSelectMember = False: Exit Function
        TxtMemberID.Text = SchMember.ParaOutMemberID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Members where MemberID=" & Val(TxtMemberID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtMemberName.Text = !MemberName
          FunSelectMember = True
          .Close
          Exit Function
      Else
          FunSelectMember = False
          .Close
          TxtMemberID.Text = ""
          TxtMemberName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtMemberID_Change()
   If TxtMemberID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtMemberID.Name Then Exit Sub
   If TxtMemberName.Text <> "" Then TxtMemberName.Text = ""
End Sub

Private Sub TxtMemberID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtMemberID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtMemberID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectMember(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectMember(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectDepartment(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchDepartment.Show vbModal, Me
        If SchDepartment.ParaOutDepartmentID = "" Then FunSelectDepartment = False: Exit Function
        TxtDepartmentID.Text = SchDepartment.ParaOutDepartmentID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Departments where DepartmentID=" & Val(TxtDepartmentID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtDepartmentName.Text = !Department
          FunSelectDepartment = True
          .Close
          Exit Function
      Else
          FunSelectDepartment = False
          .Close
          TxtDepartmentID.Text = ""
          TxtDepartmentName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnDepartment_Click()
   If FunSelectDepartment(ssButton, False) = True Then
      TxtSubDepartmentID.SetFocus
   Else
      TxtDepartmentID.SetFocus
   End If
End Sub

Private Sub TxtDepartmentID_Change()
   If ActiveControl.Name <> TxtDepartmentID.Name Then Exit Sub
   If TxtDepartmentName.Text <> "" Then TxtDepartmentName.Text = ""
End Sub

Private Sub TxtDepartmentID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtDepartmentID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtDepartmentID.Text = "" Then Exit Sub
   If TxtDepartmentName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectDepartment(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectDepartment(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
      TxtStoreID.SetFocus
   Else
      TxtVenderID.SetFocus
   End If
End Sub

Private Function FunSelectVender(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
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
    VStrSQL = " Select c.* FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " where BarCode = '" & (TxtVenderID.Text) & "' or (c.AccountNo = '" & (TxtVenderID.Text) & "' and (c.AccountNo like '6%') and c.isDetailed = 1 and c.isLocked = 0)"
    With CN.Execute(VStrSQL)
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

Private Function FunSelectCustomerType(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCustomerType.Show vbModal, Me
        If SchCustomerType.ParaOutID = "" Then FunSelectCustomerType = False: Exit Function
        TxtCustomerTypeID.Text = SchCustomerType.ParaOutID
    End If
    '---------------------------
    VStrSQL = " Select * FROM CustomerTypes where CustomerTypeID = '" & TxtCustomerTypeID.Text & "'"
    With CN.Execute(VStrSQL)
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

Private Sub BtnCustomerType_Click()
   If FunSelectCustomerType(ssButton, False) = True Then
      RdoInv.SetFocus
   Else
      If TxtCustomerTypeID.Enabled Then TxtCustomerTypeID.SetFocus
   End If
End Sub

Private Function FunSelectItemCode(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchItemCode.Show vbModal, Me
        If SchItemCode.ParaOutItemCode = "" Then FunSelectItemCode = False: Exit Function
        TxtItemCode.Text = SchItemCode.ParaOutItemCode
    End If
    '---------------------------
    VStrSQL = " Select * FROM Products where ItemCode=" & Val(TxtItemCode.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtItemCodeName.Text = !ProductName
          FunSelectItemCode = True
          .Close
          Exit Function
      Else
          FunSelectItemCode = False
          .Close
          TxtItemCode.Text = ""
          TxtItemCodeName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectDescription(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchDescription.Show vbModal, Me
        If SchDescription.ParaOutDescriptionID = "" Then FunSelectDescription = False: Exit Function
        TxtDescriptionID.Text = SchDescription.ParaOutDescriptionID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Descriptions where DescriptionID=" & Val(TxtDescriptionID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtDescriptionName.Text = !DescriptionName
          FunSelectDescription = True
          .Close
          Exit Function
      Else
          FunSelectDescription = False
          .Close
          TxtDescriptionID.Text = ""
          TxtDescriptionName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Sub TxtDescriptionID_Change()
   If TxtDescriptionID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtDescriptionID.Name Then Exit Sub
   If TxtDescriptionName.Text <> "" Then TxtDescriptionName.Text = ""
End Sub

Private Sub TxtDescriptionID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtDescriptionID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtDescriptionID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectDescription(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectDescription(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectItemDesc(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchItemDesc.Show vbModal, Me
        If SchItemDesc.ParaOutItemDescID = "" Then FunSelectItemDesc = False: Exit Function
        TxtItemDescID.Text = SchItemDesc.ParaOutItemDescID
    End If
    '---------------------------
    VStrSQL = " Select * FROM ItemDescription where ItemDescID=" & Val(TxtItemDescID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtItemDescName.Text = !ItemDescName
          FunSelectItemDesc = True
          .Close
          Exit Function
      Else
          FunSelectItemDesc = False
          .Close
          TxtItemDescID.Text = ""
          TxtItemDescName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Sub TxtItemDescID_Change()
   If TxtItemDescID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtItemDescID.Name Then Exit Sub
   If TxtItemDescName.Text <> "" Then TxtItemDescName.Text = ""
End Sub

Private Sub TxtItemDescID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtItemDescID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtItemDescID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectItemDesc(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectItemDesc(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Function FunSelectSubDepartment(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSubDepartments.Show vbModal, Me
        If SchSubDepartments.ParaOutSubDepartmentID = "" Then FunSelectSubDepartment = False: Exit Function
        TxtSubDepartmentID.Text = SchSubDepartments.ParaOutSubDepartmentID
    End If
    '---------------------------
    VStrSQL = " Select * FROM SubDepartments where SubDepartmentID=" & Val(TxtSubDepartmentID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtSubDepartmentName.Text = !SubDepartmentName
          FunSelectSubDepartment = True
          .Close
          Exit Function
      Else
          FunSelectSubDepartment = False
          .Close
          TxtSubDepartmentID.Text = ""
          TxtSubDepartmentName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtSubDepartmentID_Change()
   If ActiveControl.Name <> TxtSubDepartmentID.Name Then Exit Sub
   If TxtSubDepartmentName.Text <> "" Then TxtSubDepartmentName.Text = ""
End Sub
Private Sub TxtSubDepartmentID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSubDepartmentID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSubDepartmentName.Text <> "" Then Exit Sub
   If TxtSubDepartmentID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSubDepartment(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSubDepartment(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Function FunSelectSeason(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSeasons.Show vbModal, Me
        If SchSeasons.ParaOutSeasonID = "" Then FunSelectSeason = False: Exit Function
        TxtSeasonID.Text = SchSeasons.ParaOutSeasonID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Seasons where SeasonID=" & Val(TxtSeasonID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtSeasonName.Text = !SeasonName
          FunSelectSeason = True
          .Close
          Exit Function
      Else
          FunSelectSeason = False
          .Close
          TxtSeasonID.Text = ""
          TxtSeasonName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtSeasonID_Change()
   If TxtSeasonID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSeasonID.Name Then Exit Sub
   If TxtSeasonName.Text <> "" Then TxtSeasonName.Text = ""
End Sub

Private Sub TxtSeasonID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtSeasonID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSeasonID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSeason(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSeason(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubSelectedParameterFields()
   On Error GoTo ErrorHandler
   vSelectedParameter = ""
   If TxtVenderName.Text <> "" Then vSelectedParameter = "Vendor: " & TxtVenderID.Text & " - " & TxtVenderName.Text & " " & vbCrLf
   If TxtStoreName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Store: " & TxtStoreID.Text & " - " & TxtStoreName.Text & " " & vbCrLf
   If TxtSectorName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Sector: " & TxtSectorID.Text & " - " & TxtSectorName.Text & " " & vbCrLf
   If TxtPartyName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Party: " & TxtPartyID.Text & " - " & TxtPartyName.Text & " " & vbCrLf
   If TxtEmpName.Text <> "" Then vSelectedParameter = vSelectedParameter & "EmpLoyee: " & TxtEmpID.Text & " - " & TxtEmpName.Text & " " & vbCrLf
   If TxtUserName.Text <> "" Then vSelectedParameter = vSelectedParameter & "User: " & TxtUserNo.Text & " - " & TxtUserName.Text & " " & vbCrLf
   If TxtCompanyName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Company: " & TxtCompanyID.Text & " - " & TxtCompanyName.Text & " " & vbCrLf
   If TxtGroupName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Group: " & TxtGroupID.Text & " - " & TxtGroupName.Text & " " & vbCrLf
   If TxtSubGroupName.Text <> "" Then vSelectedParameter = vSelectedParameter & "SubGroup: " & TxtSubGroupID.Text & " - " & TxtSubGroupName.Text & " " & vbCrLf
   If TxtBrandName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Brand: " & TxtBrandID.Text & " - " & TxtBrandName.Text & " " & vbCrLf
   If TxtProductName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Product: " & TxtProductID.Text & " - " & TxtProductName.Text & " " & vbCrLf
   If TxtItemCodeName.Text <> "" Then vSelectedParameter = vSelectedParameter & "ItemCode: " & TxtItemCode.Text & " - " & TxtItemCodeName.Text & " " & vbCrLf
   If TxtDescriptionName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Description: " & TxtDescriptionID.Text & " - " & TxtDescriptionName.Text & " " & vbCrLf
   If TxtItemDescName.Text <> "" Then vSelectedParameter = vSelectedParameter & "ItemDesc: " & TxtItemDescID.Text & " - " & TxtItemDescName.Text & " " & vbCrLf
   If TxtMemberName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Member: " & TxtMemberID.Text & " - " & TxtMemberName.Text & " " & vbCrLf
   If TxtDepartmentName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Department: " & TxtDepartmentID.Text & " - " & TxtDepartmentName.Text & " " & vbCrLf
   If TxtSubDepartmentName.Text <> "" Then vSelectedParameter = vSelectedParameter & "SubDepartment: " & TxtSubDepartmentID.Text & " - " & TxtSubDepartmentName.Text & " " & vbCrLf
   If TxtOrganizatonName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Organization: " & TxtOrganizationID.Text & " - " & TxtOrganizatonName.Text & " " & vbCrLf
   If TxtSeasonName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Season: " & TxtSeasonID.Text & " - " & TxtSeasonName.Text & " " & vbCrLf
   If TxtSessionName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Session: " & TxtSessionID.Text & " - " & TxtSessionName.Text & " "
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

