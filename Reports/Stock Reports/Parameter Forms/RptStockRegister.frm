VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptStockRegister 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptStockRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9855
      TabIndex        =   90
      Top             =   8280
      Width           =   3690
      Begin VB.OptionButton OptLastPrice 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Last Price"
         Height          =   195
         Left            =   2565
         TabIndex        =   93
         Top             =   45
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton OptWeightedAvg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Weighted Avg"
         Height          =   195
         Left            =   1215
         TabIndex        =   92
         ToolTipText     =   "Weighted Mean"
         Top             =   45
         Width           =   1350
      End
      Begin VB.OptionButton OptMovingAvg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Moving Avg"
         Height          =   195
         Left            =   45
         TabIndex        =   91
         ToolTipText     =   "Simple Moving Average"
         Top             =   45
         Visible         =   0   'False
         Width           =   1170
      End
   End
   Begin VB.ComboBox CmbSortType 
      Height          =   315
      ItemData        =   "RptStockRegister.frx":0ECA
      Left            =   11655
      List            =   "RptStockRegister.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   7020
      Width           =   1275
   End
   Begin VB.ComboBox CmbSortName 
      Height          =   315
      ItemData        =   "RptStockRegister.frx":0ECE
      Left            =   9765
      List            =   "RptStockRegister.frx":0ED0
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   7020
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11190
      TabIndex        =   14
      Top             =   4740
      Width           =   2250
      Begin VB.OptionButton RdoDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detail"
         Height          =   255
         Left            =   105
         TabIndex        =   3
         Top             =   10
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton RdoSummary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary"
         Height          =   255
         Left            =   1230
         TabIndex        =   4
         Top             =   10
         Width           =   960
      End
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   9285
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4740
      Width           =   1950
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   12420
      TabIndex        =   8
      Top             =   8925
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
      MICON           =   "RptStockRegister.frx":0ED2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9690
      TabIndex        =   6
      Top             =   8925
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
      MICON           =   "RptStockRegister.frx":0EEE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11025
      TabIndex        =   7
      Top             =   8925
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
      MICON           =   "RptStockRegister.frx":0F0A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   10515
      TabIndex        =   9
      Top             =   7785
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpDate 
      Height          =   315
      Left            =   10710
      TabIndex        =   5
      Top             =   5490
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
   Begin SITextBox.Txt TxtMinLimit 
      Height          =   315
      Left            =   10530
      TabIndex        =   0
      Top             =   6210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtMaxLimit 
      Height          =   315
      Left            =   11430
      TabIndex        =   1
      Top             =   6210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   2010
      TabIndex        =   20
      Top             =   4590
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
      Masked          =   1
      IntegralPoint   =   15
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   3390
      TabIndex        =   21
      Top             =   4575
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
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   2010
      TabIndex        =   22
      Top             =   2655
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
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   3390
      TabIndex        =   23
      Tag             =   "nc"
      Top             =   2640
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
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   2010
      TabIndex        =   24
      Top             =   1995
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
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   3390
      TabIndex        =   25
      Tag             =   "nc"
      Top             =   1995
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
      Left            =   2010
      TabIndex        =   26
      Top             =   5235
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
      Masked          =   1
      IntegralPoint   =   15
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   3390
      TabIndex        =   27
      Tag             =   "nc"
      Top             =   5235
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
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   2010
      TabIndex        =   28
      Top             =   3300
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
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   3390
      TabIndex        =   29
      Tag             =   "nc"
      Top             =   3300
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
   Begin SITextBox.Txt TxtVenderName 
      Height          =   315
      Left            =   3390
      TabIndex        =   30
      Tag             =   "nc"
      Top             =   5925
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
      Left            =   2010
      TabIndex        =   31
      Top             =   5925
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
      Masked          =   1
      IntegralPoint   =   15
   End
   Begin SITextBox.Txt TxtBrandID 
      Height          =   315
      Left            =   2010
      TabIndex        =   38
      Top             =   3975
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
      Left            =   3390
      TabIndex        =   39
      Tag             =   "nc"
      Top             =   3975
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
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   2010
      TabIndex        =   40
      Top             =   1365
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
      Left            =   3390
      TabIndex        =   41
      Tag             =   "nc"
      Top             =   1365
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
   Begin SITextBox.Txt TxtDepartmentID 
      Height          =   315
      Left            =   2010
      TabIndex        =   35
      Top             =   8895
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
   Begin SITextBox.Txt TxtDepartmentName 
      Height          =   315
      Left            =   3390
      TabIndex        =   58
      Tag             =   "nc"
      Top             =   8895
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
   Begin SITextBox.Txt TxtSubDepartmentID 
      Height          =   315
      Left            =   2010
      TabIndex        =   36
      Top             =   9630
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
   Begin SITextBox.Txt TxtSubDepartmentName 
      Height          =   315
      Left            =   3390
      TabIndex        =   59
      Tag             =   "nc"
      Top             =   9630
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
   Begin SITextBox.Txt TxtSeasonID 
      Height          =   315
      Left            =   2010
      TabIndex        =   37
      Top             =   10395
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
   Begin SITextBox.Txt TxtSeasonName 
      Height          =   315
      Left            =   3390
      TabIndex        =   60
      Tag             =   "nc"
      Top             =   10395
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
      Left            =   2010
      TabIndex        =   32
      Top             =   6660
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
   Begin SITextBox.Txt TxtItemCodeName 
      Height          =   315
      Left            =   3390
      TabIndex        =   61
      Top             =   6660
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
   Begin SITextBox.Txt TxtItemDescID 
      Height          =   315
      Left            =   2010
      TabIndex        =   34
      Top             =   8100
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
   Begin SITextBox.Txt TxtItemDescName 
      Height          =   315
      Left            =   3390
      TabIndex        =   62
      Tag             =   "nc"
      Top             =   8100
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
   Begin SITextBox.Txt TxtDescriptionID 
      Height          =   315
      Left            =   2010
      TabIndex        =   33
      Top             =   7425
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
   Begin SITextBox.Txt TxtDescriptionName 
      Height          =   315
      Left            =   3390
      TabIndex        =   63
      Tag             =   "nc"
      Top             =   7425
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
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   3030
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   4575
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
      MICON           =   "RptStockRegister.frx":0F26
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   76
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5235
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
      MICON           =   "RptStockRegister.frx":0F42
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   77
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2640
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
      MICON           =   "RptStockRegister.frx":0F5E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnCompany 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   78
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1995
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
      MICON           =   "RptStockRegister.frx":0F7A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSubGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   79
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3300
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
      MICON           =   "RptStockRegister.frx":0F96
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnVender 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   80
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5925
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
      MICON           =   "RptStockRegister.frx":0FB2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnBrand 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   81
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3990
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
      MICON           =   "RptStockRegister.frx":0FCE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   82
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1365
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
      MICON           =   "RptStockRegister.frx":0FEA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   83
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   8895
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
      MICON           =   "RptStockRegister.frx":1006
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSubDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   84
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   9630
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
      MICON           =   "RptStockRegister.frx":1022
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton TxtSeason 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   85
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   10395
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
      MICON           =   "RptStockRegister.frx":103E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnItemCode 
      Height          =   330
      Left            =   3030
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   6660
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
      MICON           =   "RptStockRegister.frx":105A
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnItemDesc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   87
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   8100
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
      MICON           =   "RptStockRegister.frx":1076
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDescription 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3030
      TabIndex        =   88
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7440
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
      MICON           =   "RptStockRegister.frx":1092
      BC              =   14737632
      FC              =   0
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
      Left            =   2010
      TabIndex        =   89
      Top             =   8685
      Width           =   1245
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
      Left            =   3390
      TabIndex        =   74
      Top             =   8685
      Width           =   1530
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
      Left            =   3390
      TabIndex        =   73
      Top             =   9420
      Width           =   1410
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
      Left            =   2010
      TabIndex        =   72
      Top             =   9420
      Width           =   1125
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
      Left            =   3390
      TabIndex        =   71
      Top             =   10185
      Width           =   1185
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
      Left            =   2010
      TabIndex        =   70
      Top             =   10185
      Width           =   900
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
      Left            =   2010
      TabIndex        =   69
      Top             =   6450
      Width           =   870
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
      Left            =   3390
      TabIndex        =   68
      Top             =   6450
      Width           =   990
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
      Left            =   3390
      TabIndex        =   67
      Top             =   7890
      Width           =   1470
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
      Left            =   2010
      TabIndex        =   66
      Top             =   7890
      Width           =   930
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
      Left            =   3390
      TabIndex        =   65
      Top             =   7215
      Width           =   1515
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
      Left            =   2010
      TabIndex        =   64
      Top             =   7215
      Width           =   1230
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
      Left            =   3390
      TabIndex        =   57
      Top             =   4380
      Width           =   1215
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
      Left            =   2010
      TabIndex        =   56
      Top             =   4380
      Width           =   930
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
      Left            =   3390
      TabIndex        =   55
      Top             =   3075
      Width           =   1455
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
      Left            =   3390
      TabIndex        =   54
      Top             =   1800
      Width           =   1320
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
      Left            =   2010
      TabIndex        =   53
      Top             =   1800
      Width           =   1035
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
      Left            =   2010
      TabIndex        =   52
      Top             =   3075
      Width           =   1170
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
      Left            =   3390
      TabIndex        =   51
      Top             =   2445
      Width           =   1065
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
      Left            =   2010
      TabIndex        =   50
      Top             =   2445
      Width           =   780
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
      Left            =   2010
      TabIndex        =   49
      Top             =   5040
      Width           =   720
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
      Left            =   3390
      TabIndex        =   48
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label Label1 
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
      Left            =   2010
      TabIndex        =   47
      Top             =   5700
      Width           =   870
   End
   Begin VB.Label Label17 
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
      Left            =   3390
      TabIndex        =   46
      Top             =   5700
      Width           =   1155
   End
   Begin VB.Label Label5 
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
      Left            =   2010
      TabIndex        =   45
      Top             =   3750
      Width           =   765
   End
   Begin VB.Label Label14 
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
      Left            =   3390
      TabIndex        =   44
      Top             =   3750
      Width           =   1050
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
      Left            =   3390
      TabIndex        =   43
      Top             =   1140
      Width           =   1590
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
      Left            =   2010
      TabIndex        =   42
      Top             =   1140
      Width           =   1290
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
      Left            =   11655
      TabIndex        =   19
      Top             =   6795
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
      Left            =   9765
      TabIndex        =   18
      Top             =   6795
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Limit"
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
      Left            =   10845
      TabIndex        =   15
      Top             =   5940
      Width           =   960
   End
   Begin VB.Label Label2 
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
      Left            =   9285
      TabIndex        =   13
      Top             =   4530
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      TabIndex        =   12
      Top             =   5265
      Width           =   420
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Register"
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
      TabIndex        =   11
      Top             =   270
      Width           =   1680
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   10515
      TabIndex        =   10
      Top             =   7590
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
Attribute VB_Name = "RptStockRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql As String

Private Sub BtnBrand_Click()
   If FunSelectBrand(ssButton, False) = True Then
      TxtCode.SetFocus
   Else
      TxtBrandID.SetFocus
   End If
End Sub

Private Sub BtnCompany_Click()
   If FunSelectCompany(ssButton, False) = True Then
      TxtGroupID.SetFocus
   Else
      TxtCompanyID.SetFocus
   End If
End Sub

Private Sub BtnGroup_Click()
   If FunSelectGroup(ssButton, False) = True Then
      TxtSubGroupID.SetFocus
   Else
      TxtGroupID.SetFocus
   End If
End Sub

Private Sub BtnDescription_Click()
If FunSelectDescription(ssButton, False) = True Then
     TxtItemDescID.SetFocus
   Else
      TxtDescriptionID.SetFocus
   End If
End Sub
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
          TxtItemDescName.Text = !ItemdescName
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

Private Sub BtnItemCode_Click()
   If FunSelectItemCode(ssButton, True) = True Then
      TxtItemCode.SetFocus
   Else
      TxtDescriptionID.SetFocus
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


Private Sub BtnItemDesc_Click()
If FunSelectItemDesc(ssButton, False) = True Then
     TxtDepartmentID.SetFocus
   Else
      TxtItemDescID.SetFocus
   End If
End Sub
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

Private Sub BtnSubDepartment_Click()
   If FunSelectSubDepartment(ssButton, False) = True Then
     TxtOrganizationID.SetFocus
   Else
      TxtSubDepartmentID.SetFocus
   End If
End Sub
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

Private Sub TxtSeason_Click()
   If FunSelectSeason(ssButton, False) = True Then
     RdoDetail.SetFocus
   Else
      TxtSeasonID.SetFocus
   End If
End Sub
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


Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtStoreID.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtVenderID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub BtnSubGroup_Click()
   If FunSelectSubGroup(ssButton, False) = True Then
      TxtBrandID.SetFocus
   Else
      TxtSubGroupID.SetFocus
   End If
End Sub

Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
      CmbGroup.SetFocus
   Else
      TxtVenderID.SetFocus
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
    VStrSQL = "Select * FROM Brands where BrandID = " & Val(TxtBrandID.Text)
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
           + " where p.productid = " & Val(TxtCode.Text) & " or code='" & TxtCode.Text & "'"
  
   With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
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

Private Sub CmbGroup_Click()
   If CmbGroup.Text = "Product Wise" Then
      RdoDetail.Visible = False
      RdoSummary.Visible = True
      RdoSummary.Value = True
   ElseIf CmbGroup.Text = "Product Wise All Fields" Then
      RdoDetail.Visible = True
      RdoDetail.Value = True
      RdoSummary.Visible = False
   Else
      RdoDetail.Visible = True
      RdoSummary.Visible = True
   End If
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtCode.Text = ""
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

Private Sub TxtProductName_Change()
   If ActiveControl.Name <> TxtProductName.Name Then Exit Sub
   If TxtProductID.Text <> "" Then TxtProductID.Text = ""
End Sub

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

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   On Error GoTo ErrorHandler
   If SetReport Then
     If RdoDetail.Value = True Then
        RptReportViewer.Caption = "Stock Register Detail (" & CmbGroup.Text & ")"
     Else
        RptReportViewer.Caption = "Stock Register Summary (" & CmbGroup.Text & ")"
     End If
     RptReportViewer.Show vbModal
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus
         Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtGroupID.SetFocus
         Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtSubGroupID.SetFocus
         Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then TxtBrandID.SetFocus
         Case TxtBrandID.Name: If FunSelectBrand(ssFunctionKey, True) = True Then TxtCode.SetFocus
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtStoreID.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then TxtVenderID.SetFocus
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, True) = True Then CmbGroup.SetFocus
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
   SetWindowText Me.hWnd, "Stock Register"
   
   TxtOrganizationID.Visible = ObjRegistry.ShowOrganizationWiseStock
   BtnOrganization.Visible = ObjRegistry.ShowOrganizationWiseStock
   TxtOrganizationName.Visible = ObjRegistry.ShowOrganizationWiseStock
   LblOrganizationID.Visible = ObjRegistry.ShowOrganizationWiseStock
   LblOrganizationName.Visible = ObjRegistry.ShowOrganizationWiseStock
   
   CmbSortName.Clear
   CmbSortName.AddItem "ProductName"
   CmbSortName.AddItem "ProductID"
   CmbSortType.Clear
   CmbSortType.AddItem "Ascending"
   CmbSortType.AddItem "Descending"
   
   CmbSortName.ListIndex = 0
   
   CmbGroup.AddItem ("Company Wise")
   CmbGroup.AddItem ("Group Wise")
   CmbGroup.AddItem ("Product Wise All Fields")
   CmbGroup.AddItem ("SubGroup Wise")
   CmbGroup.AddItem ("Brand Wise")
   CmbGroup.AddItem ("Product Wise")
   CmbGroup.AddItem ("Store Wise")
   CmbGroup.AddItem ("Vendor Wise")
   'CmbGroup.AddItem ("Sale Detail (All Wise)")
   CmbGroup.ListIndex = 0
   vQty = True
   vShowRetailPrice = Not ObjRegistry.ShowRetailPriceStockRegister
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
   Set RptStockRegister = Nothing
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
         Case "Brand Wise"
'            Set RptReportViewer.Report = New CrpStockValueRegisterDetailBrandWise
             Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\StockReports\CrpStockValueRegisterDetailBrandWise.rpt")
         Case "Company Wise"
'            Set RptReportViewer.Report = New CrpStockValueRegisterDetailCompanyWise
             Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\StockReports\CrpStockValueRegisterDetailCompanyWise.rpt")
         Case "Group Wise"
             'Set RptReportViewer.Report = New CrpStockValueRegisterDetailGroupWise
             Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\StockReports\CrpStockValueRegisterDetailGroupWise.rpt")
         Case "Product Wise All Fields"
            Set RptReportViewer.Report = New CrpStockValueRegisterDetailproductWiseAllFields
         Case "SubGroup Wise"
            'Set RptReportViewer.Report = New CrpStockValueRegisterDetailSubGroupWise
             Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\StockReports\CrpStockValueRegisterDetailSubGroupWise.rpt")
         Case "Store Wise"
            Set RptReportViewer.Report = New CrpStockValueRegisterDetailStoreWise
         Case "Vendor Wise"
            Set RptReportViewer.Report = New CrpStockValueRegisterDetailVendorWise
      End Select
   Else
      Select Case CmbGroup.Text
         Case "Brand Wise"
            Set RptReportViewer.Report = New CrpStockValueRegisterBrandWise
         Case "Company Wise"
            Set RptReportViewer.Report = New CrpStockValueRegisterCompanyWise
         Case "Group Wise"
            Set RptReportViewer.Report = New CrpStockValueRegisterGroupWise
         Case "SubGroup Wise"
            Set RptReportViewer.Report = New CrpStockValueRegisterSubGroupWise
         Case "Product Wise"
            Set RptReportViewer.Report = New CrpStockValueRegisterProductWise
         Case "Store Wise"
            Set RptReportViewer.Report = New CrpStockValueRegisterStoreWise
         Case "Vendor Wise"
            Set RptReportViewer.Report = New CrpStockValueRegisterVendorWise
      End Select
   End If
   'If OptMovingAvg.Value = True Then
   '   sSql = "EXEC ProdRptStockValue '" & DtpDate.DateValue & "'," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", TxtStoreID.Text) & "," & IIf(Trim(TxtVenderID.Text) = "", "Null", "'" & TxtVenderID.Text & "'") & "," & IIf(Trim(TxtProductID.Text) <> "", "Null", "'%" & TxtProductName.Text & "%'")
   'ElseIf OptLastPrice.Value = True Then
   '   CN.Execute "exec SPProductPurchase '" & DtpDate.DateValue & "'"
   '   sSql = "EXEC ProdRptStockValueLP '" & DtpDate.DateValue & "'," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", TxtStoreID.Text) & "," & IIf(Trim(TxtVenderID.Text) = "", "Null", "'" & TxtVenderID.Text & "'") & "," & IIf(Trim(TxtProductID.Text) <> "", "Null", "'%" & TxtProductName.Text & "%'")
   'ElseIf OptWeightedAvg.Value = True Then
   
      'CN.Execute "exec SPAverageCost '" & DtpDate.DateValue & "'"
      If OptLastPrice.Value = True Then
         CN.Execute "exec SPProductPurchase '" & DtpDate.DateValue & "'"
      Else
         CN.Execute "exec SPAverageCost '" & DtpDate.DateValue & "'"
      End If
      sSql = "EXEC ProdRptStockRegister '" & DtpDate.DateValue & "'," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtBrandID.Text) = "", "Null", TxtBrandID.Text) & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", TxtStoreID.Text) & "," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text) & "," & IIf(Trim(TxtVenderID.Text) = "", "Null", "'" & TxtVenderID.Text & "'") & vbCrLf _
      & "," & IIf(Trim(TxtItemCode.Text) = "", "Null", "'" & TxtItemCode.Text & "'") & "," & IIf(Trim(TxtDepartmentID.Text) = "", "Null", TxtDepartmentID.Text) & "," & IIf(Trim(TxtSubDepartmentID.Text) = "", "Null", TxtSubDepartmentID.Text) & "," & IIf(Trim(TxtDescriptionID.Text) = "", "Null", TxtDescriptionID.Text) & "," & IIf(Trim(TxtItemDescID.Text) = "", "Null", TxtItemDescID.Text) & "," & IIf(Trim(TxtSeasonID.Text) = "", "Null", TxtSeasonID.Text) & vbCrLf _
      & "," & IIf(Trim(TxtProductID.Text) <> "", "Null", "'%" & TxtProductName.Text & "%'") & "," & IIf(Trim(TxtMinLimit.Text) = "", "-9999999", Val(TxtMinLimit.Text)) & "," & IIf(Trim(TxtMaxLimit.Text) = "", "9999999", Val(TxtMaxLimit.Text)) & vbCrLf _
      & "," & IIf(ObjRegistry.ShowWholeSaleMargin, "'WSPrice'", "'RetailPrice'") & ",Null,Null,Null,Null,Null,Null,Null," & "'" & CmbSortName.Text & " " & CmbSortType.Text & "'"
      
   
'      CN.Execute "exec SPAverageCost '" & DtpDate.DateValue & "'"
'      sSql = "EXEC ProdRptStockValueWeightedAvg '" & DtpDate.DateValue & "'," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtBrandID.Text) = "", "Null", TxtBrandID.Text) & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", TxtStoreID.Text) & "," & IIf(Trim(TxtVenderID.Text) = "", "Null", "'" & TxtVenderID.Text & "'") & "," & IIf(Trim(TxtProductID.Text) <> "", "Null", "'%" & TxtProductName.Text & "%'," & IIf(Trim(TxtMinLimit.Text) = "", "-9999999", Val(TxtMinLimit.Text)) & "," & IIf(Trim(TxtMaxLimit.Text) = "", "9999999", Val(TxtMaxLimit.Text)))
'      sSql = "EXEC ProdRptStockValueWeightedAvg '" & DtpDate.DateValue & "'," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtBrandID.Text) = "", "Null", TxtBrandID.Text) & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", TxtStoreID.Text) & "," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text) & "," & IIf(Trim(TxtVenderID.Text) = "", "Null", "'" & TxtVenderID.Text & "'") & vbCrLf _
      & "," & IIf(Trim(TxtDepartmentID.Text) = "", "Null", TxtDepartmentID.Text) & "," & IIf(Trim(TxtSubDepartmentID.Text) = "", "Null", TxtSubDepartmentID.Text) & "," & IIf(Trim(TxtItemCode.Text) = "", "Null", "'" & TxtItemCode.Text & "'") & "," & IIf(Trim(TxtDescriptionID.Text) = "", "Null", TxtDescriptionID.Text) & "," & IIf(Trim(TxtItemDescID.Text) = "", "Null", TxtItemDescID.Text) & "," & IIf(Trim(TxtSeasonID.Text) = "", "Null", TxtSeasonID.Text) & vbCrLf _
      & "," & IIf(Trim(TxtProductID.Text) <> "", "Null", "'%" & TxtProductName.Text & "%'") & "," & IIf(Trim(TxtMinLimit.Text) = "", "-9999999", Val(TxtMinLimit.Text)) & "," & IIf(Trim(TxtMaxLimit.Text) = "", "9999999", Val(TxtMaxLimit.Text)) & vbCrLf _
      & "," & IIf(ObjRegistry.ShowWholeSaleMargin, "'WSPrice'", "'RetailPrice'") & ",Null,Null,Null,Null,Null,Null,Null," & "'" & CmbSortName.Text & " " & CmbSortType.Text & "'"
      

   'End If
   
   Set RsReport = CN.Execute(sSql)
   
   If RsReport.BOF Then
      MsgBox "No record exists.", vbInformation, Me.Caption
      Me.MousePointer = vbDefault
      Exit Function
   End If
   
   If RdoDetail.Value = True Then
      RptReportViewer.Report.ReportTitle = "Stock Register Detail (" & CmbGroup.Text & ")"
   Else
      RptReportViewer.Report.ReportTitle = "Stock Register Summary (" & CmbGroup.Text & ")"
   End If
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport

'   RptReportViewer.Report.ReportTitle = "Stock Value Register (" & CmbGroup.Text & ")"
    
'   With CN.Execute("Select CompanyName,Address,City,PhoneNo,email from Company")
'      If .RecordCount > 0 Then
'         RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(IsNull(!CompanyName), "", CStr(!CompanyName))
'         RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(IsNull(!Address), "", !Address) & IIf(IsNull(!City), "", ", " & !City & ".")
'         RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(IsNull(!PhoneNo), "", CStr(!PhoneNo))
'      End If
'      .Close
'   End With
'   RptReportViewer.Report.ParameterFields(4).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
'   RptReportViewer.Report.ParameterFields(5).AddCurrentValue "Date : " & Format(DtpDate.DateValue, "dd/MM/yyyy")
    
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName & IIf(ObjRegistry.CompanyCity = "", "", " - " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue "Date :" & Format(DtpDate.DateValue, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(6).AddCurrentValue IIf(ObjRegistry.ShowWholeSaleMargin, "WSPrice", "Retail")
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   If CmbGroup.Text = "Product Wise All Fields" Then
      RptReportViewer.Report.PaperOrientation = crLandscape
   Else
      RptReportViewer.Report.PaperOrientation = crPortrait
   End If
   SetReport = True
   Me.MousePointer = vbDefault
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnOrganization_Click()
   If FunSelectOrganization(ssButton, False) = True Then
      TxtCompanyID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
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

