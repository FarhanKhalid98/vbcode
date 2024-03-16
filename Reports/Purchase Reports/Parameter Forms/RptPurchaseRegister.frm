VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptPurchaseRegister 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptPurchaseRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbSortName 
      Height          =   315
      ItemData        =   "RptPurchaseRegister.frx":0ECA
      Left            =   8355
      List            =   "RptPurchaseRegister.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   5505
      Width           =   1815
   End
   Begin VB.ComboBox CmbSortType 
      Height          =   315
      ItemData        =   "RptPurchaseRegister.frx":0ECE
      Left            =   10245
      List            =   "RptPurchaseRegister.frx":0ED0
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   5505
      Width           =   1275
   End
   Begin VB.CheckBox ChkZeroPriceInvoices 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zero Price Invoices"
      Height          =   255
      Left            =   11460
      TabIndex        =   16
      Top             =   4770
      Width           =   1740
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   7260
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4710
      Width           =   1950
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9210
      TabIndex        =   39
      Top             =   4710
      Width           =   2250
      Begin VB.OptionButton RdoSummary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary"
         Height          =   255
         Left            =   975
         TabIndex        =   15
         Top             =   10
         Width           =   960
      End
      Begin VB.OptionButton RdoDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detail"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   10
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   7275
      TabIndex        =   38
      Top             =   4335
      Width           =   5940
      Begin VB.OptionButton RdoNet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Net Purchase"
         Height          =   300
         Left            =   4275
         TabIndex        =   12
         Top             =   45
         Width           =   1365
      End
      Begin VB.OptionButton RdoBoth 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Both"
         Height          =   255
         Left            =   3420
         TabIndex        =   11
         Top             =   90
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton RdoReturn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Return Invioce"
         Height          =   255
         Left            =   1845
         TabIndex        =   10
         Top             =   90
         Width           =   1410
      End
      Begin VB.OptionButton RdoInv 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Purchase Invoice"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   90
         Width           =   1560
      End
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10560
      TabIndex        =   21
      Top             =   7905
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
      MICON           =   "RptPurchaseRegister.frx":0ED2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7785
      TabIndex        =   19
      Top             =   7905
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
      MICON           =   "RptPurchaseRegister.frx":0EEE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9165
      TabIndex        =   20
      Top             =   7905
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
      MICON           =   "RptPurchaseRegister.frx":0F0A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   8978
      TabIndex        =   40
      Top             =   2438
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
      Left            =   1575
      TabIndex        =   8
      Top             =   5640
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
      Left            =   2595
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5640
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
      MICON           =   "RptPurchaseRegister.frx":0F26
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2955
      TabIndex        =   37
      Top             =   5640
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
      Masked          =   5
   End
   Begin JeweledBut.JeweledButton BtnOrganizaton 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2595
      TabIndex        =   22
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
      MICON           =   "RptPurchaseRegister.frx":0F42
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   1575
      TabIndex        =   0
      Top             =   1350
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
   Begin SITextBox.Txt TxtOrganizatonName 
      Height          =   315
      Left            =   2955
      TabIndex        =   23
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2595
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1875
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
      MICON           =   "RptPurchaseRegister.frx":0F5E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   2955
      TabIndex        =   25
      Tag             =   "nc"
      Top             =   1875
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
      Left            =   8280
      TabIndex        =   17
      Top             =   7305
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
      Left            =   10035
      TabIndex        =   18
      Top             =   7305
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
      Left            =   2595
      TabIndex        =   32
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
      MICON           =   "RptPurchaseRegister.frx":0F7A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   1575
      TabIndex        =   5
      Top             =   3990
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
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   2955
      TabIndex        =   33
      Tag             =   "nc"
      Top             =   3990
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
      Left            =   2595
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4515
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
      MICON           =   "RptPurchaseRegister.frx":0F96
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   1575
      TabIndex        =   6
      Top             =   4515
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
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   2955
      TabIndex        =   35
      Tag             =   "nc"
      Top             =   4500
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
      Left            =   2595
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3450
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
      MICON           =   "RptPurchaseRegister.frx":0FB2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   1575
      TabIndex        =   4
      Top             =   3450
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
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   2955
      TabIndex        =   31
      Tag             =   "nc"
      Top             =   3450
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
      Left            =   2595
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2940
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
      MICON           =   "RptPurchaseRegister.frx":0FCE
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtUserNo 
      Height          =   315
      Left            =   1575
      TabIndex        =   3
      Top             =   2940
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
   Begin SITextBox.Txt TxtUserName 
      Height          =   315
      Left            =   2955
      TabIndex        =   29
      Tag             =   "nc"
      Top             =   2940
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
      Left            =   1575
      TabIndex        =   1
      Top             =   1875
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
   Begin JeweledBut.JeweledButton BtnVendor 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2595
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2430
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
      MICON           =   "RptPurchaseRegister.frx":0FEA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtVendorID 
      Height          =   315
      Left            =   1575
      TabIndex        =   2
      Top             =   2430
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
   Begin SITextBox.Txt TxtPartyName 
      Height          =   315
      Left            =   2955
      TabIndex        =   27
      Tag             =   "nc"
      Top             =   2430
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
   Begin JeweledBut.JeweledButton BtnBrand 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2580
      TabIndex        =   61
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5055
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
      MICON           =   "RptPurchaseRegister.frx":1006
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBrandID 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   5055
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
      Left            =   2940
      TabIndex        =   62
      Tag             =   "nc"
      Top             =   5055
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
      Left            =   9030
      TabIndex        =   69
      Top             =   3660
      Width           =   915
      _ExtentX        =   1614
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
      Left            =   10110
      TabIndex        =   70
      Top             =   3660
      Width           =   870
      _ExtentX        =   1535
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
   Begin SITextBox.Txt TxtDepartmentID 
      Height          =   315
      Left            =   1575
      TabIndex        =   73
      Top             =   8610
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
      Left            =   2955
      TabIndex        =   74
      Tag             =   "nc"
      Top             =   8610
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
      Left            =   1575
      TabIndex        =   75
      Top             =   9345
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
      Left            =   2955
      TabIndex        =   76
      Tag             =   "nc"
      Top             =   9345
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
      Left            =   1575
      TabIndex        =   77
      Top             =   10110
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
      Left            =   2955
      TabIndex        =   78
      Tag             =   "nc"
      Top             =   10110
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
      Left            =   1575
      TabIndex        =   79
      Top             =   6375
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
      Left            =   2955
      TabIndex        =   80
      Top             =   6375
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
      Left            =   1575
      TabIndex        =   81
      Top             =   7815
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
      Left            =   2955
      TabIndex        =   82
      Tag             =   "nc"
      Top             =   7815
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
      Left            =   1575
      TabIndex        =   83
      Top             =   7140
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
      Left            =   2955
      TabIndex        =   84
      Tag             =   "nc"
      Top             =   7140
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
      Left            =   2595
      TabIndex        =   85
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   8610
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
      MICON           =   "RptPurchaseRegister.frx":1022
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSubDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2595
      TabIndex        =   86
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   9345
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
      MICON           =   "RptPurchaseRegister.frx":103E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton TxtSeason 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2595
      TabIndex        =   87
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   10110
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
      MICON           =   "RptPurchaseRegister.frx":105A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnItemCode 
      Height          =   330
      Left            =   2595
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   6375
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
      MICON           =   "RptPurchaseRegister.frx":1076
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnItemDesc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2595
      TabIndex        =   89
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7815
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
      MICON           =   "RptPurchaseRegister.frx":1092
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDescription 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2595
      TabIndex        =   90
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7155
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
      MICON           =   "RptPurchaseRegister.frx":10AE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSession 
      Height          =   330
      Left            =   8340
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   1335
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
      MICON           =   "RptPurchaseRegister.frx":10CA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSessionID 
      Height          =   315
      Left            =   7335
      TabIndex        =   104
      Top             =   1335
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
   Begin SITextBox.Txt TxtSessionName 
      Height          =   315
      Left            =   8700
      TabIndex        =   105
      Top             =   1335
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
      Left            =   8700
      TabIndex        =   107
      Top             =   1125
      Width           =   1215
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
      Left            =   7335
      TabIndex        =   106
      Top             =   1125
      Width           =   930
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
      Left            =   1575
      TabIndex        =   102
      Top             =   6930
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
      Left            =   2955
      TabIndex        =   101
      Top             =   6930
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
      Left            =   1575
      TabIndex        =   100
      Top             =   7605
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
      Left            =   2955
      TabIndex        =   99
      Top             =   7605
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
      Left            =   2955
      TabIndex        =   98
      Top             =   6165
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
      Left            =   1575
      TabIndex        =   97
      Top             =   6165
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
      Left            =   1575
      TabIndex        =   96
      Top             =   9900
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
      Left            =   2955
      TabIndex        =   95
      Top             =   9900
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
      Left            =   1575
      TabIndex        =   94
      Top             =   9135
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
      Left            =   2955
      TabIndex        =   93
      Top             =   9135
      Width           =   1410
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
      Left            =   2955
      TabIndex        =   92
      Top             =   8400
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
      Left            =   1575
      TabIndex        =   91
      Top             =   8400
      Width           =   1245
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From ID"
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
      Left            =   9030
      TabIndex        =   72
      Top             =   3465
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To ID"
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
      Left            =   10110
      TabIndex        =   71
      Top             =   3465
      Width           =   495
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
      Left            =   8355
      TabIndex        =   68
      Top             =   5265
      Width           =   900
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
      Left            =   10245
      TabIndex        =   67
      Top             =   5265
      Width           =   840
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
      Left            =   2940
      TabIndex        =   64
      Top             =   4860
      Width           =   1050
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
      Left            =   1560
      TabIndex        =   63
      Top             =   4860
      Width           =   765
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor ID"
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
      Left            =   1575
      TabIndex        =   60
      Top             =   2235
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
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
      Left            =   2955
      TabIndex        =   59
      Top             =   2235
      Width           =   1155
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
      Left            =   1575
      TabIndex        =   58
      Top             =   1680
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
      Left            =   2955
      TabIndex        =   57
      Top             =   2745
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
      Left            =   1575
      TabIndex        =   56
      Top             =   2745
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
      Left            =   1575
      TabIndex        =   55
      Top             =   3795
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
      Left            =   2955
      TabIndex        =   54
      Top             =   3795
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
      Left            =   1575
      TabIndex        =   53
      Top             =   4320
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
      Left            =   1575
      TabIndex        =   52
      Top             =   3255
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
      Left            =   2955
      TabIndex        =   51
      Top             =   3255
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
      Left            =   2955
      TabIndex        =   50
      Top             =   4305
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   1575
      TabIndex        =   49
      Top             =   5445
      Width           =   450
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
      Left            =   2955
      TabIndex        =   48
      Top             =   5445
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
      Left            =   1575
      TabIndex        =   47
      Top             =   1155
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
      Left            =   2955
      TabIndex        =   46
      Top             =   1155
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
      Left            =   2955
      TabIndex        =   45
      Top             =   1680
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
      Left            =   10050
      TabIndex        =   44
      Top             =   7080
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
      Left            =   8280
      TabIndex        =   43
      Top             =   7080
      Width           =   885
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Register"
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
      TabIndex        =   42
      Top             =   270
      Width           =   2115
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
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
      Left            =   8978
      TabIndex        =   41
      Top             =   2243
      Visible         =   0   'False
      Width           =   870
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
Attribute VB_Name = "RptPurchaseRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql As String

Private Function FunSelectVendor(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.CmbFilter = "Vendors"
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectVendor = False: Exit Function
        TxtVendorID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    VStrSQL = " Select c.AccountNo, c.AccountName as AccountName, Address, City" & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " where c.AccountNo = '" & (TxtVendorID.Text) & "' and (c.AccountNo like '6%') and isDetailed = 1 and isLocked = 0"
    
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtPartyName.Text = !AccountName
          FunSelectVendor = True
          .Close
          Exit Function
      Else
          FunSelectVendor = False
          .Close
          TxtVendorID.Text = ""
          TxtPartyName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectOrganizaton(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganizaton = False: Exit Function
       TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    If TxtOrganizationID.Text = "" Then FunSelectOrganizaton = False: Exit Function
    VStrSQL = " Select * FROM Organizations where OrganizationID='" & TxtOrganizationID.Text & "'"
    With CN.Execute(VStrSQL)
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
    If Len(TxtCode.Text) <= 5 Then
      TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
    End If
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    VStrSQL = " SELECT p.productid, code, ProductName" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
  
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

Private Sub BtnItemCode_Click()
   If FunSelectItemCode(ssButton, True) = True Then
      TxtItemCode.SetFocus
   Else
      TxtDescriptionID.SetFocus
   End If
End Sub

Private Sub BtnSession_Click()
   If FunSelectSession(ssButton, False) = True Then
      RdoDetail.SetFocus
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
          TxtSessionName.Text = "All Sessions"
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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
     TxtSessionID.SetFocus
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
     TxtItemCode.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnOrganizaton_Click()
   If FunSelectOrganizaton(ssButton, False) = True Then
      TxtStoreID.SetFocus
   Else
     TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnVendor_Click()
   If FunSelectVendor(ssButton, False) = True Then
      TxtCode.SetFocus
   Else
      TxtVendorID.SetFocus
   End If
End Sub

Private Sub BtnStore_Click()
If FunSelectStore(ssButton, False) = True Then
     TxtVendorID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub BtnUser_Click()
   If FunSelectUser(ssButton, False) = True Then
      TxtCompanyID.SetFocus
   Else
      TxtUserNo.SetFocus
   End If
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
'      TxtCode.Text = ""
'      TxtProductID.Text = ""
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

Private Sub TxtVendorID_Change()
   If TxtVendorID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtVendorID.Name Then Exit Sub
   If TxtPartyName.Text <> "" Then TxtPartyName.Text = ""
End Sub

Private Sub TxtVendorID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtVendorID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtVendorID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVendor(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVendor(ssButton, False)
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
   vTemp = Not FunSelectOrganizaton(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganizaton(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
    If RdoDetail.Value = True Then
        If RdoInv.Value = True Then
            RptReportViewer.Caption = "Purchase Detail (" & CmbGroup.Text & ")"
        ElseIf RdoReturn.Value = True Then
            RptReportViewer.Caption = "Purchase Return Detail (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Caption = "Purchase & Purchase Return Detail (" & CmbGroup.Text & ")"
        End If
    Else
        If RdoInv.Value = True Then
            RptReportViewer.Caption = "Purchase Summary (" & CmbGroup.Text & ")"
        ElseIf RdoReturn.Value = True Then
            RptReportViewer.Caption = "Purchase Return Summary (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Caption = "Purchase & Purchase Return Summary (" & CmbGroup.Text & ")"
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
        Case TxtOrganizationID.Name: If FunSelectOrganizaton(ssFunctionKey, True) = True Then TxtStoreID.SetFocus
        Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then TxtVendorID.SetFocus
        Case TxtVendorID.Name: If FunSelectVendor(ssFunctionKey, True) = True Then TxtUserNo.SetFocus
        Case TxtUserNo.Name: If FunSelectUser(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus
        Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtGroupID.SetFocus
        Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtSubGroupID.SetFocus
        Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then TxtBrandID.SetFocus
        Case TxtBrandID.Name: If FunSelectBrand(ssFunctionKey, True) = True Then TxtCode.SetFocus
        Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then RdoInv.SetFocus
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
   SetWindowText Me.hWnd, "Purchase Register"
   
   CmbGroup.AddItem ("Brand Wise")
   CmbGroup.AddItem ("Company Wise")
'   CmbGroup.AddItem ("Customer Wise")
   CmbGroup.AddItem ("Date Wise")
   CmbGroup.AddItem ("Group Wise")
   CmbGroup.AddItem ("Invoice Wise")
   CmbGroup.AddItem ("Organization Wise")
   CmbGroup.AddItem ("Product Wise")
   CmbGroup.AddItem ("Product Wise All Fields")
   CmbGroup.AddItem ("Store Wise")
   CmbGroup.AddItem ("SubGroup Wise")
   CmbGroup.AddItem ("User Wise")
   CmbGroup.AddItem ("Vendor Wise")
   CmbGroup.AddItem ("Vendor Wise Profit Margin")
   
   CmbSortName.Clear
   CmbSortName.AddItem "ProductName"
   CmbSortName.AddItem "ProductID"
   CmbSortType.Clear
   CmbSortType.AddItem "Ascending"
   CmbSortType.AddItem "Descending"
   
   'CmbGroup.AddItem ("Purchase Detail (All Wise)")
   CmbGroup.ListIndex = 0
   TxtSessionID.Text = vSessionID
   FunSelectSession ssValidate, True
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
'   Set RsReport = Nothing
   Set RptPurchaseRegister = Nothing
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
    
'    sSql = "EXEC ProdRptPurchaseRegister '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", "'" & TxtStoreID.Text & "'") & "," & IIf(Trim(TxtVendorID.Text) = "", "Null", "'" & TxtVendorID.Text & "'") & "," & IIf(Trim(TxtUserNo.Text) = "", "Null", "'" & TxtUserNo.Text & "'") & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtBrandID.Text) = "", "Null", "'" & TxtBrandID.Text & "'") & "," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & vbCrLf _
'    & "," & IIf(Trim(TxtDepartmentID.Text) = "", "Null", TxtDepartmentID.Text) & "," & IIf(Trim(TxtSubDepartmentID.Text) = "", "Null", TxtSubDepartmentID.Text) & "," & IIf(Trim(TxtItemCode.Text) = "", "Null", "'" & TxtItemCode.Text & "'") & "," & IIf(Trim(TxtDescriptionID.Text) = "", "Null", TxtDescriptionID.Text) & "," & IIf(Trim(TxtItemDescID.Text) = "", "Null", TxtItemDescID.Text) & "," & IIf(Trim(TxtSeasonID.Text) = "", "Null", TxtSeasonID.Text) & vbCrLf _
'    & "," & IIf(RdoBoth.Value = True Or RdoNet.Value = True, "Null", IIf(RdoInv.Value = True, 0, 1)) & "," & IIf(ChkZeroPriceInvoices.Value = 1, 0, "Null") & "," & "'" & CmbSortName.Text & " " & CmbSortType.Text & "'," & Val(TxtFrom.Text) & "," & Val(TxtTo.Text)
'    Set RsReport = CN.Execute(sSql)
    
    Set RsReport = CN.Execute("EXEC ProdRptPurchaseRegister '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", "'" & TxtStoreID.Text & "'") & "," & IIf(Trim(TxtVendorID.Text) = "", "Null", "'" & TxtVendorID.Text & "'") & "," & IIf(Trim(TxtUserNo.Text) = "", "Null", "'" & TxtUserNo.Text & "'") & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtBrandID.Text) = "", "Null", "'" & TxtBrandID.Text & "'") & "," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & vbCrLf _
    & "," & IIf(Trim(TxtDepartmentID.Text) = "", "Null", TxtDepartmentID.Text) & "," & IIf(Trim(TxtItemCode.Text) = "", "Null", "'" & TxtItemCode.Text & "'") & "," & IIf(Trim(TxtSubDepartmentID.Text) = "", "Null", TxtSubDepartmentID.Text) & "," & IIf(Trim(TxtDescriptionID.Text) = "", "Null", TxtDescriptionID.Text) & "," & IIf(Trim(TxtItemDescID.Text) = "", "Null", TxtItemDescID.Text) & "," & IIf(Trim(TxtSeasonID.Text) = "", "Null", TxtSeasonID.Text) & vbCrLf _
    & "," & IIf(RdoBoth.Value = True Or RdoNet.Value = True, "Null", IIf(RdoInv.Value = True, 0, 1)) & "," & IIf(ChkZeroPriceInvoices.Value = 1, 0, "Null") & "," & IIf(Trim(TxtSessionID.Text) = "", "Null", TxtSessionID.Text) & "," & "'" & CmbSortName.Text & " " & CmbSortType.Text & "'", Val(TxtFrom.Text), Val(TxtTo.Text))
    
    
    If RdoDetail.Value = True Then
        Select Case CmbGroup.Text
            Case "Organization Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailOrgWise
            Case "Brand Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailBrandWise
            Case "Store Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailStoreWise
            Case "Vendor Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailVendorWise
            Case "Vendor Wise Profit Margin"
                Set RptReportViewer.Report = New CrptPurchaseDetailVendorWiseProfitMargin
            Case "Vendor Wise Without Price"
                Set RptReportViewer.Report = New CrptPurchaseDetailVendorWiseWithoutPrice
            Case "User Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailUserWise
            Case "Company Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailCompanyWise
            Case "Group Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailGroupWise
            Case "SubGroup Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailSubGroupWise
            Case "Product Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailProductWise
            Case "Product Wise All Fields"
                Set RptReportViewer.Report = New CrptPurchaseDetailProductWiseAllFields
            Case "Date Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailDateWise
            Case "Invoice Wise"
                Set RptReportViewer.Report = New CrptPurchaseDetailInvoiceWise
            Case "Purchase (All Wise)"
                Set RptReportViewer.Report = New CrptPurchaseDetailAllWise
        End Select
    Else
        Select Case CmbGroup.Text
            Case "Organization Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummaryOrgWise
            Case "Brand Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummaryBrandWise
            Case "Store Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummaryStoreWise
            Case "Vendor Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummaryVendorWise
            Case "User Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummaryUserWise
            Case "Company Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummaryCompanyWise
            Case "Group Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummaryGroupWise
            Case "SubGroup Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummarySubGroupWise
            Case "Product Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummaryProductWise
            Case "Date Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummaryDateWise
            Case "Invoice Wise"
                Set RptReportViewer.Report = New CrptPurchaseSummaryInvoiceWise
        End Select
    End If
    
    If RsReport.BOF Then
        MsgBox "No record exists.", vbInformation, Me.Caption
        Me.MousePointer = vbDefault
        Exit Function
    End If
    
    If RdoDetail.Value = True Then
        If RdoInv.Value = True Then
            RptReportViewer.Report.ReportTitle = "Purchase Detail (" & CmbGroup.Text & ")"
        ElseIf RdoReturn.Value = True Then
            RptReportViewer.Report.ReportTitle = "Purchase Return Detail (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Report.ReportTitle = "Purchase & Purchase Return Detail (" & CmbGroup.Text & ")"
        End If
    Else
        If RdoInv.Value = True Then
            RptReportViewer.Report.ReportTitle = "Purchase Summary (" & CmbGroup.Text & ")"
        ElseIf RdoReturn.Value = True Then
            RptReportViewer.Report.ReportTitle = "Purchase Return Summary (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Report.ReportTitle = "Purchase & Purchase Return Summary (" & CmbGroup.Text & ")"
        End If
        
    End If
    
    RptReportViewer.Report.Database.SetDataSource RsReport
    If RdoNet.Value Then
      RptReportViewer.Report.DeleteGroup 1
    End If
    RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName & IIf(ObjRegistry.CompanyCity = "", "", " - " & ObjRegistry.CompanyCity)
    RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress)
    RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", "Phone # " & ObjRegistry.CompanyPhoneNo)
    RptReportViewer.Report.ParameterFields(5).AddCurrentValue "Date From " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To " & Format(DtpTo.DateValue, "dd/MM/yyyy")
    RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
    RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
    RptReportViewer.Report.PaperOrientation = crLandscape
    
    SetReport = True
    Me.MousePointer = vbDefault
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

