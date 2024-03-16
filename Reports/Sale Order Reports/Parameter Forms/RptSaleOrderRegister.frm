VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptSaleOrderRegister 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptSaleOrderRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkRemOrder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remaining Sale Order"
      Height          =   255
      Left            =   9233
      TabIndex        =   73
      Top             =   3210
      Width           =   2415
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   8333
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   2895
      Width           =   1950
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   10298
      TabIndex        =   38
      Top             =   2895
      Width           =   2250
      Begin VB.OptionButton RdoSummary 
         BackColor       =   &H80000009&
         Caption         =   "Summary"
         Height          =   255
         Left            =   975
         TabIndex        =   40
         Top             =   10
         Width           =   960
      End
      Begin VB.OptionButton RdoDetail 
         Caption         =   "Detail"
         Height          =   255
         Left            =   120
         TabIndex        =   39
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
      Left            =   8333
      TabIndex        =   33
      Top             =   2565
      Width           =   4215
      Begin VB.OptionButton RdoBoth 
         BackColor       =   &H80000009&
         Caption         =   "Both Order"
         Height          =   255
         Left            =   2790
         TabIndex        =   36
         Top             =   10
         Width           =   1335
      End
      Begin VB.OptionButton RdoReturn 
         BackColor       =   &H80000009&
         Caption         =   "Return Order"
         Height          =   255
         Left            =   1335
         TabIndex        =   35
         Top             =   10
         Width           =   1455
      End
      Begin VB.OptionButton RdoInv 
         BackColor       =   &H80000009&
         Caption         =   "Sale Order"
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   10
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11303
      TabIndex        =   45
      Top             =   8895
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
      MICON           =   "RptSaleOrderRegister.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8528
      TabIndex        =   43
      Top             =   8895
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
      MICON           =   "RptSaleOrderRegister.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9908
      TabIndex        =   44
      Top             =   8895
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
      MICON           =   "RptSaleOrderRegister.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   11183
      TabIndex        =   46
      Top             =   1905
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
      Left            =   2318
      TabIndex        =   30
      Top             =   9090
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
      IntegralPoint   =   15
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   3338
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   9090
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
      MICON           =   "RptSaleOrderRegister.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   3698
      TabIndex        =   32
      Top             =   9090
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
      Left            =   3338
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2580
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
      MICON           =   "RptSaleOrderRegister.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   2318
      TabIndex        =   0
      Top             =   2580
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
      Left            =   3698
      TabIndex        =   2
      Tag             =   "nc"
      Top             =   2580
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
      Left            =   3338
      TabIndex        =   4
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
      MICON           =   "RptSaleOrderRegister.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   3698
      TabIndex        =   5
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
      Left            =   9023
      TabIndex        =   41
      Top             =   8325
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
      Left            =   10778
      TabIndex        =   42
      Top             =   8325
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
      Left            =   3338
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7785
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
      MICON           =   "RptSaleOrderRegister.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   2318
      TabIndex        =   24
      Top             =   7785
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
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   3698
      TabIndex        =   26
      Tag             =   "nc"
      Top             =   7785
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
      Left            =   3338
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   8430
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
      MICON           =   "RptSaleOrderRegister.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   2318
      TabIndex        =   27
      Top             =   8430
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
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   3698
      TabIndex        =   29
      Tag             =   "nc"
      Top             =   8430
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
      Left            =   3338
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7125
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
      MICON           =   "RptSaleOrderRegister.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   2318
      TabIndex        =   21
      Top             =   7125
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
      Left            =   3698
      TabIndex        =   23
      Tag             =   "nc"
      Top             =   7125
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
      Left            =   3338
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   6480
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
      MICON           =   "RptSaleOrderRegister.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtUserNo 
      Height          =   315
      Left            =   2318
      TabIndex        =   18
      Top             =   6480
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
      Left            =   3698
      TabIndex        =   20
      Tag             =   "nc"
      Top             =   6480
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
      Left            =   2318
      TabIndex        =   3
      Top             =   3225
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
      Left            =   3338
      TabIndex        =   13
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
      MICON           =   "RptSaleOrderRegister.frx":0FE2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPartyID 
      Height          =   315
      Left            =   2318
      TabIndex        =   12
      Top             =   5175
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
   Begin JeweledBut.JeweledButton BtnEmpName 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3338
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5835
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
      MICON           =   "RptSaleOrderRegister.frx":0FFE
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmpID 
      Height          =   315
      Left            =   2318
      TabIndex        =   15
      Top             =   5835
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
   Begin SITextBox.Txt TxtPartyName 
      Height          =   315
      Left            =   3698
      TabIndex        =   14
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
   Begin SITextBox.Txt TxtEmpName 
      Height          =   315
      Left            =   3698
      TabIndex        =   17
      Tag             =   "nc"
      Top             =   5835
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
      Left            =   3338
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4530
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
      MICON           =   "RptSaleOrderRegister.frx":101A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSectorID 
      Height          =   315
      Left            =   2318
      TabIndex        =   9
      Top             =   4530
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
   Begin SITextBox.Txt TxtSectorName 
      Height          =   315
      Left            =   3653
      TabIndex        =   11
      Tag             =   "nc"
      Top             =   4530
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
      Left            =   3338
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3870
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
      MICON           =   "RptSaleOrderRegister.frx":1036
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtZoneID 
      Height          =   315
      Left            =   2318
      TabIndex        =   6
      Top             =   3870
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
   Begin SITextBox.Txt TxtZoneName 
      Height          =   315
      Left            =   3698
      TabIndex        =   8
      Tag             =   "nc"
      Top             =   3870
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
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   72
      Top             =   3675
      Width           =   585
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   71
      Top             =   3675
      Width           =   840
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   70
      Top             =   4320
      Width           =   930
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   69
      Top             =   4320
      Width           =   675
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   68
      Top             =   5625
      Width           =   780
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   67
      Top             =   5625
      Width           =   525
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   66
      Top             =   4965
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   65
      Top             =   4965
      Width           =   1125
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   64
      Top             =   3045
      Width           =   585
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   63
      Top             =   6285
      Width           =   795
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   62
      Top             =   6285
      Width           =   540
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   61
      Top             =   7575
      Width           =   645
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   60
      Top             =   7575
      Width           =   900
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   59
      Top             =   8205
      Width           =   975
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   58
      Top             =   6915
      Width           =   870
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   57
      Top             =   6915
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   56
      Top             =   8205
      Width           =   1230
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   55
      Top             =   8895
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   54
      Top             =   8895
      Width           =   1020
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   2318
      TabIndex        =   53
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   52
      Top             =   2385
      Width           =   1350
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   3698
      TabIndex        =   51
      Top             =   3045
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   195
      Left            =   10793
      TabIndex        =   50
      Top             =   8100
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      Height          =   195
      Left            =   9023
      TabIndex        =   49
      Top             =   8100
      Width           =   735
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Order Register"
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
      TabIndex        =   48
      Top             =   270
      Width           =   2220
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   11183
      TabIndex        =   47
      Top             =   1710
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
Attribute VB_Name = "RptSaleOrderRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql As String

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCustomer.Show vbModal, Me
        If SchCustomer.ParaOutCustomerID = "" Then FunSelectCustomer = False: Exit Function
        TxtPartyID.Text = SchCustomer.ParaOutCustomerID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Parties WHERE PartyType='C' And PartyID=" & Val(TxtPartyID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtPartyName.Text = !PartyName
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

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
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
    vStrSQL = " SELECT p.productid, code, ProductName" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
  
   With CN.Execute(vStrSQL)
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

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
     RdoInv.SetFocus
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

Private Sub BtnStore_Click()
If FunSelectStore(ssButton, False) = True Then
     TxtZoneID.SetFocus
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



Private Sub Txt1_Change()

End Sub

Private Sub BtnZone_Click()
 If FunSelectZone(ssButton, False) = True Then
      TxtSectorID.SetFocus
   Else
      TxtZoneID.SetFocus
   End If
End Sub

Private Sub TxtSectorID_Change()
   If TxtSectorID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   If TxtPartyName.Text <> "" Then TxtPartyName.Text = ""
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
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCompany.Show vbModal, Me
        If SchCompany.ParaOutCompanyID = "" Then FunSelectCompany = False: Exit Function
        TxtCompanyID.Text = SchCompany.ParaOutCompanyID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Companies where CompanyID=" & Val(TxtCompanyID.Text)
    With CN.Execute(vStrSQL)
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
    Dim vStrSQL As String
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
    vStrSQL = " Select * FROM Groups where GroupID = '" & TxtGroupID.Text & "'"
    With CN.Execute(vStrSQL)
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
   vTemp = Not FunSelectOrganizaton(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganizaton(ssButton, False)
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
   vTemp = Not FunSelectUser(ssValidate, True)
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
            RptReportViewer.Caption = "Sale Order Detail (" & CmbGroup.Text & ")"
        ElseIf RdoReturn.Value = True Then
            RptReportViewer.Caption = "Sale Return Order Detail (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Caption = "Sale Order & Sale Return Order Detail (" & CmbGroup.Text & ")"
        End If
    Else
        If RdoInv.Value = True Then
            RptReportViewer.Caption = "Sale Order Summary (" & CmbGroup.Text & ")"
        ElseIf RdoReturn.Value = True Then
            RptReportViewer.Caption = "Sale Order Return Summary (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Caption = "Sale Order & Sale Return Order Summary (" & CmbGroup.Text & ")"
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
        Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then TxtZoneID.SetFocus
        Case TxtZoneID.Name: If FunSelectZone(ssFunctionKey, True) = True Then TxtSectorID.SetFocus
        Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, True) = True Then TxtPartyID.SetFocus
        Case TxtPartyID.Name: If FunSelectCustomer(ssFunctionKey, True) = True Then TxtEmpID.SetFocus
        Case TxtEmpID.Name: If FunSelectEmployee(ssFunctionKey, True) = True Then TxtUserNo.SetFocus
        Case TxtUserNo.Name: If FunSelectUser(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus
        Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtGroupID.SetFocus
        Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtSubGroupID.SetFocus
        Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then TxtCode.SetFocus
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
   SetWindowText Me.hWnd, "Sale Order Register"
   
   CmbGroup.AddItem ("Organization Wise")
   CmbGroup.AddItem ("Store Wise")
   CmbGroup.AddItem ("Zone Wise")
   CmbGroup.AddItem ("Sector Wise")
   CmbGroup.AddItem ("Customer Wise")
   CmbGroup.AddItem ("Employee Wise")
   CmbGroup.AddItem ("User Wise")
   CmbGroup.AddItem ("Company Wise")
   CmbGroup.AddItem ("Group Wise")
   CmbGroup.AddItem ("SubGroup Wise")
   CmbGroup.AddItem ("Product Wise")
   CmbGroup.AddItem ("Date Wise")
   CmbGroup.AddItem ("Invoice Wise")
   
   'CmbGroup.AddItem ("Sale Detail (All Wise)")
   CmbGroup.ListIndex = 0
   ChkRemOrder.Value = 1
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
   Set RptSaleOrderRegister = Nothing
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
    sSql = "EXEC ProdRptSaleOrderRegister '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", "'" & TxtStoreID.Text & "'") & "," & IIf(Trim(TxtZoneID.Text) = "", "Null", "'" & TxtZoneID.Text & "'") & "," & IIf(Trim(TxtSectorID.Text) = "", "Null", "'" & TxtSectorID.Text & "'") & "," & IIf(Trim(TxtPartyID.Text) = "", "Null", "'" & TxtPartyID.Text & "'") & "," & vbCrLf _
                              & IIf(Trim(TxtEmpID.Text) = "", "Null", "'" & TxtEmpID.Text & "'") & "," & IIf(Trim(TxtUserNo.Text) = "", "Null", "'" & TxtUserNo.Text & "'") & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & "," & IIf(ChkRemOrder.Value = 1, 0, "Null") & "," & IIf(RdoBoth.Value = True, "Null", IIf(RdoInv.Value = True, 0, 1))
    If RdoDetail.Value = True Then
        Select Case CmbGroup.Text
            Case "Organization Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailOrgWise
            Case "Store Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailStoreWise
            Case "Zone Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailZoneWise
            Case "Sector Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailSectorWise
            Case "Customer Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailCustomerWise
            Case "Employee Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailEmpWise
            Case "User Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailUserWise
            Case "Company Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailCompanyWise
            Case "Group Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailGroupWise
            Case "SubGroup Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailSubGroupWise
            Case "Product Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailProductWise
            Case "Date Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailDateWise
            Case "Invoice Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailInvoiceWise
            Case "Sale (All Wise)"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleDetailAllWise
        End Select
    
    Else
        Select Case CmbGroup.Text
            Case "Organization Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryOrgWise
            Case "Store Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryStoreWise
            Case "Zone Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryZoneWise
            Case "Sector Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummarySectorWise
            Case "Customer Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryCustomerWise
            Case "Employee Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryEmpWise
            Case "User Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryUserWise
            Case "Company Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryCompanyWise
            Case "Group Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryGroupWise
            Case "SubGroup Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummarySubGroupWise
            Case "Product Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryProductWise
            Case "Date Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryDateWise
            Case "Invoice Wise"
                Set RsReport = CN.Execute(sSql)
                Set RptReportViewer.Report = New CrptSaleSummaryInvoiceWise
        End Select
    
    End If
     
    
    
    If RsReport.BOF Then
        MsgBox "No record exists.", vbInformation, Me.Caption
        Me.MousePointer = vbDefault
        Exit Function
    End If
    
    If RdoDetail.Value = True Then
        If RdoInv.Value = True Then
            RptReportViewer.Report.ReportTitle = "Sale Order Detail (" & CmbGroup.Text & ")"
        ElseIf RdoReturn.Value = True Then
            RptReportViewer.Report.ReportTitle = "Sale Return Order Detail (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Report.ReportTitle = "Sale Order & Sale Return Order Detail (" & CmbGroup.Text & ")"
        End If
    Else
        If RdoInv.Value = True Then
            RptReportViewer.Report.ReportTitle = "Sale Order Summary (" & CmbGroup.Text & ")"
        ElseIf RdoReturn.Value = True Then
            RptReportViewer.Report.ReportTitle = "Sale Return Order Summary (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Report.ReportTitle = "Sale Order & Sale Return Order Summary (" & CmbGroup.Text & ")"
        End If
        
    End If
    
    RptReportViewer.Report.Database.SetDataSource RsReport
    RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
    RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
    RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
    RptReportViewer.Report.ParameterFields(4).AddCurrentValue " Date From :" & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To : " & Format(DtpTo.DateValue, "dd/MM/yyyy")
    RptReportViewer.Report.ParameterFields(5).AddCurrentValue ObjRegistry.DevelopedBy
    RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
    RptReportViewer.Report.PaperOrientation = crPortrait

   
    
    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    'RptReportViewer.Report.PaperOrientation = crLandscape
    SetReport = True
    Me.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function
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
    If Len(TxtStoreID.Text) <= 3 Then
      TxtStoreID.Text = Right("000" + CStr(Val(TxtStoreID.Text)), 3)
    End If
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
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSubGroup.Show vbModal, Me
        If SchSubGroup.ParaOutSubGroupID = "" Then FunSelectSubGroup = False: Exit Function
        TxtSubGroupID.Text = SchSubGroup.ParaOutSubGroupID
    End If
    '---------------------------
    vStrSQL = " Select * FROM SubGroups where SubGroupID = " & Val(TxtSubGroupID.Text)
    With CN.Execute(vStrSQL)
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
      TxtCode.SetFocus
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
   If TxtPartyName.Text <> "" Then TxtPartyName.Text = ""
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
          TxtSectorName.Text = !ZoneName
          FunSelectZone = True
          .Close
          Exit Function
             FunSelectZone = True
   Else
          FunSelectZone = False
          .Close
          TxtZoneID.Text = ""
          TxtSectorName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

