VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmCustomerDemand 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   15360
   ClipControls    =   0   'False
   Icon            =   "FrmCustomerDemand.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtDemandID 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      CausesValidation=   0   'False
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   4617
      TabIndex        =   0
      Top             =   2903
      Width           =   1020
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7695
      TabIndex        =   9
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6390
      TabIndex        =   10
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10305
      TabIndex        =   14
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5085
      TabIndex        =   11
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9000
      TabIndex        =   13
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   3780
      TabIndex        =   12
      Top             =   8036
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCustomerDemand.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpDemandDate 
      Height          =   315
      Left            =   5835
      TabIndex        =   1
      Top             =   2903
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
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
   Begin SITextBox.Txt TxtRemarks 
      Height          =   315
      Left            =   5063
      TabIndex        =   8
      Top             =   6934
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   100
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   7410
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   2903
      Width           =   675
      _ExtentX        =   1191
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
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   8445
      TabIndex        =   16
      Tag             =   "NC"
      Top             =   2903
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   50
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8085
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2903
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
      MICON           =   "FrmCustomerDemand.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   5033
      TabIndex        =   3
      Top             =   3664
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   100
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
   Begin SITextBox.Txt TxtVendorID 
      Height          =   315
      Left            =   5063
      TabIndex        =   7
      Top             =   6289
      Width           =   780
      _ExtentX        =   1376
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
   Begin SITextBox.Txt TxtVendorName 
      Height          =   315
      Left            =   6203
      TabIndex        =   24
      Top             =   6289
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   50
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
   Begin JeweledBut.JeweledButton BtnVendor 
      Height          =   330
      Left            =   5843
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6274
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
      MICON           =   "FrmCustomerDemand.frx":0F8E
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5843
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4369
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
      MICON           =   "FrmCustomerDemand.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   5063
      TabIndex        =   4
      Top             =   4369
      Width           =   780
      _ExtentX        =   1376
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
      Left            =   6203
      TabIndex        =   27
      Tag             =   "nc"
      Top             =   4369
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
      Left            =   5843
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5644
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
      MICON           =   "FrmCustomerDemand.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   5063
      TabIndex        =   6
      Top             =   5644
      Width           =   780
      _ExtentX        =   1376
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
      Left            =   6203
      TabIndex        =   29
      Tag             =   "nc"
      Top             =   5644
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
      Left            =   5843
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4999
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
      MICON           =   "FrmCustomerDemand.frx":0FE2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   5063
      TabIndex        =   5
      Top             =   4999
      Width           =   780
      _ExtentX        =   1376
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
      Left            =   6203
      TabIndex        =   31
      Tag             =   "nc"
      Top             =   4999
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
      Caption         =   "Sub Group Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6278
      TabIndex        =   39
      Top             =   5419
      Width           =   1455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6203
      TabIndex        =   38
      Top             =   4774
      Width           =   1320
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5063
      TabIndex        =   37
      Top             =   4774
      Width           =   1035
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5063
      TabIndex        =   36
      Top             =   5419
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6203
      TabIndex        =   35
      Top             =   4144
      Width           =   1065
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5063
      TabIndex        =   34
      Top             =   4144
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5063
      TabIndex        =   33
      Top             =   6064
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6203
      TabIndex        =   32
      Top             =   6064
      Width           =   1155
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
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
      Left            =   7410
      TabIndex        =   23
      Top             =   2711
      Width           =   720
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
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
      Left            =   8460
      TabIndex        =   22
      Top             =   2711
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Demand Date"
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
      Left            =   5835
      TabIndex        =   21
      Top             =   2689
      Width           =   1170
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
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
      Left            =   5063
      TabIndex        =   20
      Top             =   6694
      Width           =   750
   End
   Begin VB.Label Label2 
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
      Height          =   225
      Left            =   5033
      TabIndex        =   19
      Top             =   3454
      Width           =   1335
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Demand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   18
      Top             =   270
      Width           =   2460
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Demand ID"
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
      Left            =   4620
      TabIndex        =   17
      Top             =   2674
      Width           =   960
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove this Row"
      End
   End
End
Attribute VB_Name = "FrmCustomerDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vCounter As Integer
Dim Flag As Boolean
Dim sSql As String
Dim vStrSQL As String
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vStrComp As String, vCompanyName As String, vAddress As String, vemail As String
'----------------------------------

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
    vStrSQL = " Select * FROM Groups where GroupID='" & TxtGroupID.Text & "'"
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

Private Function FunSelectSubGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSubGroup.Show vbModal, Me
        If SchSubGroup.ParaOutSubGroupID = "" Then FunSelectSubGroup = False: Exit Function
        TxtSubGroupID.Text = SchSubGroup.ParaOutSubGroupID
    End If
    '---------------------------
    vStrSQL = " Select * FROM SubGroups where SubGroupID=" & Val(TxtSubGroupID.Text)
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

Private Function FunSelectVendor(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchVendor.Show vbModal, Me
        If SchVendor.ParaOutVendorID = "" Then FunSelectVendor = False: Exit Function
        TxtVendorID.Text = SchVendor.ParaOutVendorID
    End If
    '---------------------------
    If Trim(TxtVendorID.Text) = "" Then Exit Function
    vStrSQL = " Select * FROM Parties where PartyID = '" & TxtVendorID.Text & "' AND PartyType <> 'C'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtVendorName.Text = !PartyName
          FunSelectVendor = True
          .Close
          Exit Function
      Else
          FunSelectVendor = False
          .Close
          TxtVendorID.Text = ""
          TxtVendorName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnCompany_Click()
   If FunSelectCompany(ssButton, False) = True Then
      TxtSubGroupID.SetFocus
   Else
      TxtCompanyID.SetFocus
   End If
End Sub

Private Sub BtnGroup_Click()
   If FunSelectGroup(ssButton, False) = True Then
      TxtCompanyID.SetFocus
   Else
      TxtGroupID.SetFocus
   End If
End Sub

Private Sub BtnSubGroup_Click()
   If FunSelectSubGroup(ssButton, False) = True Then
      TxtVendorID.SetFocus
   Else
      TxtSubGroupID.SetFocus
   End If
End Sub

Private Sub BtnVendor_Click()
   If FunSelectVendor(ssButton, False) = True Then
      TxtRemarks.SetFocus
   Else
      TxtVendorID.SetFocus
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
   If Trim(TxtCompanyID.Text) = "" Then Exit Sub
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

Private Sub TxtSubGroupID_Change()
   If TxtSubGroupID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSubGroupID.Name Then Exit Sub
   If TxtSubGroupName.Text <> "" Then TxtSubGroupName.Text = ""
End Sub

Private Sub TxtSubGroupID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSubGroupID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtSubGroupID.Text) = "" Then Exit Sub
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

Private Sub TxtVendorID_Change()
   If TxtVendorID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtVendorID.Name Then Exit Sub
   If TxtVendorName.Text <> "" Then TxtVendorName.Text = ""
End Sub

Private Sub TxtVendorID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtVendorID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtVendorName.Text <> "" Then Exit Sub
   If Trim(TxtVendorID.Text) = "" Then Exit Sub
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

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   BtnPrint.Enabled = False
   
   vStrSQL = " Select D.*, StoreName, GroupName, SubGroupName, CompanyName, PartyName from CustomerDemands D Left Outer Join Stores st on st.storeid = d.storeid" & vbCrLf _
            + " Left outer Join Groups g on D.GroupID = g.GroupID " & vbCrLf _
            + " Left outer Join Companies c on D.CompanyID = c.CompanyID " & vbCrLf _
            + " Left outer Join SubGroups s on D.SubGroupID = s.SubGroupID " & vbCrLf _
            + " Left outer Join Parties p on d.VendorID = p.PartyID" & vbCrLf _
            + " where D.DemandID = " & Val(TxtDemandID.Text)

   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
  
   Set RptReportViewer.Report = New CrptCustomerDemand
   
   RptReportViewer.Report.ReportTitle = "Customer Demand"
   
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   With CN.Execute("Select CompanyName,Address,City,PhoneNo,email from Company")
      If .RecordCount > 0 Then
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(IsNull(!CompanyName), "", CStr(!CompanyName))
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(IsNull(!Address), "", !Address) & IIf(IsNull(!City), "", ", " & !City & ".")
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(IsNull(!PhoneNo), "", CStr(!PhoneNo))
      End If
    .Close
    End With
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
   Dim vDevice As String, vDriver As String, vPort As String
   vStrSQL = "Select * from Registry"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         vDevice = IIf(IsNull(!DeviceName), "Abc", !DeviceName)
         vDriver = IIf(IsNull(!DriverName), "Xyz", !DriverName)
         vPort = IIf(IsNull(!Port), "LPT1", !Port)
         RptReportViewer.Report.SelectPrinter vDriver, vDevice, vPort
      End If
   End With
   'RptReportViewer.Report.PaperOrientation = crPortrait
   RptReportViewer.Show
   'RptReportViewer.Report.PrintOut False
   BtnPrint.Enabled = True
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
    BtnPrint.Enabled = True
End Sub

Private Sub BtnClear_Click()
  On Error GoTo ErrorHandler
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
  End If
  If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
  CN.BeginTrans
  CN.Execute "Delete from CustomerDemands Where DemandID = " & Val(TxtDemandID.Text)
  CN.CommitTrans
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  If CN.Errors.Count > 0 Then CN.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Sub GetCustomerDemand()
   On Error GoTo ErrorHandler
   sSql = "Select D.*, StoreName, GroupName, SubGroupName, CompanyName, PartyName from CustomerDemands D Left outer Join Stores st on st.storeid = D.storeid Left outer Join Groups g on D.GroupID = g.GroupID Left outer Join SubGroups s on D.SubGroupID = s.SubGroupID Left outer Join Companies c on D.CompanyID = c.CompanyID Left outer Join Parties p on d.VendorID = p.PartyID where DemandID = " & Val(TxtDemandID.Text)
   With CN.Execute(sSql)
      If Not .BOF Then
          DtpDemandDate.DateValue = !DemandDate
          TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
          TxtStoreName.Text = IIf(IsNull(!StoreName), "", !StoreName)
          TxtGroupID.Text = IIf(IsNull(!GroupID), "", !GroupID)
          TxtGroupName.Text = IIf(IsNull(!GroupName), "", !GroupName)
          TxtCompanyID.Text = IIf(IsNull(!companyid), "", !companyid)
          TxtCompanyName.Text = IIf(IsNull(!CompanyName), "", !CompanyName)
          TxtSubGroupID.Text = IIf(IsNull(!SubGroupID), "", !SubGroupID)
          TxtSubGroupName.Text = IIf(IsNull(!SubGroupName), "", !SubGroupName)
          TxtVendorID.Text = IIf(IsNull(!VendorID), "", !VendorID)
          TxtVendorName.Text = IIf(IsNull(!PartyName), "", !PartyName)
          TxtProductName.Text = !ProductName
          TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
      End If
      .Close
   End With
'   Call PopulateDataToGrid
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchCustomerDemands.Show vbModal
   If SchCustomerDemands.ParaOutDemandID <> "" Then
      TxtDemandID.Text = SchCustomerDemands.ParaOutDemandID
      GetCustomerDemand
   End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If vIsNewRecord Then
      If CN.Execute("Select * from CustomerDemands where DemandID = " & Val(TxtDemandID.Text)).RecordCount > 0 Then
         MsgBox "This voucher already exists. A new voucher No. has been generated. Please try again", vbCritical, "Alert"
         TxtDemandID.Text = FunGetMaxID
         Exit Sub
      End If
   End If
'   RsBody.Filter = 0
'   If RsBody.RecordCount = 0 Then
'       MsgBox "Please enter at least one entry to save", vbExclamation, "Alert"
'       If TxtProductID.Visible And TxtProductID.Enabled Then TxtProductID.SetFocus
'       Exit Sub
'   End If
   
   
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
  
   CN.BeginTrans
   sSql = "Select * From CustomerDemands Where DemandID =" & Val(TxtDemandID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, CN, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !DemandID = Val(TxtDemandID.Text)
      End If
      !DemandDate = DtpDemandDate.DateValue
      !StoreID = IIf(Trim(TxtStoreID.Text) = "", Null, Val(TxtStoreID.Text))
      !GroupID = IIf(Trim(TxtGroupID.Text) = "", Null, (TxtGroupID.Text))
      !companyid = IIf(Trim(TxtCompanyID.Text) = "", Null, Val(TxtCompanyID.Text))
      !SubGroupID = IIf(Trim(TxtSubGroupID.Text) = "", Null, Val(TxtSubGroupID.Text))
      !VendorID = IIf(Trim(TxtVendorID.Text) = "", Null, Val(TxtVendorID.Text))
      !ProductName = Trim(TxtProductName.Text)
      !Remarks = IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text)
      !UserNo = vUser
      .Update
      .Close
   End With
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Property Get FormStatus() As FormMode
  'Nothing
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
  'Based upon the value of vNewValue, we shall decide what controls to enable/disable
  On Error GoTo ErrorHandler
  vMode = vNewValue
  Select Case vNewValue
    Case Is = NewMode
      Call SubClearFields
'      Call PopulateDataToGrid
      BtnPrint.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtDemandID.Text = FunGetMaxID
      If DtpDemandDate.Enabled And DtpDemandDate.Visible Then DtpDemandDate.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      BtnPrint.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      DtpDemandDate.SetFocus
      vIsNewRecord = False
    Case Is = changeMode
      BtnPrint.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
'  ElseIf KeyCode = vbKeyEscape And (Me.ActiveControl.Name = TxtProductID.Name Or Me.ActiveControl.Name = TxtProductName.Name Or Me.ActiveControl.Name = TxtProductName.Name Or Me.ActiveControl.Name = TxtUnderQty.Name Or Me.ActiveControl.Name = Grid.Name) Then
'    Call ClearDetailArea
  ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSave.Enabled Then BtnSave_Click
            KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
         Case vbKeyO
            If BtnOpen.Enabled Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
      End Select
  ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtProductName.SetFocus
         Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus
         Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtSubGroupID.SetFocus
         Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then TxtVendorID.SetFocus
         Case TxtVendorID.Name: If FunSelectVendor(ssFunctionKey, True) = True Then TxtRemarks.SetFocus
      End Select
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If BtnSave.Enabled Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then FormStatus = changeMode
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Customer Demand"
   DtpDemandDate.DateValue = Date
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
  On Error GoTo ErrorHandler
  FunGetMaxID = CN.Execute("Select isnull(max(DemandID),0) from CustomerDemands").Fields(0) + 1
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub SubClearFields()
  On Error GoTo ErrorHandler
  Dim ctl As Control
  For Each ctl In Me.Controls
    If TypeOf ctl Is TextBox Or TypeOf ctl Is SITextBox.txt Then
      ctl.Text = ""
    ElseIf TypeOf ctl Is ComboBox Then
    
    End If
  Next
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Are you sure to close without save?", vbQuestion + vbApplicationModal + vbYesNo, "Alert") = vbNo Then
         Cancel = 1
      End If
   Else
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set RsReport = Nothing
    Set FrmCustomerDemand = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      DtpDemandDate.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub TxtStoreID_Change()
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtStoreID.Text = "" Then Exit Sub
   If TxtStoreName.Text <> "" Then Exit Sub
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

Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectStore = False: Exit Function
        TxtStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Stores where StoreID=" & Val(TxtStoreID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtStoreName.Text = !StoreName
          FunSelectStore = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = changeMode
          Exit Function
      Else
          FunSelectStore = False
          .Close
          TxtStoreID.Text = ""
          TxtStoreName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = changeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
