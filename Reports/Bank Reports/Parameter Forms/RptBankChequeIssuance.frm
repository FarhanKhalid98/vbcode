VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptBankChequeIssuance 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptBankChequeIssuance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkVooucherDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Voucher Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10200
      TabIndex        =   39
      Top             =   3720
      Width           =   1290
   End
   Begin VB.CheckBox ChkChequeDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Cheque Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10200
      TabIndex        =   37
      Top             =   3000
      Width           =   1290
   End
   Begin VB.OptionButton OptAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8843
      TabIndex        =   10
      Top             =   6548
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton OptReturn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Return"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7658
      TabIndex        =   9
      Top             =   6548
      Width           =   1185
   End
   Begin VB.OptionButton OptBounce 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bounce"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6488
      TabIndex        =   8
      Top             =   6548
      Width           =   1185
   End
   Begin VB.OptionButton OptReconcile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reconcile"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5318
      TabIndex        =   7
      Top             =   6548
      Width           =   1185
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8430
      TabIndex        =   13
      Top             =   7268
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
      MICON           =   "RptBankChequeIssuance.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   5655
      TabIndex        =   11
      Top             =   7268
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
      MICON           =   "RptBankChequeIssuance.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7035
      TabIndex        =   12
      Top             =   7268
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
      MICON           =   "RptBankChequeIssuance.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnACPayee 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6938
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5123
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
      MICON           =   "RptBankChequeIssuance.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtACPayeeID 
      Height          =   315
      Left            =   5318
      TabIndex        =   4
      Top             =   5123
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   15
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtACPayeeName 
      Height          =   315
      Left            =   7298
      TabIndex        =   17
      Tag             =   "nc"
      Top             =   5123
      Width           =   2730
      _ExtentX        =   4815
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
   Begin JeweledBut.JeweledButton BtnReceive 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9683
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5873
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
      MICON           =   "RptBankChequeIssuance.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnBank 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6938
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4418
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
      MICON           =   "RptBankChequeIssuance.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt txtBankID 
      Height          =   315
      Left            =   5318
      TabIndex        =   3
      Top             =   4418
      Width           =   1620
      _ExtentX        =   2858
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
      Masked          =   1
      IntegralPoint   =   3
      Mandatory       =   1
   End
   Begin SITextBox.Txt txtBankName 
      Height          =   315
      Left            =   7298
      TabIndex        =   15
      Tag             =   "nc"
      Top             =   4418
      Width           =   2640
      _ExtentX        =   4657
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
   Begin SITextBox.Txt TxtChequeNo 
      Height          =   315
      Left            =   5318
      TabIndex        =   5
      Top             =   5873
      Width           =   1620
      _ExtentX        =   2858
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
   Begin JeweledBut.JeweledButton BtnChequeNo 
      Height          =   330
      Left            =   6938
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5873
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
      MICON           =   "RptBankChequeIssuance.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtReceiveBy 
      Height          =   315
      Left            =   7298
      TabIndex        =   6
      Top             =   5873
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   40
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtVoucherID 
      Height          =   315
      Left            =   5318
      TabIndex        =   0
      Top             =   3623
      Width           =   1620
      _ExtentX        =   2858
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
      Masked          =   1
      IntegralPoint   =   15
      Mandatory       =   1
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   195
         Left            =   1995
         TabIndex        =   32
         Top             =   0
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   195
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   30
      End
   End
   Begin JeweledBut.JeweledButton BtnVoucher 
      Height          =   330
      Left            =   6938
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3623
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
      MICON           =   "RptBankChequeIssuance.frx":0F8E
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   315
      Left            =   7298
      TabIndex        =   1
      Top             =   3623
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpToDate 
      Height          =   315
      Left            =   8603
      TabIndex        =   2
      Top             =   3623
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpChequeDateFrom 
      Height          =   315
      Left            =   7200
      TabIndex        =   34
      Top             =   2940
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpChequeDateTo 
      Height          =   315
      Left            =   8505
      TabIndex        =   38
      Top             =   2940
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
   Begin VB.Label Label11 
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
      Left            =   8505
      TabIndex        =   36
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label Label7 
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
      Left            =   7200
      TabIndex        =   35
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label Label6 
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
      Left            =   7305
      TabIndex        =   30
      Top             =   3443
      Width           =   885
   End
   Begin VB.Label Label3 
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
      Left            =   8610
      TabIndex        =   29
      Top             =   3443
      Width           =   705
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receive By"
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
      Left            =   7298
      TabIndex        =   28
      Top             =   5693
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Payee Name"
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
      Left            =   7298
      TabIndex        =   27
      Top             =   4928
      Width           =   1470
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No"
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
      Left            =   5325
      TabIndex        =   26
      Top             =   5693
      Width           =   960
   End
   Begin VB.Label LblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Cheque Issuance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2700
      TabIndex        =   25
      Top             =   270
      Width           =   3510
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Payee ID"
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
      Left            =   5325
      TabIndex        =   24
      Top             =   4928
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C  Name"
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
      Left            =   7298
      TabIndex        =   23
      Top             =   4208
      Width           =   1440
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C ID"
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
      Left            =   5325
      TabIndex        =   22
      Top             =   4208
      Width           =   1095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   -1410
      TabIndex        =   21
      Top             =   7635
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher ID"
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
      Left            =   5325
      TabIndex        =   20
      Top             =   3443
      Width           =   975
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
Attribute VB_Name = "RptBankChequeIssuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql As String
Dim VStrSQL As String

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchBankAct.Show vbModal, Me
      If SchBankAct.ParaOutID = "" Then FunSelectAccount = False: Exit Function
      txtBankID.Text = SchBankAct.ParaOutID
   End If
   Dim VStrSQL As String
   VStrSQL = "select * from ChartofAccounts where AccountNo =  '" & Val(txtBankID.Text) & "'"
   With cn.Execute(VStrSQL)
         If .RecordCount > 0 Then
            txtBankName.Text = !AccountName
            .Close
            FunSelectAccount = True
            Exit Function
         Else
            FunSelectAccount = False
            .Close
            txtBankName.Text = ""
           End If
      End With
      Exit Function
ErrorHandler:
      Call ShowErrorMessage
End Function

Private Function FunSelectPayee(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchActPayee.Show vbModal, Me
      If SchActPayee.ParaOutID = "" Then FunSelectPayee = False: Exit Function
      TxtACPayeeID.Text = SchActPayee.ParaOutID
   End If
   Dim VStrSQL As String
   VStrSQL = "select * from Parties where PartyID =  '" & Val(TxtACPayeeID.Text) & "'"
   With cn.Execute(VStrSQL)
         If .RecordCount > 0 Then
            TxtACPayeeName.Text = !PartyName
'            TxtVendorAddress.Text = IIf(IsNull(!address), "", !address)
'            TxtVendorCity.Text = IIf(IsNull(!city), "", !city)
            .Close
            FunSelectPayee = True
           
            Exit Function
         Else
            FunSelectPayee = False
            .Close
            TxtACPayeeName.Text = ""
'            TxtACPayeeID.Text = ""
'            TxtVendorAddress.Text = ""
'            TxtVendorCity.Text = ""
            
         End If
      End With
      Exit Function
ErrorHandler:
      Call ShowErrorMessage
End Function

Private Function FunSelectReceiveBy(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchChqReceiveBy.Show vbModal, Me
        If SchChqReceiveBy.ParaOutDepositBy = "" Then FunSelectReceiveBy = False: Exit Function
        TxtReceiveBy.Text = SchChqReceiveBy.ParaOutDepositBy
    End If
    '---------------------------
    VStrSQL = "Select ReceiveBy FROM BankChequeIssueBody where ReceiveBy = '" & TxtReceiveBy.Text & "'"
    With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
          FunSelectReceiveBy = True
          .Close
          Exit Function
      Else
          FunSelectReceiveBy = False
          .Close
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectCheque(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchChequeIssue.Show vbModal, Me
      If SchChequeIssue.ParaOutID = "" Then FunSelectCheque = False: Exit Function
      TxtChequeNo.Text = SchChequeIssue.ParaOutID
   End If
   Dim VStrSQL As String
   VStrSQL = "select ChequeNo from BankChequeIssueBody Where ChequeNo =  '" & Val(TxtChequeNo.Text) & "'"
   With cn.Execute(VStrSQL)
         If .RecordCount > 0 Then
            .Close
            FunSelectCheque = True
            Exit Function
         Else
            FunSelectCheque = False
            .Close
            TxtChequeNo.Text = ""
         End If
      End With
      Exit Function
ErrorHandler:
      Call ShowErrorMessage
End Function

Private Function FunSelectVoucher(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchChequeIssue.Show vbModal, Me
      If SchChequeIssue.ParaOutID = Null Then FunSelectVoucher = False: Exit Function
      TxtVoucherID.Text = SchChequeIssue.ParaOutID
   End If
    '---------------------------
    If Trim(TxtVoucherID.Text) = "" Then Exit Function
'    If Len(TxtVoucherID.Text) <= 5 Then
'      TxtVoucherID.Text = Right("00000" + CStr(Val(TxtVoucherID.Text)), 5)
'    End If
    If TxtVoucherID.Text = "" Then FunSelectVoucher = False: Exit Function
    VStrSQL = " SELECT VoucherID, VoucherDate From BankChequeIssueHeader Where VoucherID = " & TxtVoucherID.Text
  
   With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtVoucherID.Text = !VoucherID
         DtpFromDate.Date = !VoucherDate
         DtpToDate.Date = !VoucherDate
         FunSelectVoucher = True
         .Close
         Exit Function
      Else
         FunSelectVoucher = False
         .Close
        ' MsgBox "Invalid VoucherID ID.", vbOKOnly, "Alert"
         TxtVoucherID.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnChequeNo_Click()
   If FunSelectCheque(ssButton, True) = True Then
    TxtReceiveBy.SetFocus
   Else
      TxtChequeNo.SetFocus
   End If
End Sub

Private Sub BtnVoucher_Click()
   If FunSelectVoucher(ssButton, True) = True Then
      DtpFromDate.SetFocus
   Else
      TxtVoucherID.SetFocus
   End If
End Sub

Private Sub BtnACPayee_Click()
   If FunSelectPayee(ssButton, False) = True Then
      TxtChequeNo.SetFocus
   Else
      TxtACPayeeID.SetFocus
   End If
End Sub

Private Sub btnBank_Click()
   If FunSelectAccount(ssButton, False) = True Then
      TxtACPayeeID.SetFocus
   Else
      txtBankID.SetFocus
   End If
End Sub

Private Sub BtnReceive_Click()
   If FunSelectReceiveBy(ssButton, False) = True Then
      BtnPreview.SetFocus
   Else
      TxtReceiveBy.SetFocus
   End If
End Sub

Private Sub TxtChequeNo_Change()
   If TxtChequeNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtChequeNo.Name Then Exit Sub
End Sub

Private Sub TxtChequeNo_Validate(Cancel As Boolean)
  If Me.ActiveControl.Name <> TxtChequeNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtChequeNo.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCheque(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCheque(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtReceiveBy_Change()
 If TxtReceiveBy.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtReceiveBy.Name Then Exit Sub
   'If txtBankName.Text <> "" Then txtBankName.Text = ""
End Sub

Private Sub TxtReceiveBy_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtReceiveBy.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtReceiveBy.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectReceiveBy(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectReceiveBy(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtVoucherID_Change()
    If TxtVoucherID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtVoucherID.Name Then Exit Sub
End Sub

Private Sub TxtVoucherID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtVoucherID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtVoucherID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVoucher(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVoucher(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub txtBankID_Change()
   If txtBankID.Visible = False Then Exit Sub
   If ActiveControl.Name <> txtBankID.Name Then Exit Sub
   If txtBankName.Text <> "" Then txtBankName.Text = ""
End Sub

Private Sub txtBankID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> txtBankID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If txtBankID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtACPayeeID_change()
   If TxtACPayeeID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtACPayeeID.Name Then Exit Sub
   If TxtACPayeeName.Text <> "" Then TxtACPayeeName.Text = ""
End Sub

Private Sub TxtACPayeeID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtACPayeeID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtACPayeeID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectPayee(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectPayee(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtACPayeeName_Change()
   If TxtACPayeeName.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtACPayeeName.Name Then Exit Sub
   
End Sub

Private Sub TxtACPayeeName_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtACPayeeName.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtACPayeeName.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectReceiveBy(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectReceiveBy(ssButton, False)
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
       RptReportViewer.Caption = "Cheque Issuance"
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
         Case TxtVoucherID.Name: If FunSelectVoucher(ssFunctionKey, True) = True Then DtpFromDate.SetFocus
         Case txtBankID.Name: If FunSelectAccount(ssFunctionKey, True) = True Then TxtACPayeeID.SetFocus
         Case TxtACPayeeID.Name: If FunSelectPayee(ssFunctionKey, True) = True Then TxtChequeNo.SetFocus
         Case TxtChequeNo.Name: If FunSelectCheque(ssFunctionKey, True) = True Then TxtReceiveBy.SetFocus
         Case TxtReceiveBy.Name: If FunSelectReceiveBy(ssFunctionKey, True) = True Then BtnPreview.SetFocus
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
   SetWindowText Me.hWnd, "Bank Cheque Issuance"
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
   Set RptBankChequeIssuance = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function SetReport() As Boolean
    On Error GoTo ErrorHandler
    Dim RsReport As New ADODB.Recordset
    SetReport = False
'    VStrSQL = "Select OS.*, (Isnull(OS.QtyPack,0) * Isnull(OS.Multiplier,0)) + Isnull(OS.QtyLoose,0) NetQty, P.ProductName, PK.PackingName, S.StoreName, C.CompanyName, G.GroupName, SG.SubGroupName" & vbCrLf _
'    + "from OpeningStock OS" & vbCrLf _
'    + "Left Outer Join  Products P on P.ProductID = OS.ProductID" & vbCrLf _
'    + "Left Outer Join   Packings PK on OS.PackingID = PK.PackingID" & vbCrLf _
'    + "Left Outer Join   Stores S on S.VoucherID = OS.VoucherID" & vbCrLf _
'    + "Left Outer Join Companies C on P.comPanyID = C.comPanyID" & vbCrLf _
'    + "Left Outer Join Groups G on P.GroupID = G.GroupID" & vbCrLf _
'    + "Left Outer Join SubGroups SG on P.SubGroupID = SG.SubGroupID Where 1=1 " & IIf(Trim(TxtVoucherID.Text) = "", "", " And S.VoucherID = " & TxtVoucherID.Text) & vbCrLf _
'    + "" & IIf(Trim(txtBankID.Text) = "", "", " And C.CompanyID = " & txtBankID.Text) & vbCrLf _
'    + "" & IIf(Trim(TxtACPayeeID.Text) = "", "", " And G.GroupID = " & TxtACPayeeID.Text) & vbCrLf _
'    + "" & IIf(Trim(TxtACPayeeName.Text) = "", "", " And SG.SubGroupID = " & TxtACPayeeName.Text) & vbCrLf _
'    + "" & IIf(Trim(TxtChequeNo.Text) = "", "", " And P.ProductID = " & TxtChequeNo.Text)
    Me.MousePointer = vbHourglass
'    If RsReport.State = adStateOpen Then RsReport.Close
'   RsReport.Open VStrSQL, CN, adOpenStatic, adLockReadOnly

   CrptBankChequeIssuParameter.DiscardSavedData
   
   VStrSQL = "ProdRptBankChequeIssue '" & DtpFromDate.Date & "','" & DtpToDate.Date & "','" & DtpChequeDateFrom.Date & "','" & DtpChequeDateTo.Date & "','" & TxtVoucherID.Text & "','" & txtBankID.Text & "','" & TxtACPayeeID.Text & "','" & TxtReceiveBy.Text & "','" & TxtChequeNo.Text & "'," & ChkChequeDate.Value & "," & ChkVooucherDate.Value & "," & Abs(OptReconcile.Value) & "," & Abs(OptBounce.Value) & "," & Abs(OptReturn.Value)
   Set RsReport = cn.Execute(VStrSQL)
  
   Set RptReportViewer.Report = CrptBankChequeIssuParameter
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
        
   If RsReport.BOF Then
       MsgBox "No record exists.", vbInformation, Me.Caption
       Me.MousePointer = vbDefault
       Exit Function
   End If
    
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.PaperOrientation = crLandscape
   SetReport = True
   Me.MousePointer = vbDefault
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
   Me.MousePointer = vbDefault
End Function
