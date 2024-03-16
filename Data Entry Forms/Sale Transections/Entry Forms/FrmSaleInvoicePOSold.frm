VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmSaleInvoicePOSold 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10950
   ClientLeft      =   -3210
   ClientTop       =   480
   ClientWidth     =   15360
   Icon            =   "FrmSaleInvoicePOSold.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   758
   ScaleMode       =   0  'User
   ScaleWidth      =   901.419
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   4770
      TabIndex        =   122
      Top             =   3645
      Width           =   6090
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   1080
         TabIndex        =   133
         Top             =   135
         Width           =   3525
         Begin VB.OptionButton OptCash 
            Caption         =   "&Cash"
            Height          =   285
            Left            =   210
            TabIndex        =   104
            Tag             =   "F"
            Top             =   240
            Width           =   765
         End
         Begin VB.OptionButton OptBankCard 
            Caption         =   "&Bank Card"
            Height          =   285
            Left            =   2100
            TabIndex        =   106
            Tag             =   "F"
            Top             =   240
            Width           =   1125
         End
         Begin VB.OptionButton OptCredit 
            Caption         =   "&Credit"
            Height          =   285
            Left            =   1200
            TabIndex        =   105
            Tag             =   "F"
            Top             =   240
            Value           =   -1  'True
            Width           =   765
         End
      End
      Begin VB.CheckBox ChkPrint 
         Caption         =   "&Print"
         Height          =   285
         Left            =   180
         TabIndex        =   102
         Tag             =   "F"
         Top             =   315
         Value           =   1  'Checked
         Width           =   705
      End
      Begin JeweledBut.JeweledButton BtnCancel 
         Height          =   420
         Left            =   2940
         TabIndex        =   129
         Tag             =   "F"
         Top             =   3975
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   741
         TX              =   "Cancel"
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
         MICON           =   "FrmSaleInvoicePOSold.frx":0ECA
         BC              =   14737632
         FC              =   0
      End
      Begin JeweledBut.JeweledButton BtnOk 
         Height          =   420
         Left            =   1635
         TabIndex        =   127
         Tag             =   "F"
         Top             =   3975
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   741
         TX              =   "OK"
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
         MICON           =   "FrmSaleInvoicePOSold.frx":0EE6
         BC              =   14737632
         FC              =   0
      End
      Begin VB.Frame FrameCredit 
         BorderStyle     =   0  'None
         Height          =   2385
         Left            =   15
         TabIndex        =   145
         Top             =   1305
         Width           =   5805
         Begin VB.TextBox TxtNetAmountCredit 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   146
            Tag             =   "F"
            Top             =   180
            Width           =   2025
         End
         Begin SITextBox.Txt TxtCustomerID 
            Height          =   315
            Left            =   375
            TabIndex        =   109
            Tag             =   "F"
            Top             =   1380
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
            IntegralPoint   =   10
            Mandatory       =   1
         End
         Begin SITextBox.Txt TxtCustomerName 
            Height          =   315
            Left            =   1755
            TabIndex        =   147
            Tag             =   "F"
            Top             =   1380
            Width           =   3150
            _ExtentX        =   5556
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
         Begin JeweledBut.JeweledButton BtnCustomer 
            CausesValidation=   0   'False
            Height          =   330
            Left            =   1395
            TabIndex        =   148
            TabStop         =   0   'False
            Tag             =   "F"
            Top             =   1380
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
            MICON           =   "FrmSaleInvoicePOSold.frx":0F02
            BC              =   12632256
            FC              =   0
         End
         Begin SITextBox.Txt TxtCashReceivedCredit 
            Height          =   315
            Left            =   2025
            TabIndex        =   108
            Tag             =   "F"
            Top             =   675
            Width           =   2025
            _ExtentX        =   3572
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
         End
         Begin JeweledBut.JeweledButton BtnAddCustomer 
            CausesValidation=   0   'False
            Height          =   330
            Left            =   4920
            TabIndex        =   210
            TabStop         =   0   'False
            Tag             =   "F"
            Top             =   1380
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   582
            TX              =   "+"
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
            MICON           =   "FrmSaleInvoicePOSold.frx":0F1E
            BC              =   12632256
            FC              =   0
         End
         Begin SITextBox.Txt TxtPreviousReceivable 
            Height          =   315
            Left            =   4260
            TabIndex        =   213
            Top             =   855
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Alignment       =   1
            Appearance      =   0
            Enabled         =   0   'False
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
            Masked          =   2
         End
         Begin SITextBox.Txt TxtRefID 
            Height          =   315
            Left            =   4200
            TabIndex        =   215
            Top             =   315
            Visible         =   0   'False
            Width           =   600
            _ExtentX        =   1058
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
         Begin SITextBox.Txt TxtRefComm 
            Height          =   315
            Left            =   5130
            TabIndex        =   216
            Top             =   330
            Visible         =   0   'False
            Width           =   600
            _ExtentX        =   1058
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
         Begin SITextBox.Txt TxtBankMachineCreditID 
            Height          =   315
            Left            =   405
            TabIndex        =   219
            Tag             =   "F"
            Top             =   2010
            Width           =   1335
            _ExtentX        =   2355
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
            Mandatory       =   1
         End
         Begin SITextBox.Txt TxtBankMachineCreditName 
            Height          =   315
            Left            =   2100
            TabIndex        =   220
            Tag             =   "F"
            Top             =   2010
            Width           =   2280
            _ExtentX        =   4022
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
         Begin JeweledBut.JeweledButton BtnBankMachineCredit 
            CausesValidation=   0   'False
            Height          =   330
            Left            =   1740
            TabIndex        =   221
            TabStop         =   0   'False
            Tag             =   "F"
            Top             =   2010
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
            MICON           =   "FrmSaleInvoicePOSold.frx":0F3A
            BC              =   12632256
            FC              =   0
         End
         Begin SITextBox.Txt TxtBankAmount 
            Height          =   315
            Left            =   4455
            TabIndex        =   225
            Tag             =   "F"
            Top             =   2010
            Width           =   1125
            _ExtentX        =   1984
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
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4455
            TabIndex        =   224
            Top             =   1800
            Width           =   1110
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Machine Name"
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
            Left            =   2100
            TabIndex        =   223
            Top             =   1800
            Width           =   1770
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Machine ID"
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
            Left            =   405
            TabIndex        =   222
            Top             =   1800
            Width           =   1485
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Reference ID Comm %"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4200
            TabIndex        =   217
            Top             =   90
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label lblPayable 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Previous Receivable"
            Height          =   195
            Left            =   4275
            TabIndex        =   214
            Top             =   630
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Net Amount"
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
            Left            =   870
            TabIndex        =   123
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
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
            Left            =   1755
            TabIndex        =   151
            Top             =   1170
            Width           =   1335
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
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
            Left            =   375
            TabIndex        =   150
            Top             =   1170
            Width           =   1050
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Received"
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
            Left            =   585
            TabIndex        =   149
            Top             =   705
            Width           =   1305
         End
      End
      Begin VB.Frame FrameBank 
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   270
         TabIndex        =   134
         Top             =   720
         Width           =   5025
         Begin VB.TextBox TxtCommision 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   136
            Top             =   270
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox TxtNetAmountBank 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   135
            Tag             =   "F"
            Top             =   720
            Width           =   1080
         End
         Begin SITextBox.Txt TxtBankMachineID 
            Height          =   315
            Left            =   675
            TabIndex        =   113
            Tag             =   "F"
            Top             =   2235
            Width           =   1335
            _ExtentX        =   2355
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
            Mandatory       =   1
         End
         Begin SITextBox.Txt TxtBankMachineName 
            Height          =   315
            Left            =   2340
            TabIndex        =   137
            Tag             =   "F"
            Top             =   2235
            Width           =   2685
            _ExtentX        =   4736
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
         Begin JeweledBut.JeweledButton BtnBankMachine 
            CausesValidation=   0   'False
            Height          =   330
            Left            =   1980
            TabIndex        =   138
            TabStop         =   0   'False
            Tag             =   "F"
            Top             =   2235
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
            MICON           =   "FrmSaleInvoicePOSold.frx":0F56
            BC              =   12632256
            FC              =   0
         End
         Begin SITextBox.Txt TxtBankCustomer 
            Height          =   315
            Left            =   675
            TabIndex        =   110
            Tag             =   "F"
            Top             =   285
            Width           =   4350
            _ExtentX        =   7673
            _ExtentY        =   556
            Appearance      =   0
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
         Begin SITextBox.Txt TxtInvoiceNo 
            Height          =   315
            Left            =   1560
            TabIndex        =   112
            Tag             =   "F"
            Top             =   1605
            Width           =   2025
            _ExtentX        =   3572
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
         End
         Begin SITextBox.Txt TxtCashReceivedBank 
            Height          =   315
            Left            =   1590
            TabIndex        =   111
            Tag             =   "F"
            Top             =   1170
            Width           =   2025
            _ExtentX        =   3572
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
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Machine ID"
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
            Left            =   645
            TabIndex        =   144
            Top             =   2025
            Width           =   1485
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Machine Name"
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
            TabIndex        =   143
            Top             =   2025
            Width           =   1770
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Net Amount"
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
            Left            =   420
            TabIndex        =   142
            Top             =   750
            Width           =   1005
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice No"
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
            Left            =   480
            TabIndex        =   141
            Top             =   1635
            Width           =   945
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   675
            TabIndex        =   140
            Top             =   45
            Width           =   1665
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Received"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   210
            TabIndex        =   139
            Top             =   1215
            Width           =   1215
         End
      End
      Begin VB.Frame FrameCash 
         BorderStyle     =   0  'None
         Height          =   2790
         Left            =   450
         TabIndex        =   124
         Top             =   900
         Width           =   4425
         Begin VB.TextBox TxtCashReturn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   203
            TabStop         =   0   'False
            Tag             =   "F"
            Top             =   2295
            Width           =   2385
         End
         Begin VB.TextBox TxtCashReceivedCash 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1470
            MaxLength       =   9
            TabIndex        =   126
            Top             =   1800
            Width           =   2385
         End
         Begin VB.TextBox TxtNetAmountCash 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   125
            TabStop         =   0   'False
            Tag             =   "F"
            Top             =   1305
            Width           =   2385
         End
         Begin SITextBox.Txt TxtCashCustomer 
            Height          =   315
            Left            =   90
            TabIndex        =   107
            Tag             =   "F"
            Top             =   360
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   556
            Appearance      =   0
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
         Begin SITextBox.Txt TxtCNIC 
            Height          =   315
            Left            =   510
            TabIndex        =   204
            Tag             =   "F"
            Top             =   720
            Width           =   1665
            _ExtentX        =   2937
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
            Masked          =   5
         End
         Begin SITextBox.Txt TxtCellNo 
            Height          =   315
            Left            =   3015
            TabIndex        =   205
            Tag             =   "F"
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
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
            Masked          =   5
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Cell: No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2280
            TabIndex        =   202
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "CNIC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   30
            TabIndex        =   201
            Top             =   720
            Width           =   390
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Return"
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
            Left            =   270
            TabIndex        =   132
            Top             =   2400
            Width           =   1065
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Received"
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
            Left            =   30
            TabIndex        =   131
            Top             =   1935
            Width           =   1305
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Net Amount"
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
            Left            =   330
            TabIndex        =   130
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00DEAB97&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   30
            TabIndex        =   128
            Top             =   120
            Width           =   1665
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   1170
      TabIndex        =   188
      Top             =   4365
      Visible         =   0   'False
      Width           =   2295
      Begin SITextBox.Txt TxtSerial 
         Height          =   315
         Left            =   120
         TabIndex        =   189
         Top             =   240
         Width           =   2040
         _ExtentX        =   3598
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
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridSerial 
         Height          =   1500
         Left            =   120
         TabIndex        =   190
         Top             =   555
         Width           =   2040
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   2
         stylesets.count =   1
         stylesets(0).Name=   "SelectedRow"
         stylesets(0).ForeColor=   -2147483634
         stylesets(0).BackColor=   -2147483635
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "FrmSaleInvoicePOSold.frx":0F72
         AllowDelete     =   -1  'True
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   1
         ForeColorEven   =   0
         BackColorOdd    =   15724527
         RowHeight       =   423
         ExtraHeight     =   26
         ActiveRowStyleSet=   "SelectedRow"
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "ProductID"
         Columns(0).Name =   "ProductID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3096
         Columns(1).Caption=   "Serial No"
         Columns(1).Name =   "Serial"
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   3598
         _ExtentY        =   2646
         _StockProps     =   79
         BackColor       =   15724527
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ComboBox cmbSizeName 
      Height          =   315
      Left            =   7815
      Style           =   2  'Dropdown List
      TabIndex        =   182
      Top             =   3165
      Width           =   840
   End
   Begin VB.ComboBox CmbColourName 
      Height          =   315
      Left            =   6615
      Style           =   2  'Dropdown List
      TabIndex        =   181
      Top             =   3165
      Width           =   1200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   4860
      Top             =   585
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   4140
      Top             =   585
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      Left            =   3525
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   9555
      Width           =   1950
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   735
      MaxLength       =   50
      TabIndex        =   75
      Top             =   10695
      Visible         =   0   'False
      Width           =   4125
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4875
      Top             =   10500
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame FraHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Left            =   13995
      TabIndex        =   67
      Top             =   810
      Visible         =   0   'False
      Width           =   4200
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4200
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   68
         Tag             =   "NC"
         Text            =   "FrmSaleInvoicePOSold.frx":0F8E
         Top             =   300
         Width           =   3975
      End
      Begin VB.Label LblClose 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3915
         TabIndex        =   69
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.CheckBox ChkIsProduct 
      Caption         =   "Is Product"
      Height          =   255
      Left            =   7035
      TabIndex        =   54
      Top             =   735
      Visible         =   0   'False
      Width           =   1050
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   2340
      TabIndex        =   34
      Top             =   1410
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   9
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
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8025
      TabIndex        =   32
      Top             =   10080
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
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
      MICON           =   "FrmSaleInvoicePOSold.frx":1105
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6750
      TabIndex        =   28
      Top             =   10080
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
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
      MICON           =   "FrmSaleInvoicePOSold.frx":1121
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4095
      TabIndex        =   30
      Top             =   10080
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
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
      MICON           =   "FrmSaleInvoicePOSold.frx":113D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9345
      TabIndex        =   33
      Top             =   10080
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
      MICON           =   "FrmSaleInvoicePOSold.frx":1159
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5385
      TabIndex        =   29
      Top             =   10080
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
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
      MICON           =   "FrmSaleInvoicePOSold.frx":1175
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2715
      TabIndex        =   31
      Top             =   10095
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
      MICON           =   "FrmSaleInvoicePOSold.frx":1191
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtActualAmount 
      Height          =   315
      Left            =   9150
      TabIndex        =   41
      Top             =   10635
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   556
      Alignment       =   1
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
      Left            =   4380
      TabIndex        =   0
      Tag             =   "NC"
      Top             =   1410
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   5415
      TabIndex        =   43
      Tag             =   "NC"
      Top             =   1410
      Width           =   1395
      _ExtentX        =   2461
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
      Left            =   5055
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1410
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
      MICON           =   "FrmSaleInvoicePOSold.frx":11AD
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpBillDate 
      Height          =   315
      Left            =   3060
      TabIndex        =   152
      Tag             =   "NC"
      Top             =   1410
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      Enabled         =   0   'False
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
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SITextBox.Txt TxtPID 
      Height          =   315
      Left            =   10095
      TabIndex        =   47
      Top             =   10680
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
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
   Begin SITextBox.Txt TxtCost 
      Height          =   315
      Left            =   8610
      TabIndex        =   49
      Top             =   10680
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   11595
      TabIndex        =   7
      Top             =   2085
      Width           =   750
      _ExtentX        =   1323
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
   End
   Begin SITextBox.Txt TxtEmployeeName 
      Height          =   315
      Left            =   12705
      TabIndex        =   59
      Top             =   2085
      Width           =   1530
      _ExtentX        =   2699
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
   Begin JeweledBut.JeweledButton BtnEmployee 
      Height          =   330
      Left            =   12345
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2085
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
      MICON           =   "FrmSaleInvoicePOSold.frx":11C9
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCommission 
      Height          =   315
      Left            =   8160
      TabIndex        =   63
      Top             =   10680
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtMemberID 
      Height          =   315
      Left            =   8250
      TabIndex        =   2
      Top             =   1410
      Width           =   1440
      _ExtentX        =   2540
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
   Begin SITextBox.Txt TxtMemberName 
      Height          =   315
      Left            =   10050
      TabIndex        =   71
      Top             =   1410
      Width           =   1395
      _ExtentX        =   2461
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
   Begin JeweledBut.JeweledButton BtnMember 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   9690
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   1410
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
      MICON           =   "FrmSaleInvoicePOSold.frx":11E5
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtManualBillNo 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10995
      TabIndex        =   27
      Top             =   10230
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Alignment       =   1
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
      IntegralPoint   =   4
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   8610
      TabIndex        =   6
      Tag             =   "NC"
      Top             =   2085
      Width           =   705
      _ExtentX        =   1244
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
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   9675
      TabIndex        =   78
      Tag             =   "NC"
      Top             =   2085
      Width           =   1845
      _ExtentX        =   3254
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
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   9315
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   2085
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
      MICON           =   "FrmSaleInvoicePOSold.frx":1201
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillDisc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   915
      TabIndex        =   17
      Top             =   7845
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtBillDiscPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1755
      TabIndex        =   18
      Top             =   7845
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtServiceCharges 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   915
      TabIndex        =   21
      Top             =   8970
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtServiceChargesPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1755
      TabIndex        =   22
      Top             =   8970
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtSTax 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   915
      TabIndex        =   19
      Top             =   8415
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtSTaxPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1755
      TabIndex        =   20
      Top             =   8415
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtTableID 
      Height          =   315
      Left            =   915
      TabIndex        =   23
      Top             =   9555
      Width           =   525
      _ExtentX        =   926
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
   End
   Begin SITextBox.Txt TxtTableName 
      Height          =   315
      Left            =   1800
      TabIndex        =   82
      Top             =   9555
      Width           =   1530
      _ExtentX        =   2699
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
   Begin JeweledBut.JeweledButton BtnTable 
      Height          =   330
      Left            =   1440
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   9555
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
      MICON           =   "FrmSaleInvoicePOSold.frx":121D
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtDiscVal 
      Height          =   315
      Left            =   11505
      TabIndex        =   14
      Top             =   3165
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   8655
      TabIndex        =   10
      Top             =   3165
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   3
      IntegralPoint   =   5
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtPrice 
      Height          =   315
      Left            =   9300
      TabIndex        =   11
      Tag             =   "D"
      Top             =   3165
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   12885
      TabIndex        =   16
      Tag             =   "D"
      Top             =   3165
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Alignment       =   1
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
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2790
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   3165
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
      MICON           =   "FrmSaleInvoicePOSold.frx":1239
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   3135
      TabIndex        =   93
      Tag             =   "D"
      Top             =   3165
      Width           =   3495
      _ExtentX        =   6165
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
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   10260
      TabIndex        =   12
      Top             =   3165
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   10950
      TabIndex        =   13
      Top             =   3165
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtSC 
      Height          =   315
      Left            =   12195
      TabIndex        =   15
      Top             =   3165
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtRemarks 
      Height          =   315
      Left            =   5595
      TabIndex        =   25
      Top             =   9585
      Width           =   6915
      _ExtentX        =   12197
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
      IntegralPoint   =   4
   End
   Begin JeweledBut.JeweledButton BtnSaleOrder 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   4380
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2085
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
      MICON           =   "FrmSaleInvoicePOSold.frx":1255
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrderID 
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   2085
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   9
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
      Mandatory       =   1
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpOrderDate 
      Height          =   315
      Left            =   3030
      TabIndex        =   4
      Top             =   2085
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      Enabled         =   0   'False
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
   Begin SITextBox.Txt TxtEmpComm 
      Height          =   315
      Left            =   7260
      TabIndex        =   120
      Top             =   10680
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   7
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpDeliveryDate 
      Height          =   315
      Left            =   12510
      TabIndex        =   154
      Top             =   7815
      Visible         =   0   'False
      Width           =   1260
      _Version        =   65543
      _ExtentX        =   2222
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
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin MSComCtl2.DTPicker DTPDeliveryTime 
      Height          =   315
      Left            =   13875
      TabIndex        =   155
      Top             =   7800
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "hh:mm tt"
      Format          =   119734275
      UpDown          =   -1  'True
      CurrentDate     =   39224.0416666667
   End
   Begin SITextBox.Txt TxtStampID 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7755
      TabIndex        =   158
      Top             =   465
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
      Alignment       =   1
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
      IntegralPoint   =   4
   End
   Begin SITextBox.Txt TxtBatchNo 
      Height          =   315
      Left            =   8595
      TabIndex        =   160
      Top             =   405
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Alignment       =   1
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
   End
   Begin SITextBox.Txt TxtTokenVal 
      Height          =   315
      Left            =   6540
      TabIndex        =   161
      Top             =   10680
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   3
      IntegralPoint   =   5
   End
   Begin JeweledBut.JeweledButton BtnSaveAS 
      Height          =   420
      Left            =   1305
      TabIndex        =   165
      Top             =   10095
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save As"
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
      MICON           =   "FrmSaleInvoicePOSold.frx":1271
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpPromiseDate 
      Height          =   315
      Left            =   6840
      TabIndex        =   1
      Top             =   1395
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
   Begin SITextBox.Txt TxtSyllabusID 
      Height          =   315
      Left            =   5175
      TabIndex        =   5
      Top             =   2085
      Width           =   705
      _ExtentX        =   1244
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
   Begin SITextBox.Txt TxtSyllabusName 
      Height          =   315
      Left            =   6240
      TabIndex        =   169
      Top             =   2085
      Width           =   2430
      _ExtentX        =   4286
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
   Begin JeweledBut.JeweledButton BtnSyllabus 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5880
      TabIndex        =   170
      TabStop         =   0   'False
      Top             =   2085
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
      MICON           =   "FrmSaleInvoicePOSold.frx":128D
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtPurAmount 
      Height          =   315
      Left            =   10935
      TabIndex        =   173
      Tag             =   "D"
      Top             =   420
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Alignment       =   1
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
   Begin SITextBox.Txt TxtLastPurPrice 
      Height          =   315
      Left            =   9810
      TabIndex        =   175
      Tag             =   "D"
      Top             =   465
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtProdProfit 
      Height          =   315
      Left            =   13140
      TabIndex        =   177
      Tag             =   "D"
      Top             =   375
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Alignment       =   1
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
   Begin SITextBox.Txt TxtTotalProdProfit 
      Height          =   315
      Left            =   13050
      TabIndex        =   179
      Tag             =   "D"
      Top             =   10125
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Alignment       =   1
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
   Begin SITextBox.Txt TxtMemberBarCode 
      Height          =   315
      Left            =   11430
      TabIndex        =   186
      Top             =   1410
      Width           =   1440
      _ExtentX        =   2540
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      CausesValidation=   0   'False
      Height          =   4095
      Left            =   915
      TabIndex        =   94
      Top             =   3480
      Width           =   13650
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   32
      stylesets.count =   6
      stylesets(0).Name=   "Yellow"
      stylesets(0).ForeColor=   65535
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "FrmSaleInvoicePOSold.frx":12A9
      stylesets(1).Name=   "Blue"
      stylesets(1).ForeColor=   16711680
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "FrmSaleInvoicePOSold.frx":12C5
      stylesets(2).Name=   "Red"
      stylesets(2).ForeColor=   665589
      stylesets(2).HasFont=   -1  'True
      BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(2).Picture=   "FrmSaleInvoicePOSold.frx":12E1
      stylesets(3).Name=   "Select"
      stylesets(3).ForeColor=   16777215
      stylesets(3).BackColor=   8388608
      stylesets(3).HasFont=   -1  'True
      BeginProperty stylesets(3).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(3).Picture=   "FrmSaleInvoicePOSold.frx":12FD
      stylesets(4).Name=   "Orange"
      stylesets(4).ForeColor=   33023
      stylesets(4).HasFont=   -1  'True
      BeginProperty stylesets(4).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(4).Picture=   "FrmSaleInvoicePOSold.frx":1319
      stylesets(5).Name=   "Green"
      stylesets(5).ForeColor=   2135858
      stylesets(5).HasFont=   -1  'True
      BeginProperty stylesets(5).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(5).Picture=   "FrmSaleInvoicePOSold.frx":1335
      AllowUpdate     =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   2
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   609
      ExtraHeight     =   609
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   32
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Sr"
      Columns(0).Name =   "Sr"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "Product ID"
      Columns(1).Name =   "ProductID"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3916
      Columns(2).Caption=   "Code"
      Columns(2).Name =   "Code"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   6165
      Columns(3).Caption=   "ProductName"
      Columns(3).Name =   "ProductName"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2117
      Columns(4).Caption=   "Colour"
      Columns(4).Name =   "ColourName"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1455
      Columns(5).Caption=   "Size"
      Columns(5).Name =   "SizeName"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1138
      Columns(6).Caption=   "Qty"
      Columns(6).Name =   "Qty"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   4
      Columns(6).FieldLen=   256
      Columns(7).Width=   1693
      Columns(7).Caption=   "Price"
      Columns(7).Name =   "Price"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   1217
      Columns(8).Caption=   "Disc/Pc"
      Columns(8).Name =   "DiscPC"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   979
      Columns(9).Caption=   "Disc%"
      Columns(9).Name =   "DiscPer"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1217
      Columns(10).Caption=   "Dis. Val"
      Columns(10).Name=   "DiscVal"
      Columns(10).Alignment=   1
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   4
      Columns(10).FieldLen=   256
      Columns(11).Width=   1217
      Columns(11).Caption=   "SC"
      Columns(11).Name=   "SC"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   2461
      Columns(12).Caption=   "Amount"
      Columns(12).Name=   "Amount"
      Columns(12).Alignment=   1
      Columns(12).CaptionAlignment=   2
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   5
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "TotalAmount"
      Columns(13).Name=   "TotalAmount"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   3200
      Columns(14).Visible=   0   'False
      Columns(14).Caption=   "Cost"
      Columns(14).Name=   "Cost"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   4
      Columns(14).FieldLen=   256
      Columns(15).Width=   3200
      Columns(15).Visible=   0   'False
      Columns(15).Caption=   "QtyOrigional"
      Columns(15).Name=   "QtyOrigional"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   4
      Columns(15).FieldLen=   256
      Columns(16).Width=   3200
      Columns(16).Visible=   0   'False
      Columns(16).Caption=   "IsProduct"
      Columns(16).Name=   "IsProduct"
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   11
      Columns(16).FieldLen=   256
      Columns(16).Style=   2
      Columns(17).Width=   3200
      Columns(17).Visible=   0   'False
      Columns(17).Caption=   "EmpComm"
      Columns(17).Name=   "EmpComm"
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      Columns(18).Width=   3200
      Columns(18).Visible=   0   'False
      Columns(18).Caption=   "ExpiryTime"
      Columns(18).Name=   "ExpiryTime"
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(19).Width=   3200
      Columns(19).Visible=   0   'False
      Columns(19).Caption=   "TokenVal"
      Columns(19).Name=   "TokenVal"
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(20).Width=   1429
      Columns(20).Caption=   "EmpID"
      Columns(20).Name=   "EmpID"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      Columns(21).Width=   1720
      Columns(21).Caption=   "EmpName"
      Columns(21).Name=   "EmpName"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   8
      Columns(21).FieldLen=   256
      Columns(22).Width=   1429
      Columns(22).Caption=   "StoreID"
      Columns(22).Name=   "StoreID"
      Columns(22).DataField=   "Column 22"
      Columns(22).DataType=   8
      Columns(22).FieldLen=   256
      Columns(23).Width=   1852
      Columns(23).Caption=   "StoreName"
      Columns(23).Name=   "StoreName"
      Columns(23).DataField=   "Column 23"
      Columns(23).DataType=   8
      Columns(23).FieldLen=   256
      Columns(24).Width=   3200
      Columns(24).Visible=   0   'False
      Columns(24).Caption=   "PurAmount"
      Columns(24).Name=   "PurAmount"
      Columns(24).DataField=   "Column 24"
      Columns(24).DataType=   8
      Columns(24).FieldLen=   256
      Columns(25).Width=   3200
      Columns(25).Visible=   0   'False
      Columns(25).Caption=   "ProdProfit"
      Columns(25).Name=   "ProdProfit"
      Columns(25).DataField=   "Column 25"
      Columns(25).DataType=   8
      Columns(25).FieldLen=   256
      Columns(26).Width=   3200
      Columns(26).Visible=   0   'False
      Columns(26).Caption=   "LastPurPrice"
      Columns(26).Name=   "LastPurPrice"
      Columns(26).DataField=   "Column 26"
      Columns(26).DataType=   8
      Columns(26).FieldLen=   256
      Columns(27).Width=   3200
      Columns(27).Visible=   0   'False
      Columns(27).Caption=   "ColourID"
      Columns(27).Name=   "ColourID"
      Columns(27).DataField=   "Column 27"
      Columns(27).DataType=   8
      Columns(27).FieldLen=   256
      Columns(28).Width=   3200
      Columns(28).Visible=   0   'False
      Columns(28).Caption=   "SizeID"
      Columns(28).Name=   "SizeID"
      Columns(28).DataField=   "Column 28"
      Columns(28).DataType=   8
      Columns(28).FieldLen=   256
      Columns(29).Width=   3200
      Columns(29).Caption=   "DiscAmount"
      Columns(29).Name=   "DiscAmount"
      Columns(29).DataField=   "Column 29"
      Columns(29).DataType=   8
      Columns(29).FieldLen=   256
      Columns(30).Width=   3200
      Columns(30).Caption=   "SaletaxVal"
      Columns(30).Name=   "SaletaxVal"
      Columns(30).DataField=   "Column 30"
      Columns(30).DataType=   8
      Columns(30).FieldLen=   256
      Columns(31).Width=   3200
      Columns(31).Caption=   "SaletaxPer"
      Columns(31).Name=   "SaletaxPer"
      Columns(31).DataField=   "Column 31"
      Columns(31).DataType=   8
      Columns(31).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   24077
      _ExtentY        =   7223
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtSID 
      Height          =   315
      Left            =   1080
      TabIndex        =   191
      Top             =   1395
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   9
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtExtraTaxVal 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   12840
      TabIndex        =   193
      Top             =   9060
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtExtraTaxPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   14040
      TabIndex        =   194
      Top             =   9060
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtAdvTaxVal 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   12840
      TabIndex        =   195
      Top             =   8505
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtAdvTaxPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   14040
      TabIndex        =   196
      Top             =   8505
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtSumDiscAmount 
      Height          =   315
      Left            =   3960
      TabIndex        =   206
      Top             =   7680
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
      Masked          =   2
   End
   Begin SITextBox.Txt TxtDiscAmount 
      Height          =   315
      Left            =   120
      TabIndex        =   207
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   1
   End
   Begin SITextBox.Txt TxtAvgDisc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   211
      Top             =   7920
      Visible         =   0   'False
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   915
      TabIndex        =   8
      Top             =   3165
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Appearance      =   0
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
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtSaleTaxPer 
      Height          =   315
      Left            =   270
      TabIndex        =   227
      Top             =   825
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   3
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtSaleTaxValue 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   810
      TabIndex        =   228
      Top             =   810
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtTotalSaleTaxValue 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1710
      TabIndex        =   229
      Top             =   810
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   5
   End
   Begin VB.Label LblSaleTaxPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax%"
      Height          =   195
      Left            =   270
      TabIndex        =   232
      Top             =   630
      Width           =   390
   End
   Begin VB.Label Label60 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Value"
      Height          =   195
      Left            =   855
      TabIndex        =   231
      Top             =   630
      Width           =   720
   End
   Begin VB.Label Label61 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Tax Value"
      Height          =   195
      Left            =   1800
      TabIndex        =   230
      Top             =   630
      Width           =   1125
   End
   Begin VB.Label LblAllStock 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Store Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   7245
      TabIndex        =   226
      Top             =   2655
      Width           =   1905
   End
   Begin VB.Label LblMultiplier 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Multiplier"
      Height          =   195
      Left            =   9720
      TabIndex        =   218
      Top             =   2790
      Width           =   615
   End
   Begin VB.Label Label55 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Avg (%)"
      Height          =   195
      Left            =   90
      TabIndex        =   212
      Top             =   7680
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Amount"
      Height          =   195
      Left            =   120
      TabIndex        =   209
      Top             =   2400
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label53 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sum of Disc Amount"
      Height          =   195
      Left            =   2400
      TabIndex        =   208
      Top             =   7800
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   14040
      TabIndex        =   200
      Top             =   8280
      Width           =   210
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Withholding Tax"
      Height          =   195
      Left            =   12840
      TabIndex        =   199
      Top             =   8280
      Width           =   1155
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Sale Tax"
      Height          =   195
      Left            =   12840
      TabIndex        =   198
      Top             =   8835
      Width           =   1035
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   14040
      TabIndex        =   197
      Top             =   8835
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "SID"
      Height          =   195
      Left            =   1080
      TabIndex        =   192
      Top             =   1170
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label LblMemberBarCode 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member BarCode"
      Height          =   195
      Left            =   11430
      TabIndex        =   187
      Top             =   1185
      Width           =   1230
   End
   Begin VB.Label LblMemberName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      Height          =   195
      Left            =   10050
      TabIndex        =   185
      Top             =   1185
      Width           =   1035
   End
   Begin VB.Label LblSize 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   7815
      TabIndex        =   184
      Top             =   2970
      Width           =   300
   End
   Begin VB.Label LblColour 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
      Height          =   195
      Left            =   6615
      TabIndex        =   183
      Top             =   2970
      Width           =   450
   End
   Begin VB.Label LblTotalProdProfit 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "PPxxxxx"
      Height          =   195
      Left            =   13065
      TabIndex        =   180
      Top             =   9930
      Width           =   585
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Prod Profit"
      Height          =   195
      Left            =   13155
      TabIndex        =   178
      Top             =   180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Price"
      Height          =   195
      Left            =   9810
      TabIndex        =   176
      Top             =   270
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label49 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Amount"
      Height          =   195
      Left            =   10950
      TabIndex        =   174
      Top             =   225
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblSyllabusName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Syllabus Name"
      Height          =   195
      Left            =   6240
      TabIndex        =   172
      Top             =   1890
      Width           =   1050
   End
   Begin VB.Label LblSyllabusID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Syllabus ID"
      Height          =   195
      Left            =   5175
      TabIndex        =   171
      Top             =   1890
      Width           =   795
   End
   Begin VB.Label LblLastBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Bill Nos."
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
      Left            =   9675
      TabIndex        =   168
      Top             =   945
      Width           =   1140
   End
   Begin VB.Label LblPromiseDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Promise Date"
      Height          =   195
      Left            =   6885
      TabIndex        =   167
      Top             =   1185
      Width           =   945
   End
   Begin VB.Label LblLastPurPrice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   11130
      TabIndex        =   166
      Top             =   10680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   2400
      TabIndex        =   164
      Top             =   8070
      Width           =   1365
   End
   Begin VB.Label TxtTotalItems 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   915
      Left            =   2400
      TabIndex        =   163
      Top             =   8355
      Width           =   1380
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Token Val"
      Height          =   195
      Left            =   6540
      TabIndex        =   162
      Top             =   10500
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label47 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Stamp ID"
      Height          =   195
      Left            =   7080
      TabIndex        =   159
      Top             =   555
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
      Height          =   195
      Left            =   12510
      TabIndex        =   157
      Top             =   7635
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Time"
      Height          =   195
      Left            =   13830
      TabIndex        =   156
      Top             =   7620
      Visible         =   0   'False
      Width           =   960
   End
   Begin MSForms.TextBox TxtRemarksUrdu 
      Height          =   435
      Left            =   5595
      TabIndex        =   26
      ToolTipText     =   "Textbox1"
      Top             =   9510
      Visible         =   0   'False
      Width           =   6945
      VariousPropertyBits=   752896027
      ForeColor       =   0
      MaxLength       =   100
      BorderStyle     =   1
      Size            =   "12250;767"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label LblType 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Left            =   3525
      TabIndex        =   153
      Top             =   9330
      Width           =   360
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "EmpComm"
      Height          =   195
      Left            =   7260
      TabIndex        =   121
      Top             =   10455
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ord ID"
      Height          =   195
      Left            =   2340
      TabIndex        =   119
      Top             =   1860
      Width           =   465
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
      Height          =   195
      Left            =   3030
      TabIndex        =   118
      Top             =   1860
      Width           =   780
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Order"
      Height          =   195
      Left            =   4380
      TabIndex        =   117
      Top             =   1860
      Width           =   390
   End
   Begin VB.Label LblRemarks 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   5595
      TabIndex        =   116
      Top             =   9285
      Width           =   630
   End
   Begin VB.Label LblSC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "S.C."
      Height          =   195
      Left            =   12225
      TabIndex        =   115
      Top             =   2970
      Width           =   300
   End
   Begin VB.Label LblCost 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   1395
      TabIndex        =   114
      Top             =   2340
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. %"
      Height          =   195
      Left            =   10935
      TabIndex        =   103
      Top             =   2970
      Width           =   525
   End
   Begin VB.Label LblProdPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   9300
      TabIndex        =   101
      Top             =   2970
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3135
      TabIndex        =   100
      Top             =   2970
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   915
      TabIndex        =   99
      Top             =   2970
      Width           =   375
   End
   Begin VB.Label LblDiscPC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc / PC"
      Height          =   195
      Left            =   10200
      TabIndex        =   98
      Top             =   2970
      Width           =   690
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   8655
      TabIndex        =   97
      Top             =   2970
      Width           =   240
   End
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   12900
      TabIndex        =   96
      Top             =   2970
      Width           =   540
   End
   Begin VB.Label LblDiscVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. Val"
      Height          =   195
      Left            =   11505
      TabIndex        =   95
      Top             =   2970
      Width           =   630
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Disc."
      Height          =   195
      Left            =   915
      TabIndex        =   91
      Top             =   7620
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   1755
      TabIndex        =   90
      Top             =   7620
      Width           =   210
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   1755
      TabIndex        =   89
      Top             =   8745
      Width           =   210
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Ch."
      Height          =   195
      Left            =   915
      TabIndex        =   88
      Top             =   8745
      Width           =   825
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
      Height          =   195
      Left            =   915
      TabIndex        =   87
      Top             =   8190
      Width           =   705
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   1755
      TabIndex        =   86
      Top             =   8190
      Width           =   210
   End
   Begin VB.Label LblTableID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Table ID"
      Height          =   195
      Left            =   915
      TabIndex        =   85
      Top             =   9375
      Width           =   615
   End
   Begin VB.Label LblTableName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name"
      Height          =   195
      Left            =   1755
      TabIndex        =   84
      Top             =   9375
      Width           =   870
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   8610
      TabIndex        =   81
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   9795
      TabIndex        =   80
      Top             =   1860
      Width           =   1350
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill No"
      Height          =   195
      Left            =   10995
      TabIndex        =   77
      Top             =   10005
      Width           =   1020
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
      Height          =   225
      Left            =   750
      TabIndex        =   76
      Top             =   10455
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label LblNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   810
      TabIndex        =   74
      Top             =   10125
      Width           =   165
   End
   Begin VB.Label LblMemberID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      Height          =   195
      Left            =   8250
      TabIndex        =   73
      Top             =   1185
      Width           =   780
   End
   Begin VB.Label LblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12975
      TabIndex        =   70
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label LblCaptionPrice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   11850
      TabIndex        =   66
      Top             =   600
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label LblPrice 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   12075
      TabIndex        =   65
      Top             =   915
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      Height          =   195
      Left            =   8025
      TabIndex        =   64
      Top             =   10455
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   12705
      TabIndex        =   62
      Top             =   1860
      Width           =   780
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   11595
      TabIndex        =   61
      Top             =   1860
      Width           =   525
   End
   Begin VB.Label TxtTotalQty 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   915
      Left            =   3855
      TabIndex        =   58
      Top             =   8340
      Width           =   1380
   End
   Begin VB.Label TxtTotalAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   915
      Left            =   5325
      TabIndex        =   57
      Top             =   8340
      Width           =   2550
   End
   Begin VB.Label TxtTotalDiscount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   7965
      TabIndex        =   56
      Top             =   8340
      Width           =   2010
   End
   Begin VB.Label TxtNetAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   10020
      TabIndex        =   55
      Top             =   8340
      Width           =   2550
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Invoice (POS)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   53
      Top             =   270
      Width           =   3180
   End
   Begin VB.Label LblStockCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   4140
      TabIndex        =   52
      Top             =   2655
      Width           =   720
   End
   Begin VB.Label LblStock 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   5040
      TabIndex        =   51
      Top             =   2655
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   8835
      TabIndex        =   50
      Top             =   10455
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   10140
      TabIndex        =   48
      Top             =   10500
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   4380
      TabIndex        =   46
      Top             =   1185
      Width           =   585
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   5415
      TabIndex        =   45
      Top             =   1185
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Amount"
      Height          =   195
      Left            =   9195
      TabIndex        =   42
      Top             =   10485
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Discount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   7965
      TabIndex        =   40
      Top             =   8040
      Width           =   1755
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   5325
      TabIndex        =   39
      Top             =   8040
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3855
      TabIndex        =   38
      Top             =   8055
      Width           =   1020
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   13290
      Top             =   1065
      Width           =   345
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   10020
      TabIndex        =   37
      Top             =   8040
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   3030
      TabIndex        =   36
      Top             =   1185
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " Bill ID"
      Height          =   195
      Left            =   2340
      TabIndex        =   35
      Top             =   1185
      Width           =   450
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
      Begin VB.Menu MniCostPrice 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmSaleInvoicePOSold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public cnSalePOS As New ADODB.Connection
Dim Application1 As New CRAXDRT.Application
Dim vMode As FormMode
Dim vDisplay As String, vPrice As String
Dim vPOSID As String, vFBRInvoiceNo As String
Dim vCounter As Integer, vChange As Boolean
Dim vDate, vNow, vServerDate As Date, vHDiff As Integer, vSystemDate As Boolean
Dim RsDetail As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset
Dim RsBodyStore As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vMaxBinID As Integer
Dim vIsNewRecord As Boolean
Dim Flag As Boolean, vNegativeSale As Boolean
Dim vFlag As Boolean
Dim vBm As Variant, vExpiryTime As String
Dim UniCode As Variant
Dim DateFlag As Boolean, DiscPerFlag As Boolean
Dim vProducts As String, vHeader As String
Dim ParaOutPrevious As Double
Public ParaOutSelection As Boolean
Public ParaInChoice As String
Public objFSO As New Scripting.FileSystemObject
'Dim vSystemDate As Date
Dim sSQL, vRemarks, vDescription As String
Dim vStrSQL As String, vAutoEnterBeforeQty As Boolean, vEmptyEnterGotoSave As Boolean
Dim vPrevious As Double, vCurrent As Double
Dim vQtyLoose As Double, vTotalAmount As Double
Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String, vTotDisc As Double
Dim i As Integer, vCustomerPoleDisplay As Boolean, vLaserInvoice As Boolean, vPrintHeader  As Boolean, vNoofPrints As Byte, vX As Integer, vY As Integer, vBlankFooter As Integer
Dim vCash, vCredit As Integer
Dim vStrPara As String
Dim vMasterID As Long
Dim vContactNo As String
Dim vBarcode As String
Dim vStrDetail As String
Dim vMobileNo() As String, vMobile As String
Dim vUnitPrice, vUnitRetailPrice, vAmount As Double
Dim vIsChangedPrice As Boolean
Dim vColour, vEmiNo As Boolean, vAlreadySerial, vIsRemarksCompulsory As Boolean
Dim vIsAdministrator, vIsEdit, vOrganizationMandatory, vEmployeeMandatory, visEntryDate As Boolean
Dim vCurrentDateDataEntry, vNotEditingAfterPrinting, vChangeQtyOnChangedPrice, vIsDisableCreditSale, vEmployeeCommision As Boolean
Dim vIsCreditSale, vAutoPrintinInvoices, vUpdateStock, vLaserPrintofSaleInvoice As Boolean
Dim vPrintHeadersSaleInvoice, vPreviousBalanceVisible, vPrintCondition1, vPrintCondition2, vPrintCondition3, vPrintCondition4   As Boolean
Dim vCompanyAddress, vCompanyCity, vCompanyPhoneNo, vAddSpace, vStatement, vExpiryColor   As String
Dim vCashReceived, vBottomPrice, vNetAmount As Double

Private Sub BtnAddCustomer_Click()
    DefCustomers.Show
End Sub

Private Sub Form_Activate()
   If Trim(ParaCustID) <> "" Then
       TxtCustomerID.Text = ParaCustID
       TxtCustomerName.Text = ParaCustName
       If FunSelectCustomer(1, False) = True Then
       End If
       ParaCustID = ""
       ParaCustName = ""
    End If
'    If TxtCode.Text = "" And TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
End Sub

Private Sub Form_GotFocus()
On Error GoTo ErrorHandler
'   If cnSalePOS.State = adStateClosed Then cnSalePOS.Open
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name = TxtRemarksUrdu.Name Then
      Call Textbox1_KeyDown(KeyCode, Shift)
      Exit Sub
   End If
   If KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      Call SubEnable(True)
      Frame1.Visible = False
      If TxtCode.Enabled Then TxtCode.SetFocus
      Select Case ActiveControl.Name
         Case TxtCode.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name, TxtSC.Name
         If TxtCode.Enabled Then TxtCode.SetFocus: Call SubClearDetailArea
      End Select
   ElseIf Shift = vbCtrlMask Then
      If ActiveControl.Name = Grid.Name Then
         If KeyCode = vbKeyDelete Then
            If Trim(Grid.Columns("ProductID").Text <> "") Then Call mniRemoveRow_Click
            KeyCode = 0
         Else
            KeyCode = 0: Exit Sub
         End If
      End If
      Select Case KeyCode
         Case vbKeyS
            If BtnOk.Enabled And BtnOk.Visible Then BtnOk_Click: KeyCode = 0: Exit Sub
            If BtnSave.Enabled And BtnSave.Visible Then BtnSave_Click
            KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled And BtnClear.Visible Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled And BtnClose.Visible Then BtnClose_Click
            KeyCode = 0
         Case vbKeyU
            Call SubMakePackageDeal
         Case vbKeyE
               If TxtEmployeeID.Visible = True And TxtEmployeeID.Enabled = True Then TxtEmployeeID.SetFocus
               KeyCode = 0
'         Case vbKeyY
'               If lblComPort.Visible = False Then lblComPort.Visible = True
'               If TxtComPort.Visible = False Then TxtComPort.Visible = True
'               If vCustomerPoleDisplay = True Then
'                  If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
'                  MSComm1.CommPort = Val(TxtComPort.Text)                     'Use com1 port
'                  MSComm1.Settings = "9600,N,8,1"                             'Port Settings
'                  If MSComm1.PortOpen = False Then MSComm1.PortOpen = True    'open port
'               End If
'               KeyCode = 0
         Case vbKeyT
               If TxtStoreID.Visible = True And TxtStoreID.Enabled = True Then TxtStoreID.SetFocus
               KeyCode = 0
         Case vbKeyM
               If TxtMemberID.Visible = True And TxtMemberID.Enabled = True Then TxtMemberID.SetFocus
               KeyCode = 0
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
               KeyCode = 0
         Case vbKeyO
            If BtnOpen.Enabled And BtnOpen.Visible Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled And BtnDelete.Visible Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled And BtnPrint.Visible Then BtnPrint_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyC And Shift = vbAltMask Then
      ParaInChoice = "Credit"
      FrmPrint.Show vbModal, Me
   ElseIf KeyCode = vbKeyReturn And Shift = vbShiftMask Then
      Select Case ActiveControl.Name
      Case TxtCode.Name
         If FunSelectProduct(ssValidate, False) = True Then TxtQty.SetFocus
      Case TxtQty.Name
         If TxtPrice.Visible = False Then TxtDiscPC.SetFocus Else TxtPrice.SetFocus
      Case TxtPrice.Name
         TxtDiscPC.SetFocus
      Case TxtDiscPC.Name
         TxtDiscPer.SetFocus
      Case TxtDiscPer.Name
         TxtDiscVal.SetFocus
      Case TxtDiscVal.Name
         TxtSC.SetFocus
      End Select
      KeyCode = 0
      Shift = 0
   ElseIf KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name
         Grid_DblClick
      Case TxtCode.Name
         If vEmptyEnterGotoSave = True Then If Trim(TxtCode.Text) = "" Then If BtnSave.Enabled And BtnSave.Visible Then BtnSave.SetFocus
         If FunSelectProduct(ssValidate, False) = True Then
            If vAutoEnterBeforeQty = True And vIsChangedPrice = False Then GetDataFromTexBoxesToGrid Else keybd_event 9, 1, 1, 1: KeyCode = 0
         End If
      Case TxtProductName.Name
            Call FindRow
      Case TxtQty.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name, TxtPrice.Name, TxtSC.Name, TxtAmount.Name
         GetDataFromTexBoxesToGrid
      Case Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtMemberID.Visible And TxtMemberID.Enabled Then TxtMemberID.SetFocus Else TxtStoreID.SetFocus
         Case TxtTableID.Name: If FunSelectTable(ssFunctionKey, False) = True Then TxtTableID.SetFocus
         Case TxtMemberID.Name: If FunSelectMember(ssFunctionKey, True) = True Then If TxtEmployeeID.Visible And TxtEmployeeID.Enabled Then TxtEmployeeID.SetFocus Else TxtMemberID.SetFocus
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus Else TxtEmployeeID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then If TxtEmployeeID.Visible And TxtEmployeeID.Enabled Then TxtEmployeeID.SetFocus Else TxtOrganizationID.SetFocus
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then If TxtQty.Enabled And TxtQty.Visible Then TxtQty.SetFocus Else TxtCode.SetFocus
         Case TxtBankMachineID.Name: If FunSelectBankMachine(ssFunctionKey, True) = True Then BtnOk.SetFocus
         Case TxtBankMachineCreditID.Name: If FunSelectBankMachineCredit(ssFunctionKey, True) = True Then TxtBankAmount.SetFocus
         Case TxtCustomerID.Name: If FunSelectCustomer(ssFunctionKey, False) = True Then TxtBankAmount.SetFocus
      End Select
   ElseIf KeyCode = vbKeyF2 Then
         If Frame3.Visible = True Then
            Frame3.Visible = False
            If TxtCode.Enabled = True Then TxtCode.SetFocus Else Grid.SetFocus
        Else
            Frame3.Visible = True
            Frame3.ZOrder 0
            KeyCode = 0
            If TxtSerial.Enabled = True And TxtSerial.Visible = True Then TxtSerial.SetFocus
        End If
   ElseIf KeyCode = vbKeyF3 Then
      TxtProductName.Enabled = True
      If TxtProductName.Enabled = True And TxtProductName.Visible = True Then TxtProductName.SetFocus
      'Call FindRow
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         If Grid.Visible And Grid.Enabled Then Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
         KeyCode = 0
         TxtBillDisc.SetFocus
      End If
   ''' Show Product Purchase Price
   ElseIf KeyCode = vbKeyF4 Then
      If TxtPID.Text <> "" And (ObjUserSecurity.ShowPrice = True Or ObjUserSecurity.IsAdministrator = True) Then
         Select Case ActiveControl.Name
         Case TxtCode.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name, TxtSC.Name, Grid.Name
            LblCost.Caption = cn.Execute("select PurPrice from products where productid='" & TxtPID.Text & "'").Fields(0).Value
            Call MniCostPrice_Click
'            LblCost.Visible = True
         End Select
      End If
   ''' Show Last Purchase Price Price
   ElseIf KeyCode = vbKeyF5 And (ObjUserSecurity.LastPurchasePrice = True Or ObjUserSecurity.IsAdministrator = True) Then
      If TxtPID.Text <> "" Then
         Select Case ActiveControl.Name
         Case TxtCode.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name, TxtSC.Name, Grid.Name
            LblCost.Caption = cn.Execute("select dbo.FunPurPrice('" & TxtPID.Text & "')").Fields(0).Value
            Call MniCostPrice_Click
'            LblCost.Visible = True
         End Select
      End If
    ''' Show Weighted Price
    ElseIf KeyCode = vbKeyF6 And (ObjUserSecurity.WeightedPrice = True Or ObjUserSecurity.IsAdministrator = True) Then
      If TxtPID.Text <> "" Then
         Select Case ActiveControl.Name
         Case TxtCode.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name, TxtSC.Name, Grid.Name
            cn.Execute "exec SPProductAverageCost '" & DtpBillDate.DateValue & "','" & TxtPID.Text & "'"
            LblCost.Caption = cn.Execute("Select Price from TempPurchase Where Productid = '" & TxtPID.Text & "'").Fields(0).Value
            Call MniCostPrice_Click
'            LblCost.Visible = True
         End Select
      End If
   ''' Show WS Price
    ElseIf KeyCode = vbKeyF7 And (ObjUserSecurity.WSPrice = True Or ObjUserSecurity.IsAdministrator = True) Then
      If TxtPID.Text <> "" Then
         Select Case ActiveControl.Name
         Case TxtCode.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name, TxtSC.Name, Grid.Name
'            cn.Execute "exec SPProductAverageCost '" & DtpBillDate.DateValue & "','" & TxtPID.Text & "'"
            LblCost.Caption = cn.Execute("Select WSPrice from Products Where Productid = '" & TxtPID.Text & "'").Fields(0).Value
            Call MniCostPrice_Click
'            LblCost.Visible = True
         End Select
      End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrorHandler
   If KeyAscii = vbKeyReturn Then Exit Sub
   If Me.ActiveControl.Name = TxtRemarksUrdu.Name Then
      Call Textbox1_KeyPress(KeyAscii)
      Exit Sub
   End If
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Or vChange = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF7 Or KeyCode = vbKeyF6 Or KeyCode = vbKeyF5 Or KeyCode = vbKeyF4 Then
      LblCost.Visible = False
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   
'   Set cnSalePOS = CN
'   If cnSalePOS.State = adStateOpen Then cnSalePOS.Close
'
'   cnSalePOS.Open
'   cnSalePOS.CursorLocation = adUseClient

   If objFSO.FileExists(vTmp & "\Settings.ini") Then
      Open vTmp & "\Settings.ini" For Input As #1
      Line Input #1, vPOSID
      Close #1
   Else
      vPOSID = ""
   End If
   
   
ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   vFlag = True
   Call InvoiceNo
   SetWindowText Me.hWnd, "Sale Invoice (" & LblNo & ")"
   HelpLocation Me
   
   Grid.RowHeight = ObjRegistry.GridRowHeight
   
   LblCost.Left = ObjRegistry.CostX
   LblCost.Top = ObjRegistry.CostY
   vIsAdministrator = ObjUserSecurity.IsAdministrator
   vIsEdit = ObjUserSecurity.IsEdit
   vOrganizationMandatory = ObjRegistry.OrganizationMandatory
   vEmployeeMandatory = ObjRegistry.EmployeeMandatory
   visEntryDate = ObjRegistry.isEntryDate
   vCurrentDateDataEntry = ObjRegistry.CurrentDateDataEntry
   vNotEditingAfterPrinting = ObjUserSecurity.NotEditingAfterPrinting
   vChangeQtyOnChangedPrice = ObjRegistry.ChangeQtyOnChangedPrice
   vIsDisableCreditSale = ObjUserSecurity.IsDisableCreditSale
   vIsCreditSale = ObjUserSecurity.IsCreditSale
   vEmployeeCommision = ObjRegistry.EmployeeCommision
   vAutoPrintinInvoices = ObjRegistry.AutoPrintinInvoices
   vUpdateStock = ObjRegistry.UpdateStockSaleBodyInsert
   vLaserPrintofSaleInvoice = ObjRegistry.LaserPrintofSaleInvoice
   vPrintHeadersSaleInvoice = ObjRegistry.PrintHeadersSaleInvoice
   vCompanyName = ObjRegistry.CompanyName
   vCompanyAddress = ObjRegistry.CompanyAddress
   vCompanyCity = IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   vCompanyPhoneNo = IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   vAddSpace = IIf(ObjRegistry.AddSpace = True, Left(".......................................", Val(ObjRegistry.BlankFooter)), "")
   vCashReceived = ObjRegistry.CashReceived
   vStatement = ObjRegistry.Statement
   vPreviousBalanceVisible = IIf(ObjRegistry.PreviousBalanceVisible = True, ParaOutPrevious, 0)
   
   
   vPrintCondition1 = False
   vPrintCondition2 = False
   vPrintCondition3 = False
   vPrintCondition4 = False
   
   
'   TxtBillDisc.Enabled = vIsAdministrator
   
   
   If InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "CANON") > 0 Or InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "HP") > 0 Then
      vPrintCondition1 = True
   End If
   
   If InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "CBM1000") > 0 Then
      vPrintCondition2 = True
   End If
   
   If InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "AB-80K") > 0 Or InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "ARP-808K") > 0 Then
      vPrintCondition3 = True
   End If
   
   If (InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0) Then
      vPrintCondition4 = True
   End If
   
   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
      
   vEmiNo = False
   vColour = ObjRegistry.ShowColourSize
   
   LblColour.Visible = vColour
   CmbColourName.Visible = vColour
   LblSize.Visible = vColour
   cmbSizeName.Visible = vColour
   Grid.Columns("ColourName").Visible = vColour
   Grid.Columns("SizeName").Visible = vColour
   
   If vColour = False Then
      LblQty.Left = LblQty.Left - CmbColourName.Width - cmbSizeName.Width
      TxtQty.Left = TxtQty.Left - CmbColourName.Width - cmbSizeName.Width
      LblProdPrice.Left = LblProdPrice.Left - CmbColourName.Width - cmbSizeName.Width
      TxtPrice.Left = TxtPrice.Left - CmbColourName.Width - cmbSizeName.Width
      LblDiscPC.Left = LblDiscPC.Left - CmbColourName.Width - cmbSizeName.Width
      TxtDiscPC.Left = TxtDiscPC.Left - CmbColourName.Width - cmbSizeName.Width
      LblDiscPer.Left = LblDiscPer.Left - CmbColourName.Width - cmbSizeName.Width
      TxtDiscPer.Left = TxtDiscPer.Left - CmbColourName.Width - cmbSizeName.Width
      LblDiscVal.Left = LblDiscVal.Left - CmbColourName.Width - cmbSizeName.Width
      TxtDiscVal.Left = TxtDiscVal.Left - CmbColourName.Width - cmbSizeName.Width
      LblSC.Left = LblSC.Left - CmbColourName.Width - cmbSizeName.Width
      TxtSC.Left = TxtSC.Left - CmbColourName.Width - cmbSizeName.Width
      LblAmount.Left = LblAmount.Left - CmbColourName.Width - cmbSizeName.Width
      TxtAmount.Left = TxtAmount.Left - CmbColourName.Width - cmbSizeName.Width
      Grid.Width = Grid.Width - CmbColourName.Width - cmbSizeName.Width
   End If
   vServerDate = cn.Execute("Select CONVERT(datetime, CONVERT(varchar, GETDATE(), 110)) ServerDate").Fields(0).Value
   vSystemDate = Abs(ObjRegistry.SystemDate)
   vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)

   TxtStoreID.Text = IIf((ObjRegistry.StoreID = ""), "", ObjRegistry.StoreID)
   FunSelectStore ssValidate, True
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible

   LblEmpID.Visible = ObjRegistry.EmpVisible
   LblEmpName.Visible = ObjRegistry.EmpVisible
   TxtEmployeeID.Visible = ObjRegistry.EmpVisible
   TxtEmployeeName.Visible = ObjRegistry.EmpVisible
   BtnEmployee.Visible = ObjRegistry.EmpVisible
   
   LblMemberID.Visible = ObjRegistry.MemberVisible
   LblMemberName.Visible = ObjRegistry.MemberVisible
   TxtMemberID.Visible = ObjRegistry.MemberVisible
   TxtMemberName.Visible = ObjRegistry.MemberVisible
   BtnMember.Visible = ObjRegistry.MemberVisible
   LblMemberBarCode.Visible = ObjRegistry.MemberVisible
   TxtMemberBarCode.Visible = ObjRegistry.MemberVisible
         
   LblTableID.Visible = ObjRegistry.TableVisible
   LblTableName.Visible = ObjRegistry.TableVisible
   TxtTableID.Visible = ObjRegistry.TableVisible
   TxtTableName.Visible = ObjRegistry.TableVisible
   BtnTable.Visible = ObjRegistry.TableVisible
      
   TxtManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   LblManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   
   
   TxtRemarks.Visible = ObjRegistry.RemarksVisible
   LblRemarks.Visible = ObjRegistry.RemarksVisible
   
   LblSyllabusID.Visible = ObjRegistry.ShowSyllabus
   LblSyllabusName.Visible = ObjRegistry.ShowSyllabus
   TxtSyllabusID.Visible = ObjRegistry.ShowSyllabus
   TxtSyllabusName.Visible = ObjRegistry.ShowSyllabus
   BtnSyllabus.Visible = ObjRegistry.ShowSyllabus
   
   LblTotalProdProfit.Visible = ObjRegistry.ShowProdProfit
   TxtTotalProdProfit.Visible = ObjRegistry.ShowProdProfit
   
   If LblRemarks.Visible = True Then
      TxtRemarks.Visible = Not ObjRegistry.AllowUrduProduct
      TxtRemarksUrdu.Visible = ObjRegistry.AllowUrduProduct
   End If
   
   If ObjRegistry.ShowPromiseDateInSalaPurchase = True Then
      LblPromiseDate.Visible = True
      DtpPromiseDate.Visible = True
      DtpPromiseDate.DateValue = Null
   Else
      LblPromiseDate.Visible = False
      DtpPromiseDate.Visible = False
      DtpPromiseDate.DateValue = Null
   End If
   If ObjUserSecurity.IsAdministrator = False Then
      TxtDiscPC.Enabled = ObjRegistry.DiscAllowed
      TxtDiscPer.Enabled = ObjRegistry.DiscAllowed
      TxtDiscVal.Enabled = ObjRegistry.DiscAllowed
      TxtBillDisc.Enabled = ObjRegistry.DiscAllowed
      TxtBillDiscPer.Enabled = ObjRegistry.DiscAllowed
      TxtSTax.Enabled = ObjRegistry.DiscAllowed
      TxtSTaxPer.Enabled = ObjRegistry.DiscAllowed
      TxtServiceCharges.Enabled = ObjRegistry.DiscAllowed
      TxtServiceChargesPer.Enabled = ObjRegistry.DiscAllowed
      If ObjRegistry.DiscAllowed = False Then
         TxtDiscPC.Tag = "NC"
         TxtDiscPer.Tag = "NC"
         TxtDiscVal.Tag = "NC"
         TxtBillDisc.Tag = "NC"
         TxtBillDiscPer.Tag = "NC"
      End If
   End If
   
   LblType.Visible = ObjRegistry.InvType
   CmbType.Visible = ObjRegistry.InvType
   
   CmbType.Clear
   CmbType.AddItem ""
   With cn.Execute("select * from InvTypes")
      If .RecordCount > 0 Then
         While Not .EOF
            CmbType.AddItem ![InvType]
            .MoveNext
         Wend
      End If
   End With
   Frame1.Visible = False
   vNegativeSale = ObjRegistry.NegativeSale
   vAutoEnterBeforeQty = ObjRegistry.AutoEnterBeforeQty
   vEmptyEnterGotoSave = ObjRegistry.EmptyEnterGotoSave
   vX = IIf(IsNull(ObjRegistry.x), 0, Val(ObjRegistry.x))
   vY = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))
   vLaserInvoice = ObjRegistry.LaserPrintofSaleInvoice
   vPrintHeader = ObjRegistry.PrintHeadersSaleInvoice
   vNoofPrints = IIf(IsNull(ObjRegistry.NoofPrints) Or Val(ObjRegistry.NoofPrints) = 0, 1, ObjRegistry.NoofPrints)
   MniCostPrice.Visible = ObjRegistry.CostVisible
   If ObjUserSecurity.IsAdministrator = True Then
      TxtPrice.Enabled = True
      TxtPrice.Tag = ""
   Else
      TxtPrice.Enabled = ObjUserSecurity.IsChangeRetail
      TxtPrice.Tag = IIf(TxtPrice.Enabled = True, "", "D")
   End If
   DateFlag = True
   vCustomerPoleDisplay = False
   With cn.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
         If ObjRegistry.ChangePrice = True Then TxtPrice.Enabled = True
         vCustomerPoleDisplay = IIf(IsNull(!CustomerPoleDisplay), False, !CustomerPoleDisplay)
         If vCustomerPoleDisplay = True Then
            MSComm1.CommPort = !CommPort
            If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
            MSComm1.Settings = "9600,N,8,1"
            Timer2.Enabled = True
         End If
      End If
      .Close
   End With
   
   ChkPrint.Enabled = Not ObjRegistry.HideAutoPrint
   ChkPrint.Value = Abs(ObjRegistry.AutoPrintinInvoices)
   ChkPrint.Tag = IIf(ChkPrint.Enabled = True, "F", "NC")
   
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus

   DateFlag = True
   
   vCompanyName = ObjRegistry.CompanyName
   If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
'   Frame3.Visible = False
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Are you sure to close without save?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then
         Cancel = 1
         
      End If
   Else
   cn.Execute "delete from tempno where tempno = " & Val(Right(LblNo.Caption, 1))
    
    'cn.Execute ("exec spcurrentstock")
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set RsBody = Nothing
    Set RsDetail = Nothing
    Set RsReport = Nothing
    Set FrmSaleInvoicePOS = Nothing
   End If
   
    If Grid.rows > 1 And BtnSave.Enabled = True Then
        vRemarks = "Closed"
        Call Sub_Bin_Save
        vStrSQL = "INSERT INTO ActivityLog(userno,FormType,EntryDate,Description,isnew,isedit,isdelete,isClear) values(" & vUser & ",'Sale Invoice', GetDate()," & "'BillID = " & TxtBillID.Text & " BillDate = " & DtpBillDate.DateValue & " Close' ,0,0,0,1" & ")"
        'cnSalePOS.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Clear'," & vUser & ")")
        cn.Execute (vStrSQL)
    End If
    If Grid.rows > 1 And BtnSave.Enabled = True Then
    If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnClear And vIsNewRecord = True Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = ObjRegistry.PrefixPhoneNo + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               sSQL = ObjUserSecurity.UserName & " Closed ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & " Time: " & Time & IIf(Val(TxtTotalDiscount.Caption) = 0, "", " Disc:" & TxtTotalDiscount.Caption) & vbCrLf & " NetAmt" & TxtNetAmount.Caption
               sSQL = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & sSQL & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               cn.Execute sSQL
            End If
         Next
   End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
'----------------------------------
Private Property Get FormStatus() As FormMode
  'Nothing
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
   'Based upon the value of vNewValue, we shall decide what controls to enable/disable
   On Error GoTo ErrorHandler
'   If cnSalePOS.State = adStateClosed Then cnSalePOS.Open
   If ObjRegistry.HideClearButton = True Then BtnClear.Visible = False Else BtnClear.Visible = True
   vMode = vNewValue
   Select Case vNewValue
   Case Is = NewMode
      If ObjRegistry.SaveAsNewBill = True Then BtnSaveAS.Visible = True Else BtnSaveAS.Visible = False
      Call SubClearFields
      
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      TxtStoreID.Enabled = True
      BtnStore.Enabled = True
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      BtnSaveAS.Enabled = False
      
      If Not ObjUserSecurity.IsAdministrator Then BtnOpen.Visible = ObjUserSecurity.OpenForm
      
      
      vServerDate = cn.Execute("Select CONVERT(datetime, CONVERT(varchar, GETDATE(), 110)) ServerDate").Fields(0).Value
      vSystemDate = Abs(ObjRegistry.SystemDate)
      vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
   
      vDate = IIf(vSystemDate = True, cn.Execute("Select SystemDate From SystemDate").Fields(0).Value, vServerDate)

      
      If vSystemDate = True Then
         If IsNull(vDate) Then
            If Format(Now, "hh") >= vHDiff Then
               vDate = Date
            Else
               vDate = DateAdd("d", -1, Date)
            End If
         Else
            If Format(Now, "hh") >= vHDiff Then
               vDate = vDate
            Else
               vDate = DateAdd("d", -1, vDate)
            End If
         End If
      Else
         If Format(cn.Execute("Select getdate()").Fields(0).Value, "hh") >= vHDiff Then
            vDate = vDate
         Else
            vDate = DateAdd("d", -1, vDate)
         End If
      End If
      
    
      
      DtpBillDate.DateValue = vDate
      DtpBillDate.Enabled = True
      
      If ObjUserSecurity.IsAdministrator = False Then DtpBillDate.Enabled = ObjUserSecurity.ChangeDate
      
      DtpOrderDate.DateValue = DtpBillDate.DateValue
      TxtBillID.Text = FunGetMaxID()
      LblLastBillNo.Caption = "Last Bill Nos" & FunGetLastBillID
      TxtStampID.Text = StampID()
      Call PopulateDataToGrid
      'TxtCustomerID.Text = "621"
      'TxtCustomerName.Text = "Counter Sale"
      LblStock.Visible = False
      LblAllStock.Visible = False
      LblStockCaption.Visible = False
      LblCaptionPrice.Visible = False
      LblPrice.Visible = False
      TxtCode.Enabled = True
      TxtProductName.Enabled = False
      BtnProduct.Enabled = True
      'TxtCode.Enabled = True
      'If TxtEmployeeID.Visible And TxtEmployeeID.Enabled Then TxtEmployeeID.SetFocus Else If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      vIsNewRecord = True
      vChange = False
   Case Is = OpenMode
      DtpBillDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      TxtStoreID.Enabled = False
      BtnStore.Enabled = False
      BtnPrint.Enabled = True
      BtnSaveAS.Enabled = True
      TxtCode.Enabled = True
      TxtProductName.Enabled = False
      BtnProduct.Enabled = True
      TxtCode.SetFocus
      LblStock.Visible = False
      LblAllStock.Visible = False
      LblStockCaption.Visible = False
      LblCaptionPrice.Visible = False
      LblPrice.Visible = False
      vIsNewRecord = False
      vChange = False
   Case Is = ChangeMode
      BtnPrint.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
      vChange = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub BtnProduct_Click()
   On Error GoTo ErrorHandler
   If FunSelectProduct(ssButton, True) = True Then
      If TxtQty.Enabled And TxtQty.Visible Then TxtQty.SetFocus Else TxtCode.SetFocus
   Else
      TxtCode.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCashReceivedCash_GotFocus()
   With TxtCashReceivedCash
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub TxtCashReceivedCash_Validate(Cancel As Boolean)
 Cancel = Not IsNumeric(TxtCashReceivedCash.Text)
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
      TxtPrice.Text = ""
      TxtDiscPC.Text = ""
   End If
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

'Private Sub TxtCode_LostFocus()
'   If Len(TxtCode.Text) > 7 Then
'      GetDataFromTexBoxesToGrid
'   End If
'End Sub

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

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      If vColour = True Then
         SchItemCode.ParaInWhere = " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ") "
         SchItemCode.Show vbModal, Me
         TxtCode.Text = SchItemCode.ParaOutItemCode
      Else
         SchProduct.ParaInWhere = " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ") "
         SchProduct.Show vbModal, Me
         TxtCode.Text = SchProduct.ParaOutID
      End If
      vFlag = False
   End If
    '---------------------------
   If TxtCode.Enabled = False Then FunSelectProduct = False: Exit Function
   If Trim(TxtCode.Text) = "" Then FunSelectProduct = False: Exit Function
    
   If IsNumeric(TxtCode.Text) = True Then
      If Len(TxtCode.Text) < 5 Then
         TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
      End If
   End If
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    
   If vColour = True Then
      sSQL = "select c.ColourID, ColourName from productcolours pc inner join Colours c on pc.colourid = c.colourid " & vbCrLf _
             & "inner join products p on p.productid = pc.productid " & vbCrLf _
             & "where ItemCode = '" & IIf(Len(TxtCode.Text) = 9, TxtCode.Text & "'", Mid(TxtCode.Text, 1, 9) & "' and c.colourid = " & Val(Mid(TxtCode.Text, 10, 2)))
      With cn.Execute(sSQL)
         If .RecordCount > 0 Then
            CmbColourName.AddItem !ColourName
            CmbColourName.ItemData(CmbColourName.NewIndex) = !ColourID
            CmbColourName.ListIndex = 0
         End If
      End With
      
      sSQL = "select s.SizeID, SizeName from productSizes pz inner join Sizes s on pz.Sizeid = s.Sizeid " & vbCrLf _
      & "inner join products p on p.productid = pz.productid " & vbCrLf _
      & "where ItemCode = '" & IIf(Len(TxtCode.Text) = 13, Mid(TxtCode.Text, 1, 9) & "' and s.sizeid = " & Val(Mid(TxtCode.Text, 12, 2)), TxtCode.Text & "'")
      With cn.Execute(sSQL)
         If .RecordCount > 0 Then
            cmbSizeName.AddItem !SizeName
            cmbSizeName.ItemData(cmbSizeName.NewIndex) = !SizeID
            cmbSizeName.ListIndex = 0
         End If
      End With
      TxtCode.Text = CStr(Left(TxtCode.Text, 9))
   End If
   
   ''''''''***********   Prefix BarCode For Label Weight Machine   ***********''''''''
   vBarcode = TxtCode.Text
   If ObjRegistry.BarCodePrefix = Mid(vBarcode, 1, 2) And Len(vBarcode) > 5 Then
      TxtCode.Text = Mid(vBarcode, 3, 5)
   End If
   '''''''''''''''''''''''''''''''''''''''''''
   
    ''''''''''' Show Multiplier
    vStrSQL = "Select Multiplier from  productpacking Where ProductId = '" & TxtCode.Text & "'"
    With cn.Execute(vStrSQL)
      If Not .EOF Then
         LblMultiplier.Caption = IIf(IsNull(.Fields(0).Value), "", "Pack: " & .Fields(0).Value)
      Else
         LblMultiplier.Caption = ""
      End If
    End With
   
   ''''''''***********   Checking PackageDeal   ***********''''''''
   vStrSQL = " SELECT p.productid, Code, ProductName, ServiceCharges, RetailPrice, DiscPer, DiscPC, EmpComm, isChangedPrice" & vbCrLf _
         + " from PackageDealInfoHeader un inner join Products p on un.PackageDealid = p.productid" & vbCrLf _
         + " left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
         + " where (p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "')" & " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
         
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtProductName.Text = !ProductName
         vIsChangedPrice = !isChangedPrice
         TxtPrice.Text = !RetailPrice
         vUnitPrice = TxtPrice.Text
         TxtLastPurPrice.Text = cn.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "','" & TxtPID.Text & "')").Fields(0).Value
         TxtEmpComm.Text = IIf(IsNull(!EmpComm), "", !EmpComm)
         TxtQty.Text = IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         
         
         
         
         vStrSQL = " select sum(isnull(Cost,PurPrice)* b.QtyLoose) as Cost from PackageDealInfoHeader h inner join PackageDealInfoBody b on h.id = b.id" & vbCrLf _
               + " inner join Products p on p.productid = b.productid" & vbCrLf _
               + " left outer join CurrentStock cs on cs.productid = p.productid " & vbCrLf _
               + " where h.PackageDealid ='" & TxtPID.Text & "'"
         With cn.Execute(vStrSQL)
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
   
   
         
'         VStrSQL = " select (min(css.QtyLoose/b.QtyLoose)) as QtyLoose " & vbCrLf _
'                  + " from PackageDealInfoHeader h inner join PackageDealInfoBody b on h.id = b.id" & vbCrLf _
'                  + " inner join Products p on p.productid = b.productid" & vbCrLf _
'                  + " left outer join CurrentStockStore css on css.productid = p.productid " & vbCrLf _
'                  + " where h.PackageDealid ='" & TxtPID.Text & "' and css.storeid = " & TxtStoreID.Text
'         With cnSalePOS.Execute(VStrSQL)
'            If .RecordCount > 0 Then
'               vQtyLoose = !QtyLoose
'               LblStock.Caption = IIf(IsNull(!QtyLoose), 0, !QtyLoose) & " " & cnSalePOS.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'            Else
'               vQtyLoose = 0
'               LblStock.Caption = 0
'            End If
'         End With
         If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock('" & TxtPID.Text & "',Null,0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
            LblAllStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtPID.Text & "',(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtPID.Text & "',(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & "Loose"
            LblAllStock.Visible = True
         Else
            LblAllStock.Visible = False
         End If
         
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = '" & TxtPID.Text & "'"
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock('" & TxtPID.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtPID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtPID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
         
         If .RecordCount > 0 Then
            If ObjRegistry.NegativeSale = False Then
               If vQtyLoose <= 0 Then
                  MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                  FunSelectProduct = False
                  Exit Function
               End If
            End If
            If ObjRegistry.LastRateVisible = True Then
               If TxtCustomerID.Text <> "" Then
                  LblPrice = cn.Execute("Select dbo.FunLastPrice('S','" & DtpBillDate.DateValue & "','" & TxtPID.Text & "','" & TxtCustomerID.Text & "')").Fields(0).Value
                  LblCaptionPrice.Visible = True
                  LblPrice.Visible = True
               End If
            End If
         End If
         
        
         TxtSC.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtDiscPC.Text) <> 0 Then
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / IIf(Val(TxtPrice.Text) = 0, 1, Val(TxtPrice.Text)), 2)
         End If
         ChkIsProduct.Value = 0
         SubCalculateBody
'         Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Or vChange = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      End If
   End With
   
         
    
'   MSComm1.PortOpen = False
   
   
     
''''''''***********   Checking Product  ***********''''''''
   If vEmiNo = True Then
   vStrSQL = " SELECT top 1 p.productid, code, PBS.Serial, Qty, ProductName, ServiceCharges, RetailPrice, BottomPrice, DiscPer, DiscPC, EmpComm, SaletaxPer, TokenVal, isChangedPrice" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join PurchaseBodySerial PBS on PBS.ProductID = p.productid" & vbCrLf _
           + " where (p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "' or Serial = '" & TxtCode.Text & "')" & " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
           
   Else
   vStrSQL = " SELECT p.productid, code, Qty, ProductName, ServiceCharges, p.salepackingid, Multiplier, RetailPrice, BottomPrice, DiscPer, DiscPC, EmpComm, SaletaxPer, TokenVal, isChangedPrice" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join productpacking pp on pp.productid = p.productid and p.salepackingid = pp.packingid" & vbCrLf _
           + " where (p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "')" & " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
   End If
   
   
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         
'   'Initialize customer display (clear display buffer & set row to 'upper row) with ESC & @ = Chr(27) & Chr(64)
'   MSComm1.Output = Chr(27) & Chr(64)
'   'Show item and price on first line (limited to 20 chars)
'   For i = 1 To Len(!ProductName & "  Rs." & !RetailPrice)
'   'MSComm1.Output = !ProductName & "  Rs." & !RetailPrice
'      MSComm1.Output = Space(i) & !ProductName & "  Rs." & !RetailPrice
'   Next i
'   'Show total with tax on second line (limited to 20 chars) with 'ESC & D  = Chr(27) & Chr(68)
'   'MSComm1.Output = Chr(27) & Chr(68) & "Total w/tax" & Space((20 - Len("Total w/tax")) - Len(strTotal)) & strTotal
'
         If vEmiNo = True Then TxtCode.Text = !Productid
         If vCustomerPoleDisplay = True Then
            vCounter = 0
            MSComm1.Output = Chr(CInt((&HB))) 'for home cursor
            vDisplay = !ProductName & Space(5) & "Rs." & (!RetailPrice) & Space(20)
         End If
         TxtPID.Text = !Productid
         TxtProductName.Text = !ProductName
         vIsChangedPrice = !isChangedPrice
         TxtPrice.Text = !RetailPrice
          '''''' Divide R
         If ObjRegistry.DivideRetailWithPacking = True Then
            If Not IsNull(!salepackingid) And !Multiplier <> 0 Then
                  vUnitPrice = !RetailPrice / !Multiplier
                  TxtPrice.Text = vUnitPrice
            End If
         End If
         
         '''''''
        
         
         
         vBottomPrice = IIf(IsNull(!BottomPrice), 0, !BottomPrice)
         vUnitPrice = Val(TxtPrice.Text)
         TxtLastPurPrice.Text = cn.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "','" & TxtPID.Text & "')").Fields(0).Value
         TxtEmpComm.Text = IIf(IsNull(!EmpComm), "", !EmpComm)
         TxtTokenVal.Text = IIf(IsNull(!TokenVal), "", !TokenVal)
         TxtSaleTaxPer.Text = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
         
         ''''''''***********   Prefix BarCode For Label Weight Machine   ***********''''''''
         If ObjRegistry.BarCodePrefix = Mid(vBarcode, 1, 2) And Len(vBarcode) > 5 Then
            TxtQty.Text = Round(Val(Mid(vBarcode, 8, 5)) / 1000, 3)
         Else
            TxtQty.Text = IIf(Len(TxtCode.Text) = 5 And IsNumeric(TxtCode.Text), 1, IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty))  'IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         End If
'         TxtQty.Text = IIf(Len(TxtCode.Text) = 5 And IsNumeric(TxtCode.Text), 1, IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty))  'IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         With cn.Execute("select cost from currentstock where productid ='" & TxtPID.Text & "'")
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
         If !isChangedPrice = True Then
            TxtPrice.Enabled = True
            TxtPrice.Tag = ""
         End If
         
        If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock('" & TxtPID.Text & "',Null,0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
            LblAllStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtPID.Text & "',(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtPID.Text & "',(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & "Loose"
            LblAllStock.Visible = True
         Else
            LblAllStock.Visible = False
         End If
        If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = '" & TxtPID.Text & "'"
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock('" & TxtPID.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         End If
          LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtPID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtPID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = IIf(Val(LblStock.Caption) = 0, "", LblStock.Caption) & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtPID.Text & "',(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
         If vEmiNo = True Then
            If !Serial <> "Null" Then
               TxtSerial.Text = !Serial
               GetDataFromTexBoxesToGridSerial
            End If
         End If
         
'         LblStock.Caption = vQtyLoose & " " & cnSalePOS.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'         LblStock.Visible = True
'         LblStockCaption.Visible = True
'         LblStock.Visible = True
'         LblStockCaption.Visible = True
         
                     
         If ObjRegistry.NegativeSale = False Then
            If vQtyLoose <= 0 Then
               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
               FunSelectProduct = False
               Exit Function
            End If
         End If
         If ObjRegistry.LastRateVisible = True Then
            If TxtCustomerID.Text <> "" Then
               LblPrice = cn.Execute("Select dbo.FunLastPrice('S','" & DtpBillDate.DateValue & "','" & TxtPID.Text & "','" & TxtCustomerID.Text & "')").Fields(0).Value
               LblCaptionPrice.Visible = True
               LblPrice.Visible = True
            End If
         End If
         '''''' GetExpiryTime
         vExpiryTime = 0
         sSQL = "Select dbo.GetExpiryTime('" & TxtPID.Text & "', " & IIf(TxtBatchNo.Text = "", "Null", "'" & TxtBatchNo.Text & "'") & " , getdate()) as Day "
         With cn.Execute(sSQL)
            If .RecordCount > 0 Then
               vExpiryTime = !Day
            End If
         End With
         
         
         TxtSC.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtDiscPC.Text) <> 0 Then
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / IIf(Val(TxtPrice.Text) = 0, 1, Val(TxtPrice.Text)), 2)
         End If
         ChkIsProduct.Value = 1
         If Val(TxtQty.Text) > 1 Then FindRebate
         SubCalculateBody
'         Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Or vChange = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtPID.Text = ""
         TxtCode.Text = ""
         TxtProductName.Text = ""
         TxtPrice.Text = ""
         TxtSC.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtAmount.Text = ""
         TxtCost.Text = ""
         LblStock.Visible = False
         LblStockCaption.Visible = False
         If BtnSave.Enabled = False Or vChange = False Then FormStatus = ChangeMode
         Exit Function
      End If
      If vEmiNo = True Then GetDataFromTexBoxesToGridSerial
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub GetDataFromTexBoxesToGrid()
      On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   
'   Dim x As Long
'   x = 2000 * 365
'   x = CLng(2000) * 365
   If ObjUserSecurity.IsAdministrator = True Then
      TxtPrice.Enabled = True
      TxtPrice.Tag = ""
   Else
      TxtPrice.Enabled = ObjUserSecurity.IsChangeRetail
      TxtPrice.Tag = IIf(TxtPrice.Enabled = True, "", "D")
   End If
   
   If Trim(TxtCode.Text) = "" Then
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
      TxtCode.SetFocus
      Exit Sub
   End If
   If Val(TxtQty.Text) = 0 Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      TxtQty.SetFocus
      Exit Sub
   End If
   If vBottomPrice > 0 Then
    If Round(Val(TxtAmount.Text) / Val(TxtQty.Text), 2) < vBottomPrice Then
        MsgBox "Sale Price is less than Bottom Price.", vbExclamation, "Alert"
        Exit Sub
    End If
   End If
   If Val(TxtPrice.Text) <> 0 Then
'      If Round(Val(TxtDiscPer.Text), 2) <> Round(Val(TxtDiscPer.Text), 0) Then
         Dim a As Double
         a = (Round(Val(TxtDiscPer.Text), 2) - Round((Val(TxtDiscPC.Text) * 100) / IIf(Val(TxtPrice.Text) = 0, 1, Val(TxtPrice.Text)), 2))
         If a > 1 Or a < -1 Then
            MsgBox "Please update the Discount for change Price.", vbExclamation, "Alert"
            If TxtDiscPer.Enabled And TxtDiscPer.Visible Then TxtDiscPer.SetFocus
            Exit Sub
         End If
'      End If
   End If
   
   If (CmbColourName.Text = "" Or cmbSizeName.Text = "") And vColour = True Then
      MsgBox "Please Select Colour and Size", vbInformation + vbOKOnly, "Error"
      Exit Sub
   End If
   LblLastPurPrice.Caption = cn.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "','" & TxtPID.Text & "')").Fields(0).Value
   
   If Round(Val(TxtAmount.Text) / Val(TxtQty.Text), 2) < Round(Val(LblLastPurPrice.Caption), 3) And ObjRegistry.SalePriceLessThanPurchase = False And vIsChangedPrice = False Then
      MsgBox "Sale Price is Less than Last (" & Round(Val(LblLastPurPrice.Caption), 3) & ").", vbInformation + vbOKOnly, "Alert"
      Exit Sub
   End If
   
   If Round(Val(TxtAmount.Text) / Val(TxtQty.Text), 2) < Round(Val(LblLastPurPrice.Caption), 3) And ObjRegistry.SalePriceLessThanPurchase = True And vIsChangedPrice = False Then
      If MsgBox("Sale Price is Less than Last (" & Round(Val(LblLastPurPrice.Caption), 3) & "). Do You want to continue?", vbQuestion + vbYesNo, "Alert") = vbNo Then Exit Sub
   End If
   
   If vNegativeSale = False Then
      If vIsNewRecord = True Then
         If (Val(vQtyLoose) - Val(TxtQty.Text)) < 0 Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            Grid.Redraw = True
            Call SubClearDetailArea
            If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
            Exit Sub
         End If
      Else
         If (Val(vQtyLoose) - Val(TxtQty.Text) + Val(Grid.Columns("QtyOrigional").Value)) < 0 Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            Grid.Redraw = True
            Call SubClearDetailArea
            If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
            Exit Sub
         End If
      End If
   End If
   
'   If Trim(Grid.Columns("ProductID").Text) = "" Then
'      RsBody.Filter = "ProductID = '" & TxtPID.Text & "'" & IIf(ObjRegistry.AllowEmployeProductWise = True, IIf(Trim(TxtEmployeeID.Text) = "", "", " and EmpID = '" & Trim(TxtEmployeeID.Text) & "'"), "") & IIf(ObjRegistry.AllowStoreProductWise = True, " and StoreID = " & Val(TxtStoreID.Text), "")
'   Else
'      RsBody.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'" & IIf(ObjRegistry.AllowEmployeProductWise = True, IIf(Trim(Grid.Columns("EmpID").Text) = "", "", " and EmpID = '" & Trim(Grid.Columns("EmpID").Text) & "'"), "") & IIf(ObjRegistry.AllowStoreProductWise = True, " and StoreID = " & Val(Trim(Grid.Columns("StoreID").Text)), "")
'   End If
   
   If TxtCode.Enabled Then
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.rows
               If Grid.Columns("Productid").Text = TxtPID.Text And vIsChangedPrice = False Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  If vNegativeSale = False Then
                     If vIsNewRecord = True Then
                        If (Val(vQtyLoose) - Val(TxtQty.Text) - Val(Grid.Columns("Qty").Value)) < 0 Then
                           MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                           Grid.MoveLast
                           Grid.Redraw = True
                           Exit Sub
                        End If
                     Else
                        If (Val(vQtyLoose) - Val(TxtQty.Text) - Val(Grid.Columns("Qty").Value) + Val(Grid.Columns("QtyOrigional").Value)) < 0 Then
                           MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                           Grid.MoveLast
                           Grid.Redraw = True
                           Exit Sub
                        End If
                     End If
                  End If
                                                 
                                  
                  TxtQty.Text = Val(TxtQty.Text) + Val(Grid.Columns("Qty").Value)
                  
                  Call FindRebate
                  
                  vTotDisc = vTotDisc - Val(Grid.Columns("DiscVal").Text)
                  Call SubCalculateBody
                  vTotDisc = vTotDisc + Val(TxtDiscVal.Text)
                  TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text) - Val(Grid.Columns("TotalAmount").Text)
                  TxtSumDiscAmount.Text = Val(TxtSumDiscAmount.Text) + Val(TxtDiscAmount.Text) - Val(Grid.Columns("DiscAmount").Text)
                  Call SubCalculateFooter
                  
                  
                  TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text) - Val(Grid.Columns("Qty").Text)
                  TxtTotalSaleTaxValue.Text = Val(TxtTotalSaleTaxValue.Text) + Val(TxtSaleTaxValue.Text) - Val(Grid.Columns("SaleTaxVal").Text)

                   
'                  vTotalAmount = vTotalAmount + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
'                  TxtTotalAmount.Caption = vTotalAmount
'                  TxtTotalDiscount.Caption = vTotDisc
                                  
                                    
                  TxtTotalProdProfit.Text = Round(Val(TxtTotalProdProfit.Text) + Val(TxtProdProfit.Text), 2)
'                  TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  If Val(TxtDiscVal.Text) = 0 Then
                      TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - Val(Grid.Columns("discval").Text)
                  End If
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  If ObjRegistry.AllowEmployeProductWise Then
                     Grid.Columns("EmpID").Text = TxtEmployeeID.Text
                     Grid.Columns("EmpName").Text = TxtEmployeeName.Text
                  End If
                  If ObjRegistry.AllowStoreProductWise Then
                     Grid.Columns("StoreID").Text = TxtStoreID.Text
                     Grid.Columns("StoreName").Text = TxtStoreName.Text
                  End If
                  
                  Grid.Columns("Qty").Value = Val(TxtQty.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("SC").Value = Val(TxtSC.Text)
                  Grid.Columns("DiscPC").Value = Val(TxtDiscPC.Text)
                  Grid.Columns("DiscPer").Value = Val(TxtDiscPer.Text)
                  Grid.Columns("DiscVal").Value = Val(TxtDiscVal.Text)
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  
                  Grid.Columns("SaleTaxPer").Value = Val(TxtSaleTaxPer.Text)
                  Grid.Columns("SaleTaxVal").Value = Val(TxtSaleTaxValue.Text)
                  
                  Grid.Columns("DiscAmount").Value = Val(TxtDiscAmount.Text)
                  Grid.Columns("LastPurPrice").Value = Val(TxtLastPurPrice.Text)
                  Grid.Columns("PurAmount").Value = Val(TxtPurAmount.Text)
                  Grid.Columns("ProdProfit").Value = Val(TxtProdProfit.Text)
                  Grid.Columns("Cost").Value = Val(TxtCost.Text)
                  Grid.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
                  Grid.Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
                  Grid.Columns("EmpComm").Value = IIf(Val(TxtEmpComm.Text) = 0, 0, Val(TxtEmpComm.Text))
                  Grid.Columns("TokenVal").Value = IIf(Val(TxtTokenVal.Text) = 0, 0, Val(TxtTokenVal.Text))
                  Grid.Columns("ExpiryTime").Value = Val(vExpiryTime)
                  If ObjRegistry.AllowEmployeProductWise Then
'                     RsBody!EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text))
                  End If
                  If ObjRegistry.AllowStoreProductWise Then
'                     RsBody!StoreID = IIf(Trim(TxtStoreID.Text) = "", Null, Val(TxtStoreID.Text))
                  End If
'                  RsBody!Qty = Val(TxtQty.Text)
'                  RsBody!Price = Val(TxtPrice.Text)
'                  RsBody!SC = IIf(Val(TxtSC.Text) = 0, Null, Val(TxtSC.Text))
'                  RsBody!DiscPC = Val(TxtDiscPC.Text)
'                  RsBody!DiscPer = Val(TxtDiscPer.Text)
'                  RsBody!DiscVal = Val(TxtDiscVal.Text)
'                  RsBody!Cost = Val(TxtCost.Text)
'                  RsBody!isProduct = Abs(ChkIsProduct.Value)
'                  RsBody!Amount = Val(TxtAmount.Text)
'                  RsBody!EmpComm = Val(TxtEmpComm.Text)
'                  RsBody!TokenVal = IIf(Val(TxtTokenVal.Text) = 0, Null, Val(TxtTokenVal.Text))
                  Grid.MoveLast
                  Call SubClearDetailArea
                  TxtCode.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         Grid.Columns("ProductID").Text = TxtPID.Text
         Grid.Columns("Code").Text = TxtCode.Text
'         SubClearDetailArea
         Grid.MoveLast
         TxtCode.SetFocus
'         Exit Sub
         
   End If
   'Grid.Redraw = False
   With Grid
      If TxtCode.Enabled = True Then
         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text)
         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text)
         TxtTotalSaleTaxValue.Text = Val(TxtTotalSaleTaxValue.Text) + Val(TxtSaleTaxValue.Text)
         
         TxtTotalItems.Caption = Val(TxtTotalItems.Caption) + 1
         vTotDisc = vTotDisc + Val(TxtDiscVal.Text)
         vTotalAmount = CDbl(vTotalAmount) + Val(TxtActualAmount.Text)
         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text)
         Call SubCalculateFooter
         TxtTotalDiscount.Caption = vTotDisc + Val(TxtBillDisc.Text)
         TxtSumDiscAmount.Text = Val(TxtSumDiscAmount.Text) + Val(TxtDiscAmount.Text) - Val(Grid.Columns("DiscAmount").Text)
         TxtTotalProdProfit.Text = Round(Val(TxtTotalProdProfit.Text) + Val(TxtProdProfit.Text), 2)

      Else
         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text) - Val(.Columns("Qty").Text)
         TxtTotalSaleTaxValue.Text = Val(TxtTotalSaleTaxValue.Text) + Val(TxtSaleTaxValue.Text) - Val(Grid.Columns("SaleTaxVal").Text)
         vTotDisc = TxtTotalDiscount.Caption + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text) - Val(TxtBillDisc.Text)
         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text) - Val(Grid.Columns("TotalAmount").Text)
         Call SubCalculateFooter
         TxtTotalDiscount.Caption = vTotDisc + Val(TxtBillDisc.Text)
         TxtTotalProdProfit.Text = Round(Val(TxtTotalProdProfit.Text) + Val(TxtProdProfit.Text) - Val(Grid.Columns("ProdProfit").Text), 2)
      End If
'      Call FindRebate
      .Columns("ProductName").Text = TxtProductName.Text
      If ObjRegistry.AllowEmployeProductWise Then
         .Columns("EmpID").Text = TxtEmployeeID.Text
         .Columns("EmpName").Text = TxtEmployeeName.Text
      End If
      If ObjRegistry.AllowStoreProductWise Then
         .Columns("StoreID").Text = TxtStoreID.Text
         .Columns("StoreName").Text = TxtStoreName.Text
      End If
     
      Grid.Columns("ColourName").Text = CmbColourName.Text
      If CmbColourName.Text <> "" Then Grid.Columns("ColourID").Value = CmbColourName.ItemData(CmbColourName.ListIndex)
      Grid.Columns("SizeName").Text = cmbSizeName.Text
      If cmbSizeName.Text <> "" Then Grid.Columns("SizeID").Value = cmbSizeName.ItemData(cmbSizeName.ListIndex)
      
      .Columns("Qty").Value = Val(TxtQty.Text)
      .Columns("Price").Value = Val(TxtPrice.Text)
      Grid.Columns("LastPurPrice").Value = Val(TxtLastPurPrice.Text)
      Grid.Columns("PurAmount").Value = Val(TxtPurAmount.Text)
      Grid.Columns("ProdProfit").Value = Val(TxtProdProfit.Text)
      .Columns("SC").Value = Val(TxtSC.Text)
      .Columns("DiscPC").Value = Val(TxtDiscPC.Text)
      .Columns("DiscPer").Value = Val(TxtDiscPer.Text)
      .Columns("DiscVal").Value = Val(TxtDiscVal.Text)
      .Columns("SaleTaxPer").Value = Val(TxtSaleTaxPer.Text)
      .Columns("SaleTaxVal").Value = Val(TxtSaleTaxValue.Text)
      
      If Trim(TxtCost.Text) <> "" Then
         .Columns("Cost").Value = Val(TxtCost.Text)
      End If
      .Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("DiscAmount").Value = Val(TxtDiscAmount.Text)
      .Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
      .Columns("EmpComm").Value = IIf(Val(TxtEmpComm.Text) = 0, 0, Val(TxtEmpComm.Text))
      .Columns("TokenVal").Value = IIf(Val(TxtTokenVal.Text) = 0, 0, Val(TxtTokenVal.Text))
      .Columns("ExpiryTime").Value = Val(vExpiryTime)
      If ObjRegistry.AllowEmployeProductWise Then
'         RsBody!EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text))
      End If
      If ObjRegistry.AllowStoreProductWise Then
'         RsBody!StoreID = IIf(Trim(TxtStoreID.Text) = "", Null, Val(TxtStoreID.Text))
      End If
'      RsBody!Qty = Val(TxtQty.Text)
'      RsBody!Price = Val(TxtPrice.Text)
'      RsBody!SC = IIf(Val(TxtSC.Text) = 0, Null, Val(TxtSC.Text))
'      RsBody!DiscPC = Val(TxtDiscPC.Text)
'      RsBody!DiscPer = Val(TxtDiscPer.Text)
'      RsBody!DiscVal = Val(TxtDiscVal.Text)
      If Trim(TxtCost.Text) <> "" Then
'         RsBody!Cost = Val(TxtCost.Text)
      End If
'      If IsNull(RsBody!Cost) Then RsBody!Cost = 0
'      RsBody!isProduct = Abs(ChkIsProduct.Value)
'      RsBody!Amount = Val(TxtAmount.Text)
'      RsBody!EmpComm = Val(TxtEmpComm.Text)
'      RsBody!TokenVal = IIf(Val(TxtTokenVal.Text) = 0, Null, Val(TxtTokenVal.Text))
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   vNetAmount = Val(TxtNetAmount.Caption)
   Call SubClearDetailArea
   TxtCode.SetFocus
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGridSerial()
   On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   
   If Trim(TxtSerial.Text) = "" Then
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
'      TxtSerial.SetFocus
      Exit Sub
   End If
'   RsBodySerial.Filter = "ProductID ='" & Grid.Columns("ProductID").Text & "' And Serial='" & TxtSerial.Text & "'"
        
         GridSerial.Redraw = False
         GridSerial.MoveFirst
            For vrowcounter = 1 To GridSerial.rows
               If GridSerial.Columns("Serial").Text = TxtSerial.Text Then
'                  MsgBox "The Product cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
                  vAlreadySerial = True
                  'SubClearDetailArea
                  GridSerial.MoveLast
                  TxtSerial.SetFocus
                  GridSerial.Redraw = True
                  Exit Sub
               End If
               GridSerial.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         
  If TxtSerial.Enabled Then
'         RsBodySerial.AddNew
         GridSerial.Columns("ProductID").Text = TxtCode.Text
         GridSerial.Columns("Serial").Text = TxtSerial.Text
'         RsBodySerial!Productid = TxtCode.Text
'         RsBodySerial!Serial = TxtSerial.Text
         TxtSerial.Text = ""
  End If
   'GridSerial.Redraw = False
   With GridSerial
      If Trim(.Columns("Serial").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Serial").Text = " "
         .AllowAddNew = False
      End If
   End With
   If TxtSerial.Visible = True Then TxtSerial.SetFocus
   Frame3.Visible = True
   GridSerial.Redraw = True
   Exit Sub
ErrorHandler:
   GridSerial.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   If LblStock.Visible = False Then
         LblStock.Visible = True
         LblStockCaption.Visible = True
'         LblCaptionPrice.Visible = True
'         LblPrice.Visible = True
   End If
   With Grid
      TxtPID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("code").Text
      TxtProductName.Text = .Columns("ProductName").Text
      If ObjRegistry.AllowEmployeProductWise Then
         TxtEmployeeID.Text = .Columns("EmpID").Text
         TxtEmployeeName.Text = .Columns("EmpName").Text
      End If
      If ObjRegistry.AllowStoreProductWise And (.Columns("StoreID").Text <> "") Then
         TxtStoreID.Text = .Columns("StoreID").Text
         TxtStoreName.Text = .Columns("StoreName").Text
      End If
      
      If Trim(.Columns("ColourName").Text) <> "" Then
         CmbColourName.AddItem .Columns("ColourName").Text
         CmbColourName.ItemData(CmbColourName.NewIndex) = .Columns("ColourID").Text
         CmbColourName.ListIndex = 0
      End If
      
       ''''''''''' Show Multiplier
      vStrSQL = "Select Multiplier from  productpacking Where ProductId = '" & TxtCode.Text & "'"
      With cn.Execute(vStrSQL)
         If Not .EOF Then
            LblMultiplier.Caption = IIf(IsNull(.Fields(0).Value), "", "Pack: " & .Fields(0).Value)
         Else
            LblMultiplier.Caption = ""
         End If
      End With
      If Trim(.Columns("SizeName").Text) <> "" Then
         cmbSizeName.AddItem .Columns("ColourName").Text
         cmbSizeName.ItemData(cmbSizeName.NewIndex) = .Columns("SizeID").Text
         cmbSizeName.ListIndex = 0
      End If
      If vIsChangedPrice = True Then TxtPrice.Enabled = True
      TxtQty.Text = .Columns("Qty").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtLastPurPrice.Text = .Columns("LastPurPrice").Text
      TxtPurAmount.Text = .Columns("PurAmount").Text
      TxtProdProfit.Text = .Columns("ProdProfit").Text
      TxtSC.Text = .Columns("SC").Value
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      
      TxtSaleTaxPer.Text = .Columns("SaleTaxPer").Value
      TxtSaleTaxValue.Text = .Columns("SaleTaxVal").Value

      TxtCost.Text = .Columns("Cost").Value
      TxtEmpComm.Text = .Columns("EmpComm").Value
      TxtTokenVal.Text = .Columns("TokenVal").Value
      TxtAmount.Text = .Columns("Amount").Text
      TxtDiscAmount.Text = .Columns("DiscAmount").Text
      TxtActualAmount.Text = .Columns("TotalAmount").Text
      ChkIsProduct.Value = Abs(.Columns("IsProduct").Value)
      vUnitPrice = Round((Val(TxtAmount.Text) - Val(TxtDiscVal.Text)) / IIf(Val(TxtQty.Text) = 0, 1, Val(TxtQty.Text)), 3)
      If ObjRegistry.ShowStockFromTableGridDataMovement = True Then
         If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock('" & TxtPID.Text & "',Null,0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
            LblAllStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtPID.Text & "',(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtPID.Text & "',(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & "Loose"
            LblAllStock.Visible = True
         Else
            LblAllStock.Visible = False
         End If
         
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = '" & TxtPID.Text & "'"
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock('" & TxtPID.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         End If
'         LblStock.Caption = vQtyLoose & " " & cnSalePOS.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'         LblStock.Visible = True
'         LblStockCaption.Visible = True
      
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtPID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtPID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
         
'      vUnitPrice = Val(.Columns("Price").Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))
'      vUnitRetailPrice = Val(.Columns("RetailPrice").Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))
'      If Trim(TxtPID.Text) <> "" Then
'         LblPrice.Caption = cnSalePOS.Execute("Select RetailPrice from Products where ProductID = '" & TxtPID.Text & "'").Fields(0).Value
'         LblLastPurPrice.Caption = cnSalePOS.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "','" & TxtPID.Text & "')").Fields(0).Value
'      End If
      Else
         LblStock.Visible = False
         LblStockCaption.Visible = False
      End If
   End With
   
   If Grid.rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOrganization_Click()
   If FunSelectOrganization(ssButton, False) = True Then
'      If TxtCustomerID.Enabled Then TxtCustomerID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub TxtExtraTaxPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtExtraTaxPer.Name Then Exit Sub
   TxtExtraTaxVal.Text = SelfRound((Val(TxtSumDiscAmount.Text) * Val(TxtExtraTaxPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtExtraTaxVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtExtraTaxVal.Name Then Exit Sub
   TxtExtraTaxPer.Text = Round((Val(TxtExtraTaxVal.Text) * 100) / IIf(Val(TxtSumDiscAmount.Text) = 0, 1, Val(TxtSumDiscAmount.Text)), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtOrganizationID_Change()
   On Error GoTo ErrorHandler
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizationName.Text <> "" Then TxtOrganizationName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtOrganizationID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtOrganizationID.Text) = "" Then Exit Sub
   If TxtOrganizationName.Text <> "" Then Exit Sub
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
    '---------------------------
    vStrSQL = " Select * FROM Organizations where OrganizationID=" & Val(TxtOrganizationID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizationName.Text = !OrganizationName
          FunSelectOrganization = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectOrganization = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizationName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag <> "NC" Then
            ctl.Text = ""
         End If
      End If
   Next
   CmbColourName.Clear
   cmbSizeName.Clear
   TxtBillDisc.Text = ""
   TxtBillDiscPer.Text = ""
   TxtRemarksUrdu.Text = ""
   TxtTotalQty.Caption = 0
   TxtTotalItems.Caption = 0
   TxtTotalDiscount.Caption = 0
   TxtTotalAmount.Caption = 0
   TxtNetAmount.Caption = 0
   TxtCashReceivedCash.Text = 0
   LblMultiplier.Caption = ""
   ParaOutPrevious = 0
   vTotDisc = 0
   vTotalAmount = 0
   vUnitPrice = 0
   vAmount = 0
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ProductID").Text = " "
   Grid.Update
   DtpPromiseDate.DateValue = Null
   If vCustomerPoleDisplay = True Then
      MSComm1.Output = Chr(CInt((&HB)))
      vDisplay = ""
      MSComm1.Output = Space(40)
      'Show Company Name on first line (limited to 20 chars)
      MSComm1.Output = vCompanyName
   End If
   OptCash.Value = True
   If ObjRegistry.ChangeQtyOnChangedPrice = True Then TxtAmount.Enabled = True
   Call SubClearSerialFields
'   Unload FrmPrint
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearSerialFields()
   TxtSerial.Text = ""
'   TxtSerial.Enabled = False
   GridSerial.CancelUpdate
   GridSerial.RemoveAll
   GridSerial.AddNew
   GridSerial.Columns("Serial").Text = " "
   GridSerial.Update
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpBillDate.IsDateValid = False Then FunGetMaxID = 1: Exit Function
   If ObjRegistry.AllowContinuousBillNo = True Then
      FunGetMaxID = cn.Execute("Select isnull(max(BillID),0)+1 from SaleHeader").Fields(0)
   ElseIf ObjRegistry.AllowMonthlyBillNo = True Then
      FunGetMaxID = cn.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where Month(BillDate) = '" & Month(DtpBillDate.DateValue) & "' and  year(BillDate) ='" & Year(DtpBillDate.DateValue) & "'").Fields(0)
   ElseIf ObjRegistry.AllowDailyBillNo = True Then
      FunGetMaxID = cn.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where BillDate = '" & DtpBillDate.DateValue & "'").Fields(0)
   Else
      FunGetMaxID = cn.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where BillDate = '" & DtpBillDate.DateValue & "' and StoreID = " & TxtStoreID.Text).Fields(0)
   End If
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetLastBillID() As String
   On Error GoTo ErrorHandler
      FunGetLastBillID = ""
      With cn.Execute("Select top 3 billId from  saleheader where userno = " & vUser & " order by billdate desc, billid desc ")
         While Not .EOF
            FunGetLastBillID = FunGetLastBillID & " :- " & .Fields(0)
            .MoveNext
         Wend
      End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function StampID() As Long
   On Error GoTo ErrorHandler
   StampID = cn.Execute("Select isnull(max(SID),0)+1 from Stamp").Fields(0)
   cn.Execute "update Stamp set SID = " & StampID
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
'   RsBody.Open "Select * from SaleBody where BillId = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "' --and StampID = " & TxtStampID.Text, cnSalePOS, adOpenStatic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      sSQL = " Select p.ProductName, b.code, HeaderStoreID, b.ColourID, b.SizeID, ColourName, SizeName, EmpName, StoreName, b.ProductID, ProductName, isProduct," & vbCrLf & _
             " IsNull(Multiplier,0)*IsNull(QtyPack,0)+ Qty as Qty, round(Price/isnull(multiplier,1),2) as Price," & vbCrLf & _
             " SC, b.DiscPC, b.DiscPer, b.DiscVal, Amount, b.DiscAmount, Cost, b.EmpComm, b.TokenVal" & vbCrLf & _
             " from Salebody b Left Outer join products p on p.productid = b.productid left outer join Employees e on e.empid = b.empid " & vbCrLf & _
             " Left outer join Colours Col on Col.Colourid = b.ColourID Left Outer join Sizes Sz on Sz.SizeID = b.SizeID left outer join Stores s on s.StoreID = b.StoreID " & vbCrLf & _
             " where SID=" & Val(TxtSID.Text) & " Order by SerialNo asc "
      With cn.Execute(sSQL)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         'TxtGrossAmount.Text = 0
         TxtTotalQty.Caption = 0
         TxtTotalItems.Caption = 0
         'TxtTotalDiscount.Caption = 0
         vTotDisc = 0
         vTotalAmount = 0
         vNetAmount = 0
         TxtTotalAmount.Caption = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("StoreID").Text = IIf(IsNull(!HeaderStoreID), "", !HeaderStoreID)
            Grid.Columns("ProductName").Text = !ProductName
            If ObjRegistry.AllowEmployeProductWise Then
               Grid.Columns("EmpID").Text = IIf(IsNull(!EmpID), "", !EmpID)
               Grid.Columns("EmpName").Text = IIf(IsNull(!empname), "", !empname)
            End If
            If ObjRegistry.AllowStoreProductWise Then
               Grid.Columns("StoreID").Text = IIf(IsNull(!StoreID), "", !StoreID)
               Grid.Columns("StoreName").Text = IIf(IsNull(!StoreName), "", !StoreName)
            End If
            Grid.Columns("Qty").Value = !Qty
            Grid.Columns("QtyOrigional").Value = !Qty
            Grid.Columns("Price").Value = !Price
            Grid.Columns("SC").Value = IIf(IsNull(!SC), "", !SC)
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)

            Grid.Columns("ColourID").Value = IIf(IsNull(!ColourID), "", !ColourID)
            Grid.Columns("ColourName").Value = IIf(IsNull(!ColourName), "", !ColourName)
            Grid.Columns("SizeID").Value = IIf(IsNull(!SizeID), "", !SizeID)
            Grid.Columns("SizeName").Value = IIf(IsNull(!SizeName), "", !SizeName)


            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("DiscAmount").Value = !DiscAmount
            vNetAmount = vNetAmount + !Amount

            '''''''''''''''''''''''''' get prod proft
            TxtLastPurPrice.Text = cn.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "','" & Grid.Columns("ProductID").Text & "')").Fields(0).Value
            Grid.Columns("LastPurPrice").Value = Val(TxtLastPurPrice.Text)
            TxtPurAmount.Text = Round(Val(Grid.Columns("Qty").Value) * (Val(TxtLastPurPrice.Text) + Val(Grid.Columns("SC").Value)), 2)
            Grid.Columns("PurAmount").Value = Val(TxtPurAmount.Text)
            TxtProdProfit.Text = Round(Val(Grid.Columns("Amount").Value) - Val(TxtPurAmount.Text), 2)
            Grid.Columns("ProdProfit").Value = Val(TxtProdProfit.Text)
            TxtTotalProdProfit.Text = Round(Val(TxtTotalProdProfit.Text) + Val(TxtProdProfit.Text), 2)

            ''''''''''''''''''''''''''''''''''''''


            Grid.Columns("IsProduct").Value = Abs(!isProduct)
            Grid.Columns("TotalAmount").Value = Val(!Qty) * (Val(!Price) + Val(IIf(IsNull(!SC), "", !SC)))
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            Grid.Columns("EmpComm").Value = IIf(IsNull(!EmpComm), "", !EmpComm)
            Grid.Columns("TokenVal").Value = IIf(IsNull(!TokenVal), "", !TokenVal)
            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
            TxtTotalItems.Caption = Val(TxtTotalItems.Caption) + 1
            'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(!DiscVal)
            vTotDisc = vTotDisc + Val(!DiscVal)
            vTotalAmount = vTotalAmount + !Amount
            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
            .MoveNext
         Wend
         .Close
      End With
      TxtTotalAmount.Caption = SelfRound(TxtTotalAmount.Caption)
      Call SubCalculateBody
      Grid.AddNew
      Grid.Columns("ProductID").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
'   End If

   If RsBodyStore.State = adStateOpen Then RsBodyStore.Close
   RsBodyStore.Filter = 0
   RsBodyStore.Open "Select * from SaleBodyStore where BillID = " & Val(TxtBillID.Text) & " And SID = " & Val(TxtSID.Text) & " and Billdate = '" & DtpBillDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
   RsDetail.Filter = 0
   If RsDetail.State = adStateOpen Then RsDetail.Close
   RsDetail.Open "Select * from SaleUnionUsed where BillId=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
   PopulateDataToGridserial
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub PopulateSaleOrderDataToGrid()
   On Error GoTo ErrorHandler
   RsBody.Filter = 0
'   If RsBody.State = adStateOpen Then RsBody.Close
'   RsBody.Open "Select * from SaleBody where BillId=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      sSQL = " select sob.ProductID, ProductName, (isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty - isnull(uqty,0) as Qtyloose, sob.*" & vbCrLf _
      + " from (select OrderID, OrderDate, ProductID, Sum(Qty) as UQty from SaleBody b inner join SaleHeader h on h.BillID = b.BillID and h.BillDate = b.BillDate Group By OrderID, OrderDate, ProductID) b " & vbCrLf _
      + " right outer join SaleOrderBody sob on sob.OrderID = b.orderid and sob.OrderDate = b.orderdate and b.ProductID = sob.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = sob.productid" & vbCrLf _
      + " where sob.OrderID = " & Val(TxtOrderID.Text) & " and sob.OrderDate = '" & DtpOrderDate.DateValue & "' and (isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty - isnull(uqty,0) <> 0 order by sob.serialno"
      With cn.Execute(sSQL)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         'TxtGrossAmount.Text = 0
         TxtTotalQty.Caption = 0
         TxtTotalItems.Caption = 0
         'TxtTotalDiscount.Caption = 0
         vTotDisc = 0
         vTotalAmount = 0
         TxtTotalAmount.Caption = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !QtyLoose
            Grid.Columns("QtyOrigional").Value = !QtyLoose
            Grid.Columns("Price").Value = !Price
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = (Val(!Price) - IIf(IsNull(!DiscPC), 0, !DiscPC)) * Val(!QtyLoose)
            Grid.Columns("IsProduct").Value = Abs(!isProduct)
            Grid.Columns("TotalAmount").Value = Val(!Price) * Val(!QtyLoose)
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            Grid.Columns("EmpComm").Value = IIf(IsNull(!EmpComm), "", !EmpComm)
            Grid.Columns("DiscAmount").Value = (Val(!QtyLoose) * (Val(!Price))) + Val(!DiscVal)
            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!QtyLoose)
            TxtTotalItems.Caption = Val(TxtTotalItems.Caption) + 1
            'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(!DiscVal)
            vTotDisc = vTotDisc + Val(!DiscVal)
            vTotalAmount = vTotalAmount + (Val(!Price) - IIf(IsNull(!DiscPC), 0, !DiscPC)) * Val(!QtyLoose)
            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
                  
'            RsBody.AddNew
'            RsBody!Productid = !Productid
'            RsBody!code = IIf(IsNull(!code), "", !code)
'            RsBody!Qty = !QtyLoose
'            RsBody!Price = !Price
'            RsBody!SC = Null
'            RsBody!DiscPC = IIf(IsNull(!DiscPC), 0, !DiscPC)
'            RsBody!DiscPer = IIf(IsNull(!DiscPer), 0, !DiscPer)
'            RsBody!DiscVal = IIf(IsNull(!DiscVal), 0, !DiscVal)
'            RsBody!Cost = 0
'            RsBody!isProduct = Abs(!isProduct)
'            RsBody!Amount = (Val(!Price) - IIf(IsNull(!DiscPC), 0, !DiscPC)) * Val(!QtyLoose)
'            RsBody!EmpComm = IIf(IsNull(!EmpComm), 0, !EmpComm)
'            RsBody.Update
            .MoveNext
         Wend
         .Close
      End With
      Call SubCalculateBody
      Grid.AddNew
      Grid.Columns("ProductID").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
'   End If
   RsDetail.Filter = 0
   If RsDetail.State = adStateOpen Then RsDetail.Close
   RsDetail.Open "Select * from SaleUnionUsed where BillId=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub
Private Sub PopulateDataToGridserial()
'   RsBodySerial.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
'   If RsBodySerial.RecordCount > 0 Then
       sSQL = "select d.* from SaleBodySerial d  where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
      With cn.Execute(sSQL)
'       With RsBodySerial
         GridSerial.Redraw = False
         GridSerial.MoveFirst
         GridSerial.RemoveAll
         GridSerial.AllowAddNew = True
'         .MoveFirst
         While Not .EOF
            GridSerial.AddNew
            GridSerial.Columns("ProductID").Text = !Productid
            GridSerial.Columns("Serial").Text = !Serial
            .MoveNext
         Wend
      .Close
      End With
      GridSerial.AddNew
      GridSerial.Columns("Serial").Text = " "
      GridSerial.AllowAddNew = False
      GridSerial.Redraw = True
'   Else
'    Call SubClearSerialFields
'   End If
'   RsBodySerial.Filter = 0
End Sub

Private Sub SubCalculateBody()
   TxtActualAmount.Text = Round(Val(TxtQty.Text) * (Val(TxtPrice.Text) + Val(TxtSC.Text)), 0)
   TxtDiscVal.Text = Round(Val(TxtQty.Text) * Val(TxtDiscPC.Text), 0)
   TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
   If ObjRegistry.IsRoundFigure = True Then TxtAmount.Text = SelfRound(TxtAmount.Text)
   TxtPurAmount.Text = Round(Val(TxtQty.Text) * (Val(TxtLastPurPrice.Text) + Val(TxtSC.Text)), 2)
   TxtProdProfit.Text = Round(Val(TxtAmount.Text) - Val(TxtPurAmount.Text), 2)
   TxtTotalDiscount.Caption = Round(vTotDisc, 2)
   TxtDiscAmount.Text = (Val(TxtQty.Text) * (Val(TxtPrice.Text) + Val(TxtSC.Text))) + Val(TxtDiscVal.Text)
   SubCalculateFooter
   
End Sub

Private Sub SubCalculateFooter()
   If Val(TxtBillDisc.Text) <> 0 Then
      If DiscPerFlag = False Then
         TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / IIf(Val(TxtTotalAmount.Caption) = 0, 1, Val(TxtTotalAmount.Caption)), 2)
      Else
         TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtBillDiscPer.Text) / 100))
      End If
   End If
   TxtTotalDiscount.Caption = Round(Val(TxtBillDisc.Text) + vTotDisc, 2)
   TxtNetAmount.Caption = SelfRound(Val(TxtTotalAmount.Caption) - Val(TxtTotalDiscount.Caption) + Val(TxtServiceCharges.Text) + Val(TxtSTax.Text) + Val(TxtAdvTaxVal.Text) + Val(TxtExtraTaxVal.Text))
'   TxtTotalProdProfit.Text = TxtProdProfit.Text
   'If TxtGrossAmount.Text = "" Then Exit Sub
   'TxtNetAmount.Caption = Round(Val(TxtGrossAmount.Text)) - Val(TxtBillDisc.Text)
   'TxtCashReturn.Text = IIf(Val(TxtCashReceived.Text) > 0, Val(TxtCashReceived.Text) - Val(TxtNetAmount.Caption), "")
End Sub
Private Sub Textbox1_KeyDown(KeyCode As Integer, Shift As Integer)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyDown-Event behaviours for  ''
        '' Enter, Space, Delete & Tab keys to set      ''
        '' Behavior in TxtRemarksUrdu.Text, keys will behave ''
        '' as Normal Text writing behavior.            ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''
     
        'Space Key Behavior
        If KeyCode = 32 Then
        UniCode = &H20
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        KeyCode = 0
        
        'Enter Key Behavior
        ElseIf KeyCode = 13 Then
        UniCode = &HA
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        KeyCode = 0
        
        'Horizontal Tab Behavior
        ElseIf KeyCode = 9 Then
        UniCode = &H9
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        KeyCode = 0
        
        'Delete Key Behavior
        ElseIf KeyCode = 127 Then
        UniCode = &H7F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        KeyCode = 0
        
        End If
        If BtnSave.Enabled = False Then FormStatus = ChangeMode
        
        'This Function Got End There

End Sub

Private Sub Textbox1_KeyPress(KeyAscii As Integer)

        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyPress-Event behaviours for ''
        '' Alfabatic, Numaric & Symbolic keys to write ''
        '' Urdu. I've tried to make it near with Urdu  ''
        '' Phonetic Keyboard Layout.                   ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''
       
'If ModeValue = False Then

        'For Small Letter's Behaviors

        'a Key Behavior
        If KeyAscii = 97 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H627
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'b Key Behavior
        ElseIf KeyAscii = 98 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H628
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'c Key Behavior
        ElseIf KeyAscii = 99 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H686
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'd Key Behavior
        ElseIf KeyAscii = 100 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'e Key Behavior
        ElseIf KeyAscii = 101 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H639
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'f Key Behavior
        ElseIf KeyAscii = 102 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H641
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'g Key Behavior
        ElseIf KeyAscii = 103 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6AF
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'h Key Behavior
        ElseIf KeyAscii = 104 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6BE
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'i Key Behavior
        ElseIf KeyAscii = 105 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6CC
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'j Key Behavior
        ElseIf KeyAscii = 106 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62C
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'k Key Behavior
        ElseIf KeyAscii = 107 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6A9
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'l Key Behavior
        ElseIf KeyAscii = 108 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H644
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'm Key Behavior
        ElseIf KeyAscii = 109 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H645
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'n Key Behavior
        ElseIf KeyAscii = 110 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H646
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'o Key Behavior
        ElseIf KeyAscii = 111 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6C1
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'p Key Behavior
        ElseIf KeyAscii = 112 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H67E
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'q Key Behavior
        ElseIf KeyAscii = 113 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H642
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'r Key Behavior
        ElseIf KeyAscii = 114 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H631
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        's Key Behavior
        ElseIf KeyAscii = 115 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H633
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        't Key Behavior
        ElseIf KeyAscii = 116 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62A
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'u Key Behavior
        ElseIf KeyAscii = 117 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H621
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'v Key Behavior
        ElseIf KeyAscii = 118 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H637
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'w Key Behavior
        ElseIf KeyAscii = 119 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H648
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'x Key Behavior
        ElseIf KeyAscii = 120 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H634
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'y Key Behavior
        ElseIf KeyAscii = 121 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6D2
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'z Key Behavior
        ElseIf KeyAscii = 122 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H632
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        
        ' For Capital Latter's Behaviors
        
        'A Key Behavior
        ElseIf KeyAscii = 65 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H622
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'B Key Behavior
        ElseIf KeyAscii = 66 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &HFBB0
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'C Key Behavior
        ElseIf KeyAscii = 67 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62B
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'D Key Behavior
        ElseIf KeyAscii = 68 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H688
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'E Key Behavior
        ElseIf KeyAscii = 69 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H650
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'F Key Behavior
        ElseIf KeyAscii = 70 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H652
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'G Key Behavior
        ElseIf KeyAscii = 71 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H63A
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'H Key Behavior
        ElseIf KeyAscii = 72 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'I Key Behavior
        ElseIf KeyAscii = 73 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H649
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'J Key Behavior
        ElseIf KeyAscii = 74 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H636
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'K Key Behavior
        ElseIf KeyAscii = 75 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62E
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'L Key Behavior
        ElseIf KeyAscii = 76 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &HFEFB
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'M Key Behavior
        ElseIf KeyAscii = 77 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H66B
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'N Key Behavior
        ElseIf KeyAscii = 78 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6BA
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'O Key Behavior
        ElseIf KeyAscii = 79 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H629
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'P Key Behavior
        ElseIf KeyAscii = 80 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H64F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'Q Key Behavior
        ElseIf KeyAscii = 81 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H626
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'R Key Behavior
        ElseIf KeyAscii = 82 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H691
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'S Key Behavior
        ElseIf KeyAscii = 83 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H635
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'T Key Behavior
        ElseIf KeyAscii = 84 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H679
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'U Key Behavior
        ElseIf KeyAscii = 85 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H626
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'V Key Behavior
        ElseIf KeyAscii = 86 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H638
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'W Key Behavior
        ElseIf KeyAscii = 87 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H624
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'Z Key Behavior
        ElseIf KeyAscii = 88 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H698
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'Y Key Behavior
        ElseIf KeyAscii = 89 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &HFBAF
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'Z Key Behavior
        ElseIf KeyAscii = 90 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H630
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        
        'For Numaric Key's Behaviors
                
        '0 Key Behavior
        ElseIf KeyAscii = 48 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 48
'        UniCode = &H660
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '1 Key Behavior
        ElseIf KeyAscii = 49 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 49
'        UniCode = &H661
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '2 Key Behavior
        ElseIf KeyAscii = 50 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 50
'        UniCode = &H662
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '3 Key Behavior
        ElseIf KeyAscii = 51 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 51
'        UniCode = &H663
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '4 Key Behavior
        ElseIf KeyAscii = 52 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 52
'        UniCode = &H664
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '5 Key Behavior
        ElseIf KeyAscii = 53 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 53
'        UniCode = &H665
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '6 Key Behavior
        ElseIf KeyAscii = 54 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 54
'        UniCode = &H666
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '7 Key Behavior
        ElseIf KeyAscii = 55 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 55
'        UniCode = &H667
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '8 Key Behavior
        ElseIf KeyAscii = 56 Or TxtRemarksUrdu.SelText <> "" Then
        UniCode = 56
'        UniCode = &H668
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '9 Key Behavior
        ElseIf KeyAscii = 57 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 57
'        UniCode = &H669
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)

        ' Numaric Keys with 'Shift' Behavior
        
        ') Key Behavior
        ElseIf KeyAscii = 41 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &HFD3F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '! Key Behavior
        ElseIf KeyAscii = 33 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H21
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '@ Key Behavior
        ElseIf KeyAscii = 64 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H40
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '# Key Behavior
        ElseIf KeyAscii = 35 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H23
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '$ Key Behavior
        ElseIf KeyAscii = 36 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H24
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '% Key Behavior
        ElseIf KeyAscii = 37 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H66A
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '^ Key Behavior
        ElseIf KeyAscii = 94 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H5E
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '& Key Behavior
        ElseIf KeyAscii = 38 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H26
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '* Key Behavior
        ElseIf KeyAscii = 42 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H66D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '( Key Behavior
        ElseIf KeyAscii = 40 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &HFD3E
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        
        'For Special Characters
        
        'Symbols
        
        '? Key Behavior
        ElseIf KeyAscii = 63 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H61F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '/ Key Behavior
        ElseIf KeyAscii = 47 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        ', Key Behavior
        ElseIf KeyAscii = 44 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H60C
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '. Key Behavior
        ElseIf KeyAscii = 46 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H640
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '_ Key Behavior
        ElseIf KeyAscii = 95 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H5F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '- Key Behavior
        ElseIf KeyAscii = 45 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '+ Key Behavior
        ElseIf KeyAscii = 43 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2B
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '= Key Behavior
        ElseIf KeyAscii = 61 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H3D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        ': Key Behavior
        ElseIf KeyAscii = 58 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H3A
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '; Key Behavior
        ElseIf KeyAscii = 59 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H201C
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '< Key Behavior
        ElseIf KeyAscii = 60 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H64E
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '> Key Behavior
        ElseIf KeyAscii = 62 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H650
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '{ Key Behavior
        ElseIf KeyAscii = 123 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2018
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '} Key Behavior
        ElseIf KeyAscii = 125 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2019
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '[ Key Behavior
        ElseIf KeyAscii = 91 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H5B
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '] Key Behavior
        ElseIf KeyAscii = 93 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H5D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '| Key Behavior
        ElseIf KeyAscii = 124 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H7C
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '\ Key Behavior
        ElseIf KeyAscii = 92 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H5C
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '~ Key Behavior
        ElseIf KeyAscii = 126 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H64B
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '` Key Behavior
        ElseIf KeyAscii = 96 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H64D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '" Key Behavior
        ElseIf KeyAscii = 34 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2190
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '' Key Behavior
        ElseIf KeyAscii = 39 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H201D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        End If
        KeyAscii = 0
 '       End If

        'This Function Got End There
End Sub

Private Sub SubEnable(vFlag As Boolean)
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is SITextBox.txt Or TypeOf ctl Is JeweledButton Or TypeOf ctl Is OptionButton Or TypeOf ctl Is CheckBox Or TypeOf ctl Is SSDateCombo Or TypeOf ctl Is SSOleDBGrid Then
         If ctl.Tag = "F" Then
            ctl.Enabled = Not vFlag
         ElseIf ctl.Tag = "D" Or ctl.Tag = "NC" Then
            ctl.Enabled = False
         Else
            ctl.Enabled = vFlag
         End If
      End If
   Next
   If vChangeQtyOnChangedPrice = True Then TxtAmount.Enabled = True
   TxtStoreID.Enabled = True
   TxtOrganizationID.Enabled = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   CmbColourName.Clear
   cmbSizeName.Clear
   TxtCode.Enabled = True
   BtnProduct.Enabled = True
   TxtCode.Text = ""
   TxtProductName.Text = ""
   TxtQty.Text = ""
   TxtPrice.Text = ""
   TxtSC.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtDiscVal.Text = ""
   TxtAmount.Text = ""
   TxtCost.Text = ""
   TxtActualAmount.Text = ""
   TxtEmpComm.Text = ""
   TxtTokenVal.Text = ""
   ChkIsProduct.Value = 1
   LblMultiplier.Caption = ""
End Sub
Private Sub Grid_DblClick()
On Error GoTo ErrorHandler
   If Flag Then Call GetDataBackFromGridToTexBoxes
   Call Grid_LostFocus
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Call ShowErrorMessage
End Sub

Private Sub Grid_GotFocus()
On Error GoTo ErrorHandler
   Flag = True
   TxtCode.Enabled = False
   BtnProduct.Enabled = False
   'TxtCode.BackColor = TxtProductName.BackColor
   'TxtCode.TabStop = False
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Call ShowErrorMessage
End Sub

Private Sub Grid_LostFocus()
On Error GoTo ErrorHandler
   'If BtnSave.Enabled = False Then Call SubFrameLoad:    Exit Sub
   If Grid.Visible = False Or Grid.Enabled = False Then Exit Sub
   Flag = False
   LblCost.Visible = False
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      TxtCode.Text = ""
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      TxtCode.Enabled = False
      BtnProduct.Enabled = False
      If TxtQty.Enabled = True And TxtQty.Visible Then TxtQty.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Call ShowErrorMessage
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Call ShowErrorMessage
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo ErrorHandler
   If Grid.Enabled = False Or Grid.Visible = False Then Exit Sub
   If ActiveControl.Name <> Grid.Name Then Exit Sub
   'If Flag Then
   Call GetDataBackFromGridToTexBoxes
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Call ShowErrorMessage
End Sub
Private Sub Grid_RowLoaded(ByVal Bookmark As Variant)
On Error GoTo ErrorHandler
   With Grid
'      If Val(.Columns("ExpiryTime").Value) = 0 Then
'         .Columns("ProductName").CellStyleSet ""
'      ElseIf Val(.Columns("ExpiryTime").Value) <= 90 And Val(.Columns("ExpiryTime").Value) > 30 Then
'         .Columns("ProductName").CellStyleSet "Green"
'      ElseIf Val(.Columns("ExpiryTime").Value) <= 30 And Val(.Columns("ExpiryTime").Value) > 0 Then
'         .Columns("ProductName").CellStyleSet "Orange"
'      ElseIf Val(.Columns("ExpiryTime").Value) < 0 Then
'         .Columns("ProductName").CellStyleSet "Red"
'      End If
      '''''' Get ExpiryColor
      If Val(.Columns("ExpiryTime").Value) < 0 Then
         .Columns("ProductName").CellStyleSet "Red"
      Else
         sSQL = "Select * from ExpiryDayColor Where " & Val(.Columns("ExpiryTime").Value) & " >= DayFrom and " & Val(.Columns("ExpiryTime").Value) & " <= DayTo"
         With cn.Execute(sSQL)
            If .RecordCount <> 0 Then vExpiryColor = !ExpiryColor Else vExpiryColor = ""
         End With
      .Columns("ProductName").CellStyleSet vExpiryColor
      End If
   End With
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   'TxtGrossAmount.Text = Val(TxtGrossAmount.Text) - Grid.Columns("Amount").Value
   TxtTotalQty.Caption = Round(Val(TxtTotalQty.Caption) - Grid.Columns("Qty").Value, 2)
   TxtTotalItems.Caption = Val(TxtTotalItems.Caption) - 1
   TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) - Grid.Columns("TotalAmount").Value
   vTotalAmount = TxtTotalAmount.Caption
   If Val(TxtBillDisc.Text) <> 0 Then
      If DiscPerFlag = False Then
         TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / IIf(Val(TxtTotalAmount.Caption) = 0, 1, Val(TxtTotalAmount.Caption)), 2)
      Else
         TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtBillDiscPer.Text) / 100))
      End If
   End If
   vTotDisc = vTotDisc - Val(Grid.Columns("DiscVal").Value) + Val(TxtBillDisc.Text)
   TxtTotalDiscount.Caption = Round(Val(vTotDisc), 2)
   vTotDisc = vTotDisc - Val(TxtBillDisc.Text)
   TxtNetAmount.Caption = SelfRound(Val(TxtTotalAmount.Caption) - Val(TxtTotalDiscount.Caption) + Val(TxtServiceCharges.Text) + Val(TxtSTax.Text))
   Call ActivityLogSale("Sale Invoice", eDelete, TxtBillID.Text, DtpBillDate.DateValue)
   'SubCalculateFooter
   FormStatus = ChangeMode
   Exit Sub
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
'   RsBody.Filter = "ProductID = '" & TxtPID.Text & "'" & IIf(ObjRegistry.AllowEmployeProductWise = True, IIf(Trim(TxtEmployeeID.Text) = "", "", " and EmpID = '" & Trim(TxtEmployeeID.Text) & "'"), "") & IIf(ObjRegistry.AllowStoreProductWise = True, " and StoreID = " & Val(TxtStoreID.Text), "")
   'RsBody.Filter = "Code='" & TxtCode.Text & "'"
'   If RsBody.RecordCount > 0 Then RsBody.Delete
    If ObjRegistry.UsePasswordForm = True And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False Then
      If UsePasswordForm = False Then Exit Sub
   End If
   ''''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '/******* Mobile SMS *************/
   vStrDetail = ""
   vStrDetail = vStrDetail & " (P" & Grid.Columns("ProductID").Text & " Q" & Grid.Columns("Qty").Text & " A" & Grid.Columns("Amount").Text & ")"
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnClear Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = ObjRegistry.PrefixPhoneNo + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               sSQL = ObjUserSecurity.UserName & " Removed Item ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & " Time: " & Time & IIf(Val(TxtTotalDiscount.Caption) = 0, "", " Disc:" & TxtTotalDiscount.Caption) & vbCrLf & " NetAmt" & TxtNetAmount.Caption
               sSQL = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & sSQL & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               cn.Execute sSQL
            End If
         Next
   End If
   '''' delete serial grid
   GridSerial.Redraw = False
   GridSerial.MoveFirst
    For vCounter = 1 To GridSerial.rows
      If Trim(GridSerial.Columns("ProductID").Text) <> "" Then
         If GridSerial.Columns("ProductID").Text = Grid.Columns("ProductID").Text Then
            GridSerial.SelBookmarks.Add GridSerial.Row
            GridSerial.SelBookmarks.Add GridSerial.Bookmark
            GridSerial.DeleteSelected
         End If
      End If
      GridSerial.MoveNext
   Next vCounter
  GridSerial.Redraw = True
   
   '''''' delete sale body grid
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   
   
'   RsBody.Filter = 0
   Grid.MoveLast
   GetDataBackFromGridToTexBoxes
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub BtnCancel_Click()
   On Error GoTo ErrorHandler
   Call SubEnable(True)
   Frame1.Visible = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub BtnSaveAS_Click()
vIsNewRecord = True
DtpBillDate.Date = Date
Call BtnSave_Click
End Sub

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   If ObjRegistry.UsePasswordForm = True And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False Then
      If UsePasswordForm = False Then Exit Sub
   End If
   
'   If Grid.Rows <= 1 And TxtPID.Text = "" Then Exit Sub
   If MsgBox("Do you want to Clear this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   
   If vIsNewRecord = True And vChange = True And TxtOrderID.Text <> "" And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a Sale Order.", vbCritical, "Error"
      Exit Sub
   End If
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
'      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
   
'   If cnSalePOS.State = adStateClosed Then cnSalePOS.Open
   'Header Validation
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      TxtStoreID.SetFocus
      Exit Sub
   End If
   
   If Trim(TxtEmployeeID.Text) = "" Then
      SubDestroyEmployeeCommision
   Else
      SubApplyEmployeeCommision
   End If
   
  
    
   TxtNetAmountCash.Text = TxtNetAmount.Caption
   If ObjRegistry.CashReceived = True Then
      TxtCashReceivedCash.Text = TxtNetAmount.Caption
   End If
   vRemarks = "Clear"
   Call Sub_Bin_Save
   If Grid.rows > 1 Then
      vStrSQL = "INSERT INTO ActivityLog(userno,FormType,EntryDate,Description,isnew,isedit,isdelete,isClear) values(" & vUser & ",'Sale Invoice', GetDate()," & "'BillID = " & TxtBillID.Text & " BillDate = " & DtpBillDate.DateValue & " Clear' ,0,0,0,1" & ")"
    'cnSalePOS.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Clear'," & vUser & ")")
      cn.Execute (vStrSQL)
   End If
'   ChkPrint.Value = Abs(ObjRegistry.AutoPrintinInvoices)
   
   vDescription = "ID " & TxtBillID.Text & " Date " & DtpBillDate.DateValue
   
      
   '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnClear And vIsNewRecord = True Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = ObjRegistry.PrefixPhoneNo + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               sSQL = ObjUserSecurity.UserName & " Cleared ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & " Time: " & Time & IIf(Val(TxtTotalDiscount.Caption) = 0, "", " Disc:" & TxtTotalDiscount.Caption) & vbCrLf & " NetAmt" & TxtNetAmount.Caption
               sSQL = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & sSQL & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               cn.Execute sSQL
            End If
         Next
   End If
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   ''''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cnSalePOS.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnDelete_Click()
    On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   vRemarks = "Delete"
   vDescription = "ID " & TxtBillID.Text & " Date " & DtpBillDate.DateValue
   Call Sub_Bin_Save
   cn.BeginTrans
   
   Call BinData("Click On Delete Button", eDelete)
'   vMaxBinID = FunGetMaxBinID
'   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   cnSalePOS.Execute ("Insert Into Bin_SaleHeader Select " & vMaxBinID & ",'" & Date & "',*," & 1 & "," & 0 & " from SaleHeader Where BillID = " & TxtBillID.Text & " And BillDate ='" & DtpBillDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   cnSalePOS.Execute ("Insert Into Bin_SaleBody Select " & vMaxBinID & ",'" & Date & "', * from SaleBody Where BillID = " & TxtBillID.Text & " And BillDate ='" & DtpBillDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cnSalePOS.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Grid.Redraw = False
   Grid.MoveFirst
   Call ActivityLogSale("Sale Invoice", eDelete, TxtBillID.Text, DtpBillDate.DateValue)
'   For vCounter = 1 To RsDetail.RecordCount
'      cnSalePOS.Execute "Delete from SaleUnionUsed where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "' and Productid ='" & RsDetail!Productid & "'"
'      RsDetail.MoveNext
'   Next vCounter
   vStrDetail = ""
   ''''''''''''''''''' Delete salebodyserial '''''''''''''''''''''''
   cn.Execute "Delete from salebodyserial where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   
   vStrSQL = "Delete from SaleBody where SID = " & Val(TxtSID.Text)
   cn.Execute vStrSQL
   For vCounter = 1 To Grid.rows
      If Trim(Grid.Columns("ProductID").Text) <> "" Then
         cn.Execute "Exec UpdateStockPlus " & TxtStoreID.Text & ",'" & Grid.Columns("ProductID").Text & "'," & Grid.Columns("Qty").Value & "," & Val(TxtBillID.Text) & ",'" & DtpBillDate.DateValue & "'"
         vStrDetail = vStrDetail & " (P" & Grid.Columns("ProductID").Text & " Q" & Grid.Columns("Qty").Text & " A" & Grid.Columns("Amount").Text & ")"
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   vStrSQL = "Delete from SaleHeader where SID = " & Val(TxtSID.Text)
   cn.Execute vStrSQL
   cn.Execute ("Update SaleOrderHeader set IsSale = 0 Where OrderID = " & Val(TxtOrderID.Text) & "And OrderDate = '" & DtpOrderDate.DateValue & "' and StoreID = " & Val(TxtStoreID.Text))
      
   '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnDelete Then
   vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = ObjRegistry.PrefixPhoneNo + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               sSQL = ObjUserSecurity.UserName & " Deleted ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & " Time: " & Time & IIf(Val(TxtTotalDiscount.Caption) = 0, "", " Disc:" & TxtTotalDiscount.Caption) & vbCrLf & " NetAmt" & TxtNetAmount.Caption
               sSQL = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & sSQL & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               cn.Execute sSQL
            End If
         Next
   End If
   cn.CommitTrans
   
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub
Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchSale.ParaInBillDate = DtpBillDate.DateValue
   SchSale.Show vbModal
   If SchSale.ParaOutBillID <> -1 Then
      TxtSID.Text = SchSale.ParaOutSID
      TxtBillID.Text = SchSale.ParaOutBillID
      'Dim a
      'a = Split(SchSale.ParaOutBillDate, "/")
      DtpBillDate.DateValue = SchSale.ParaOutBillDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      TxtStoreID.Text = SchSale.ParaOutStoreID
      ''''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      cnSalePOS.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      GetSale
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub BtnPrint_Click()
'   On Error GoTo ErrorHandler
'
''   VStrSQL = " Select UserName, h.billid, h.BillDate, isnull(h.BillTime,0) as BillTime, h.Description, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, isnull(h.STax,0) as STax, isnull(h.cashReceived,0) as CashReceived, case when isdeadproduct = 1 then 'Book' else  p.ProductName end /*case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.billid, h.BillDate) end */ ProductName, unitname, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + b.qty as Qty, b.price/isnull(multiplier,1) as price, b.amount, b.DiscPC, b.DiscPer, b.DiscVal, isnull(b.SC,0) as SC, InvoiceNo, " & IIf(ObjRegistry.AllowUrduProduct = False, " isnull(Remarks,'')", " isnull(RemarksUrdu,'')") & " as Remarks" & vbCrLf _
'            + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName End as Customer, H.empid, isnull(EmpName,'') as EmpName, Cash, Credit, BankCard, b.ProductID, h.MemberID, isnull(cast(h.MemberID as varchar(6)) + '-' + MemberName,'') as MemberName, h.TableID, isnull(TableName,'') as TableName, null as DeliveryDate, isnull(h.DeliveryTime,0) as DeliveryTime, h.InvType, isnull(h.isPrinted,0) as isPrinted, b.Code" & vbCrLf _
'            + " from saleHeader h inner join salebody b on h.billid = b.billid and h.BillDate = b.BillDate" & vbCrLf _
'            + " inner join products p on p.productid = b.productid" & vbCrLf _
'            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
'            + " left outer join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
'            + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
'            + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
'            + " left outer join Members m on m.MemberID = h.MemberID" & vbCrLf _
'            + " left outer join Tables t on t.TableID = h.TableID " & vbCrLf _
'            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
'            + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate = '" & DtpBillDate.DateValue & "' Order By SerialNo"
'   VStrSQL = "  select h.BillID, h.BillDate, EntryDate, h.OrganizationID, OrganizationName, Customerid, isnull(Pr.PartyName,AccountName) + ' - ' + H.CustomerID as Customer_Name_ID," & vbCrLf _
'               + " pr.address, StoreName, BiltyNo, VehicleNo, h.Description," & vbCrLf _
'               + " Isnull(H.BillDiscPer, 0) BillDiscPer, Isnull(H.BillDisc,0) BillDisc, isnull(OtherCharges,0) as OtherCharges," & vbCrLf _
'               + " TotalAmount,  isnull(TotalExpense,0) as TotalExpense, b.ProductID as Code,  ProductName, dbo.FunSaleBodySerial(b.BillID,b.BillDate, b.ProductId) Serial," & vbCrLf _
'               + " dbo.FunSaleBodyOffer(b.BillID,b.BillDate, b.ProductId) ProductOffer, isnull(QtyPack,0)QtyPack, isnull(b.Multiplier,0)Multiplier, Qty," & vbCrLf _
'               + " Bonus,b.DiscPc, b.DiscPer, DiscVal, Offer, b.SaleTaxPer, SaleTaxval," & vbCrLf _
'               + " h.Empid, empname, price, Amount, previousAmount, CashReceived, b.RetailPrice, isnull(BatchNo,'') as BatchNo, BillNo," & vbCrLf _
'               + " Abbreviation + '/' + cast(b.Multiplier as varchar(10)) as packing, " & vbCrLf _
'               + " isnull( pr.Phone1  + ', ','') + isnull( pr.Phone2 + ', ','')  + isnull( pr.mobile + ', ','') +  isnull( pr.mobile2 + ', ','') as Moblie, packingname, pr.city" & vbCrLf _
'               + " from SaleBody b inner join products p on b.productid = p.productid" & vbCrLf _
'               + " inner join SaleHeader h on h.BillID = b.BillID and h.BillDate = b.BillDate" & vbCrLf _
'               + " LEFT OUTER JOIN packings pak on pak.packingid = b.packingid" & vbCrLf _
'               + " left outer join Organizations o on o.OrganizationID = h.OrganizationID" & vbCrLf _
'               + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
'               + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
'               + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
'               + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
'               + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate='" & DtpBillDate.DateValue & "'" & IIf(ObjRegistry.AllowOrderByCodeinInvoices, "Order By Code", "Order By SerialNo")
'
'   If ObjRegistry.LaserPrintofSaleInvoice = True Then
'      VStrSQL = "Select UserName, h.billid, h.BillDate, isnull(h.BillTime,0) as BillTime, h.Description, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, isnull(h.STax,0) as STax, isnull(h.cashReceived,0) as CashReceived, case when isdeadproduct = 1 then 'Book' else  p.ProductName end  /*case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.billid, h.BillDate) end */ ProductName, unitname, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + Qty as Qty, b.price/isnull(multiplier,1) as price, b.amount, b.DiscVal, InvoiceNo" & vbCrLf _
'            + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else h.CustomerID + ' - ' + AccountName End as Customer, isnull(pr.Address,'') as Address, Cash, Credit, BankCard, b.ProductID, PreviousAmount, isnull(OtherCharges,0) as OtherCharges, h.Empid, e.empname, dbo.FunSaleBodySerial(b.BillID,b.BillDate, b.ProductId) Serial, h.TableID, isnull(TableName,'') as TableName, null as DeliveryDate, isnull(h.isPrinted,0) as isPrinted," & IIf(ObjRegistry.AllowUrduProduct = False, " isnull(Remarks,'')", " isnull(RemarksUrdu,'')") & " as Remarks " & vbCrLf _
'            + " from saleHeader h inner join salebody b on h.billid = b.billid and h.BillDate = b.BillDate" & vbCrLf _
'            + " inner join products p on p.productid = b.productid" & vbCrLf _
'            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
'            + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
'            + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
'            + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
'            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
'            + " left outer join Tables t on t.TableID = h.TableID " & vbCrLf _
'            + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
'            + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate ='" & DtpBillDate.DateValue & "' Order By SerialNo"
'   End If
'
'   If RsReport.State = adStateOpen Then RsReport.Close
'   RsReport.Open VStrSQL, cnSalePOS, adOpenStatic, adLockReadOnly
'
'   If RsReport.RecordCount = 0 Then Exit Sub
'
'   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
'
'   If vLaserInvoice = True Then
'      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleInvoiceHalf1.rpt")
''      Set RptReportViewer.Report = New CrpSaleInvoiceHalf1
'      RptReportViewer.Report.PaperSize = crPaperA4
'      RptReportViewer.Report.PaperOrientation = crLandscape
'      RptReportViewer.Report.TopMargin = vY
'      RptReportViewer.Report.LeftMargin = vX
'      RptReportViewer.Report.RightMargin = 225
'   Else
'      If InStr(1, Printer.DeviceName, "CBM1000") > 0 Then
'         Set RptReportViewer.Report = New CrpSaleInvoiceCBM
'      ElseIf InStr(1, Printer.DeviceName, "AB-80K") > 0 Or InStr(1, Printer.DeviceName, "ARP-808K") > 0 Then
''         Set RptReportViewer.Report = New CrpSaleInvoiceAurora
'         Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleInvoiceAurora.rpt")
'         RptReportViewer.Report.LeftMargin = 225
'         RptReportViewer.Report.RightMargin = 0
'         RptReportViewer.Report.TopMargin = 255
'      ElseIf InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
'      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrptSaleInvoice.rpt")
''         Set RptReportViewer.Report = New CrpSaleInvoice
''         RptReportViewer.Report.TopMargin = 225
''         RptReportViewer.Report.LeftMargin = 225
''         RptReportViewer.Report.RightMargin = 225
'          RptReportViewer.Report.PaperOrientation = crPortrait
'      Else 'InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
'         Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleInvoiceAurora.rpt")
''         Set RptReportViewer.Report = New CrpSaleInvoiceAurora
'         RptReportViewer.Report.TopMargin = vY
'         RptReportViewer.Report.LeftMargin = vX
'         RptReportViewer.Report.RightMargin = 0
''         RptReportViewer.Report.BottomMargin = 100
'      End If
'      'RptReportViewer.Report.PaperOrientation = crPortrait
'    End If
'
'    RptReportViewer.Report.DiscardSavedData
'    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
'    RptReportViewer.Report.ReportTitle = "Sale Invoice"
'
'    'RptReportViewer.Report.LeftMargin = 0
'    'RptReportViewer.Report.RightMargin = 0
'
''   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
'
''   If ObjRegistry.PrintHeadersSaleInvoice = True Then
''      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
''      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
''      RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
''   Else
''      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ""
''      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ""
''      RptReportViewer.Report.ParameterFields(4).AddCurrentValue ""
''   End If
''   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy  'cnSalePOS.Execute("Select Name from Manufacturer").Fields(0).Value
''   RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.AddSpace = True, Left(".......................................", Val(ObjRegistry.BlankFooter)), "")
''   RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(ObjRegistry.CashReceived)
''   RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(ObjRegistry.Statement)
''
''   If vLaserInvoice = True Then
''      RptReportViewer.Report.ParameterFields(8).AddCurrentValue ""
''      RptReportViewer.Report.ParameterFields(9).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
''   Else
'''      RptReportViewer.Report.ParameterFields(8).AddCurrentValue CStr(IIf(IsNull(!ProdDesc1), "Description", !ProdDesc1))
''      RptReportViewer.Report.ParameterFields(8).AddCurrentValue IIf(ObjRegistry.AddSpace = True, Left(".......................................", Val(ObjRegistry.BlankFooter)), "")
''      RptReportViewer.Report.ParameterFields(9).AddCurrentValue IIf(ObjRegistry.PreviousBalanceVisible = True, ParaOutPrevious, 0)
''   End If
''
''   'RptReportViewer.Report.SelectPrinter "RASDD.DLL", "CBM1000 Partial Cut", "Com1" 'RptReportViewer.Report.SelectPrinter  "RASDD.DLL", "CBM1000 Partial Cut", "Com1"
''   cnSalePOS.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
''
''   RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
''   cnSalePOS.Execute "update SaleHeader set isPrinted = 1 where isnull(isPrinted,0) = 0 and BillID = " & Val(TxtBillID.Text) & " and BillDate ='" & DtpBillDate.DateValue & "'"
'
'   If ObjRegistry.LaserPrintofSaleInvoice = True Then
'      RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy  'cnSale.Execute("Select Name from Manufacturer").Fields(0).Value
'      RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
'      RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.AddSpace = True, ".", "")
'      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(ObjRegistry.CashReceived)
'      RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(ObjRegistry.Statement)
'      RptReportViewer.Report.ParameterFields(8).AddCurrentValue ""
'      RptReportViewer.Report.ParameterFields(9).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
'   Else
'      RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
'      RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
'      RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
'      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CStr(ObjRegistry.Statement)
'   End If
'   If ObjRegistry.PrintHeadersSaleInvoice = True Then
'      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
'      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
'   Else
'      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ""
'      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ""
'   End If
'   If ObjRegistry.PreviewSaleInoice Then
'      RptReportViewer.Show vbModal, Me
'   Else
'      RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
'   End If
'   cnSalePOS.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
'
'   'RptReportViewer.Show
'
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   
'   VStrSQL = " Select UserName, h.billid, h.BillDate, isnull(h.BillTime,0) as BillTime, h.Description, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, isnull(h.STax,0) as STax, isnull(h.cashReceived,0) as CashReceived, case when isdeadproduct = 1 then 'Book' else  p.ProductName end /*case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.billid, h.BillDate) end */ ProductName, p.productName1, unitname, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + b.qty as Qty, b.price/isnull(multiplier,1) as price, b.amount, b.DiscPC, b.DiscPer, b.DiscVal, isnull(b.SC,0) as SC, InvoiceNo, " & IIf(ObjRegistry.AllowUrduProduct = False, " isnull(Remarks,'')", " isnull(RemarksUrdu,'')") & " as Remarks" & vbCrLf _
            + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName End as Customer, H.empid, isnull(EmpName,'') as EmpName, Cash, Credit, BankCard, b.ProductID, h.MemberID, isnull(cast(h.MemberID as varchar(6)) + '-' + MemberName,'') as MemberName, h.TableID, isnull(TableName,'') as TableName, null as DeliveryDate, isnull(h.DeliveryTime,0) as DeliveryTime, h.InvType, isnull(h.isPrinted,0) as isPrinted, b.Code, p.ItemCode, right('00'+ cast(b.ColourID as varchar(2)),2) as ColourID, right('00'+ cast(b.SizeID as varchar(2)),2) as SizeID, ColourName, SizeName" & vbCrLf _
            + " from saleHeader h inner join salebody b on H.SID = B.SID " & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
            + " left outer join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
            + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
            + " left outer join Members m on m.MemberID = h.MemberID" & vbCrLf _
            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
            + " left outer join Tables t on t.TableID = h.TableID " & vbCrLf _
            + " Left outer join Colours Col on Col.Colourid = b.ColourID" & vbCrLf _
            + " Left Outer join Sizes Sz on Sz.SizeID = b.SizeID " & vbCrLf _
            + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate = '" & DtpBillDate.DateValue & "' and h.storeid = " & Val(TxtStoreID.Text) & " Order By SerialNo"
  
   
   vStrSQL = "Exec ProdPrintSalePos " & Val(TxtSID.Text)

   If InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "CANON") > 0 Or InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "HP") > 0 Or vLaserInvoice = True Then
      vStrSQL = "Select h.BillID, h.BillDate,Billtime, h.StoreID, UserName, ExpiryInvoice, EntryDate, h.OrganizationID, OrganizationName, Customerid, isnull(Pr.PartyName,AccountName) + ' - ' + H.CustomerID as Customer_Name_ID," & vbCrLf _
                + " Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName End as Customer, pr.address, LicenceNo, pr.Description PartyDescription,SectorName, ZoneName, StoreName, BiltyNo, VehicleNo, h.Description," & vbCrLf _
                + " Isnull(H.BillDiscPer, 0) BillDiscPer, Isnull(H.BillDisc,0) BillDisc, isnull(OtherCharges,0) as OtherCharges,  isnull(h.ServiceCharges,0) as ServiceCharges," & vbCrLf _
                + " TotalAmount,  isnull(TotalExpense,0) as TotalExpense,  CompanyName, P.GroupID, " & IIf(ObjRegistry.AllowUrduProduct = False, "GroupName", "GroupName1") & "  as GroupName, SubGroupName, BrandName, SeasonName, b.ProductID as Code, p.ProductName, p.ProductName1, dbo.FunSaleBodySerial(b.SID,b.BillDate, b.ProductId) Serial," & vbCrLf _
                + " dbo.FunSaleBodyOffer(b.BillID,b.BillDate,b.ProductId) ProductOffer, isnull(QtyPack,0)QtyPack, isnull(b.Multiplier,0)Multiplier, isnull(b.GrossQty,0)GrossQty, isnull(b.GrossUnit,0)GrossUnit, Qty," & vbCrLf _
                + " P.RetailPrice, P.PurPrice, Bonus,b.DiscPc, b.DiscPer, DiscVal, Cash, Credit, BankCard, isPrinted, Stax, Offer, Cast(b.Tradeoffer1 as varchar(5)) + ' + ' + cast(b.tradeoffer2 as varchar(5)) TradeOffer_12, tradevalue, Extraschemevalue, b.ExtraSchemePer," & vbCrLf _
                + " b.SaleTaxPer, SaleTaxval, AdvTaxVal, AdvTaxPer, ExtraTaxVal, ExtraTaxPer, h.CNIC, h.MobileNo,  b.SC, h.Empid, empname, price, Amount, previousAmount, CashReceived, isnull(h.BankAmount,0) as BankAmount, b.RetailPrice, isnull(BatchNo,'') as BatchNo, BillNo," & vbCrLf _
                + " Abbreviation + '/' + cast(b.Multiplier as varchar(10)) as packing," & vbCrLf _
                + " isnull( pr.Phone1  + ', ','') + isnull( pr.Phone2 + ', ','')  + isnull( pr.mobile + ', ','') +  isnull( pr.mobile2 + ', ','') as Moblie, packingname, pr.city, " & vbCrLf _
                + " AmountType = " & vbCrLf _
                + " CASE  " & vbCrLf _
                + " WHEN bankcard = 1 THEN ' Through Bank Card' " & vbCrLf _
                + " WHEN Cash = 1 THEN ' Through Cash' " & vbCrLf _
                + " WHEN Credit = 1 THEN ' Through Credit' " & vbCrLf _
                + " End "
                 vStrSQL = vStrSQL + " from SaleBody b inner join products p on b.productid = p.productid" & vbCrLf _
                + " inner join SaleHeader h on H.SID = B.SID" & vbCrLf _
                + " inner join users ur on ur.UserNo = h.UserNo Left Outer jOin companies cmp on cmp.companyid = p.companyid " & vbCrLf _
                + " Left Outer jOin Groups g on g.Groupid = p.Groupid" & vbCrLf _
                + " Left Outer jOin SubGroups sg on sg.subGroupid = p.subGroupid" & vbCrLf _
                + " Left Outer jOin Brands bd on bd.brandid = p.brandid" & vbCrLf _
                + " Left Outer jOin Seasons se on se.Seasonid = p.Seasonid" & vbCrLf _
                + " LEFT OUTER JOIN packings pak on pak.packingid = b.packingid" & vbCrLf _
                + " left outer join Organizations o on o.OrganizationID = h.OrganizationID" & vbCrLf _
                + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
                + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
                + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
                + " left outer join Sectors Sec on Sec.SectorID = Pr.SectorID" & vbCrLf _
                + " left outer join Zones Z on Z.ZoneID = Sec.ZoneID" & vbCrLf _
                + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
                + " where h.SID = " & Val(TxtSID.Text) & " and h.BillDate='" & DtpBillDate.DateValue & "'" & IIf(ObjRegistry.AllowOrderByCodeinInvoices, "Order By Code", "Order By SerialNo")
   End If
  
  
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
   
   If RsReport.RecordCount = 0 Then Exit Sub
   
'   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
   
   If vLaserInvoice = True Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleInvoiceHalf1.rpt")
      RptReportViewer.Report.PaperOrientation = crLandscape
      RptReportViewer.Report.TopMargin = vY
      RptReportViewer.Report.LeftMargin = vX
      RptReportViewer.Report.RightMargin = 225
   ElseIf InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "CANON") > 0 Or InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "HP") > 0 Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrptSaleInvoice.rpt")
   Else
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleInvoiceAurora.rpt")
         RptReportViewer.Report.TopMargin = 0
         RptReportViewer.Report.LeftMargin = 0
         RptReportViewer.Report.RightMargin = 0
    End If
    
    
    RptReportViewer.Report.DiscardSavedData
    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
    
    RptReportViewer.Report.ReportTitle = "Sale Invoice"
    
    'RptReportViewer.Report.LeftMargin = 0
    'RptReportViewer.Report.RightMargin = 0
    
   'RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
 
   If vLaserInvoice = True Then
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue vCompanyAddress & vCompanyCity
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
      RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CStr(ObjRegistry.Statement)
    ElseIf InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "CANON") > 0 Or InStr(1, StrConv(Printer.DeviceName, vbUpperCase), "HP") > 0 Then
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue vCompanyAddress & vCompanyCity
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
      RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CStr(ObjRegistry.Statement)
   Else
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue vCompanyAddress & vCompanyCity
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
      RptReportViewer.Report.ParameterFields(5).AddCurrentValue vAddSpace
      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(vCashReceived)
      RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(vStatement)
      RptReportViewer.Report.ParameterFields(8).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(9).AddCurrentValue (IIf(ObjRegistry.PreviousBalanceVisible = True, ParaOutPrevious, 0))
   End If

   If ObjRegistry.IsPortrait = False Then RptReportViewer.Report.PaperOrientation = crLandscape
   cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
   
   If vLaserInvoice = True Then
      RptReportViewer.Report.PaperOrientation = crLandscape
   End If
   
'   If ObjRegistry.PreviewSaleInoice Then
'      RptReportViewer.Show vbModal, Me
'   Else
      RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
'   End If
  
   cn.Execute "update SaleHeader set isPrinted = 1 where isnull(isPrinted,0) = 0 and BillID = " & Val(TxtBillID.Text) & " and BillDate ='" & DtpBillDate.DateValue & "'"
'   RptReportViewer.Show
    Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSaleOrder_Click()
   On Error GoTo ErrorHandler
   SchSaleOrder.ParaInOrderDate = DtpOrderDate.DateValue
   SchSaleOrder.Show vbModal
   If SchSaleOrder.ParaOutOrderID <> -1 Then
      TxtOrderID.Text = SchSaleOrder.ParaOutOrderID
      'Dim a
      'a = Split(SchSale.ParaOutBillDate, "/")
      DtpOrderDate.DateValue = SchSaleOrder.ParaOutOrderDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
'      cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetSaleOrder
      If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
'   If cnSalePOS.State = adStateClosed Then cnSalePOS.Open
   If vIsNewRecord = True And vChange = True And TxtOrderID.Text <> "" And vIsAdministrator = False And vIsEdit = False Then
      MsgBox "You are not authorized to modify a Sale Order.", vbCritical, "Error"
      Exit Sub
   End If
   If vIsNewRecord = False And vIsAdministrator = False And vIsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
   '''''''''''''''''''''''Check Organization'''''''''''''''''''''''''''''''''
   If vOrganizationMandatory = True And TxtOrganizationID.Text = "" Then
      MsgBox "Please Select Organization", vbInformation, Me.Caption
      If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
      Exit Sub
   End If
   '''''''''''''''''''''''Check Employee'''''''''''''''''''''''''''''''''
   If vEmployeeMandatory = True And TxtEmployeeID.Text = "" Then
      MsgBox "Please Select Employee", vbInformation, Me.Caption
      If TxtEmployeeID.Visible = True Then TxtEmployeeID.SetFocus
      Exit Sub
   End If
    '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
'    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where ToUserNo = " & vUser & " and Entrydate <='" & Date & "'"
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where ToUserNo = " & vUser
    With cn.Execute(vStrSQL)
        If .Fields(0).Value >= DtpBillDate.DateValue Then
            MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
    '''''''''''''''''''''''Check Entry Date'''''''''''''''''''''''''''''''''
    If visEntryDate = True Then
       If ObjRegistry.FromDate > Date Or ObjRegistry.ToDate < Date Then
         MsgBox "Data can not be saved Because Date is not set according to the Software's Entry date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
   '''''''''''''''''''''''Check Current Date'''''''''''''''''''''''''''''''''
    If vCurrentDateDataEntry = True And ObjUserSecurity.IsAdministrator = False Then
       If DtpBillDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
   If Grid.rows < 2 Then
      MsgBox "Please enter at least one product to sale", vbExclamation, "Alert"
     If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
     Exit Sub
   End If
   
   '''''''''''''''''''''''Check Printed Bill'''''''''''''''''''''''''''''''''
   If vNotEditingAfterPrinting = True Then
    vStrSQL = "Select isPrinted from saleheader where isprinted = 1 and SID = " & TxtSID.Text & " and billDate = '" & DtpBillDate.DateValue & "'"
    With cn.Execute(vStrSQL)
        If .RecordCount > 0 Then
            MsgBox "Data can not be edit becuase bill has been Printed ", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
   End If
   
'  Body Validation
'  validation has been performed when a row is added to the grid
   
'   RsBody.Filter = 0
'   If RsBody.RecordCount = 0 Then
'      MsgBox "Please enter at least one product to sale", vbExclamation, "Alert"
'      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
'      Exit Sub
'   End If

   'Header Validation
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      TxtStoreID.SetFocus
      Exit Sub
   End If
   If vEmployeeCommision = True Then
      If Trim(TxtEmployeeID.Text) = "" Then
         SubDestroyEmployeeCommision
      Else
         SubApplyEmployeeCommision
      End If
   End If
   
'   TxtTotalAmount.Caption = "0"
'   vTotalAmount = 0
'   With RsBody
'      .Filter = 0
'      .MoveFirst
'      For vCounter = 1 To .RecordCount
'         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + (!Qty * !Price)
'         vTotalAmount = vTotalAmount + Val(!Amount)
'         .MoveNext
'      Next vCounter
'   End With

'   Call SubCalculateFooter
  
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
'   ssql = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where ToUserNo = " & vUser & " and Entrydate <='" & Date & "'"
'   With cn.Execute(ssql)
'       If .Fields(0).Value >= DtpBillDate.DateValue Then
'           MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
'           Exit Sub
'       End If
'   End With

'   If DtpBillDate.Enabled = True Then
'      If OptCash.Visible Then OptCash.SetFocus
'      SubClearFields
'   End If

   TxtNetAmountCash.Text = SelfRound(TxtNetAmount.Caption)
   If ObjRegistry.CashReceived = True Then
      TxtCashReceivedCash.Text = SelfRound(TxtNetAmount.Caption)
   End If
   
   '''' Check Stock of Each Product not go to negative stock
If ObjRegistry.CheckStockOnSave = True Then
   With Grid
      .Redraw = False
      .MoveFirst
      For vCounter = 1 To .rows
         If Trim(.Columns("Productid").Text) <> "" Then
            vStrSQL = "select isnull(dbo.FunStock('" & .Columns("ProductID").Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
            If ObjRegistry.NegativeSale = False Then
               If vQtyLoose - Val(.Columns("Qty").Value) < 0 Then
                  MsgBox "Insufficient Stock Of " & .Columns("ProductName").Text & " ", vbInformation + vbOKOnly, "Error"
                  Exit Sub
               End If
            End If
         End If
      .MoveNext
      Next vCounter
      .Redraw = True
   End With
End If

vIsRemarksCompulsory = False
'''' Check Remarks Compulsory in Group
If ObjRegistry.RemarksVisible = True Then
   With Grid
      .Redraw = False
      .MoveFirst
      For vCounter = 1 To .rows
         If Trim(.Columns("Productid").Text) <> "" Then
            vStrSQL = "select isRemarksCompulsory from Groups G inner join Products P on p.GroupID = G.GroupID where productid = '" & .Columns("ProductId").Text & "' and isRemarksCompulsory = 1"
            If Not cn.Execute(vStrSQL).EOF Then
               vIsRemarksCompulsory = True
               If Trim(TxtRemarks.Text) = "" Then
                  MsgBox "Please Enter Remarks", vbInformation + vbOKOnly, "Error"
                  TxtRemarks.SetFocus
                  Exit Sub
               End If
            End If
         End If
      .MoveNext
      Next vCounter
      .Redraw = True
   End With
End If
   
   ChkPrint.Value = Abs(vAutoPrintinInvoices)
   
  
   
   'ParaInPrint = True
   'ParaInChoice = "Cash"
   Call SubEnable(False)
   Call SubFrameLoad
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub


Private Sub BtnOk_Click()
   On Error GoTo ErrorHandler
   BtnOk.Enabled = False
    
   If FunValidation = False Then BtnOk.Enabled = True: Exit Sub
   
   If OptCash.Value = True Then
      vContactNo = TxtCashCustomer.Text
   ElseIf OptCredit.Value = True Then
      vContactNo = TxtCustomerName.Text + ""
   Else
      vContactNo = TxtBankCustomer.Text
   End If
'   If TxtCashCustomer.Text <> "" Then
'        vContactNo = TxtCustomerName.Text
'   End If
   
   Call SubSave
   Call SubEnable(True)
   BtnOk.Enabled = True
   Frame1.Visible = False
   'If TxtEmployeeID.Visible And TxtEmployeeID.Enabled Then TxtEmployeeID.SetFocus Else
   If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateSyllabusToGrid()
    sSQL = " select b.ProductID, b.code, ProductName, 0 as packingID, Null as Multiplier, Null QtyPack, Null Bonus, Null Cost, " & vbCrLf _
             + " RetailPrice, IsWSDiscb4ST, IsWSSaleTax,  TokenVal, Null DiscPC, Null Offer, SaleTaxPer, Null SaleTaxval, " & vbCrLf _
             + " Price, QtyLoose Qty, Null DiscPer, 0 DiscPC, Null DiscVal, Amount From syllabusBody b left outer join products p on p.productid = b.productid where syllabusid =" & TxtSyllabusID.Text & " and isShow = 1"
       With cn.Execute(sSQL)
        Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         'TxtGrossAmount.Text = 0
         TxtTotalQty.Caption = 0
         TxtTotalItems.Caption = 0
         'TxtTotalDiscount.Caption = 0
         vTotDisc = 0
         vTotalAmount = 0
         TxtTotalAmount.Caption = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !Qty
            Grid.Columns("QtyOrigional").Value = !Qty
            Grid.Columns("Price").Value = !Price
            Grid.Columns("DiscPC").Value = 0 'IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = 0 'IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = 0 'IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = (Val(!Price) - IIf(IsNull(!DiscPC), 0, !DiscPC)) * Val(!Qty)
            Grid.Columns("IsProduct").Value = 1
            Grid.Columns("TotalAmount").Value = Val(!Price) * Val(!Qty)
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            Grid.Columns("EmpComm").Value = ""
            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
            TxtTotalItems.Caption = Val(TxtTotalItems.Caption) + 1
            'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(!DiscVal)
            vTotDisc = vTotDisc + 0 'Discval
            vTotalAmount = vTotalAmount + (Val(!Price) - IIf(IsNull(!DiscPC), 0, !DiscPC)) * Val(!Qty)
            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
            .MoveNext
         Wend
         .Close
      End With
      Call SubCalculateBody
      Grid.AddNew
      Grid.Columns("ProductID").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
'   End If
End Sub
Private Sub BtnSyllabus_Click()
   If FunSelectSyllabus(ssButton, False) = True Then
      TxtCode.SetFocus
   Else
      TxtSyllabusID.SetFocus
   End If
End Sub
Private Sub TxtSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridSerial.SetFocus
End Sub

Private Sub TxtSerial_LostFocus()
   GetDataFromTexBoxesToGridSerial
End Sub

Private Sub TxtSyllabusID_Change()
    If TxtSyllabusID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSyllabusID.Name Then Exit Sub
   If TxtSyllabusName.Text <> "" Then TxtSyllabusName.Text = ""
End Sub

Private Sub TxtSyllabusID_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler
    If TxtSyllabusName.Text <> "" Then Exit Sub
    If TxtSyllabusID.Text = "" Then Exit Sub
    Dim vTemp As Boolean
    vTemp = Not FunSelectSyllabus(ssValidate, True)
    If vTemp = True Then
        vTemp = Not FunSelectSyllabus(ssButton, False)
    End If
    Cancel = vTemp
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Function FunSelectSyllabus(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        FrmSyllabusSelection.Show vbModal, Me
        If FrmSyllabusSelection.ParaOutID = "" Then FunSelectSyllabus = False: Exit Function
        TxtSyllabusID.Text = FrmSyllabusSelection.ParaOutID
    End If
    '---------------------------
    vStrSQL = " Select * FROM syllabusheader where SyllabusID=" & Val(TxtSyllabusID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSyllabusName.Text = !SyllabusName
          FunSelectSyllabus = True
          .Close
          GetSyllabus
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSyllabus = False
          .Close
          TxtSyllabusID.Text = ""
          TxtSyllabusName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub GetSyllabus()
   On Error GoTo ErrorHandler
   sSQL = "select h.* from SyllabusHeader h Where h.SyllabusID=" & Val(TxtSyllabusID.Text)
   With cn.Execute(sSQL)
      If Not .BOF Then
          ' Ok
      End If
      .Close
   End With
   Call PopulateSyllabusToGrid
'   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub BtnEmployee_Click()
   On Error GoTo ErrorHandler
   If FunSelectEmployee(ssButton, False) = True Then
      If TxtMemberID.Visible And TxtMemberID.Enabled Then TxtMemberID.SetFocus Else TxtEmployeeID.SetFocus
   Else
      TxtEmployeeID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub TxtEmployeeID_Change()
   If TxtEmployeeID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtEmployeeID.Name Then Exit Sub
   If TxtEmployeeName.Text <> "" Then TxtEmployeeName.Text = ""
End Sub

Private Sub TxtEmployeeID_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If TxtEmployeeName.Text <> "" Then Exit Sub
    If TxtEmployeeID.Text = "" Then Exit Sub
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
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchEmployee.Show vbModal, Me
        If SchEmployee.ParaOutEmployeeID = "" Then FunSelectEmployee = False: Exit Function
        TxtEmployeeID.Text = SchEmployee.ParaOutEmployeeID
    End If
    '---------------------------
    If Trim(TxtEmployeeID.Text) = "" Then Exit Function
    sSQL = "Select *" & vbCrLf _
            + " from Employees" & vbCrLf _
            + " where isLockEmployee = 0 and EmpID = " & Val(TxtEmployeeID.Text)
    With cn.Execute(sSQL)
      If .RecordCount > 0 Then
        TxtEmployeeName.Text = !empname
        TxtCommission.Text = !Commission
        FunSelectEmployee = True
        .Close
        Exit Function
      Else
        FunSelectEmployee = False
        .Close
        TxtEmployeeID.Text = ""
        TxtEmployeeName.Text = ""
        TxtCommission.Text = ""
        Exit Function
      End If
    End With
Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function
Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      BtnOk.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

Private Sub TxtCustomerID_Change()
   If TxtCustomerID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   If TxtCustomerName.Text <> "" Then TxtCustomerName.Text = ""
End Sub

Private Sub TxtCustomerID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerName.Text <> "" Then Exit Sub
   If Trim(TxtCustomerID.Text) = "" Then Exit Sub
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

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.CmbFilter = "-- ALL PARENT ACCOUNTS --" '"Customers"
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo like '3%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select c.*, p.RefID, P.RefComm, isnull(isnull(p.mobile,p.mobile2),m.mobile) mobile FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " Left Outer join Members m on c.AccountNo = cast(m.Prefix as varchar(2))  + cast(m.MemberID as varchar(10)) " & vbCrLf & _
              " where p.BarCode = '" & (TxtCustomerID.Text) & "' or m.BarCode = '" & (TxtCustomerID.Text) & "' or (c.AccountNo = '" & (TxtCustomerID.Text) & "' and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo like '3%') and c.isDetailed = 1 and c.isLocked = 0)"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerID.Text = !AccountNo
          TxtCustomerName.Text = !AccountName + " " + IIf(IsNull(!Mobile), " ", !Mobile)
          TxtRefID.Text = IIf(IsNull(!RefID), "", !RefID)
          TxtRefComm.Text = IIf(IsNull(!RefComm), "", !RefComm)
          TxtPreviousReceivable.Text = cn.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtCustomerID.Text & "','" & DtpBillDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value

          vStrSQL = " Select isnull(Sum(round(B.TTLValue,0) - isnull(BillDisc,0) + isnull(OtherCharges,0) + Isnull(TotalExpense,0) + isnull(servicecharges,0) + isnull(STax,0)),0) as Amount " & vbCrLf _
                  + " FROM SaleHeader h INNER JOIN (Select SID, Sum(Amount) TTLValue FROM SaleBody Group By SID)b " & vbCrLf _
                  + " ON H.SID = B.SID " & vbCrLf _
                  + " where CustomerID = '" & (TxtCustomerID.Text) & "' and h.BillDate = '" & DtpBillDate.DateValue & "' and h.BillID >= " & Val(TxtBillID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
          TxtPreviousReceivable.Text = TxtPreviousReceivable.Text - cn.Execute(vStrSQL).Fields(0).Value

          FunSelectCustomer = True
          .Close
          Exit Function
      Else
          FunSelectCustomer = False
          .Close
          TxtCustomerID.Text = ""
          TxtCustomerName.Text = ""
          TxtRefID.Text = ""
          TxtRefComm.Text = ""
          TxtPreviousReceivable.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

'Private Sub Timer1_Timer()
'   cn.Execute "Select Getdate()"
'End Sub

Private Sub Timer2_Timer()
'   vDisplay = Space(20) + "abcdefghijklmnopqrstuvwxyz" + Space(20)
   vCounter = vCounter + 1
   If vCounter = Len(vDisplay) - 20 Then
      vCounter = 1
'   MSComm1.Output = Chr(27) & Chr(64)
      MSComm1.Output = Chr(CInt((&HB))) 'for home cursor
   End If
   If vDisplay <> "" Then
      MSComm1.Output = Mid(vDisplay, vCounter, 20) & "Total Bill" & Space(10 - Len("Rs." & TxtNetAmount.Caption)) & "Rs." & TxtNetAmount.Caption
   End If
End Sub

Private Sub TxtAmount_Change()
If TxtAmount.Visible = False Then Exit Sub
If ActiveControl.Name <> TxtAmount.Name Then Exit Sub
   If ObjRegistry.ChangeQtyOnChangedPrice = True Then
      vUnitPrice = Val(TxtPrice.Text)
      vAmount = Val(TxtAmount.Text) + Val(TxtDiscVal.Text) - (Round(vAmount / IIf(Val(vUnitPrice) = 0, IIf(vAmount = 0, 1, vAmount), Val(vUnitPrice)), 1) * Val(TxtSC.Text))
      TxtQty.Text = Round(vAmount / IIf(Val(vUnitPrice) = 0, IIf(vAmount = 0, 1, vAmount), Val(vUnitPrice)), 3)
      TxtActualAmount.Text = Val(TxtQty.Text) * (Val(TxtPrice.Text) + Val(TxtSC.Text))
      TxtTotalDiscount.Caption = Round(vTotDisc, 2)
      SubCalculateFooter
   End If
End Sub

Private Sub TxtBillDisc_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrorHandler
   If KeyAscii = 8 Then Exit Sub
   If (Val(TxtBillDisc.Text & Chr(KeyAscii)) + vTotDisc) > SelfRound(Val(TxtTotalAmount.Caption) + Val(TxtServiceCharges.Text) + Val(TxtSTax.Text)) Then
      KeyAscii = 0
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDiscVal_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrorHandler
   If KeyAscii = 8 Then Exit Sub
   If Val(TxtDiscVal.Text & Chr(KeyAscii)) > Val(TxtQty.Text) * (Val(TxtPrice.Text) + Val(TxtSC.Text)) Then
      KeyAscii = 0
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub FindRow()
   Dim vBm As Variant
   Dim lTotal As Long
   Dim i As Integer, vFind As String
   Dim vStart As Integer
   vBm = Grid.Bookmark
   
   If IsNumeric(TxtProductName.Text) = True Then
      If Len(TxtProductName.Text) <= 5 Then
         TxtProductName.Text = Right("00000" + CStr(Val(TxtProductName.Text)), 5)
         Grid.MoveFirst
         For i = 0 To Grid.rows - 1
            If (Grid.Columns("ProductID").CellValue(Grid.GetBookmark(i))) = TxtProductName.Text Then
               Grid.Bookmark = Grid.GetBookmark(i)
               Exit Sub
            End If
         Next i
      ElseIf Len(TxtProductName.Text) > 5 Then
         Grid.MoveFirst
         For i = 0 To Grid.rows - 1
            If (Grid.Columns("Code").CellValue(Grid.GetBookmark(i))) = TxtProductName.Text Then
               Grid.Bookmark = Grid.GetBookmark(i)
               Exit Sub
            End If
         Next i
      End If
   Else
      If Grid.Columns("Code").Text = "" Then
         Grid.MoveFirst
         vStart = 0
      Else
         vStart = 1
      End If
      For i = vStart To Grid.rows
         If UCase(Grid.Columns("ProductName").CellValue(Grid.GetBookmark(i))) Like "*" & UCase(TxtProductName.Text) & "*" Then
            Grid.Bookmark = Grid.GetBookmark(i)
            Exit Sub
         End If
      Next i
   End If
   Grid.Bookmark = vBm
End Sub



Private Sub SubMakePackageDeal()
   Dim RsTemp As New ADODB.Recordset
   'Grid.Redraw = False
   vBm = Grid.Bookmark
   Grid.MoveFirst
   sSQL = " select * " & vbCrLf _
         + " from PackageDealInfoBody b inner join PackageDealInfoHeader h on h.id = b.id"
   With cn.Execute(sSQL)
      Grid.MoveFirst
      While Grid.Columns("ProductID").Text <> ""
         .Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
         If .RecordCount > 0 Then
            RsDetail.AddNew
            RsDetail!Productid = Grid.Columns("ProductID").Text
            RsDetail!Rate = Grid.Columns("Price").Text
            RsDetail!QtyLoose = Grid.Columns("Qty").Text
            RsDetail!Amount = Grid.Columns("Amount").Text
            RsDetail.Update
            RsBody.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
            If RsBody.RecordCount > 0 Then RsBody.Delete
            Grid.SelBookmarks.RemoveAll
            Grid.SelBookmarks.Add Grid.Bookmark
            Grid.DeleteSelected
         Else
            Grid.MoveNext
         End If
      Wend
      .Filter = "ProductID = '" & RsDetail!Productid & "'"
      If .RecordCount > 0 Then
         If RsTemp.State = adStateOpen Then RsTemp.Close
         vStrSQL = " SELECT p.productid, ProductName, RetailPrice, DiscPer, DiscPC, EmpComm, ServiceCharges" & vbCrLf _
               + " from PackageDealInfoHeader un inner join Products p on un.PackageDealid = p.productid" & vbCrLf _
               + " where p.productid = '" & !PackageDealID & "'"
         
         RsTemp.Open vStrSQL, cn, adOpenDynamic, adLockReadOnly
         If RsTemp.RecordCount > 0 Then
            TxtCode.Text = RsTemp!Productid
            TxtPID.Text = RsTemp!Productid
            TxtProductName.Text = RsTemp!ProductName
            TxtPrice.Text = RsTemp!RetailPrice
            TxtQty.Text = RsDetail!QtyLoose
            TxtCost.Text = 0
            TxtDiscPC.Text = IIf(IsNull(RsTemp!DiscPC), 0, RsTemp!DiscPC)
            TxtDiscPer.Text = IIf(IsNull(RsTemp!DiscPer), 0, RsTemp!DiscPer)
            TxtSC.Text = IIf(IsNull(RsTemp!ServiceCharges), 0, RsTemp!ServiceCharges)
            TxtEmpComm.Text = IIf(IsNull(RsTemp!EmpComm), 0, RsTemp!EmpComm)
            If Val(TxtDiscPC.Text) <> 0 Then
               TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
            End If
            ChkIsProduct.Value = 0
            SubCalculateBody
            Grid.MoveLast
            GetDataFromTexBoxesToGrid
         End If
      End If
      .Close
   End With
   
   'RsDetail.Filter = 0
   'Grid.Bookmark = vBm
   'Grid.Redraw = True
End Sub

Private Sub SubApplyEmployeeCommision()
   On Error GoTo ErrorHandler
   Grid.Redraw = False
   Grid.MoveFirst
   sSQL = " select * " & vbCrLf _
         + " from Products"
   With cn.Execute(sSQL)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         Grid.Columns("EmpComm").Value = IIf(IsNull(!EmpComm), "", !EmpComm)
         Grid.MoveNext
      Wend
      .Close
   End With
   Grid.MoveLast
   Grid.Redraw = True
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubDestroyEmployeeCommision()
   On Error GoTo ErrorHandler
   Grid.Redraw = False
   Grid.MoveFirst
'   ssql = " select * " & vbCrLf _
         + " from Products"
   
   For vCounter = 1 To Grid.rows
      If Trim(Grid.Columns("ProductID").Text) <> "" Then
'         .Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
'         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
'            RsBody.Filter = "ProductID='" & !Productid & "'"
            Grid.Columns("EmpComm").Value = 0
'            RsBody!EmpComm = Null
'         End If
    End If
    Grid.MoveNext
    Next vCounter
   
   Grid.Redraw = True
   Grid.MoveLast
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubApplyMember()
   On Error GoTo ErrorHandler
   Dim vAmount, vDiscVal As Double
   Grid.MoveFirst
   sSQL = " select * " & vbCrLf _
         + " from MembersDiscount "
   With cn.Execute(sSQL)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
         If .RecordCount > 0 Then
'            vDiscVal = Val(Grid.Columns("DiscVal").Value)
            'GetDataBackFromGridToTexBoxes
'            RsBody.Filter = "ProductID='" & !Productid & "'"
            vDiscVal = Val(Grid.Columns("DiscVal").Value)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
            Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("Price").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
            'Grid.Columns("SC").Value = IIf(IsNull(!Sc), 0, !Sc)
            vAmount = Val(Grid.Columns("Amount").Value)
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) + Val(Grid.Columns("SC").Value) - Val(Grid.Columns("DiscVal").Value)
            
            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - vAmount + Val(Grid.Columns("Amount").Text)
            vTotDisc = vTotDisc - vDiscVal + Val(Grid.Columns("DiscVal").Text)
            vTotalAmount = vTotalAmount - vAmount + Val(Grid.Columns("Amount").Text)
            
'            RsBody!DiscPC = Val(Grid.Columns("DiscPC").Value)
'            RsBody!DiscPer = Val(Grid.Columns("DiscPer").Value)
'            RsBody!DiscVal = Val(Grid.Columns("DiscVal").Value)
'            RsBody!Amount = Val(Grid.Columns("Amount").Value)
     
         End If
         Grid.MoveNext
      Wend
      .Close
   End With
   Grid.MoveLast
'   TxtBillDisc.Text = Val(TxtBillDisc.Text) + vTotDisc
'   TxtBillDiscPer.Text = Val(TxtBillDisc.Text) / IIf(Val(TxtTotalAmount.Caption) = 0, 1, Val(TxtNetAmount.Caption)) * 100
   SubCalculateFooter
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubDestroyMember()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   sSQL = " select * " & vbCrLf _
         + " from MembersDiscount "
   With cn.Execute(sSQL)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
            
''            RsBody.Filter = "ProductID='" & !Productid & "'"
''            Grid.Columns("DiscPer").Value = 0 'IIf(IsNull(!DiscPer), 0, !DiscPer)
''            Grid.Columns("DiscPC").Value = 0 'Round((Val(RsBody!Price) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
''            Grid.Columns("DiscVal").Value = 0 'Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
''            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) + Val(Grid.Columns("SC").Value) - Val(Grid.Columns("DiscVal").Value)
'
''            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - RsBody!Amount + Val(Grid.Columns("Amount").Text)
''            vTotDisc = vTotDisc - RsBody!DiscVal + Val(Grid.Columns("DiscVal").Text)
''            vTotalAmount = vTotalAmount - RsBody!Amount + Val(Grid.Columns("Amount").Text)
            
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) + Val(Grid.Columns("SC").Value)
            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - Val(Grid.Columns("Amount").Text)
            vTotDisc = vTotDisc - Val(Grid.Columns("DiscVal").Text)
            vTotalAmount = vTotalAmount - Val(Grid.Columns("Amount").Text)
            
            Grid.Columns("DiscPer").Value = 0
            Grid.Columns("DiscPC").Value = 0 'Round((Val(RsBody!Price) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = 0 'Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
            
'
            
'            RsBody!DiscPC = Val(Grid.Columns("DiscPC").Value)
'            RsBody!DiscPer = Val(Grid.Columns("DiscPer").Value)
'            RsBody!DiscVal = Val(Grid.Columns("DiscVal").Value)
'            RsBody!Amount = Val(Grid.Columns("Amount").Value)
         End If
         Grid.MoveNext
      Wend
      .Close
   End With
   SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnMember_Click()
   On Error GoTo ErrorHandler
   If FunSelectMember(ssButton, False) = True Then
      If TxtEmployeeID.Enabled And TxtEmployeeID.Visible Then TxtEmployeeID.SetFocus Else TxtCode.SetFocus
   Else
      TxtMemberID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub TxtMemberID_Change()
   If ActiveControl.Name <> TxtMemberID.Name Then Exit Sub
   If TxtMemberName.Text <> "" Then TxtMemberName.Text = "": TxtMemberBarCode.Text = "": Call SubDestroyMember
End Sub

Private Sub TxtMemberID_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If TxtMemberName.Text <> "" Then Exit Sub
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

Private Sub SubSave()
   On Error GoTo ErrorHandler
'   If DtpBillDate.Enabled And DtpBillDate.Date <> IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date)) And DateFlag = True Then
'      If MsgBox("Are you sure to Change Bill Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
'         DtpBillDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))
'         TxtBillID.Text = FunGetMaxID()
'      End If
'      DateFlag = False
'   End If

'''''' Check Multiple Store
  RsBodyStore.Filter = 0
   If RsBodyStore.RecordCount > 0 Then
      MsgBox "Data cannot be saved because This invoice inlcude Muliple Store ", vbCritical, "Alert"
      Exit Sub
   End If
   
  'Saving record
  
   cn.BeginTrans
   
'   If vIsNewRecord = True Then
'      If cn.Execute("Select * from SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "' and StoreID = " & Val(TxtStoreID.Text) & " --and StampID <> " & TxtStampID.Text).RecordCount > 0 Then
'         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
'         TxtBillID.Text = FunGetMaxID
'         'Exit Sub
'      End If
'   End If
    
'   Call SubLastEntryDate(DtpBillDate.DateValue)
   
   If vIsNewRecord = False Then Call ActivityLogSale("Sale Invoice", eEdit, TxtBillID.Text, DtpBillDate.DateValue)
   ''''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Call UserActivities
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  ''''''''''''''''' Get Commision from commisionDisc if not exists commision in employee
If Trim(TxtEmployeeID.Text) <> "" Then
If cn.Execute("Select commission from employees where EmpID = " & TxtEmployeeID.Text).Fields(0) = 0 Then
  TxtAvgDisc.Text = Round(TxtTotalDiscount.Caption / TxtTotalAmount.Caption * 100, 3)
  sSQL = "Select * from commisionDisc Where " & Val(TxtAvgDisc.Text) & " >= DiscPerFrom and " & Val(TxtAvgDisc.Text) & " <= DiscPerTo"
   With cn.Execute(sSQL)
'      TxtAvgDisc.Text = Round(TxtTotalDiscount.Caption / TxtNetAmount.Caption * 100, 3)
      If .RecordCount <> 0 Then
         TxtCommission.Text = !Commision
         TxtRemarks.Text = !CommisionName
      End If
   End With
End If
End If
 '''''''''''''''''''''''''''''''''''''''''''
 
   
'' Sale Header

vNow = vDate & " " & Format(IIf(vSystemDate = True, Now, cn.Execute("Select getdate()").Fields(0).Value), "hh:mm:ss")

Dim vInvoiceNo, vComission, vBankMachineID, vCashReceived, vCustomerID, vCustomerName As String
      If OptBankCard.Value = True Then
         vInvoiceNo = TxtInvoiceNo.Text
         vComission = TxtCommision.Text
         vBankMachineID = TxtBankMachineID.Text
         vCashReceived = Val(TxtCashReceivedBank.Text)
         vCustomerID = "621"
         vCustomerName = IIf(Trim(TxtBankCustomer.Text) = "", Null, TxtBankCustomer.Text)
         TxtBankAmount.Text = ""
      End If
      If OptCash.Value = True Then
         vComission = "''"
         vInvoiceNo = Null
         vBankMachineID = "''"
         vCashReceived = Val(TxtCashReceivedCash.Text)
         vCustomerID = "621"
         vCustomerName = IIf(Trim(TxtCashCustomer.Text) = "", Null, TxtCashCustomer.Text)
         TxtBankAmount.Text = ""
      End If
      If OptCredit.Value = True Then
         vComission = Val(TxtCommision.Text)
         vInvoiceNo = Null
         vBankMachineID = IIf(Trim(TxtBankMachineCreditID.Text) = "", "''", TxtBankMachineCreditID.Text)
         vCashReceived = Val(TxtCashReceivedCredit.Text)
         If Val(TxtBankMachineCreditID.Text) > 1 Then vComission = TxtCommision.Text
         vCustomerID = TxtCustomerID.Text
         vCustomerName = TxtCustomerName.Text
      End If
      
vStrPara = ""
vStrPara = Abs(ObjRegistry.AllowContinuousBillNo) & ","
vStrPara = vStrPara & Abs(ObjRegistry.AllowMonthlyBillNo) & ","
vStrPara = vStrPara & Abs(ObjRegistry.AllowDailyBillNo) & "," 'AllowDailyBillNo
vStrPara = vStrPara & Val(TxtSID.Text) & "," 'SID
vStrPara = vStrPara & TxtBillID.Text & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
vStrPara = vStrPara & "'" & vCustomerID & "'," 'CustomerID
'vStrPara = vStrPara & SelfRound(vTotalAmount) & "," ' Total Amount
vStrPara = vStrPara & SelfRound(TxtNetAmount.Caption + Val(TxtBillDisc.Text) - Val(TxtServiceCharges.Text) - Val(TxtSTax.Text)) & ","    ' Total Amount
vStrPara = vStrPara & Val(TxtBillDisc.Text) & "," 'BillDisc
vStrPara = vStrPara & vCashReceived & "," ' 'CashReceived
vStrPara = vStrPara & vUser & "," 'UserNo
vStrPara = vStrPara & TxtStoreID.Text & "," 'StoreID
vStrPara = vStrPara & IIf(OptBankCard.Value = True, 1, 0) & "," 'BankCard
vStrPara = vStrPara & IIf(OptCredit.Value = True, 1, 0) & "," 'Credit
vStrPara = vStrPara & IIf(OptCash.Value = True, 1, 0) & "," 'Cash
vStrPara = vStrPara & "" & vBankMachineID & "," 'BankMachineID
vStrPara = vStrPara & "'" & vInvoiceNo & "',"  'InvoiceNo
vStrPara = vStrPara & "'" & vCustomerName & "'," 'CustomerName
vStrPara = vStrPara & Val(TxtBillDiscPer.Text) & "," 'BillDiscPer
vStrPara = vStrPara & vComission & ","   'Commision
vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtCommission.Text)) & "," 'EmpComm
vStrPara = vStrPara & "'" & IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text)) & "'," 'EmpID
vStrPara = vStrPara & 0 & "," 'isReplace
vStrPara = vStrPara & 0 & "," 'isPosted
vStrPara = vStrPara & IIf(Trim(TxtMemberID.Text) = "", "''", TxtMemberID.Text) & "," 'MemberID
vStrPara = vStrPara & "'" & vNow & "'," 'BillTime
vStrPara = vStrPara & "'" & vIsNewRecord & "'," 'Tag
vStrPara = vStrPara & "'" & IIf(Trim(TxtManualBillNo.Text) = "", Null, TxtManualBillNo.Text) & "'," 'ManualBillNo
vStrPara = vStrPara & "'" & IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text) & "',"  'Remarks
vStrPara = vStrPara & IIf(Trim(TxtOrganizationID.Text) = "", "''", TxtOrganizationID.Text) & ","  'OrganizationID
vStrPara = vStrPara & "'" & Null & "'," ' BillNo
vStrPara = vStrPara & "'" & Null & "'," ' Bilty No
vStrPara = vStrPara & "'" & Null & "'," 'Description
vStrPara = vStrPara & "''" & "," 'PAIDAMOUNT
vStrPara = vStrPara & "'" & Null & "',"  'EntryDate
vStrPara = vStrPara & IIf(OptCredit.Value = True, Val(TxtPreviousReceivable.Text), 0) & "," 'PreviousAmount
vStrPara = vStrPara & 0 & "," 'OtherCharges
vStrPara = vStrPara & "'" & Null & "'," 'SaleManID
vStrPara = vStrPara & 0 & "," 'TotalExpense
vStrPara = vStrPara & IIf(Val(TxtOrderID.Text) = 0, "''", TxtOrderID.Text) & "," 'OrderID
vStrPara = vStrPara & "'" & DtpOrderDate.DateValue & "'," 'OrderDate
vStrPara = vStrPara & 0 & "," 'Freight
vStrPara = vStrPara & 0 & "," 'IsCustomerFreight
vStrPara = vStrPara & "'" & Null & "'," 'VechicleNo
vStrPara = vStrPara & IIf(TxtServiceCharges.Text = "", "''", Val(TxtServiceCharges.Text)) & "," 'ServiceCharges
vStrPara = vStrPara & IIf(TxtServiceChargesPer.Text = "", "''", Val(TxtServiceChargesPer.Text)) & "," 'ServiceChargesPer
vStrPara = vStrPara & IIf(TxtSTax.Text = "", "''", Val(TxtSTax.Text)) & "," 'STax
vStrPara = vStrPara & IIf(TxtSTaxPer.Text = "", "''", Val(TxtSTaxPer.Text)) & "," 'STaxPer
vStrPara = vStrPara & "'" & IIf(Trim(TxtTableID.Text) = "", Null, TxtTableID.Text) & "'," 'TableID
vStrPara = vStrPara & "'" & Now & "'," 'ServerEntry
vStrPara = vStrPara & "'" & IIf(CmbType.Visible = False, Null, CmbType.Text) & "'," 'InvType
vStrPara = vStrPara & "'" & DtpDeliveryDate.DateValue & "'," 'DeliveryDate
vStrPara = vStrPara & "'" & DTPDeliveryTime.Value & "'," 'DeliveryTime
vStrPara = vStrPara & "'" & Null & "'," 'isPrinted
vStrPara = vStrPara & "'" & Null & "'," 'RemarksUrdu
'vStrPara = vStrPara & "Default" & ","  'StampID
vStrPara = vStrPara & 0 & "," 'isTransfer
vStrPara = vStrPara & IIf(DtpPromiseDate.DateValue = Empty, "Null", "'" & DtpPromiseDate.DateValue & "'") & "," 'PromiseDate
vStrPara = vStrPara & "Null," 'Expiry Invoice
vStrPara = vStrPara & "Null," 'Syllabus
vStrPara = vStrPara & "'" & IIf(Trim(vSessionID) = 0, Null, Val(vSessionID)) & "',"  'vSessionID
vStrPara = vStrPara & IIf(TxtAdvTaxVal.Text = "", "''", Val(TxtAdvTaxVal.Text)) & "," 'AdvTaxVal
vStrPara = vStrPara & IIf(TxtAdvTaxPer.Text = "", "''", Val(TxtAdvTaxPer.Text)) & "," 'AdvTaxPer
vStrPara = vStrPara & IIf(TxtExtraTaxVal.Text = "", "''", Val(TxtExtraTaxVal.Text)) & "," 'ExtraTaxVal
vStrPara = vStrPara & IIf(TxtExtraTaxPer.Text = "", "''", Val(TxtExtraTaxPer.Text)) & "," 'ExtraTaxPer
vStrPara = vStrPara & "'" & IIf(Trim(TxtCNIC.Text) = "", Null, TxtCNIC.Text) & "',"  'CNIC
vStrPara = vStrPara & "'" & IIf(Trim(TxtCellNo.Text) = "", Null, TxtCellNo.Text) & "',"  'CellNo
vStrPara = vStrPara & Val(TxtSumDiscAmount.Text) & "," 'Sum Disc Amount
vStrPara = vStrPara & "Null," 'DispatchDate
vStrPara = vStrPara & "Null," 'Terms
vStrPara = vStrPara & "'" & IIf(Trim(TxtRefID.Text) = "", Null, TxtRefID.Text) & "',"  'RefID
vStrPara = vStrPara & "'" & IIf(Trim(TxtRefComm.Text) = "", Null, TxtRefComm.Text) & "',"  'Refcomm
vStrPara = vStrPara & Val(TxtBankAmount.Text) 'Bank Amount in Credit Option
vStrPara = Replace(vStrPara, "''", "Null")

vStrPara = "DECLARE @returnvalue INT EXEC @returnvalue = saleheaderinsert " & vStrPara & " Select @returnvalue"
   vMasterID = cn.Execute(vStrPara).Fields(0).Value
   TxtSID.Text = vMasterID
'   MsgBox vMasterID
   

''' insert Sale Body
vStrDetail = ""

With Grid
 .Redraw = False
 .MoveFirst
   For vCounter = 1 To .rows
      If Trim(.Columns("Productid").Text) <> "" Then
      
      ''''''''''''''''''''''''''''
 vStrPara = ""
TxtBillID.Text = cn.Execute("Select billID from Saleheader where SID = " & vMasterID).Fields(0).Value
vStrPara = vStrPara & "'" & vUpdateStock & "'," 'check stock update or not
vStrPara = vStrPara & vMasterID & ","
vStrPara = vStrPara & TxtBillID.Text & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
'vStrPara = vStrPara & .Columns("SerialNo").Text & ","
'vStrPara = vStrPara & .Columns("BillID").Text & ","
'vStrPara = vStrPara & .Columns("BillDate").Text & ","
vStrPara = vStrPara & "'" & .Columns("ProductID").Text & "',"
vStrPara = vStrPara & .Columns("Qty").Text & ","
vStrPara = vStrPara & .Columns("Price").Text & ","
vStrPara = vStrPara & .Columns("DiscPC").Text & ","
vStrPara = vStrPara & .Columns("Amount").Text & ","
vStrPara = vStrPara & "'" & .Columns("Code").Text & "',"
vStrPara = vStrPara & .Columns("DiscPer").Text & ","
vStrPara = vStrPara & .Columns("DiscVal").Text & ","

vStrPara = vStrPara & 0 & "," ' isDiscB4TradeOffer
vStrPara = vStrPara & 0 & ","   'isDiscB4ExtraScheme
vStrPara = vStrPara & 0 & "," 'isDiscB4SaleTax
vStrPara = vStrPara & "''" & ","  'TradeOffer1
vStrPara = vStrPara & "''" & ","   'TradeOffer2
vStrPara = vStrPara & "''" & ","   'ExtraSchemePer
vStrPara = vStrPara & "''" & ","   'TradeValue
vStrPara = vStrPara & "''" & ","   'ExtraSchemeValue

vStrPara = vStrPara & .Columns("Cost").Text & ","
vStrPara = vStrPara & .Columns("isProduct").Text & ","
vStrPara = vStrPara & "''" & "," ' Pack Name
vStrPara = vStrPara & "''" & "," ' Qty Pack
vStrPara = vStrPara & "''" & "," ' Pack
vStrPara = vStrPara & "''" & "," ' Bonus
vStrPara = vStrPara & "''" & "," 'Offer
vStrPara = vStrPara & "''" & "," 'SaleTaxPer
vStrPara = vStrPara & "''" & "," ' SaleTaxVal
vStrPara = vStrPara & Val(.Columns("TokenVal").Text) & ","
vStrPara = vStrPara & Val(TxtPrice.Text) & "," 'RetailPrice
vStrPara = vStrPara & 0 & "," 'IsWSSaleTax
vStrPara = vStrPara & 0 & "," 'IsRetailSaleTax
vStrPara = vStrPara & 0 & "," 'IsWSDiscb4ST
vStrPara = vStrPara & Val(.Columns("SC").Text) & "," 'SC
vStrPara = vStrPara & Val(TxtEmpComm.Text) & "," 'EmpComm
vStrPara = vStrPara & "''" & "," 'BatchNo
'vStrPara = vStrPara & "''" & "," 'StampID
vStrPara = vStrPara & TxtStoreID.Text & ","                  'StoreID
If ObjRegistry.AllowEmployeProductWise Then
   vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtEmployeeID.Text)) & "," 'EmpID
Else
   vStrPara = vStrPara & "''" & "," 'EmpID
End If
vStrPara = vStrPara & "'" & IIf(Trim(.Columns("ColourID").Text) = "", Null, Val(.Columns("ColourID").Text)) & "'," ' ColourID
vStrPara = vStrPara & "'" & IIf(Trim(.Columns("SizeID").Text) = "", Null, Val(.Columns("SizeID").Text)) & "'," ' SizeID
vStrPara = vStrPara & "null" & ","  'Gross Qty
vStrPara = vStrPara & "null" & ","  'Gross Unit
If ObjRegistry.AllowStoreProductWise Then
   vStrPara = vStrPara & "'" & IIf(Trim(.Columns("ColourID").Text) = "", Null, Val(.Columns("ColourID").Text)) & "'," 'HeaderStoreID
Else
   vStrPara = vStrPara & "''," 'HeaderStoreID
End If
vStrPara = vStrPara & .Columns("DiscAmount").Value & "," ' Disc Amount
vStrPara = vStrPara & "Null" & "," ' isLastPrice
vStrPara = vStrPara & "Null" & ","   'Re SPrice
vStrPara = vStrPara & "Null" & ""   'Re SAmount
vStrPara = Replace(vStrPara, "''", "Null")
cn.Execute ("Exec SaleBodyInsert " & vStrPara)
vStrDetail = vStrDetail & " (P" & .Columns("ProductID").Text & " Q" & .Columns("Qty").Text & " A" & .Columns("Amount").Text & ")"
      ''''''''''''''''''''''''''''
      
      End If
      .MoveNext
   Next vCounter
   .RemoveAll
   .Redraw = True
   End With
   
'''''' Sale Body Serial
 
 With GridSerial
 .Redraw = False
 .MoveFirst
  For vCounter = 1 To .rows
      If Trim(.Columns("Productid").Text) <> "" Then
      
      ''''''''''''''''''''''''''''
 vStrPara = ""
vStrPara = vStrPara & vMasterID & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
vStrPara = vStrPara & "'" & .Columns("ProductID").Text & "',"
vStrPara = vStrPara & "'" & .Columns("Serial").Text & "'"

vStrPara = Replace(vStrPara, "''", "Null")
cn.Execute ("Exec SaleBodySerialInsert " & vStrPara)

      ''''''''''''''''''''''''''''
      
      End If
      .MoveNext
   Next vCounter
   .RemoveAll
   .Redraw = True
   End With
   
'    ssql = "select * from SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
'    Dim Rs As New ADODB.Recordset
'    With Rs
'      .Open ssql, cn, adOpenDynamic, adLockPessimistic
'      If .BOF Then
'         .AddNew
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         !OrderID = IIf(Val(TxtOrderID.Text) = 0, Null, TxtOrderID.Text)
'         !OrderDate = DtpOrderDate.DateValue
''         !StampID = TxtStampID.Text
'         !BillTime = Now
'         !UserNo = vUser
'      End If
'      !isReplace = 0
'      !isPosted = 0
'      !isTransfer = 0
'      !InvType = IIf(CmbType.Visible = False, Null, CmbType.Text)
'      !StoreID = Val(TxtStoreID.Text)
'      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
'      !TableId = IIf(Trim(TxtTableID.Text) = "", Null, TxtTableID.Text)
'      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
'      !EmpComm = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text))
'      !MemberID = IIf(Trim(TxtMemberID.Text) = "", Null, TxtMemberID.Text)
'      !TotalAmount = SelfRound(vTotalAmount)
'      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
'      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
'      !ServiceCharges = IIf(TxtServiceCharges.Text = "", Null, Val(TxtServiceCharges.Text))
'      !ServiceChargesPer = IIf(TxtServiceCharges.Text = "", Null, Val(TxtServiceCharges.Text))
'      !STax = IIf(TxtSTax.Text = "", Null, Val(TxtSTax.Text))
'      !STaxPer = IIf(TxtSTax.Text = "", Null, Val(TxtSTax.Text))
'      !DeliveryDate = DtpDeliveryDate.DateValue
'      !DeliveryTime = DTPDeliveryTime.Value
'      If OptBankCard.Value = True Then
'         !InvoiceNo = TxtInvoiceNo.Text
'         !Commision = TxtCommision.Text
'         !BankMachineID = TxtBankMachineID.Text
'         !CashReceived = Val(TxtCashReceivedBank.Text)
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(TxtBankCustomer.Text) = "", Null, TxtBankCustomer.Text)
'      End If
'      If OptCash.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !BankMachineID = Null
'         !CashReceived = Val(TxtCashReceivedCash.Text)
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(TxtCashCustomer.Text) = "", Null, TxtCashCustomer.Text)
'      End If
'      If OptCredit.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !BankMachineID = Null
'         !CashReceived = Val(TxtCashReceivedCredit.Text)
'         !CustomerID = TxtCustomerID.Text
'         !CustomerName = Null
'      End If
'      !BankCard = OptBankCard.Value
'      !Cash = OptCash.Value
'      !Credit = OptCredit.Value
'      '!Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
'      !Remarks = IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text)
'      !RemarksUrdu = IIf(Trim(TxtRemarksUrdu.Text) = "", Null, TxtRemarksUrdu.Text)
'      !ManualBillNo = IIf(Trim(TxtManualBillNo.Text) = "", "", TxtManualBillNo.Text)
''      .Update
'      .Close
'   End With
'
'   With RsBody
'      .Filter = 0
'      .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
''         !StampID = TxtStampID.Text
'         .MoveNext
'      Next vCounter
''      .UpdateBatch
'   End With
   
'   With RsDetail
'      .Filter = 0
'      If .RecordCount > 0 Then .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         .MoveNext
'      Next vCounter
''      .UpdateBatch
'   End With
   
   sSQL = " select sob.ProductID, ProductName, (isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty - isnull(uqty,0) as Qtyloose, sob.*" & vbCrLf _
      + " from (select OrderID, OrderDate, ProductID, Sum(Qty) as UQty from SaleBody b inner join SaleHeader h on H.SID = b.SID Group By OrderID, OrderDate, ProductID) b " & vbCrLf _
      + " right outer join SaleOrderBody sob on sob.OrderID = b.orderid and sob.OrderDate = b.orderdate and b.ProductID = sob.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = sob.productid" & vbCrLf _
      + " where sob.OrderID = " & Val(TxtOrderID.Text) & " and sob.OrderDate = '" & DtpOrderDate.DateValue & "' and (isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty - isnull(uqty,0) <> 0"
   
   With cn.Execute(sSQL)
      If .RecordCount = 0 Then
         cn.Execute ("Update SaleOrderHeader set IsSale = 1 Where OrderID = " & Val(TxtOrderID.Text) & " And Orderdate ='" & DtpOrderDate.DateValue & "'")
      End If
   End With
   If vIsNewRecord = True Then Call ActivityLogSale("Sale Invoice", eAdd, TxtBillID.Text, DtpBillDate.DateValue)
'   If vIsNewRecord = True Then Call ActivityLogBin("Sale Invoice", eAdd, TxtBillID.Text, DtpBillDate.DateValue, "")
   
   cn.CommitTrans
   
   Dim vMobileNoCust As String
   vMobileNoCust = Val(Replace(vContactNo, "-", ""))
     
   If ObjRegistry.CustomerSalesMessage <> "" And ObjRegistry.AllowSMSThroughDevice = True And Len(vMobileNoCust) = 10 Then
      vMobileNoCust = ObjRegistry.PrefixPhoneNo + Right(Replace(vContactNo, "-", ""), 10)
      vMobileNoCust = (Replace(vMobileNoCust, "-", ""))
      If Len(vMobileNoCust) >= 9 Then
         sSQL = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobileNoCust & "','',N'" & ReplaceSMS(ObjRegistry.CustomerSalesMessage) & "','')"
'          ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobileNoCust & "','',N'" & ReplaceSMS(cn.Execute("Select  Value From sysindexs Where RegistryKey = 'CustomerSalesMessage'").Fields(0).Value) & "','')"
         cn.Execute sSQL
      End If
   End If
   
   '/******* WEB SMS *************/
   
   
   If ObjRegistry.CustomerSalesMessage <> "" And ObjRegistry.WebLinkForSMS <> "" Then
      vMobileNoCust = ObjRegistry.PrefixPhoneNo + Right(Replace(vContactNo, "-", ""), 10)
      If Val(Replace(vMobileNoCust, "-", "")) >= 9 Then
         Call WebSMS(ObjRegistry.WebLinkForSMS, ReplaceSMS(ObjRegistry.CustomerSalesMessage), (Replace(vMobileNoCust, "-", "")))
      End If
   End If
   
   '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = ObjRegistry.PrefixPhoneNo + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               sSQL = ObjUserSecurity.UserName & " Saved ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & " Time: " & Time & IIf(Val(TxtTotalDiscount.Caption) = 0, "", " Disc:" & TxtTotalDiscount.Caption) & vbCrLf & " NetAmt" & TxtNetAmount.Caption
               sSQL = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & sSQL & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               cn.Execute sSQL
            End If
         Next
   End If
   
   '/***********************/
   
'   Char.Speak "Thank you for comming"
   'If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
   
   If ChkPrint.Value = 1 Then Call BtnPrint_Click
   If ObjRegistry.ShowLastInvoiceMsgAtSave = True Then
'      sSql = " Bill ID = " & TxtBillID.Text & vbCrLf & " Total Amount = " & TxtNetAmountCash.Text & vbCrLf & " Cash Received = " & TxtCashReceivedCash.Text & vbCrLf & " Cash Return = " & TxtCashReturn.Text
'      MsgBox sSql, vbOKOnly, "Information"  ' for al habib it is blocked
      FrmLastInvoiceInfo.LblBillID.Caption = "Bill ID = " & TxtBillID.Text
      FrmLastInvoiceInfo.LblNetAmountCash.Caption = "Total Amount = " & TxtNetAmountCash.Text
      FrmLastInvoiceInfo.LblCashReceivedCash.Caption = "Cash Received = " & TxtCashReceivedCash.Text
      FrmLastInvoiceInfo.LblCashReturn.Caption = "Cash Return = " & TxtCashReturn.Text
      FrmLastInvoiceInfo.Show vbModal, Me
      
   End If
'   cnSalePOS.Close
   
   
   FormStatus = NewMode
   'End If
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub


Private Sub Sub_Bin_Save()
   On Error GoTo ErrorHandler
   If Grid.rows <= 1 Then Exit Sub
'   If DtpBillDate.Enabled And DtpBillDate.Date <> IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date)) And DateFlag = True Then
'      If MsgBox("Are you sure to Change Bill Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
'         DtpBillDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))
'         TxtBillID.Text = FunGetMaxID()
'      End If
'      DateFlag = False
'   End If
  'Saving record
   cn.BeginTrans
    
   If vIsNewRecord = True Then
      If cn.Execute("Select * from bin_SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "' --and StampID <> " & TxtStampID.Text).RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtBillID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
    
'   Call SubLastEntryDate(DtpBillDate.DateValue)
   
'   If vIsNewRecord = False Then Call ActivityLogSale("Bin Sale Invoice", eEdit, TxtBillID.Text, DtpBillDate.DateValue)
   ''''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Call UserActivities
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  
'' Sale Header

Dim vInvoiceNo, vComission, vBankMachineID, vCashReceived, vCustomerID, vCustomerName As String
      If OptBankCard.Value = True Then
         vInvoiceNo = TxtInvoiceNo.Text
         vComission = TxtCommision.Text
         vBankMachineID = TxtBankMachineID.Text
         vCashReceived = Val(TxtCashReceivedBank.Text)
         vCustomerID = "621"
         vCustomerName = IIf(Trim(TxtBankCustomer.Text) = "", Null, TxtBankCustomer.Text)
      End If
      If OptCash.Value = True Then
         vComission = "''"
         vInvoiceNo = Null
         vBankMachineID = "''"
         vCashReceived = Val(TxtCashReceivedCash.Text)
         vCustomerID = "621"
         vCustomerName = IIf(Trim(TxtCashCustomer.Text) = "", "''", TxtCashCustomer.Text)
      End If
      If OptCredit.Value = True Then
         vComission = "''"
         vInvoiceNo = Null
         vBankMachineID = "''"
         vCashReceived = Val(TxtCashReceivedCredit.Text)
         vCustomerID = TxtCustomerID.Text
'         vCustomerName = Null
      End If
      
vStrPara = ""
vStrPara = vStrPara & TxtBillID.Text & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
vStrPara = vStrPara & "'" & vCustomerID & "'," 'CustomerID
vStrPara = vStrPara & SelfRound(vTotalAmount) & "," ' Total Amount
vStrPara = vStrPara & Val(TxtBillDisc.Text) & "," 'BillDisc
vStrPara = vStrPara & vCashReceived & "," ' 'CashReceived
vStrPara = vStrPara & vUser & "," 'UserNo
vStrPara = vStrPara & TxtStoreID.Text & "," 'StoreID
vStrPara = vStrPara & IIf(OptBankCard.Value = True, 1, 0) & "," 'BankCard
vStrPara = vStrPara & IIf(OptCash.Value = True, 1, 0) & "," 'Credit
vStrPara = vStrPara & IIf(OptCredit.Value = True, 1, 0) & "," 'Cash
vStrPara = vStrPara & vBankMachineID & "," 'BankMachineID
vStrPara = vStrPara & "'" & vInvoiceNo & "',"  'InvoiceNo
vStrPara = vStrPara & "'" & vCustomerName & "'," 'CustomerName
vStrPara = vStrPara & Val(TxtBillDiscPer.Text) & "," 'BillDiscPer
vStrPara = vStrPara & vComission & ","   'Commision
vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtCommission.Text)) & "," 'EmpComm
vStrPara = vStrPara & "'" & IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text)) & "'," 'EmpID
vStrPara = vStrPara & 0 & "," 'isReplace
vStrPara = vStrPara & 0 & "," 'isPosted
vStrPara = vStrPara & IIf(Trim(TxtMemberID.Text) = "", "''", TxtMemberID.Text) & "," 'MemberID
vStrPara = vStrPara & "'" & Now & "'," 'BillTime
vStrPara = vStrPara & "'True'," 'Tag
vStrPara = vStrPara & "'" & IIf(Trim(TxtManualBillNo.Text) = "", Null, TxtManualBillNo.Text) & "'," 'ManualBillNo
'vStrPara = vStrPara & "'" & IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text) & "',"  'Remarks
vStrPara = vStrPara & "'" & vRemarks & "',"  'Remarks
vStrPara = vStrPara & IIf(Trim(TxtOrganizationID.Text) = "", "''", TxtOrganizationID.Text) & ","  'OrganizationID
vStrPara = vStrPara & "'" & Null & "'," ' BillNo
vStrPara = vStrPara & "'" & Null & "'," ' Bilty No
vStrPara = vStrPara & "'" & vDescription & "'," 'Description
vStrPara = vStrPara & "''" & "," 'PAIDAMOUNT
vStrPara = vStrPara & "'" & Null & "',"  'EntryDate
vStrPara = vStrPara & 0 & "," 'PreviousAmount
vStrPara = vStrPara & 0 & "," 'OtherCharges
vStrPara = vStrPara & "'" & Null & "'," 'SaleManID
vStrPara = vStrPara & 0 & "," 'TotalExpense
vStrPara = vStrPara & IIf(Val(TxtOrderID.Text) = 0, "''", TxtOrderID.Text) & "," 'OrderID
vStrPara = vStrPara & "'" & DtpOrderDate.DateValue & "'," 'OrderDate
vStrPara = vStrPara & 0 & "," 'Freight
vStrPara = vStrPara & 0 & "," 'IsCustomerFreight
vStrPara = vStrPara & "'" & Null & "'," 'VechicleNo
vStrPara = vStrPara & IIf(TxtServiceCharges.Text = "", "''", Val(TxtServiceCharges.Text)) & "," 'ServiceCharges
vStrPara = vStrPara & IIf(TxtServiceCharges.Text = "", "''", Val(TxtServiceCharges.Text)) & "," 'ServiceChargesPer
vStrPara = vStrPara & IIf(TxtSTax.Text = "", "''", Val(TxtSTax.Text)) & "," 'STax
vStrPara = vStrPara & IIf(TxtSTax.Text = "", "''", Val(TxtSTax.Text)) & "," 'STaxPer
vStrPara = vStrPara & "'" & IIf(Trim(TxtTableID.Text) = "", Null, TxtTableID.Text) & "'," 'TableID
vStrPara = vStrPara & "'" & Null & "'," 'ServerEntry
vStrPara = vStrPara & "'" & IIf(CmbType.Visible = False, Null, CmbType.Text) & "'," 'InvType
vStrPara = vStrPara & "'" & DtpDeliveryDate.DateValue & "'," 'DeliveryDate
vStrPara = vStrPara & "'" & DTPDeliveryTime.Value & "'," 'DeliveryTime
vStrPara = vStrPara & "'" & Null & "'," 'isPrinted
vStrPara = vStrPara & "'" & Null & "'," 'RemarksUrdu
vStrPara = vStrPara & "'" & Null & "'," 'StampID
vStrPara = vStrPara & 0 & "," 'isTransfer
vStrPara = vStrPara & IIf(DtpPromiseDate.DateValue = Empty, "Null", "'" & DtpPromiseDate.DateValue & "'") 'PromiseDate
vStrPara = Replace(vStrPara, "''", "Null")

vStrPara = "DECLARE @returnvalue INT EXEC @returnvalue = bin_saleheaderinsert " & vStrPara & " Select @returnvalue"
   vMasterID = cn.Execute(vStrPara).Fields(0).Value
'   MsgBox vMasterID
   
vStrDetail = ""
''' insert Sale Body
With Grid
 .Redraw = False
 .MoveFirst
   For vCounter = 1 To .rows
      If Trim(.Columns("Productid").Text) <> "" Then
      
      ''''''''''''''''''''''''''''
 vStrPara = ""

vStrPara = vStrPara & vMasterID & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
'vStrPara = vStrPara & .Columns("SerialNo").Text & ","
'vStrPara = vStrPara & .Columns("BillID").Text & ","
'vStrPara = vStrPara & .Columns("BillDate").Text & ","
vStrPara = vStrPara & "'" & .Columns("ProductID").Text & "',"
vStrPara = vStrPara & .Columns("Qty").Text & ","
vStrPara = vStrPara & .Columns("Price").Text & ","
vStrPara = vStrPara & .Columns("DiscPC").Text & ","
vStrPara = vStrPara & .Columns("Amount").Text & ","
vStrPara = vStrPara & "'" & .Columns("Code").Text & "',"
vStrPara = vStrPara & .Columns("DiscPer").Text & ","
vStrPara = vStrPara & .Columns("DiscVal").Text & ","
vStrPara = vStrPara & .Columns("Cost").Text & ","
vStrPara = vStrPara & .Columns("isProduct").Text & ","
vStrPara = vStrPara & "''" & "," ' Pack Name
vStrPara = vStrPara & "''" & "," ' Qty Pack
vStrPara = vStrPara & "''" & "," ' Pack
vStrPara = vStrPara & "''" & "," ' Bonus
vStrPara = vStrPara & "''" & "," 'Offer
vStrPara = vStrPara & "''" & "," 'SaleTaxPer
vStrPara = vStrPara & "''" & "," ' SaleTaxVal
vStrPara = vStrPara & Val(.Columns("TokenVal").Text) & ","
vStrPara = vStrPara & Val(TxtPrice.Text) & "," 'RetailPrice
vStrPara = vStrPara & 0 & "," 'IsWSSaleTax
vStrPara = vStrPara & 0 & "," 'IsRetailSaleTax
vStrPara = vStrPara & 0 & "," 'IsWSDiscb4ST
vStrPara = vStrPara & "''" & "," 'SC
vStrPara = vStrPara & Val(TxtEmpComm.Text) & "," 'EmpComm
vStrPara = vStrPara & "''" & "," 'BatchNo
vStrPara = vStrPara & "''" & "," 'StampID
vStrPara = vStrPara & IIf(Trim(TxtStoreID.Text) = "", "''", Val(TxtStoreID.Text)) & "," 'StoreID
vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtEmployeeID.Text))  'EmpID
vStrPara = Replace(vStrPara, "''", "Null")
cn.Execute ("Exec Bin_SaleBodyInsert " & vStrPara)
vStrDetail = vStrDetail & " (P" & .Columns("ProductID").Text & " Q" & .Columns("Qty").Text & " A" & .Columns("Amount").Text & ")"
      ''''''''''''''''''''''''''''
      
      End If
      .MoveNext
   Next vCounter
'   .RemoveAll
   .Redraw = True
   End With
   

   
'    ssql = "select * from SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
'    Dim Rs As New ADODB.Recordset
'    With Rs
'      .Open ssql, cnSalePOS, adOpenDynamic, adLockPessimistic
'      If .BOF Then
'         .AddNew
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         !OrderID = IIf(Val(TxtOrderID.Text) = 0, Null, TxtOrderID.Text)
'         !OrderDate = DtpOrderDate.DateValue
''         !StampID = TxtStampID.Text
'         !BillTime = Now
'         !UserNo = vUser
'      End If
'      !isReplace = 0
'      !isPosted = 0
'      !isTransfer = 0
'      !InvType = IIf(CmbType.Visible = False, Null, CmbType.Text)
'      !StoreID = Val(TxtStoreID.Text)
'      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
'      !TableId = IIf(Trim(TxtTableID.Text) = "", Null, TxtTableID.Text)
'      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
'      !EmpComm = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text))
'      !MemberID = IIf(Trim(TxtMemberID.Text) = "", Null, TxtMemberID.Text)
'      !TotalAmount = SelfRound(vTotalAmount)
'      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
'      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
'      !ServiceCharges = IIf(TxtServiceCharges.Text = "", Null, Val(TxtServiceCharges.Text))
'      !ServiceChargesPer = IIf(TxtServiceCharges.Text = "", Null, Val(TxtServiceCharges.Text))
'      !STax = IIf(TxtSTax.Text = "", Null, Val(TxtSTax.Text))
'      !STaxPer = IIf(TxtSTax.Text = "", Null, Val(TxtSTax.Text))
'      !DeliveryDate = DtpDeliveryDate.DateValue
'      !DeliveryTime = DTPDeliveryTime.Value
'      If OptBankCard.Value = True Then
'         !InvoiceNo = TxtInvoiceNo.Text
'         !Commision = TxtCommision.Text
'         !BankMachineID = TxtBankMachineID.Text
'         !CashReceived = Val(TxtCashReceivedBank.Text)
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(TxtBankCustomer.Text) = "", Null, TxtBankCustomer.Text)
'      End If
'      If OptCash.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !BankMachineID = Null
'         !CashReceived = Val(TxtCashReceivedCash.Text)
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(TxtCashCustomer.Text) = "", Null, TxtCashCustomer.Text)
'      End If
'      If OptCredit.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !BankMachineID = Null
'         !CashReceived = Val(TxtCashReceivedCredit.Text)
'         !CustomerID = TxtCustomerID.Text
'         !CustomerName = Null
'      End If
'      !BankCard = OptBankCard.Value
'      !Cash = OptCash.Value
'      !Credit = OptCredit.Value
'      '!Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
'      !Remarks = IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text)
'      !RemarksUrdu = IIf(Trim(TxtRemarksUrdu.Text) = "", Null, TxtRemarksUrdu.Text)
'      !ManualBillNo = IIf(Trim(TxtManualBillNo.Text) = "", "", TxtManualBillNo.Text)
''      .Update
'      .Close
'   End With
'
'   With RsBody
'      .Filter = 0
'      .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
''         !StampID = TxtStampID.Text
'         .MoveNext
'      Next vCounter
''      .UpdateBatch
'   End With
   
'   With RsDetail
'      .Filter = 0
'      If .RecordCount > 0 Then .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         .MoveNext
'      Next vCounter
''      .UpdateBatch
'   End With
   
'   sSql = " select sob.ProductID, ProductName, (isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty - isnull(uqty,0) as Qtyloose, sob.*" & vbCrLf _
      + " from (select OrderID, OrderDate, ProductID, Sum(Qty) as UQty from SaleBody b inner join SaleHeader h on h.BillID = b.BillID and h.BillDate = b.BillDate Group By OrderID, OrderDate, ProductID) b " & vbCrLf _
      + " right outer join SaleOrderBody sob on sob.OrderID = b.orderid and sob.OrderDate = b.orderdate and b.ProductID = sob.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = sob.productid" & vbCrLf _
      + " where sob.OrderID = " & Val(TxtOrderID.Text) & " and sob.OrderDate = '" & DtpOrderDate.DateValue & "' and (isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty - isnull(uqty,0) <> 0"
   
'   With cnSalePOS.Execute(sSql)
'      If .RecordCount = 0 Then
'         cnSalePOS.Execute ("Update SaleOrderHeader set IsSale = 1 Where OrderID = " & Val(TxtOrderID.Text) & " And Orderdate ='" & DtpOrderDate.DateValue & "'")
'      End If
'   End With
'   If vIsNewRecord = True Then Call ActivityLogSale("Sale Invoice", eAdd, TxtBillID.Text, DtpBillDate.DateValue)
   cn.CommitTrans
'   Dim vMobileNo As String
'   If ObjRegistry.CustomerSalesMessage <> "" Then
'      vMobileNo = ObjRegistry.PrefixPhoneNo + Right(Replace(TxtCashCustomer.Text, "-", ""), 10)
'      If Val(Replace(TxtCashCustomer.Text, "-", "")) >= 9 Then
'         sSql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobileNo & "','','" & ObjRegistry.CustomerSalesMessage & "','')"
'         cnSalePOS.Execute sSql
'      End If
'   End If
'   '/******* Mobile SMS *************/
'   If ObjRegistry.OwnerMobileNo <> "" Then
'      sSql = " Bill ID = " & TxtBillID.Text & vbCrLf & " Bill Date = " & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & vbCrLf & " Total Amount = " & TxtNetAmountCash.Text
'      sSql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & ObjRegistry.OwnerMobileNo & "','','" & sSql & "','')"
'      cnSalePOS.Execute sSql
'   End If
   
   '/***********************/
      
'   Char.Speak "Thank you for comming"
   'If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
'   If ChkPrint.Value = 1 Then Call BtnPrint_Click
   If ObjRegistry.ShowLastInvoiceMsgAtSave = True Then
      sSQL = " Bill ID = " & TxtBillID.Text & vbCrLf & " Total Amount = " & TxtNetAmountCash.Text & vbCrLf & " Cash Received = " & TxtCashReceivedCash.Text & vbCrLf & " Cash Return = " & TxtCashReturn.Text
      MsgBox sSQL, vbOKOnly, "Information" ' for al habib it is blocked
   End If
   
'   cnSalePOS.Close
'   If cnSalePOS.State = adStateClosed Then cnSalePOS.Open
'   FormStatus = NewMode
   'End If
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub


Private Function FunSelectMember(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchMember.Show vbModal, Me
        If SchMember.ParaOutMemberID = "" Then FunSelectMember = False: Exit Function
        TxtMemberID.Text = SchMember.ParaOutMemberID
    End If
    '---------------------------
    If Trim(TxtMemberID.Text) = "" Then Exit Function
    
        vStrSQL = " Select c.* FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " Left Outer join Members m on c.AccountNo = cast(m.Prefix as varchar(2))  + cast(m.MemberID as varchar(10)) " & vbCrLf & _
              " where p.BarCode = '" & (TxtCustomerID.Text) & "' or m.BarCode = '" & (TxtCustomerID.Text) & "' or (c.AccountNo = '" & (TxtCustomerID.Text) & "' and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo like '3%') and c.isDetailed = 1 and c.isLocked = 0)"

    sSQL = "Select * " & vbCrLf _
            + " from Members" & vbCrLf _
            + " where IsLockMember = 0 and ( MemberID = case when isnumeric('" & Trim(TxtMemberID.Text) & " ')=1 then '" & Trim(TxtMemberID.Text) & " ' else '' end or BarCode = '" & Trim(TxtMemberID.Text) & "')"
    With cn.Execute(sSQL)
      If .RecordCount > 0 Then
        TxtMemberID.Text = !MemberID
        TxtMemberName.Text = !MemberName
        TxtMemberBarCode.Text = IIf(IsNull(!BarCode), "", !BarCode)
        If !ExpiryDate > Date Or IsNull(!ExpiryDate) = True Then Call SubApplyMember Else MsgBox "Discount Not Applied Because Member's Discount Expired ", vbExclamation, "Alert"
        FunSelectMember = True
        .Close
        Exit Function
      Else
        FunSelectMember = False
        .Close
        TxtMemberID.Text = ""
        TxtMemberName.Text = ""
        Exit Function
      End If
    End With
Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function
Private Sub BtnStore_Click()
   On Error GoTo ErrorHandler
   If FunSelectStore(ssButton, False) = True Then
      If TxtMemberID.Visible And TxtMemberID.Enabled Then TxtMemberID.SetFocus Else TxtStoreID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
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
   If TxtStoreName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectStore(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectStore(ssButton, False)
   End If
   TxtBillID.Text = FunGetMaxID()
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
    vStrSQL = "Select * FROM Stores where isLock = 0 and StoreID = " & Val(TxtStoreID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtStoreName.Text = !StoreName
          FunSelectStore = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectStore = False
          .Close
          TxtStoreID.Text = ""
          TxtStoreName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function


Private Sub DtpBillDate_LostFocus()
   On Error GoTo ErrorHandler
   TxtBillID.Text = FunGetMaxID()
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbType_Click()
   On Error GoTo ErrorHandler
   With cn.Execute("Select ServiceChargesPer, IsEdit from InvTypes where InvType = '" & CmbType.Text & "'")
      If .RecordCount > 0 Then
         'vIsEdit = !IsEdit
         TxtServiceChargesPer.Enabled = False
         TxtServiceCharges.Enabled = False
         TxtServiceCharges.Tag = "NC"
         TxtServiceChargesPer.Tag = "NC"
         TxtServiceChargesPer.Text = !ServiceChargesPer
      Else
         TxtServiceChargesPer.Enabled = True
         TxtServiceCharges.Enabled = True
         TxtServiceCharges.Tag = ""
         TxtServiceChargesPer.Tag = ""
         TxtServiceChargesPer.Text = ""
      End If
      .Close
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub DtpPromiseDate_Change()
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub DtpPromiseDate_DblClick()
   DtpPromiseDate.DateValue = Null
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub


Private Sub LblClose_Click()
   FraHelp.Visible = False
End Sub

Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   LblHelp.ForeColor = &H800000
   FraHelp.ZOrder 0
   FraHelp.Visible = True
End Sub

Private Sub LblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If LblHelp.FontUnderline = True Then Exit Sub
   LblHelp.FontUnderline = True
End Sub

Private Sub LblHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   LblHelp.ForeColor = vbWhite
End Sub

Private Sub InvoiceNo()
   On Error GoTo ErrorHandler
   Dim vC As Byte, LoopFlag As Boolean
   vC = 1: LoopFlag = True
   With cn.Execute("Select * from TempNo where UserNo = " & vUser & " order by tempno")
      While (Not .EOF) And LoopFlag = True
         If vC <> !TempNo And Not .EOF Then
            LoopFlag = False
         Else
            vC = vC + 1
         End If
         .MoveNext
      Wend
      LblNo.Caption = " Inv. Open # " & CStr(vC)
      cn.Execute "INSERT INTO TempNo(TempNo,UserNo) VALUES (" & vC & "," & vUser & ")"
      .Close
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub MniCostPrice_Click()
   On Error GoTo ErrorHandler
'   If Trim(Grid.Columns("Cost").Text) = "" Then Exit Sub
'   If ObjUserSecurity.ShowPurchasePriceInInvoice = True Or ObjUserSecurity.IsAdministrator = True Then
'      LblCost.Caption = Grid.Columns("Cost").Value
      LblCost.Visible = True
'   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub GetSale()
   On Error GoTo ErrorHandler
   sSQL = "select h.*, c.AccountName, BankMachineName, StoreName, EmpName, TableName, MemberName, OrganizationName FROM SaleHeader h left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join Organizations o on o.Organizationid = h.Organizationid left outer join BankMachines b on b.BankMachineid = h.BankMachineid left outer join Members m on m.MemberID = h.MemberID left outer join Tables t on t.TableID = h.TableID inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID where isReplace=0 and h.SID=" & Val(TxtSID.Text) & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With cn.Execute(sSQL)
      If Not .BOF Then
         TxtOrderID.Text = IIf(IsNull(!OrderID), "", !OrderID)
         DtpOrderDate.DateValue = IIf(IsNull(!OrderDate), "01/01/1990", !OrderDate)
         If IsNull(!InvType) Or !InvType = "" Then
            CmbType.ListIndex = 0
         Else
            CmbType.Text = !InvType
         End If
'         TxtStampID.Text = IIf(IsNull(!StampID), "1", !StampID)
         DtpPromiseDate.DateValue = !PromiseDate
         TxtStoreID.Text = !StoreID
         TxtStoreName.Text = !StoreName
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         TxtTableID.Text = IIf(IsNull(!TableId), "", !TableId)
         TxtTableName.Text = IIf(IsNull(!TableName), "", !TableName)
         TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
         TxtEmployeeName.Text = IIf(IsNull(!empname), "", !empname)
         TxtCommission.Text = IIf(IsNull(!EmpComm), "", !EmpComm)
         TxtMemberID.Text = IIf(IsNull(!MemberID), "", !MemberID)
         TxtMemberName.Text = IIf(IsNull(!MemberName), "", !MemberName)
         TxtTotalAmount.Caption = !TotalAmount
         TxtSumDiscAmount.Text = !SumDiscAmount
         TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
         TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
         TxtServiceChargesPer.Text = IIf(IsNull(!ServiceChargesPer), "", !ServiceChargesPer)
         TxtServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtSTaxPer.Text = IIf(IsNull(!STaxPer), "", !STaxPer)
         TxtSTax.Text = IIf(IsNull(!STax), "", !STax)
         TxtAdvTaxVal.Text = IIf(IsNull(!AdvTaxVal), "", !AdvTaxVal)
         TxtAdvTaxPer.Text = IIf(IsNull(!AdvTaxPer), "", !AdvTaxPer)
         TxtExtraTaxVal.Text = IIf(IsNull(!ExtraTaxVal), "", !ExtraTaxVal)
         TxtExtraTaxPer.Text = IIf(IsNull(!ExtraTaxPer), "", !ExtraTaxPer)
         TxtManualBillNo.Text = IIf(IsNull(!ManualBillNo), "", !ManualBillNo)
         TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
         TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
         TxtRemarksUrdu.Text = IIf(IsNull(!RemarksUrdu), "", !RemarksUrdu)
         OptBankCard.Value = !BankCard
         OptCash.Value = !Cash
         OptCredit.Value = !Credit
         If OptBankCard.Value = True Then
            TxtInvoiceNo.Text = IIf(IsNull(!InvoiceNo), "", !InvoiceNo)
            TxtCommision.Text = !Commision
            TxtBankMachineID.Text = !BankMachineID
            TxtBankMachineName.Text = !BankMachineName
            TxtCashReceivedCash.Text = ""
            TxtCashReceivedCredit.Text = ""
            TxtCashReceivedBank.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            TxtCustomerID.Text = ""
            TxtCustomerName.Text = ""
            TxtCashCustomer.Text = ""
            TxtBankCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
         End If
         If OptCash.Value = True Then
            TxtCommision.Text = ""
            TxtInvoiceNo.Text = ""
            TxtBankMachineID.Text = ""
            TxtBankMachineName.Text = ""
            TxtCashReceivedCash.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            TxtCashReceivedCredit.Text = ""
            TxtCashReceivedBank.Text = ""
            TxtCustomerID.Text = ""
            TxtCustomerName.Text = ""
            TxtCashCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
            TxtBankCustomer.Text = ""
            TxtCNIC.Text = IIf(IsNull(!CNIC), "", !CNIC)
            TxtCellNo.Text = IIf(IsNull(!MobileNo), "", !MobileNo)
         End If
         If OptCredit.Value = True Then
            TxtBankMachineCreditID.Text = IIf(IsNull(!BankMachineID), "", !BankMachineID)
            TxtBankMachineCreditName.Text = IIf(IsNull(!BankMachineName), "", !BankMachineName)
            TxtBankAmount.Text = IIf(IsNull(!BankAmount), "", !BankAmount)
            TxtCommision.Text = IIf(IsNull(!Commision), "", !Commision)
            TxtInvoiceNo.Text = ""
            TxtBankMachineID.Text = ""
            TxtBankMachineName.Text = ""
            TxtCashReceivedCash.Text = ""
            TxtCashReceivedCredit.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            TxtCashReceivedBank.Text = ""
            TxtCustomerID.Text = !CustomerID
            TxtCustomerName.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
            TxtRefID.Text = IIf(IsNull(!RefID), "", !RefID)
            TxtRefComm.Text = IIf(IsNull(!RefComm), "", !RefComm)
            TxtCashCustomer.Text = ""
            TxtBankCustomer.Text = ""
         End If
         TxtNetAmount.Caption = !TotalAmount
         Call PopulateDataToGrid
      End If
      .Close
   End With
   vPrevious = cn.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtCustomerID.Text & "','" & DtpBillDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
   vCurrent = cn.Execute("Select isnull(sum(TotalAmount-isnull(BillDisc,0)+isnull(STax,0)+isnull(ServiceCharges,0)-isnull(CashReceived,0)),0) from SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'" & IIf(TxtOrganizationID.Text = "", "", " and OrganizationID = '" & Val(TxtOrganizationID.Text) & "'")).Fields(0).Value
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub GetSaleOrder()
   On Error GoTo ErrorHandler
   TxtBillID.Text = FunGetMaxID
   sSQL = "select h.*, c.AccountName, BankMachineName, StoreName, TableName, EmpName, MemberName FROM SaleOrderHeader h left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join BankMachines b on b.BankMachineid = h.BankMachineid left outer join Members m on m.MemberID = h.MemberID left outer join Tables t on t.TableID = h.TableID inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID where isReplace=0 and h.OrderID=" & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'"
   With cn.Execute(sSQL)
      If Not .BOF Then
         DtpOrderDate.DateValue = !OrderDate
         DtpDeliveryDate.DateValue = IIf(IsNull(!DeliveryDate), "", !DeliveryDate)
         DTPDeliveryTime.Value = IIf(IsNull(!DeliveryTime), Now, !DeliveryTime)
         If IsNull(!InvType) Or (!InvType = "") Then
            CmbType.ListIndex = 0
         Else
            CmbType.Text = !InvType
         End If
         TxtStoreID.Text = !StoreID
         TxtStoreName.Text = !StoreName
         TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
         TxtEmployeeName.Text = IIf(IsNull(!empname), "", !empname)
         TxtTableID.Text = IIf(IsNull(!TableId), "", !TableId)
         TxtTableName.Text = IIf(IsNull(!TableName), "", !TableName)
         TxtMemberID.Text = IIf(IsNull(!MemberID), "", !MemberID)
         TxtMemberName.Text = IIf(IsNull(!MemberName), "", !MemberName)
         TxtTotalAmount.Caption = !TotalAmount
         TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
         TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
         TxtServiceChargesPer.Text = IIf(IsNull(!ServiceChargesPer), "", !ServiceChargesPer)
         TxtServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtSTaxPer.Text = IIf(IsNull(!STaxPer), "", !STaxPer)
         TxtSTax.Text = IIf(IsNull(!STax), "", !STax)
         TxtManualBillNo.Text = IIf(IsNull(!ManualBillNo), "", !ManualBillNo)
         TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
         TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
         TxtRemarksUrdu.Text = IIf(IsNull(!RemarksUrdu), "", !RemarksUrdu)
         OptBankCard.Value = !BankCard
         OptCash.Value = !Cash
         OptCredit.Value = !Credit
         If OptBankCard.Value = True Then
            TxtInvoiceNo.Text = !InvoiceNo
            TxtCommision.Text = !Commision
            TxtBankMachineID.Text = !BankMachineID
            TxtBankMachineName.Text = !BankMachineName
            TxtCashReceivedCash.Text = ""
            TxtCustomerID.Text = ""
            TxtCustomerName.Text = ""
            TxtCashCustomer.Text = ""
            TxtBankCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
         End If
         If OptCash.Value = True Then
            TxtCommision.Text = ""
            TxtInvoiceNo.Text = ""
            TxtBankMachineID.Text = ""
            TxtBankMachineName.Text = ""
            TxtCashReceivedCash.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            TxtCustomerID.Text = ""
            TxtCustomerName.Text = ""
            TxtCashCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
            TxtBankCustomer.Text = ""
         End If
         If OptCredit.Value = True Then
            TxtCommision.Text = ""
            TxtInvoiceNo.Text = ""
            TxtBankMachineID.Text = ""
            TxtBankMachineName.Text = ""
            TxtCashReceivedCredit.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            TxtCustomerID.Text = !CustomerID
            TxtCustomerName.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
            TxtCashCustomer.Text = ""
            TxtBankCustomer.Text = ""
         End If
         TxtNetAmount.Caption = !TotalAmount
         Call PopulateSaleOrderDataToGrid
      End If
      .Close
   End With
   'FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtBillDisc_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtBillDisc.Name Then Exit Sub
   DiscPerFlag = False
   TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / IIf(Val(TxtTotalAmount.Caption) = 0, 1, Val(TxtTotalAmount.Caption)), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtBillDiscPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtBillDiscPer.Name Then Exit Sub
   DiscPerFlag = True
   TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtBillDiscPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCode_GotFocus()
   Grid.MoveLast
   Grid.MoveNext
End Sub

Private Sub TxtDiscPC_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDiscPC.Name Then Exit Sub
   If Val(TxtPrice.Text) = 0 Then Exit Sub
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub TxtDiscPC_LostFocus()
'   Select Case Me.ActiveControl.Name
'   Case TxtCode.Name, TxtQty.Name, TxtDiscPC.Name
'      Exit Sub
'   End Select
'   Call GetDataFromTexBoxesToGrid
'End Sub

Private Sub TxtDiscPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDiscPer.Name Then Exit Sub
   TxtDiscPC.Text = Round((Val(TxtPrice.Text) * Val(TxtDiscPer.Text) / 100), 2)
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDiscVal_Change()
   If TxtDiscVal.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtDiscVal.Name Then Exit Sub
   If Val(TxtPrice.Text) = 0 Then Exit Sub
   If Val(TxtQty.Text) = 0 Then Exit Sub
   TxtActualAmount.Text = Val(TxtQty.Text) * (Val(TxtPrice.Text) + Val(TxtSC.Text))
   TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (TxtQty.Text), 3)
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
   TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
   TxtPurAmount.Text = Round(Val(TxtQty.Text) * (Val(TxtLastPurPrice.Text) + Val(TxtSC.Text)), 2)
   TxtProdProfit.Text = Round(Val(TxtAmount.Text) - Val(TxtPurAmount.Text), 2)
   TxtDiscAmount.Text = (Val(TxtQty.Text) * (Val(TxtPrice.Text) + Val(TxtSC.Text))) + Val(TxtDiscVal.Text)
'   TxtTotalDiscount.Caption = Round(vTotDisc, 2)
   SubCalculateFooter
End Sub

Private Sub TxtPrice_Change()
If ActiveControl.Name <> TxtPrice.Name Then Exit Sub
   Call SubCalculateBody
End Sub

'Private Sub TxtProductName_Change()
'   If ActiveControl.Name <> TxtProductName.Name Then Exit Sub
'   Call FindRow
'End Sub

Private Sub TxtQty_Change()
If ActiveControl.Name <> TxtQty.Name Then Exit Sub
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub TxtDiscPC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 8 Then Exit Sub
   If Len(Trim(TxtPrice.Text)) - 1 = Len(Trim(TxtDiscPC.Text)) Then
      If Val(TxtPrice.Text) < Val(TxtDiscPC.Text & Chr(KeyAscii)) Then
         KeyAscii = 0
      End If
   ElseIf Len(Trim(TxtPrice.Text)) = Len(Trim(TxtDiscPC.Text)) Then
         KeyAscii = 0
   End If
End Sub

Private Sub TxtDiscPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 8 Then Exit Sub
   If Len(Trim(TxtDiscPer.Text)) = 2 Then
      If Val(TxtDiscPer.Text & Chr(KeyAscii)) > 100 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub TxtSC_Change()
 On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSC.Name Then Exit Sub
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtServiceCharges_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtServiceCharges.Name Then Exit Sub
'   TxtServiceChargesPer.Text = Round((Val(TxtServiceCharges.Text) * 100) / IIf(Val(TxtTotalAmount.Caption) = 0, 1, Val(TxtTotalAmount.Caption)), 2)
    TxtServiceChargesPer.Text = Round((Val(TxtServiceCharges.Text) * 100) / IIf(Val(vNetAmount) = 0, 1, Val(vNetAmount)), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtServiceCharges_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrorHandler
   If KeyAscii = 8 Then Exit Sub
   If (Val(TxtServiceCharges.Text & Chr(KeyAscii))) > SelfRound(Val(TxtTotalAmount.Caption)) Then
      KeyAscii = 0
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtServiceChargesPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtServiceChargesPer.Name Then Exit Sub
'   TxtServiceCharges.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtServiceChargesPer.Text) / 100))
    TxtServiceCharges.Text = SelfRound((Val(vNetAmount) * Val(TxtServiceChargesPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSTax_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSTax.Name Then Exit Sub
   TxtSTaxPer.Text = Round((Val(TxtSTax.Text) * 100) / IIf(Val(TxtTotalAmount.Caption) = 0, 1, Val(TxtTotalAmount.Caption)), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSTax_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrorHandler
   If KeyAscii = 8 Then Exit Sub
   If (Val(TxtSTax.Text & Chr(KeyAscii))) > SelfRound(Val(TxtTotalAmount.Caption)) Then
      KeyAscii = 0
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSTaxPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSTaxPer.Name Then Exit Sub
   TxtSTax.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtSTaxPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpBillDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = cn.Execute("Select isnull(max(Bin_BillID),0)+1 from Bin_SaleHeader ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub FindRebate()
   Dim Rebate
   On Error GoTo ErrorHandler
    With cn.Execute("Select * from ProductOffers where Rebate <> 0 and ProductID = '" & TxtPID.Text & "'")
        If .RecordCount > 0 Then
            Rebate = Val(TxtQty.Text)
            
            If !FixedRebate Then
               Rebate = IIf(Val(TxtQty.Text) <= !Qty And Val(TxtQty.Text) > 1, 1, 0)
            Else
               Rebate = Rebate \ !Qty
            End If
            
            Rebate = Rebate * !Rebate
            TxtDiscVal.Text = Rebate
            If Val(TxtPrice.Text) = 0 Then Exit Sub
            If Val(TxtQty.Text) = 0 Then Exit Sub
            TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (TxtQty.Text), 3)
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
            TxtActualAmount.Text = Val(TxtQty.Text) * Val(TxtPrice.Text)
            TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
            If ObjRegistry.IsRoundFigure = True Then TxtAmount.Text = SelfRound(TxtAmount.Text)
            TxtTotalDiscount.Caption = Round(vTotDisc, 2)
            SubCalculateFooter
        End If
    End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub UserActivities()
   If vIsNewRecord = False Then
      With cn.Execute("Select  * from SaleHeader where BillID =" & TxtBillID.Text & " And BillDate = '" & DtpBillDate.DateValue & "'")
          If Val(TxtEmployeeID.Text) <> IIf(IsNull(!EmpID), 0, !EmpID) Then
              cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Updated EmpID-" & !EmpID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
          End If
          If TxtMemberID.Text <> !MemberID Then
              cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Updated MemberID-" & !MemberID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
          End If
          If TxtStoreID.Text <> !StoreID Then
              cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
          End If
      End With
      Grid.MoveFirst
      For i = 1 To Grid.rows - 1
         With cn.Execute("Select * from SaleBody Where billID = " & TxtBillID.Text & " and billdate ='" & DtpBillDate.DateValue & "' and Productid ='" & Grid.Columns("Productid").Text & "'")
            If .EOF = True Then
               cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
            Else
               If Grid.Columns("Qty").Text <> !Qty Or Grid.Columns("Price").Text <> !Price Or Grid.Columns("discper").Text <> !DiscPer Then
                  cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Updated ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & !Qty & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
               End If
            End If
         End With
      Grid.MoveNext
      Next
   Else
      cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub TxtNetAmount_Change()
   On Error GoTo ErrorHandler
   If Len(TxtNetAmount.Caption) > 5 Then
      TxtNetAmount.FontSize = 36
   Else
      TxtNetAmount.FontSize = 48
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub



Private Sub TxtTotalAmount_Change()
   On Error GoTo ErrorHandler
   If Len(TxtTotalAmount.Caption) > 5 Then
      TxtTotalAmount.FontSize = 36
   Else
      TxtTotalAmount.FontSize = 48
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTotalDiscount_Change()
   On Error GoTo ErrorHandler
   If Len(TxtTotalDiscount.Caption) >= 4 Then
      TxtTotalDiscount.FontSize = 36
   Else
      TxtTotalDiscount.FontSize = 48
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTotalItems_Change()
   On Error GoTo ErrorHandler
   If Len(TxtTotalItems.Caption) >= 3 Then
      TxtTotalItems.FontSize = 36
   Else
      TxtTotalItems.FontSize = 48
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTotalQty_Change()
   On Error GoTo ErrorHandler
   If Len(TxtTotalQty.Caption) >= 3 Then
      TxtTotalQty.FontSize = 36
   Else
      TxtTotalQty.FontSize = 48
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnTable_Click()
   On Error GoTo ErrorHandler
   If FunSelectTable(ssButton, False) = True Then
      TxtTableID.SetFocus
   Else
      TxtTableID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectTable(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchTable.Show vbModal, Me
        If SchTable.ParaOutTableID = "" Then FunSelectTable = False: Exit Function
        TxtTableID.Text = SchTable.ParaOutTableID
    End If
    '---------------------------
    If Trim(TxtTableID.Text) = "" Then Exit Function
    sSQL = "Select * " & vbCrLf _
            + " from Tables" & vbCrLf _
            + " where TableID = " & Val(TxtTableID.Text)
    With cn.Execute(sSQL)
      If .RecordCount > 0 Then
        TxtTableName.Text = !TableName
        FunSelectTable = True
        .Close
        Exit Function
      Else
        FunSelectTable = False
        .Close
        TxtTableID.Text = ""
        TxtTableName.Text = ""
        Exit Function
      End If
    End With
Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub TxtTableID_Change()
   If TxtTableID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtTableID.Name Then Exit Sub
   If TxtTableName.Text <> "" Then TxtTableName.Text = ""
End Sub

Private Sub TxtTableID_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If TxtTableName.Text <> "" Then Exit Sub
    If TxtTableID.Text = "" Then Exit Sub
    Dim vTemp As Boolean
    vTemp = Not FunSelectTable(ssValidate, True)
    If vTemp = True Then
        vTemp = Not FunSelectTable(ssButton, False)
    End If
    Cancel = vTemp
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub


' Print Form functions
Private Function FunCreditLimit() As Boolean
   FunCreditLimit = False
   vPrevious = cn.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtCustomerID.Text & "','" & DtpBillDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0) ").Fields(0).Value
   vCurrent = cn.Execute("Select isnull(sum(TotalAmount-isnull(BillDisc,0)+isnull(STax,0)+isnull(ServiceCharges,0)-isnull(CashReceived,0)),0) from SaleHeader where SID = " & Val(TxtSID.Text) & "").Fields(0).Value
   ParaOutPrevious = vPrevious - vCurrent
   If OptCredit.Value = True Then
      If Trim(TxtCustomerID.Text) = "" Then
         MsgBox "Please specify a Customer ID", vbExclamation, "Alert"
         TxtCustomerID.SetFocus
         Exit Function
      End If
      With cn.Execute("Select * from Employees where CreditLimit <> 0 and CreditLimit is not null and EmpID = '" & TxtCustomerID.Text & "'")
         If .RecordCount > 0 Then
            If !CreditLimit < (vPrevious - vCurrent + Val(TxtNetAmountCredit.Text) - Val(TxtCashReceivedCredit.Text)) Then
               MsgBox "Credit Limit (" & !CreditLimit & ") is Exceed Balance (" & ParaOutPrevious & ") in this month for this Employee.", vbExclamation, "Alert"
               Exit Function
            End If
         End If
      End With
      With cn.Execute("Select * from Parties where CreditLimit <> 0 and CreditLimit is not null and PartyID = '" & TxtCustomerID.Text & "'")
         If .RecordCount > 0 Then
            If !CreditLimit < (vPrevious - vCurrent + Val(TxtNetAmountCredit.Text) - Val(TxtCashReceivedCredit.Text)) Then
               MsgBox "Credit Limit (" & !CreditLimit & ") is Exceed Balance (" & ParaOutPrevious & ") for this Customer.", vbExclamation, "Alert"
               Exit Function
            End If
         End If
      End With
      With cn.Execute("Select * from Members where CreditLimit <> 0 and CreditLimit is not null and MemberID = substring('" & TxtCustomerID.Text & "',3,10)")
         If .RecordCount > 0 Then
            If !CreditLimit < (vPrevious - vCurrent + Val(TxtNetAmountCredit.Text) - Val(TxtCashReceivedCredit.Text)) Then
               MsgBox "Credit Limit (" & !CreditLimit & ") is Exceed Balance (" & ParaOutPrevious & ") for this Member.", vbExclamation, "Alert"
               Exit Function
            End If
         End If
      End With
   End If
   FunCreditLimit = True
End Function

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   FunValidation = False
   If OptBankCard.Value = True Then
      If Trim(TxtBankMachineID.Text) = "" Then
         MsgBox "Please specify a Bank Machine ID", vbExclamation, "Alert"
         TxtBankMachineID.SetFocus
         Exit Function
      End If
   End If
   If OptCredit.Value = True Then
      If FunCreditLimit = False Then Exit Function
      If Trim(TxtBankMachineCreditID.Text) <> "" Then
         If Val(TxtBankAmount.Text) = 0 Then
            MsgBox "Please Enter Bank Amount", vbExclamation, "Alert"
            Exit Function
         End If
         If Val(TxtBankAmount.Text) + Val(TxtCashReceivedBank.Text) > Val(TxtNetAmountCredit.Text) Then
            MsgBox "Amount Greater than Net Amount", vbExclamation, "Alert"
            TxtBankAmount.SetFocus
            Exit Function
         End If
      End If
      If Val(TxtBankAmount.Text) > 0 Then
          If Trim(TxtBankMachineCreditID.Text) = "" Then
            MsgBox "Please specify a Bank Machine ID OR Bank Amount is 0", vbExclamation, "Alert"
            TxtBankMachineCreditID.SetFocus
            Exit Function
          End If
      End If
   End If
   If OptCash.Value = True Then
      If Val(TxtCashReceivedCash.Text) = 0 Then
         MsgBox "Please specify Cash Received", vbExclamation, "Alert"
         TxtCashReceivedCash.SetFocus
         Exit Function
      End If
      If Val(TxtCashReturn.Text) < 0 Then
         MsgBox "Cash Received not less than Net Amount", vbExclamation, "Alert"
         TxtCashReceivedCash.SetFocus
         Exit Function
      End If
      If Trim(TxtCashCustomer.Text) = "" Then
         MsgBox "Please Enter Customer Name", vbExclamation, "Alert"
         TxtCashCustomer.SetFocus
         Exit Function
      End If
   End If
   
   FunValidation = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectBankMachine(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchBankMachine.Show vbModal, Me
        If SchBankMachine.ParaOutBankMachineID = "" Then FunSelectBankMachine = False: Exit Function
        TxtBankMachineID.Text = SchBankMachine.ParaOutBankMachineID
    End If
    '---------------------------
    vStrSQL = "Select * FROM BankMachines where BankMachineID=" & Val(TxtBankMachineID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtBankMachineName.Text = !BankMachineName
          TxtCommision.Text = !Commision
          FunSelectBankMachine = True
          .Close
          Exit Function
      Else
          FunSelectBankMachine = False
          .Close
          TxtBankMachineID.Text = ""
          TxtBankMachineName.Text = ""
          TxtCommision.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function



Private Sub BtnBankMachine_Click()
   On Error GoTo ErrorHandler
   If FunSelectBankMachine(ssButton, False) = True Then
      If BtnOk.Visible And BtnOk.Enabled Then BtnOk.SetFocus Else TxtBankMachineID.SetFocus
   Else
      TxtBankMachineID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub




Private Sub OptBankCard_Click()
   FrameCash.Visible = False
   FrameCredit.Visible = False
   FrameBank.Visible = True
   TxtBankCustomer.Text = IIf(TxtBankCustomer.Text = "", "Counter Sale", TxtBankCustomer.Text)
   If Trim(TxtBankMachineID.Text) <> "" Then Exit Sub
   If Trim(TxtBankMachineCreditID.Text) <> "" Then Exit Sub
   TxtBankMachineID.Text = ObjRegistry.BankMachineID
   FunSelectBankMachine ssValidate, True
End Sub

Private Sub OptCash_Click()
   FrameCash.Visible = True
   FrameCredit.Visible = False
   FrameBank.Visible = False
   TxtCashCustomer.Text = IIf(TxtCashCustomer.Text = "", "Counter Sale", TxtCashCustomer.Text)
   If vIsRemarksCompulsory = True Then TxtCashCustomer.Text = ""
   If TxtCashReceivedCash.Enabled And TxtCashReceivedCash.Visible Then TxtCashReceivedCash.SetFocus
End Sub

Private Sub OptCredit_Click()
   FrameCash.Visible = False
   FrameCredit.Visible = True
   FrameBank.Visible = False
   If Trim(TxtBankMachineID.Text) <> "" Then Exit Sub
   If Trim(TxtBankMachineCreditID.Text) <> "" Then Exit Sub
   TxtBankMachineCreditID.Text = ObjRegistry.BankMachineID
   FunSelectBankMachineCredit ssValidate, True
End Sub

Private Sub TxtCashReceivedCash_Change()
   If Not IsNumeric(TxtCashReceivedCash.Text) Then
    TxtCashReceivedCash.Text = ""
   End If
   TxtCashReturn.Text = Val(TxtCashReceivedCash.Text) - Val(TxtNetAmountCash.Text)
End Sub

Private Sub TxtNetAmountCash_Change()
   Call TxtCashReceivedCash_Change
   TxtNetAmountBank.Text = TxtNetAmountCash.Text
   TxtNetAmountCredit.Text = TxtNetAmountCash.Text
End Sub

Private Sub TxtBankMachineID_Change()
   If TxtBankMachineID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtBankMachineID.Name Then Exit Sub
   If TxtBankMachineName.Text <> "" Then
      TxtBankMachineName.Text = ""
      TxtCommision.Text = ""
   End If
End Sub

Private Sub TxtBankMachineID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtBankMachineID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtBankMachineName.Text <> "" Then Exit Sub
   If Trim(TxtBankMachineID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBankMachine(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBankMachine(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub SubFrameLoad()
   On Error GoTo ErrorHandler
   Frame1.Visible = True
   Frame1.ZOrder 0
   
   FrameCash.Left = 900
   FrameCash.Top = 900
   
   FrameBank.Left = 200
   FrameBank.Top = 900
   
   FrameCredit.Left = 200
   FrameCredit.Top = 1200
   
  
If vIsDisableCreditSale = True Then OptCredit.Enabled = False
'   OptCredit.Value = vIsCreditSale
    If OptCredit.Value = False Then
      OptCredit.Value = False
      OptCash.Value = True
    End If
   '''''''''''''''''''''
   
   
   
   
   If OptCash.Value = True Then Call OptCash_Click
   If OptCredit.Value = True Then OptCredit_Click: OptCredit.SetFocus
   If OptBankCard.Value = True Then OptBankCard_Click: OptBankCard.SetFocus
   
   
'   ChkPrint.Value = Abs(ParaInPrint)
'   If ParaInChoice = "Cash" Or ParaInChoice = "" Then
'      OptCash.Value = True
'      Call OptCash_Click
'   ElseIf ParaInChoice = "Credit" Then
'      OptCredit.Value = True
'      Call OptCredit_Click
'   ElseIf ParaInChoice = "Bank" Then
'      OptBankCard.Value = True
'      Call OptBankCard_Click
'   End If
   If TxtCashCustomer.Text = "" Then TxtCashCustomer.Text = "Counter Sale"
   If vIsRemarksCompulsory = True Then TxtCashCustomer.Text = ""
   If TxtBankCustomer.Text = "" Then TxtBankCustomer.Text = "Counter Sale"
'   If OptCredit.Value = True Then If FunValidation = False Then Exit Sub
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Function ReplaceSMS(vStr As String) As String
   Dim StrSMS As String, Count As Integer, Start As Integer, i As Integer
   Dim distinctChr As String, StrReplace As String, StrColumn As String
   distinctChr = "["
   StrSMS = vStr
   Count = Len(StrSMS) - Len(Replace(StrSMS, distinctChr, ""))
   Start = 1
   StrReplace = StrSMS
   For i = 1 To Count
      StrColumn = Mid(StrSMS, FindFirst(StrSMS, Start) + 1, FindSecond(StrSMS, Start) - FindFirst(StrSMS, Start) - 1)
      If UCase(StrColumn) = UCase("BillDate") Then
'         vStr = "Select CONVERT( varchar(20),billdate,3 ) BillDate from saleheader Where Billid = " & TxtBillID.Text & " and billdate = '" & DtpBillDate.DateValue & "'"
         vStr = "Select CONVERT( varchar(20),billdate,3 ) BillDate from saleheader Where SID = " & TxtSID.Text
      Else
'         vStr = "Select " & StrColumn & " from saleheader Where Billid = " & TxtBillID.Text & " and billdate = '" & DtpBillDate.DateValue & "'"
          vStr = "Select " & StrColumn & " from saleheader Where SID = " & TxtSID.Text
      End If
      
      StrReplace = Replace(StrReplace, "[" & StrColumn & "]", cn.Execute(vStr).Fields(0))
      Start = FindSecond(StrSMS, Start) + 1
   Next
   ReplaceSMS = StrReplace
End Function

Private Function FindFirst(vStr As String, Start As Integer) As Integer
   FindFirst = InStr(Start, vStr, "[")
End Function

Private Function FindSecond(vStr As String, Start As Integer) As Integer
   FindSecond = InStr(Start, vStr, "]")
End Function

Private Sub ActivityLogSale(FormType As String, Mode As EntryMode, Optional Key1 As Long = 0, Optional Key2 As Date = "01-01-1900", Optional Key3 As String = "")
   Dim vSQL As String
   vSQL = "Exec ProdActivityLog '" & FormType & "'," & ObjUserSecurity.UserNo & "," & Mode & "," & Key1 & ",'" & Key2 & "','" & Key3 & "'"
   'vSQL = "INSERT into ActivityLogSale(userno,FormType,EntryDate,Description,isnew,isedit,isdelete) values(" & ObjUserSecurity.UserNo & ",'" & FormType & "',getdate(),'" & Desc & "'," & IIf(Mode = eAdd, 1, 0) & "," & IIf(Mode = eEdit, 1, 0) & "," & IIf(Mode = eDelete, 1, 0) & ")"
   cn.Execute vSQL
End Sub

Private Sub WebSMS(vStrUrl As String, vMessage As String, vCustNo As String)
   Dim postData As String, SendSMS_Text As String
   Dim winHttpReq As Object
   Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
   'http://sms4connect.com/api/sendsms.php/sendsms/url?id=alsanagarments&pass=11221122az&mask=alsana&to=923346136881&lang=English&msg=Hello%20Customer%20&type=xml
   vStrUrl = Replace(vStrUrl, "[ToNumber]", vCustNo)
   vStrUrl = Replace(vStrUrl, "[Message]", vMessage)
   vStrUrl = Replace(vStrUrl, " ", "%20")
   winHttpReq.Open "GET", vStrUrl, False
   winHttpReq.Send
   SendSMS_Text = winHttpReq.responseText
'   winHttpReq.Open "POST", vStrUrl, False
'   winHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'   winHttpReq.Send (postData)
'   SendSMS_Text = winHttpReq.responseText
End Sub

Private Sub TxtAdvTaxPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtAdvTaxPer.Name Then Exit Sub
   TxtAdvTaxVal.Text = SelfRound((Val(TxtSumDiscAmount.Text) * Val(TxtAdvTaxPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtAdvTaxVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtAdvTaxVal.Name Then Exit Sub
   TxtAdvTaxPer.Text = Round((Val(TxtAdvTaxVal.Text) * 100) / IIf(Val(TxtSumDiscAmount.Text) = 0, 1, Val(TxtSumDiscAmount.Text)), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectBankMachineCredit(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchBankMachine.Show vbModal, Me
        If SchBankMachine.ParaOutBankMachineID = "" Then FunSelectBankMachineCredit = False: Exit Function
        TxtBankMachineCreditID.Text = SchBankMachine.ParaOutBankMachineID
    End If
    '---------------------------
    vStrSQL = "Select * FROM BankMachines where BankMachineID=" & Val(TxtBankMachineCreditID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtBankMachineCreditName.Text = !BankMachineName
          TxtCommision.Text = !Commision
          FunSelectBankMachineCredit = True
          .Close
          Exit Function
      Else
          FunSelectBankMachineCredit = False
          .Close
          TxtBankMachineCreditID.Text = ""
          TxtBankMachineCreditName.Text = ""
          TxtCommision.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function



Private Sub BtnBankMachineCredit_Click()
   On Error GoTo ErrorHandler
   If FunSelectBankMachineCredit(ssButton, False) = True Then
      TxtBankAmount.SetFocus
   Else
      TxtBankMachineCreditID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtBankMachineCreditID_Change()
   If TxtBankMachineCreditID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtBankMachineCreditID.Name Then Exit Sub
   If TxtBankMachineCreditName.Text <> "" Then
      TxtBankMachineCreditName.Text = ""
      TxtCommision.Text = ""
   End If
End Sub

Private Sub TxtBankMachineCreditID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtBankMachineCreditID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtBankMachineCreditName.Text <> "" Then Exit Sub
   If Trim(TxtBankMachineCreditID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBankMachineCredit(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBankMachineCredit(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub HeaderTableColumn()
'vStrPara = ""
'vStrPara = vStrPara & "BillID, "
'vStrPara = vStrPara & "BillDate, "
'vStrPara = vStrPara & "CustomerID, "
'vStrPara = vStrPara & "TotalAmount, "
'vStrPara = vStrPara & "BillDisc, "
'vStrPara = vStrPara & "CashReceived, "
'vStrPara = vStrPara & "UserNo, "
'vStrPara = vStrPara & "StoreID, "
'vStrPara = vStrPara & "BankCard, "
'vStrPara = vStrPara & "Credit, "
'vStrPara = vStrPara & "Cash, "
'vStrPara = vStrPara & "BankMachineID, "
'vStrPara = vStrPara & "InvoiceNo, "
'vStrPara = vStrPara & "CustomerName, "
'vStrPara = vStrPara & "BillDiscPer, "
'vStrPara = vStrPara & "Commision, "
'vStrPara = vStrPara & "EmpComm, "
'vStrPara = vStrPara & "EmpID, "
'vStrPara = vStrPara & "isReplace, "
'vStrPara = vStrPara & "isPosted, "
'vStrPara = vStrPara & "MemberID, "
'vStrPara = vStrPara & "BillTime, "
'vStrPara = vStrPara & "Tag, "
'vStrPara = vStrPara & "ManualBillNo, "
'vStrPara = vStrPara & "Remarks, "
'vStrPara = vStrPara & "OrganizationID, "
'vStrPara = vStrPara & "BillNO, "
'vStrPara = vStrPara & "BILTYNO, "
'vStrPara = vStrPara & "DESCRIPTION, "
'vStrPara = vStrPara & "PAIDAMOUNT, "
'vStrPara = vStrPara & "EntryDate, "
'vStrPara = vStrPara & "PreviousAmount, "
'vStrPara = vStrPara & "OtherCharges, "
'vStrPara = vStrPara & "SaleManID, "
'vStrPara = vStrPara & "TotalExpense, "
'vStrPara = vStrPara & "OrderID, "
'vStrPara = vStrPara & "OrderDate, "
'vStrPara = vStrPara & "Freight, "
'vStrPara = vStrPara & "IsCustomerFreight, "
'vStrPara = vStrPara & "VehicleNo, "
'vStrPara = vStrPara & "ServiceCharges, "
'vStrPara = vStrPara & "ServiceChargesPer, "
'vStrPara = vStrPara & "STax, "
'vStrPara = vStrPara & "STaxPer, "
'vStrPara = vStrPara & "TableID, "
'vStrPara = vStrPara & "ServerEntry, "
'vStrPara = vStrPara & "InvType, "
'vStrPara = vStrPara & "DeliveryDate, "
'vStrPara = vStrPara & "DeliveryTime, "
'vStrPara = vStrPara & "isPrinted, "
'vStrPara = vStrPara & "RemarksUrdu, "
'vStrPara = vStrPara & "isTransfer, "
'vStrPara = vStrPara & "PromiseDate, "
'vStrPara = vStrPara & "SyllabusID, "
'vStrPara = vStrPara & "Modified_On, "
'vStrPara = vStrPara & "SessionID, "
'vStrPara = vStrPara & "ExpiryInvoice, "
'vStrPara = vStrPara & "SID, "
'vStrPara = vStrPara & "AdvTaxVal, "
'vStrPara = vStrPara & "AdvTaxPer, "
'vStrPara = vStrPara & "ExtraTaxVal, "
'vStrPara = vStrPara & "ExtraTaxPer, "
'vStrPara = vStrPara & "CNIC, "
'vStrPara = vStrPara & "MobileNo, "
'vStrPara = vStrPara & "SumDiscAmount, "
'vStrPara = vStrPara & "DispatchDate, "
'vStrPara = vStrPara & "Terms, "
'vStrPara = vStrPara & "RefID, "
'vStrPara = vStrPara & "RefComm, "
'vStrPara = vStrPara & "BankAmount, "
'vStrPara = vStrPara & "FBRInvoiceNo"
'End Sub
'
'Private Sub BodyTableColumn()
'vStrPara = ""
'vStrPara = vStrPara & "BillID, "
'vStrPara = vStrPara & "BillDate, "
'vStrPara = vStrPara & "ProductID, "
'vStrPara = vStrPara & "Qty, "
'vStrPara = vStrPara & "Price, "
'vStrPara = vStrPara & "DiscPC, "
'vStrPara = vStrPara & "Amount, "
'vStrPara = vStrPara & "Code, "
'vStrPara = vStrPara & "DiscPer, "
'vStrPara = vStrPara & "DiscVal, "
'vStrPara = vStrPara & "Cost, "
'vStrPara = vStrPara & "isProduct, "
'vStrPara = vStrPara & "PackingID, "
'vStrPara = vStrPara & "QtyPack, "
'vStrPara = vStrPara & "Multiplier, "
'vStrPara = vStrPara & "Bonus, "
'vStrPara = vStrPara & "Offer, "
'vStrPara = vStrPara & "SaleTaxPer, "
'vStrPara = vStrPara & "SaleTaxVal, "
'vStrPara = vStrPara & "TokenVal, "
'vStrPara = vStrPara & "RetailPrice, "
'vStrPara = vStrPara & "IsWSSaleTax, "
'vStrPara = vStrPara & "IsRetailSaleTax, "
'vStrPara = vStrPara & "IsWSDiscb4ST, "
'vStrPara = vStrPara & "SC, "
'vStrPara = vStrPara & "EmpComm, "
'vStrPara = vStrPara & "BatchNo, "
'vStrPara = vStrPara & "StoreID, "
'vStrPara = vStrPara & "EmpID, "
'vStrPara = vStrPara & "ColourID, "
'vStrPara = vStrPara & "SizeID, "
'vStrPara = vStrPara & "GrossQty, "
'vStrPara = vStrPara & "GrossUnit, "
'vStrPara = vStrPara & "HeaderStoreID, "
'vStrPara = vStrPara & "Modified_On, "
'vStrPara = vStrPara & "isDiscB4TradeOffer, "
'vStrPara = vStrPara & "isDiscB4ExtraScheme, "
'vStrPara = vStrPara & "isDiscB4SaleTax, "
'vStrPara = vStrPara & "TradeOffer1, "
'vStrPara = vStrPara & "TradeOffer2, "
'vStrPara = vStrPara & "ExtraSchemePer, "
'vStrPara = vStrPara & "TradeValue, "
'vStrPara = vStrPara & "ExtraSchemeValue, "
'vStrPara = vStrPara & "SID, "
'vStrPara = vStrPara & "DiscAmount, "
'vStrPara = vStrPara & "isLastPrice, "
'vStrPara = vStrPara & "ReSPrice, "
'vStrPara = vStrPara & "ReSAmount "
'End Sub
'
Private Sub BinData(vActionDesc As String, vActionNO As Integer)
'   If ObjRegistry.UseBin = True Then
'      Call HeaderTableColumn
'      vBinSql = "Insert Into " & vBinDataBase & ".dbo.SaleHeaderBin (BinDate, ActionNo, ActionUserNo, ActionDesc, " & vStrPara & ")" & vbCrLf _
'             & "Select '" & Now & "', " & vActionNO & ", " & vUser & ", '" & vActionDesc & "', " & vStrPara & " from SaleHeader " & vbCrLf _
'             & "Where SID = " & TxtSID.Text
'      cn.Execute vBinSql
'      Call BodyTableColumn
'      vBinSql = "Insert Into " & vBinDataBase & ".dbo.SaleBodyBin (" & vStrPara & ")" & vbCrLf _
'             & "Select " & vStrPara & " from SaleBody " & vbCrLf _
'             & "Where SID = " & TxtSID.Text
'      cn.Execute vBinSql
'   End If
End Sub
