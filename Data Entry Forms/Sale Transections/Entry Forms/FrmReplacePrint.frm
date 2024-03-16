VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmReplacePrint 
   BorderStyle     =   0  'None
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkPrint 
      Caption         =   "&Print"
      Height          =   285
      Left            =   780
      TabIndex        =   18
      Top             =   270
      Value           =   1  'Checked
      Width           =   705
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   1680
      TabIndex        =   13
      Top             =   90
      Width           =   3525
      Begin VB.OptionButton OptCredit 
         Caption         =   "&Credit"
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton OptBankCard 
         Caption         =   "&Bank Card"
         Height          =   285
         Left            =   2100
         TabIndex        =   2
         Top             =   240
         Width           =   1125
      End
      Begin VB.OptionButton OptCash 
         Caption         =   "&Cash"
         Height          =   285
         Left            =   210
         TabIndex        =   0
         Top             =   240
         Width           =   765
      End
   End
   Begin JeweledBut.JeweledButton BtnCancel 
      Height          =   420
      Left            =   3630
      TabIndex        =   12
      Top             =   3600
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
      MICON           =   "FrmReplacePrint.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOk 
      Height          =   420
      Left            =   2325
      TabIndex        =   11
      Top             =   3600
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
      MICON           =   "FrmReplacePrint.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Frame FrameCredit 
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   540
      TabIndex        =   21
      Top             =   765
      Width           =   6105
      Begin VB.TextBox TxtNetAmountCredit 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   180
         Width           =   2025
      End
      Begin SITextBox.Txt TxtCustomerID 
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Appearance      =   0
         MaxLength       =   13
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
         TabIndex        =   24
         Top             =   1380
         Width           =   4350
         _ExtentX        =   7673
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
         TabIndex        =   25
         TabStop         =   0   'False
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
         MICON           =   "FrmReplacePrint.frx":0038
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtCashReceivedCredit 
         Height          =   315
         Left            =   1485
         TabIndex        =   5
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
      Begin SITextBox.Txt TxtBankMachineCreditID 
         Height          =   315
         Left            =   135
         TabIndex        =   41
         Top             =   2055
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
         Left            =   1830
         TabIndex        =   42
         Top             =   2055
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
         Left            =   1470
         TabIndex        =   43
         TabStop         =   0   'False
         Tag             =   "F"
         Top             =   2055
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
         MICON           =   "FrmReplacePrint.frx":0054
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtBankAmount 
         Height          =   315
         Left            =   4185
         TabIndex        =   44
         Top             =   2055
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
         Left            =   135
         TabIndex        =   47
         Top             =   1845
         Width           =   1485
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
         Left            =   1830
         TabIndex        =   46
         Top             =   1845
         Width           =   1770
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
         Left            =   4185
         TabIndex        =   45
         Top             =   1845
         Width           =   1110
      End
      Begin VB.Label LblCreditCaption 
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
         Left            =   45
         TabIndex        =   38
         Top             =   705
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   195
         Left            =   1755
         TabIndex        =   26
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Label5 
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
         TabIndex        =   23
         Top             =   210
         Width           =   1005
      End
   End
   Begin VB.Frame FrameBank 
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   495
      TabIndex        =   28
      Top             =   720
      Width           =   6195
      Begin VB.TextBox TxtCommision 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4725
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   855
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox TxtNetAmountBank 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   720
         Width           =   2025
      End
      Begin SITextBox.Txt TxtBankMachineID 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   2265
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
         Left            =   1755
         TabIndex        =   30
         Top             =   2265
         Width           =   4350
         _ExtentX        =   7673
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
         Left            =   1395
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2265
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
         MICON           =   "FrmReplacePrint.frx":0070
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtInvoiceNo 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   1200
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
      Begin SITextBox.Txt TxtBankCustomer 
         Height          =   315
         Left            =   675
         TabIndex        =   7
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
      Begin SITextBox.Txt TxtCashReceivedBank 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Tag             =   "F"
         Top             =   1650
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
      Begin VB.Label LblBankCaption 
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
         Left            =   180
         TabIndex        =   40
         Top             =   1695
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         TabIndex        =   37
         Top             =   45
         Width           =   1665
      End
      Begin VB.Label Label9 
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
         TabIndex        =   35
         Top             =   1230
         Width           =   945
      End
      Begin VB.Label Label8 
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
         Left            =   390
         TabIndex        =   34
         Top             =   750
         Width           =   1005
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Machine Name"
         Height          =   195
         Left            =   1755
         TabIndex        =   33
         Top             =   2055
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Machine ID"
         Height          =   195
         Left            =   60
         TabIndex        =   32
         Top             =   2055
         Width           =   1245
      End
   End
   Begin VB.Frame FrameCash 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   1215
      TabIndex        =   14
      Top             =   855
      Width           =   4425
      Begin VB.TextBox TxtNetAmount 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   2025
      End
      Begin VB.TextBox TxtCashReturn 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1770
         Width           =   2025
      End
      Begin SITextBox.Txt TxtCashCustomer 
         Height          =   315
         Left            =   30
         TabIndex        =   3
         Top             =   360
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
      Begin SITextBox.Txt TxtCashReceivedCash 
         Height          =   315
         Left            =   1470
         TabIndex        =   4
         Top             =   1365
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
      Begin VB.Label Label10 
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
         TabIndex        =   36
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
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
         Left            =   375
         TabIndex        =   20
         Top             =   990
         Width           =   1005
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
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
         Left            =   75
         TabIndex        =   19
         Top             =   1395
         Width           =   1305
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   315
         TabIndex        =   17
         Top             =   1800
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmReplacePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ParaOutSelection As Boolean
Public ParaInPrint As Boolean
Public ParaInChoice As String
Public ParaInID As String
Public ParaInOgtanizationID As String
Public ParaInDate As Date
Dim vPrevious As Double
Dim vCurrent As Double

Public Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is SITextBox.txt Then
         ctl.Text = ""
      ElseIf TypeOf ctl Is TextBox Then
         ctl.Text = ""
      End If
   Next
   OptCash.Value = True
   TxtCashCustomer.Text = "Counter Sale"
   TxtBankCustomer.Text = "Counter Sale"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

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
      If Trim(TxtCustomerID.Text) = "" Then
         MsgBox "Please specify a Customer ID", vbExclamation, "Alert"
         TxtCustomerID.SetFocus
         Exit Function
      End If
      If Trim(TxtBankMachineCreditID.Text) <> "" Then
         If Val(TxtBankAmount.Text) = 0 Then
            MsgBox "Please Enter Bank Amount", vbExclamation, "Alert"
            Exit Function
         End If
'         If Val(TxtBankAmount.Text) + Val(TxtCashReceivedBank.Text) > Val(TxtNetAmountCredit.Text) Then
'            MsgBox "Amount Greater than Net Amount", vbExclamation, "Alert"
'            TxtBankAmount.SetFocus
'            Exit Function
'         End If
      End If
      If Val(TxtBankAmount.Text) > 0 Then
          If Trim(TxtBankMachineCreditID.Text) = "" Then
            MsgBox "Please specify a Bank Machine ID OR Bank Amount is 0", vbExclamation, "Alert"
            TxtBankMachineCreditID.SetFocus
            Exit Function
          End If
      End If
      With CN.Execute("Select * from Employees where CreditLimit <> 0 and CreditLimit is not null and EmpID = '" & TxtCustomerID.Text & "'")
         If .RecordCount > 0 Then
            vPrevious = CN.Execute("SELECT isnull(sum(TotalAmount) - isnull(sum(BillDisc),0) - isnull(sum(CashReceived),0),0) as Balance FROM SaleHeader where month(BillDate)=" & Month(ParaInDate) & " and year(BillDate)=" & Year(ParaInDate) & " and CustomerID = '" & TxtCustomerID.Text & "'").Fields(0).Value
            vCurrent = CN.Execute("Select sum(TotalAmount-isnull(BillDisc,0)+isnull(STax,0)+isnull(ServiceCharges,0)-isnull(CashReceived,0)) from SaleHeader where BillID = " & ParaInID & " and BillDate = '" & ParaInDate & "'").Fields(0).Value
            If !CreditLimit < (vPrevious - vCurrent + Val(TxtNetAmountCredit.Text) - Val(TxtCashReceivedCredit.Text)) Then
               MsgBox "Credit Limit (" & !CreditLimit & ") is Exceed Balance (" & vPrevious & ") in this month for this Employee.", vbExclamation, "Alert"
               Exit Function
            End If
         End If
      End With
      With CN.Execute("Select * from Parties where CreditLimit <> 0 and CreditLimit is not null and PartyID = '" & TxtCustomerID.Text & "'")
         If .RecordCount > 0 Then
            vPrevious = CN.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtCustomerID.Text & "','" & DateAdd("d", 0, ParaInDate) & "'," & IIf(Val(ParaInOgtanizationID) = 0, "Null", Val(ParaInOgtanizationID)) & "),0) ").Fields(0).Value
            vCurrent = CN.Execute("Select isnull(sum(TotalAmount-isnull(BillDisc,0)+isnull(STax,0)+isnull(ServiceCharges,0)-isnull(CashReceived,0)),0) from SaleHeader where BillID = " & ParaInID & " and BillDate = '" & ParaInDate & "'").Fields(0).Value
            If !CreditLimit < (vPrevious - vCurrent + Val(TxtNetAmountCredit.Text) - Val(TxtCashReceivedCredit.Text)) Then
               MsgBox "Credit Limit (" & !CreditLimit & ") is Exceed Balance (" & vPrevious & ") in this month for this Customer.", vbExclamation, "Alert"
               Exit Function
            End If
         End If
      End With
      With CN.Execute("Select * from Members where CreditLimit <> 0 and CreditLimit is not null and MemberID = substring('" & TxtCustomerID.Text & "',3,10)")
         If .RecordCount > 0 Then
            vPrevious = CN.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtCustomerID.Text & "','" & DateAdd("d", 0, ParaInDate) & "'," & IIf(Val(ParaInOgtanizationID) = 0, "Null", Val(ParaInOgtanizationID)) & "),0) ").Fields(0).Value
            vCurrent = CN.Execute("Select isnull(sum(TotalAmount-isnull(BillDisc,0)+isnull(STax,0)+isnull(ServiceCharges,0)-isnull(CashReceived,0)),0) from SaleHeader where BillID = " & ParaInID & " and BillDate = '" & ParaInDate & "'").Fields(0).Value
            If !CreditLimit < (vPrevious - vCurrent + Val(TxtNetAmountCredit.Text) - Val(TxtCashReceivedCredit.Text)) Then
               MsgBox "Credit Limit (" & !CreditLimit & ") is Exceed Balance (" & vPrevious & ") in this month for this Customer.", vbExclamation, "Alert"
               Exit Function
            End If
         End If
      End With
   End If
   If OptCash.Value = True Then
      If Val(TxtNetAmount.Text) <> 0 Then
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
    vStrSQL = " Select * FROM BankMachines where BankMachineID=" & Val(TxtBankMachineID.Text)
    With CN.Execute(vStrSQL)
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

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.CmbFilter = "Customers"
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo like '3%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select c.* FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " Left Outer join Members m on c.AccountNo = cast(m.Prefix as varchar(2))  + cast(m.MemberID as varchar(10)) " & vbCrLf & _
              " where p.BarCode = '" & (TxtCustomerID.Text) & "' or m.BarCode = '" & (TxtCustomerID.Text) & "' or (c.AccountNo = '" & (TxtCustomerID.Text) & "' and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo like '3%') and c.isDetailed = 1 and c.isLocked = 0)"
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

Private Sub BtnBankMachine_Click()
   If FunSelectBankMachine(ssButton, False) = True Then
      BtnOk.SetFocus
   Else
      TxtBankMachineID.SetFocus
   End If
End Sub

Private Sub BtnCancel_Click()
   On Error GoTo ErrorHandler
   ParaOutSelection = False
   Me.Hide
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      BtnOk.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

Private Sub BtnOk_Click()
   On Error GoTo ErrorHandler
   If FunValidation = False Then Exit Sub
   ParaOutSelection = True
   Me.Hide
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkPrint_Click()
   If ChkPrint.Value = 1 Then
      ParaInPrint = True
   Else
      ParaInPrint = False
   End If
End Sub

Private Sub Form_Activate()

   If ObjUserSecurity.IsDisableCreditSale = True Then OptCredit.Enabled = False
'   If ObjUserSecurity.IsCreditSale = False Then OptCredit.Enabled = False
'   OptCredit.Value = ObjUserSecurity.IsCreditSale
    If OptCredit.Enabled = False Then
      OptCredit.Value = False
      OptCash.Value = True
    End If
   ''''''''''''''''''''
   
   If OptCash.Value = True And OptCash.Visible = True Then OptCash.SetFocus: Call OptCash_Click
   If OptCredit.Value = True And OptCredit.Visible = True Then OptCredit_Click: OptCredit.SetFocus
   If OptBankCard.Value = True And OptBankCard.Visible = True Then OptBankCard_Click: OptBankCard.SetFocus
   ChkPrint.Value = Abs(ParaInPrint)
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
   If Shift = vbCtrlMask And KeyCode = vbKeyReturn Then
      BtnOk_Click
   ElseIf KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyEscape Then
      BtnCancel_Click
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnOk.Enabled Then BtnOk_Click
            KeyCode = 0
         Case vbKeyW
            If BtnCancel.Enabled Then BtnCancel_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtBankMachineID.Name: If FunSelectBankMachine(ssFunctionKey, True) = True Then BtnOk.SetFocus
         Case TxtCustomerID.Name: If FunSelectCustomer(ssFunctionKey, False) = True Then BtnOk.SetFocus
         Case TxtBankMachineCreditID.Name: If FunSelectBankMachineCredit(ssFunctionKey, True) = True Then TxtBankAmount.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_Load()
''''' Edit by Rizwan
   FrameBank.Top = 765
   FrameCash.Top = 765
   FrameCredit.Top = 765
   ChkPrint.Value = Abs(ParaInPrint)
   If ParaInChoice = "Cash" Or ParaInChoice = "" Then
      OptCash.Value = True
      Call OptCash_Click
   ElseIf ParaInChoice = "Credit" Then
      OptCredit.Value = True
      Call OptCredit_Click
   ElseIf ParaInChoice = "Bank" Then
      OptBankCard.Value = True
      Call OptBankCard_Click
   End If
   TxtCashCustomer.Text = "Counter Sale"
   TxtBankCustomer.Text = "Counter Sale"
   ChkPrint.Enabled = Not ObjRegistry.HideAutoPrint
End Sub

Private Sub OptBankCard_Click()
   If OptBankCard.Visible = False Then Exit Sub
   If ActiveControl.Name <> OptBankCard.Name Then Exit Sub
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
   If OptCash.Visible = False Then Exit Sub
   If ActiveControl.Name <> OptCash.Name Then Exit Sub
   FrameCash.Visible = True
   FrameCredit.Visible = False
   FrameBank.Visible = False
   TxtCashCustomer.Text = IIf(TxtCashCustomer.Text = "", "Counter Sale", TxtCashCustomer.Text)
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
   TxtCashReturn.Text = Val(TxtCashReceivedCash.Text) - Val(TxtNetAmount.Text)
End Sub

Private Sub TxtNetAmount_Change()
   Call TxtCashReceivedCash_Change
   TxtNetAmountBank.Text = TxtNetAmount.Text
   TxtNetAmountCredit.Text = TxtNetAmount.Text
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
    With CN.Execute(vStrSQL)
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


