VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSalary 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmSalary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2850
      Left            =   12480
      TabIndex        =   48
      Top             =   1560
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
         Height          =   2445
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   49
         Tag             =   "NC"
         Text            =   "FrmSalary.frx":0ECA
         Top             =   360
         Width           =   3930
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
         TabIndex        =   50
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8956
      TabIndex        =   30
      Top             =   8618
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
      MICON           =   "FrmSalary.frx":0F55
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7636
      TabIndex        =   26
      Top             =   8618
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
      MICON           =   "FrmSalary.frx":0F71
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4996
      TabIndex        =   28
      Top             =   8618
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
      MICON           =   "FrmSalary.frx":0F8D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10276
      TabIndex        =   31
      Top             =   8618
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
      MICON           =   "FrmSalary.frx":0FA9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6316
      TabIndex        =   27
      Top             =   8618
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
      MICON           =   "FrmSalary.frx":0FC5
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpMonth 
      Height          =   345
      Left            =   3233
      TabIndex        =   7
      Top             =   4778
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "MMM-yyyy"
      Format          =   114688003
      CurrentDate     =   38595
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3646
      TabIndex        =   29
      Top             =   8618
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
      MICON           =   "FrmSalary.frx":0FE1
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpEntryDate 
      Height          =   315
      Left            =   3143
      TabIndex        =   0
      Top             =   2348
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   114688003
      CurrentDate     =   38595
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   3188
      TabIndex        =   1
      Top             =   3083
      Width           =   930
      _ExtentX        =   1640
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtEmployeeName 
      Height          =   315
      Left            =   4478
      TabIndex        =   3
      Top             =   3083
      Width           =   2880
      _ExtentX        =   5080
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
   Begin SITextBox.Txt TxtAddress 
      Height          =   315
      Left            =   3143
      TabIndex        =   6
      Top             =   3923
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin SITextBox.Txt TxtFName 
      Height          =   315
      Left            =   7358
      TabIndex        =   4
      Top             =   3083
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
      Left            =   4118
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3068
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
      MICON           =   "FrmSalary.frx":0FFD
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtDesignation 
      Height          =   315
      Left            =   10238
      TabIndex        =   5
      Top             =   3083
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtSalary 
      Height          =   315
      Left            =   4373
      TabIndex        =   8
      Top             =   4793
      Width           =   1380
      _ExtentX        =   2434
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
      Masked          =   5
   End
   Begin SITextBox.Txt TxtTTLWorkingDays 
      Height          =   315
      Left            =   5753
      TabIndex        =   9
      Top             =   4793
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtWorkingDays 
      Height          =   315
      Left            =   3233
      TabIndex        =   11
      Top             =   5663
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   5
   End
   Begin SITextBox.Txt TxtSalaryOneDay 
      Height          =   315
      Left            =   6398
      TabIndex        =   13
      Top             =   5663
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
      Masked          =   5
   End
   Begin SITextBox.Txt TxtSalaryPerHrs 
      Height          =   315
      Left            =   7778
      TabIndex        =   14
      Top             =   5663
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
      Masked          =   5
   End
   Begin SITextBox.Txt TxtTTLWorkingHours 
      Height          =   315
      Left            =   7253
      TabIndex        =   10
      Top             =   4793
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtHoursPerDay 
      Height          =   315
      Left            =   5783
      TabIndex        =   56
      Top             =   2348
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtWorkingHours 
      Height          =   315
      Left            =   4853
      TabIndex        =   12
      Top             =   5663
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   4
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
   Begin SITextBox.Txt TxtTTLSalary 
      Height          =   315
      Left            =   9038
      TabIndex        =   15
      Top             =   5663
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtSalaryID 
      Height          =   315
      Left            =   4463
      TabIndex        =   59
      Top             =   2348
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin SITextBox.Txt TxtRemainingLoan 
      Height          =   315
      Left            =   7118
      TabIndex        =   18
      Top             =   6503
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtLoanInstallment 
      Height          =   315
      Left            =   5093
      TabIndex        =   17
      Top             =   6503
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtPreviousLoan 
      Height          =   315
      Left            =   3233
      TabIndex        =   16
      Top             =   6503
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtPrevious 
      Height          =   315
      Left            =   3233
      TabIndex        =   20
      Top             =   7403
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtAdvance 
      Height          =   315
      Left            =   6293
      TabIndex        =   22
      Top             =   7403
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtLess 
      Height          =   315
      Left            =   9323
      TabIndex        =   24
      Top             =   7403
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtTotal 
      Height          =   315
      Left            =   10838
      TabIndex        =   25
      Top             =   7403
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtItemsPurchase 
      Height          =   315
      Left            =   7808
      TabIndex        =   23
      Top             =   7403
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtSaleCommision 
      Height          =   315
      Left            =   4778
      TabIndex        =   21
      Top             =   7403
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   7748
      TabIndex        =   66
      Tag             =   "NC"
      Top             =   2288
      Width           =   945
      _ExtentX        =   1667
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
      Left            =   9053
      TabIndex        =   67
      Tag             =   "NC"
      Top             =   2288
      Width           =   1980
      _ExtentX        =   3493
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
      Left            =   8693
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   2288
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
      MICON           =   "FrmSalary.frx":1019
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBonus 
      Height          =   315
      Left            =   9090
      TabIndex        =   19
      Top             =   6503
      Width           =   1380
      _ExtentX        =   2434
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
      Masked          =   5
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bonus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9090
      TabIndex        =   71
      Top             =   6233
      Width           =   660
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
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
      Left            =   7658
      TabIndex        =   70
      Top             =   2093
      Width           =   1335
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
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
      Left            =   9053
      TabIndex        =   69
      Top             =   2093
      Width           =   1620
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Commision"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4508
      TabIndex        =   65
      Top             =   7133
      Width           =   1695
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Items Purchase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7793
      TabIndex        =   64
      Top             =   6908
      Width           =   1035
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Loan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3233
      TabIndex        =   63
      Top             =   6233
      Width           =   1500
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Installment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5018
      TabIndex        =   62
      Top             =   6233
      Width           =   1695
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Loan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7148
      TabIndex        =   61
      Top             =   6233
      Width           =   1695
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Salary ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4463
      TabIndex        =   60
      Top             =   2108
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hours Worked "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4853
      TabIndex        =   58
      Top             =   5393
      Width           =   1560
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hours / Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5798
      TabIndex        =   57
      Top             =   2108
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7268
      TabIndex        =   55
      Top             =   4508
      Width           =   1545
   End
   Begin VB.Label HoursPerDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary / Hrs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7778
      TabIndex        =   54
      Top             =   5393
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Days Worked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3233
      TabIndex        =   53
      Top             =   5393
      Width           =   1425
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   52
      Top             =   270
      Width           =   1110
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
      Left            =   11250
      TabIndex        =   51
      Top             =   765
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10238
      TabIndex        =   47
      Top             =   2828
      Width           =   1260
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3173
      TabIndex        =   46
      Top             =   2828
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4478
      TabIndex        =   45
      Top             =   2828
      Width           =   1155
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3173
      TabIndex        =   44
      Top             =   3668
      Width           =   885
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7358
      TabIndex        =   43
      Top             =   2828
      Width           =   1350
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3233
      TabIndex        =   42
      Top             =   4508
      Width           =   630
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3143
      TabIndex        =   41
      Top             =   2108
      Width           =   1095
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10838
      TabIndex        =   40
      Top             =   7133
      Width           =   1110
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9323
      TabIndex        =   39
      Top             =   7133
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6293
      TabIndex        =   38
      Top             =   7133
      Width           =   930
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3233
      TabIndex        =   37
      Top             =   7133
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4373
      TabIndex        =   36
      Top             =   4508
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9038
      TabIndex        =   35
      Top             =   5393
      Width           =   1290
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary / Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6398
      TabIndex        =   34
      Top             =   5393
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5753
      TabIndex        =   33
      Top             =   4508
      Width           =   1470
   End
   Begin VB.Label LblAdvance 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   270
      Left            =   3548
      TabIndex        =   32
      Top             =   8558
      Width           =   60
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11595
      Top             =   45
      Width           =   375
   End
End
Attribute VB_Name = "FrmSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim RsHeader As ADODB.Recordset
Dim vMode As FormMode
Dim vid As String
Dim sSql As String, vStrSQL As String, SqlWorking As String
Dim vCounter As Integer
Dim vCounter1 As Integer
Dim PreviousDate As Date

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

Private Sub TxtBonus_Change()
On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtBonus.Name Then Exit Sub
      SubCalculateSalary
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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

Private Sub BtnOrganization_Click()
   If FunSelectOrganization(ssButton, False) = True Then
      TxtEmployeeID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub SubPrevious()
   On Error GoTo ErrorHandler
   ' Get Previous Loan
   TxtPreviousLoan.Text = cn.Execute("SELECT isnull(dbo.FunPreviousLoan(" & TxtEmployeeID.Text & ",'" & DtpEntryDate.Value & "'),0)").Fields(0).Value
   
   ' Get Salary Month Advance
   sSql = "Select isnull(sum(Amount),0) from AdvanceVouchersBody b inner join AdvanceVouchers h on b.VoucherNo = h.VoucherNo" & vbCrLf _
        + " left outer join Employees e on b.AccountNo = AdvanceAccountNo" & vbCrLf _
        + " where isnull(EmpID,AccountNo) = '" & TxtEmployeeID.Text & "' and month(VoucherDate) = " & DtpMonth.Month & " and Year(VoucherDate) = " & DtpMonth.Year
   TxtAdvance.Text = cn.Execute(sSql).Fields(0).Value
   
   ' Get Salary Month Sale
   sSql = " Select isnull(round(sum(Amount - isnull(CashReceived,0) - isnull(BillDisc,0) ),0),0) as CreditSale" & vbCrLf _
      + " from SaleHeader h inner join (select BillID, BillDate, sum(Amount) as Amount From SaleBody Group By BillId, BillDate) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
      + " where CustomerID = '" & TxtEmployeeID.Text & "' and month(h.BillDate) = " & DtpMonth.Month & " and Year(h.BillDate) = " & DtpMonth.Year
      
   TxtItemsPurchase.Text = cn.Execute(sSql).Fields(0).Value
   
   ' Get Salary Month Sale Return
   sSql = " Select isnull(round(sum(Amount - isnull(CashPaid,0) - isnull(BillDisc,0) ),0),0) as CreditSale" & vbCrLf _
      + " from SaleReturnHeader h inner join (select ReturnID, ReturnDate, sum(Amount) as Amount From SaleReturnBody Group By ReturnId, ReturnDate) b on h.ReturnID = b.ReturnID and h.ReturnDate = b.ReturnDate " & vbCrLf _
      + " where CustomerID = '" & TxtEmployeeID.Text & "' and month(h.ReturnDate) = " & DtpMonth.Month & " and Year(h.ReturnDate) = " & DtpMonth.Year
      
   TxtItemsPurchase.Text = Val(TxtItemsPurchase.Text) - cn.Execute(sSql).Fields(0).Value
   
   ' Get Month Sale Commision
   sSql = " Select isnull(sum(round(((TotalAmount - isnull(billdisc,0)+ isnull(OtherCharges,0) + isnull(servicecharges,0) + isnull(STax,0)) * isnull(EmpComm,0) * 0.01) + Comm,0)),0) as CreditSale" & vbCrLf _
      + " from SaleHeader h inner join (select BillID, BillDate, sum(Amount) as Amount, isnull(sum(EmpComm*Qty),0) as Comm From SaleBody Group By BillId, BillDate) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
      + " where EmpID = '" & TxtEmployeeID.Text & "' and month(h.BillDate) = " & DtpMonth.Month & " and Year(h.BillDate) = " & DtpMonth.Year
   
   TxtSaleCommision.Text = cn.Execute(sSql).Fields(0).Value
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
   
'   With CN.Execute("Select max(EntryDate)as EntryDate from Salaries where Employeeid = " & TxtEmployeeID.Text)
'      If Not IsNull(!EntryDate) Then
'         PreviousDate = !EntryDate
'      Else
'         PreviousDate = DtpMonth.Value
'      End If
'   End With
'   TxtPrevious.Text = CN.Execute("SELECT isnull(dbo.FunCurrentBalance(" & TxtEmployeeID.Text & ",'" & PreviousDate & "'),0)").Fields(0).Value
'   sSQL = "select Groupid, sum(amount) as Amount from PaymentVouchersBody b inner join PaymentVouchers h on b.voucherno = h.voucherno where voucherdate >='" & PreviousDate & "' and voucherdate<'" & DtpEntryDate.Value & "' and Employeeid = " & TxtEmployeeID.Text & " Group By GroupID"
'   With CN.Execute(sSQL)
'      While Not .EOF
'         Select Case !GroupID
'         Case 2
'            TxtAdvance.Text = !Amount
'         End Select
'         .MoveNext
'      Wend
'   End With
 '  Call SubCalculateSalary
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
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSalary", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
  Dim vtbl As String
  'If RsHeader.RecordCount > 0 Then
    cn.BeginTrans
   Call BinData
   Call ActivityLogBin("", eFrmSalaries, eDelete, TxtSalaryID.Text, DtpEntryDate.Value, "Salary Deleted Amount: " & Val(TxtSalary.Text))
    
    cn.Execute "delete from Salaries where EmpID='" & TxtEmployeeID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'"
    'RsHeader.Requery
    cn.CommitTrans
    Call SubClearFields
    FormStatus = NewMode
  'End If
  Exit Sub
ErrorHandler:
  If cn.Errors.Count > 0 Then cn.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchSalary.Show vbModal
   If SchSalary.ParaOutEmpID <> "" Then
      TxtEmployeeID.Text = SchSalary.ParaOutEmpID
      'Dim a
      'a = Split(SchSalary.ParaOutDate, "/")
      'DtpMonth.Value = Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      DtpMonth.Value = SchSalary.ParaOutDate
      GetSalary
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetSalary()
   Dim n As Integer
   On Error GoTo ErrorHandler
   sSql = "Select s.*, OrganizationName, e.empname,  Address, designation from salaries s inner join Employees e on s.EmpID = e.EmpID" & _
          " left outer join Organizations o on o.OrganizationID = s.OrganizationID Inner Join Designations D on D.Designationid = e.Designationid where s.EmpID='" & TxtEmployeeID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'"
   With cn.Execute(sSql)
      If Not .BOF Then
         TxtSalaryID.Text = !SalaryID
         TxtEmployeeID.Text = !EmpID
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         DtpMonth.Value = !SalaryMonth
         DtpEntryDate.Value = !EntryDate
         TxtEmployeeName.Text = IIf(IsNull(!EmpName), "", !EmpName)
         TxtTTLWorkingDays.Text = !TTLWorkingDays
         TxtTTLWorkingHours.Text = !TTLWorkingHours
         TxtSalaryOneDay.Text = !SalaryOneDay
         TxtSalaryPerHrs.Text = !SalaryPerHour
         TxtWorkingHours.Text = !WorkingHours
         'TxtFName.Text = IIf(IsNull(!FName), "", !FName)
         TxtDesignation.Text = IIf(IsNull(!Designation), "", !Designation)
         TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
         TxtWorkingDays.Text = !WorkingDays
         TxtSalary.Text = !Salary
         TxtBonus.Text = IIf(IsNull(!Bonus), "", !Bonus)
         TxtTTLSalary.Text = !TTLSalary
         TxtLoanInstallment.Text = IIf(IsNull(!LoanInstallment), "", !LoanInstallment)
         TxtPreviousLoan.Text = IIf(IsNull(!PreviousLoan), "", !PreviousLoan)
         TxtPrevious.Text = IIf(IsNull(!Previous), "", !Previous)
         TxtSaleCommision.Text = IIf(IsNull(!SaleCommision), "", !SaleCommision)
         TxtAdvance.Text = IIf(IsNull(!Advance), "", !Advance)
         TxtItemsPurchase.Text = IIf(IsNull(!ItemsPurchase), "", !ItemsPurchase)
         TxtLess.Text = IIf(IsNull(!Less), "", !Less)
         SubCalculateSalary
      End If
      .Close
   End With
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   vStrSQL = " select s.*, s.EmpID + ' - ' + EmpName as Empname, ttlsalary - isnull(s.LoanInstallment,0) + isnull(SaleCommision,0)  - isnull(advance,0) - isnull(ItemsPurchase,0) - isnull(less,0) - isnull(previous,0) as total " & vbCrLf _
      + " from Salaries s " & vbCrLf _
      + " inner join Employees e on e.Empid = s.Empid" & vbCrLf _
      + " where SalaryID = " & Val(TxtSalaryID.Text)
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, cn, adOpenDynamic, adLockReadOnly
   Set RptReportViewer.Report = New CrptEmpSalary
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ReportTitle = "Employee Salary"
   
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy  'CN.Execute("Select Name from Manufacturer").Fields(0).Value

   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
    
   RptReportViewer.Report.PaperOrientation = crPortrait
'   If MsgBox("Do you want to print directly this Salary", vbQuestion + vbYesNo, "Alert") = vbYes Then
   RptReportViewer.Report.PrintOut False
'   Else
'      RptReportViewer.Show vbModal
'   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateWorkingDays()
   On Error GoTo ErrorHandler
   TxtTTLWorkingDays.Text = DateDiff("d", DtpMonth.Value, DateAdd("m", 1, DtpMonth.Value))
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnEmployee_Click()
  If FunSelectEmployee(ssButton, False) = True Then
      If DtpMonth.Enabled Then DtpMonth.SetFocus
   Else
      TxtEmployeeID.SetFocus
   End If
End Sub

Private Sub DtpEntryDate_Change()
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> DtpEntryDate.Name Then Exit Sub
   SubPrevious
   SubCalculateHeaders
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpMonth_Change()
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> DtpMonth.Name Then Exit Sub
   DtpMonth.Day = 1
   GetWorkingDays
   GetWorkingHours
   SubCalculateWorkingDays
   SubCalculateWorkingHours
   If Val(TxtWorkingHours.Text) > 0 Then
      SubCalculateWorkingHours
      SubSalaryPerHoursOrPerDay
      SubCalculateSalaryHrsWise
   Else
      SubSalaryPerHoursOrPerDay
      SubCalculateSalary
      TxtWorkingHours.Text = ""
   End If
   SubPrevious
   SubCalculateHeaders
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Salary"
   HelpLocation Me
   DtpMonth.Value = Date
   DtpEntryDate.Value = Date
   'DtpMonth.Value = DateAdd("d", -1, DateAdd("m", 1, DtpMonth.Value))
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   With cn.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
      End If
      .Close
   End With

   FormStatus = NewMode
'   BtnSave.Visible = Not vReadOnly
'   BtnDelete.Visible = Not vReadOnly
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, True) = True Then If DtpMonth.Enabled Then DtpMonth.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtEmployeeID.SetFocus Else TxtOrganizationID.SetFocus
      End Select
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
         Case vbKeyH
            FraHelp.ZOrder 0
            FraHelp.Visible = True
            KeyCode = 0
         Case vbKeyO
            If BtnOpen.Enabled Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
    
   
  
   
   '  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
'    MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
'    Exit Sub
'  End If
   If FunValidation = False Then Exit Sub
   Set RsHeader = New ADODB.Recordset
   sSql = "Select * FROM Salaries where SalaryMonth = '" & IIf(TxtEmployeeID.Enabled = True, DateAdd("d", -1, DateAdd("m", 1, DtpMonth.Value)), DtpMonth.Value) & "' and Empid='" & TxtEmployeeID.Text & "'"
   With RsHeader
          .Open sSql, cn, adOpenDynamic, adLockReadOnly
        If .RecordCount > 0 Then
           If !SalaryID <> Val(TxtSalaryID.Text) Then
            MsgBox "This employee already taken his salary of given month.", vbExclamation, Me.Caption
            If TxtEmployeeID.Enabled = True Then TxtEmployeeID.SetFocus
           Exit Sub
           End If
        End If
   End With
   RsHeader.Close
   cn.BeginTrans
   
   RsHeader.Open "Select * FROM Salaries where SalaryMonth='" & IIf(TxtEmployeeID.Enabled = True, DateAdd("d", -1, DateAdd("m", 1, DtpMonth.Value)), DtpMonth.Value) & "' and Empid='" & TxtEmployeeID.Text & "'", cn, adOpenDynamic, adLockOptimistic
   
   With RsHeader
      If .RecordCount <> 0 Then Call ActivityLogBin("", eFrmSalaries, eEdit, TxtSalaryID.Text, DtpEntryDate.Value, "Effected Emp Code-" & !EmpID & " Salary-" & Val(!Salary) & " Bonus-" & Val(!Bonus) & " TTLSalary-" & Val(!TTLSalary))
      If .RecordCount <> 0 Then
         Call ActivityLogBin("", eFrmSalaries, eEdit, TxtSalaryID.Text, DtpEntryDate.Value, "Updated Emp Code-" & TxtEmployeeID.Text & " Salary-" & Val(TxtSalary.Text) & " Bonus-" & Val(TxtBonus.Text) & " TotalSalary-" & Val(TxtTTLSalary.Text))
         ''''''''''''' User Authentication ''''''''''''''
         vUserAction = UserAuthentication("MniSalary", vUser, ObjUserSecurity.IsAdministrator, eUserEdit)
         If vUserAction <> "" Then
            MsgBox vUserAction, vbCritical, "Error"
            Exit Sub
         End If
         ''''''''''''' '''''''''''''''''''' ''''''''''''''
      End If
      If .RecordCount = 0 Then Call ActivityLogBin("", eFrmSalaries, eAdd, TxtSalaryID.Text, DtpEntryDate.Value, "Saved New Emp Code -" & TxtEmployeeID.Text & " Salary-" & Val(TxtSalary.Text) & " Bonus-" & Val(TxtBonus.Text) & " TotalSalary-" & Val(TxtTTLSalary.Text))
      If .RecordCount = 0 Then
         .AddNew
         !SalaryID = TxtSalaryID.Text
         !EmpID = TxtEmployeeID.Text
         !SalaryMonth = DateAdd("d", -1, DateAdd("m", 1, DtpMonth.Value))
      End If
      
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !EntryDate = DtpEntryDate.Value
      !Salary = Val(TxtSalary.Text)
      !SalaryOneDay = Val(TxtSalaryOneDay.Text)
      !SalaryPerHour = Val(TxtSalaryPerHrs.Text)
      !WorkingDays = Val(TxtWorkingDays.Text)
      !WorkingHours = Val(TxtWorkingHours.Text)
      !TTLWorkingDays = Val(TxtTTLWorkingDays.Text)
      !TTLWorkingHours = Val(TxtTTLWorkingHours.Text)
      !Bonus = Val(TxtBonus.Text)
      !TTLSalary = Val(TxtTTLSalary.Text)
      !PreviousLoan = Val(TxtPreviousLoan.Text)
      !LoanInstallment = Val(TxtLoanInstallment.Text)
      !Previous = Val(TxtPrevious.Text)
      !Less = Val(TxtLess.Text)
      !Advance = Val(TxtAdvance.Text)
      !SaleCommision = Val(TxtSaleCommision.Text)
      !ItemsPurchase = Val(TxtItemsPurchase.Text)
      .Update
      .Close
      cn.CommitTrans
      
      If MsgBox("Do you want to print this Salary", vbQuestion + vbYesNo, "Alert") = vbYes Then
         Call BtnPrint_Click
      End If
   End With
   FormStatus = NewMode
   If TxtEmployeeID.Enabled And TxtEmployeeID.Visible Then TxtEmployeeID.SetFocus
   Exit Sub
ErrorHandler:
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   If Trim(TxtEmployeeID.Text) = "" Then
      MsgBox "Please specify a Employee ID", vbExclamation, "Alert"
      If TxtEmployeeID.Enabled And TxtEmployeeID.Visible Then TxtEmployeeID.SetFocus
      Exit Function
   End If
   If Val(TxtTotal.Text) < 0 Then
      If Val(TxtLess.Text) <> 0 Then
         If Abs(Val(TxtLess.Text)) > Abs(Val(TxtTotal.Text)) Then
               MsgBox "Please remove Less.", vbExclamation, "Alert"
               If TxtLess.Enabled And TxtLess.Visible Then TxtLess.SetFocus
               Exit Function
         End If
      End If
   End If
'   If Val(TxtTotal.Text) < 0 Then
'      MsgBox "Negative Salary not Saved.", vbExclamation, "Alert"
'      If TxtLess.Enabled And TxtLess.Visible Then TxtLess.SetFocus
'      Exit Function
'   End If
   If TxtEmployeeID.Enabled = True And DtpMonth.Enabled = True Then
      If cn.Execute("select * from salaries where Empid='" & TxtEmployeeID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'").RecordCount > 0 Then
        MsgBox "Salary of This Month Already Exist. ", vbExclamation, "Alert"
        If TxtEmployeeID.Enabled And TxtEmployeeID.Visible Then TxtEmployeeID.SetFocus
        Exit Function
      End If
   End If
  'All Ok, now validation is success
  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

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
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      BtnEmployee.Enabled = True
      TxtEmployeeID.Enabled = True
      TxtSalaryID.Text = FunGetMaxID()
      DtpMonth.Enabled = True
      DtpMonth.Day = 1
'      SubCalculateWorkingDays
      If TxtEmployeeID.Enabled And TxtEmployeeID.Visible Then TxtEmployeeID.SetFocus
   Case Is = OpenMode
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      BtnEmployee.Enabled = False
      TxtEmployeeID.Enabled = False
      DtpMonth.Enabled = False
     
      'SubCalculateSalary
      TxtSalary.SetFocus
   Case Is = ChangeMode
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnPrint.Enabled = False
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo ErrorHandler
      Set FrmSalary = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then
            ctl.Text = ""
         End If
      End If
   Next
   DtpEntryDate.Value = Date
   DtpMonth.Value = Date
   TxtTotal.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtLess_Change()
   If TxtLess.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtLess.Name Then Exit Sub
   SubCalculateSalary
End Sub

Private Sub TxtLoanInstallment_Change()
   SubCalculateHeaders
End Sub

Private Sub TxtPreviousLoan_Change()
   TxtRemainingLoan.Text = Val(TxtPreviousLoan.Text) - Val(TxtLoanInstallment.Text)
End Sub

Private Sub TxtSalary_Change()
   If ActiveControl.Name <> TxtSalary.Name Then Exit Sub
   If TxtSalary.Visible = False Then Exit Sub
   SubSalaryPerHoursOrPerDay
   If Val(TxtWorkingHours.Text) > 0 Then
      SubCalculateSalaryHrsWise
   Else
      SubCalculateSalary
   End If
End Sub

Private Sub TxtEmployeeID_Change()
   If TxtEmployeeID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtEmployeeID.Name Then Exit Sub
   If TxtEmployeeName.Text <> "" Then TxtEmployeeName.Text = ""
   If BtnSave.Enabled Then Exit Sub
End Sub

Private Sub TxtEmployeeID_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If Me.ActiveControl.Name <> TxtEmployeeID.Name Then Exit Sub
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
        'SchEmployee.ParaInStr = " and SalaryEmployee = 1"
        SchEmployee.Show vbModal, Me
        If SchEmployee.ParaOutEmployeeID = "" Then FunSelectEmployee = False: Exit Function
        TxtEmployeeID.Text = SchEmployee.ParaOutEmployeeID
    End If
    '---------------------------
    If Trim(TxtEmployeeID.Text) = "" Then Exit Function
'    If CN.Execute("select * from Salaries where Employeeid='" & TxtEmployeeID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'").RecordCount > 0 Then
'        MsgBox "Salary Already Exists", vbInformation, "Alert"
'    End If
    sSql = "Select *" & vbCrLf _
            + " from Employees e inner join Designations d on e.designationid = d.designationid" & vbCrLf _
            + " where EmpID=" & Val(TxtEmployeeID.Text)
    With cn.Execute(sSql)
      If .RecordCount > 0 Then
          TxtSalary.Text = !Salary
          SubCalculateWorkingDays
          TxtHoursPerDay.Text = IIf(IsNull(!HoursPerDay), 0, !HoursPerDay)
          TxtLoanInstallment.Text = Val(!LoanInstallment)
          If Val(TxtHoursPerDay.Text) > 0 Then
              GetWorkingHours
              SubCalculateWorkingHours
              SubSalaryPerHoursOrPerDay
              SubCalculateSalaryHrsWise
          Else
              GetWorkingDays
              SubSalaryPerHoursOrPerDay
              SubCalculateSalary
          End If
          TxtEmployeeName.Text = !EmpName
          'TxtFName.Text = !FName
          TxtDesignation.Text = !Designation
          TxtAddress.Text = !Address
          'TxtLess.Text = !minus
         
          SubPrevious
          
          FormStatus = ChangeMode
          'LblAdvance.Caption = CN.Execute("select sum(amount) - sum(rec) from (select sum(amount) as amount, 0 as Rec from OfficeVouchersBody  where groupid ='8' and accountno='" & TxtEmployeeID.Text & "' Union All select 0,sum(RecLoan) from Salaries where Employeeid='" & TxtEmployeeID.Text & "')d").Fields(0).Value
          FunSelectEmployee = True
          .Close
          Exit Function
      Else
         FunSelectEmployee = False
         .Close
         TxtEmployeeID.Text = ""
         TxtEmployeeName.Text = ""
         TxtFName.Text = ""
         TxtDesignation.Text = ""
         TxtAddress.Text = ""
         TxtLoanInstallment.Text = ""
         TxtSalary.Text = ""
         TxtHoursPerDay.Text = 0
         'LblAdvance.Caption = ""
         FormStatus = ChangeMode
         Exit Function
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtTTLSalary_Change()
   On Error GoTo ErrorHandler
'   SubCalculateSalary
'    SubCalculateHeaders
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtWorkingDays_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtWorkingDays.Name Then Exit Sub
   If Val(TxtWorkingDays.Text) = 0 Then Exit Sub
   If Val(TxtWorkingDays.Text) > (Val(TxtTTLWorkingDays.Text) * 3) Then TxtWorkingDays.Text = 0
   TxtWorkingHours.Text = ""
   SubSalaryPerHoursOrPerDay
   SubCalculateSalary
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateHeaders()
   On Error GoTo ErrorHandler
   TxtTotal.Text = Val(TxtTTLSalary.Text) - Val(TxtLoanInstallment.Text) + Val(TxtSaleCommision.Text) - Val(TxtAdvance.Text) - Val(TxtItemsPurchase.Text) - Val(TxtLess.Text)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateSalary()
   On Error GoTo ErrorHandler
   TxtSalaryOneDay.Text = Round(Val(TxtSalary.Text) / IIf(Val(TxtTTLWorkingDays.Text) = 0, 1, Val(TxtTTLWorkingDays.Text)), 2)
   TxtTTLSalary.Text = Round(Val(TxtWorkingDays.Text) * Val(TxtSalaryOneDay.Text)) + Val(TxtBonus.Text)
   SubCalculateHeaders
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateWorkingHours()
   On Error GoTo ErrorHandler
   If Val(TxtHoursPerDay.Text) = 0 Then TxtTTLWorkingHours.Text = 0: Exit Sub
   TxtTTLWorkingHours.Text = Val(TxtTTLWorkingDays.Text) * Val(TxtHoursPerDay.Text)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubSalaryPerHoursOrPerDay()
   On Error GoTo ErrorHandler
   TxtSalaryOneDay.Text = Round(Val(TxtSalary.Text) / Val(TxtTTLWorkingDays.Text), 3)
   If Round(Val(TxtTTLWorkingHours.Text), 2) > 0 Then TxtSalaryPerHrs.Text = Round(Val(TxtSalary.Text) / Val(TxtTTLWorkingHours.Text), 2)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetWorkingDays()
   On Error GoTo ErrorHandler
   SqlWorking = "Select isnull(sum(WorkingDays),0) From " & vbCrLf _
   + " (" & vbCrLf _
   + " Select Count(HolidayID) WorkingDays from Holidays HD " & vbCrLf _
   + " Where datepart(day,Date)  Between datepart(day,'" & DtpMonth.Value & "')  and '" & DateDiff("d", DtpMonth.Value, DateAdd("m", 1, DtpMonth.Value)) & "'" & vbCrLf _
   + " and datepart(Month,Date) =  datepart(Month,'" & DtpMonth.Value & "')" & vbCrLf _
   + " and datepart(Year,Date) =  datepart(Year,'" & DtpMonth.Value & "')" & vbCrLf _
   + " Union All " & vbCrLf _
   + " Select  Count(LeaveID) as WorkingDays from EmpLeaves " & vbCrLf _
   + " Where UnpaidLeave = 0 " & vbCrLf _
   + " And datepart(day,FromDate)  Between datepart(day,'" & DtpMonth.Value & "')  and '" & DateDiff("d", DtpMonth.Value, DateAdd("m", 1, DtpMonth.Value)) & "'" & vbCrLf _
   + " and datepart(Month,FromDate) =  datepart(Month,'" & DtpMonth.Value & "')" & vbCrLf _
   + " and datepart(Year,FromDate) =  datepart(Year,'" & DtpMonth.Value & "')" & vbCrLf _
   + " and EmpID = " & Val(TxtEmployeeID.Text) & vbCrLf _
   + " Union All " & vbCrLf _
   + " Select  Count(AttendID) WorkingDays From EmpAttendance " & vbCrLf _
   + " Where dateout Is Not Null " & vbCrLf _
   + " and datepart(day,AttendDate)  Between datepart(day,'" & DtpMonth.Value & "')  and '" & DateDiff("d", DtpMonth.Value, DateAdd("m", 1, DtpMonth.Value)) & "'" & vbCrLf _
   + " and datepart(Month,AttendDate) =  datepart(Month,'" & DtpMonth.Value & "')" & vbCrLf _
   + " and datepart(Year,AttendDate) =  datepart(Year,'" & DtpMonth.Value & "')" & vbCrLf _
   + " and Empid = " & Val(TxtEmployeeID.Text) & vbCrLf _
   + " ) WorkingDays"
   TxtWorkingDays.Text = cn.Execute(SqlWorking).Fields(0)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetWorkingHours()
   On Error GoTo ErrorHandler
   If Val(TxtHoursPerDay.Text) = 0 Then Exit Sub
   SqlWorking = " Select isnull(sum(WorkingHours),0) From " & vbCrLf _
   + " ( " & vbCrLf _
   + " Select isnull(Hoursperday,0) * " & vbCrLf _
   + " ( " & vbCrLf _
   + " Select Count(HolidayID) from Holidays " & vbCrLf _
   + " Where datepart(day,Date)  Between datepart(day,'" & DtpMonth.Value & "')  and '" & DateDiff("d", DtpMonth.Value, DateAdd("m", 1, DtpMonth.Value)) & "'" & vbCrLf _
   + " and datepart(Month,Date) =  datepart(Month,'" & DtpMonth.Value & "')" & vbCrLf _
   + " and datepart(Year,Date) =  datepart(Year,'" & DtpMonth.Value & "')" & vbCrLf _
   + " )  WorkingHours " & vbCrLf _
   + " from Employees Where EmpID =631 " & vbCrLf _
   + " Union All " & vbCrLf _
   + " Select isnull(sum(Hoursperday),0) WorkingHours from EmpLeaves EL " & vbCrLf _
   + " Inner Join  Employees Emp ON EL.EmpID = Emp.EmpID " & vbCrLf _
   + " Where UnpaidLeave = 0 " & vbCrLf _
   + " And datepart(day,FromDate)  Between datepart(day,'" & DtpMonth.Value & "')  and '" & DateDiff("d", DtpMonth.Value, DateAdd("m", 1, DtpMonth.Value)) & "'" & vbCrLf _
   + " and datepart(Month,FromDate) =  datepart(Month,'" & DtpMonth.Value & "')" & vbCrLf _
   + " and datepart(Year,FromDate) =  datepart(Year,'" & DtpMonth.Value & "')" & vbCrLf _
   + " and EL.EmpID = " & Val(TxtEmployeeID.Text) & vbCrLf _
   + " Union All " & vbCrLf _
   + " Select cast(Sum(WorkingTime/60)as int) WorkingHours From EmpAttendance " & vbCrLf _
   + " Where dateout Is Not Null " & vbCrLf _
   + " and datepart(day,AttendDate)  Between datepart(day,'" & DtpMonth.Value & "')  and '" & DateDiff("d", DtpMonth.Value, DateAdd("m", 1, DtpMonth.Value)) & "'" & vbCrLf _
   + " and datepart(Month,AttendDate) =  datepart(Month,'" & DtpMonth.Value & "')" & vbCrLf _
   + " and datepart(Year,AttendDate) =  datepart(Year,'" & DtpMonth.Value & "') and EmpID = " & Val(TxtEmployeeID.Text) & vbCrLf _
   + " ) WorkingHours"
   TxtWorkingHours.Text = cn.Execute(SqlWorking).Fields(0)
   'If Val(TxtWorkingHours.Text) > 0 Then TxtWorkingDays.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateSalaryHrsWise()
   On Error GoTo ErrorHandler
   TxtTTLSalary.Text = Round(Val(TxtWorkingHours.Text) * Val(TxtSalaryPerHrs.Text))
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtWorkingHours_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtWorkingHours.Name Then Exit Sub
   If Val(TxtWorkingHours.Text) = 0 Or Val(TxtTTLWorkingHours.Text) = 0 Then TxtWorkingHours.Text = "": Exit Sub
   TxtWorkingDays.Text = ""
   If Val(TxtWorkingHours.Text) > Val(TxtTTLWorkingHours.Text) Then TxtWorkingHours.Text = 0
   SubSalaryPerHoursOrPerDay
   SubCalculateSalaryHrsWise
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select isnull(max(SalaryID),0)+1 from Salaries").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SalariesBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmSalaries) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmSalaries & ", " & vUser & "," & TableHeaderFields(eFrmSalaries) & " from Salaries " & vbCrLf _
             & "Where SalaryID = " & TxtSalaryID.Text & " and EntryDate = '" & DtpEntryDate.Value & "'"
      cn.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub




