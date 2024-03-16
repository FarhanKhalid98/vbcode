VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CB3D1B76-3538-4AA3-B3BA-CB69FE1811D0}#1.0#0"; "SIUTextBox.ocx"
Begin VB.Form FrmSalary 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9030
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12030
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSalary.frx":0000
   ScaleHeight     =   9030
   ScaleMode       =   0  'User
   ScaleWidth      =   10777.16
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnStaff 
      Height          =   480
      Left            =   9600
      TabIndex        =   15
      Top             =   1320
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   847
      TX              =   "..."
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
      MICON           =   "FrmSalary.frx":6EBB
      BC              =   12632256
      FC              =   0
   End
   Begin SIUTextBox.UTxt TxtStaffName 
      Height          =   480
      Left            =   5865
      TabIndex        =   16
      Top             =   1320
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   847
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SIUrdu"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin SIUTextBox.Txt TxtStaffID 
      Height          =   480
      Left            =   10065
      TabIndex        =   0
      Top             =   1320
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   847
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mandatory       =   1
   End
   Begin SIUTextBox.UTxt TxtFName 
      Height          =   480
      Left            =   2505
      TabIndex        =   17
      Top             =   1320
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   847
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SIUrdu"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   100
   End
   Begin SIUTextBox.UTxt TxtDesignation 
      Height          =   480
      Left            =   225
      TabIndex        =   21
      Top             =   1320
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   847
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SIUrdu"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   50
   End
   Begin SIUTextBox.UTxt TxtAddress 
      Height          =   480
      Left            =   3525
      TabIndex        =   23
      Top             =   2415
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   847
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SIUrdu"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   200
   End
   Begin SIUTextBox.Txt TxtSalary 
      Height          =   480
      Left            =   9345
      TabIndex        =   3
      Top             =   4440
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   847
      Alignment       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mandatory       =   1
   End
   Begin SIUTextBox.Txt TxtTTLWorkingDays 
      Height          =   480
      Left            =   7260
      TabIndex        =   14
      Top             =   4440
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   847
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mandatory       =   1
   End
   Begin SIUTextBox.Txt TxtSalaryOneDay 
      Height          =   480
      Left            =   2985
      TabIndex        =   5
      Top             =   4440
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   847
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   6
      Mandatory       =   1
   End
   Begin SIUTextBox.Txt TxtTTLSalary 
      Height          =   480
      Left            =   930
      TabIndex        =   6
      Top             =   4440
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   847
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   7
      Mandatory       =   1
   End
   Begin SIUTextBox.Txt TxtWorkingDays 
      Height          =   480
      Left            =   5160
      TabIndex        =   4
      Top             =   4440
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   847
      Alignment       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
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
      Left            =   7380
      TabIndex        =   12
      Top             =   7650
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
      MICON           =   "FrmSalary.frx":6ED7
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6060
      TabIndex        =   8
      Top             =   7650
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
      MICON           =   "FrmSalary.frx":6EF3
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3420
      TabIndex        =   10
      Top             =   7650
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
      MICON           =   "FrmSalary.frx":6F0F
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8700
      TabIndex        =   13
      Top             =   7650
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
      MICON           =   "FrmSalary.frx":6F2B
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4740
      TabIndex        =   9
      Top             =   7650
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
      MICON           =   "FrmSalary.frx":6F47
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpMonth 
      Height          =   345
      Left            =   1950
      TabIndex        =   1
      Top             =   2415
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "MMM-yyyy"
      Format          =   41418755
      CurrentDate     =   38595
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2070
      TabIndex        =   11
      Top             =   7650
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
      MICON           =   "FrmSalary.frx":6F63
      BC              =   14737632
      FC              =   0
   End
   Begin SIUTextBox.Txt TxtAdvance 
      Height          =   480
      Left            =   7500
      TabIndex        =   32
      Top             =   6300
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   847
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
   End
   Begin SIUTextBox.Txt TxtLess 
      Height          =   480
      Left            =   5355
      TabIndex        =   7
      Top             =   6300
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   847
      Alignment       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
   End
   Begin SIUTextBox.Txt TxtTotal 
      Height          =   480
      Left            =   855
      TabIndex        =   33
      Top             =   6345
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   847
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SIUTextBox.Txt TxtPrevious 
      Height          =   480
      Left            =   9570
      TabIndex        =   34
      Top             =   6300
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   847
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
   End
   Begin MSComCtl2.DTPicker DtpEntryDate 
      Height          =   345
      Left            =   585
      TabIndex        =   2
      Top             =   2430
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   41418755
      CurrentDate     =   38595
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "bÐiBP"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1440
      TabIndex        =   39
      Top             =   1935
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ÈµL Bp"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10350
      TabIndex        =   38
      Top             =   5715
      Width           =   555
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "nÆAË•ÐA"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8055
      TabIndex        =   37
      Top             =   5760
      Width           =   750
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ÎPÌ’º"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6075
      TabIndex        =   36
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "BÐBµL"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1800
      TabIndex        =   35
      Top             =   5805
      Width           =   390
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
      Left            =   1965
      TabIndex        =   31
      Top             =   7410
      Width           =   60
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11595
      Top             =   45
      Width           =   375
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " ÇAÌcÅP"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10065
      TabIndex        =   30
      Top             =   3930
      Width           =   615
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "BÏº ¿Bº Ãe ¼º"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5280
      TabIndex        =   29
      Top             =   3930
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇBÂ mA Ãe ¼º"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7455
      TabIndex        =   28
      Top             =   3930
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇAÌcÅP Îº Ãe ¸ÐA"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2790
      TabIndex        =   27
      Top             =   3930
      Width           =   1665
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇAÌcÅP ¼º"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1440
      TabIndex        =   26
      Top             =   3930
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇBÂ"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2745
      TabIndex        =   25
      Top             =   1935
      Width           =   255
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÈO… "
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10710
      TabIndex        =   24
      Top             =   1905
      Width           =   345
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÇfÉª"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2055
      TabIndex        =   22
      Top             =   825
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "¿BÆ"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9330
      TabIndex        =   20
      Top             =   825
      Width           =   255
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "”Ìº"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10560
      TabIndex        =   19
      Top             =   825
      Width           =   360
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÐf¾Ë"
      BeginProperty Font 
         Name            =   "SIUrdu"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5250
      TabIndex        =   18
      Top             =   825
      Width           =   600
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
Dim vID As String
Dim sSQL As String, vStrSQL As String
Dim vCounter As Integer
Dim vCounter1 As Integer
Dim PreviousDate As Date

Private Sub SubPrevious()
   TxtPrevious.Text = ""
   TxtSoap.Text = ""
   TxtAdvance.Text = ""
   With CN.Execute("Select max(EntryDate)as EntryDate from Salaries where staffid = " & TxtStaffID.Text)
      If Not IsNull(!EntryDate) Then
         PreviousDate = !EntryDate
      Else
         PreviousDate = DtpMonth.Value
      End If
   End With
   TxtPrevious.Text = CN.Execute("SELECT isnull(dbo.FunCurrentBalance(" & TxtStaffID.Text & ",'" & PreviousDate & "'),0)").Fields(0).Value
   sSQL = "select Groupid, sum(amount) as Amount from PaymentVouchersBody b inner join PaymentVouchers h on b.voucherno = h.voucherno where voucherdate >='" & PreviousDate & "' and voucherdate<'" & DtpEntryDate.Value & "' and staffid = " & TxtStaffID.Text & " Group By GroupID"
   With CN.Execute(sSQL)
      While Not .EOF
         Select Case !GroupID
         Case 1
            TxtSoap.Text = !Amount
         Case 2
            TxtAdvance.Text = !Amount
         End Select
         .MoveNext
      Wend
   End With
   Call SubCalculateSalary
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
  Dim vTbl As String
  'If RsHeader.RecordCount > 0 Then
    CN.BeginTrans
    CN.Execute "delete from Salaries where staffid='" & TxtStaffID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'"
    'RsHeader.Requery
    CN.CommitTrans
    Call SubClearFields
    FormStatus = NewMode
  'End If
  Exit Sub
ErrorHandler:
  If CN.Errors.Count > 0 Then CN.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchSalary.Show vbModal
   If SchSalary.ParaOutStaffID <> "" Then
      TxtStaffID.Text = SchSalary.ParaOutStaffID
      'Dim a
      'a = Split(SchSalary.ParaOutDate, "/")
      'DtpMonth.Value = Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      DtpMonth.Value = SchSalary.ParaOutDate
      GetSalary
   End If
End Sub

Private Sub GetSalary()
   Dim n As Integer
   On Error GoTo ErrorHandler
   sSQL = "Select s.*, f.name, f.fname, Address, designation from salaries s inner join staff f on s.StaffID = f.StaffID" & _
          " where s.StaffID='" & TxtStaffID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'"
   With CN.Execute(sSQL)
      If Not .BOF Then
         TxtStaffID.Text = !StaffID
         DtpMonth.Value = !SalaryMonth
         DtpEntryDate.Value = !EntryDate
         TxtStaffName.Text = IIf(IsNull(!Name), "", !Name)
         TxtFName.Text = IIf(IsNull(!FName), "", !FName)
         TxtDesignation.Text = IIf(IsNull(!Designation), "", !Designation)
         TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
         TxtSalary.Text = !Salary
         TxtSalaryOneDay.Text = !SalaryOneDay
         TxtWorkingDays.Text = !WorkingDays
         TxtTTLWorkingDays.Text = !TTLWorkingDays
         TxtTTLSalary.Text = !TTLSalary
         TxtPrevious.Text = !Previous
         TxtAdvance.Text = !Advance
         TxtSoap.Text = !Soap
         TxtLess.Text = !Less
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
   vStrSQL = " select h.*,s.*, departmentname" & vbCrLf _
      + " from Salaries h " & vbCrLf _
      + " inner join staff s on h.staffid = s.staffid" & vbCrLf _
      + " inner join departments d on d.departmentid = s.departmentid" & vbCrLf _
      + " where s.StaffID='" & TxtStaffID.Text & "' and SalaryMonth='" & IIf(TxtStaffID.Enabled = True, DateAdd("d", -1, DateAdd("m", 1, DtpMonth.Value)), DtpMonth.Value) & "'"
   
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set RptReportViewer.Report = New CrpSalary
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
   RptReportViewer.Report.PaperOrientation = crPortrait
   If MsgBox("Do you want to print directly this Salary", vbQuestion + vbYesNo, "Alert") = vbYes Then
      RptReportViewer.Report.PrintOut False
   Else
      RptReportViewer.Show vbModal
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateWorkingDays()
   TxtTTLWorkingDays.Text = DateDiff("d", DtpMonth.Value, DateAdd("m", 1, DtpMonth.Value))
   TxtSalaryOneDay.Text = Round(Val(TxtSalary.Text) / Val(TxtTTLWorkingDays.Text), 2)
End Sub

Private Sub BtnStaff_Click()
  If FunSelectStaff(ssButton, False) = True Then
      If DtpMonth.Enabled Then DtpMonth.SetFocus
   Else
      TxtStaffID.SetFocus
   End If
End Sub

Private Sub DtpEntryDate_Change()
   If Me.ActiveControl.Name <> DtpEntryDate.Name Then Exit Sub
   SubPrevious
End Sub

Private Sub DtpMonth_Change()
   If Me.ActiveControl.Name <> DtpMonth.Name Then Exit Sub
   SubPrevious
   SubCalculateWorkingDays
   SubCalculateSalary
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hwnd, "Salary"
   DtpMonth.Value = Date
   DtpEntryDate.Value = Date
   'DtpMonth.Value = DateAdd("d", -1, DateAdd("m", 1, DtpMonth.Value))
   FormStatus = NewMode
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
         Case TxtStaffID.Name: If FunSelectStaff(ssFunctionKey, True) = True Then If DtpMonth.Enabled Then DtpMonth.SetFocus
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
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If FunValidation = False Then Exit Sub
   CN.BeginTrans
   Set RsHeader = New ADODB.Recordset
   
   RsHeader.Open "Select * FROM Salaries where SalaryMonth='" & IIf(TxtStaffID.Enabled = True, DateAdd("d", -1, DateAdd("m", 1, DtpMonth.Value)), DtpMonth.Value) & "' and Staffid='" & TxtStaffID.Text & "'", CN, adOpenStatic, adLockPessimistic
   With RsHeader
      If RsHeader.RecordCount = 0 Then
         .AddNew
         !StaffID = TxtStaffID.Text
         !SalaryMonth = DateAdd("d", -1, DateAdd("m", 1, DtpMonth.Value))
         !EntryDate = DtpEntryDate.Value
      End If
      !Salary = Val(TxtSalary.Text)
      !SalaryOneDay = Val(TxtSalaryOneDay.Text)
      !WorkingDays = Val(TxtWorkingDays.Text)
      !TTLWorkingDays = Val(TxtTTLWorkingDays.Text)
      !TTLSalary = Val(TxtTTLSalary.Text)
      !Previous = Val(TxtPrevious.Text)
      !Less = Val(TxtLess.Text)
      !Advance = Val(TxtAdvance.Text)
      !Soap = Val(TxtSoap.Text)
      .Update
      .Close
      CN.CommitTrans
      If MsgBox("Do you want to print this Salary", vbQuestion + vbYesNo, "Alert") = vbYes Then
         Call BtnPrint_Click
      End If
   End With
   FormStatus = NewMode
   If TxtStaffID.Enabled And TxtStaffID.Visible Then TxtStaffID.SetFocus
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   If Trim(TxtStaffID.Text) = "" Then
      MsgBox "Please specify a Staff ID", vbExclamation, "Alert"
      If TxtStaffID.Enabled And TxtStaffID.Visible Then TxtStaffID.SetFocus
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
   If Val(TxtTotal.Text) < 0 Then
      MsgBox "Negative Salary not Saved.", vbExclamation, "Alert"
      If TxtLess.Enabled And TxtLess.Visible Then TxtLess.SetFocus
      Exit Function
   End If
   If TxtStaffID.Enabled = True And DtpMonth.Enabled = True Then
      If CN.Execute("select * from salaries where staffid='" & TxtStaffID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'").RecordCount > 0 Then
        MsgBox "Salary of This Month Already Exist. ", vbExclamation, "Alert"
        If TxtStaffID.Enabled And TxtStaffID.Visible Then TxtStaffID.SetFocus
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
      BtnStaff.Enabled = True
      TxtStaffID.Enabled = True
      DtpMonth.Enabled = True
      DtpMonth.Day = 1
      DtpEntryDate.Enabled = True
      SubCalculateWorkingDays
      If TxtStaffID.Enabled And TxtStaffID.Visible Then TxtStaffID.SetFocus
   Case Is = OpenMode
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      BtnStaff.Enabled = False
      TxtStaffID.Enabled = False
      DtpMonth.Enabled = False
      DtpEntryDate.Enabled = False
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
      If TypeOf ctl Is SIUTextBox.txt Then
         If ctl.Tag = "" Then
            ctl.Text = ""
         End If
      ElseIf TypeOf ctl Is SIUTextBox.UTxt Then
         If ctl.Tag = "" Then
            ctl.Text = ""
         End If
      End If
   Next
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

Private Sub TxtSalary_Change()
   If TxtSalary.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSalary.Name Then Exit Sub
   TxtSalaryOneDay.Text = Round(Val(TxtSalary.Text) / Val(TxtTTLWorkingDays.Text), 2)
   SubCalculateSalary
End Sub

Private Sub TxtSalaryOneDay_Change()
   If TxtSalaryOneDay.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSalaryOneDay.Name Then Exit Sub
   SubCalculateSalary
End Sub

Private Sub TxtStaffID_change()
   If TxtStaffID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStaffID.Name Then Exit Sub
   If TxtStaffName.Text <> "" Then
      TxtStaffName.Text = ""
   End If
   If BtnSave.Enabled Then Exit Sub
End Sub

Private Sub TxtStaffID_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If Me.ActiveControl.Name <> TxtStaffID.Name Then Exit Sub
    Dim vTemp As Boolean
    vTemp = Not FunSelectStaff(ssValidate, True)
    If vTemp = True Then
        vTemp = Not FunSelectStaff(ssButton, False)
    End If
    Cancel = vTemp
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Function FunSelectStaff(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStaff.ParaInStr = " and SalaryStaff = 1"
        SchStaff.Show vbModal, Me
        If SchStaff.ParaOutStaffID = "" Then FunSelectStaff = False: Exit Function
        TxtStaffID.Text = SchStaff.ParaOutStaffID
    End If
    '---------------------------
    If Trim(TxtStaffID.Text) = "" Then Exit Function
'    If CN.Execute("select * from Salaries where staffid='" & TxtStaffID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'").RecordCount > 0 Then
'        MsgBox "Salary Already Exists", vbInformation, "Alert"
'    End If
    sSQL = "Select *" & vbCrLf _
            + " from Staff" & vbCrLf _
            + " where StaffID=" & Val(TxtStaffID.Text) & " and SalaryStaff = 1"
    With CN.Execute(sSQL)
      If .RecordCount > 0 Then
        TxtStaffName.Text = !Name
        TxtFName.Text = !FName
        TxtDesignation.Text = !Designation
        TxtAddress.Text = !Address
        TxtSalary.Text = !Salary
        TxtLess.Text = !minus
        SubPrevious
        SubCalculateWorkingDays
        'LblAdvance.Caption = CN.Execute("select sum(amount) - sum(rec) from (select sum(amount) as amount, 0 as Rec from OfficeVouchersBody  where groupid ='8' and accountno='" & TxtStaffID.Text & "' Union All select 0,sum(RecLoan) from Salaries where staffid='" & TxtStaffID.Text & "')d").Fields(0).Value
        FunSelectStaff = True
        .Close
        Exit Function
      Else
        FunSelectStaff = False
        .Close
        TxtStaffID.Text = ""
        TxtStaffName.Text = ""
        TxtFName.Text = ""
        TxtDesignation.Text = ""
        TxtAddress.Text = ""
        TxtSalary.Text = ""
        TxtLess.Text = ""
        'LblAdvance.Caption = ""
        Exit Function
      End If
    End With
Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub TxtTTLWorkingDays_Change()
   If Val(TxtTTLWorkingDays.Text) = 0 Then Exit Sub
   TxtSalaryOneDay.Text = Round(Val(TxtSalary.Text) / Val(TxtTTLWorkingDays.Text), 2)
End Sub

Private Sub TxtWorkingDays_Change()
   If ActiveControl.Name <> TxtWorkingDays.Name Then Exit Sub
   SubCalculateSalary
End Sub

Private Sub SubCalculateSalary()
   TxtTTLSalary.Text = Round(Val(TxtWorkingDays.Text) * Val(TxtSalaryOneDay.Text))
   TxtTotal.Text = Val(TxtTTLSalary.Text) - Val(TxtLess.Text) - Val(TxtSoap.Text) - IIf(Val(TxtPrevious.Text) + Val(TxtAdvance.Text) < 0, Val(TxtPrevious.Text) + Val(TxtAdvance.Text), 0)
End Sub
