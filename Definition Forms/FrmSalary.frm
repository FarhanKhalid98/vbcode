VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleMode       =   0  'User
   ScaleWidth      =   10777.16
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
      Left            =   11070
      TabIndex        =   40
      Top             =   1260
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
         TabIndex        =   41
         Tag             =   "NC"
         Text            =   "FrmSalary.frx":0000
         Top             =   360
         Width           =   3930
      End
      Begin VB.Label Label17 
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
         TabIndex        =   42
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7380
      TabIndex        =   9
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
      MICON           =   "FrmSalary.frx":008B
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6060
      TabIndex        =   5
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
      MICON           =   "FrmSalary.frx":00A7
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3420
      TabIndex        =   7
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
      MICON           =   "FrmSalary.frx":00C3
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8700
      TabIndex        =   10
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
      MICON           =   "FrmSalary.frx":00DF
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4740
      TabIndex        =   6
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
      MICON           =   "FrmSalary.frx":00FB
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpMonth 
      Height          =   345
      Left            =   2085
      TabIndex        =   2
      Top             =   3945
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "MMM-yyyy"
      Format          =   45875203
      CurrentDate     =   38595
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2070
      TabIndex        =   8
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
      MICON           =   "FrmSalary.frx":0117
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpEntryDate 
      Height          =   345
      Left            =   675
      TabIndex        =   1
      Top             =   3960
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   45875203
      CurrentDate     =   38595
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   690
      TabIndex        =   0
      Top             =   1695
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
      Left            =   1980
      TabIndex        =   23
      Top             =   1695
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
      Left            =   675
      TabIndex        =   24
      Top             =   2775
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
      Left            =   4860
      TabIndex        =   25
      Top             =   1695
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
      Left            =   1620
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1695
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
      MICON           =   "FrmSalary.frx":0133
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtDesignation 
      Height          =   315
      Left            =   7740
      TabIndex        =   31
      Top             =   1710
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
      Left            =   705
      TabIndex        =   3
      Top             =   5310
      Width           =   1380
      _ExtentX        =   2434
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
   Begin SITextBox.Txt TxtTTLWorkingDays 
      Height          =   315
      Left            =   2895
      TabIndex        =   33
      Top             =   5310
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
   Begin SITextBox.Txt TxtWorkingDays 
      Height          =   315
      Left            =   5490
      TabIndex        =   4
      Top             =   5310
      Width           =   1380
      _ExtentX        =   2434
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
   Begin SITextBox.Txt TxtSalaryOneDay 
      Height          =   315
      Left            =   7755
      TabIndex        =   34
      Top             =   5310
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
   Begin SITextBox.Txt TxtTTLSalary 
      Height          =   315
      Left            =   9840
      TabIndex        =   35
      Top             =   5310
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
   Begin SITextBox.Txt TxtPrevious 
      Height          =   315
      Left            =   495
      TabIndex        =   36
      Top             =   1035
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtAdvance 
      Height          =   315
      Left            =   2700
      TabIndex        =   37
      Top             =   1035
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtLess 
      Height          =   315
      Left            =   5310
      TabIndex        =   38
      Top             =   1080
      Visible         =   0   'False
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
      Left            =   7560
      TabIndex        =   39
      Top             =   1080
      Visible         =   0   'False
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
      Left            =   1935
      TabIndex        =   44
      Top             =   135
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
      TabIndex        =   43
      Top             =   765
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      Height          =   195
      Left            =   7740
      TabIndex        =   32
      Top             =   1530
      Width           =   840
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   675
      TabIndex        =   30
      Top             =   1485
      Width           =   525
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   1980
      TabIndex        =   29
      Top             =   1485
      Width           =   780
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   675
      TabIndex        =   28
      Top             =   2565
      Width           =   570
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      Height          =   195
      Left            =   4860
      TabIndex        =   27
      Top             =   1515
      Width           =   915
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
      Left            =   2070
      TabIndex        =   22
      Top             =   3645
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
      Left            =   675
      TabIndex        =   21
      Top             =   3645
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
      Left            =   7575
      TabIndex        =   20
      Top             =   810
      Visible         =   0   'False
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
      Left            =   5310
      TabIndex        =   19
      Top             =   810
      Visible         =   0   'False
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
      Left            =   2715
      TabIndex        =   18
      Top             =   810
      Visible         =   0   'False
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
      Left            =   525
      TabIndex        =   17
      Top             =   810
      Visible         =   0   'False
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
      Left            =   705
      TabIndex        =   16
      Top             =   4995
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
      Left            =   9840
      TabIndex        =   15
      Top             =   4995
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
      Left            =   7755
      TabIndex        =   14
      Top             =   4995
      Width           =   1305
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
      Left            =   5490
      TabIndex        =   13
      Top             =   4995
      Width           =   1425
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
      Left            =   2895
      TabIndex        =   12
      Top             =   4995
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
      Left            =   1965
      TabIndex        =   11
      Top             =   7410
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
Dim sSql As String, vStrSQL As String
Dim vCounter As Integer
Dim vCounter1 As Integer
Dim PreviousDate As Date

Private Sub SubPrevious()
   TxtPrevious.Text = ""
   TxtAdvance.Text = ""
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
  Dim vtbl As String
  'If RsHeader.RecordCount > 0 Then
    CN.BeginTrans
    CN.Execute "delete from Salaries where Employeeid='" & TxtEmployeeID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'"
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
   If SchSalary.ParaOutEmpID <> "" Then
      TxtEmployeeID.Text = SchSalary.ParaOutEmpID
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
   sSql = "Select s.*, f.name, f.fname, Address, designation from salaries s inner join Employee f on s.EmployeeID = f.EmployeeID" & _
          " where s.EmployeeID='" & TxtEmployeeID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'"
   With CN.Execute(sSql)
      If Not .BOF Then
         TxtEmployeeID.Text = !EmployeeID
         DtpMonth.Value = !SalaryMonth
         DtpEntryDate.Value = !EntryDate
         TxtEmployeeName.Text = IIf(IsNull(!Name), "", !Name)
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
      + " inner join Employee s on h.Employeeid = s.Employeeid" & vbCrLf _
      + " inner join departments d on d.departmentid = s.departmentid" & vbCrLf _
      + " where s.EmployeeID='" & TxtEmployeeID.Text & "' and SalaryMonth='" & IIf(TxtEmployeeID.Enabled = True, DateAdd("d", -1, DateAdd("m", 1, DtpMonth.Value)), DtpMonth.Value) & "'"
   
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   'Set RptReportViewer.Report = New CrpSalary
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

Private Sub BtnEmployee_Click()
  If FunSelectEmployee(ssButton, False) = True Then
      If DtpMonth.Enabled Then DtpMonth.SetFocus
   Else
      TxtEmployeeID.SetFocus
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

Private Sub LblClose_Click()
   FraHelp.Visible = False
End Sub

Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblHelp.ForeColor = &H800000
   FraHelp.ZOrder 0
   FraHelp.Visible = True
End Sub

Private Sub LblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If LblHelp.FontUnderline = True Then Exit Sub
   LblHelp.FontUnderline = True
End Sub

Private Sub LblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblHelp.ForeColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Salary"
   HelpLocation Me
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
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, True) = True Then If DtpMonth.Enabled Then DtpMonth.SetFocus
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
   If FunValidation = False Then Exit Sub
   CN.BeginTrans
   Set RsHeader = New ADODB.Recordset
   
   RsHeader.Open "Select * FROM Salaries where SalaryMonth='" & IIf(TxtEmployeeID.Enabled = True, DateAdd("d", -1, DateAdd("m", 1, DtpMonth.Value)), DtpMonth.Value) & "' and Empid='" & TxtEmployeeID.Text & "'", CN, adOpenStatic, adLockPessimistic
   With RsHeader
      If .RecordCount = 0 Then
         .AddNew
         !EmpID = TxtEmployeeID.Text
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
      .Update
      .Close
      CN.CommitTrans
      If MsgBox("Do you want to print this Salary", vbQuestion + vbYesNo, "Alert") = vbYes Then
         Call BtnPrint_Click
      End If
   End With
   FormStatus = NewMode
   If TxtEmployeeID.Enabled And TxtEmployeeID.Visible Then TxtEmployeeID.SetFocus
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
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
   If Val(TxtTotal.Text) < 0 Then
      MsgBox "Negative Salary not Saved.", vbExclamation, "Alert"
      If TxtLess.Enabled And TxtLess.Visible Then TxtLess.SetFocus
      Exit Function
   End If
   If TxtEmployeeID.Enabled = True And DtpMonth.Enabled = True Then
      If CN.Execute("select * from salaries where Empid='" & TxtEmployeeID.Text & "' and SalaryMonth='" & DtpMonth.Value & "'").RecordCount > 0 Then
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
      DtpMonth.Enabled = True
      DtpMonth.Day = 1
      DtpEntryDate.Enabled = True
      SubCalculateWorkingDays
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
      If TypeOf ctl Is SITextBox.txt Then
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
    With CN.Execute(sSql)
      If .RecordCount > 0 Then
        TxtEmployeeName.Text = !EmpName
        'TxtFName.Text = !FName
        TxtDesignation.Text = !Designation
        TxtAddress.Text = !Address
        TxtSalary.Text = !Salary
        'TxtLess.Text = !minus
        SubPrevious
        SubCalculateWorkingDays
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
   TxtTotal.Text = Val(TxtTTLSalary.Text) - Val(TxtLess.Text) - IIf(Val(TxtPrevious.Text) + Val(TxtAdvance.Text) < 0, Val(TxtPrevious.Text) + Val(TxtAdvance.Text), 0)
End Sub
