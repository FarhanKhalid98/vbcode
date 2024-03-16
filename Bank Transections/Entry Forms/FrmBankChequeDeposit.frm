VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmBankChequeDeposit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   DrawMode        =   1  'Blackness
   Icon            =   "FrmBankChequeDeposit.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTotalAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      CausesValidation=   0   'False
      Height          =   330
      Left            =   10283
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   8115
      Width           =   1695
   End
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   6938
      TabIndex        =   6
      Top             =   2730
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   35
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
   Begin SITextBox.Txt TxtReceiveBy 
      Height          =   315
      Left            =   7478
      TabIndex        =   9
      Top             =   3570
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   35
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
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   10493
      TabIndex        =   10
      Top             =   3570
      Width           =   1380
      _ExtentX        =   2434
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
      Masked          =   1
   End
   Begin JeweledBut.JeweledButton btnReceiving 
      Height          =   315
      Left            =   5708
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3570
      Width           =   375
      _ExtentX        =   661
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
      MICON           =   "FrmBankChequeDeposit.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnClose 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10403
      TabIndex        =   17
      Top             =   8925
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
      MICON           =   "FrmBankChequeDeposit.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7733
      TabIndex        =   12
      Top             =   8925
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
      MICON           =   "FrmBankChequeDeposit.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6398
      TabIndex        =   13
      Top             =   8925
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
      MICON           =   "FrmBankChequeDeposit.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5063
      TabIndex        =   14
      Top             =   8925
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
      MICON           =   "FrmBankChequeDeposit.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSlipNo 
      Height          =   315
      Left            =   3878
      TabIndex        =   4
      Top             =   2730
      Width           =   1755
      _ExtentX        =   3096
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtChequeNo 
      Height          =   315
      Left            =   3878
      TabIndex        =   7
      Top             =   3570
      Width           =   1830
      _ExtentX        =   3228
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
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton btndelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9068
      TabIndex        =   16
      Top             =   8925
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
      MICON           =   "FrmBankChequeDeposit.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnBank 
      Height          =   315
      Left            =   6375
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2010
      Width           =   375
      _ExtentX        =   661
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
      MICON           =   "FrmBankChequeDeposit.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBankActName 
      Height          =   315
      Left            =   6750
      TabIndex        =   19
      Top             =   2010
      Width           =   2055
      _ExtentX        =   3625
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
   End
   Begin SITextBox.Txt TxtBankActID 
      Height          =   315
      Left            =   5355
      TabIndex        =   2
      Top             =   2010
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
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
      Masked          =   1
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtVoucherID 
      Height          =   315
      Left            =   3105
      TabIndex        =   0
      Top             =   2010
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      Locked          =   -1  'True
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
      Masked          =   1
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton btnPrint 
      Height          =   420
      Left            =   3728
      TabIndex        =   15
      Top             =   8925
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
      MICON           =   "FrmBankChequeDeposit.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      CausesValidation=   0   'False
      Height          =   4215
      Left            =   3878
      TabIndex        =   11
      Top             =   3885
      Width           =   8295
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   4
      stylesets.count =   1
      stylesets(0).Name=   "style"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
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
      stylesets(0).Picture=   "FrmBankChequeDeposit.frx":0FAA
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
      ActiveRowStyleSet=   "style"
      Columns.Count   =   4
      Columns(0).Width=   3889
      Columns(0).Caption=   "Cheque No"
      Columns(0).Name =   "ChequeNo"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2461
      Columns(1).Caption=   "Cheque Date"
      Columns(1).Name =   "ChequeDate"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd/MM/yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   5292
      Columns(2).Caption=   "Deposit By"
      Columns(2).Name =   "ReceiveBy"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2461
      Columns(3).Caption=   "Amount"
      Columns(3).Name =   "Amount"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   14631
      _ExtentY        =   7435
      _StockProps     =   79
      BackColor       =   16777215
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpVoucherDate 
      Height          =   315
      Left            =   4050
      TabIndex        =   1
      Top             =   2010
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
   Begin SSCalendarWidgets_A.SSDateCombo dtpSlipDate 
      Height          =   315
      Left            =   5633
      TabIndex        =   5
      Top             =   2730
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
   Begin SSCalendarWidgets_A.SSDateCombo dtpChequeDate 
      Height          =   315
      Left            =   6083
      TabIndex        =   8
      Top             =   3570
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
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
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   8805
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   2010
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
      Left            =   10110
      TabIndex        =   33
      Tag             =   "NC"
      Top             =   2010
      Width           =   2205
      _ExtentX        =   3889
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
      Left            =   9750
      TabIndex        =   34
      TabStop         =   0   'False
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
      MICON           =   "FrmBankChequeDeposit.frx":0FC6
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      Height          =   225
      Index           =   1
      Left            =   9188
      TabIndex        =   38
      Top             =   8160
      Width           =   1020
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   10110
      TabIndex        =   36
      Top             =   1785
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   8805
      TabIndex        =   35
      Top             =   1785
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Slip Date"
      Height          =   255
      Index           =   0
      Left            =   5633
      TabIndex        =   32
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Left            =   6938
      TabIndex        =   31
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit By"
      Height          =   255
      Left            =   7478
      TabIndex        =   30
      Top             =   3330
      Width           =   855
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C Name"
      Height          =   195
      Left            =   6750
      TabIndex        =   29
      Top             =   1785
      Width           =   1170
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C  ID"
      Height          =   195
      Left            =   5355
      TabIndex        =   28
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Cheque Deposit"
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
      TabIndex        =   27
      Top             =   270
      Width           =   3810
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Date"
      Height          =   195
      Left            =   6083
      TabIndex        =   26
      Top             =   3330
      Width           =   945
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Slip No"
      Height          =   255
      Left            =   3878
      TabIndex        =   25
      Top             =   2520
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   10493
      TabIndex        =   24
      Top             =   3330
      Width           =   540
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   11610
      Top             =   60
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Date"
      Height          =   195
      Left            =   4050
      TabIndex        =   23
      Top             =   1785
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher ID"
      Height          =   195
      Left            =   3105
      TabIndex        =   22
      Top             =   1785
      Width           =   810
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No"
      Height          =   195
      Left            =   3878
      TabIndex        =   21
      Top             =   3330
      Width           =   810
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmBankChequeDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSql As String
Dim vStrSQL As String
Dim vIsNewRecord As Boolean
Dim vIsNewRow As Boolean
Dim vCounter As Integer
Dim Flag As Boolean
Dim vMode As FormMode
Dim RsReport As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset

Private Sub btnClear_Click()
   FormStatus = NewMode
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub btndelete_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   cn.BeginTrans
   Grid.Redraw = False
   Grid.RemoveAll
   cn.Execute "Delete from BankChequeDepositBody where VoucherID = " & Val(TxtVoucherID.Text)
   cn.Execute "Delete from BankChequeDepositHeader where VoucherID = " & Val(TxtVoucherID.Text)
   cn.CommitTrans
   Grid.Redraw = True
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   vStrSQL = " Select H.VoucherID, h.VoucherDate, Description, BankID, AccountName as BankName," & vbCrLf _
      + " SlipNo, SlipDate, ChequeNo, ChequeDate, DepositBy, DepositAmount" & vbCrLf _
      + " from BankChequeDepositHeader H " & vbCrLf _
      + " inner join BankChequeDepositBody B on H.VoucherID = B.VoucherID " & vbCrLf _
      + " inner Join ChartofAccounts a on a.AccountNo = h.BankID" & vbCrLf _
      + " where H.VoucherID = " & Val(TxtVoucherID.Text)


   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
   Set RptReportViewer.Report = New CRptBankChqDeposit
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1

   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port

'   RptReportViewer.Report.PaperOrientation = crPortrait
   RptReportViewer.Show
   'RptReportViewer.Report.PrintOut False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If Not (UCase(ActiveControl.Name) Like UCase("txt*")) Then Exit Sub
   If btnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If btnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
         Set frmObj = Nothing
      Next
         Set RsBody = Nothing
         Set FrmBankChequeDeposit = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Bank Cheque Deposite"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   'btnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   'btndelete.Visible = Not ObjRegistry.ReadOnlyStatus
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Property Get FormStatus() As FormMode
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
  On Error GoTo ErrorHandler
   vMode = vNewValue
   Select Case vNewValue
   Case Is = NewMode
      Call SubClearFields
      If RsBody.State = adStateOpen Then RsBody.Close
      btnOpen.Enabled = True
      btndelete.Enabled = False
      btnSave.Enabled = False
      btnClear.Enabled = True
      btnPrint.Enabled = False
      TxtVoucherID.Text = FunGetMaxID
      'DtpPurchaseDate.Value = Date
      PopulateDataToGrid
      TxtBankActID.Enabled = True
      'If DtpPurchaseDate.Enabled And DtpPurchaseDate.Visible Then DtpPurchaseDate.SetFocus
      vIsNewRecord = True
      vIsNewRow = True
   Case Is = OpenMode
      btnOpen.Enabled = True
      btndelete.Enabled = True
      btnClear.Enabled = True
      btnSave.Enabled = False
      btnPrint.Enabled = True
      vIsNewRecord = False
      vIsNewRow = True
   Case Is = ChangeMode
      btnOpen.Enabled = False
      btndelete.Enabled = False
      btnSave.Enabled = True
      btnPrint.Enabled = False
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         If ctl.Tag <> "NC" Then
            ctl.Text = ""
         End If
      ElseIf TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag <> "NC" Then
            ctl.Text = ""
         End If
      End If
   Next
   DtpVoucherDate.DateValue = Date
   dtpSlipDate.DateValue = Date
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ChequeNo").Text = " "
   Grid.Update
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select isnull(max(VoucherID),0) from BankChequeDepositHeader").Fields(0) + 1
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "select * from BankChequeDepositBody where VoucherID = ' " & Val(TxtVoucherID.Text) & " ' ", cn, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = "Select B.ChequeNo, B.ChequeDate, B.DepositBy, B.DepositAmount  From BankChequeDepositBody B  where B.VoucherID =" & Val(TxtVoucherID.Text)
      With cn.Execute(sSql)
         If .RecordCount > 0 Then
            Grid.Redraw = False
            Grid.MoveFirst
            Grid.RemoveAll
            Grid.AllowAddNew = True
            TxtTotalAmount.Text = 0
            While Not .EOF
               Grid.AddNew
               Grid.Columns("ChequeNo").Text = Val(!ChequeNo)
               Grid.Columns("ChequeDate").Text = (!ChequeDate)
               Grid.Columns("ReceiveBy").Text = IIf(IsNull(!DepositBy), "", !DepositBy)
               Grid.Columns("Amount").Value = Val(!DepositAmount)
               TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + !DepositAmount
               .MoveNext
            Wend
         End If
         .Close
      End With
      Grid.AddNew
      Grid.Columns("ChequeNo").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
      If KeyCode = vbKeyReturn Then
         If ActiveControl.Name = Grid.Name Then
            Grid_DblClick
         Else
            keybd_event 9, 1, 1, 1
            KeyCode = 0
         End If
      ElseIf KeyCode = vbKeyEscape Then
         Call SubClearDetailArea: TxtchequeNo.SetFocus
      ElseIf KeyCode = vbKeyF1 Then
         Select Case ActiveControl.Name
            Case TxtBankActID.Name: If FunSelectAccount(ssFunctionKey, False) = True Then If TxtOrganizationID.Enabled And TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus Else TxtBankActID.SetFocus
            Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtSlipNo.SetFocus Else TxtOrganizationID.SetFocus
            Case TxtchequeNo.Name: If FunSelectReceiving(ssFunctionKey, False) = True Then TxtReceiveBy.SetFocus Else TxtchequeNo.SetFocus
         End Select
      ElseIf Shift = vbCtrlMask Then
         Select Case KeyCode
            Case vbKeyS
               If btnSave.Enabled And btnSave.Visible Then btnSave_Click
               KeyCode = 0
            Case vbKeyW
               If btnClear.Enabled = True Then btnClear_Click
               KeyCode = 0
            Case vbKeyQ
               If BtnClose.Enabled = True Then BtnClose_Click
               KeyCode = 0
            Case vbKeyO
               If btnOpen.Enabled = True Then BtnOpen_Click
               KeyCode = 0
            Case vbKeyP
               If btnPrint.Enabled = True Then BtnPrint_Click
               KeyCode = 0
            Case vbKeyR
               If btndelete.Enabled And btndelete.Visible Then btndelete_Click
               KeyCode = 0
            Case vbKeyDelete
               MniRemoveRow_Click
               KeyCode = 0
         End Select
      ElseIf ActiveControl.Name = TxtchequeNo.Name Then
         If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtchequeNo.Name Then
         KeyCode = 0
      End If
   End If
   Exit Sub
ErrorHandler:
     Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   Call Grid_LostFocus
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("ChequeNo").Text) = "" Then
      TxtchequeNo.Text = ""
      TxtchequeNo.Enabled = True
      vIsNewRow = True
      If TxtchequeNo.Enabled Then TxtchequeNo.SetFocus
   Else
      TxtchequeNo.Enabled = False
      vIsNewRow = False
   End If
End Sub

Private Sub Grid_GotFocus()
   Flag = True
   TxtchequeNo.Enabled = False
End Sub

Private Sub SubClearDetailArea()
   dtpSlipDate.DateValue = Date
   TxtchequeNo.Text = ""
   TxtAmount.Text = ""
   TxtReceiveBy.Text = ""
End Sub

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchAccounts.ParaInDetail = ""
      SchAccounts.ParaInWhereClause = " and c.accountno like '1%' or c.accountno = '111'"
      'SchAccounts.cmbfilter.Text = "Banks"
      'SchAccounts.cmbfilter.Enabled = False
      SchAccounts.Show vbModal, Me
      If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
      TxtBankActID.Text = SchAccounts.ParaOutAccountNo
   End If
   Dim vStrSQL As String
   vStrSQL = "select * from ChartofAccounts where AccountNo ='" & Val(TxtBankActID.Text) & "' and (accountno like '1%' or accountno = '111')"
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtBankActName.Text = !AccountName
         .Close
         FunSelectAccount = True
         If btnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      Else
         FunSelectAccount = False
         .Close
         TxtBankActName.Text = ""
         If btnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectReceiving(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchReceivingCheque.Show vbModal, Me
      If SchReceivingCheque.ParaOutID = "" Then FunSelectReceiving = False: Exit Function
      TxtchequeNo.Text = SchReceivingCheque.ParaOutID
   End If
   Dim vStrSQL As String
   vStrSQL = "select *, AccountName as ReceivingName from BankChequeReceiveBody b inner Join ChartofAccounts c on c.AccountNo = b.ReceivingID Where ChequeNo = '" & Val(TxtchequeNo.Text) & "'"
   With cn.Execute(vStrSQL)
         If .RecordCount > 0 Then
           'TxtAct.Text = !PartyName
'            TxtVendorAddress.Text = IIf(IsNull(!address), "", !address)
'            TxtVendorCity.Text = IIf(IsNull(!city), "", !city)
            dtpChequedate.DateValue = !ChequeDate
            TxtReceiveBy.Text = IIf(IsNull(!ReceivingName), "", !ReceivingName)
            TxtAmount.Text = Val(!Amount)
            .Close
            FunSelectReceiving = True
            If btnSave.Enabled = False Then FormStatus = ChangeMode
            Exit Function
         Else
            FunSelectReceiving = False
            .Close
            'TxtActPayeeName.Text = ""
'            TxtChequeNo.Text = ""
'            TxtVendorAddress.Text = ""
'            TxtVendorCity.Text = ""
            If btnSave.Enabled = False Then FormStatus = ChangeMode
         End If
      End With
      Exit Function
ErrorHandler:
      Call ShowErrorMessage
End Function

'Private Sub TxtChequeNo_Change()
'If ActiveControl.Name <> TxtChequeNo.Name Then Exit Sub
'   If TxtActPayeeName.Text <> "" Then
'      TxtActPayeeName.Text = ""
'      TxtChequeNo.Text = ""
'   End If
'End Sub
'
'Private Sub TxtChequeNo_Validate(Cancel As Boolean)
'On Error GoTo ErrorHandler
'   If Me.ActiveControl.Name <> TxtChequeNo.Name Then Exit Sub
'   If TxtActPayeeName.Text <> "" Then Exit Sub
'   Dim vTemp As Boolean
'   vTemp = Not FunSelectReceiving(ssValidate, True)
'   If vTemp = True Then
'      vTemp = Not FunSelectReceiving(ssButton, False)
'   End If
'   Cancel = vTemp
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub

Private Sub TxtAmount_Change()
'If btnSave.Enabled = False Then FormStatus = changeMode
End Sub

Private Sub TxtBankActID_Change()
If ActiveControl.Name <> TxtBankActID.Name Then Exit Sub
If TxtBankActName.Text <> "" Then
      'TxtBankActID.Text = ""
      TxtBankActName.Text = ""
    End If
End Sub

Private Sub TxtBankActID_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtBankActID.Name Then Exit Sub
   If TxtBankActName.Text <> "" Then Exit Sub
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

Private Sub TxtAmount_LostFocus()
'   If Trim(Grid.Columns("ChequeNo").Text) = "" Then
'      vIsNewRow = True
'      Flag = True
'   Else
'      vIsNewRow = False
'   End If
Call GetDataFromTextBoxesToGrid
End Sub

Private Sub GetDataFromTextBoxesToGrid()
   On Error GoTo ErrorHandler
   If Trim(TxtchequeNo.Text) = "" Then
      MsgBox " Please Specify ChequeNo ", vbInformation + vbOKOnly, "Error"
      If TxtchequeNo.Enabled = True Then TxtchequeNo.SetFocus
      Exit Sub
   End If
   If TxtAmount.Text = "" Then
      MsgBox " Please Specify Amount ", vbInformation + vbOKOnly, "Error"
      If TxtAmount.Enabled = True Then TxtAmount.SetFocus
      Exit Sub
   End If
   If TxtchequeNo.Enabled = True Then
      If cn.Execute("Select ChequeNo from BankChequeIssueBody where ChequeNo = '" & TxtchequeNo.Text & "' and VoucherID <> " & TxtVoucherID.Text).EOF = False Then
         MsgBox "Cheque No. '" & TxtchequeNo.Text & "'  Already Exists in DataBase ", vbInformation + vbOKOnly, "Error"
         If TxtchequeNo.Enabled = True Then TxtchequeNo.SetFocus
         Exit Sub
      End If
   End If
   RsBody.Filter = "ChequeNo = '" & TxtchequeNo.Text & "'"
   If vIsNewRow = True Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ChequeNo").Text = TxtchequeNo.Text
         RsBody!VoucherID = TxtVoucherID.Text
      Else
         'If Grid.Columns("productid").Text <> TxtChequeNo.Text Then
            MsgBox "Current Record Already Exist ", vbInformation + vbOKOnly, "Alert"
            RsBody.Filter = 0
            Call SubClearDetailArea
            TxtSlipNo.SetFocus
            Exit Sub
            'Else
         End If
   End If
   With Grid
      TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
      .Columns("Amount").Text = Val(TxtAmount.Text)
      .Columns("ChequeNO").Text = TxtchequeNo.Text
      .Columns("ChequeDate").Text = dtpChequedate.DateValue
      .Columns("ReceiveBy").Text = TxtReceiveBy.Text
           
      RsBody!ChequeNo = TxtchequeNo.Text
      RsBody!ChequeDate = dtpChequedate.DateValue
      RsBody!DepositAmount = Val(TxtAmount.Text)
      RsBody!DepositBy = IIf((TxtReceiveBy.Text = ""), Null, TxtReceiveBy.Text)
            
      .MoveLast
      If Trim(.Columns("ChequeNo").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("ChequeNo").Text = " "
         .AllowAddNew = False
      End If
   End With
   vIsNewRow = True
   Call SubClearDetailArea
   If TxtchequeNo.Enabled = True Then TxtchequeNo.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Grid.Columns("ChequeNo").Text = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub
Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then GetDatabackFromGridToTextBoxes
End Sub

Private Sub MniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ChequeNo").Text) = "" Then Exit Sub
   RsBody.Filter = "ChequeNo = '" & Grid.Columns("ChequeNo").Text & " '"
   RsBody.Delete
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   RsBody.Filter = 0
   GetDatabackFromGridToTextBoxes
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("Amount").Value
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDatabackFromGridToTextBoxes()
   On Error GoTo ErrorHandler
   With Grid
      If Grid.Rows > 0 Then
         'TxtChequeNo.Text = .Columns("AcPayeeID").Text
         'TxtActPayeeName.Text = .Columns("ACPayeeName").Text
         TxtchequeNo.Text = .Columns("ChequeNO").Text
         dtpChequedate.DateValue = IIf(.Columns("ChequeDate").Value = Empty, Date, .Columns("ChequeDate").Value)
         TxtReceiveBy.Text = .Columns("ReceiveBy").Text
         TxtAmount.Text = .Columns("amount").Text
         
      End If
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub btnSave_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If TxtBankActID.Text = "" Then
      MsgBox " Please Specify Bank Account ", vbInformation + vbOKOnly, "Error"
      If TxtBankActID.Enabled = True Then TxtBankActID.SetFocus
    Exit Sub
    End If
    If TxtSlipNo.Text = "" Then
      MsgBox " Please Specify SlipNo ", vbInformation + vbOKOnly, "Error"
      If TxtSlipNo.Enabled = True Then TxtSlipNo.SetFocus
    Exit Sub
    End If
'    If TxtChequeNo.Text = "" Then
'      MsgBox " Please Specify Account PayeeID ", vbInformation + vbOKOnly, "Error"
'      If TxtChequeNo.Enabled = True Then TxtChequeNo.SetFocus
'    Exit Sub
 '   End If
   If vIsNewRecord Then
      If cn.Execute("Select * from BankChequeDepositHeader where VoucherID=" & Val(TxtVoucherID.Text)).RecordCount > 0 Then
         MsgBox "This Voucher ID already exists. A new Voucher ID. has been generated. Please try again", vbCritical, "Alert"
         TxtVoucherID.Text = FunGetMaxID
         Exit Sub
      End If
   End If
   
   
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where userno = " & vUser & " and Entrydate <='" & Date & "'"
    With cn.Execute(vStrSQL)
        If .Fields(0).Value >= DtpVoucherDate.DateValue Then
            MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
   
   Dim Rs As New ADODB.Recordset
   sSql = "select * from BankChequeDepositHeader where VoucherID = " & Val(TxtVoucherID.Text)
   Rs.Open sSql, cn, adOpenStatic, adLockPessimistic
   With Rs
      If .BOF Then
         .AddNew
         Rs!VoucherID = Val(TxtVoucherID.Text)
      End If
      !VoucherDate = DtpVoucherDate.DateValue
      !BankID = Val(TxtBankActID.Text)
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !SlipDate = dtpSlipDate.DateValue
      !SlipNo = Val(TxtSlipNo.Text)
      !Description = IIf((TxtDescription.Text = ""), Null, TxtDescription.Text)
      !UserNo = vUser
      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
      .Update
      .Close
   End With
   RsBody.Filter = 0
   RsBody.MoveFirst
   For vCounter = 1 To RsBody.RecordCount
      RsBody!VoucherID = TxtVoucherID.Text
      RsBody.Update
      RsBody.MoveNext
   Next vCounter
   RsBody.UpdateBatch
   RsBody.MoveFirst
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub btnBank_Click()
   On Error GoTo ErrorHandler
   If FunSelectAccount(ssButton, False) = True Then
      If TxtOrganizationID.Enabled And TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus
   Else
      TxtBankActID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchChequeDeposit.Show vbModal, Me
   If SchChequeDeposit.ParaOutVoucherNo <> 0 Then
      TxtVoucherID.Text = SchChequeDeposit.ParaOutVoucherNo
      GetCompeleteInfo
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetCompeleteInfo()
   On Error GoTo ErrorHandler
   sSql = "select H.VoucherID, H.VoucherDate, H.OrganizationID, OrganizationName, H.BankID, AccountName as BankName, H.Description, SlipNo, SlipDate, B.ChequeNo, B.ChequeDate, B.DepositBy, B.DepositAmount from BankChequeDepositHeader H inner Join BankChequeDepositBody B on H.VoucherID = B.VoucherID left outer join Organizations o on o.OrganizationID = h.OrganizationID inner join ChartofAccounts a on a.AccountNo = h.BankID where H.VoucherID = " & Val(TxtVoucherID.Text) & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With cn.Execute(sSql)
      If Not .BOF Then
         TxtVoucherID.Text = !VoucherID
         DtpVoucherDate.DateValue = !VoucherDate
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         TxtBankActID.Text = IIf(IsNull(!BankID), " ", !BankID)
         TxtBankActName.Text = IIf(IsNull(!BankName), "", !BankName)
         TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
         TxtSlipNo.Text = !SlipNo
         dtpSlipDate.DateValue = !SlipDate
      End If
      .Close
   End With
   PopulateDataToGrid
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub btnReceiving_Click()
   On Error GoTo ErrorHandler
   If FunSelectReceiving(ssButton, False) = True Then
      TxtReceiveBy.SetFocus
   Else
      TxtchequeNo.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtChequeNo_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtchequeNo.Name Then Exit Sub
   If TxtAmount.Text <> "" Then
      TxtReceiveBy.Text = ""
      TxtAmount.Text = ""
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtchequeNo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtChequeNo_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtchequeNo.Name Then Exit Sub
   If TxtAmount.Text <> "" Then Exit Sub
   If Trim(TxtchequeNo.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectReceiving(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectReceiving(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDescription_Change()
'If btnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub TxtReceiveBy_Change()
'If btnSave.Enabled = False Then FormStatus = ChangeMode
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
          If btnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectOrganization = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizationName.Text = ""
          If btnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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
      TxtSlipNo.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

