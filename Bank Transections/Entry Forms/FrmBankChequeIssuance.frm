VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmBankChequeIssuance 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   DrawMode        =   1  'Blackness
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTotalAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      CausesValidation=   0   'False
      Height          =   330
      Left            =   9349
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   8235
      Width           =   1695
   End
   Begin VB.CheckBox ChkReconcile 
      Caption         =   "Auto Reconcile"
      Height          =   255
      Left            =   11224
      TabIndex        =   27
      Top             =   3015
      Value           =   1  'Checked
      Width           =   1470
   End
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   7684
      TabIndex        =   5
      Top             =   2970
      Width           =   3495
      _ExtentX        =   6165
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
      Left            =   6521
      TabIndex        =   8
      Top             =   3690
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
      Left            =   9536
      TabIndex        =   9
      Top             =   3690
      Width           =   1410
      _ExtentX        =   2487
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
   Begin SITextBox.Txt TxtActPayeeName 
      Height          =   315
      Left            =   5096
      TabIndex        =   17
      Top             =   2970
      Width           =   2595
      _ExtentX        =   4577
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
   Begin JeweledBut.JeweledButton btnAccount 
      Height          =   315
      Left            =   4721
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2970
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
      MICON           =   "FrmBankChequeIssuance.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnClose 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10114
      TabIndex        =   15
      Top             =   8820
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
      MICON           =   "FrmBankChequeIssuance.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7444
      TabIndex        =   10
      Top             =   8820
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
      MICON           =   "FrmBankChequeIssuance.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6109
      TabIndex        =   11
      Top             =   8820
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
      MICON           =   "FrmBankChequeIssuance.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4774
      TabIndex        =   12
      Top             =   8820
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
      MICON           =   "FrmBankChequeIssuance.frx":0070
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtchequeNo 
      Height          =   315
      Left            =   3326
      TabIndex        =   6
      Top             =   3690
      Width           =   1815
      _ExtentX        =   3201
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
   Begin SITextBox.Txt TxtActPayeeID 
      Height          =   315
      Left            =   3326
      TabIndex        =   4
      Top             =   2970
      Width           =   1395
      _ExtentX        =   2461
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
   Begin JeweledBut.JeweledButton btndelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8779
      TabIndex        =   14
      Top             =   8820
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
      MICON           =   "FrmBankChequeIssuance.frx":008C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnPrint 
      Height          =   420
      Left            =   3439
      TabIndex        =   13
      Top             =   8820
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
      MICON           =   "FrmBankChequeIssuance.frx":00A8
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      CausesValidation=   0   'False
      Height          =   4215
      Left            =   3319
      TabIndex        =   19
      Top             =   4005
      Width           =   7965
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
      stylesets(0).Picture=   "FrmBankChequeIssuance.frx":00C4
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
      Columns(0).Width=   3200
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
      Columns(2).Caption=   "Receive By"
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
      _ExtentX        =   14049
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
   Begin SSCalendarWidgets_A.SSDateCombo dtpChequedate 
      Height          =   315
      Left            =   5141
      TabIndex        =   7
      Top             =   3690
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
   Begin JeweledBut.JeweledButton btnBank 
      Height          =   315
      Left            =   5936
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2115
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
      MICON           =   "FrmBankChequeIssuance.frx":00E0
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBankActName 
      Height          =   315
      Left            =   6311
      TabIndex        =   29
      Top             =   2115
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
      Left            =   4916
      TabIndex        =   2
      Top             =   2115
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
      Left            =   2666
      TabIndex        =   0
      Top             =   2115
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpVoucherDate 
      Height          =   315
      Left            =   3611
      TabIndex        =   1
      Top             =   2115
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
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   8366
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   2115
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
      Left            =   9671
      TabIndex        =   34
      Tag             =   "NC"
      Top             =   2115
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
      Left            =   9311
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2115
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
      MICON           =   "FrmBankChequeIssuance.frx":00FC
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      Height          =   225
      Left            =   8254
      TabIndex        =   39
      Top             =   8280
      Width           =   1020
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   9671
      TabIndex        =   37
      Top             =   1890
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   8366
      TabIndex        =   36
      Top             =   1890
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher ID"
      Height          =   195
      Left            =   2666
      TabIndex        =   33
      Top             =   1890
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Date"
      Height          =   195
      Left            =   3611
      TabIndex        =   32
      Top             =   1890
      Width           =   990
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C  ID"
      Height          =   195
      Left            =   4916
      TabIndex        =   31
      Top             =   1890
      Width           =   960
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C Name"
      Height          =   195
      Left            =   6311
      TabIndex        =   30
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Left            =   7691
      TabIndex        =   26
      Top             =   2730
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ReceiveBy"
      Height          =   255
      Left            =   6521
      TabIndex        =   25
      Top             =   3465
      Width           =   855
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Issuance"
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
      TabIndex        =   24
      Top             =   270
      Width           =   3030
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Date"
      Height          =   195
      Left            =   5141
      TabIndex        =   23
      Top             =   3465
      Width           =   945
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account payee Name"
      Height          =   195
      Left            =   5096
      TabIndex        =   22
      Top             =   2760
      Width           =   1545
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Payee ID"
      Height          =   255
      Left            =   3326
      TabIndex        =   21
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   9536
      TabIndex        =   20
      Top             =   3465
      Width           =   540
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   11610
      Top             =   60
      Width           =   345
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No"
      Height          =   195
      Left            =   3326
      TabIndex        =   18
      Top             =   3465
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
Attribute VB_Name = "FrmBankChequeIssuance"
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
   cn.Execute "Delete from BankChequeIssueBody where VoucherID = " & Val(TxtVoucherID.Text)
   cn.Execute "Delete from BankChequeIssueHeader where VoucherID = " & Val(TxtVoucherID.Text)
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
   vStrSQL = " Select H.VoucherID, H.VoucherDate, H.BankID, c.AccountName as BankName, " & vbCrLf _
    + " H.ACPayeeID, a.AccountName as ACPayeeName, " & vbCrLf _
    + " (a.AccountName + ' - ' + Cast(H.ACPayeeID as Varchar) )  Code, " & vbCrLf _
    + " B.ChequeNo, B.ChequeDate, B.ReceiveBy, B.Amount, h.Description" & vbCrLf _
    + " From BankChequeIssueHeader H " & vbCrLf _
    + " Inner Join BankChequeIssueBody B on H.VoucherID = B.VoucherID " & vbCrLf _
    + " inner join ChartofAccounts c on c.accountno = h.BankID " & vbCrLf _
    + " inner join ChartofAccounts a on a.accountno = H.ACPayeeID " & vbCrLf _
    + " where H.VoucherID = " & Val(TxtVoucherID.Text)

   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
  
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   
   Set RptReportViewer.Report = New CRptBankChequeIssu
   RptReportViewer.Report.PaperOrientation = crPortrait

   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1

   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy

   RptReportViewer.Show
   'RptReportViewer.Report.PrintOut False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkReconcile_Click()
   If ActiveControl.Name <> ChkReconcile.Name Then Exit Sub
   If btnSave.Enabled = False Then FormStatus = ChangeMode
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
         Set FrmBankChequeIssueance = Nothing
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
   SetWindowText Me.hWnd, "Bank Cheque Issuance"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
'   btnSave.Visible = Not ObjRegistry.ReadOnlyStatus
'   btndelete.Visible = Not ObjRegistry.ReadOnlyStatus
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
   dtpChequedate.DateValue = Date
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
   FunGetMaxID = cn.Execute("Select isnull(max(VoucherID),0) from BankChequeIssueHeader").Fields(0) + 1
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "select * from BankChequeIssueBody where VoucherID = ' " & Val(TxtVoucherID.Text) & " ' ", cn, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = "Select B.ChequeNo, B.ChequeDate, B.ReceiveBy, B.Amount From BankChequeIssueBody B  where B.VoucherID =" & Val(TxtVoucherID.Text)
      With cn.Execute(sSql)
         If .RecordCount > 0 Then
            Grid.Redraw = False
            Grid.MoveFirst
            Grid.RemoveAll
            Grid.AllowAddNew = True
            TxtTotalAmount.Text = 0
            While Not .EOF
               Grid.AddNew
               Grid.Columns("ChequeNo").Text = IIf(IsNull(!ChequeNo), "", !ChequeNo)
               Grid.Columns("ChequeDate").Text = (!ChequeDate)
               Grid.Columns("ReceiveBy").Text = IIf(IsNull(!ReceiveBy), "", !ReceiveBy)
               Grid.Columns("amount").Value = Val(!Amount)
               TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + !Amount
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
         Call SubClearDetailArea: TxtActPayeeID.SetFocus
      ElseIf KeyCode = vbKeyF1 Then
         Select Case ActiveControl.Name
            Case TxtBankActID.Name: If FunSelectAccount(ssFunctionKey, False) = True Then If TxtOrganizationID.Enabled And TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus Else TxtBankActID.SetFocus
            Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtActPayeeID.SetFocus Else TxtOrganizationID.SetFocus
            Case TxtActPayeeID.Name: If FunSelectPayee(ssFunctionKey, False) = True Then TxtDescription.SetFocus
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
      ElseIf ActiveControl.Name = TxtActPayeeID.Name Then
         If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtActPayeeID.Name Then
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
      TxtChequeNo.Text = ""
      TxtChequeNo.Enabled = True
      vIsNewRow = True
      If TxtChequeNo.Enabled Then TxtChequeNo.SetFocus
   Else
      TxtChequeNo.Enabled = False
      vIsNewRow = False
      dtpChequedate.SetFocus
   End If
End Sub

Private Sub Grid_GotFocus()
   Flag = True
   TxtChequeNo.Enabled = False
End Sub

Private Sub SubClearDetailArea()
   TxtChequeNo.Text = ""
   dtpChequedate.DateValue = Date
   TxtAmount.Text = ""
   TxtReceiveBy.Text = ""
End Sub

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchAccounts.ParaInDetail = ""
      SchAccounts.ParaInWhereClause = " and c.accountno like '1%'"
      'SchAccounts.cmbfilter.Text = "Banks"
      'SchAccounts.cmbfilter.Enabled = False
      SchAccounts.Show vbModal, Me
      If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
      TxtBankActID.Text = SchAccounts.ParaOutAccountNo
   End If
   Dim vStrSQL As String
   vStrSQL = "select * from ChartofAccounts where AccountNo =  '" & Val(TxtBankActID.Text) & "' and accountno like '1%'"
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

Private Function FunSelectPayee(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchAccounts.ParaInDetail = ""
      SchAccounts.ParaInWhereClause = "" '" and c.accountno like '12%'"
      SchAccounts.Show vbModal, Me
      If SchAccounts.ParaOutAccountNo = "" Then FunSelectPayee = False: Exit Function
      TxtActPayeeID.Text = SchAccounts.ParaOutAccountNo
   End If
   Dim vStrSQL As String
   vStrSQL = "select * from ChartofAccounts where isLocked = 0 and AccountNo = '" & (TxtActPayeeID.Text) & "'"
   With cn.Execute(vStrSQL)
         If .RecordCount > 0 Then
            TxtActPayeeName.Text = !AccountName
            .Close
            FunSelectPayee = True
            If btnSave.Enabled = False Then FormStatus = ChangeMode
            Exit Function
         Else
            FunSelectPayee = False
            .Close
            TxtActPayeeName.Text = ""
            If btnSave.Enabled = False Then FormStatus = ChangeMode
         End If
      End With
      Exit Function
ErrorHandler:
      Call ShowErrorMessage
End Function

Private Sub TxtActPayeeID_Change()
   If ActiveControl.Name <> TxtActPayeeID.Name Then Exit Sub
   If TxtActPayeeName.Text <> "" Then TxtActPayeeName.Text = ""
End Sub

Private Sub TxtActPayeeID_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtActPayeeID.Name Then Exit Sub
   If TxtActPayeeName.Text <> "" Then Exit Sub
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

Private Sub TxtBankActID_Change()
   If ActiveControl.Name <> TxtBankActID.Name Then Exit Sub
   If TxtBankActName.Text <> "" Then TxtBankActName.Text = ""
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
   If Trim(Grid.Columns("ChequeNo").Text) = "" Then
      vIsNewRow = True
      Flag = True
   Else
      vIsNewRow = False
   End If
   Call GetDataFromTextBoxesToGrid
End Sub

Private Sub GetDataFromTextBoxesToGrid()
   On Error GoTo ErrorHandler
   If Trim(TxtChequeNo.Text) = "" Then
      MsgBox " Please Specify ChequeNo ", vbInformation + vbOKOnly, "Error"
      If TxtChequeNo.Enabled = True Then TxtChequeNo.SetFocus
      Exit Sub
   End If
   If TxtAmount.Text = "" Then
      MsgBox " Please Specify Amount ", vbInformation + vbOKOnly, "Error"
      If TxtAmount.Enabled = True Then TxtAmount.SetFocus
      Exit Sub
   End If
   If TxtChequeNo.Enabled = True Then
      If cn.Execute("Select ChequeNo from BankChequeIssueBody where ChequeNo = '" & TxtChequeNo.Text & "' and VoucherID <> " & TxtVoucherID.Text).EOF = False Then
         MsgBox "Cheque No. '" & TxtChequeNo.Text & "'  Already Exists in DataBase ", vbInformation + vbOKOnly, "Error"
         If TxtChequeNo.Enabled = True Then TxtChequeNo.SetFocus
         Exit Sub
      End If
   End If
   RsBody.Filter = "ChequeNo = '" & TxtChequeNo.Text & "'"
   If vIsNewRow = True Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ChequeNo").Text = TxtChequeNo.Text
         RsBody!VoucherID = TxtVoucherID.Text
      Else
         'If Grid.Columns("productid").Text <> TxtActPayeeID.Text Then
            MsgBox "Current Record Already Exist ", vbInformation + vbOKOnly, "Alert"
            RsBody.Filter = 0
            Call SubClearDetailArea
            TxtActPayeeID.SetFocus
            Exit Sub
            'Else
         End If
   End If
   With Grid
      TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
      .Columns("Amount").Text = Val(TxtAmount.Text)
      .Columns("ChequeNO").Text = TxtChequeNo.Text
      .Columns("ChequeDate").Text = dtpChequedate.DateValue
      .Columns("ReceiveBy").Text = TxtReceiveBy.Text
      RsBody!ChequeNo = TxtChequeNo.Text
      RsBody!ChequeDate = dtpChequedate.DateValue
      RsBody!Amount = Val(TxtAmount.Text)
      RsBody!ReceiveBy = IIf((TxtReceiveBy.Text = ""), Null, TxtReceiveBy.Text)
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
   If TxtChequeNo.Enabled = True Then TxtChequeNo.SetFocus
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
         TxtChequeNo.Text = .Columns("ChequeNO").Text
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
    If TxtActPayeeID.Text = "" Then
      MsgBox " Please Specify Account PayeeID ", vbInformation + vbOKOnly, "Error"
      If TxtActPayeeID.Enabled = True Then TxtActPayeeID.SetFocus
    Exit Sub
    End If
   If vIsNewRecord Then
      If cn.Execute("Select * from BankChequeIssueHeader where VoucherID=" & Val(TxtVoucherID.Text)).RecordCount > 0 Then
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
   
   '''''''''''''''''''''''Check Organization'''''''''''''''''''''''''''''''''
   If ObjRegistry.OrganizationMandatory = True And TxtOrganizationID.Text = "" Then
      MsgBox "Please Select Organization", vbInformation, Me.Caption
      If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
      Exit Sub
   End If
   
   Dim Rs As New ADODB.Recordset
   sSql = "select * from BankChequeIssueHeader where VoucherID = " & Val(TxtVoucherID.Text)
   Rs.Open sSql, cn, adOpenStatic, adLockPessimistic
   With Rs
      If .BOF Then
         .AddNew
         Rs!VoucherID = Val(TxtVoucherID.Text)
      End If
      !VoucherDate = DtpVoucherDate.DateValue
      !BankID = Val(TxtBankActID.Text)
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !ACPayeeID = Val(TxtActPayeeID.Text)
      !Description = IIf((TxtDescription.Text = ""), Null, TxtDescription.Text)
      !AutoReconcile = ChkReconcile.Value
      !UserNo = vUser
      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
      .Update
      .Close
   End With
   RsBody.Filter = 0
   RsBody.MoveFirst
   For vCounter = 1 To RsBody.RecordCount
      RsBody!VoucherID = TxtVoucherID.Text
      RsBody!Reconcile = ChkReconcile.Value
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
   If FunSelectAccount(ssButton, False) = True Then
      If TxtOrganizationID.Enabled And TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus
   Else
      TxtBankActID.SetFocus
   End If
End Sub

Private Sub BtnOpen_Click()
   SchChequeIssue.Show vbModal, Me
   If SchChequeIssue.ParaOutVoucherNo <> 0 Then
      TxtVoucherID.Text = SchChequeIssue.ParaOutVoucherNo
      GetCompeleteInfo
   End If
   'dtpVoucherDate.SetFocus
End Sub

Private Sub GetCompeleteInfo()
   On Error GoTo ErrorHandler
   sSql = "select H.VoucherID, H.VoucherDate, h.OrganizationID, OrganizationName, h.AutoReconcile, H.BankID, c.AccountName as BankName, H.ACPayeeID, a.AccountName as ACPayeeName, H.Description, B.ChequeNo, B.ChequeDate, B.ReceiveBy, B.Amount from BankChequeIssueHeader H inner Join BankChequeIssueBody B on H.VoucherID = B.VoucherID inner join ChartofAccounts c on c.AccountNo = h.BankID inner join ChartofAccounts a on a.AccountNo = h.ACPayeeID left outer join Organizations o on o.OrganizationID = h.OrganizationID where H.VoucherID = " & Val(TxtVoucherID.Text) & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With cn.Execute(sSql)
      If Not .BOF Then
         TxtVoucherID.Text = !VoucherID
         DtpVoucherDate.DateValue = !VoucherDate
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         TxtBankActID.Text = IIf(IsNull(!BankID), " ", !BankID)
         TxtBankActName.Text = IIf(IsNull(!BankName), "", !BankName)
         TxtActPayeeID.Text = IIf(IsNull(!ACPayeeID), "", !ACPayeeID)
         TxtActPayeeName.Text = IIf(IsNull(!ACPayeeName), "", !ACPayeeName)
         TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
         ChkReconcile.Value = Abs(!AutoReconcile)
      End If
      .Close
   End With
   PopulateDataToGrid
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub btnAccount_Click()
   If FunSelectPayee(ssButton, False) = True Then
      TxtDescription.SetFocus
   Else
      TxtActPayeeID.SetFocus
   End If
End Sub

Private Sub TxtchequeNo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
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
      TxtActPayeeID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub
