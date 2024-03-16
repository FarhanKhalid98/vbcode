VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmCreditVoucher 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmCreditVoucher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkIsPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Print"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10995
      TabIndex        =   51
      Top             =   8730
      Width           =   1290
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmCreditVoucher.frx":0ECA
      Left            =   5220
      List            =   "FrmCreditVoucher.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Tag             =   "1"
      Top             =   8685
      Width           =   3276
   End
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   8505
      TabIndex        =   46
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   8685
      Width           =   1170
   End
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9720
      TabIndex        =   45
      Top             =   8730
      Width           =   1290
   End
   Begin VB.TextBox TxtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   11693
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3510
      Width           =   1335
   End
   Begin VB.TextBox TxtNarration 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7583
      MaxLength       =   300
      TabIndex        =   7
      Top             =   3510
      Width           =   4110
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   3728
      MaxLength       =   100
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3510
      Width           =   3855
   End
   Begin VB.TextBox TxtVoucherNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   330
      Left            =   1958
      TabIndex        =   0
      Top             =   2610
      Width           =   1020
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   6915
      MaxLength       =   50
      TabIndex        =   9
      Top             =   8190
      Visible         =   0   'False
      Width           =   4125
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
      Height          =   4110
      Left            =   13725
      TabIndex        =   19
      Top             =   585
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
         Height          =   3750
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Tag             =   "NC"
         Text            =   "FrmCreditVoucher.frx":0ECE
         Top             =   360
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
         TabIndex        =   21
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtTotalAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      CausesValidation=   0   'False
      Height          =   330
      Left            =   12270
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8205
      Width           =   1695
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   4868
      TabIndex        =   12
      Top             =   9135
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
      MICON           =   "FrmCreditVoucher.frx":0FC1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6173
      TabIndex        =   11
      Top             =   9135
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
      MICON           =   "FrmCreditVoucher.frx":0FDD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10088
      TabIndex        =   15
      Top             =   9135
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
      MICON           =   "FrmCreditVoucher.frx":0FF9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7478
      TabIndex        =   10
      Top             =   9135
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
      MICON           =   "FrmCreditVoucher.frx":1015
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8775
      TabIndex        =   14
      Top             =   9135
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
      MICON           =   "FrmCreditVoucher.frx":1031
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   3563
      TabIndex        =   13
      Top             =   9135
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
      MICON           =   "FrmCreditVoucher.frx":104D
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpVoucherDate 
      Height          =   315
      Left            =   3023
      TabIndex        =   1
      Top             =   2610
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   4373
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   2610
      Width           =   495
      _ExtentX        =   873
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   4868
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2610
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
      MICON           =   "FrmCreditVoucher.frx":1069
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   5228
      TabIndex        =   27
      Tag             =   "NC"
      Top             =   2610
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
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   6668
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   2610
      Width           =   495
      _ExtentX        =   873
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
      Left            =   7523
      TabIndex        =   28
      Tag             =   "NC"
      Top             =   2610
      Width           =   1890
      _ExtentX        =   3334
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
      Left            =   7163
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2610
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
      MICON           =   "FrmCreditVoucher.frx":1085
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBankID 
      Height          =   315
      Left            =   9458
      TabIndex        =   4
      Top             =   2610
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtBankName 
      Height          =   315
      Left            =   10523
      TabIndex        =   31
      Top             =   2610
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
   Begin JeweledBut.JeweledButton BtnBank 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   10163
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
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
      MICON           =   "FrmCreditVoucher.frx":10A1
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSearch 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   3368
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3510
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmCreditVoucher.frx":10BD
      BC              =   12632256
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      CausesValidation=   0   'False
      Height          =   3990
      Left            =   1260
      TabIndex        =   38
      Top             =   3825
      Width           =   11775
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
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
      stylesets(0).Picture=   "FrmCreditVoucher.frx":10D9
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
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
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   6
      Columns(0).Width=   1376
      Columns(0).Caption=   "Serial"
      Columns(0).Name =   "Serial"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2461
      Columns(1).Caption=   "A/c No."
      Columns(1).Name =   "ID"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   6800
      Columns(2).Caption=   "A/c Name"
      Columns(2).Name =   "Name"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   7250
      Columns(3).Caption=   "Narration"
      Columns(3).Name =   "Narration"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "Amount"
      Columns(4).Name =   "Debit"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   5
      Columns(4).NumberFormat=   "########.##"
      Columns(4).FieldLen=   256
      Columns(5).Width=   1852
      Columns(5).Caption=   "Amount"
      Columns(5).Name =   "Credit"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   5
      Columns(5).NumberFormat=   "########.##"
      Columns(5).FieldLen=   256
      _ExtentX        =   20770
      _ExtentY        =   7038
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
   Begin SITextBox.Txt TxtSID 
      Height          =   315
      Left            =   1260
      TabIndex        =   49
      Top             =   2595
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
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   2340
      TabIndex        =   5
      Top             =   3510
      Width           =   1035
      _ExtentX        =   1826
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
      Masked          =   1
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "SID"
      Height          =   195
      Left            =   1260
      TabIndex        =   50
      Top             =   2385
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printer"
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
      Left            =   4545
      TabIndex        =   48
      Top             =   8685
      Width           =   570
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No."
      Height          =   225
      Left            =   6375
      TabIndex        =   44
      Top             =   8790
      Width           =   1095
   End
   Begin VB.Label LblWords 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   11265
      TabIndex        =   43
      Top             =   8700
      Width           =   2670
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c No."
      Height          =   225
      Left            =   2333
      TabIndex        =   42
      Top             =   3300
      Width           =   1020
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   225
      Left            =   11693
      TabIndex        =   41
      Top             =   3300
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Name"
      Height          =   225
      Left            =   3728
      TabIndex        =   40
      Top             =   3300
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Narration"
      Height          =   225
      Left            =   7568
      TabIndex        =   39
      Top             =   3300
      Width           =   900
   End
   Begin VB.Label LblBankID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank ID"
      Height          =   195
      Left            =   9458
      TabIndex        =   36
      Top             =   2385
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Date"
      Height          =   225
      Left            =   3023
      TabIndex        =   35
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   4373
      TabIndex        =   34
      Top             =   2385
      Width           =   585
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   6668
      TabIndex        =   33
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No."
      Height          =   225
      Left            =   1958
      TabIndex        =   30
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label LblBalance 
      Alignment       =   2  'Center
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
      Left            =   11633
      TabIndex        =   25
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Label LblBalanceCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
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
      Left            =   11633
      TabIndex        =   24
      Top             =   1575
      Width           =   1020
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
      Height          =   225
      Left            =   5880
      TabIndex        =   23
      Top             =   8235
      Visible         =   0   'False
      Width           =   900
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
      Left            =   12983
      TabIndex        =   22
      Top             =   1665
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Received Voucher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   18
      Top             =   270
      Width           =   2430
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      Height          =   225
      Left            =   11175
      TabIndex        =   17
      Top             =   8310
      Width           =   1020
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove this Row"
      End
   End
End
Attribute VB_Name = "FrmCreditVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vCounter, vGridRows As Integer
Dim Flag As Boolean, vBalance As Boolean
Dim sSql As String
Dim vStrSQL, vRandomID As String
Dim vMode As FormMode
Dim vMobileNo() As String, vMobile As String
Dim vIsNewRecord As Boolean
Dim vPrinter() As String

'----------------------------------

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   vStrSQL = " select h.voucherno, voucherdate, c.accountno, accountname + isnull(' - ' + p.city,'') as accountname, b.narration, amount, ServerEntry " & vbCrLf _
      + " from creditvouchers h inner join creditvouchersbody b" & vbCrLf _
      + " on h.voucherno = b.voucherno and h.storeid = b.storeid" & vbCrLf _
      + " inner join chartofaccounts c on c.accountno = b.accountno" & vbCrLf _
      + " left outer join parties p on p.partyid = b.accountno" & vbCrLf _
      + " where h.SID=" & TxtSID.Text

   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
  

'   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName2, ObjRegistry.DeviceName2, ObjRegistry.Port2
   
   vPrinter = Split(CmbPrinters.Text, ",")
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   
'   If ObjRegistry.LaserPrintofSaleInvoice = True Then
    If cmbPrintType.Text = "Half Page" Then
'      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\TransactionReports\CrpVoucherHalf.rpt")
      Set RptReportViewer.Report = New CrpVoucherHalf
      'Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\CrpVoucherHalf.rpt")
      RptReportViewer.Report.TopMargin = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))
      RptReportViewer.Report.LeftMargin = IIf(IsNull(ObjRegistry.x), 0, Val(ObjRegistry.x))
      RptReportViewer.Report.RightMargin = 225
'      RptReportViewer.Report.PaperSize = crPaperA4
      RptReportViewer.Report.PaperOrientation = crLandscape
'   ElseIf InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
   ElseIf cmbPrintType.Text = "Thermal" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\TransactionReports\CrpVoucherSmall.rpt")
'      Set RptReportViewer.Report = New CrpVoucher
      RptReportViewer.Report.PaperOrientation = crPortrait
   Else
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\TransactionReports\CrpVoucher.rpt")
'      Set RptReportViewer.Report = New CrpVoucherSmall
      RptReportViewer.Report.TopMargin = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))
      RptReportViewer.Report.LeftMargin = IIf(IsNull(ObjRegistry.x), 0, Val(ObjRegistry.x))
      RptReportViewer.Report.RightMargin = 0
   End If

   RptReportViewer.Report.ReportTitle = "Cash Received Voucher"
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity) & IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy
   
  '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Credit Voucher'" & "," & TxtVoucherNo.Text & ",'" & DtpVoucherDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'RptReportViewer.Show
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   If ObjRegistry.PreviewSaleInoice = True Or ChkIsPreview.Value = 1 Then
      If ChkIsPrint.Value = 1 Then
         RptReportViewer.Report.PrintOut False
      End If
       RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False
   End If
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and c.isDetailed = 1 and c.isLocked = 0"
        SchAccounts.ParaInAllowListSelection = False 'True
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
        TxtID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    If Trim(TxtID.Text) = "" Then Exit Function
    vStrSQL = " Select c.AccountNo, c.AccountName + isnull(' (' + p.Address + ')','') + isnull(' (' + p.City + ')','') as AccountName FROM ChartofAccounts c " & vbCrLf & _
         " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
         " Left Outer join Members m on c.AccountNo = cast(m.Prefix as varchar(2))  + cast(m.MemberID as varchar(10)) " & vbCrLf & _
         " where p.BarCode = '" & (TxtID.Text) & "' or m.BarCode = '" & (TxtID.Text) & "' or (c.AccountNo = " & Val(TxtID.Text) & " and c.isDetailed = 1 and c.isLocked = 0)"

    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtName.Text = !AccountName
          LblBalance.Caption = CN.Execute("SELECT isnull(dbo.FunCurrentDebit('" & Val(TxtID.Text) & "','" & DtpVoucherDate.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
          LblBalance.Caption = Abs(LblBalance.Caption) & " " & IIf(Val(LblBalance.Caption) >= 0, "Dr", "Cr")
          LblBalance.Visible = vBalance
          LblBalanceCaption.Visible = vBalance
          FunSelectAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectAccount = False
          .Close
          LblBalance.Visible = False
          LblBalanceCaption.Visible = False
          MsgBox "Invalid Account No.", vbOKOnly, "Alert"
          TxtID.Text = ""
          TxtName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnSearch_Click()
   On Error GoTo ErrorHandler
   If FunSelectAccount(ssButton, True) = True Then
      TxtNarration.SetFocus
   Else
      TxtID.SetFocus
   End If
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
  On Error GoTo ErrorHandler
  
   '''''''''''''''''' ActivityLogBin For Clear Action
'      Call DeleteTempActivityLogBin(vRandomID)
      vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.Rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("ID").Text) <> "" Then
           sSql = "Select AccountNo From CreditVouchersbody where VoucherNo=" & Val(TxtVoucherNo.Text) & " and StoreID = " & TxtStoreID.Text & " and AccountNo = " & Val(Grid.Columns("ID").Text)
            With CN.Execute(sSql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmCreditVoucher, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtVoucherNo.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpVoucherDate.Date), "Cleared Code-" & Grid.Columns("ID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("narration").Text))
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmCreditVoucher, eClearSavedRecord, TxtVoucherNo.Text, DtpVoucherDate.DateValue, vGridRows & " Account/s Cleared")
      Grid.Redraw = True
  ''''''''''''''''''
  '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  cn.Execute ("Insert Into UserActivities values ('Credit Voucher'" & "," & TxtVoucherNo.Text & ",'" & DtpVoucherDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   sSql = "select * from FormDefaultSetting Where FormType = 'Credit Voucher' and LocalComputerName = '" & LocalComputerName & "'"
   If CN.Execute(sSql).EOF Then
      sSql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview ) Values ('" & LocalComputerName & "', 'Credit Voucher','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & ")"
   Else
      sSql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & " Where FormType = 'Credit Voucher' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   CN.Execute sSql
   ''''''''''''''''''''''''''''''''''''''''''''
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Credit Voucher'" & "," & TxtVoucherNo.Text & ",'" & DtpVoucherDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniCashReceiveVoucher", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   '''''''''''''''''''''''Check Import / Export'''''''''''''''''''''''''''''''''
    If ObjRegistry.ShowMultiBranches = True Then
      vStrSQL = "select * from CreditVouchers where Tag is not null And SID=" & Val(TxtSID.Text) & " and Voucherdate = '" & DtpVoucherDate.DateValue & "'"
      With CN.Execute(vStrSQL)
          If Not .EOF Then
              MsgBox "Import/Export Record Cannot be Deleted", vbInformation, Me.Caption
              Exit Sub
          End If
      End With
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''

  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
  End If
  If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
  CN.BeginTrans
  
   Call BinData
   Call ActivityLogBin("", eFrmCreditVoucher, eDelete, TxtVoucherNo.Text, DtpVoucherDate.DateValue, Grid.Rows - 1 & " Accounts/s Deleted Amount: " & Val(TxtTotalAmount.Text))
   
  '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Credit Voucher'" & "," & TxtVoucherNo.Text & ",'" & DtpVoucherDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Call ActivityLog("Cash Received Voucher", eDelete, TxtVoucherNo.Text)
  CN.Execute "Delete from CreditVouchersBody where SID = " & Val(TxtSID.Text)
  CN.Execute "Delete from CreditVouchers WHere SID = " & Val(TxtSID.Text)
  
  If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnDelete Then
   vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = ObjRegistry.PrefixPhoneNo + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               sSql = ObjUserSecurity.UserName & " " & FrmCreditVoucher.Caption & " Deleted ID:" & TxtVoucherNo.Text & vbCrLf & " Date:" & Format(DtpVoucherDate.DateValue, "dd-MMM-yyyy") & " Time: " & Time & vbCrLf & " NetAmt: " & TxtTotalAmount.Text
               sSql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & sSql & "','')"
               CN.Execute sSql
            End If
         Next
   End If
   
  CN.CommitTrans
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  If CN.Errors.Count > 0 Then CN.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from Creditvouchersbody where voucherno = " & Val(TxtVoucherNo.Text) & " and StoreID = " & TxtStoreID.Text, CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = "Select Creditvouchersbody.*, AccountName + isnull(' (' + p.Address + ')','') AccountName from Creditvouchersbody inner join chartofaccounts on chartofaccounts.accountno = Creditvouchersbody.accountno Left Outer join Parties p on chartofaccounts.AccountNo = p.PartyID where voucherno = " & Val(TxtVoucherNo.Text) & " and Creditvouchersbody.StoreID = " & TxtStoreID.Text & " Order by SerialNo"
      With CN.Execute(sSql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("Serial").Text = Grid.Rows
            Grid.Columns("ID").Text = !AccountNo
            Grid.Columns("Name").Text = !AccountName
            Grid.Columns("Narration").Text = !Narration
            Grid.Columns("Credit").Value = !Amount
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + !Amount
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("ID").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
End Sub

Private Sub GetVoucher()
   On Error GoTo ErrorHandler
   sSql = "Select h.*, StoreName, OrganizationName, AccountName  from CreditVouchers h left outer join Stores s on s.storeid = h.storeid left outer join Organizations o on o.OrganizationID = h.OrganizationID left outer join ChartofAccounts ca on ca.AccountNo = h.BankID where h.SID = " & Val(TxtSID.Text) & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With CN.Execute(sSql)
      If Not .BOF Then
          TxtVoucherNo.Text = !voucherno
          DtpVoucherDate.DateValue = !VoucherDate
          TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
          TxtStoreName.Text = IIf(IsNull(!StoreName), "", !StoreName)
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          TxtBankID.Text = IIf(IsNull(!BankID), "", !BankID)
          TxtBankName.Text = IIf(IsNull(!AccountName), "", !AccountName)
          TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
      End If
      .Close
   End With
   Call PopulateDataToGrid
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchCV.Show vbModal, Me
   If SchCV.ParaOutVoucherNo <> Empty Then
      TxtVoucherNo.Text = SchCV.ParaOutVoucherNo
      TxtStoreID.Text = SchCV.ParaOutStoreID
      TxtSID.Text = SchCV.ParaOutSID
      '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      CN.Execute ("Insert Into UserActivities values ('Credit Voucher'" & "," & TxtVoucherNo.Text & ",'" & DtpVoucherDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      GetVoucher
   End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniCashReceiveVoucher", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   ''''''''''''''''''''''Check Import / Export'''''''''''''''''''''''''''''''''
    If ObjRegistry.ShowMultiBranches = True Then
      vStrSQL = "select * from CreditVouchers where Tag is not null And SID=" & Val(TxtSID.Text) & " and Voucherdate = '" & DtpVoucherDate.DateValue & "'"
      With CN.Execute(vStrSQL)
          If Not .EOF Then
              MsgBox "Import/Export Record Cannot be Updated", vbInformation, Me.Caption
              Exit Sub
          End If
      End With
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If CN.Execute("Select * From AdminClosing where ToUserNo = " & vUser & " and EntryDate = '" & DtpVoucherDate.DateValue & "'").RecordCount > 0 Then
      MsgBox "You are not authorized to Add Record in Closing Dates.", vbCritical, "Alert"
      Exit Sub
   End If

   If vIsNewRecord Then
      If CN.Execute("Select * from Creditvouchers where voucherno = " & Val(TxtVoucherNo.Text) & " and StoreID = " & TxtStoreID.Text).RecordCount > 0 Then
         MsgBox "This voucher already exists. A new voucher No. has been generated. Please try again", vbCritical, "Alert"
         TxtVoucherNo.Text = FunGetMaxID
         Exit Sub
      End If
      If TxtStoreID.Text = "" Then
         MsgBox "Please select store.", vbCritical, "Alert"
         If TxtStoreID.Visible And TxtStoreID.Enabled Then TxtStoreID.SetFocus
         Exit Sub
      End If
   End If
   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
       MsgBox "Please enter at least one entry to save", vbExclamation, "Alert"
       If TxtID.Visible And TxtID.Enabled Then TxtID.SetFocus
       Exit Sub
   End If
   
   ''''''''''''''''''''''''Check Organization '''''''''''''''''''''''''''''''''
  If ObjRegistry.OrganizationMandatory = True And TxtOrganizationID.Text = "" Then
    MsgBox "Please Select Organization", vbInformation, Me.Caption
    If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
    Exit Sub
  End If
  
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where touserno = " & vUser & " and Entrydate <='" & Date & "'"
    With CN.Execute(vStrSQL)
        If .Fields(0).Value >= DtpVoucherDate.DateValue Then
            MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
   '''''''''''''''''''''''Check Entry Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.isEntryDate = True Then
       If ObjRegistry.FromDate > Date Or ObjRegistry.ToDate < Date Then
         MsgBox "Data can not be saved Because Date is not set according to the Software's Entry date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
    
    '''''''''''''''''''''''Check Current Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.CurrentDateDataEntry = True And ObjUserSecurity.IsAdministrator = False Then
       If DtpVoucherDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
  
   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   sSql = "select * from FormDefaultSetting Where FormType = 'Credit Voucher' and LocalComputerName = '" & LocalComputerName & "'"
   If CN.Execute(sSql).EOF Then
      sSql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview, IsPrint ) Values ('" & LocalComputerName & "', 'Credit Voucher','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & "," & ChkIsPreview.Value & ")"
   Else
      sSql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & ", IsPrint = " & ChkIsPrint.Value & " Where FormType = 'Credit Voucher' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   CN.Execute sSql
   ''''''''''''''''''''''''''''''''''''''''''''
   CN.BeginTrans
   
   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmCreditVoucher, eEdit, TxtVoucherNo.Text, DtpVoucherDate.DateValue, "Amount: " & Val(TxtTotalAmount.Text))
   
'   Call UserActivities

   sSql = "Select * From CreditVouchers Where VoucherNo =" & Val(TxtVoucherNo.Text) & " and StoreID = " & TxtStoreID.Text
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, CN, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !voucherno = Val(TxtVoucherNo.Text)
         !StoreID = Val(TxtStoreID.Text)
         !UserNo = vUser
      End If
      !VoucherDate = DtpVoucherDate.DateValue
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !BankID = IIf(Val(TxtBankID.Text) = 0, Null, TxtBankID.Text)
      !Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
'      !UserNo = vUser
      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
      !IsSync = 0
      .Update
      .Close
      If vIsNewRecord = True Then TxtSID.Text = CN.Execute("select @@identity").Fields(0).Value
   End With
   If vIsNewRecord = False Then Call ActivityLog("Cash Received Voucher", eEdit, TxtVoucherNo.Text)
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !SID = Val(TxtSID.Text)
         !voucherno = Val(TxtVoucherNo.Text)
         !StoreID = Val(TxtStoreID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmCreditVoucher, eAdd, TxtVoucherNo.Text, DtpVoucherDate.DateValue, Grid.Rows - 1 & " New Account/s Added Amount: " & Val(TxtTotalAmount.Text))

'   If vIsNewRecord = True Then Call ActivityLog("Cash Received Voucher", eAdd, TxtVoucherNo.Text)
'   If MsgBox("Do you want to print this Voucher", vbQuestion + vbYesNo, "Alert") = vbYes Then
'      Call BtnPrint_Click
'   End If
   
   
   CN.CommitTrans
   If ChkIsPreview.Value = 1 Or ChkIsPrint.Value = 1 Then
      Call BtnPrint_Click
   End If
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
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
      vRandomID = Rnd() * 11111 & " " & Format(Now, "dd/mm hh:mm:ss")
      Call PopulateDataToGrid
      BtnPrint.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnSearch.Enabled = True
      LblBalance.Visible = False
      LblBalanceCaption.Visible = False
      TxtVoucherNo.Text = FunGetMaxID
      TxtStoreID.Enabled = True
      BtnStore.Enabled = True
      If DtpVoucherDate.Enabled And DtpVoucherDate.Visible Then DtpVoucherDate.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      BtnPrint.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnSearch.Enabled = True
      TxtStoreID.Enabled = False
      BtnStore.Enabled = False
      LblBalance.Visible = False
      LblBalanceCaption.Visible = False
      If DtpVoucherDate.Enabled And DtpVoucherDate.Visible Then DtpVoucherDate.SetFocus Else TxtID.SetFocus
      vIsNewRecord = False
    Case Is = ChangeMode
      BtnPrint.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub DtpVoucherDate_Change()
  If DtpVoucherDate.Enabled And DtpVoucherDate.Visible Then FormStatus = ChangeMode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = "Grid" Then
         Grid_DblClick
      Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtID.Name: If FunSelectAccount(ssFunctionKey, False) = True Then TxtNarration.SetFocus Else TxtID.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtOrganizationID.Enabled And TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus Else TxtStoreID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then If TxtBankID.Enabled Then TxtBankID.SetFocus Else TxtOrganizationID.SetFocus
         Case TxtBankID.Name: If FunSelectBank(ssFunctionKey, False) = True Then If TxtID.Enabled Then TxtID.SetFocus Else TxtBankID.SetFocus
     End Select
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtID.Enabled Then TxtID.SetFocus: Call SubClearDetailArea
   ElseIf Shift = vbCtrlMask Then
      If ActiveControl.Name = Grid.Name Then
         If KeyCode = vbKeyDelete Then
            If Trim(Grid.Columns("ID").Text <> "") Then Call mniRemoveRow_Click
            KeyCode = 0
         Else
            KeyCode = 0: Exit Sub
         End If
      End If
      Select Case KeyCode
      Case vbKeyS
            If BtnSave.Enabled And BtnSave.Visible Then BtnSave_Click
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
            If BtnDelete.Enabled And BtnDelete.Visible Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
      End Select
   Else
      If UCase(Me.ActiveControl.Name) Like "TXT*" Or UCase(Me.ActiveControl.Name) Like "DTP*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
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
   SetWindowText Me.hWnd, "Cash Received Vouchers"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   HelpLocation Me
   
   cmbPrintType.Clear
   cmbPrintType.AddItem "Full Page"
   cmbPrintType.AddItem "Half Page"
   cmbPrintType.AddItem "Thermal"
   cmbPrintType.ListIndex = 0
   
   CmbPrinters.Clear
   CmbPrinters.AddItem "Default,winspool,LPT1"
   Dim p
   For Each p In Printers
      CmbPrinters.AddItem p.DeviceName & "," & p.DriverName & "," & p.Port
   Next p
   CmbPrinters.ListIndex = 0

   
   '''''''''''''''' Form Default Setting  ''''''''''''''''''''''
   sSql = "select * from FormDefaultSetting Where FormType = 'Credit Voucher' and LocalComputerName = '" & LocalComputerName & "'"
   With CN.Execute(sSql)
     If .RecordCount > 0 Then
        cmbPrintType.Text = !Size
        ChkIsPreview.Value = Abs(!IsPreview)
        ChkIsPrint.Value = Abs(!IsPrint)
        If Not IsNull(!DeviceName) Then
            CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
        Else
            CmbPrinters.ListIndex = 0
        End If
     End If
     .Close
   End With
   ''''''''''''''''''''''''''''''''''''''''''''''
   
   DtpVoucherDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))
   
   TxtStoreID.Text = ObjRegistry.StoreID
   FunSelectStore ssValidate, True
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   LblStoreID.Visible = ObjRegistry.StoreVisible
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   
   With CN.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
      End If
      .Close
   End With
   DtpVoucherDate.Enabled = False
'   DtpVoucherDate.Enabled = ObjUserSecurity.IsAdministrator
   If ObjUserSecurity.IsAdministrator = True Or ObjRegistry.ChangeTransactionDate = True Then
      DtpVoucherDate.Enabled = True
   End If

   If ObjUserSecurity.IsAdministrator = True Or ObjRegistry.ShowBankInTransection = True Then
      LblBankID.Visible = True
      TxtBankID.Visible = True
      BtnBank.Visible = True
      TxtBankName.Visible = True
   End If
   
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus

   vBalance = ObjRegistry.PreviousBalanceVisible
   
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
  On Error GoTo ErrorHandler
   If ObjRegistry.AllowDailyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(voucherno),0) from Creditvouchers").Fields(0) + 1
   Else
      FunGetMaxID = CN.Execute("Select isnull(max(voucherno),0) from Creditvouchers where StoreID = " & TxtStoreID.Text).Fields(0) + 1
   End If
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is ComboBox Then
      End If
   Next
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ID").Text = " "
   Grid.Update
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
    Set RsBody = Nothing
    Set RsReport = Nothing
    Set FrmCreditVoucher = Nothing
   End If
   '''''''''''''''''' ActivityLogBin For Close Action
'      Call DeleteTempActivityLogBin(vRandomID)
      If Grid.Rows > 1 And Cancel = 0 Then
         vGridRows = 0
         Grid.Redraw = False
         Grid.MoveFirst
         For vCounter = 2 To Grid.Rows
            vGridRows = vGridRows + 1
            If Trim(Grid.Columns("ID").Text) <> "" Then
               sSql = "Select AccountNo From CreditVouchersbody where VoucherNo=" & Val(TxtVoucherNo.Text) & " and StoreID = " & TxtStoreID.Text & " and AccountNo = " & Val(Grid.Columns("ID").Text)
               With CN.Execute(sSql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmCreditVoucher, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtVoucherNo.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpVoucherDate.Date), "Closed Code-" & Grid.Columns("ID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("narration").Text))
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmCreditVoucher, eCloseSavedRecord, TxtVoucherNo.Text, DtpVoucherDate.DateValue, vGridRows & " Account/s Closed")
         Grid.Redraw = True
      End If
  ''''''''''''''''''
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("Credit").Value
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   Call Grid_LostFocus
End Sub

Private Sub Grid_GotFocus()
   Flag = True
   TxtID.Enabled = False
   BtnSearch.Enabled = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("ID").Text) = "" Then
      'TxtID.Text = ""
      TxtID.Enabled = True
      BtnSearch.Enabled = True
      TxtID.SetFocus
   Else
      TxtID.Enabled = False
      BtnSearch.Enabled = False
      TxtNarration.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(TxtID.Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu mnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ID").Text) = "" Then Exit Sub
   sSql = "Select AccountNo From CreditVouchersbody where VoucherNo=" & Val(TxtVoucherNo.Text) & " and StoreID = " & TxtStoreID.Text & " and AccountNo = " & Val(Grid.Columns("ID").Text)
   With CN.Execute(sSql)
      If .EOF Then
         Call ActivityLogBin("", eFrmCreditVoucher, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtVoucherNo.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpVoucherDate.Date), "Removed Code-" & Grid.Columns("ID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("narration").Text))
      Else
         Call ActivityLogBin("", eFrmCreditVoucher, eRemoveRow, TxtVoucherNo.Text, DtpVoucherDate.DateValue, "Removed Code-" & Grid.Columns("ID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("narration").Text))
         Call ActivityLogBin(vRandomID, eFrmCreditVoucher, eAddTempRecord, TxtVoucherNo.Text, DtpVoucherDate.DateValue, "Pending Remove Code-" & Grid.Columns("ID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("narration").Text))
      End If
   End With
   RsBody.Filter = " AccountNo = " & Val(Trim(TxtID.Text)) & " and Narration = '" & Trim(TxtNarration.Text) & "'"
   If RsBody.RecordCount > 0 Then RsBody.Delete
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Credit Voucher'" & "," & TxtVoucherNo.Text & ",'" & DtpVoucherDate.DateValue & "','Removed AccountNo-" & Grid.Columns("ID").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   Grid.Refresh
   RsBody.Filter = 0
   Grid.MoveLast
   GetDataBackFromGridToTexBoxes
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   Dim vrowcounter As Integer
   If Trim(TxtID.Text) = "" And Val(TxtAmount.Text) = 0 And Trim(TxtNarration.Text) = "" Then If TxtID.Enabled Then TxtID.SetFocus: Exit Sub
   If Trim(TxtID.Text) = "" Then
      MsgBox "Please provide an Account No.", vbExclamation, "Alert"
      TxtID.SetFocus
      Exit Sub
   End If
   If Val(TxtAmount.Text) <= 0 Then
      MsgBox "The amount must be greater than zero", vbExclamation, "Alert"
      TxtAmount.SetFocus
      Exit Sub
   End If
On Error GoTo ErrorHandler
   RsBody.Filter = " AccountNo = " & Val(TxtID.Text) & " and Narration = '" & Trim(TxtNarration.Text) & "'"
   
   If TxtID.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("Serial").Text = Grid.Rows
         Grid.Columns("ID").Text = Val(TxtID.Text)
         RsBody!AccountNo = Val(TxtID.Text)
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmCreditVoucher, eAddNewRowByEdit, TxtVoucherNo.Text, DtpVoucherDate.DateValue, "Add New Code-" & Val(TxtID.Text) & " Amount-" & " " & TxtAmount.Text & " " & Trim(TxtNarration.Text))
         Call ActivityLogBin(vRandomID, eFrmCreditVoucher, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtVoucherNo.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpVoucherDate.Date), "Pending Add New Code-" & Val(TxtID.Text) & " Amount-" & TxtAmount.Text & " " & Trim(TxtNarration.Text))
      Else
         MsgBox "This Record Already Exists. Please change the narration.", vbOKOnly + vbInformation, "Alert"
         TxtNarration.SetFocus
         Exit Sub
      End If
   Else
      If RsBody.RecordCount = 1 Then
         If Not (Trim(TxtID.Text) = Trim(Grid.Columns("ID").Text) And Trim(TxtNarration.Text) = Trim(Grid.Columns("narration").Text)) Then
            MsgBox "This Record Already Exists. Please change the narration.", vbOKOnly + vbInformation, "Alert"
            TxtNarration.SetFocus
            Exit Sub
         End If
      End If
      RsBody.Filter = " AccountNo = " & Val(Grid.Columns("ID").Text) & " and Narration = '" & Trim(Grid.Columns("narration").Text) & "'"
      sSql = "Select AccountNo From CreditVouchersbody where VoucherNo=" & Val(TxtVoucherNo.Text) & " and StoreID = " & TxtStoreID.Text & " and AccountNo = " & Val(Grid.Columns("ID").Text)
         With CN.Execute(sSql)
            If .EOF Then
               Call ActivityLogBin("", eFrmCreditVoucher, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtVoucherNo.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpVoucherDate.Date), "Effected Code-" & Grid.Columns("ID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("narration").Text))
               Call ActivityLogBin("", eFrmCreditVoucher, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtVoucherNo.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpVoucherDate.Date), "Updated Code-" & Val(TxtID.Text) & " Amount-" & TxtAmount.Text & " " & Trim(TxtNarration.Text))
            Else
               Call ActivityLogBin("", eFrmCreditVoucher, eEdit, TxtVoucherNo.Text, DtpVoucherDate.Date, "Effected Code-" & Grid.Columns("ID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("narration").Text))
               Call ActivityLogBin("", eFrmCreditVoucher, eEdit, TxtVoucherNo.Text, DtpVoucherDate.Date, "Updated Code-" & Val(TxtID.Text) & " Amount-" & TxtAmount.Text & " " & Trim(TxtNarration.Text))
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmCreditVoucher, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtVoucherNo.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpVoucherDate.Date), "Pending Update Code-" & Val(TxtID.Text) & " Amount-" & TxtAmount.Text & " " & Trim(TxtNarration.Text))
   End If
            
                  'MsgBox "The ID cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  Grid.Columns("Name").Text = TxtName.Text
                  Grid.Columns("Narration").Text = Trim(TxtNarration.Text)
                  Grid.Columns("Credit").Value = Val(TxtAmount.Text)
                  'RsBody!AccountNo = Grid.Columns("ID").Text
                  RsBody!Narration = Grid.Columns("narration").Text
                  RsBody!Amount = Val(Grid.Columns("Credit").Text)
                  Grid.MoveLast
                  With Grid
                    If Trim(.Columns("ID").Text) <> "" Then
                    .AllowAddNew = True
                    .AddNew
                    .Columns("ID").Text = " "
                    .AllowAddNew = False
                    End If
                End With
                  Call SubClearDetailArea
                
            
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearDetailArea
         TxtID.SetFocus
         Grid.Redraw = True
    Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtID.Text = .Columns("ID").Text
      TxtName.Text = .Columns("Name").Text
      TxtNarration.Text = .Columns("Narration").Text
      TxtAmount.Text = .Columns("Amount").Value
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtID.Enabled = True
   BtnSearch.Enabled = True
   TxtID.Text = ""
   TxtName.Text = ""
   TxtNarration.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub TxtAmount_LostFocus()
   On Error GoTo ErrorHandler
   Select Case ActiveControl.Name
   Case TxtID.Name, TxtNarration.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtID_Change()
   If ActiveControl.Name <> TxtID.Name Then Exit Sub
   If TxtName.Text <> "" Then TxtName.Text = ""
End Sub

Private Sub TxtID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtID.Text) = "" Then Exit Sub
   If Trim(TxtName.Text) <> "" Then Exit Sub
   vTemp = FunSelectAccount(ssValidate, False)
   If vTemp = False Then
      Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtNarration_GotFocus()
   If ActiveControl.Name <> TxtNarration.Name Then Exit Sub
   If Trim(TxtNarration.Text) = "" Then TxtNarration.Text = "Cash Received "
   TxtNarration.SelStart = 0
   TxtNarration.SelLength = Len(TxtNarration.Text)
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
   TxtVoucherNo.Text = FunGetMaxID
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
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
    With CN.Execute(vStrSQL)
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
      If TxtBankID.Enabled And TxtBankID.Visible Then TxtBankID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub TxtTotalAmount_Change()
   LblWords.Caption = StrConv(Words_Money_Only(Val(TxtTotalAmount.Text)), vbProperCase)
End Sub

Private Sub UserActivities()
    On Error GoTo ErrorHandler
    If vIsNewRecord = False Then
    With CN.Execute("Select  * from CreditVouchers where VoucherNo =" & TxtVoucherNo.Text & " And VoucherDate = '" & DtpVoucherDate.DateValue & "'")
        If TxtStoreID.Text <> !StoreID Then
            CN.Execute ("Insert Into UserActivities values ('Credit Voucher'" & "," & TxtVoucherNo.Text & ",'" & DtpVoucherDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.Rows - 1
        With CN.Execute("Select * from CreditVouchersBody Where VoucherNo = " & TxtVoucherNo.Text & " and AccountNo = " & Val(Grid.Columns("ID").Text))
             If .EOF = True Then
                CN.Execute ("Insert Into UserActivities values ('Credit Voucher'" & "," & TxtVoucherNo.Text & ",'" & DtpVoucherDate.DateValue & "','Inserted New AccountNo-" & Grid.Columns("ID").Text & " Amount-" & Grid.Columns("Credit").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("Credit").Text <> !Amount Then
                   CN.Execute ("Insert Into UserActivities values ('Credit Voucher'" & "," & TxtVoucherNo.Text & ",'" & DtpVoucherDate.DateValue & "','Updated AccountNo-" & Grid.Columns("ID").Text & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
   Else
    CN.Execute ("Insert Into UserActivities values ('Credit Voucher'" & "," & TxtVoucherNo.Text & ",'" & DtpVoucherDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectBank(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchAccounts.ParaInDetail = ""
      SchAccounts.ParaInWhereClause = " and (c.accountno like '1%' or c.accountno like '6%' or c.accountno like '5%')"
      'SchAccounts.cmbfilter.Text = "Banks"
      'SchAccounts.cmbfilter.Enabled = False
      SchAccounts.Show vbModal, Me
      If SchAccounts.ParaOutAccountNo = "" Then FunSelectBank = False: Exit Function
      TxtBankID.Text = SchAccounts.ParaOutAccountNo
   End If
   Dim vStrSQL As String
   vStrSQL = "select * from ChartofAccounts where AccountNo = " & Val(TxtBankID.Text) & " and (accountno like '1%' or accountno like '6%' or accountno like '5%')"
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtBankName.Text = !AccountName
         .Close
         FunSelectBank = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      Else
         FunSelectBank = False
         .Close
         TxtBankName.Text = ""
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnBank_Click()
   If FunSelectBank(ssButton, False) = True Then
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
   Else
      TxtBankID.SetFocus
   End If
End Sub

Private Sub TxtBankID_Change()
   If ActiveControl.Name <> TxtBankID.Name Then Exit Sub
   If TxtBankName.Text <> "" Then TxtBankName.Text = ""
End Sub

Private Sub TxtBankID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtBankID.Name Then Exit Sub
   If TxtBankID.Text = "" Then Exit Sub
   If TxtBankName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBank(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBank(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.CreditVouchersBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmCreditVoucher) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmCreditVoucher & ", " & vUser & "," & TableHeaderFields(eFrmCreditVoucher) & " from CreditVouchers " & vbCrLf _
             & "Where VoucherNo = " & TxtVoucherNo.Text & " and VoucherDate = '" & DtpVoucherDate.DateValue & "' and StoreID = " & TxtStoreID.Text
      CN.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.CreditVouchersBodyBin (" & TableBodyFields(eFrmCreditVoucher) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmCreditVoucher) & " from CreditVouchersBody " & vbCrLf _
             & "Where VoucherNo = " & TxtVoucherNo.Text & " and StoreID = " & TxtStoreID.Text
      CN.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub



