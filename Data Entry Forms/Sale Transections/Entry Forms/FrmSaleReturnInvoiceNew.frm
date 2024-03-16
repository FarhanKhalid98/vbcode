VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmSaleReturnInvoiceNew 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmSaleReturnInvoiceNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   2070
      MaxLength       =   50
      TabIndex        =   104
      Top             =   7950
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.CheckBox ChkIsProduct 
      Caption         =   "Is Product"
      Height          =   255
      Left            =   5280
      TabIndex        =   97
      Top             =   360
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   3825
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3345
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   45
      TabIndex        =   72
      Top             =   6840
      Width           =   2295
      Begin SITextBox.Txt TxtSerial 
         Height          =   315
         Left            =   120
         TabIndex        =   73
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
         TabIndex        =   74
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
         stylesets(0).Picture=   "FrmSaleReturnInvoiceNew.frx":000C
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
         Columns(1).Caption=   "Serial (s)"
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
      Height          =   4380
      Left            =   11640
      TabIndex        =   66
      Top             =   840
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
         Height          =   3930
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   67
         Tag             =   "NC"
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
         TabIndex        =   68
         Top             =   90
         Width           =   135
      End
   End
   Begin SITextBox.Txt TxtPaidAmount 
      Height          =   315
      Left            =   10005
      TabIndex        =   32
      Top             =   7905
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
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
   Begin SITextBox.Txt TxtReturnID 
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   1185
      Width           =   1050
      _ExtentX        =   1852
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
      Left            =   7320
      TabIndex        =   37
      Top             =   8295
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":0028
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6015
      TabIndex        =   33
      Top             =   8295
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":0044
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3405
      TabIndex        =   35
      Top             =   8295
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":0060
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8625
      TabIndex        =   38
      Top             =   8295
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":007C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4710
      TabIndex        =   34
      Top             =   8295
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":0098
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   2955
      TabIndex        =   42
      Top             =   7185
      Width           =   1185
      _ExtentX        =   2090
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
   Begin SITextBox.Txt TxtBillDiscPer 
      Height          =   315
      Left            =   4140
      TabIndex        =   29
      Top             =   7185
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtNetAmount 
      Height          =   315
      Left            =   7275
      TabIndex        =   45
      Top             =   7185
      Width           =   1395
      _ExtentX        =   2461
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
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   900
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3345
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":00B4
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpReturnDate 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   1185
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   2610
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   1185
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
      Left            =   3645
      TabIndex        =   49
      Tag             =   "NC"
      Top             =   1200
      Width           =   1740
      _ExtentX        =   3069
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
      Left            =   3285
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1185
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":00D0
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2100
      TabIndex        =   36
      Top             =   8295
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":00EC
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPreviousPayable 
      Height          =   315
      Left            =   7215
      TabIndex        =   56
      Top             =   7905
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
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
   Begin SITextBox.Txt TxtTotalPayable 
      Height          =   315
      Left            =   8610
      TabIndex        =   58
      Top             =   7905
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
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
   Begin SITextBox.Txt TxtBillDisc 
      Height          =   315
      Left            =   5115
      TabIndex        =   30
      Top             =   7185
      Width           =   975
      _ExtentX        =   1720
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   6510
      TabIndex        =   60
      Top             =   345
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
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
   Begin SITextBox.Txt TxtTotalItems 
      Height          =   315
      Left            =   2070
      TabIndex        =   63
      Top             =   7185
      Width           =   885
      _ExtentX        =   1561
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
   Begin SITextBox.Txt TxtOtherCharges 
      Height          =   315
      Left            =   6090
      TabIndex        =   31
      Top             =   7185
      Width           =   1185
      _ExtentX        =   2090
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
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   1260
      TabIndex        =   16
      Top             =   3345
      Width           =   2565
      _ExtentX        =   4524
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
   Begin SITextBox.Txt TxtMultiplier 
      Height          =   315
      Left            =   5250
      TabIndex        =   18
      Top             =   3345
      Width           =   510
      _ExtentX        =   900
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
      Masked          =   1
   End
   Begin SITextBox.Txt TxtQtyLoose 
      Height          =   315
      Left            =   6270
      TabIndex        =   20
      Top             =   3345
      Width           =   540
      _ExtentX        =   953
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
   Begin SITextBox.Txt TxtQtyPack 
      Height          =   315
      Left            =   5760
      TabIndex        =   19
      Top             =   3345
      Width           =   510
      _ExtentX        =   900
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
      DecimalPoint    =   3
      IntegralPoint   =   4
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtDiscVal 
      Height          =   315
      Left            =   10140
      TabIndex        =   27
      Top             =   3345
      Width           =   675
      _ExtentX        =   1191
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
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtPrice 
      Height          =   315
      Left            =   7830
      TabIndex        =   23
      Top             =   3345
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
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   9660
      TabIndex        =   26
      Top             =   3345
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   8475
      TabIndex        =   24
      Top             =   3345
      Width           =   675
      _ExtentX        =   1191
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
   Begin SITextBox.Txt TxtBonus 
      Height          =   315
      Left            =   6810
      TabIndex        =   21
      Top             =   3345
      Width           =   540
      _ExtentX        =   953
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
   Begin SITextBox.Txt TxtOffer 
      Height          =   315
      Left            =   7350
      TabIndex        =   22
      Top             =   3345
      Width           =   480
      _ExtentX        =   847
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
   Begin SITextBox.Txt TxtSaleTaxVal 
      Height          =   315
      Left            =   9360
      TabIndex        =   86
      Top             =   360
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtSaleTaxPer 
      Height          =   315
      Left            =   9150
      TabIndex        =   25
      Top             =   3345
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   10815
      TabIndex        =   28
      Top             =   3345
      Width           =   990
      _ExtentX        =   1746
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
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   5400
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   1200
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
      Left            =   6480
      TabIndex        =   91
      Tag             =   "NC"
      Top             =   1200
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
      Left            =   6105
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   1200
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":0108
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtMemberID 
      Height          =   315
      Left            =   8280
      TabIndex        =   4
      Top             =   1200
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
   Begin SITextBox.Txt TxtMemberName 
      Height          =   315
      Left            =   9315
      TabIndex        =   93
      Top             =   1200
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
      Left            =   8955
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   1200
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":0124
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCost 
      Height          =   315
      Left            =   8355
      TabIndex        =   98
      Top             =   345
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
   Begin SITextBox.Txt TxtCommission 
      Height          =   315
      Left            =   7470
      TabIndex        =   100
      Top             =   345
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   45
      TabIndex        =   14
      Top             =   3345
      Width           =   855
      _ExtentX        =   1508
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3120
      Left            =   45
      TabIndex        =   103
      Top             =   3660
      Width           =   11895
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   19
      stylesets.count =   1
      stylesets(0).Name=   "Select"
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
      stylesets(0).Picture=   "FrmSaleReturnInvoiceNew.frx":0140
      AllowUpdate     =   0   'False
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
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   19
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ProductID"
      Columns(0).Name =   "ProductID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2143
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4498
      Columns(2).Caption=   "Product Name"
      Columns(2).Name =   "ProductName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2514
      Columns(3).Caption=   "Pack Name"
      Columns(3).Name =   "PackName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   900
      Columns(4).Caption=   "Pack"
      Columns(4).Name =   "Pack"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   900
      Columns(5).Caption=   "Q(P)"
      Columns(5).Name =   "QtyPack"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   4
      Columns(5).FieldLen=   256
      Columns(6).Width=   953
      Columns(6).Caption=   "Q(L)"
      Columns(6).Name =   "QtyLoose"
      Columns(6).Alignment=   1
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   4
      Columns(6).FieldLen=   256
      Columns(7).Width=   953
      Columns(7).Caption=   "Bns"
      Columns(7).Name =   "Bonus"
      Columns(7).Alignment=   1
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   847
      Columns(8).Caption=   "Offer"
      Columns(8).Name =   "Offer"
      Columns(8).Alignment=   1
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1138
      Columns(9).Caption=   "Price"
      Columns(9).Name =   "Price"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   4
      Columns(9).FieldLen=   256
      Columns(10).Width=   1191
      Columns(10).Caption=   "DiscPC"
      Columns(10).Name=   "DiscPC"
      Columns(10).Alignment=   1
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   926
      Columns(11).Caption=   "Tax%"
      Columns(11).Name=   "SaleTaxPer"
      Columns(11).Alignment=   1
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   847
      Columns(12).Caption=   "Dis%"
      Columns(12).Name=   "DiscPer"
      Columns(12).Alignment=   1
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   1191
      Columns(13).Caption=   "Dis.Val"
      Columns(13).Name=   "DiscVal"
      Columns(13).Alignment=   1
      Columns(13).CaptionAlignment=   2
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   4
      Columns(13).FieldLen=   256
      Columns(14).Width=   1508
      Columns(14).Caption=   "Amount"
      Columns(14).Name=   "Amount"
      Columns(14).Alignment=   1
      Columns(14).CaptionAlignment=   2
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   5
      Columns(14).FieldLen=   256
      Columns(15).Width=   3200
      Columns(15).Visible=   0   'False
      Columns(15).Caption=   "PackingID"
      Columns(15).Name=   "PackingID"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   3200
      Columns(16).Visible=   0   'False
      Columns(16).Caption=   "SaleTaxVal"
      Columns(16).Name=   "SaleTaxVal"
      Columns(16).Alignment=   1
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      Columns(17).Width=   3200
      Columns(17).Visible=   0   'False
      Columns(17).Caption=   "IsProduct"
      Columns(17).Name=   "IsProduct"
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   11
      Columns(17).FieldLen=   256
      Columns(18).Width=   3200
      Columns(18).Visible=   0   'False
      Columns(18).Caption=   "Cost"
      Columns(18).Name=   "Cost"
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20981
      _ExtentY        =   5503
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
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6660
      Top             =   7755
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin SITextBox.Txt TxtRemarks 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8640
      TabIndex        =   107
      Top             =   7185
      Width           =   1830
      _ExtentX        =   3228
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
      IntegralPoint   =   4
   End
   Begin SITextBox.Txt TxtManualBillNo 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10080
      TabIndex        =   109
      Top             =   8505
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridOffer 
      Height          =   1365
      Left            =   45
      TabIndex        =   99
      Top             =   5520
      Width           =   3705
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   4
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
      stylesets(0).Picture=   "FrmSaleReturnInvoiceNew.frx":015C
      UseGroups       =   -1  'True
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowColumnMoving=   2
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Groups(0).Width =   6059
      Groups(0).Caption=   "Product Offer"
      Groups(0).Columns.Count=   4
      Groups(0).Columns(0).Width=   2090
      Groups(0).Columns(0).Visible=   0   'False
      Groups(0).Columns(0).Caption=   "Product ID"
      Groups(0).Columns(0).Name=   "ProductID"
      Groups(0).Columns(0).CaptionAlignment=   2
      Groups(0).Columns(0).DataField=   "Column 0"
      Groups(0).Columns(0).DataType=   8
      Groups(0).Columns(0).FieldLen=   256
      Groups(0).Columns(0).Locked=   -1  'True
      Groups(0).Columns(1).Width=   1693
      Groups(0).Columns(1).Caption=   "Product ID"
      Groups(0).Columns(1).Name=   "ProductOfferID"
      Groups(0).Columns(1).DataField=   "Column 1"
      Groups(0).Columns(1).DataType=   8
      Groups(0).Columns(1).FieldLen=   256
      Groups(0).Columns(1).Locked=   -1  'True
      Groups(0).Columns(2).Width=   3440
      Groups(0).Columns(2).Caption=   "Product Name"
      Groups(0).Columns(2).Name=   "ProductName"
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   8
      Groups(0).Columns(2).FieldLen=   256
      Groups(0).Columns(2).Locked=   -1  'True
      Groups(0).Columns(3).Width=   926
      Groups(0).Columns(3).Caption=   "Qty"
      Groups(0).Columns(3).Name=   "Qty"
      Groups(0).Columns(3).Alignment=   1
      Groups(0).Columns(3).CaptionAlignment=   2
      Groups(0).Columns(3).DataField=   "Column 3"
      Groups(0).Columns(3).DataType=   2
      Groups(0).Columns(3).FieldLen=   256
      _ExtentX        =   6535
      _ExtentY        =   2408
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
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   330
      Left            =   255
      TabIndex        =   5
      Top             =   1890
      Width           =   930
      _ExtentX        =   2143
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
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   1545
      TabIndex        =   111
      Top             =   1890
      Width           =   3645
      _ExtentX        =   6429
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
      Left            =   5190
      TabIndex        =   112
      Top             =   1890
      Width           =   4530
      _ExtentX        =   7990
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
   Begin SITextBox.Txt TxtCity 
      Height          =   315
      Left            =   9720
      TabIndex        =   113
      Top             =   1890
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
   Begin JeweledBut.JeweledButton BtnCustomer 
      Height          =   330
      Left            =   1185
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   1890
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":0178
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   240
      TabIndex        =   119
      Top             =   675
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
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
      Masked          =   1
      Mandatory       =   1
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpBillDate 
      Height          =   315
      Left            =   885
      TabIndex        =   120
      Top             =   675
      Visible         =   0   'False
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
   Begin JeweledBut.JeweledButton BtnReturnAll 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2550
      TabIndex        =   121
      Top             =   660
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      TX              =   "Return All"
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":0194
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSale 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2190
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   660
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":01B0
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   8790
      TabIndex        =   12
      Top             =   2595
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
      Left            =   9900
      TabIndex        =   125
      Top             =   2595
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
      Left            =   9540
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2595
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":01CC
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillNo 
      Height          =   315
      Left            =   1395
      TabIndex        =   7
      Top             =   2610
      Width           =   1140
      _ExtentX        =   2011
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
   Begin SITextBox.Txt TxtBiltyNo 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   2610
      Width           =   1155
      _ExtentX        =   2037
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
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   2535
      TabIndex        =   8
      Top             =   2610
      Width           =   3495
      _ExtentX        =   6165
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
   Begin JeweledBut.JeweledButton BtnSaleMan 
      Height          =   330
      Left            =   6840
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2595
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
      MICON           =   "FrmSaleReturnInvoiceNew.frx":01E8
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtSaleManName 
      Height          =   315
      Left            =   7200
      TabIndex        =   11
      Top             =   2595
      Width           =   1590
      _ExtentX        =   2805
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
   Begin SITextBox.Txt TxtSaleManID 
      Height          =   315
      Left            =   6030
      TabIndex        =   9
      Top             =   2610
      Width           =   810
      _ExtentX        =   1429
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
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   9900
      TabIndex        =   132
      Top             =   2415
      Width           =   780
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   8790
      TabIndex        =   131
      Top             =   2415
      Width           =   525
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   2535
      TabIndex        =   130
      Top             =   2415
      Width           =   795
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bilty No."
      Height          =   195
      Left            =   240
      TabIndex        =   129
      Top             =   2400
      Width           =   585
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      Height          =   195
      Left            =   1395
      TabIndex        =   128
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Man ID"
      Height          =   255
      Left            =   6030
      TabIndex        =   127
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Man Name"
      Height          =   375
      Left            =   7200
      TabIndex        =   126
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   885
      TabIndex        =   124
      Top             =   480
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " Bill ID"
      Height          =   195
      Left            =   240
      TabIndex        =   123
      Top             =   480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      Height          =   195
      Left            =   240
      TabIndex        =   118
      Top             =   1680
      Width           =   870
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   195
      Left            =   1545
      TabIndex        =   117
      Top             =   1680
      Width           =   1125
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   5190
      TabIndex        =   116
      Top             =   1680
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   9720
      TabIndex        =   115
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill No"
      Height          =   195
      Left            =   10080
      TabIndex        =   110
      Top             =   8280
      Width           =   1020
   End
   Begin VB.Label LblRemarks 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   8640
      TabIndex        =   108
      Top             =   6960
      Width           =   630
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
      Left            =   2070
      TabIndex        =   106
      Top             =   7440
      Width           =   165
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
      Height          =   225
      Left            =   2070
      TabIndex        =   105
      Top             =   7710
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   8400
      TabIndex        =   102
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      Height          =   195
      Left            =   7395
      TabIndex        =   101
      Top             =   120
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblMemberID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      Height          =   195
      Left            =   8280
      TabIndex        =   96
      Top             =   960
      Width           =   780
   End
   Begin VB.Label LblMemberName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      Height          =   195
      Left            =   9315
      TabIndex        =   95
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   6585
      TabIndex        =   90
      Top             =   960
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   5400
      TabIndex        =   89
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   10815
      TabIndex        =   88
      Top             =   3150
      Width           =   540
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax%"
      Height          =   195
      Left            =   9150
      TabIndex        =   87
      Top             =   3150
      Width           =   390
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Val"
      Height          =   195
      Left            =   9360
      TabIndex        =   85
      Top             =   165
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Offer"
      Height          =   195
      Left            =   7350
      TabIndex        =   84
      Top             =   3150
      Width           =   345
   End
   Begin VB.Label LblPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   7830
      TabIndex        =   83
      Top             =   3150
      Width           =   405
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   5280
      TabIndex        =   82
      Top             =   3150
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   3855
      TabIndex        =   81
      Top             =   3150
      Width           =   840
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   6270
      TabIndex        =   80
      Top             =   3150
      Width           =   465
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   5760
      TabIndex        =   79
      Top             =   3150
      Width           =   480
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.Val"
      Height          =   195
      Left            =   10140
      TabIndex        =   78
      Top             =   3150
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Dis%"
      Height          =   195
      Left            =   9660
      TabIndex        =   77
      Top             =   3150
      Width           =   345
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc/PC"
      Height          =   195
      Left            =   8475
      TabIndex        =   76
      Top             =   3150
      Width           =   600
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bns(L)"
      Height          =   195
      Left            =   6810
      TabIndex        =   75
      Top             =   3150
      Width           =   450
   End
   Begin VB.Label LblRetailPrice 
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
      Left            =   8040
      TabIndex        =   71
      Top             =   1560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label LblCaptionRetailPrice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retail Price"
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
      Left            =   6480
      TabIndex        =   70
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   11205
      TabIndex        =   69
      Top             =   540
      Width           =   435
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      Height          =   195
      Left            =   6090
      TabIndex        =   65
      Top             =   6960
      Width           =   1020
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items"
      Height          =   195
      Left            =   2085
      TabIndex        =   64
      Top             =   6960
      Width           =   780
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Return Invoice"
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
      Left            =   1920
      TabIndex        =   62
      Top             =   150
      Width           =   3420
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   6510
      TabIndex        =   61
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LblTtlPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payable"
      Height          =   195
      Left            =   7320
      TabIndex        =   59
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Payable"
      Height          =   195
      Left            =   8674
      TabIndex        =   57
      Top             =   7680
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   5115
      TabIndex        =   55
      Top             =   6960
      Width           =   630
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
      Left            =   10800
      TabIndex        =   54
      Top             =   960
      Width           =   720
   End
   Begin VB.Label LblStock 
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
      Left            =   10800
      TabIndex        =   53
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   3645
      TabIndex        =   52
      Top             =   990
      Width           =   840
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   2610
      TabIndex        =   51
      Top             =   990
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   45
      TabIndex        =   48
      Top             =   3150
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   1260
      TabIndex        =   47
      Top             =   3150
      Width           =   1020
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label LblNetAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      Height          =   195
      Left            =   7290
      TabIndex        =   46
      Top             =   6960
      Width           =   840
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount (%)"
      Height          =   195
      Left            =   4125
      TabIndex        =   44
      Top             =   6960
      Width           =   885
   End
   Begin VB.Label LblTotalAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amount"
      Height          =   195
      Left            =   2985
      TabIndex        =   43
      Top             =   6960
      Width           =   990
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Amount"
      Height          =   195
      Left            =   10005
      TabIndex        =   41
      Top             =   7680
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      Height          =   195
      Left            =   1275
      TabIndex        =   40
      Top             =   990
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ReturnID"
      Height          =   195
      Left            =   225
      TabIndex        =   39
      Top             =   990
      Width           =   645
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmSaleReturnInvoiceNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vUnitPrice As Double
Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim vMaxBinID As Integer
Dim RsBody As New ADODB.Recordset
Dim RsBodySerial As New ADODB.Recordset
Dim RsProductOffer As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim QtyOffer As Integer
Dim Rebate As Integer
Dim DateFlag As Boolean
Dim Flag As Boolean
Dim sSql As String
Dim vStrSQL As String
Dim vBillID  As Integer
Dim vBillDate  As Date
Dim i As Integer
'----------------------------------

Private Sub SubCalculateBody()
   TxtDiscVal.Text = (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPC.Text)
   If Val(TxtDiscVal.Text) = 0 Then TxtDiscVal.Text = ""
   TxtSaleTaxVal.Text = Round(Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 2)
   TxtAmount.Text = Round((Val(vUnitPrice) - Val(TxtDiscPC.Text)) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 2) + Val(TxtSaleTaxVal.Text) - Val(TxtOffer.Text)
End Sub

Private Sub SubCalculateFooter()
   If TxtTotalAmount.Text = "" Then Exit Sub
   TxtNetAmount.Text = SelfRound(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtOtherCharges.Text)
   TxtTotalPayable.Text = Abs(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousPayable.Text, Val(TxtPreviousPayable.Text) * -1)))
   LblTtlPayable.Caption = IIf(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousPayable.Text, Val(TxtPreviousPayable.Text) * -1)) < 0, "Total Receivable", "Total Payable")
End Sub

Private Sub SubClearSerialFields()
   TxtSerial.Text = ""
   TxtSerial.Enabled = False
   GridSerial.CancelUpdate
   GridSerial.RemoveAll
   GridSerial.AddNew
   GridSerial.Columns("Serial").Text = " "
   GridSerial.Update
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


Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.ParaInWhere = "" 'and isLocked = 0"
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCode.Text) = "" Then Exit Function
    If Len(TxtCode.Text) <= 5 Then
      TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
    End If
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
        vStrSQL = " SELECT p.productid, Code, ProductName, PurPrice, RetailPrice, PurDiscPC, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
 
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !ProductID
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !RetailPrice
         TxtDiscPC.Text = IIf(IsNull(!PurDiscPC), "", !PurDiscPC)
         LblRetailPrice.Caption = !RetailPrice
         TxtCost.Text = 0
         ChkIsProduct.Value = 1
         If IsNull(!Packingname) Then
            vUnitPrice = !RetailPrice
            TxtMultiplier.Text = ""
            CmbPackName.ListIndex = 0
         Else
            TxtMultiplier.Text = !Multiplier
            If !Multiplier <> 0 Then
               vUnitPrice = !RetailPrice / !Multiplier
            End If
            CmbPackName.Text = !Packingname
         End If
         With CN.Execute("select QtyLoose from CurrentStockStore where ProductID ='" & TxtProductID.Text & "' and StoreID = " & TxtStoreID.Text)
            If .RecordCount > 0 Then
               LblStock.Caption = !QtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtProductID.Text & "')").Fields(0).Value
            Else
               LblStock.Caption = 0
            End If
         End With
'         With CN.Execute("select * from registry")
'         If .RecordCount > 0 Then
'            If !NegativeSale = False Then
'               If LblStock.Caption = "0" Then
'                  MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                  FunSelectProduct = False
'                  Exit Function
'               End If
'            End If
'         End If
'         .Close
'         End With
         LblStock.Visible = True
         LblStockCaption.Visible = True
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
         SubCalculateBody
         'Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtProductID.Text = ""
         TxtCode.Text = ""
         If CmbPackName.ListCount > 0 Then CmbPackName.ListIndex = 0
         TxtProductName.Text = ""
         TxtMultiplier.Text = ""
         TxtPrice.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtAmount.Text = ""
         LblStock.Visible = False
         LblStockCaption.Visible = False
         LblCaptionRetailPrice.Visible = False
         LblRetailPrice.Visible = False
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   FormStatus = NewMode
   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
   
   
   vMaxBinID = FunGetMaxBinID
   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
   CN.Execute ("Insert Into Bin_SaleReturnHeader Select " & vMaxBinID & ",'" & Date & "',* from SaleReturnHeader Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
   '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into Bin_SaleReturnBody Select " & vMaxBinID & ",'" & Date & "', * from SaleReturnBody Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
   '''''''''''''''''''''''''''''''''''''''''''''''Bin Serial''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into Bin_SaleReturnSerial Select " & vMaxBinID & ",'" & Date & "', * from SaleReturnSerial Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
   '''''''''''''''''''''''''''''''''''''''''''''''Bin ProductOffer''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into Bin_SaleReturnOffer Select " & vMaxBinID & ",'" & Date & "', * from SaleReturnOffer Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")

   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
   ''''''''''''''''''''''''''Delete Product OFfer'''''''''''''''''''''
   GridOffer.Redraw = False
   GridOffer.MoveFirst
   For vCounter = 1 To GridOffer.Rows
      If Trim(GridOffer.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from SaleReturnOffer where ReturnID = " & Val(TxtReturnID.Text) & " And ReturnDate ='" & DtpReturnDate.DateValue & "' and productid ='" & GridOffer.Columns("Productid").Text & "'"
      End If
      GridOffer.MoveNext
   Next vCounter
   GridOffer.RemoveAll
   GridOffer.Redraw = True
   
   ''''''''''''''''''''''''''Delete Serials'''''''''''''''''''''
   
   
    If RsBodySerial.RecordCount > 0 Then
        RsBodySerial.MoveFirst
        For vCounter = 1 To RsBodySerial.RecordCount
            CN.Execute "Delete from SaleReturnSerial where ReturnID = " & Val(TxtReturnID.Text) & " And ReturnDate ='" & DtpReturnDate.DateValue & "' and productid ='" & RsBodySerial!ProductID & "' and Serial ='" & RsBodySerial!Serial & "'"
            RsBodySerial.MoveNext
        Next vCounter
   End If
  
   ''''''''''''''''''''''''''Delete Purchase Body'''''''''''''''''''''
   Grid.Redraw = False
   Grid.MoveFirst
   Call ActivityLog("Purchase Invoice", eDelete, TxtReturnID.Text, DtpReturnDate.DateValue)
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from SaleReturnBody where ReturnID = " & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "' and productid ='" & Grid.Columns("Productid").Text & "'"
'          CN.Execute ("Insert Into Bin_SaleReturnBody Select " & FunGetMaxBinID & ", * from SaleReturnBody Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "' and productid ='" & Grid.Columns("Productid").Text & "'")
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   CN.Execute "Delete from SaleReturnHeader where ReturnID = " & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
   
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnEmployee_Click()
   On Error GoTo ErrorHandler
   If FunSelectEmployee(ssButton, False) = True Then
      If TxtMemberID.Visible = True Then TxtMemberID.SetFocus Else If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      TxtEmployeeID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      TxtBiltyNo.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

Private Sub BtnSaleman_Click()
   If FunSelectSaleman(ssButton, False) = True Then
      TxtCode.SetFocus
   Else
      TxtSalemanID.SetFocus
   End If

End Sub

Private Sub TxtSalemanID_Change()
 If TxtSalemanID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSalemanID.Name Then Exit Sub
   If TxtSalemanName.Text <> "" Then TxtSalemanName.Text = ""
End Sub

Private Sub TxtSalemanID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSalemanID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSalemanName.Text <> "" Then Exit Sub
   If Trim(TxtSalemanID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSaleman(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSaleman(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSaleman(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSaleman.Show vbModal, Me
        If SchSaleman.ParaOutSalemanID = "" Then FunSelectSaleman = False: Exit Function
        TxtSalemanID.Text = SchSaleman.ParaOutSalemanID
    End If
    '---------------------------
    vStrSQL = " Select * FROM SalesMan where SaleManID=" & Val(TxtSalemanID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSalemanName.Text = !SalemanName
          FunSelectSaleman = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSaleman = False
          .Close
          TxtSalemanID.Text = ""
          TxtSalemanName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function


Private Sub TxtCustomerID_Change()
   If TxtCustomerID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   If TxtCustomerName.Text <> "" Then
      TxtCustomerName.Text = ""
   End If
End Sub

Private Sub TxtCustomerID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerName.Text <> "" Then Exit Sub
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
        SchCustomer.Show vbModal, Me
        If SchCustomer.ParaOutCustomerID = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchCustomer.ParaOutCustomerID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Parties where PartyID = '" & TxtCustomerID.Text & "' AND PartyType <> 'v' and isLockParty=0"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerName.Text = !PartyName
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtPreviousPayable.Text = CN.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtCustomerID.Text & "','" & DtpReturnDate.DateValue & "'),0)").Fields(0).Value
          lblPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Previous Receivable", "Previous Payable")
          TxtPreviousPayable.Text = Abs(TxtPreviousPayable.Text)
          FunSelectCustomer = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectCustomer = False
          .Close
          TxtCustomerID.Text = ""
          TxtCustomerName.Text = ""
          TxtPreviousPayable.Text = ""
          lblPayable.Caption = "Previous Payable"
          LblTtlPayable.Caption = "Total Payable"
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtEmployeeID_Change()
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
    sSql = "Select *" & vbCrLf _
            + " from Employees" & vbCrLf _
            + " where isLockEmployee = 0 and EmpID=" & Val(TxtEmployeeID.Text)
    With CN.Execute(sSql)
      If .RecordCount > 0 Then
        TxtEmployeeName.Text = !EmpName
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


Private Sub BtnMember_Click()
  On Error GoTo ErrorHandler
   If FunSelectMember(ssButton, False) = True Then
      If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      TxtMemberID.SetFocus
   End If
   Exit Sub
ErrorHandler:
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
    sSql = "Select * " & vbCrLf _
            + " from Members" & vbCrLf _
            + " where IsLockMember = 0 and MemberID = " & Val(TxtMemberID.Text)
    With CN.Execute(sSql)
      If .RecordCount > 0 Then
        TxtMemberName.Text = !MemberName
        Call SubApplyMember
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

Private Sub TxtMemberID_Change()
   If ActiveControl.Name <> TxtMemberID.Name Then Exit Sub
   If TxtMemberName.Text <> "" Then TxtMemberName.Text = "": Call SubDestroyMember
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

Private Sub SubApplyMember()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   sSql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With CN.Execute(sSql)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
            RsBody.Filter = "ProductID='" & !ProductID & "'"
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
            Grid.Columns("DiscPC").Value = Round((Val(RsBody!Price) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) - Val(Grid.Columns("DiscVal").Value)
            
'            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - RsBody!Amount + Val(Grid.Columns("Amount").Text)
'            vTotDisc = vTotDisc - RsBody!DiscVal + Val(Grid.Columns("DiscVal").Text)
'            vTotalAmount = vTotalAmount - RsBody!Amount + Val(Grid.Columns("Amount").Text)
            
            RsBody!DiscPC = Val(Grid.Columns("DiscPC").Value)
            RsBody!DiscPer = Val(Grid.Columns("DiscPer").Value)
            RsBody!DiscVal = Val(Grid.Columns("DiscVal").Value)
            RsBody!Amount = Val(Grid.Columns("Amount").Value)
     
         End If
         Grid.MoveNext
      Wend
      .Close
   End With
   Grid.MoveLast
   SubCalculateFooter
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubDestroyMember()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   sSql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With CN.Execute(sSql)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
            
            RsBody.Filter = "ProductID='" & !ProductID & "'"
            Grid.Columns("DiscPer").Value = 0 'IIf(IsNull(!DiscPer), 0, !DiscPer)
            Grid.Columns("DiscPC").Value = 0 'Round((Val(RsBody!Price) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = 0 'Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) - Val(Grid.Columns("DiscVal").Value)
            
'            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - RsBody!Amount + Val(Grid.Columns("Amount").Text)
'            vTotDisc = vTotDisc - RsBody!DiscVal + Val(Grid.Columns("DiscVal").Text)
'            vTotalAmount = vTotalAmount - RsBody!Amount + Val(Grid.Columns("Amount").Text)
            
            RsBody!DiscPC = Val(Grid.Columns("DiscPC").Value)
            RsBody!DiscPer = Val(Grid.Columns("DiscPer").Value)
            RsBody!DiscVal = Val(Grid.Columns("DiscVal").Value)
            RsBody!Amount = Val(Grid.Columns("Amount").Value)
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


Private Sub BtnOpen_Click()
   SchSaleReturn.ParaInReturnDate = DtpReturnDate.DateValue
   SchSaleReturn.Show vbModal
   If SchSaleReturn.ParaOutReturnID <> -1 Then
      TxtReturnID.Text = SchSaleReturn.ParaOutReturnID
      'Dim a
      'a = Split(SchSaleReturn.ParaOutReturnDate, "/")
      DtpReturnDate.DateValue = SchSaleReturn.ParaOutReturnDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      CN.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetSaleReturn
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
'      If TxtCustomerID.Enabled Then TxtCustomerID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub


Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   vStrSQL = " select h.ReturnID, h.ReturnDate, Pr.PartyName + ' - ' + H.CustomerID as Cust_Name_ID, Pr.Address, StoreName, BillNo, Description," & vbCrLf _
            + " Isnull(H.TotalAmount, 0) TotalAmount, code, ProductName, dbo.FunSaleReturnSerial(b.ReturnID,b.ReturnDate,b.ProductId) Serial," & vbCrLf _
            + " dbo.FunSaleReturnOffer(b.ReturnID,b.ReturnDate, b.ProductId) ProductOffer, QtyPack, Multiplier, Qty," & vbCrLf _
            + " Isnull(B.Price,0) Price, Bonus,b.DiscPc, b.DiscPer, DiscVal, Offer, SaleTaxPer, SaleTaxval,Isnull(B.Amount,0) Amount, RetailPrice" & vbCrLf _
            + " from SaleReturnBody b inner join products p on b.productid = p.productid" & vbCrLf _
            + " inner join SaleReturnHeader h on h.ReturnID = b.ReturnID and h.ReturnDate = b.ReturnDate" & vbCrLf _
            + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
            + " inner join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
            + " where h.ReturnID = " & Val(TxtReturnID.Text) & " and h.ReturnDate='" & DtpReturnDate.DateValue & "'"
      
    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
  
    Set RptReportViewer.Report = New CryRptSaleReturnInvoiceNew
    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String
    vStrComp = "Select CompanyName,Address,City,PhoneNo,email from Company"
    With CN.Execute(vStrComp)
      If .RecordCount > 0 Then
         vCompanyName = !CompanyName
         vAddress = !Address & IIf(IsNull(!City), "", ", " & !City)
         vPhone = IIf(IsNull(!PhoneNo), "", "Phone # " & !PhoneNo)
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue vAddress
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue vPhone
      End If
   End With
   'RptReportViewer.Report.ParameterFields(4).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
   'RptReportViewer.Report.PrintOut False
   RptReportViewer.Show vbModal, Me
   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      CmbPackName.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
'  Header Validation
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      TxtStoreID.SetFocus
      Exit Sub
   End If
   If CN.Execute("Select * From AdminClosing where ToUserNo = " & vUser & " and EntryDate = '" & DtpReturnDate.DateValue & "'").RecordCount > 0 Then
      MsgBox "You are not authorized to Add Record in Closing Dates.", vbCritical, "Alert"
      Exit Sub
   End If
''''''''''''''''''''''''''''''''''
   If FrmReturnPrint.OptCash.Visible Then FrmReturnPrint.OptCash.SetFocus
   FrmReturnPrint.SubClearFields
   FrmReturnPrint.TxtNetAmount.Text = TxtNetAmount.Text
   FrmReturnPrint.Show vbModal, Me
   If FrmReturnPrint.ParaOutSelection = False Then Exit Sub
   'Body Validation
   'validation has been performed when a row is added to the grid
   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Sale Return", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
   End If
   
   If Val(TxtBillID.Text) <> 0 Then
      RsBody.Filter = 0
      RsBody.MoveFirst
      For vCounter = 1 To RsBody.RecordCount
         With CN.Execute("select isnull(Qty,0) as Qty from salebody where BillID = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "' and ProductID = '" & RsBody!ProductID & "'")
            If .RecordCount > 0 Then
               If !Qty < RsBody!Qty Then
                  MsgBox "Sale Quantity is less than Sale Return Quantity of Code = " & RsBody!Code, vbExclamation, "Alert"
                  If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
                  Exit Sub
               End If
            Else
               MsgBox "Sale of This Bill ID can not be found", vbExclamation, "Alert"
               If TxtBillID.Visible And TxtBillID.Enabled Then TxtBillID.SetFocus
               Exit Sub
            End If
         End With
         RsBody.MoveNext
      Next vCounter
   End If
   If DtpReturnDate.Date <> IIf(Format(Now, "hh") > 2, Date, DateAdd("d", -1, Date)) And DateFlag = True Then
      If MsgBox("Are you sure to Change Return Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
         DtpReturnDate.DateValue = IIf(Format(Now, "hh") > 2, Date, DateAdd("d", -1, Date))
      End If
      DateFlag = False
   End If
   If DtpReturnDate.Enabled Then
      If CN.Execute("Select * from SaleReturnHeader where ReturnID = " & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'").RecordCount > 0 Then
         'MsgBox "This Return ID already exists. A new Return ID. has been generated. Please try again", vbCritical, "Alert"
         TxtReturnID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
  
  'Saving record
   CN.BeginTrans
   If vIsNewRecord = False Then Call ActivityLog("Sale Return Invoice", eEdit, TxtReturnID.Text, DtpReturnDate.DateValue)
   Call UserActivities
   sSql = "select * from SaleReturnHeader where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, CN, adOpenStatic, adLockOptimistic
      If .BOF Then
         .AddNew
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
      End If
      !isReplace = 0
      !isPosted = 0
      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
      !SalemanID = IIf(Trim(TxtSalemanID.Text) = "", Null, TxtSalemanID.Text)
      !EmpComm = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text))
      !BillID = IIf(Val(TxtBillID.Text) = 0, Null, Val(TxtBillID.Text))
      !BillDate = IIf(IsNull(!BillID), Null, DtpBillDate.DateValue)
      !StoreID = TxtStoreID.Text
      !TotalAmount = Round(Val(TxtTotalAmount.Text))
      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
      If FrmReturnPrint.OptCash.Value = True Then
         !CashPaid = IIf(FrmReturnPrint.TxtCashPaid.Text = "", Null, Val(FrmReturnPrint.TxtCashPaid.Text))
         !CustomerID = "621"
         !CustomerName = IIf(Trim(FrmReturnPrint.TxtCashCustomer.Text) = "", Null, FrmReturnPrint.TxtCashCustomer.Text)
      End If
      If FrmReturnPrint.OptCredit.Value = True Then
         !CashPaid = 0
         !CustomerID = FrmReturnPrint.TxtCustomerID.Text
         !CustomerName = Null
      End If
      !Cash = FrmReturnPrint.OptCash.Value
      !Credit = FrmReturnPrint.OptCredit.Value
      !ManualBillNo = IIf(Trim(TxtManualBillNo.Text) = "", "", TxtManualBillNo.Text)
      !Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
      !UserNo = vUser
      .Update
      .Close
   End With
   
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   
   If RsBodySerial.RecordCount > 0 Then
     With RsBodySerial
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
         .MoveNext
      Next vCounter
      .UpdateBatch
     End With
   End If
   
   With RsProductOffer
      .Filter = 0
      If Not .EOF Then
        .MoveFirst
        For vCounter = 1 To .RecordCount
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
         .MoveNext
        Next vCounter
      End If
      .UpdateBatch
   End With

   If vIsNewRecord = True Then Call ActivityLog("Sale Return Invoice", eAdd, TxtReturnID.Text, DtpReturnDate.DateValue)
   CN.CommitTrans
'   Char.Speak "Thank you for comming"
   'If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
   If FrmReturnPrint.ChkPrint.Value = 1 Then Call BtnPrint_Click
   'End If
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleReturnBody where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = "select p.ProductName, code, b.* from SaleReturnBody b join products p on p.productid = b.productid where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
      With CN.Execute(sSql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !ProductID
            Grid.Columns("Code").Text = !Code
            Grid.Columns("ProductName").Text = !ProductName
            If !PackingID = 0 Or IsNull(!PackingID) Then
               Grid.Columns("PackingID").Value = ""
            Else
               Grid.Columns("PackingID").Value = !PackingID
            End If
            If !PackingID = 0 Or IsNull(!PackingID) Then
               Grid.Columns("PackName").Text = ""
            Else
               Grid.Columns("PackName").Text = CN.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
            End If
            Grid.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            Grid.Columns("QtyLoose").Value = !Qty
            Grid.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            Grid.Columns("Price").Value = !Price
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(!Amount)
            TxtTotalItems.Text = Val(TxtTotalItems.Text) + !Qty + IIf(IsNull(!Bonus), "0", !Bonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack))
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
   
   RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   RsBodySerial.Open "Select * from SaleReturnSerial where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
   Call PopulateDataToGridOffer
End Sub

Private Sub PopulateDataToGridserial()
   RsBodySerial.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
   If RsBodySerial.RecordCount > 0 Then
'       sSql = "select d.* from SaleReturnSerial d  where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "' and ProductID = '" & Grid.Columns("ProductID").Text & "'"
'      With CN.Execute(sSql)
       With RsBodySerial
         GridSerial.Redraw = False
         GridSerial.MoveFirst
         GridSerial.RemoveAll
         GridSerial.AllowAddNew = True
         .MoveFirst
         While Not .EOF
            GridSerial.AddNew
            GridSerial.Columns("ProductID").Text = !ProductID
            GridSerial.Columns("Serial").Text = !Serial
            .MoveNext
         Wend
'         .Close
      End With
      GridSerial.AddNew
      GridSerial.Columns("Serial").Text = " "
      GridSerial.AllowAddNew = False
      GridSerial.Redraw = True
   Else
    Call SubClearSerialFields
   End If
   RsBodySerial.Filter = 0
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
      TxtReturnID.Text = FunGetMaxID()
      Call PopulateDataToGrid
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      TxtCode.Enabled = True
      TxtStoreID.Enabled = True
      BtnStore.Enabled = True
      LblStock.Visible = False
      LblStockCaption.Visible = False
      BtnProduct.Enabled = True
      'TxtReturnID.Enabled = True
      DtpReturnDate.Enabled = True
      If DtpReturnDate.Enabled And DtpReturnDate.Visible Then DtpReturnDate.SetFocus
      GridOffer.Visible = False
      vIsNewRecord = True
   Case Is = OpenMode
      'TxtReturnID.Enabled = False
      DtpReturnDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      'TxtStoreID.Enabled = False
      'BtnStore.Enabled = False
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      'DtpReturnDate.SetFocus
      
      vIsNewRecord = False
   Case Is = ChangeMode
      BtnPrint.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtOrganizationID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub


Private Sub CmbPackName_Click()
   If CmbPackName.Text = "" Then
      TxtMultiplier.Enabled = False
      TxtQtyPack.Enabled = False
      TxtMultiplier.Text = ""
      TxtQtyPack.Text = ""
      TxtPrice.Text = Round(vUnitPrice, 3)
   Else
      TxtMultiplier.Enabled = True
      TxtQtyPack.Enabled = True
      If Trim(TxtCode.Text) <> "" Then
         With CN.Execute("select * from ProductPacking where productid='" & TxtProductID.Text & "' and packingid=" & CmbPackName.ItemData(CmbPackName.ListIndex))
            TxtMultiplier.Text = IIf(.RecordCount = 0, "", !Multiplier)
            If Val(TxtMultiplier.Text) <> 0 Then
               TxtPrice.Text = Round(vUnitPrice * !Multiplier, 3)
            Else
               TxtPrice.Text = Round(vUnitPrice, 3)
            End If
         .Close
         End With
      End If
   End If
End Sub

Private Sub DtpReturnDate_Validate(Cancel As Boolean)
   TxtReturnID.Text = FunGetMaxID()
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = "Grid" Then
         Grid_DblClick
      ElseIf ActiveControl.Name = "GridSerial" Then
         GridSerial_DblClick
      Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtCode.Enabled Then TxtCode.SetFocus: Call SubClearDetailArea
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
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then CmbPackName.SetFocus
         Case TxtCustomerID.Name: If FunSelectCustomer(ssFunctionKey, False) = True Then TxtBiltyNo.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtOrganizationID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtStoreID.SetFocus
         Case TxtSalemanID.Name: If FunSelectSaleman(ssFunctionKey, True) = True Then TxtProductID.SetFocus
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then If TxtMemberID.Visible = True Then If TxtCode.Enabled Then TxtCode.SetFocus
         Case TxtMemberID.Name: If FunSelectMember(ssFunctionKey, True) = True Then If TxtCode.Enabled Then TxtCode.SetFocus
      End Select
   ElseIf KeyCode = vbKeyF2 Then
         If Frame1.Visible = True Then
            Frame1.Visible = False
            If TxtCode.Enabled = True Then TxtCode.SetFocus Else Grid.SetFocus
        Else
            Frame1.Visible = True
            If TxtSerial.Enabled = True Then TxtSerial.SetFocus
        End If
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
         KeyCode = 0
         TxtBillDiscPer.SetFocus
      End If
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub



Private Sub Grid_Change()
If GridOffer.Visible = False Then Exit Sub
    If Me.ActiveControl.Name <> GridOffer.Name Then Exit Sub
    If GridOffer.Columns("Qty").Value = Val("") Then GridOffer.Columns("qty").Value = 0
End Sub

Private Sub GridOffer_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
    DispPromptMsg = 0
End Sub

Private Sub GridOffer_BeforeUpdate(Cancel As Integer)
    If GridOffer.Visible = False Then Exit Sub
'    If Me.ActiveControl.Name <> GridOffer.Name Then Exit Sub
    
    If GridOffer.Columns("ProductName").Text = "" Then Exit Sub
    If RsProductOffer.RecordCount = 0 Then Exit Sub
    RsProductOffer.Filter = 0
    RsProductOffer.Filter = "ProductID='" & GridOffer.Columns("ProductID").Text & "'"
    If RsProductOffer.RecordCount = 0 Then Exit Sub
    RsProductOffer!QtyOffer = Val(GridOffer.Columns("Qty").Value)
    RsProductOffer.Update
    RsProductOffer.Filter = 0
    If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub GridOffer_Click()
'    If Me.ActiveControl.Name <> GridOffer.Name Then Exit Sub
'    If GridOffer.Columns("ProductName").Text = "" Then GridOffer.Columns("Qty").Locked = True Else GridOffer.Columns("Qty").Locked = False: Exit Sub
End Sub

Private Sub GridOffer_DblClick()
    If Me.ActiveControl.Name <> GridOffer.Name Then Exit Sub
'    If GridOffer.Top = 5700 Then
'        GridOffer.Top = 1750
'        GridOffer.Left = 8250
'    Else
'        GridOffer.Top = 5700
'        GridOffer.Left = 45
'    End If
End Sub

Private Sub GridOffer_GotFocus()
    GridOffer.Row = 0
   GridOffer.Col = 0
   SendKeys "{Right}"
End Sub

Private Sub GridOffer_LostFocus()
'    If Me.ActiveControl.Name <> GridOffer.Name Then Exit Sub
    If Trim(Grid.Columns("Code").Text) <> "" Then Exit Sub
    GridOffer.MoveLast
End Sub

Private Sub GridOffer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(GridOffer.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
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
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Purchase Invoice"
   HelpLocation Me
   DtpReturnDate.DateValue = IIf(Format(Now, "hh") > 3, Date, DateAdd("d", -1, Date))
   
   With CN.Execute("Select * from Packings")
      CmbPackName.AddItem ""
      While Not .EOF
         CmbPackName.AddItem !Packingname
         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
         .MoveNext
      Wend
      .Close
   End With
   With CN.Execute("select * from registry")
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         If !StoreVisible = True Then
            LblStoreID.Visible = True
            LblStoreName.Visible = True
            TxtStoreID.Visible = True
            TxtStoreName.Visible = True
            BtnStore.Visible = True
         Else
            LblStoreID.Visible = False
            LblStoreName.Visible = False
            TxtStoreID.Visible = False
            TxtStoreName.Visible = False
            BtnStore.Visible = False
         End If
         LblEmpID.Visible = !EmpVisible
         LblEmpName.Visible = !EmpVisible
         TxtEmployeeID.Visible = !EmpVisible
         TxtEmployeeName.Visible = !EmpVisible
         BtnEmployee.Visible = !EmpVisible
         
         LblMemberID.Visible = !MemberVisible
         LblMemberName.Visible = !MemberVisible
         TxtMemberID.Visible = !MemberVisible
         TxtMemberName.Visible = !MemberVisible
         BtnMember.Visible = !MemberVisible
         
         TxtManualBillNo.Visible = !ManualBillNoVisible
         LblManualBillNo.Visible = !ManualBillNoVisible
         
         TxtRemarks.Visible = !RemarksVisible
         LblRemarks.Visible = !RemarksVisible
         If ObjUserSecurity.IsAdministrator = False Then
            LblPrice.Visible = Not !HidePurchaseAmount
            TxtPrice.Visible = Not !HidePurchaseAmount
            LblAmount.Visible = Not !HidePurchaseAmount
            TxtAmount.Visible = Not !HidePurchaseAmount
            LblTotalAmount.Visible = Not !HidePurchaseAmount
            TxtTotalAmount.Visible = Not !HidePurchaseAmount
            LblNetAmount.Visible = Not !HidePurchaseAmount
            TxtNetAmount.Visible = Not !HidePurchaseAmount
            lblPayable.Visible = Not !HidePurchaseAmount
            TxtPreviousPayable.Visible = Not !HidePurchaseAmount
            LblTtlPayable.Visible = Not !HidePurchaseAmount
            TxtTotalPayable.Visible = Not !HidePurchaseAmount
            Grid.Columns("Price").Visible = Not !HidePurchaseAmount
            Grid.Columns("Amount").Visible = Not !HidePurchaseAmount
         End If
      End If
      .Close
   End With
   DateFlag = True
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpReturnDate.IsDateValid = False Then Exit Function
   FunGetMaxID = CN.Execute("Select isnull(max(ReturnID),0)+1 from SaleReturnHeader where ReturnDate = '" & DtpReturnDate.DateValue & "'").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpReturnDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = CN.Execute("Select isnull(max(BinID),0)+1 from Bin_SaleReturnHeader ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function


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
   TxtNetAmount.Text = 0
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("Code").Text = " "
   Grid.Update
   
   GridOffer.CancelUpdate
   GridOffer.RemoveAll
   GridOffer.AddNew
   GridOffer.Columns("ProductID").Text = " "
   GridOffer.Update
   GridOffer.Visible = False
   Call SubClearSerialFields
   Frame1.Visible = False
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
    Set RsBodySerial = Nothing
    Set FrmSaleReturnInvoice = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("Amount").Value
   TxtTotalItems.Text = Val(TxtTotalItems.Text) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Grid.Columns("QtyPack").Value)))
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
   TxtCode.Enabled = False
   BtnProduct.Enabled = False
   'TxtCode.BackColor = TxtProductName.BackColor
   'TxtCode.TabStop = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("Code").Text) = "" Then
      TxtCode.Text = ""
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      TxtCode.SetFocus
   Else
      TxtCode.Enabled = False
      BtnProduct.Enabled = False
      If Me.ActiveControl.Name = Grid.Name Then CmbPackName.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(Grid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    If Flag Then Call GetDataBackFromGridToTexBoxes
    Call PopulateDataToGridserial
    If Trim(Grid.Columns("Code").Text) = "" Then
        TxtSerial.Enabled = False
    Else
        TxtSerial.Enabled = True
    End If
    
    GridOffer.MoveFirst
    For vCounter = 1 To GridOffer.Rows
    If GridOffer.Columns("ProductID").Text = Grid.Columns("ProductID").Text Then Exit Sub
        GridOffer.MoveNext
    Next vCounter
   
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name = "Grid" Then
    If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
    RsBody.Filter = "ProductID = '" & TxtProductID.Text & "' and Price = " & Val(TxtPrice.Text)
    If RsBody.RecordCount > 0 Then RsBody.Delete
    RsProductOffer.Filter = "ProductID='" & GridOffer.Columns("ProductID").Text & "'"
    If RsProductOffer.RecordCount > 0 Then
        RsProductOffer.Delete
        GridOffer.SelBookmarks.RemoveAll
        GridOffer.SelBookmarks.Add GridOffer.Bookmark
        GridOffer.DeleteSelected
        GridOffer.Refresh
        RsProductOffer.Filter = 0
    End If
    RsBodySerial.Filter = "ProductID ='" & TxtProductID.Text & "' And Serial = '" & TxtSerial.Text & "'"
    If RsBodySerial.RecordCount > 0 Then
      With RsBodySerial
'      .Filter = 0
      .MoveFirst
        For vCounter = 1 To .RecordCount
            RsBodySerial.Delete
            .MoveNext
        Next vCounter
      End With
    End If
    CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
    Grid.SelBookmarks.RemoveAll
    Grid.SelBookmarks.Add Grid.Bookmark
    Grid.DeleteSelected
    Grid.Refresh
    RsBody.Filter = 0
    Grid.MoveLast
    GetDataBackFromGridToTexBoxes
   ElseIf Me.ActiveControl.Name = "GridSerial" Then
    If Trim(GridSerial.Columns("Serial").Text) = "" Then Exit Sub
    RsBodySerial.Filter = "Serial = '" & TxtSerial.Text & "'"
    If RsBodySerial.RecordCount > 0 Then RsBodySerial.Delete
    CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed ProdcutID-" & GridSerial.Columns("ProductID").Text & " Serial-" & GridSerial.Columns("Serial").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
    GridSerial.SelBookmarks.RemoveAll
    GridSerial.SelBookmarks.Add GridSerial.Bookmark
    GridSerial.DeleteSelected
    GridSerial.Refresh
    RsBodySerial.Filter = 0
'    GridSerial.MoveLast
    GetDataBackFromGridSerialToTexBoxes
   ElseIf Me.ActiveControl.Name = GridOffer.Name Then
    RsProductOffer.Filter = "ProductID='" & GridOffer.Columns("ProductID").Text & "'"
    If RsProductOffer.RecordCount > 0 Then
        RsProductOffer.Delete
        GridOffer.SelBookmarks.RemoveAll
        GridOffer.SelBookmarks.Add GridOffer.Bookmark
        GridOffer.DeleteSelected
        GridOffer.Refresh
        RsProductOffer.Filter = 0
    End If
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataBackFromGridSerialToTexBoxes()
   On Error GoTo ErrorHandler
   With GridSerial
      TxtSerial.Text = .Columns("Serial").Text
   End With
   If GridSerial.Rows = 1 Then GridSerial.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   Dim vrowcounter As Integer
   If Trim(TxtCode.Text) = "" Then
      TxtCode.SetFocus
      Exit Sub
   End If
   If CmbPackName.ListIndex > 0 Then
      If Trim(TxtMultiplier.Text) = 0 Then
         TxtMultiplier.SetFocus
         Exit Sub
      End If
   End If
   If Val(TxtQtyPack.Text) = 0 And Val(TxtQtyLoose.Text) = 0 Then
      If TxtQtyPack.Enabled Then TxtQtyPack.SetFocus Else TxtQtyLoose.SetFocus
      Exit Sub
   End If
On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Productid").Text) = "" Then
      RsBody.Filter = "ProductID='" & TxtProductID.Text & "' and Price = " & Val(TxtPrice.Text)
   Else
      RsBody.Filter = "ProductID='" & Grid.Columns("Productid").Text & "' and Price = " & Val(Grid.Columns("Price").Text)
   End If
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ProductID").Text = TxtProductID.Text
         Grid.Columns("Code").Text = TxtCode.Text
         Grid.Columns("Price").Value = Val(TxtPrice.Text)
         RsBody!ProductID = TxtProductID.Text
         RsBody!Code = TxtCode.Text
         RsBody!Price = Val(TxtPrice.Text)
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("Productid").Text = TxtProductID.Text And Val(Grid.Columns("Price").Text) = Val(TxtPrice.Text) Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  
                  '''''''''''''''''''''''''This QtyOffer Is used for DetailGrid
                  QtyOffer = Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value) + Val(Grid.Columns("QtyLoose").Value)
                  GetDataFromTextBoxesToGridOffer
                  TxtOffer.Text = Val(TxtOffer.Text) + Val(Grid.Columns("Offer").Text)
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  
                  TxtQtyLoose.Text = Val(TxtQtyLoose.Text) + Val(Grid.Columns("QtyLoose").Value)
                  TxtQtyPack.Text = Val(TxtQtyPack.Text) + Val(Grid.Columns("QtyPack").Value)
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Val(Grid.Columns("QtyLoose").Value) + Val(Grid.Columns("Bonus").Value) + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
                  Grid.Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
                  Grid.Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, 0, Val(TxtMultiplier.Text))
                  Grid.Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
                  Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
                  Grid.Columns("Bonus").Value = Val(TxtBonus.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  Grid.Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  Grid.Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  Grid.Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  Grid.Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  Grid.Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("Cost").Value = Val(TxtCost.Text)
                  Grid.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
                  RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
                  RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
                  RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
                  RsBody!Qty = Val(TxtQtyLoose.Text)
                  RsBody!Bonus = Val(TxtBonus.Text)
                  RsBody!Price = Val(TxtPrice.Text)
                  RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  RsBody!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  RsBody!SaleTaxval = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  RsBody!Amount = Val(TxtAmount.Text)
                  RsBody!Cost = Val(TxtCost.Text)
                  RsBody!IsProduct = Abs(Grid.Columns("isProduct").Value)
                  Grid.MoveLast

                  Call SubClearDetailArea
                  TxtCode.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearDetailArea
         Grid.MoveLast
         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   Grid.Redraw = False
   With Grid
      If TxtCode.Enabled = True Then
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text)))
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Val(Grid.Columns("Pack").Value)) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
      End If
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
      .Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, 0, Val(TxtMultiplier.Text))
      .Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
      .Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
      .Columns("Bonus").Value = Val(TxtBonus.Text)
      .Columns("Price").Value = Val(TxtPrice.Text)
      .Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      .Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
      .Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
      .Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      .Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      .Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("Cost").Value = Val(TxtCost.Text)
      .Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
      RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
      RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
      RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
      RsBody!Qty = Val(TxtQtyLoose.Text)
      RsBody!Bonus = Val(TxtBonus.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      RsBody!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
      RsBody!SaleTaxval = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
      RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      RsBody!Amount = Val(TxtAmount.Text)
      RsBody!Cost = Val(TxtCost.Text)
      RsBody!IsProduct = Abs(Grid.Columns("isProduct").Value)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   QtyOffer = 0
   GetDataFromTextBoxesToGridOffer
   Call SubClearDetailArea
   TxtCode.SetFocus
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtCode.Enabled = True
   BtnProduct.Enabled = True
   TxtCode.Text = ""
   TxtProductName.Text = ""
   CmbPackName.ListIndex = 0
   TxtMultiplier.Text = ""
   TxtQtyPack.Text = ""
   TxtQtyLoose.Text = ""
   TxtBonus.Text = ""
   TxtPrice.Text = ""
   TxtOffer.Text = ""
   TxtSaleTaxPer.Text = ""
   TxtSaleTaxVal.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtDiscVal.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub GetDataFromTextBoxesToGridOffer()
On Error GoTo ErrorHandler
    With CN.Execute("Select * from ProductOffers where Rebate = 0 and ProductID = '" & TxtProductID.Text & "'")
            
        If .RecordCount > 0 Then
            QtyOffer = QtyOffer + Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)
            QtyOffer = QtyOffer \ !Qty * !QtyOffer
            If QtyOffer > 0 Then
                
                RsProductOffer.Filter = "ProductID='" & TxtProductID.Text & "'"
                If TxtProductID.Enabled Then
                    If RsProductOffer.RecordCount = 0 Then
                        RsProductOffer.AddNew
                        GridOffer.Columns("ProductID").Text = TxtProductID.Text
                        GridOffer.Columns("ProductOfferID").Text = !ProductOfferID
                        RsProductOffer!ProductID = TxtProductID.Text
                        RsProductOffer!ProductOfferID = !ProductOfferID
                    Else
                        GridOffer.Redraw = False
                        GridOffer.MoveFirst
                        For vCounter = 1 To GridOffer.Rows
                        If GridOffer.Columns("ProductID").Text = TxtProductID.Text Then
                            GridOffer.Columns("ProductName").Text = CN.Execute("Select ProductName from products where productid = '" & GridOffer.Columns("ProductOfferID").Text & "'").Fields(0)
                            GridOffer.Columns("Qty").Value = QtyOffer
                            RsProductOffer!QtyOffer = QtyOffer
                            GridOffer.MoveLast
                            If TxtCode.Enabled = True Then TxtCode.SetFocus
                            GridOffer.Redraw = True
                            Exit Sub
                        End If
                            GridOffer.MoveNext
                        Next vCounter
                        GridOffer.MoveLast
                        Exit Sub
                    End If
                End If
                GridOffer.Redraw = False
                If RsProductOffer.RecordCount = 0 Then
                        RsProductOffer.AddNew
                        GridOffer.Columns("ProductID").Text = TxtProductID.Text
                        GridOffer.Columns("ProductOfferID").Text = !ProductOfferID
                        RsProductOffer!ProductID = TxtProductID.Text
                        RsProductOffer!ProductOfferID = !ProductOfferID
                End If
                    GridOffer.Columns("ProductName").Text = CN.Execute("Select ProductName from products where productid = '" & GridOffer.Columns("ProductOfferID").Text & "'").Fields(0)
                    GridOffer.Columns("Qty").Value = QtyOffer
                    RsProductOffer!QtyOffer = QtyOffer
                    GridOffer.MoveLast
                If Trim(GridOffer.Columns("ProductID").Text) <> "" Then
                    GridOffer.AllowAddNew = True
                    GridOffer.AddNew
                    GridOffer.Columns("ProductID").Text = " "
                    GridOffer.AllowAddNew = False
                End If
            GridOffer.Redraw = True
            
            ''''''''''*******QtyOffer Less than total qty than Delete Row in GridOffer
        Else
            RsProductOffer.Filter = "ProductID='" & TxtProductID.Text & "'"
            If RsProductOffer.RecordCount > 0 Then
                GridOffer.MoveFirst
                For vCounter = 1 To GridOffer.Rows
                If GridOffer.Columns("ProductID").Text = TxtProductID.Text Then
                    If RsProductOffer.RecordCount > 0 Then RsProductOffer.Delete
                    GridOffer.SelBookmarks.RemoveAll
                    GridOffer.SelBookmarks.Add GridOffer.Bookmark
                    GridOffer.DeleteSelected
                    GridOffer.Refresh
                    GridOffer.MoveLast
                    Exit Sub
                End If
                    GridOffer.MoveNext
                    Next vCounter
                End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
        
    End If
    End With
   Exit Sub
ErrorHandler:
   GridOffer.Redraw = True
   Call ShowErrorMessage
End Sub


Private Sub PopulateDataToGridOffer()
    If RsProductOffer.State = adStateOpen Then RsProductOffer.Close
    RsProductOffer.Open "Select * from SaleReturnOffer where ReturnID =" & Val(TxtReturnID.Text) & " And ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
    If RsProductOffer.RecordCount > 0 Then
    GridOffer.Visible = True
    sSql = "select p.productname, D.* from SaleReturnOffer D Inner join products p on p.productid = D.productOfferid where ReturnID =" & Val(TxtReturnID.Text) & " And ReturnDate = '" & DtpReturnDate.DateValue & "'"
      With CN.Execute(sSql)
         GridOffer.Redraw = False
         GridOffer.MoveFirst
         GridOffer.RemoveAll
         GridOffer.AllowAddNew = True
         While Not .EOF
            GridOffer.AddNew
            GridOffer.Columns("ProductID").Text = !ProductID
            GridOffer.Columns("ProductOfferID").Text = !ProductOfferID
            GridOffer.Columns("ProductName").Text = !ProductName
            GridOffer.Columns("Qty").Value = !QtyOffer
            .MoveNext
         Wend
      End With
      GridOffer.AddNew
      GridOffer.Columns("ProductID").Text = " "
      GridOffer.AllowAddNew = False
      GridOffer.Redraw = True
    Else
      GridOffer.CancelUpdate
      GridOffer.RemoveAll
      GridOffer.AddNew
      GridOffer.Columns("ProductID").Text = " "
      GridOffer.Update
      GridOffer.Visible = False
    End If
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("Code").Text
      TxtProductName.Text = .Columns("ProductName").Text
      If Trim(.Columns("PackName").Text) = "" Then
         CmbPackName.ListIndex = 0
      Else
         CmbPackName.Text = .Columns("PackName").Text
      End If
      TxtMultiplier.Text = .Columns("Pack").Text
      TxtQtyLoose.Text = .Columns("QtyLoose").Text
      TxtQtyPack.Text = .Columns("QtyPack").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtBonus.Text = .Columns("Bonus").Text
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtOffer.Text = .Columns("Offer").Value
      TxtSaleTaxPer.Text = .Columns("SaleTaxPer").Value
      TxtSaleTaxVal.Text = .Columns("SaleTaxVal").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      TxtAmount.Text = .Columns("Amount").Value
      TxtCost.Text = .Columns("Cost").Value
      ChkIsProduct.Value = Abs(.Columns("isProduct").Value)
      If Val(TxtMultiplier.Text) = 0 Then
         vUnitPrice = IIf(.Columns("Price").Text = "", 0, .Columns("Price").Text)
      Else
         vUnitPrice = .Columns("Price").Text / Val(TxtMultiplier.Text)
      End If
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetSaleReturn()
   On Error GoTo ErrorHandler
   sSql = "select h.*, SaleManName, OrganizationName, p.partyname, P.Address, P.City, c.AccountName,  StoreName, EmpName, MemberName FROM SaleReturnHeader h join parties p on h.CustomerID = p.partyid left outer join Organizations o on o.OrganizationID = h.OrganizationID left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join Members m on m.MemberID = h.MemberID inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID left outer join SalesMan SM on SM.SaleManID = h.SaleManID  where isReplace=0 and h.ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
   With CN.Execute(sSql)
      If Not .BOF Then
          DtpReturnDate.DateValue = !ReturnDate
          TxtCustomerID.Text = !CustomerID
          TxtCustomerName.Text = !PartyName
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          TxtSalemanID.Text = IIf(IsNull(!SalemanID), "", !SalemanID)
          TxtSalemanName.Text = IIf(IsNull(!SalemanName), "", !SalemanName)
          TxtBillNo.Text = IIf(IsNull(!BillNo), "", !BillNo)
          TxtBiltyNo.Text = IIf(IsNull(!BiltyNo), "", !BiltyNo)
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtTotalAmount.Text = !TotalAmount
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtPaidAmount.Text = IIf(IsNull(!PAIDAMOUNT), "", !PAIDAMOUNT)
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

Private Sub GridSerial_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub GridSerial_DblClick()
   If Grid.Columns("Code").Text <> " " And GridSerial.Columns("Serial").Text = " " Then
        TxtSerial.Enabled = True
        TxtSerial.SetFocus
    Else
        TxtSerial.Enabled = False
    End If
End Sub

Private Sub GridSerial_GotFocus()
   If Grid.Columns("Code").Text <> " " And GridSerial.Columns("Serial").Text = " " Then
        TxtSerial.Enabled = True
    Else
        TxtSerial.Enabled = False
    End If
End Sub

Private Sub GridSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub GridSerial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(GridSerial.Columns("Serial").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub GridSerial_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    GetDataBackFromGridSerialToTexBoxes
End Sub

Private Sub TxtBillDisc_Change()
   If ActiveControl.Name <> TxtBillDisc.Name Then Exit Sub
   TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / Val(TxtTotalAmount.Text), 2)
   Call SubCalculateFooter
End Sub

Private Sub TxtBillDiscPer_Change()
   If ActiveControl.Name <> TxtBillDiscPer.Name Then Exit Sub
   TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Text) * Val(TxtBillDiscPer.Text) / 100))
   Call SubCalculateFooter
End Sub

Private Sub TxtDiscPC_Change()
   If ActiveControl.Name <> TxtDiscPC.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 2)
   If Val(TxtDiscPer.Text) = 0 Then TxtDiscPer.Text = ""
   Call SubCalculateBody
End Sub

Private Sub TxtDiscPer_Change()
   If ActiveControl.Name <> TxtDiscPer.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   TxtDiscPC.Text = Round((vUnitPrice * Val(TxtDiscPer.Text) / 100), 2)
   If Val(TxtDiscPC.Text) = 0 Then TxtDiscPC.Text = ""
   Call SubCalculateBody
End Sub

Private Sub TxtDiscVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDiscVal.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   If (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) = 0 Then Exit Sub
   TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 3)
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 2)
   TxtAmount.Text = Round((Val(vUnitPrice) - Val(TxtDiscPC.Text)) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 2)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDiscVal_LostFocus()
   Select Case ActiveControl.Name
   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtBonus.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtOffer.Name, TxtSaleTaxPer.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtMultiplier_Change()
   If ActiveControl.Name <> TxtMultiplier.Name Then Exit Sub
   If Val(TxtMultiplier.Text) <> 0 Then
      TxtPrice.Text = Round(vUnitPrice * Val(TxtMultiplier.Text), 3)
   Else
      TxtPrice.Text = Round(vUnitPrice, 3)
   End If
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub TxtMultiplier_Validate(Cancel As Boolean)
If ActiveControl.Name <> TxtQtyPack.Name Then Exit Sub
End Sub

Private Sub TxtOffer_Change()
'    If ActiveControl.Name <> TxtOffer.Name Then Exit Sub
    Call SubCalculateBody
End Sub

Private Sub TxtOtherCharges_Change()
   Call SubCalculateFooter
End Sub

Private Sub TxtPrice_Change()
   If ActiveControl.Name <> TxtPrice.Name Then Exit Sub
   If Val(TxtMultiplier.Text) <> 0 Then
      vUnitPrice = Val(TxtPrice.Text) / Val(TxtMultiplier.Text)
   Else
      vUnitPrice = Val(TxtPrice.Text)
   End If
   Call SubCalculateBody
End Sub

Private Sub TxtQtyLoose_Change()
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub TxtQtyLoose_Validate(Cancel As Boolean)
If ActiveControl.Name <> TxtQtyLoose.Name Then Exit Sub
End Sub

Private Sub TxtQtyPack_Change()
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub TxtSaleTaxPer_Change()
If ActiveControl.Name <> TxtSaleTaxPer.Name Then Exit Sub
   Call SubCalculateBody
End Sub

Private Sub TxtSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridSerial.SetFocus
End Sub

Private Sub TxtSerial_LostFocus()
    GetDataFromTexBoxesToGridSerial
End Sub

Private Sub GetDataFromTexBoxesToGridSerial()
   On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   If Trim(TxtSerial.Text) = "" Then
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
'      TxtSerial.SetFocus
      Exit Sub
   End If
   RsBodySerial.Filter = "ProductID ='" & Grid.Columns("ProductID").Text & "' And Serial='" & TxtSerial.Text & "'"
   If TxtSerial.Enabled Then
      If RsBodySerial.RecordCount = 0 Then
         RsBodySerial.AddNew
         GridSerial.Columns("ProductID").Text = TxtCode.Text
         GridSerial.Columns("Serial").Text = TxtSerial.Text
         RsBodySerial!ProductID = TxtCode.Text
         RsBodySerial!Serial = TxtSerial.Text
         TxtSerial.Text = ""
      Else
         GridSerial.Redraw = False
         GridSerial.MoveFirst
            For vrowcounter = 1 To GridSerial.Rows
               If GridSerial.Columns("Serial").Text = TxtSerial.Text Then
                  MsgBox "The Product cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  GridSerial.MoveLast
                  TxtSerial.SetFocus
                  GridSerial.Redraw = True
                  Exit Sub
               End If
               GridSerial.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         
         GridSerial.MoveLast
         TxtSerial.SetFocus
         Exit Sub
      End If
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
   TxtSerial.SetFocus
   GridSerial.Redraw = True
   Exit Sub
ErrorHandler:
   GridSerial.Redraw = True
   Call ShowErrorMessage
End Sub
Private Sub TxtTotalAmount_Change()
   TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Text) * Val(TxtBillDiscPer.Text) / 100))
   Call SubCalculateFooter
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
      TxtPrice.Text = ""
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
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = FunSelectProduct(ssValidate, False)
   If vTemp = False Then   '   vTemp = FunSelectProduct(ssButton, False)
      Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub TxtStoreID_Change()
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then
      TxtStoreName.Text = ""
   End If
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
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub FindRebate()
   On Error GoTo ErrorHandler
    With CN.Execute("Select * from ProductOffers where Rebate <> 0 and ProductID = '" & TxtProductID.Text & "'")
        If .RecordCount > 0 Then
            Rebate = Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)
            Rebate = Rebate \ !Qty
            Rebate = Rebate * !Rebate
            TxtOffer.Text = Rebate
        End If
    End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub UserActivities()
     If vIsNewRecord = False Then
    With CN.Execute("Select  * from SaleReturnHeader where ReturnID =" & TxtReturnID.Text & " And ReturnDate = '" & DtpReturnDate.DateValue & "'")
        
        If TxtStoreID.Text <> !StoreID Then
            CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.Rows - 1
        With CN.Execute("Select * from SaleReturnBody Where ReturnID = " & TxtReturnID.Text & " and ReturnDate ='" & DtpReturnDate.DateValue & "' and Productid ='" & Grid.Columns("Productid").Text & "'")
        
             If .EOF = True Then
                sSql = "Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")"
                CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("QtyLoose").Text <> !Qty Or Grid.Columns("Price").Text <> !Price Or Grid.Columns("discper").Text <> !DiscPer Then
                   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
    
   Else
    CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

