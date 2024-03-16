VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmRecoveryInvoiceWise 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmRecoveryInvoiceWise.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmRecoveryInvoiceWise.frx":0ECA
      Left            =   5123
      List            =   "FrmRecoveryInvoiceWise.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Tag             =   "1"
      Top             =   8370
      Width           =   3276
   End
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   8408
      TabIndex        =   45
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   8370
      Width           =   1170
   End
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9623
      TabIndex        =   44
      Top             =   8415
      Width           =   1290
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
      Left            =   13800
      TabIndex        =   31
      Top             =   1320
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
         TabIndex        =   32
         Tag             =   "NC"
         Text            =   "FrmRecoveryInvoiceWise.frx":0ECE
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
         TabIndex        =   33
         Top             =   90
         Width           =   135
      End
   End
   Begin SITextBox.Txt TxtRecoveryID 
      Height          =   315
      Left            =   1845
      TabIndex        =   0
      Top             =   2198
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
      Left            =   9027
      TabIndex        =   11
      Top             =   8738
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0FC1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7704
      TabIndex        =   8
      Top             =   8738
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0FDD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5058
      TabIndex        =   10
      Top             =   8738
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0FF9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10350
      TabIndex        =   13
      Top             =   8738
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":1015
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6381
      TabIndex        =   9
      Top             =   8738
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":1031
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   2438
      TabIndex        =   5
      Top             =   3255
      Width           =   945
      _ExtentX        =   1667
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
   Begin JeweledBut.JeweledButton BtnSale 
      Height          =   330
      Left            =   3383
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3240
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":104D
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   3743
      TabIndex        =   15
      Top             =   3255
      Width           =   3810
      _ExtentX        =   6720
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4650
      Left            =   1133
      TabIndex        =   19
      Top             =   3570
      Width           =   13095
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   9
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
      stylesets(0).Picture=   "FrmRecoveryInvoiceWise.frx":1069
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
      Columns.Count   =   9
      Columns(0).Width=   2275
      Columns(0).Caption=   "Bill Date"
      Columns(0).Name =   "BillDate"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).NumberFormat=   "dd/MM/yyyy"
      Columns(0).FieldLen=   256
      Columns(1).Width=   2328
      Columns(1).Caption=   "Bill ID"
      Columns(1).Name =   "BillID"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Customer ID"
      Columns(2).Name =   "CustomerID"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   6694
      Columns(3).Caption=   "Customer Name"
      Columns(3).Name =   "CustomerName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2275
      Columns(4).Caption=   "Sale Value"
      Columns(4).Name =   "SaleValue"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2302
      Columns(5).Caption=   "Received"
      Columns(5).Name =   "Received"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2302
      Columns(6).Caption=   "Amount"
      Columns(6).Name =   "Amount"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1826
      Columns(7).Caption=   "Discount"
      Columns(7).Name =   "Discount"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   2514
      Columns(8).Caption=   "Final Debit"
      Columns(8).Name =   "FinalDebit"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   23098
      _ExtentY        =   8202
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpRecoveryDate 
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   2198
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
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   10148
      TabIndex        =   6
      Top             =   3255
      Width           =   1305
      _ExtentX        =   2302
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtDiscount 
      Height          =   315
      Left            =   11453
      TabIndex        =   7
      Top             =   3255
      Width           =   1035
      _ExtentX        =   1826
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
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtFinalDebit 
      Height          =   315
      Left            =   12488
      TabIndex        =   16
      Top             =   3255
      Width           =   1440
      _ExtentX        =   2540
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3735
      TabIndex        =   12
      Top             =   8738
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":1085
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpBillDate 
      Height          =   315
      Left            =   1133
      TabIndex        =   4
      Top             =   3255
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
   Begin SITextBox.Txt TxtSaleValue 
      Height          =   315
      Left            =   7553
      TabIndex        =   29
      Top             =   3255
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtReceived 
      Height          =   315
      Left            =   8843
      TabIndex        =   30
      Top             =   3255
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin JeweledBut.JeweledButton BtnEmployee 
      Height          =   330
      Left            =   5880
      TabIndex        =   35
      Top             =   2213
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":10A1
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmployeeName 
      Height          =   315
      Left            =   6240
      TabIndex        =   36
      Top             =   2213
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   4695
      TabIndex        =   2
      Top             =   2213
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   8730
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   2213
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
      Left            =   9795
      TabIndex        =   37
      Tag             =   "NC"
      Top             =   2213
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
      Left            =   9435
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2213
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":10BD
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   315
      Left            =   1845
      TabIndex        =   48
      Top             =   1230
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerID"
      Height          =   195
      Left            =   1890
      TabIndex        =   49
      Top             =   1035
      Width           =   825
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
      Left            =   4448
      TabIndex        =   47
      Top             =   8370
      Width           =   570
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No."
      Height          =   225
      Left            =   6278
      TabIndex        =   43
      Top             =   8475
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      Height          =   255
      Left            =   4695
      TabIndex        =   42
      Top             =   2018
      Width           =   1215
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      Height          =   375
      Left            =   6225
      TabIndex        =   41
      Top             =   2018
      Width           =   1215
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   9915
      TabIndex        =   40
      Top             =   1973
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   8730
      TabIndex        =   39
      Top             =   1973
      Width           =   1095
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
      Left            =   10980
      TabIndex        =   34
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   1133
      TabIndex        =   28
      Top             =   3060
      Width           =   585
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recovery (Invoice Wise)"
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
      TabIndex        =   27
      Top             =   270
      Width           =   4245
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   11453
      TabIndex        =   26
      Top             =   3060
      Width           =   630
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Debit"
      Height          =   195
      Left            =   12488
      TabIndex        =   25
      Top             =   3060
      Width           =   750
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      Height          =   195
      Left            =   8843
      TabIndex        =   24
      Top             =   3060
      Width           =   690
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   10148
      TabIndex        =   23
      Top             =   3060
      Width           =   540
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Value"
      Height          =   195
      Left            =   7553
      TabIndex        =   22
      Top             =   3060
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill ID"
      Height          =   195
      Left            =   2438
      TabIndex        =   21
      Top             =   3060
      Width           =   405
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   195
      Left            =   3743
      TabIndex        =   20
      Top             =   3060
      Width           =   1125
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recovery Date"
      Height          =   195
      Left            =   3135
      TabIndex        =   18
      Top             =   2003
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Recovery ID"
      Height          =   195
      Left            =   1845
      TabIndex        =   17
      Top             =   2003
      Width           =   900
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmRecoveryInvoiceWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter, vGridRows As Integer
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim ssql As String
Dim i As Integer
Dim vStrSQL, vRandomID As String
Dim vMaxBinID As Integer
Dim vPrinter() As String

'----------------------------------

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
      If DtpBillDate.Enabled Then DtpBillDate.SetFocus Else TxtOrganizationID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Function FunSelectEmployee(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchEmployee.Show vbModal, Me
        If SchEmployee.ParaOutEmployeeID = "" Then FunSelectEmployee = False: Exit Function
        TxtEmployeeID.Text = SchEmployee.ParaOutEmployeeID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Employees where isLockEmployee and EmpID=" & Val(TxtEmployeeID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtEmployeeName.Text = !empname
          FunSelectEmployee = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectEmployee = False
          .Close
          TxtEmployeeID.Text = ""
          TxtEmployeeName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnEmployee_Click()
   If FunSelectEmployee(ssButton, False) = True Then
      If TxtOrganizationID.Enabled And TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus Else If DtpBillDate.Enabled Then DtpBillDate.SetFocus
   Else
      TxtEmployeeID.SetFocus
   End If
End Sub

Private Sub TxtEmployeeID_Change()
   If TxtEmployeeID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtEmployeeID.Name Then Exit Sub
   If TxtEmployeeName.Text <> "" Then
      TxtEmployeeName.Text = ""
   End If
End Sub

Private Sub TxtEmployeeID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtEmployeeID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtEmployeeName.Text <> "" Then Exit Sub
   If Trim(TxtEmployeeID.Text) = "" Then Exit Sub
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

Private Sub CalculateBody()
   TxtFinalDebit.Text = Val(TxtSaleValue.Text) - Val(TxtReceived.Text) - Val(TxtAmount.Text) - Val(TxtDiscount.Text)
End Sub

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
    '''''''''''''''''' ActivityLogBin For Clear Action
'      Call DeleteTempActivityLogBin(vRandomID)
      vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("BillID").Text) <> "" Then
           ssql = "Select RecoveryID From RecoveryInvoice where RecoveryID=" & Val(TxtRecoveryID.Text) & " and BillID = " & Grid.Columns("BillID").Text & " and BillDate =  '" & Grid.Columns("BillDate").Text & "'"
            With cn.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Cleared Code-" & Grid.Columns("BillID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("CustomerName").Text))
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eClearSavedRecord, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, vGridRows & " Recovery Invoice/s Cleared")
      Grid.Redraw = True
  ''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Recovery Invoice Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Recovery Invoice Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
'      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
    ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniRecoveryInvoiceWise", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   cn.BeginTrans
   
   Call BinData
   Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eDelete, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, Grid.rows - 1 & " Recovery Invoice/s Deleted ")
   
'   vMaxBinID = FunGetMaxBinID
   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_RecoveryHeader Select " & vMaxBinID & ",'" & Date & "',* from RecoveryHeader Where RecoveryID = " & TxtRecoveryID.Text & " And RecoveryDate ='" & DtpRecoveryDate.DateValue & "'")
    '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_RecoveryInvoice Select " & vMaxBinID & ",'" & Date & "', * from RecoveryInvoice Where RecoveryID = " & TxtRecoveryID.Text)
    
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Recovery Invoice Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Call ActivityLog("Recovery Invoice Voice", eDelete, Val(TxtRecoveryID.Text))
   Grid.Redraw = False
   Grid.RemoveAll
   cn.Execute "Delete from RecoveryInvoice where RecoveryID = " & Val(TxtRecoveryID.Text)
   Grid.Redraw = True
   cn.Execute "Delete from RecoveryHeader where RecoveryID = " & Val(TxtRecoveryID.Text)
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchRecoveryInvoice.Show vbModal
   If SchRecoveryInvoice.ParaOutRecoveryID <> 0 Then
      TxtRecoveryID.Text = SchRecoveryInvoice.ParaOutRecoveryID
      cn.Execute ("Insert Into UserActivities values ('Recovery Invoice Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetRecovery
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   vStrSQL = " Select  H.RecoveryID, H.RecoveryDate, Pty.PartyName, B.BillID, Sv.SaleValue , SV.Paid," & vbCrLf _
      + " Isnull(B.Amount,0) Amount, Isnull(B.Discount,0) Discount" & vbCrLf _
      + " from RecoveryHeader H " & vbCrLf _
      + " Inner join RecoveryInvoice B on H.RecoveryID = B.RecoveryID " & vbCrLf _
      + " Inner join SaleHeader SH on SH.BillID = H.RecoveryID" & vbCrLf _
      + " Left Join Parties pty on pty.partyID = SH.CustomerID" & vbCrLf _
      + " inner Join (select h.BillID, totalamount - isnull(billdisc,0) as SaleValue, (totalamount - isnull(billdisc,0) - isnull(PaidAmount,0) - isnull(PrePaid,0) ) as  Paid " & vbCrLf _
      + " from SaleHeader h left outer join " & vbCrLf _
      + " (select b.BillID, sum(b.amount-isnull(b.discount,0)) PrePaid from RecoveryInvoice b  inner join RecoveryHeader H on b.RecoveryID = H.RecoveryID where RecoveryDate < '" & DtpRecoveryDate.DateValue & "' group by BillID) i on h.BillID = i.BillID ) Sv" & vbCrLf _
      + " on b.BillID = Sv.BillID" & vbCrLf _
      + " where h.RecoveryID=" & Val(TxtRecoveryID.Text)


    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly

    'Set RptReportViewer.Report = New CrpRecoveryInvoice
    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
'    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
   vPrinter = Split(CmbPrinters.Text, ",")
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
    
    Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String
    
    vStrComp = "Select CompanyName,Address,City,PhoneNo,email from Company"
    With cn.Execute(vStrComp)
      If .RecordCount > 0 Then
         vCompanyName = !CompanyName
         vAddress = !Address & IIf(IsNull(!City), "", ", " & !City)
         vPhone = IIf(IsNull(!PhoneNo), "", "Phone # " & !PhoneNo)
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue vAddress
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue vPhone
      End If
   End With
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue cn.Execute("Select Name from Manufacturer").Fields(0).Value
   Dim vDevice As String, vDriver As String, vPort As String
   vStrSQL = "Select * from Registry"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         vDevice = IIf(IsNull(!DeviceName), "Abc", !DeviceName)
         vDriver = IIf(IsNull(!DriverName), "Xyz", !DriverName)
         vPort = IIf(IsNull(!Port), "LPT1", !Port)
         RptReportViewer.Report.SelectPrinter vDriver, vDevice, vPort
      End If
   End With
   'RptReportViewer.Report.PrintOut False
   cn.Execute ("Insert Into UserActivities values ('Recovery Invoice Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
   If ChkIsPreview.Value = 1 Then
      RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSale_Click()
   If FunSelectSale(ssButton, False) = True Then
      TxtAmount.SetFocus
   Else
      TxtBillID.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniRecoveryInvoiceWise", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
    MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
    Exit Sub
  End If
'  Header Validation
'   If Trim(TxtBillID.Text) = "" Then
'      MsgBox "Enter PurchaseMan ID.", vbExclamation, Me.Caption
'      TxtBillID.SetFocus
'      Exit Sub
'   End If
   If TxtRecoveryID.Enabled Then
      If cn.Execute("Select * from RecoveryHeader where RecoveryID = " & Val(TxtRecoveryID.Text)).RecordCount > 0 Or Val(TxtRecoveryID.Text) = 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtRecoveryID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
   
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where ToUserNo = " & vUser & " and Entrydate <='" & Date & "'"
    With cn.Execute(vStrSQL)
        If .Fields(0).Value >= DtpRecoveryDate.DateValue Then
            MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
    
  RsBody.Filter = 0
  If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one entry for Recovery", vbExclamation, "Alert"
      If TxtRecoveryID.Visible And TxtRecoveryID.Enabled Then TxtRecoveryID.SetFocus
      Exit Sub
  End If
  'Body Validation
  ' validation has been performed when a row is added to the grid

  'Saving record
  
  ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'Recovery Invoice Wise' and LocalComputerName = '" & LocalComputerName & "'"
   If cn.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview ) Values ('" & LocalComputerName & "', 'Recovery Invoice Wise','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & " Where FormType = 'Recovery Invoice Wise' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   cn.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   
   cn.BeginTrans
   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eEdit, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, "")
'   Call UserActivities
   
   ssql = "select * from RecoveryHeader where RecoveryID =" & Val(TxtRecoveryID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, cn, adOpenDynamic, adLockPessimistic
      If .BOF Then
         .AddNew
         !RecoveryID = Val(TxtRecoveryID.Text)
         !UserNo = vUser
      End If
      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !RecoveryDate = DtpRecoveryDate.DateValue
'      !UserNo = vUser
      .Update
      .Close
   End With
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !RecoveryID = Val(TxtRecoveryID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eAdd, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, Grid.rows - 1 & " New New Recovery Invoice/s Added ")
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
  ' If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = "Grid" Then
         Grid_DblClick
      Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtBillID.Enabled Then TxtBillID.SetFocus: Call SubClearDetailArea
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
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then If TxtOrganizationID.Enabled And TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus Else If DtpBillDate.Enabled Then DtpBillDate.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then If DtpBillDate.Enabled Then DtpBillDate.SetFocus Else TxtOrganizationID.SetFocus
         Case TxtBillID.Name: If FunSelectSale(ssFunctionKey, False) = True Then TxtAmount.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtRecoveryID.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtRecoveryID.Name Then
         KeyCode = 0
      End If
   End If
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
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Recovery (Invoice Wise)"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   
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
   ssql = "select * from FormDefaultSetting Where FormType = 'Recovery Invoice Wise' and LocalComputerName = '" & LocalComputerName & "'"
   With cn.Execute(ssql)
     If .RecordCount > 0 Then
        cmbPrintType.Text = !Size
        ChkIsPreview.Value = Abs(!IsPreview)
        If Not IsNull(!DeviceName) Then
            CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
        Else
            CmbPrinters.ListIndex = 0
        End If
     End If
     .Close
   End With
   ''''''''''''''''''''''''''''''''''''''''''''''
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   HelpLocation Me
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select isnull(max(RecoveryID),0)+1 from RecoveryHeader").Fields(0)
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
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("BillID").Text = " "
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
    Set FrmRecoveryInvoiceWise = Nothing
   End If
    '''''''''''''''''' ActivityLogBin For Close Action
'      Call DeleteTempActivityLogBin(vRandomID)
      If Grid.rows > 1 And Cancel = 0 Then
         vGridRows = 0
         Grid.Redraw = False
         Grid.MoveFirst
         For vCounter = 2 To Grid.rows
            vGridRows = vGridRows + 1
            If Trim(Grid.Columns("BillID").Text) <> "" Then
               ssql = "Select RecoveryID From RecoveryInvoice where RecoveryID=" & Val(TxtRecoveryID.Text) & " and BillID = " & Grid.Columns("BillID").Text & " and BillDate =  '" & Grid.Columns("BillDate").Text & "'"
               With cn.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Closed Code-" & Grid.Columns("BillID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("CustomerName").Text))
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eCloseSavedRecord, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, vGridRows & " Recovery Invoice/s Closed")
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
   'TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("FinalDebit").Value
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
   TxtBillID.Enabled = False
   BtnSale.Enabled = False
   DtpBillDate.Enabled = False
   'TxtRecoveryID.BackColor = TxtCustomerName.BackColor
   'TxtRecoveryID.TabStop = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("BillID").Text) = "" Then
      TxtBillID.Text = ""
      TxtBillID.Enabled = True
      BtnSale.Enabled = True
      DtpBillDate.Enabled = True
      TxtBillID.SetFocus
   Else
      TxtBillID.Enabled = False
      BtnSale.Enabled = False
      DtpBillDate.Enabled = False
      TxtAmount.SetFocus
      If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(Grid.Columns("BillID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub GetRecovery()
   On Error GoTo ErrorHandler
   ssql = " select h.*, OrganizationName, EmpName FROM RecoveryHeader h inner join RecoveryInvoice b on h.recoveryid = b.recoveryid left outer join Organizations o on o.OrganizationID = h.OrganizationID left outer join Employees sm on sm.Empid = h.Empid where h.RecoveryID=" & Val(TxtRecoveryID.Text)
   With cn.Execute(ssql)
      If Not .BOF Then
          DtpRecoveryDate.DateValue = !RecoveryDate
          TxtEmployeeID.Text = IIf(IsNull(!EmpID) = True, "", !EmpID)
          TxtEmployeeName.Text = IIf(IsNull(!empname) = True, "", !empname)
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
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

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("BillID").Text) = "" And Trim(Grid.Columns("BillDate").Text) = "" Then Exit Sub
   ssql = "Select RecoveryID From RecoveryInvoice where RecoveryID=" & Val(TxtRecoveryID.Text) & " and BillID = " & Grid.Columns("BillID").Text & " and BillDate =  '" & Grid.Columns("BillDate").Text & "'"
   With cn.Execute(ssql)
      If .EOF Then
         Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Removed Code-" & Grid.Columns("BillID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("CustomerName").Text))
      Else
         Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eRemoveRow, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, "Removed Code-" & Grid.Columns("BillID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("CustomerName").Text))
         Call ActivityLogBin(vRandomID, eFrmRecoveryInvoiceWise, eAddTempRecord, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, "Pending Remove Code-" & Grid.Columns("BillID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("CustomerName").Text))
      End If
   End With
   RsBody.Filter = "BillID = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'"
   If RsBody.RecordCount > 0 Then RsBody.Delete
   cn.Execute ("Insert Into UserActivities values ('Recovery Invoice Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Removed PurcahseID-" & Grid.Columns("BillID").Text & " BillDate-" & Grid.Columns("BillDate").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   Grid.Refresh
   RsBody.Filter = 0
   GetDataBackFromGridToTexBoxes
Exit Sub
ErrorHandler:
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
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      BtnSale.Enabled = True
      TxtRecoveryID.Enabled = True
      TxtRecoveryID.Text = FunGetMaxID()
      If DtpRecoveryDate.Enabled And DtpRecoveryDate.Visible Then DtpRecoveryDate.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      TxtRecoveryID.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      BtnSale.Enabled = True
      TxtBillID.Enabled = True
      DtpBillDate.Enabled = True
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

Private Sub PopulateDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from RecoveryInvoice where RecoveryID =" & Val(TxtRecoveryID.Text), cn, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "  Select B.BillID, b.BillDate, B.CustomerID, ca.AccountName + isnull(' (' + p.Address + ')','') as CustomerName, Sv.SaleValue , SV.Received," & vbCrLf _
      + "  Isnull(B.Amount,0) Amount, Isnull(B.Discount,0) Discount " & vbCrLf _
      + "  from RecoveryHeader H Inner join RecoveryInvoice B on H.RecoveryID = B.RecoveryID " & vbCrLf _
      + "  Inner join SaleHeader SH on SH.BillID = b.BillID and sh.BillDate = b.BillDate" & vbCrLf _
      + "  Left Join ChartofAccounts ca on ca.AccountNo = SH.CustomerID Left Outer join Parties p on ca.AccountNo = p.PartyID" & vbCrLf _
      + "  inner Join (" & vbCrLf _
      + "  select h.BillID, h.BillDate, TotalAmount - isnull(BillDisc,0) + isnull(OtherCharges,0) as SaleValue, (isnull(CashReceived,0) + isnull(PrePaid,0) ) as  Received " & vbCrLf _
      + "  from SaleHeader h left outer join " & vbCrLf _
      + "  (select BillID, BillDate, sum(b.amount + isnull(b.discount,0)) PrePaid from RecoveryInvoice b inner join RecoveryHeader H on b.RecoveryID = H.RecoveryID where RecoveryDate < '" & DtpRecoveryDate.DateValue & "' group by BillID, BillDate) i on h.BillID = i.BillID and h.BillDate = i.BillDate) Sv" & vbCrLf _
      + "  on b.BillID = Sv.BillID and b.BillDate = sv.BillDate" & vbCrLf _
      + " where h.RecoveryID = " & Val(TxtRecoveryID.Text)

      'sSql = "select b.* from RecoveryInvoice b where RecoveryID =" & Val(TxtRecoveryID.Text)
      With cn.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            Grid.AddNew
            Grid.Columns("BillID").Value = !BillID
            Grid.Columns("BillDate").Value = !BillDate
            Grid.Columns("CustomerID").Value = !CustomerID
            Grid.Columns("CustomerName").Value = IIf(IsNull(!CustomerName), "", !CustomerName)
            Grid.Columns("SaleValue").Value = !SaleValue
            Grid.Columns("Received").Value = IIf(IsNull(!Received), "", !Received)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("Discount").Value = IIf(IsNull(!Discount), "", !Discount)
            Grid.Columns("FinalDebit").Value = Val(Grid.Columns("SaleValue").Value) - Val(Grid.Columns("Received").Value) - Val(Grid.Columns("Amount").Value) - Val(Grid.Columns("Discount").Value)
            'TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(!FinalDebit)
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("BillID").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtBillID.Text = .Columns("BillID").Text
      DtpBillDate.DateValue = .Columns("BillDate").Value
      TxtCustomerID.Text = .Columns("CustomerID").Text
      TxtCustomerName.Text = .Columns("CustomerName").Text
      TxtSaleValue.Text = .Columns("SaleValue").Value
      TxtReceived.Text = .Columns("Received").Value
      TxtAmount.Text = .Columns("Amount").Value
      TxtDiscount.Text = .Columns("Discount").Value
      TxtFinalDebit.Text = .Columns("FinalDebit").Value
   End With
   If Grid.rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtBillID.Enabled = True
   BtnSale.Enabled = True
   TxtBillID.Text = ""
   'DtpBillDate.DateValue
   TxtCustomerID.Text = ""
   TxtCustomerName.Text = ""
   TxtSaleValue.Text = ""
   TxtAmount.Text = ""
   TxtReceived.Text = ""
   TxtDiscount.Text = ""
   TxtFinalDebit.Text = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub TxtAmount_Change()
   Call CalculateBody
End Sub

Private Sub TxtDiscount_Change()
   Call CalculateBody
End Sub

Private Sub txtDiscount_LostFocus()
   Select Case ActiveControl.Name
   Case TxtBillID.Name, TxtCustomerName.Name, TxtSaleValue.Name, TxtReceived.Name, TxtAmount.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtBillID_Change()
   If TxtBillID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtBillID.Name Then Exit Sub
   If TxtCustomerName.Text <> "" Then
      TxtCustomerName.Text = ""
      TxtSaleValue.Text = ""
      TxtReceived.Text = ""
   End If
End Sub

Private Sub TxtBillID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtBillID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtBillID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerName.Text <> "" Then Exit Sub
   If Trim(TxtBillID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSale(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSale(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   Dim vrowcounter As Integer
   If Trim(TxtBillID.Text) = "" Then
      TxtBillID.SetFocus
      Exit Sub
   End If
'   If Val(TxtAmount.Text) = 0 Then
'      TxtAmount.SetFocus
'      Exit Sub
'   End If
'   If Val(TxtFinalDebit.Text) < 0 Then
'      TxtAmount.SetFocus
'      Exit Sub
'   End If
On Error GoTo ErrorHandler
      RsBody.Filter = "BillID =" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'"
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("CustomerID").Text = TxtCustomerID.Text
         Grid.Columns("BillID").Text = TxtBillID.Text
         Grid.Columns("BillDate").Value = DtpBillDate.DateValue
         RsBody!BillID = Val(TxtBillID.Text)
         RsBody!BillDate = DtpBillDate.DateValue
         RsBody!CustomerID = TxtCustomerID.Text
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eAddNewRowByEdit, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, "Add New Code-" & TxtBillID.Text & " Amount-" & " " & TxtAmount.Text & " " & Trim(TxtCustomerName.Text))
         Call ActivityLogBin(vRandomID, eFrmRecoveryInvoiceWise, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Pending Add New Code-" & TxtBillID.Text & " Amount-" & TxtAmount.Text & " " & Trim(TxtCustomerName.Text))
       Else
         ssql = "Select RecoveryID From RecoveryInvoice where RecoveryID=" & Val(TxtRecoveryID.Text) & " and BillID = " & Grid.Columns("BillID").Text & " and BillDate =  '" & Grid.Columns("BillDate").Text & "'"
         With cn.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Effected Code-" & Grid.Columns("BillID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("CustomerName").Text))
               Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Updated Code-" & TxtBillID.Text & " Amount-" & TxtAmount.Text & " " & Trim(TxtCustomerName.Text))
            Else
               Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eEdit, TxtRecoveryID.Text, DtpRecoveryDate.Date, "Effected Code-" & Grid.Columns("BillID").Text & " Amount-" & Grid.Columns("Amount").Text & " " & Trim(Grid.Columns("CustomerName").Text))
               Call ActivityLogBin("", eFrmRecoveryInvoiceWise, eEdit, TxtRecoveryID.Text, DtpRecoveryDate.Date, "Updated Code-" & TxtBillID.Text & " Amount-" & TxtAmount.Text & " " & Trim(TxtCustomerName.Text))
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmRecoveryInvoiceWise, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Pending Update Code-" & TxtBillID.Text & " Amount-" & TxtAmount.Text & " " & Trim(TxtCustomerName.Text))
       End If
       
       With Grid
         .Columns("CustomerID").Text = TxtCustomerID.Text
         .Columns("CustomerName").Text = TxtCustomerName.Text
         .Columns("SaleValue").Value = Val(TxtSaleValue.Text)
         .Columns("Received").Value = Val(TxtReceived.Text)
         .Columns("Amount").Value = Val(TxtAmount.Text)
         .Columns("Discount").Value = IIf(Val(TxtDiscount.Text) = 0, 0, Val(TxtDiscount.Text))
         .Columns("FinalDebit").Value = Val(TxtFinalDebit.Text)
         RsBody!CustomerID = TxtCustomerID.Text
         RsBody!Amount = Val(TxtAmount.Text)
         RsBody!Discount = IIf(Val(TxtDiscount.Text) = 0, 0, Val(TxtDiscount.Text))
         
         .MoveLast
         If Trim(.Columns("BillID").Text) <> "" Then
            .AllowAddNew = True
            .AddNew
            .Columns("BillID").Text = " "
            .AllowAddNew = False
         End If
      End With
   Call SubClearDetailArea
   TxtBillID.SetFocus
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Function FunSelectSale(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchUnReceivedSale.Show vbModal, Me
        If SchUnReceivedSale.ParaOutBillDate = "" Then FunSelectSale = False: Exit Function
        TxtBillID.Text = SchUnReceivedSale.ParaOutBillID
        DtpBillDate.DateValue = SchUnReceivedSale.ParaOutBillDate
    End If
    '---------------------------
    vStrSQL = "Select h.BillID, h.BillDate, totalamount - isnull(billdisc,0) + isnull(OtherCharges,0) as SaleValue, " & vbCrLf _
      + " isnull(CashReceived,0) + isnull(i.amount,0) + isnull(discount,0) as ReceivedAmount, h.CustomerID, ca.AccountName + isnull(' (' + p.Address + ')','')  as CustomerName, " & vbCrLf _
      + " case when LastReceivedDate is not null then LastReceivedDate when isnull(CashReceived,0) <> 0 then h.BillDate end as LastReceivedDate," & vbCrLf _
      + " (totalamount - IsNull(billdisc, 0) + isnull(OtherCharges,0)) - (isnull(CashReceived,0) + isnull(i.amount,0) + isnull(discount,0)) bal" & vbCrLf _
      + " from SaleHeader h left outer join " & vbCrLf _
      + " (select BillID, BillDate, max(BillDate) as LastReceivedDate, sum(amount) as amount, sum(discount) as Discount from RecoveryInvoice group by BillID, BillDate)i " & vbCrLf _
      + " on i.BillID = h.BillID and i.BillDate = h.BillDate" & vbCrLf _
      + " inner join (select BillID, BillDate from SaleBody Group By BillID, BillDate)b on h.BillID = b.BillID and h.BillDate = b.BillDate" & vbCrLf _
      + " inner join ChartofAccounts ca on h.CustomerID = ca.AccountNo Left Outer join Parties p on ca.AccountNo = p.PartyID " & vbCrLf _
      + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate = '" & DtpBillDate.DateValue & "' and (totalamount - IsNull(billdisc, 0) + isnull(OtherCharges,0)) - (isnull(CashReceived,0) + isnull(i.amount,0) + isnull(discount,0)) > 0 "

    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerID.Text = !CustomerID
          TxtCustomerName.Text = !CustomerName
          TxtSaleValue.Text = !SaleValue
          TxtReceived.Text = !ReceivedAmount
          FunSelectSale = True
          Call CalculateBody
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSale = False
          .Close
          TxtCustomerName.Text = ""
          TxtSaleValue.Text = ""
          TxtReceived.Text = ""
          Call CalculateBody
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpRecoveryDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_RecoveryHeader ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
     If vIsNewRecord = False Then
    With cn.Execute("Select  * from RecoveryHeader where RecoveryID =" & TxtRecoveryID.Text & " And RecoveryDate = '" & DtpRecoveryDate.DateValue & "'")
        If Val(TxtEmployeeID.Text) <> IIf(IsNull(!EmpID), 0, !EmpID) Then
            cn.Execute ("Insert Into UserActivities values ('Recovery Invoice Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Updated EmpID-" & !EmpID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    Grid.MoveFirst
    
    For i = 1 To Grid.rows - 1
        With cn.Execute("Select * from RecoveryInvoice Where RecoveryID = " & TxtRecoveryID.Text & " And BillDate = '" & Grid.Columns("BillDate").Text & "' and BillID =" & Grid.Columns("BillID").Text)
             If .EOF = True Then
                cn.Execute ("Insert Into UserActivities values ('Recovery Invoice Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Inserted New BillID-" & Grid.Columns("BillID").Text & " BillDate-" & Grid.Columns("BillDate").Text & " SaleValue-" & Grid.Columns("SaleValue").Text & " Paid-" & Grid.Columns("Paid").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("Amount").Text <> !Amount Or Grid.Columns("Discount").Text <> !Discount Then
                   cn.Execute ("Insert Into UserActivities values ('Recovery Invoice Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Updated BillID-" & Grid.Columns("BillID").Text & " BillDate-" & Grid.Columns("BillDate").Text & " SaleValue-" & Grid.Columns("SaleValue").Text & " Paid-" & Grid.Columns("Paid").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
   Else
    cn.Execute ("Insert Into UserActivities values ('Recovery Invoice Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.RecoveryHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmRecoveryInvoiceWise) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmRecoveryInvoiceWise & ", " & vUser & "," & TableHeaderFields(eFrmRecoveryInvoiceWise) & " from RecoveryHeader " & vbCrLf _
             & "Where RecoveryID = " & TxtRecoveryID.Text & " and RecoveryDate = '" & DtpRecoveryDate.DateValue & "'"
      cn.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.RecoveryInvoiceBin (" & TableBodyFields(eFrmRecoveryInvoiceWise) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmRecoveryInvoiceWise) & " from RecoveryInvoice " & vbCrLf _
             & "Where RecoveryID = " & TxtRecoveryID.Text
      cn.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

