VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmOpeningStockInvoice 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmOpeningStockInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   5625
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3278
      Width           =   2070
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9016
      TabIndex        =   19
      Top             =   8978
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
      MICON           =   "FrmOpeningStockInvoice.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7696
      TabIndex        =   16
      Top             =   8978
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
      MICON           =   "FrmOpeningStockInvoice.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5056
      TabIndex        =   18
      Top             =   8978
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
      MICON           =   "FrmOpeningStockInvoice.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10336
      TabIndex        =   20
      Top             =   8978
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
      MICON           =   "FrmOpeningStockInvoice.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6376
      TabIndex        =   17
      Top             =   8978
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
      MICON           =   "FrmOpeningStockInvoice.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   12570
      TabIndex        =   15
      Top             =   7785
      Width           =   1215
      _ExtentX        =   2143
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
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Top             =   2550
      Width           =   8100
      _ExtentX        =   14288
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1050
      TabIndex        =   5
      Top             =   3285
      Width           =   1635
      _ExtentX        =   2884
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
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   9795
      TabIndex        =   10
      Top             =   3285
      Width           =   960
      _ExtentX        =   1693
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
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   2685
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3285
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
      MICON           =   "FrmOpeningStockInvoice.frx":0F56
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   3045
      TabIndex        =   26
      Top             =   3285
      Width           =   2580
      _ExtentX        =   4551
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
      Height          =   3840
      Left            =   1050
      TabIndex        =   27
      Top             =   3600
      Width           =   13215
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   13
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
      stylesets(0).Picture=   "FrmOpeningStockInvoice.frx":0F72
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
      Columns.Count   =   13
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ProductID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3519
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4551
      Columns(2).Caption=   "Product Name"
      Columns(2).Name =   "ProductName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3651
      Columns(3).Caption=   "Pack Name"
      Columns(3).Name =   "PackName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   953
      Columns(4).Caption=   "Pack"
      Columns(4).Name =   "Pack"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1323
      Columns(5).Caption=   "Qt.Pack"
      Columns(5).Name =   "QtyPack"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1429
      Columns(6).Caption=   "Qt.Loose"
      Columns(6).Name =   "QtyLoose"
      Columns(6).Alignment=   1
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1693
      Columns(7).Caption=   "Pur Price"
      Columns(7).Name =   "PurPrice"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   1191
      Columns(8).Caption=   "DiscPc"
      Columns(8).Name =   "DiscPc"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   847
      Columns(9).Caption=   "Dis%"
      Columns(9).Name =   "DiscPer"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1191
      Columns(10).Caption=   "DiscVal"
      Columns(10).Name=   "DiscVal"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   2461
      Columns(11).Caption=   "Amount"
      Columns(11).Name=   "Amount"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   5
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "PackingID"
      Columns(12).Name=   "PackingID"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   23310
      _ExtentY        =   6773
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   11640
      TabIndex        =   30
      Top             =   1095
      Visible         =   0   'False
      Width           =   1860
      _ExtentX        =   3281
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
   Begin SITextBox.Txt TxtMultiplier 
      Height          =   315
      Left            =   7695
      TabIndex        =   7
      Top             =   3285
      Width           =   540
      _ExtentX        =   953
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
   Begin SITextBox.Txt TxtQtyLoose 
      Height          =   315
      Left            =   8985
      TabIndex        =   9
      Top             =   3285
      Width           =   810
      _ExtentX        =   1429
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
      Left            =   8235
      TabIndex        =   8
      Top             =   3285
      Width           =   750
      _ExtentX        =   1323
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
      Masked          =   1
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   12585
      TabIndex        =   14
      Top             =   3285
      Width           =   1680
      _ExtentX        =   2963
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpOpeningDate 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
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
   Begin SITextBox.Txt TxtOpeningID 
      Height          =   315
      Left            =   1710
      TabIndex        =   0
      Top             =   1680
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3766
      TabIndex        =   40
      Top             =   8978
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
      MICON           =   "FrmOpeningStockInvoice.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   4065
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   1680
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
      Left            =   5100
      TabIndex        =   41
      Tag             =   "NC"
      Top             =   1680
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   4740
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1680
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
      MICON           =   "FrmOpeningStockInvoice.frx":0FAA
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   6540
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   1680
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
      Left            =   7845
      TabIndex        =   43
      Tag             =   "NC"
      Top             =   1680
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
      Left            =   7485
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1680
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
      MICON           =   "FrmOpeningStockInvoice.frx":0FC6
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtDiscVal 
      Height          =   315
      Left            =   11910
      TabIndex        =   13
      Top             =   3285
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
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   11430
      TabIndex        =   12
      Top             =   3285
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
      Left            =   10755
      TabIndex        =   11
      Top             =   3285
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
   Begin SITextBox.Txt TxtTotalItems 
      Height          =   315
      Left            =   1050
      TabIndex        =   53
      Top             =   7815
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
   Begin SITextBox.Txt TxtTotalPack 
      Height          =   315
      Left            =   8235
      TabIndex        =   55
      Top             =   7740
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
   Begin SITextBox.Txt TxtTotalLoose 
      Height          =   315
      Left            =   9180
      TabIndex        =   57
      Top             =   7740
      Width           =   1425
      _ExtentX        =   2514
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Loose"
      Height          =   195
      Left            =   9180
      TabIndex        =   58
      Top             =   7515
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pack"
      Height          =   195
      Left            =   8235
      TabIndex        =   56
      Top             =   7515
      Width           =   780
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items"
      Height          =   195
      Left            =   1050
      TabIndex        =   54
      Top             =   7590
      Width           =   780
   End
   Begin VB.Label LblDiscPC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc/PC"
      Height          =   195
      Left            =   10721
      TabIndex        =   52
      Top             =   3083
      Width           =   600
   End
   Begin VB.Label LblDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Dis%"
      Height          =   195
      Left            =   11426
      TabIndex        =   51
      Top             =   3083
      Width           =   345
   End
   Begin VB.Label LblDiscVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.Val"
      Height          =   195
      Left            =   11906
      TabIndex        =   50
      Top             =   3083
      Width           =   585
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Price"
      Height          =   195
      Left            =   9855
      TabIndex        =   49
      Top             =   3090
      Width           =   645
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   4065
      TabIndex        =   48
      Top             =   1485
      Width           =   585
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   5100
      TabIndex        =   47
      Top             =   1485
      Width           =   840
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   6540
      TabIndex        =   46
      Top             =   1485
      Width           =   1095
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   7845
      TabIndex        =   45
      Top             =   1485
      Width           =   1350
   End
   Begin VB.Label LblStock 
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
      Height          =   450
      Left            =   10350
      TabIndex        =   39
      Top             =   1080
      Width           =   975
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
      Left            =   10350
      TabIndex        =   38
      Top             =   750
      Width           =   720
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Stock Invoice"
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
      Left            =   2520
      TabIndex        =   37
      Top             =   270
      Width           =   3900
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   12596
      TabIndex        =   36
      Top             =   3083
      Width           =   540
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (Pack)"
      Height          =   195
      Left            =   8194
      TabIndex        =   35
      Top             =   3083
      Width           =   750
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (Loose)"
      Height          =   195
      Left            =   8989
      TabIndex        =   34
      Top             =   3083
      Width           =   810
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   5629
      TabIndex        =   33
      Top             =   3083
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   7699
      TabIndex        =   32
      Top             =   3083
      Width           =   375
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   11640
      TabIndex        =   31
      Top             =   900
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   1054
      TabIndex        =   29
      Top             =   3083
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3049
      TabIndex        =   28
      Top             =   3083
      Width           =   1020
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   1710
      TabIndex        =   24
      Top             =   2325
      Width           =   795
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   12570
      TabIndex        =   23
      Top             =   7560
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Date"
      Height          =   195
      Left            =   2760
      TabIndex        =   22
      Top             =   1485
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Opening ID"
      Height          =   195
      Left            =   1710
      TabIndex        =   21
      Top             =   1485
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
Attribute VB_Name = "FrmOpeningStockInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
'Dim RsReport As New ADODB.Recordset
Dim vUnitPrice As Double
Dim i As Integer
Dim Flag As Boolean
Dim ssql As String
Dim vStrSQL As String
'----------------------------------

Private Sub SubCalculateBody()
   TxtDiscVal.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) * Val(TxtDiscPer.Text) / 100, 2)
   If Val(TxtDiscVal.Text) = 0 Then TxtDiscVal.Text = ""
   TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), 2)
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
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.ParaInPurchase = True
      SchProduct.ParaInWhere = " and isLocked = 0 and isNoCostProduct = 0 and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
   If TxtCode.Enabled = False Then FunSelectProduct = False: Exit Function
   If Trim(TxtCode.Text) = "" Then FunSelectProduct = False: Exit Function
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
        vStrSQL = " SELECT p.productid, Code, ProductName, PurPrice, WSPrice, RetailPrice, IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST, SaleTaxPer, PurDiscPC, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where p.productid = " & Val(TxtCode.Text) & " or code='" & TxtCode.Text & "'"
 
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPurPrice.Text = !PurPrice
         'LblRetailPrice.Caption = !RetailPrice
         If IsNull(!PackingName) Then
            vUnitPrice = !PurPrice
            TxtMultiplier.Text = ""
            CmbPackName.ListIndex = 0
         Else
            TxtMultiplier.Text = !Multiplier
            If !Multiplier <> 0 Then
               vUnitPrice = !PurPrice / !Multiplier
            Else
               vUnitPrice = !PurPrice
            End If
            CmbPackName.Text = !PackingName
         End If
         TxtDiscPC.Text = IIf(IsNull(!PurDiscPC), "", !PurDiscPC)
         If vUnitPrice = 0 Then
            TxtDiscPer.Text = "0"
         Else
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
         End If
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         TxtProductID.Text = ""
         TxtCode.Text = ""
         If CmbPackName.ListCount > 0 Then CmbPackName.ListIndex = 0
         TxtProductName.Text = ""
         TxtMultiplier.Text = ""
         TxtPurPrice.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtAmount.Text = ""
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
    CN.Execute ("Insert Into UserActivities values ('Opening Invoice'" & "," & TxtOpeningID.Text & ",'" & DtpOpeningDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
      '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Opening Invoice'" & "," & TxtOpeningID.Text & ",'" & DtpOpeningDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub btnConvert_Click()
   On Error GoTo ErrorHandler
   FrmOpeningProductGrid.ParaInOpeningDate = DtpOpeningDate.DateValue
   FrmOpeningProductGrid.ParaInOpeningID = TxtOpeningID.Text
   FrmOpeningProductGrid.Show vbModal
   RsTemp.Filter = ""
   If RsTemp.RecordCount > 0 Then
      PopulateDataToGrid
      PopulateTempToGrid
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateTempToGrid()
    On Error GoTo ErrorHandler
'    RsBody.Filter = 0
'    For i = 1 To RsBody.RecordCount
'       RsBody.Delete
'       RsBody.MoveNext
'    Next i
    Grid.Redraw = False
    Grid.CancelUpdate
    Grid.RemoveAll
    Grid.AddNew
    Grid.Columns("Code").Text = " "
    With RsTemp
      RsTemp.MoveFirst
      Grid.AllowAddNew = True
      TxtTotalAmount.Text = 0
      TxtTotalItems.Text = 0
      While Not .EOF
        vUnitPrice = 0
         If Val(!Multiplier) <> 0 Then
            vUnitPrice = Round(Val(!Price) / Val(!Multiplier), 2)
         Else
            vUnitPrice = Val(!Price)
         End If
         Grid.Columns("ProductID").Text = !Productid
         Grid.Columns("Code").Text = !Productid
         Grid.Columns("ProductName").Text = !ProductName
         Grid.Columns("PackName").Text = !PackingName
         Grid.Columns("QtyPack").Text = IIf(Val(!QtyPack) = 0, "", Val(!QtyPack))
         Grid.Columns("Pack").Text = !Multiplier
         RsBody.Filter = "ProductID = " & !Productid
         'RsBody.AddNew
         
         'RsBody!Productid = !Productid
         Grid.Columns("QtyLoose").Value = !QtyLoose
         Grid.Columns("PurPrice").Value = !Price
         Grid.Columns("Amount").Value = Round((vUnitPrice * ((!QtyPack * !Multiplier) + !QtyLoose)), 2)

         '''''
         
         RsBody!Multiplier = !Multiplier
         RsBody!QtyPack = !QtyPack
         RsBody!QtyLoose = !QtyLoose
         RsBody!PurPrice = !Price
         RsBody!DiscPC = !DiscPC
         RsBody!DiscPer = !DiscPer
         RsBody!DiscVal = 0
         RsBody!Amount = Round((vUnitPrice * ((!QtyPack * !Multiplier) + !QtyLoose)), 2)
         RsBody.Update
         ''''
         
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Round((vUnitPrice * ((!QtyPack * !Multiplier) + !QtyLoose)), 2)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + 1
         .MoveNext
         Grid.AddNew
      Wend
      .Close
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End With
'   Grid.AddNew
   Grid.Columns("Code").Text = " "
   Grid.AllowAddNew = False
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
'      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   CN.BeginTrans
   
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Opening Invoice'" & "," & TxtOpeningID.Text & ",'" & DtpOpeningDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
     ''''''''''''''''''''''''''Delete Opening Stock Body'''''''''''''''''''''
   Grid.Redraw = False
   Grid.MoveFirst
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from OpeningStock where OpeningID = " & Val(TxtOpeningID.Text) & " and OpeningDate ='" & DtpOpeningDate.DateValue & "' and productid ='" & Grid.Columns("ProductID").Text & "'"
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   CN.Execute "Delete from OpeningInvoiceHeader where OpeningID = " & Val(TxtOpeningID.Text) & " and OpeningDate ='" & DtpOpeningDate.DateValue & "'"
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchOpeningInvoice.ParaOutOpeningDate = DtpOpeningDate.DateValue
   SchOpeningInvoice.Show vbModal
   If SchOpeningInvoice.ParaOutOpeningID <> 0 Then
      TxtOpeningID.Text = SchOpeningInvoice.ParaOutOpeningID
      DtpOpeningDate.DateValue = SchOpeningInvoice.ParaOutOpeningDate
      CN.Execute ("Insert Into UserActivities values ('Opening Invoice'" & "," & TxtOpeningID.Text & ",'" & DtpOpeningDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetOpeningInvoice
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   vStrSQL = " Select h.OpeningID, h.OpeningDate, h.Description, h.TotalAmount, h.storeid, s.storeName, b.ProductID,  ProductName, " & vbCrLf _
            + " isnull(pa.PackingName,'PC') as PackingName, b.qtyloose as NetQtyLoose, b.QtyPack, b.QtyLoose, b.Multiplier, b.PurPrice, b.DiscPC, b.DiscPer, b.DiscVal, b.QtyLoose*b.PurPrice as Value, amount " & vbCrLf _
            + " from OpeningInvoiceHeader h inner join OpeningStock b on h.Openingid = b.OpeningID and h.Openingdate = b.Openingdate " & vbCrLf _
            + " inner join stores s on s.storeid = h.storeid " & vbCrLf _
            + " Inner Join Products p on p.productid = b.productid " & vbCrLf _
            + " left outer join packings pa on pa.packingid = p.purchasepackingid " & vbCrLf _
            + " where h.Openingid = " & Val(TxtOpeningID.Text) & " and h.OpeningDate = '" & DtpOpeningDate.DateValue & "' order by Serialno"
            
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   'Set RptReportViewer.Report = New CRptOpeningInvoice
   Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\CRptOpeningInvoice.rpt")
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ReportTitle = "Opening Invoice"
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   CN.Execute ("Insert Into UserActivities values ('Opening Invoice'" & "," & TxtOpeningID.Text & ",'" & DtpOpeningDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
   RptReportViewer.Report.PaperOrientation = crPortrait
   If MsgBox("Do you want to print directly this Invoice.", vbQuestion + vbYesNo, "Alert") = vbYes Then
      RptReportViewer.Report.PrintOut False
   Else
      RptReportViewer.Show vbModal
   End If
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
'  '''''''''''''''''''''''Check Entry Date'''''''''''''''''''''''''''''''''
'    If ObjRegistry.isEntryDate = True Then
'       If ObjRegistry.FromDate > Date Or ObjRegistry.ToDate < Date Then
'         MsgBox "Data can not be saved Because Date is not set according to the Software's Entry date", vbInformation, Me.Caption
'         Exit Sub
'       End If
'    End If
'  Header Validation
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      TxtStoreID.SetFocus
      Exit Sub
   End If
   
   If DtpOpeningDate.Enabled Then
      If CN.Execute("Select * from OpeningInvoiceHeader where OpeningID = " & Val(TxtOpeningID.Text) & " and OpeningDate = '" & DtpOpeningDate.DateValue & "'").RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtOpeningID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
  RsBody.Filter = 0
  If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Opening Stock Invoice", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
  End If
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
  
   ''''''''''''''''''''''''Check Organization '''''''''''''''''''''''''''''''''
  If ObjRegistry.OrganizationMandatory = True And TxtOrganizationID.Text = "" Then
    MsgBox "Please Select Organization", vbInformation, Me.Caption
    If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
    Exit Sub
  End If
   CN.BeginTrans
   
   Call UserActivities
   
   ssql = "select * from OpeningInvoiceHeader where OpeningID = " & Val(TxtOpeningID.Text) & " and OpeningDate = '" & DtpOpeningDate.DateValue & "'"
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, CN, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !OpeningID = Val(TxtOpeningID.Text)
         !OpeningDate = DtpOpeningDate.DateValue
      End If
      !StoreID = TxtStoreID.Text
      !OrganizationID = IIf(Trim(TxtOrganizationID.Text) = "", Null, TxtOrganizationID.Text)
      !TotalAmount = Val(TxtTotalAmount.Text)
      !Description = IIf(Trim(TxtDescription.Text) = "", Null, TxtDescription.Text)
      '!UserNo = vUser
      !UserNo = 1
      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
      .Update
      .Close
   End With
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !OpeningID = Val(TxtOpeningID.Text)
         !OpeningDate = DtpOpeningDate.DateValue
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   CN.CommitTrans
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
   RsBody.Open "Select * from OpeningStock where OpeningID = " & Val(TxtOpeningID.Text) & " and OpeningDate = '" & DtpOpeningDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.productname, b.* from OpeningStock b join products p on p.productid = b.productid where OpeningID = " & Val(TxtOpeningID.Text) & " and OpeningDate = '" & DtpOpeningDate.DateValue & "' order by Serialno"
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         TxtTotalItems.Text = 0
         TxtTotalPack.Text = 0
         TxtTotalLoose.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = IIf(IsNull(!code), !Productid, !code)
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
            TxtTotalPack.Text = TxtTotalPack.Text + Grid.Columns("QtyPack").Value
            Grid.Columns("QtyLoose").Value = !QtyLoose
            TxtTotalLoose.Text = TxtTotalLoose.Text + Grid.Columns("QtyLoose").Value
            Grid.Columns("PurPrice").Value = !PurPrice
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(!Amount)
            TxtTotalItems.Text = Val(TxtTotalItems.Text) + 1
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
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
      Call PopulateDataToGrid
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtOpeningID.Text = FunGetMaxID()
      DtpOpeningDate.Enabled = True
      If DtpOpeningDate.Enabled And DtpOpeningDate.Visible Then DtpOpeningDate.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      DtpOpeningDate.Enabled = False
      BtnPrint.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      'BtnPrint.Enabled = True
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      vIsNewRecord = False
   Case Is = ChangeMode
      'BtnPrint.Enabled = False
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

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub CmbPackName_Click()
   On Error GoTo ErrorHandler
   If CmbPackName.Text = "" Then
      TxtMultiplier.Text = ""
      TxtQtyPack.Text = ""
      TxtMultiplier.Enabled = False
      TxtQtyPack.Enabled = False
      TxtPurPrice.Text = Round(vUnitPrice, 3)
      TxtQtyLoose.Enabled = True
   Else
      If ObjRegistry.ChangeQtyPack = True Then TxtMultiplier.Enabled = True Else TxtMultiplier.Enabled = False
      TxtQtyPack.Enabled = True
      TxtQtyLoose.Enabled = True 'Not ObjRegistry.EitherPackORLooseEnter
      If Trim(TxtCode.Text) <> "" Then
         With CN.Execute("select * from ProductPacking where ProductID = " & Val(TxtProductID.Text) & " and packingid = " & CmbPackName.ItemData(CmbPackName.ListIndex))
            TxtMultiplier.Text = IIf(.RecordCount = 0, "", !Multiplier)
            If Val(TxtMultiplier.Text) <> 0 Then
               TxtPurPrice.Text = Round(vUnitPrice * !Multiplier, 3)
            Else
               TxtPurPrice.Text = Round(vUnitPrice, 3)
            End If
            Call SubCalculateBody
         .Close
         End With
      End If
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub DtpOpeningDate_Change()
   TxtOpeningID.Text = FunGetMaxID()
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
      Call SubClearDetailArea: TxtCode.SetFocus
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
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then CmbPackName.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtDescription.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtOrganizationID.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
         KeyCode = 0
         TxtDescription.SetFocus
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

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Opening Stock Invoice"
   
   DtpOpeningDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))

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
   
   With CN.Execute("Select * from Packings")
      CmbPackName.AddItem ""
      While Not .EOF
         CmbPackName.AddItem !PackingName
         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
         .MoveNext
      Wend
      .Close
   End With
   CmbPackName.ListIndex = 0
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
'   If DtpOpeningDate.IsDateValid = False Then Exit Function
'   FunGetMaxID = cn.Execute("Select isnull(max(OpeningID),0)+1 from OpeningInvoiceHeader Where OpeningDate = '" & DtpOpeningDate.DateValue & "'").Fields(0)
   FunGetMaxID = CN.Execute("Select isnull(max(OpeningID),0)+1 from OpeningInvoiceHeader ").Fields(0)
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
   DtpOpeningDate.DateValue = Date
   TxtTotalAmount.Text = 0
   TxtTotalItems.Text = 0
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("Code").Text = " "
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
    Set FrmOpeningStockInvoice = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("Amount").Value
   TxtTotalItems.Text = Val(TxtTotalItems.Text) - 1
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
      CmbPackName.SetFocus
      If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(Grid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
   RsBody.Filter = "ProductID = " & Val(TxtCode.Text)
   If RsBody.RecordCount > 0 Then RsBody.Delete
   CN.Execute ("Insert Into UserActivities values ('Opening Invoice'" & "," & TxtOpeningID.Text & ",'" & DtpOpeningDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Cost-" & Grid.Columns("PurPrice").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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
   If Trim(TxtQtyPack.Text) = "" And Trim(TxtQtyLoose.Text) = "" Then
      If TxtQtyPack.Enabled = True Then TxtQtyPack.SetFocus
      Exit Sub
   End If
   If Val(TxtPurPrice.Text) <> 0 Then
      If Round(Val(TxtDiscPer.Text), 2) <> Round((Val(TxtDiscPC.Text) * 100) / (Val(TxtPurPrice.Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))), 2) Then
         MsgBox "Please update the Discount for change Price.", vbExclamation, "Alert"
         If TxtDiscPer.Enabled And TxtDiscPer.Visible Then TxtDiscPer.SetFocus
         Exit Sub
      End If
   End If

On Error GoTo ErrorHandler
   RsBody.Filter = "ProductID = " & Val(TxtProductID.Text)
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ProductID").Text = TxtProductID.Text
         Grid.Columns("Code").Text = TxtCode.Text
         RsBody!code = TxtCode.Text
         RsBody!Productid = TxtProductID.Text
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("ProductID").Text = TxtProductID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  TxtQtyLoose.Text = Val(TxtQtyLoose.Text) + Val(Grid.Columns("QtyLoose").Value)
                  TxtQtyPack.Text = Val(TxtQtyPack.Text) + Val(Grid.Columns("QtyPack").Value)
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
                  Grid.Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
                  Grid.Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, 0, Val(TxtMultiplier.Text))
                  Grid.Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
                  Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
                  Grid.Columns("PurPrice").Value = Val(TxtPurPrice.Text)
                  Grid.Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  Grid.Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  Grid.Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  RsBody!StoreID = TxtStoreID.Text
                  RsBody!PackingID = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), Null)
                  RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
                  RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
                  RsBody!QtyLoose = Val(TxtQtyLoose.Text)
                  RsBody!PurPrice = Val(TxtPurPrice.Text)
                  RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  RsBody!Amount = Val(TxtAmount.Text)
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
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + 1
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
      End If
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
      .Columns("Pack").Value = Val(TxtMultiplier.Text)
      .Columns("QtyPack").Text = TxtQtyPack.Text
      .Columns("QtyLoose").Text = TxtQtyLoose.Text
      .Columns("PurPrice").Value = Val(TxtPurPrice.Text)
      .Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      .Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      .Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      .Columns("Amount").Value = Val(TxtAmount.Text)
      RsBody!StoreID = TxtStoreID.Text
      RsBody!PackingID = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), Null)
      RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
      RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
      RsBody!QtyLoose = Val(TxtQtyLoose.Text)
      RsBody!PurPrice = Val(TxtPurPrice.Text)
      RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      RsBody!Amount = Val(TxtAmount.Text)
      .MoveNext
'      If TxtCode.Enabled = True Then
'         If Trim(.Columns("ProductID").Text) <> "" Then
'            .AllowAddNew = True
'            .AddNew
'            .Columns("ProductID").Text = " "
'            .AllowAddNew = False
'         End If
'      Else
'      CmbPackName.SetFocus
'      End If
    If TxtCode.Enabled = False And ObjRegistry.AfterRowEditFocusNextGridLine = True Then
         
'         Grid.MoveNext
         Call Grid_GotFocus
         CmbPackName.SetFocus
      Else
        
         Grid.MoveLast
         If Trim(.Columns("Code").Text) <> "" Then
         Grid.AllowAddNew = True
         Grid.AddNew
         Grid.Columns("Code").Text = " "
         Grid.AllowAddNew = False
         
      End If
   End If
   End With
   
   If Trim(Grid.Columns("ProductID").Text) = "" Then
            Call SubClearDetailArea
            TxtCode.SetFocus
   End If
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
   TxtPurPrice.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtDiscVal.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("Code").Text
      TxtProductName.Text = .Columns("ProductName").Text
      If .Columns("PackName").Text = "" Then
         CmbPackName.ListIndex = 0
      Else
         CmbPackName.Text = .Columns("PackName").Text
      End If
      TxtMultiplier.Text = .Columns("Pack").Text
      TxtQtyLoose.Text = .Columns("QtyLoose").Text
      TxtQtyPack.Text = .Columns("QtyPack").Text
      TxtPurPrice.Text = .Columns("PurPrice").Text
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      TxtAmount.Text = .Columns("Amount").Value
      If Val(TxtMultiplier.Text) = 0 Then
         vUnitPrice = IIf(.Columns("PurPrice").Text = "", 0, .Columns("PurPrice").Text)
      Else
         vUnitPrice = .Columns("PurPrice").Text / Val(TxtMultiplier.Text)
      End If
   End With
   If Grid.Rows = 1 Then Grid.MoveNext
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetOpeningInvoice()
   On Error GoTo ErrorHandler
   ssql = "select h.*, StoreName, OrganizationName FROM OpeningInvoiceHeader h inner join stores s on s.storeid = h.storeid left outer join Organizations o on o.organizationid = h.organizationid where h.OpeningID = " & Val(TxtOpeningID.Text) & " and OpeningDate ='" & DtpOpeningDate.DateValue & "'"
   With CN.Execute(ssql)
      If Not .BOF Then
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          TxtTotalAmount.Text = !TotalAmount
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
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

Private Sub TxtQtyLoose_Change()
   Call SubCalculateBody
End Sub

Private Sub TxtQtyLoose_LostFocus()
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtQtyPack_Change()
   Call SubCalculateBody
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
      TxtPurPrice.Text = ""
   End If
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

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
      TxtCode.SetFocus
   Else
      TxtOrganizationID.SetFocus
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

Private Sub TxtDiscPC_Change()
   If ActiveControl.Name <> TxtDiscPC.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
   If Val(TxtDiscPer.Text) = 0 Then TxtDiscPer.Text = ""
   Call SubCalculateBody
End Sub

Private Sub TxtDiscPer_Change()
   If ActiveControl.Name <> TxtDiscPer.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   TxtDiscPC.Text = Round((vUnitPrice * Val(TxtDiscPer.Text) / 100), 4)
   If Val(TxtDiscPC.Text) = 0 Then TxtDiscPC.Text = ""
   Call SubCalculateBody
End Sub

Private Sub TxtDiscVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDiscVal.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   If (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) = 0 Then Exit Sub
   TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 4)
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
   TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), 2)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDiscVal_LostFocus()
'   Select Case ActiveControl.Name
'   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtBonus.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPurPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtOffer.Name, TxtSaleTaxPer.Name
'      Exit Sub
'   End Select
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtMultiplier_Change()
   If ActiveControl.Name <> TxtMultiplier.Name Then Exit Sub
   If Val(TxtMultiplier.Text) <> 0 Then
      TxtPurPrice.Text = Round(vUnitPrice * Val(TxtMultiplier.Text), 3)
   Else
      TxtPurPrice.Text = Round(vUnitPrice, 3)
   End If
   Call SubCalculateBody
End Sub

Private Sub TxtMultiplier_Validate(Cancel As Boolean)
    If ActiveControl.Name <> TxtQtyPack.Name Then Exit Sub
End Sub

Private Sub UserActivities()
     If vIsNewRecord = False Then
     
    With CN.Execute("Select  * from OpeningInvoiceHeader where OpeningID =" & TxtOpeningID.Text & " And OpeningDate = '" & DtpOpeningDate.DateValue & "'")
        If TxtStoreID.Text <> !StoreID Then
            CN.Execute ("Insert Into UserActivities values ('Opening Invoice'" & "," & TxtOpeningID.Text & ",'" & DtpOpeningDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.Rows - 1
        With CN.Execute("Select * from OpeningStock Where OpeningID = " & TxtOpeningID.Text & " and OpeningDate ='" & DtpOpeningDate.DateValue & "' and Productid ='" & Grid.Columns("Productid").Text & "'")
        
             If .EOF = True Then
                CN.Execute ("Insert Into UserActivities values ('Opening Invoice'" & "," & TxtOpeningID.Text & ",'" & DtpOpeningDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Cost-" & Grid.Columns("PurPrice").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("QtyPack").Text <> !QtyPack Or Grid.Columns("QtyLoose").Text <> !QtyLoose Then
                   CN.Execute ("Insert Into UserActivities values ('Opening Invoice'" & "," & TxtOpeningID.Text & ",'" & DtpOpeningDate.DateValue & "','Updated ProdcutID-" & Grid.Columns("Code").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Cost-" & Grid.Columns("PurPrice").Text & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
    
   Else
    CN.Execute ("Insert Into UserActivities values ('Opening Invoice'" & "," & TxtOpeningID.Text & ",'" & DtpOpeningDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
   
End Sub
