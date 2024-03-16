VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmSaleReturnInvoice 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11970
   ControlBox      =   0   'False
   Icon            =   "FrmSaleReturnInvoice1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkIsProduct 
      Caption         =   "Is Product"
      Height          =   255
      Left            =   5760
      TabIndex        =   56
      Top             =   240
      Visible         =   0   'False
      Width           =   1365
   End
   Begin SITextBox.Txt TxtReturnID 
      Height          =   315
      Left            =   165
      TabIndex        =   18
      Top             =   1035
      Width           =   915
      _ExtentX        =   1614
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
   Begin SITextBox.Txt TxtDiscVal 
      Height          =   315
      Left            =   9135
      TabIndex        =   21
      Top             =   1845
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   165
      TabIndex        =   5
      Top             =   1845
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
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   5880
      TabIndex        =   6
      Top             =   1845
      Width           =   780
      _ExtentX        =   1376
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
   Begin SITextBox.Txt TxtPrice 
      Height          =   315
      Left            =   6660
      TabIndex        =   7
      Top             =   1845
      Width           =   960
      _ExtentX        =   1693
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
      DecimalPoint    =   2
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   10125
      TabIndex        =   19
      Top             =   1845
      Width           =   1650
      _ExtentX        =   2910
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
   Begin JeweledBut.JeweledButton BtnProduct 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2025
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1845
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
      MICON           =   "FrmSaleReturnInvoice1.frx":0ECA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7335
      TabIndex        =   16
      Top             =   8040
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
      MICON           =   "FrmSaleReturnInvoice1.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6015
      TabIndex        =   12
      Top             =   8040
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
      MICON           =   "FrmSaleReturnInvoice1.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3375
      TabIndex        =   14
      Top             =   8040
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
      MICON           =   "FrmSaleReturnInvoice1.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8655
      TabIndex        =   17
      Top             =   8040
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
      MICON           =   "FrmSaleReturnInvoice1.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4695
      TabIndex        =   13
      Top             =   8040
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
      MICON           =   "FrmSaleReturnInvoice1.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillDisc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   150
      TabIndex        =   10
      Top             =   6390
      Width           =   1515
      _ExtentX        =   2672
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
      Masked          =   1
      IntegralPoint   =   4
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2385
      TabIndex        =   31
      Top             =   1845
      Width           =   3495
      _ExtentX        =   6165
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
      CausesValidation=   0   'False
      Height          =   3945
      Left            =   165
      TabIndex        =   32
      Top             =   2160
      Width           =   11625
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
      stylesets(0).Picture=   "FrmSaleReturnInvoice1.frx":0F72
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
      RowHeight       =   503
      ExtraHeight     =   132
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   13
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Sr"
      Columns(0).Name =   "Sr"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "Product ID"
      Columns(1).Name =   "ProductID"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3916
      Columns(2).Caption=   "Code"
      Columns(2).Name =   "Code"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   6165
      Columns(3).Caption=   "Product Name"
      Columns(3).Name =   "ProductName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1376
      Columns(4).Caption=   "Qty"
      Columns(4).Name =   "Qty"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   4
      Columns(4).FieldLen=   256
      Columns(5).Width=   1693
      Columns(5).Caption=   "Price"
      Columns(5).Name =   "Price"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   4
      Columns(5).FieldLen=   256
      Columns(6).Width=   1455
      Columns(6).Caption=   "Disc / Pc"
      Columns(6).Name =   "DiscPC"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1217
      Columns(7).Caption=   "Disc%"
      Columns(7).Name =   "DiscPer"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1746
      Columns(8).Caption=   "Disc. Val"
      Columns(8).Name =   "DiscVal"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   4
      Columns(8).FieldLen=   256
      Columns(9).Width=   2487
      Columns(9).Caption=   "Amount"
      Columns(9).Name =   "Amount"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   5
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "TotalAmount"
      Columns(10).Name=   "TotalAmount"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Visible=   0   'False
      Columns(11).Caption=   "Cost"
      Columns(11).Name=   "Cost"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   4
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "IsProduct"
      Columns(12).Name=   "IsProduct"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   11
      Columns(12).FieldLen=   256
      Columns(12).Style=   2
      TabNavigation   =   1
      _ExtentX        =   20505
      _ExtentY        =   6959
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2055
      TabIndex        =   15
      Top             =   8040
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
      MICON           =   "FrmSaleReturnInvoice1.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   7620
      TabIndex        =   8
      Top             =   1845
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
   Begin SITextBox.Txt TxtActualAmount 
      Height          =   315
      Left            =   7125
      TabIndex        =   38
      Top             =   195
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   6990
      TabIndex        =   4
      Tag             =   "NC"
      Top             =   1035
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
      Left            =   8025
      TabIndex        =   40
      Tag             =   "NC"
      Top             =   1035
      Width           =   2340
      _ExtentX        =   4128
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
      Left            =   7665
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1035
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
      MICON           =   "FrmSaleReturnInvoice1.frx":0FAA
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpReturnDate 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   1035
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
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   8445
      TabIndex        =   9
      Top             =   1845
      Width           =   690
      _ExtentX        =   1217
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
   Begin SITextBox.Txt TxtPID 
      Height          =   315
      Left            =   8775
      TabIndex        =   45
      Top             =   195
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
   Begin SITextBox.Txt TxtCost 
      Height          =   315
      Left            =   10650
      TabIndex        =   57
      Top             =   195
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
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   2790
      TabIndex        =   1
      Top             =   1035
      Width           =   915
      _ExtentX        =   1614
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
      Left            =   3705
      TabIndex        =   2
      Top             =   1035
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
   Begin SITextBox.Txt TxtBillDiscPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   150
      TabIndex        =   11
      Top             =   7065
      Width           =   1515
      _ExtentX        =   2672
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
   Begin JeweledBut.JeweledButton BtnReturnAll 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5370
      TabIndex        =   3
      Top             =   1020
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
      MICON           =   "FrmSaleReturnInvoice1.frx":0FC6
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSale 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5010
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   1020
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
      MICON           =   "FrmSaleReturnInvoice1.frx":0FE2
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Discount (%)"
      Height          =   195
      Left            =   150
      TabIndex        =   61
      Top             =   6840
      Width           =   1125
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " Bill ID"
      Height          =   195
      Left            =   2790
      TabIndex        =   60
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   3705
      TabIndex        =   59
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   10650
      TabIndex        =   58
      Top             =   -30
      Visible         =   0   'False
      Width           =   315
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
      Left            =   10410
      TabIndex        =   55
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
      Left            =   10365
      TabIndex        =   54
      Top             =   765
      Width           =   720
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Return Invoice"
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
      Left            =   1890
      TabIndex        =   53
      Top             =   135
      Width           =   3420
   End
   Begin VB.Label TxtLastRate 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "VCRSCapsSSK"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   915
      Left            =   2040
      TabIndex        =   52
      Top             =   6540
      Width           =   1785
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Item Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   1980
      TabIndex        =   51
      Top             =   6240
      Width           =   1830
   End
   Begin VB.Label TxtNetAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "VCRSCapsSSK"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   9435
      TabIndex        =   50
      Top             =   6540
      Width           =   2370
   End
   Begin VB.Label TxtTotalDiscount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "VCRSCapsSSK"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   7665
      TabIndex        =   49
      Top             =   6540
      Width           =   1740
   End
   Begin VB.Label TxtTotalAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "VCRSCapsSSK"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   5265
      TabIndex        =   48
      Top             =   6540
      Width           =   2370
   End
   Begin VB.Label TxtTotalQty 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "VCRSCapsSSK"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   915
      Left            =   3855
      TabIndex        =   47
      Top             =   6540
      Width           =   1380
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   8775
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. %"
      Height          =   195
      Left            =   8430
      TabIndex        =   44
      Top             =   1650
      Width           =   525
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   6990
      TabIndex        =   43
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   8025
      TabIndex        =   42
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Amount"
      Height          =   195
      Left            =   7110
      TabIndex        =   39
      Top             =   -15
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Discount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   7605
      TabIndex        =   37
      Top             =   6270
      Width           =   1755
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   5415
      TabIndex        =   36
      Top             =   6270
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   6660
      TabIndex        =   35
      Top             =   1650
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   3840
      TabIndex        =   34
      Top             =   6240
      Width           =   1365
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   11610
      Top             =   60
      Width           =   345
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   2385
      TabIndex        =   33
      Top             =   1650
      Width           =   1020
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   9900
      TabIndex        =   30
      Top             =   6270
      Width           =   1440
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Discount"
      Height          =   195
      Left            =   150
      TabIndex        =   29
      Top             =   6165
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      Height          =   195
      Left            =   1170
      TabIndex        =   28
      Top             =   840
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Return ID"
      Height          =   195
      Left            =   165
      TabIndex        =   27
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   165
      TabIndex        =   26
      Top             =   1650
      Width           =   375
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc / PC"
      Height          =   195
      Left            =   7620
      TabIndex        =   25
      Top             =   1650
      Width           =   690
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   5880
      TabIndex        =   24
      Top             =   1650
      Width           =   240
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   10125
      TabIndex        =   23
      Top             =   1650
      Width           =   540
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. Val"
      Height          =   195
      Left            =   9135
      TabIndex        =   22
      Top             =   1650
      Width           =   630
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmSaleReturnInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vCounter As Integer
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vIsNewRecord As Boolean
Dim Flag As Boolean
Dim DateFlag As Boolean
Dim sSql As String
Dim vStrSQL As String
Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String, vTotDisc As Double
'----------------------------------

Private Sub SubCalculateBody()
    TxtDiscVal.Text = Val(TxtQty.Text) * Val(TxtDiscPC.Text)
    TxtActualAmount.Text = Val(TxtQty.Text) * Val(TxtPrice.Text)
    TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
    TxtTotalDiscount.Caption = vTotDisc
    SubCalculateFooter
End Sub

Private Sub SubCalculateFooter()
   TxtTotalDiscount.Caption = Val(TxtBillDisc.Text) + vTotDisc
   TxtNetAmount.Caption = Val(TxtTotalAmount.Caption) - Val(TxtTotalDiscount.Caption)
End Sub

Private Function FunSelectSale(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSale.Show vbModal, Me
        If SchSale.ParaOutBillID = "" Then FunSelectSale = False: Exit Function
        TxtBillID.Text = SchSale.ParaOutBillID
        DtpBillDate.DateValue = SchSale.ParaOutBillDate
    End If
    '---------------------------
    vStrSQL = " Select * FROM SaleHeader where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          DtpBillDate.DateValue = !BillDate
          FunSelectSale = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSale = False
          .Close
          TxtBillID.Text = ""
          DtpBillDate.DateValue = Date
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSale.Show vbModal, Me
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
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If TxtCode.Enabled = False Then FunSelectProduct = False: Exit Function
    If Trim(TxtCode.Text) = "" Then FunSelectProduct = False: Exit Function
    If Len(TxtCode.Text) <= 5 Then
      TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
    End If
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    
    ''''''''***********   Checking Union   ***********''''''''
    vStrSQL = " SELECT p.productid, Code, ProductName, RetailPrice, DiscPer, DiscPC" & vbCrLf _
         + " from UnionInfoHeader u inner join Products p on u.Unionid = p.productid" & vbCrLf _
         + " left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
         + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
         
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !ProductID
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !RetailPrice
         TxtQty.Text = IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         vStrSQL = " select sum(isnull(Cost,PurPrice)* b.QtyLoose) as Cost from UnionInfoHeader h inner join UnionInfoBody b on h.id = b.id" & vbCrLf _
               + " inner join Products p on p.productid = b.productid" & vbCrLf _
               + " left outer join CurrentStock cs on cs.productid = p.productid " & vbCrLf _
               + " where h.unionid ='" & TxtPID.Text & "'"
         With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
         vStrSQL = " select Floor(min(css.QtyLoose/b.QtyLoose)) as QtyLoose " & vbCrLf _
                  + " from UnionInfoHeader h inner join UnionInfoBody b on h.id = b.id" & vbCrLf _
                  + " inner join Products p on p.productid = b.productid" & vbCrLf _
                  + " left outer join CurrentStockStore css on css.productid = p.productid " & vbCrLf _
                  + " where h.unionid ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text
         With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               LblStock.Caption = IIf(IsNull(!QtyLoose), 0, !QtyLoose)
            Else
               LblStock.Caption = 0
            End If
         End With
         LblStock.Visible = True
         LblStockCaption.Visible = True
         With CN.Execute("select * from registry")
            If .RecordCount > 0 Then
               If !NegativeSale = False Then
                  If LblStock.Caption <= 0 Then
                     MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                     FunSelectProduct = False
                     Exit Function
                  End If
               End If
            End If
            .Close
         End With
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtDiscPC.Text) <> 0 Then
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
         End If
         ChkIsProduct.Value = 0
         SubCalculateBody
         Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      End If
   End With
    
   ''''''''***********   Checking Product  ***********''''''''
    vStrSQL = " SELECT p.productid, code, ProductName, RetailPrice, DiscPer, DiscPC" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"

   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !ProductID
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !RetailPrice
         TxtQty.Text = IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         With CN.Execute("select cost from currentstock where productid ='" & TxtPID.Text & "'")
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
         With CN.Execute("select qtyloose from currentstockStore where productid ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text)
            If .RecordCount > 0 Then
               LblStock.Caption = !QtyLoose
            Else
               LblStock.Caption = 0
            End If
         End With
         LblStock.Visible = True
         LblStockCaption.Visible = True
         With CN.Execute("select * from registry")
         If .RecordCount > 0 Then
            If !NegativeSale = False Then
               If LblStock.Caption <= 0 Then
                  MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                  FunSelectProduct = False
                  Exit Function
               End If
            End If
         End If
         .Close
         End With
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtDiscPC.Text) <> 0 Then
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
         End If
         ChkIsProduct.Value = 1
         SubCalculateBody
         Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtPID.Text = ""
         TxtCode.Text = ""
         TxtProductName.Text = ""
         TxtPrice.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtAmount.Text = ""
         TxtCost.Text = ""
         LblStock.Visible = False
         LblStockCaption.Visible = False
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
   If MsgBox("Are you sure to Clear the Data?", vbQuestion + vbApplicationModal + vbYesNo, "Alert") = vbNo Then Exit Sub
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
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   CN.BeginTrans
   Grid.Redraw = False
   Grid.MoveFirst
   Call ActivityLog("Sale Return Invoice", eDelete, TxtReturnID.Text, DtpReturnDate.DateValue)
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from SaleReturnBody where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "' and productid ='" & Grid.Columns("Productid").Text & "'"
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

Private Sub BtnOpen_Click()
   SchSaleReturn.ParaInReturnDate = DtpReturnDate.DateValue
   SchSaleReturn.Show vbModal
   If SchSaleReturn.ParaOutReturnID <> -1 Then
      TxtReturnID.Text = SchSaleReturn.ParaOutReturnID
      'Dim a
      'a = Split(SchSaleReturn.ParaOutReturnDate, "/")
      DtpReturnDate.DateValue = SchSaleReturn.ParaOutReturnDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      GetSaleReturn
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   'VStrSQL = "select u.username, h.billid, h.BillDate, h.TotalAmount as tbill, isnull(h.discount,0) as discount, isnull(h.cashReceived,0) as cashReceived, p.productname, b.qty, b.price as price, b.amount, b.Disc" _
            + " from saleHeader h inner join salebody b on h.billid = b.billid and h.BillDate = b.BillDate" _
            + " inner join products p on p.productid = b.productid" _
            + " inner join users u on u.UserNo = h.UserNo" _
            + " where h.billid= " & Val(TxtBillID.Text) & " and h.BillDate='" & DtpBillDate.DateValue & "' order by SerialNo"
    
    vStrSQL = " select u.username, h.Returnid as BillID, h.ReturnDate as BillDate, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.CashPaid,0) as cashReceived, p.productname, b.qty, b.price as price, b.amount, b.DiscVal" & vbCrLf _
            + " from SaleReturnHeader h inner join saleReturnbody b on h.Returnid = b.Returnid and h.ReturnDate = b.ReturnDate" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " inner join users u on u.UserNo = h.UserNo" & vbCrLf _
            + " where h.ReturnId= " & Val(TxtReturnID.Text) & " and h.ReturnDate='" & DtpReturnDate.DateValue & "' order by SerialNo"

    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
  
    Set RptReportViewer.Report = New CrpSaleReturnInvoice
    RptReportViewer.Report.DiscardSavedData
    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    vStrComp = "Select CompanyName,Address,City,PhoneNo,email from Company"
    With CN.Execute(vStrComp)
      If .RecordCount > 0 Then
         vCompanyName = !CompanyName
         vAddress = !Address
         vAddress = !Address & IIf(!City = "", "", IIf(!Address = "", "", ", ") & !City)
         vPhone = IIf(!PhoneNo = "", "", !PhoneNo)
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue vAddress
         RptReportViewer.Report.ParameterFields(4).AddCurrentValue vPhone
      End If
   End With
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
   With CN.Execute("select * from registry")
      If .RecordCount > 0 Then
         RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(!AddSpace = True, ".", "")
         'RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(!CashReceived)
      End If
      .Close
   End With
   RptReportViewer.Report.SelectPrinter "RASDD.DLL", "CBM1000 Partial Cut", "Com1" 'RptReportViewer.Report.SelectPrinter  "RASDD.DLL", "CBM1000 Partial Cut", "Com1"
   RptReportViewer.Report.PaperOrientation = crPortrait
   RptReportViewer.Show
   'RptReportViewer.Report.PrintOut False
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtQty.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnReturnAll_Click()
   PopulateSaleDataToGrid
End Sub

Private Sub BtnSale_Click()
   If FunSelectSale(ssButton, False) = True Then
      BtnReturnAll.SetFocus
   Else
      TxtBillID.SetFocus
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
   If FrmReturnPrint.OptCash.Visible Then FrmReturnPrint.OptCash.SetFocus
   FrmReturnPrint.SubClearFields
   FrmReturnPrint.TxtNetAmount.Text = TxtNetAmount.Caption
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
   If DtpReturnDate.Date <> Date And DateFlag = True Then
      If MsgBox("Are you sure to Change Return Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
         DtpReturnDate.DateValue = Date
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
   sSql = "select * from SaleReturnHeader where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, CN, adOpenStatic, adLockOptimistic
      If .BOF Then
         .AddNew
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
      End If
      !BillId = IIf(Val(TxtBillID.Text) = 0, Null, Val(TxtBillID.Text))
      !BillDate = IIf(IsNull(!BillId), Null, DtpBillDate.DateValue)
      !StoreID = TxtStoreID.Text
      !TotalAmount = Round(Val(TxtTotalAmount.Caption))
      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
      !CashPaid = IIf(FrmReturnPrint.TxtCashPaid.Text = "", Null, Val(FrmReturnPrint.TxtCashPaid.Text))
      !CustomerID = IIf(Trim(FrmReturnPrint.TxtCustomerID.Text) = "", "621", FrmReturnPrint.TxtCustomerID.Text)
      !CustomerName = FrmReturnPrint.TxtCashCustomer.Text
      !Cash = FrmReturnPrint.OptCash.Value
      !Credit = FrmReturnPrint.OptCredit.Value
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
   If vIsNewRecord = True Then Call ActivityLog("Sale Return Invoice", eAdd, TxtReturnID.Text, DtpReturnDate.DateValue)
   CN.CommitTrans
   Char.Speak "Thank you for comming"
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
   RsBody.Open "Select * from SaleReturnBody where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = "select p.productname, b.code,b.* from SaleReturnBody b join products p on p.productid = b.productid where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
      With CN.Execute(sSql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         'TxtGrossAmount.Text = 0
         TxtTotalQty.Caption = 0
         'TxtTotalDiscount.Caption = 0
         vTotDisc = 0
         TxtTotalAmount.Caption = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !ProductID
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !Qty
            Grid.Columns("Price").Value = !Price
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("TotalAmount").Value = Val(!Price) * Val(!Qty)
            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
            'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(!DiscVal)
            vTotDisc = vTotDisc + Val(!DiscVal)
            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
            TxtLastRate.Caption = !Price
            .MoveNext
         Wend
         .Close
      End With
      Call SubCalculateBody
      Grid.AddNew
      Grid.Columns("productid").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
End Sub

Private Sub PopulateSaleDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleReturnBody where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
   sSql = "select p.productname, b.code,b.* from SaleBody b join products p on p.productid = b.productid where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   With CN.Execute(sSql)
      If .RecordCount > 0 Then
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         'TxtGrossAmount.Text = 0
         TxtTotalQty.Caption = 0
         'TxtTotalDiscount.Caption = 0
         vTotDisc = 0
         TxtTotalAmount.Caption = 0
         While Not .EOF
            RsBody.AddNew
            RsBody!ProductID = !ProductID
            RsBody!Code = !Code
            RsBody!Qty = !Qty
            RsBody!Price = !Price
            RsBody!DiscPC = !DiscPC
            RsBody!DiscPer = !DiscPer
            RsBody!DiscVal = !DiscVal
            RsBody!Cost = !Cost
            RsBody!IsProduct = !IsProduct
            RsBody!Amount = !Amount
            RsBody.Update
            Grid.AddNew
            Grid.Columns("ProductID").Text = !ProductID
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !Qty
            Grid.Columns("Price").Value = !Price
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("IsProduct").Value = Abs(!IsProduct)
            Grid.Columns("TotalAmount").Value = Val(!Price) * Val(!Qty)
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
            'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(!DiscVal)
            vTotDisc = vTotDisc + Val(!DiscVal)
            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
            TxtLastRate.Caption = !Price
            .MoveNext
         Wend
         .Close
         Call SubCalculateBody
         Grid.AddNew
         Grid.Columns("productid").Text = " "
         Grid.AllowAddNew = False
         Grid.Redraw = True
      End If
   End With
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
      'If RsBody.State = adStateOpen Then RsBody.Close
      TxtBillID.Enabled = True
      DtpBillDate.Enabled = True
      BtnReturnAll.Enabled = True
      BtnSale.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      TxtReturnID.Text = FunGetMaxID()
      'TxtCustomerID.Text = "621"
      'TxtCustomerName.Text = "Counter Sale"
      'DtpReturnDate.DateValue = Date
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      DtpReturnDate.Enabled = True
      'If DtpReturnDate.Enabled And DtpReturnDate.Visible Then DtpReturnDate.SetFocus
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      BtnSale.Enabled = False
      TxtBillID.Enabled = False
      DtpBillDate.Enabled = False
      BtnReturnAll.Enabled = False
      DtpReturnDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      TxtCode.SetFocus
      LblStock.Visible = False
      LblStockCaption.Visible = False
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
      If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub GetSaleReturn()
   On Error GoTo ErrorHandler
   sSql = "select h.*, p.partyname, StoreName FROM SaleReturnHeader h left outer join parties p on h.customerid=p.partyid inner join stores s on s.storeid = h.storeid where h.ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
   With CN.Execute(sSql)
      If Not .BOF Then
          TxtBillID.Text = IIf(IsNull(!BillId), "", !BillId)
          DtpBillDate.DateValue = IIf(IsNull(!BillDate), "", !BillDate)
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
          TxtTotalAmount.Caption = !TotalAmount
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          FrmReturnPrint.TxtCustomerID.Text = IIf(IsNull(!CustomerID), "", !CustomerID)
          FrmReturnPrint.TxtCustomerName.Text = IIf(IsNull(!PartyName), "", !PartyName)
          FrmReturnPrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
          TxtNetAmount.Caption = !TotalAmount
      End If
      .Close
   End With
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleReturnBody where ReturnId=" & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = "select p.productname, b.code,b.* from SaleReturnBody b join products p on p.productid = b.productid where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
      With CN.Execute(sSql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         'TxtGrossAmount.Text = 0
         TxtTotalQty.Caption = 0
         'TxtTotalDiscount.Caption = 0
         vTotDisc = 0
         TxtTotalAmount.Caption = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !ProductID
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !Qty
            'Grid.Columns("QtyOrigional").Value = !Qty
            Grid.Columns("Price").Value = !Price
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("IsProduct").Value = Abs(!IsProduct)
            Grid.Columns("TotalAmount").Value = Val(!Price) * Val(!Qty)
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
            'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(!DiscVal)
            vTotDisc = vTotDisc + Val(!DiscVal)
            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
            TxtLastRate.Caption = !Price
            .MoveNext
         Wend
         .Close
      End With
      Call SubCalculateBody
      Grid.AddNew
      Grid.Columns("ProductID").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub DtpReturnDate_Validate(Cancel As Boolean)
   TxtReturnID.Text = FunGetMaxID()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn And Shift = vbShiftMask Then
      Select Case ActiveControl.Name
      Case TxtCode.Name
         If FunSelectProduct(ssValidate, False) = True Then TxtQty.SetFocus
      Case TxtQty.Name
         TxtDiscPC.SetFocus
      Case TxtDiscPC.Name
         TxtDiscPer.SetFocus
      End Select
      KeyCode = 0
      Shift = 0
   ElseIf KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name
         Grid_DblClick
      Case TxtCode.Name
         FunSelectProduct ssValidate, False
         GetDataFromTexBoxesToGrid
      Case TxtQty.Name, TxtDiscPC.Name, TxtDiscPer.Name
         GetDataFromTexBoxesToGrid
      Case Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End Select
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
         Case TxtBillID.Name: If FunSelectSale(ssFunctionKey, False) = True Then BtnReturnAll.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtCode.Enabled Then TxtCode.SetFocus
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtQty.SetFocus
      End Select
   'ElseIf KeyCode = vbKeyF9 Then
  '       Call abc
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
         KeyCode = 0
         TxtBillDisc.SetFocus
      End If
   ElseIf ActiveControl.Name = Grid.Name And (Shift = vbCtrlMask + vbShiftMask + vbAltMask) And KeyCode = 46 Then  ' 46 is del button
      If Trim(Grid.Columns("ProductID").Text <> "") Then Call mniRemoveRow_Click
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
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hwnd, "Sale Return Invoice"
   DtpReturnDate.DateValue = Date
   With CN.Execute("select * from registry")
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         If !ChangePrice = True Then
            If ObjUserSecurity.IsAdministrator = True Then TxtPrice.Enabled = True
         End If
      End If
      .Close
   End With
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
   TxtLastRate.Caption = 0
   TxtTotalQty.Caption = 0
   TxtTotalDiscount.Caption = 0
   TxtTotalAmount.Caption = 0
   TxtNetAmount.Caption = 0
   vTotDisc = 0
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ProductID").Text = " "
   Grid.Update
   Unload FrmPrint
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
    'CN.Execute ("exec spcurrentstock")
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set RsBody = Nothing
    Set RsReport = Nothing
    Set FrmSaleInvoice = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   'TxtGrossAmount.Text = Val(TxtGrossAmount.Text) - Grid.Columns("Amount").Value
   TxtTotalQty.Caption = Val(TxtTotalQty.Caption) - Grid.Columns("Qty").Value
   vTotDisc = vTotDisc - Grid.Columns("DiscVal").Value
   TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) - Grid.Columns("TotalAmount").Value
   SubCalculateFooter
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
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      TxtCode.Text = ""
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      TxtCode.SetFocus
   Else
      TxtCode.Enabled = False
      BtnProduct.Enabled = False
      If TxtQty.Enabled = True And TxtQty.Visible Then TxtQty.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(Grid.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
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
   RsBody.Filter = "Code='" & TxtCode.Text & "'"
   If RsBody.RecordCount > 0 Then RsBody.Delete
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
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
      TxtCode.SetFocus
      Exit Sub
   End If
   If Val(TxtQty.Text) = 0 Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      TxtQty.SetFocus
      Exit Sub
   End If
On Error GoTo ErrorHandler
   RsBody.Filter = "ProductID='" & TxtPID.Text & "'"
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ProductID").Text = TxtPID.Text
         Grid.Columns("Code").Text = TxtCode.Text
         RsBody!ProductID = TxtPID.Text
         RsBody!Code = TxtCode.Text
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("Productid").Text = TxtPID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  TxtQty.Text = Val(TxtQty.Text) + Grid.Columns("Qty").Value
                  TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text) - Val(Grid.Columns("Qty").Text)
                  vTotDisc = vTotDisc + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
                  'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
                  TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text) - Val(Grid.Columns("TotalAmount").Text)
                  TxtLastRate.Caption = Val(TxtPrice.Text) - Val(TxtDiscPC.Text)
                  TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("Qty").Value = Val(TxtQty.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("DiscPC").Value = Val(TxtDiscPC.Text)
                  Grid.Columns("DiscPer").Value = Val(TxtDiscPer.Text)
                  Grid.Columns("DiscVal").Value = Val(TxtDiscVal.Text)
                  Grid.Columns("Cost").Value = Val(TxtCost.Text)
                  Grid.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
                  RsBody!Qty = Val(TxtQty.Text)
                  RsBody!Price = Val(TxtPrice.Text)
                  RsBody!DiscPC = Val(TxtDiscPC.Text)
                  RsBody!DiscPer = Val(TxtDiscPer.Text)
                  RsBody!DiscVal = Val(TxtDiscVal.Text)
                  RsBody!Cost = Val(TxtCost.Text)
                  RsBody!IsProduct = Abs(ChkIsProduct.Value)
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
         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text)
         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text)
         'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(TxtDiscVal.Text)
         vTotDisc = vTotDisc + Val(TxtDiscVal.Text)
         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text)
      Else
         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text) - Val(.Columns("Qty").Text)
         vTotDisc = vTotDisc + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
         'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text) - Val(Grid.Columns("TotalAmount").Text)
      End If
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("Qty").Value = Val(TxtQty.Text)
      .Columns("Price").Value = Val(TxtPrice.Text)
      .Columns("DiscPC").Value = Val(TxtDiscPC.Text)
      .Columns("DiscPer").Value = Val(TxtDiscPer.Text)
      .Columns("DiscVal").Value = Val(TxtDiscVal.Text)
      If Trim(TxtCost.Text) <> "" Then
         .Columns("Cost").Value = Val(TxtCost.Text)
      End If
      .Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
      TxtLastRate.Caption = Val(TxtPrice.Text) - Val(TxtDiscPC.Text)
      RsBody!Qty = Val(TxtQty.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!DiscPC = Val(TxtDiscPC.Text)
      RsBody!DiscPer = Val(TxtDiscPer.Text)
      RsBody!DiscVal = Val(TxtDiscVal.Text)
      If Trim(TxtCost.Text) <> "" Then
         RsBody!Cost = Val(TxtCost.Text)
      End If
      If IsNull(RsBody!Cost) Then RsBody!Cost = 0
      RsBody!Amount = Val(TxtAmount.Text)
      RsBody!IsProduct = Abs(ChkIsProduct.Value)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
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
   TxtQty.Text = ""
   TxtPrice.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtDiscVal.Text = ""
   TxtAmount.Text = ""
   TxtActualAmount.Text = ""
   ChkIsProduct.Value = 1
   TxtCost.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtPID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("code").Text
      TxtProductName.Text = .Columns("ProductName").Text
      TxtQty.Text = .Columns("Qty").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      TxtCost.Text = .Columns("Cost").Value
      TxtAmount.Text = .Columns("Amount").Text
      TxtActualAmount.Text = .Columns("TotalAmount").Text
      ChkIsProduct.Value = Abs(.Columns("IsProduct").Value)
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub GetSaleReturn()
'   On Error GoTo ErrorHandler
'   sSql = "select h.*, p.partyname, StoreName FROM SaleReturnHeader h left outer join parties p on h.customerid=p.partyid inner join stores s on s.storeid = h.storeid where h.ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
'   With CN.Execute(sSql)
'      If Not .BOF Then
'          TxtBillID.Text = IIf(IsNull(!BillId), "", !BillId)
'          If Not IsNull(!BillDate) Then
'             DtpBillDate.DateValue = !BillDate
'          End If
'          DtpBillDate.DateValue = IIf(IsNull(!BillDate), Null, !BillDate)
'          TxtStoreID.Text = !StoreID
'          TxtStoreName.Text = !StoreName
'          TxtTotalAmount.Caption = !TotalAmount
'          TxtBillDisc.Text = IIf(IsNull(!BillDiscount), "", !BillDiscount)
'          FrmReturnPrint.OptCash.Value = !Cash
'          FrmReturnPrint.OptCredit.Value = !Credit
'          FrmReturnPrint.TxtCashPaid.Text = IIf(IsNull(!CashPaid), "", !CashPaid)
'          FrmReturnPrint.TxtCustomerID.Text = IIf(IsNull(!CustomerID), "", !CustomerID)
'          FrmReturnPrint.TxtCustomerName.Text = IIf(IsNull(!PartyName), "", !PartyName)
'          FrmReturnPrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
'          TxtNetAmount.Caption = !TotalAmount
'      End If
'      .Close
'   End With
'   Call PopulateDataToGrid
'   FormStatus = OpenMode
'   Exit Sub
'ErrorHandler:
'   Grid.Redraw = True
'   Call ShowErrorMessage
'End Sub

Private Sub TxtBillDisc_Change()
   If ActiveControl.Name <> TxtBillDisc.Name Then Exit Sub
   TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / Val(TxtTotalAmount.Caption), 2)
   Call SubCalculateFooter
End Sub

Private Sub TxtBillDiscPer_Change()
   If ActiveControl.Name <> TxtBillDiscPer.Name Then Exit Sub
   TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtBillDiscPer.Text) / 100))
   Call SubCalculateFooter
End Sub

Private Sub TxtBillID_Change()
   If Trim(TxtBillID.Text) = "" Then DtpBillDate.Enabled = False Else DtpBillDate.Enabled = True
End Sub

Private Sub TxtDiscPC_Change()
   If ActiveControl.Name <> TxtDiscPC.Name Then Exit Sub
   If Val(TxtPrice.Text) = 0 Then Exit Sub
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
   Call SubCalculateBody
End Sub

'Private Sub TxtDiscPC_LostFocus()
'   Select Case Me.ActiveControl.Name
'   Case TxtCode.Name, TxtQty.Name, TxtDiscPC.Name
'      Exit Sub
'   End Select
'   Call GetDataFromTexBoxesToGrid
'End Sub

Private Sub TxtDiscPer_Change()
   If ActiveControl.Name <> TxtDiscPer.Name Then Exit Sub
   TxtDiscPC.Text = Round((Val(TxtPrice.Text) * Val(TxtDiscPer.Text) / 100), 2)
   Call SubCalculateBody
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtCode.Text = ""
      TxtPID.Text = ""
      TxtProductName.Text = ""
      TxtPrice.Text = ""
      TxtDiscPC.Text = ""
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
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = Not FunSelectProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtPrice_Change()
   Call SubCalculateBody
End Sub

Private Sub TxtQty_Change()
   Call SubCalculateBody
End Sub

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
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

