VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmReplacementInvoice1 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11970
   ControlBox      =   0   'False
   Icon            =   "FrmReplacementInvoice1.frx":0000
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
      Left            =   5730
      TabIndex        =   51
      Top             =   300
      Visible         =   0   'False
      Width           =   1365
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   165
      TabIndex        =   16
      Top             =   1125
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
   Begin SITextBox.Txt TxtSDiscVal 
      Height          =   315
      Left            =   9135
      TabIndex        =   7
      Top             =   1845
      Width           =   990
      _ExtentX        =   1746
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
   Begin SITextBox.Txt TxtSCode 
      Height          =   315
      Left            =   165
      TabIndex        =   2
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
   Begin SITextBox.Txt TxtSQty 
      Height          =   315
      Left            =   5880
      TabIndex        =   3
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
   Begin SITextBox.Txt TxtSPrice 
      Height          =   315
      Left            =   6660
      TabIndex        =   4
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
   Begin SITextBox.Txt TxtSAmount 
      Height          =   315
      Left            =   10125
      TabIndex        =   17
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
   Begin JeweledBut.JeweledButton BtnSProduct 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2025
      TabIndex        =   18
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
      MICON           =   "FrmReplacementInvoice1.frx":0ECA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7335
      TabIndex        =   14
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
      MICON           =   "FrmReplacementInvoice1.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6015
      TabIndex        =   10
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
      MICON           =   "FrmReplacementInvoice1.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3375
      TabIndex        =   12
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
      MICON           =   "FrmReplacementInvoice1.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8655
      TabIndex        =   15
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
      MICON           =   "FrmReplacementInvoice1.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4695
      TabIndex        =   11
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
      MICON           =   "FrmReplacementInvoice1.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSBillDisc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   150
      TabIndex        =   8
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
   Begin SITextBox.Txt TxtSProductName 
      Height          =   315
      Left            =   2385
      TabIndex        =   28
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridSale 
      CausesValidation=   0   'False
      Height          =   3675
      Left            =   165
      TabIndex        =   29
      Top             =   2160
      Width           =   11625
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   14
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
      stylesets(0).Picture=   "FrmReplacementInvoice1.frx":0F72
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
      ExtraHeight     =   238
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   14
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
      Columns(12).Caption=   "QtyOrigional"
      Columns(12).Name=   "QtyOrigional"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   4
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "IsProduct"
      Columns(13).Name=   "IsProduct"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   11
      Columns(13).FieldLen=   256
      Columns(13).Style=   2
      TabNavigation   =   1
      _ExtentX        =   20505
      _ExtentY        =   6482
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
      TabIndex        =   13
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
      MICON           =   "FrmReplacementInvoice1.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSDiscPC 
      Height          =   315
      Left            =   7620
      TabIndex        =   5
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
      Left            =   7980
      TabIndex        =   35
      Top             =   285
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
      Left            =   5040
      TabIndex        =   1
      Tag             =   "NC"
      Top             =   1125
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
      Left            =   6075
      TabIndex        =   37
      Tag             =   "NC"
      Top             =   1125
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
      Left            =   5715
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1125
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
      MICON           =   "FrmReplacementInvoice1.frx":0FAA
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpBillDate 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   1125
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
   Begin SITextBox.Txt TxtSDiscPer 
      Height          =   315
      Left            =   8445
      TabIndex        =   6
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
      Left            =   9585
      TabIndex        =   42
      Top             =   285
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
      Left            =   7155
      TabIndex        =   45
      Top             =   270
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
   Begin SITextBox.Txt TxtSBillDiscPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   150
      TabIndex        =   9
      Top             =   7095
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
      Left            =   3795
      TabIndex        =   57
      Top             =   6570
      Width           =   1380
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
      Left            =   5205
      TabIndex        =   56
      Top             =   6570
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
      Left            =   7605
      TabIndex        =   55
      Top             =   6570
      Width           =   1740
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
      Left            =   9375
      TabIndex        =   54
      Top             =   6570
      Width           =   2370
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
      Left            =   1980
      TabIndex        =   53
      Top             =   6570
      Width           =   1785
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Discount (%)"
      Height          =   195
      Left            =   150
      TabIndex        =   52
      Top             =   6870
      Width           =   1125
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replacement Invoice"
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
      Left            =   1920
      TabIndex        =   50
      Top             =   180
      Width           =   3570
   End
   Begin VB.Label LblCost 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   1140
      TabIndex        =   49
      Top             =   1590
      Visible         =   0   'False
      Width           =   855
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
      Left            =   8550
      TabIndex        =   48
      Top             =   765
      Width           =   720
   End
   Begin VB.Label LblStock 
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
      Left            =   8595
      TabIndex        =   47
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   7155
      TabIndex        =   46
      Top             =   45
      Visible         =   0   'False
      Width           =   315
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
      TabIndex        =   44
      Top             =   6240
      Width           =   1830
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   9585
      TabIndex        =   43
      Top             =   90
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
      TabIndex        =   41
      Top             =   1650
      Width           =   525
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   5040
      TabIndex        =   40
      Top             =   930
      Width           =   585
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   6075
      TabIndex        =   39
      Top             =   930
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Amount"
      Height          =   195
      Left            =   7965
      TabIndex        =   36
      Top             =   75
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   27
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
      TabIndex        =   26
      Top             =   6165
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   1440
      TabIndex        =   25
      Top             =   930
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " Bill ID"
      Height          =   195
      Left            =   165
      TabIndex        =   24
      Top             =   930
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   165
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   1650
      Width           =   630
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
      Begin VB.Menu MniCostPrice 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmReplacementInvoice1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim vMode As FormMode
'Dim vCounter As Integer
'Dim RsBody As New ADODB.Recordset
'Dim RsReport As New ADODB.Recordset
'Dim vIsNewRecord As Boolean
'Dim Flag As Boolean
'Dim DateFlag As Boolean
'Dim sSql As String
'Dim vStrSQL As String
'Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String, vTotDisc As Double
''----------------------------------
'
'Private Sub SubCalculateBody()
'    TxtSDiscVal.Text = Val(TxtSQty.Text) * Val(TxtSDiscPC.Text)
'    TxtActualAmount.Text = Val(TxtSQty.Text) * Val(TxtPrice.Text)
'    TxtSAmount.Text = Val(TxtActualAmount.Text) - Val(TxtSDiscVal.Text)
'    TxtSTotalDiscount.Caption = vTotDisc
'    SubCalculateFooter
'End Sub
'
'Private Sub SubCalculateFooter()
'   TxtSTotalDiscount.Caption = Val(TxtBillDisc.Text) + vTotDisc
'   TxtNetAmount.Caption = SelfRound(Val(TxtTotalAmount.Caption) - Val(TxtSTotalDiscount.Caption))
'   'If TxtGrossAmount.Text = "" Then Exit Sub
'   'TxtNetAmount.Caption = Round(Val(TxtGrossAmount.Text)) - Val(TxtBillDisc.Text)
'   'TxtCashReturn.Text = IIf(Val(TxtCashReceived.Text) > 0, Val(TxtCashReceived.Text) - Val(TxtNetAmount.Caption), "")
'End Sub
'
'Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
'    On Error GoTo ErrorHandler
'    Dim vStrSQL As String
'    If CallerName = ssButton Or CallerName = ssFunctionKey Then
'        SchStore.Show vbModal, Me
'        If SchStore.ParaOutStoreID = "" Then FunSelectStore = False: Exit Function
'        TxtStoreID.Text = SchStore.ParaOutStoreID
'    End If
'    '---------------------------
'    vStrSQL = " Select * FROM Stores where StoreID=" & Val(TxtStoreID.Text)
'    With CN.Execute(vStrSQL)
'      If .RecordCount > 0 Then
'          TxtStoreName.Text = !StoreName
'          FunSelectStore = True
'          .Close
'          If BtnSave.Enabled = False Then FormStatus = ChangeMode
'          Exit Function
'      Else
'          FunSelectStore = False
'          .Close
'          TxtStoreID.Text = ""
'          TxtStoreName.Text = ""
'          If BtnSave.Enabled = False Then FormStatus = ChangeMode
'      End If
'   End With
'   Exit Function
'ErrorHandler:
'   Call ShowErrorMessage
'End Function
'
'Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
'   On Error GoTo ErrorHandler
'   Dim vStrSQL As String
'   If CallerName = ssButton Or CallerName = ssFunctionKey Then
'      SchProduct.Show vbModal, Me
'      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
'      TxtCode.Text = SchProduct.ParaOutID
'   End If
'    '---------------------------
'    If TxtCode.Enabled = False Then FunSelectProduct = False: Exit Function
'    If Trim(TxtCode.Text) = "" Then FunSelectProduct = False: Exit Function
'    If Len(TxtCode.Text) < 5 Then
'      TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
'    End If
'    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
'
'    ''''''''***********   Checking Union   ***********''''''''
'    vStrSQL = " SELECT p.productid, Code, ProductName, RetailPrice, DiscPer, DiscPC" & vbCrLf _
'         + " from UnionInfoHeader u inner join Products p on u.Unionid = p.productid" & vbCrLf _
'         + " left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
'         + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
'
'   With CN.Execute(vStrSQL)
'      If .RecordCount > 0 Then
'         TxtPID.Text = !ProductID
'         TxtProductName.Text = !ProductName
'         If FrmPrint.TxtCustomerID.Text = "" Then
'            TxtPrice.Text = !RetailPrice
'         Else
'            TxtPrice.Text = CN.Execute("Select dbo.FunLastPrice('S','" & DtpBillDate.DateValue & "','" & TxtPID.Text & "','" & FrmPrint.TxtCustomerID.Text & "')").Fields(0).Value
'         End If
'         TxtSQty.Text = IIf(Val(TxtSQty.Text) = 0, 1, TxtSQty.Text)
'         vStrSQL = " select sum(isnull(Cost,PurPrice)* b.QtyLoose) as Cost from UnionInfoHeader h inner join UnionInfoBody b on h.id = b.id" & vbCrLf _
'               + " inner join Products p on p.productid = b.productid" & vbCrLf _
'               + " left outer join CurrentStock cs on cs.productid = p.productid " & vbCrLf _
'               + " where h.unionid ='" & TxtPID.Text & "'"
'         With CN.Execute(vStrSQL)
'            If .RecordCount > 0 Then
'               TxtCost.Text = !Cost
'            Else
'               TxtCost.Text = "0"
'            End If
'         End With
'         vStrSQL = " select Floor(min(css.QtyLoose/b.QtyLoose)) as QtyLoose " & vbCrLf _
'                  + " from UnionInfoHeader h inner join UnionInfoBody b on h.id = b.id" & vbCrLf _
'                  + " inner join Products p on p.productid = b.productid" & vbCrLf _
'                  + " left outer join CurrentStockStore css on css.productid = p.productid " & vbCrLf _
'                  + " where h.unionid ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text
'         With CN.Execute(vStrSQL)
'            If .RecordCount > 0 Then
'               LblStock.Caption = IIf(IsNull(!QtyLoose), 0, !QtyLoose)
'            Else
'               LblStock.Caption = 0
'            End If
'         End With
'         LblStock.Visible = True
'         LblStockCaption.Visible = True
'         With CN.Execute("select * from registry")
'            If .RecordCount > 0 Then
'               If !NegativeSale = False Then
'                  If LblStock.Caption <= 0 Then
'                     MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                     FunSelectProduct = False
'                     Exit Function
'                  End If
'               End If
'            End If
'            .Close
'         End With
'         TxtSDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
'         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
'         If Val(TxtSDiscPC.Text) <> 0 Then
'            TxtDiscPer.Text = Round((Val(TxtSDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
'         End If
'         ChkIsProduct.Value = 0
'         SubCalculateBody
'         Char.Speak TxtProductName.Text
'         FunSelectProduct = True
'         If BtnSave.Enabled = False Then FormStatus = ChangeMode
'         .Close
'         Exit Function
'      End If
'   End With
'
'''''''''***********   Checking Product  ***********''''''''
'    vStrSQL = " SELECT p.productid, code, ProductName, RetailPrice, DiscPer, DiscPC" & vbCrLf _
'           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
'           + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
'
'   With CN.Execute(vStrSQL)
'      If .RecordCount > 0 Then
'         TxtPID.Text = !ProductID
'         TxtProductName.Text = !ProductName
'         If FrmPrint.TxtCustomerID.Text = "" Then
'            TxtPrice.Text = !RetailPrice
'         Else
'            TxtPrice.Text = CN.Execute("Select dbo.FunLastPrice('S','" & DtpBillDate.DateValue & "','" & TxtPID.Text & "','" & FrmPrint.TxtCustomerID.Text & "')").Fields(0).Value
'         End If
'         TxtSQty.Text = IIf(Val(TxtSQty.Text) = 0, 1, TxtSQty.Text)
'         With CN.Execute("select cost from currentstock where productid ='" & TxtPID.Text & "'")
'            If .RecordCount > 0 Then
'               TxtCost.Text = !Cost
'            Else
'               TxtCost.Text = "0"
'            End If
'         End With
'         With CN.Execute("select qtyloose from currentstockStore where productid ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text)
'            If .RecordCount > 0 Then
'               LblStock.Caption = !QtyLoose
'            Else
'               LblStock.Caption = 0
'            End If
'         End With
'         LblStock.Visible = True
'         LblStockCaption.Visible = True
'         With CN.Execute("select * from registry")
'         If .RecordCount > 0 Then
'            If !NegativeSale = False Then
'               If LblStock.Caption <= 0 Then
'                  MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                  FunSelectProduct = False
'                  Exit Function
'               End If
'            End If
'         End If
'         .Close
'         End With
'         TxtSDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
'         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
'         If Val(TxtSDiscPC.Text) <> 0 Then
'            TxtDiscPer.Text = Round((Val(TxtSDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
'         End If
'         ChkIsProduct.Value = 1
'         SubCalculateBody
'         Char.Speak TxtProductName.Text
'         FunSelectProduct = True
'         If BtnSave.Enabled = False Then FormStatus = ChangeMode
'         .Close
'         Exit Function
'      Else
'         FunSelectProduct = False
'         .Close
'         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
'         TxtPID.Text = ""
'         TxtCode.Text = ""
'         TxtProductName.Text = ""
'         TxtPrice.Text = ""
'         TxtSDiscPC.Text = ""
'         TxtDiscPer.Text = ""
'         TxtSAmount.Text = ""
'         TxtCost.Text = ""
'         LblStock.Visible = False
'         LblStockCaption.Visible = False
'         If BtnSave.Enabled = False Then FormStatus = ChangeMode
'         Exit Function
'      End If
'   End With
'Exit Function
'ErrorHandler:
'   Call ShowErrorMessage
'End Function
'
'Private Sub BtnClear_Click()
'   On Error GoTo ErrorHandler
'   If MsgBox("Are you sure to Clear the Data?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Exit Sub
'   FormStatus = NewMode
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Sub BtnClose_Click()
'   Unload Me
'End Sub
'
'Private Sub BtnDelete_Click()
'   On Error GoTo ErrorHandler
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
'      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
'   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
'   CN.BeginTrans
'   Grid.Redraw = False
'   Grid.MoveFirst
'   Call ActivityLog("Sale Invoice", eDelete, TxtBillID.Text, DtpBillDate.DateValue)
'   For vCounter = 1 To Grid.Rows
'      If Trim(Grid.Columns("ProductID").Text) <> "" Then
'         CN.Execute "Delete from SaleBody where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "' and productid ='" & Grid.Columns("Productid").Text & "'"
'      End If
'      Grid.MoveNext
'   Next vCounter
'   Grid.RemoveAll
'   Grid.Redraw = True
'   CN.Execute "Delete from SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
'   CN.CommitTrans
'   FormStatus = NewMode
'   Exit Sub
'ErrorHandler:
'   Grid.Redraw = True
'   If CN.Errors.Count > 0 Then CN.RollbackTrans
'   Call ShowErrorMessage
'End Sub
'
'Private Sub BtnOpen_Click()
'   On Error Resume Next
'   SchSale.ParaInBillDate = DtpBillDate.DateValue
'   SchSale.Show vbModal
'   If SchSale.ParaOutBillID <> -1 Then
'      TxtBillID.Text = SchSale.ParaOutBillID
'      'Dim a
'      'a = Split(SchSale.ParaOutBillDate, "/")
'      DtpBillDate.DateValue = SchSale.ParaOutBillDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
'      GetSale
'   End If
'End Sub
'
'Private Sub BtnPrint_Click()
'   On Error GoTo ErrorHandler
'
'    vStrSQL = " select u.username, h.billid, h.BillDate, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.cashReceived,0) as cashReceived, p.productname, b.qty, b.price as price, b.amount, b.DiscVal, InvoiceNo" & vbCrLf _
'            + " , Case when CustomerID = '621' then isnull(CustomerName,PartyName) Else PartyName End as Customer, Cash, Credit, BankCard" & vbCrLf _
'            + " from saleHeader h inner join salebody b on h.billid = b.billid and h.BillDate = b.BillDate" & vbCrLf _
'            + " inner join products p on p.productid = b.productid" & vbCrLf _
'            + " inner join users u on u.UserNo = h.UserNo" & vbCrLf _
'            + " left outer join Parties Pr on Pr.PartyID = h.CustomerID" & vbCrLf _
'            + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate='" & DtpBillDate.DateValue & "' Order By SerialNo"
'
'    If RsReport.State = adStateOpen Then RsReport.Close
'    RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
'
'   If InStr(1, Printer.DeviceName, "CBM1000") > 0 Then
'      Set RptReportViewer.Report = New CrpSaleInvoiceCBM
'   ElseIf InStr(1, Printer.DeviceName, "Generic") > 0 Then
'      Set RptReportViewer.Report = New CrpSaleInvoiceGeneric
'      RptReportViewer.Report.PaperSize = crPaperEnvelope14
'   Else 'If InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
'      Set RptReportViewer.Report = New CrpSaleInvoiceAurora
'   End If
'    RptReportViewer.Report.DiscardSavedData
'    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
'    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
'    'RptReportViewer.Report.LeftMargin = 0
'    'RptReportViewer.Report.RightMargin = 0
'    vStrComp = "Select CompanyName,Address,City,PhoneNo,email from Company"
'    With CN.Execute(vStrComp)
'      If .RecordCount > 0 Then
'         vCompanyName = !CompanyName
'         vAddress = !Address
'         vAddress = !Address & IIf(!City = "", "", IIf(!Address = "", "", ", ") & !City)
'         vPhone = IIf(!PhoneNo = "", "", !PhoneNo)
'         RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
'         RptReportViewer.Report.ParameterFields(2).AddCurrentValue vAddress
'         RptReportViewer.Report.ParameterFields(4).AddCurrentValue vPhone
'      End If
'   End With
'   RptReportViewer.Report.ParameterFields(3).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
'   With CN.Execute("select * from registry")
'      If .RecordCount > 0 Then
'         RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(!AddSpace = True, ".", "")
'         RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(!CashReceived)
'         RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(!Statement)
'      End If
'      .Close
'   End With
'   'RptReportViewer.Report.SelectPrinter "RASDD.DLL", "CBM1000 Partial Cut", "Com1" 'RptReportViewer.Report.SelectPrinter  "RASDD.DLL", "CBM1000 Partial Cut", "Com1"
'   RptReportViewer.Report.PaperOrientation = crPortrait
'   RptReportViewer.Show
'   'RptReportViewer.Report.PrintOut False
'Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Sub BtnProduct_Click()
'   If FunSelectProduct(ssButton, True) = True Then
'      TxtSQty.SetFocus
'   Else
'      TxtCode.SetFocus
'   End If
'End Sub
'
'Private Sub BtnSave_Click()
'   On Error GoTo ErrorHandler
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
'      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
''  Header Validation
'   If Trim(TxtStoreID.Text) = "" Then
'      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
'      TxtStoreID.SetFocus
'      Exit Sub
'   End If
''   If DtpBillDate.Enabled = True Then
''      If FrmPrint.OptCash.Visible Then FrmPrint.OptCash.SetFocus
''      FrmPrint.SubClearFields
''   End If
'   FrmPrint.TxtNetAmount.Text = TxtNetAmount.Caption
'   With CN.Execute("select * from registry")
'      If .RecordCount > 0 Then
'         If !CashReceived = False Then
'            FrmPrint.TxtCashReceivedCash.Text = TxtNetAmount.Caption
'         End If
'      End If
'      .Close
'   End With
'   FrmPrint.ParaInPrint = True
'   FrmPrint.ParaInChoice = "Cash"
'   FrmPrint.Show vbModal, Me
'   If FrmPrint.ParaOutSelection = False Then Exit Sub
'   If DtpBillDate.Enabled And DtpBillDate.Date <> Date And DateFlag = True Then
'      If MsgBox("Are you sure to Change Bill Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
'         DtpBillDate.DateValue = Date
'         TxtBillID.Text = FunGetMaxID()
'      End If
'      DateFlag = False
'   End If
'   If DtpBillDate.Enabled Then
'      If CN.Execute("Select * from SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'").RecordCount > 0 Then
'         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
'         TxtBillID.Text = FunGetMaxID
'         'Exit Sub
'      End If
'   End If
'   RsBody.Filter = 0
'   If RsBody.RecordCount = 0 Then
'      MsgBox "Please enter at least one product to sale", vbExclamation, "Alert"
'      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
'      Exit Sub
'   End If
'  'Body Validation
'  ' validation has been performed when a row is added to the grid
'
'  'Saving record
'    CN.BeginTrans
'    If vIsNewRecord = False Then Call ActivityLog("Sale Invoice", eEdit, TxtBillID.Text, DtpBillDate.DateValue)
'    sSql = "select * from SaleHeader where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
'    Dim Rs As New ADODB.Recordset
'    With Rs
'      .Open sSql, CN, adOpenStatic, adLockOptimistic
'      If .BOF Then
'         .AddNew
'         !BillId = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'      End If
'      !StoreID = TxtStoreID.Text
'      !TotalAmount = Round(Val(TxtTotalAmount.Caption))
'      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
'      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
'      If FrmPrint.OptBankCard.Value = True Then
'         !InvoiceNo = FrmPrint.TxtInvoiceNo.Text
'         !Commision = FrmPrint.TxtCommision.Text
'         !BankMachineID = FrmPrint.TxtBankMachineID.Text
'         !CashReceived = Null
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(FrmPrint.TxtBankCustomer.Text) = "", Null, FrmPrint.TxtBankCustomer.Text)
'      End If
'      If FrmPrint.OptCash.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !BankMachineID = Null
'         !CashReceived = Val(FrmPrint.TxtCashReceivedCash.Text)
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(FrmPrint.TxtCashCustomer.Text) = "", Null, FrmPrint.TxtCashCustomer.Text)
'      End If
'      If FrmPrint.OptCredit.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !BankMachineID = Null
'         !CashReceived = Val(FrmPrint.TxtCashReceivedCredit.Text)
'         !CustomerID = FrmPrint.TxtCustomerID.Text
'         !CustomerName = Null
'      End If
'      !BankCard = FrmPrint.OptBankCard.Value
'      !Cash = FrmPrint.OptCash.Value
'      !Credit = FrmPrint.OptCredit.Value
'      !UserNo = vUser
'      .Update
'      .Close
'   End With
'   With RsBody
'      .Filter = 0
'      .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !BillId = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         .MoveNext
'      Next vCounter
'      .UpdateBatch
'   End With
'   If vIsNewRecord = True Then Call ActivityLog("Sale Invoice", eAdd, TxtBillID.Text, DtpBillDate.DateValue)
'   CN.CommitTrans
'   Char.Speak "Thank you for comming"
'   'If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
'   If FrmPrint.ChkPrint.Value = 1 Then Call BtnPrint_Click
'   'End If
'   FormStatus = NewMode
'   Exit Sub
'ErrorHandler:
'   Grid.Redraw = True
'   If CN.Errors.Count > 0 Then CN.RollbackTrans
'   Call ShowErrorMessage
'End Sub
'
'Private Sub PopulateDataToGrid()
'   RsBody.Filter = 0
'   If RsBody.State = adStateOpen Then RsBody.Close
'   RsBody.Open "Select * from SaleBody where BillId=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
'      sSql = "select p.productname, b.code,b.* from salebody b join products p on p.productid = b.productid where billid=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
'      With CN.Execute(sSql)
'         Grid.Redraw = False
'         Grid.MoveFirst
'         Grid.RemoveAll
'         Grid.AllowAddNew = True
'         'TxtGrossAmount.Text = 0
'         TxtTotalQty.Caption = 0
'         'TxtSTotalDiscount.Caption = 0
'         vTotDisc = 0
'         TxtTotalAmount.Caption = 0
'         While Not .EOF
'            Grid.AddNew
'            Grid.Columns("ProductID").Text = !ProductID
'            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
'            Grid.Columns("ProductName").Text = !ProductName
'            Grid.Columns("Qty").Value = !Qty
'            Grid.Columns("QtyOrigional").Value = !Qty
'            Grid.Columns("Price").Value = !Price
'            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
'            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
'            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
'            Grid.Columns("Amount").Value = !Amount
'            Grid.Columns("IsProduct").Value = Abs(!IsProduct)
'            Grid.Columns("TotalAmount").Value = Val(!Price) * Val(!Qty)
'            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
'            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
'            'TxtSTotalDiscount.Caption = Val(TxtSTotalDiscount.Caption) + Val(!DiscVal)
'            vTotDisc = vTotDisc + Val(!DiscVal)
'            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
'            TxtLastRate.Caption = !Price
'            .MoveNext
'         Wend
'         .Close
'      End With
'      Call SubCalculateBody
'      Grid.AddNew
'      Grid.Columns("ProductID").Text = " "
'      Grid.AllowAddNew = False
'      Grid.Redraw = True
'   End If
'End Sub
'
'Private Property Get FormStatus() As FormMode
'  'Nothing
'  FormStatus = vMode
'End Property
'
'Private Property Let FormStatus(ByVal vNewValue As FormMode)
'   'Based upon the value of vNewValue, we shall decide what controls to enable/disable
'   On Error GoTo ErrorHandler
'   vMode = vNewValue
'   Select Case vNewValue
'   Case Is = NewMode
'      Call SubClearFields
'      BtnOpen.Enabled = True
'      BtnDelete.Enabled = False
'      BtnSave.Enabled = False
'      BtnClear.Enabled = True
'      BtnPrint.Enabled = False
'      TxtBillID.Text = FunGetMaxID()
'      Call PopulateDataToGrid
'      'TxtCustomerID.Text = "621"
'      'TxtCustomerName.Text = "Counter Sale"
'      'DtpBillDate.DateValue = Date
'      LblStock.Visible = False
'      LblStockCaption.Visible = False
'      TxtCode.Enabled = True
'      BtnProduct.Enabled = True
'      DtpBillDate.Enabled = True
'      'If DtpBillDate.Enabled And DtpBillDate.Visible Then DtpBillDate.SetFocus
'      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
'      vIsNewRecord = True
'   Case Is = OpenMode
'      DtpBillDate.Enabled = False
'      BtnOpen.Enabled = True
'      BtnDelete.Enabled = True
'      BtnClear.Enabled = True
'      BtnSave.Enabled = False
'      BtnPrint.Enabled = True
'      TxtCode.Enabled = True
'      BtnProduct.Enabled = True
'      TxtCode.SetFocus
'      LblStock.Visible = False
'      LblStockCaption.Visible = False
'      vIsNewRecord = False
'   Case Is = ChangeMode
'      BtnPrint.Enabled = False
'      BtnOpen.Enabled = False
'      BtnDelete.Enabled = False
'      BtnSave.Enabled = True
'   Case Is = SelectionMode
'   End Select
'   Exit Property
'ErrorHandler:
'   Call ShowErrorMessage
'End Property
'
'Private Sub BtnStore_Click()
'   If FunSelectStore(ssButton, False) = True Then
'      If TxtCode.Enabled Then TxtCode.SetFocus
'   Else
'      TxtStoreID.SetFocus
'   End If
'End Sub
'
'Private Sub DtpBillDate_Validate(Cancel As Boolean)
'   TxtBillID.Text = FunGetMaxID()
'End Sub
'
'Private Sub Form_Click()
'   'MsgBox DtpBillDate.DateValue
'   'MsgBox Date + Time
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   On Error GoTo ErrorHandler
'   If Shift = vbCtrlMask Then
'      If ActiveControl.Name = Grid.Name Then KeyCode = 0: Exit Sub
'      Select Case KeyCode
'         Case vbKeyS
'            If BtnSave.Enabled Then BtnSave_Click
'            KeyCode = 0
'         Case vbKeyW
'            If BtnClear.Enabled Then BtnClear_Click
'            KeyCode = 0
'         Case vbKeyQ
'            If BtnClose.Enabled Then BtnClose_Click
'            KeyCode = 0
'         Case vbKeyO
'            If BtnOpen.Enabled Then BtnOpen_Click
'            KeyCode = 0
'         Case vbKeyR
'            If BtnDelete.Enabled Then BtnDelete_Click
'            KeyCode = 0
'         Case vbKeyP
'            If BtnPrint.Enabled Then BtnPrint_Click
'            KeyCode = 0
'      End Select
'   ElseIf KeyCode = vbKeyC And Shift = vbAltMask Then
'      FrmPrint.ParaInChoice = "Credit"
'      FrmPrint.Show vbModal, Me
'   ElseIf KeyCode = vbKeyReturn And Shift = vbShiftMask Then
'      Select Case ActiveControl.Name
'      Case TxtCode.Name
'         If FunSelectProduct(ssValidate, False) = True Then TxtSQty.SetFocus
'      Case TxtSQty.Name
'         If TxtPrice.Visible = False Then TxtSDiscPC.SetFocus Else TxtPrice.SetFocus
'      Case TxtPrice.Name
'         TxtSDiscPC.SetFocus
'      Case TxtSDiscPC.Name
'         TxtDiscPer.SetFocus
'      Case TxtDiscPer.Name
'         TxtSDiscVal.SetFocus
'      End Select
'      KeyCode = 0
'      Shift = 0
'   ElseIf KeyCode = vbKeyReturn Then
'      Select Case ActiveControl.Name
'      Case Grid.Name
'         Grid_DblClick
'      Case TxtCode.Name
'         If FunSelectProduct(ssValidate, False) = True Then GetDataFromTexBoxesToGrid
'      Case TxtSQty.Name, TxtSDiscPC.Name, TxtDiscPer.Name, TxtSDiscVal.Name
'         GetDataFromTexBoxesToGrid
'      Case Else
'         keybd_event 9, 1, 1, 1
'         KeyCode = 0
'      End Select
'   ElseIf KeyCode = vbKeyEscape Then
'      Call SubClearDetailArea: TxtCode.SetFocus
'   ElseIf KeyCode = vbKeyF1 Then
'      Select Case ActiveControl.Name
'         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtCode.Enabled Then TxtCode.SetFocus
'         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtSQty.SetFocus
'      End Select
'   ElseIf ActiveControl.Name = TxtCode.Name Then
'      If KeyCode = vbKeyDown Then
'         Grid.SetFocus
'      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
'         KeyCode = 0
'         TxtBillDisc.SetFocus
'      End If
'   ElseIf ActiveControl.Name = Grid.Name And (Shift = vbCtrlMask + vbShiftMask + vbAltMask) And KeyCode = 46 Then  ' 46 is del button
'      If Trim(Grid.Columns("ProductID").Text <> "") Then Call mniRemoveRow_Click
'   ElseIf ActiveControl.Name = Grid.Name And KeyCode = vbKeyF4 Then
'      If Trim(Grid.Columns("ProductID").Text <> "") Then
'         If MniCostPrice.Visible = True Then
'            Call MniCostPrice_Click
'         End If
'      End If
'   ElseIf KeyCode = vbKeyF5 Then
'      Select Case ActiveControl.Name
'      Case TxtCode.Name, TxtSQty.Name, TxtPrice.Name, TxtSDiscPC.Name, Grid.Name
'         LblCost.Caption = CN.Execute("select dbo.FunPurPrice('" & TxtPID.Text & "')").Fields(0).Value
'         LblCost.Visible = True
'      End Select
'   End If
'   Exit Sub
'ErrorHandler:
'    Call ShowErrorMessage
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'   If KeyAscii = vbKeyReturn Then Exit Sub
'   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
'End Sub
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'   If ActiveControl.Name = Grid.Name And (KeyCode = vbKeyF4 Or KeyCode = vbKeyF5) Then
'      LblCost.Visible = False
'   End If
'End Sub
'
'Private Sub Form_Load()
'   On Error GoTo ErrorHandler
'   ShowPicture Me
'   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
'   SetWindowText Me.hwnd, "Sale Invoice"
'   DtpBillDate.DateValue = Date
'   With CN.Execute("select * from registry")
'      If .RecordCount > 0 Then
'         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
'         FunSelectStore ssValidate, True
'         TxtStoreID.Visible = !StoreVisible
'         BtnStore.Visible = !StoreVisible
'         TxtStoreName.Visible = !StoreVisible
'         LblStoreID.Visible = !StoreVisible
'         LblStoreName.Visible = !StoreVisible
'         MniCostPrice.Visible = !CostVisible
'         If !ChangePrice = True Then
'            If ObjUserSecurity.IsAdministrator = True Then
'               TxtPrice.Enabled = True
'            End If
'         End If
'      End If
'      .Close
'   End With
'   DateFlag = True
'   FormStatus = NewMode
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Function FunGetMaxID() As Long
'   On Error GoTo ErrorHandler
'   If DtpBillDate.IsDateValid = False Then Exit Function
'   FunGetMaxID = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where BillDate = '" & DtpBillDate.DateValue & "'").Fields(0)
'   Exit Function
'ErrorHandler:
'   Call ShowErrorMessage
'End Function
'
'Private Sub SubClearFields()
'   On Error GoTo ErrorHandler
'   Dim ctl As Control
'   For Each ctl In Me.Controls
'      If TypeOf ctl Is SITextBox.txt Then
'         If ctl.Tag = "" Then
'            ctl.Text = ""
'         End If
'      End If
'   Next
'   TxtLastRate.Caption = 0
'   TxtTotalQty.Caption = 0
'   TxtSTotalDiscount.Caption = 0
'   TxtTotalAmount.Caption = 0
'   TxtNetAmount.Caption = 0
'   vTotDisc = 0
'   Grid.CancelUpdate
'   Grid.RemoveAll
'   Grid.AddNew
'   Grid.Columns("ProductID").Text = " "
'   Grid.Update
'   Unload FrmPrint
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   On Error GoTo ErrorHandler
'   If BtnSave.Enabled = True Then
'      If MsgBox("Are you sure to close without save?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then
'         Cancel = 1
'      End If
'   Else
'    'CN.Execute ("exec spcurrentstock")
'    Dim frmObj As Object
'    For Each frmObj In Forms
'        Set frmObj = Nothing
'    Next
'    Set RsBody = Nothing
'    Set RsReport = Nothing
'    Set FrmSaleInvoice = Nothing
'   End If
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
'   On Error GoTo ErrorHandler
'   DispPromptMsg = 0
'   'TxtGrossAmount.Text = Val(TxtGrossAmount.Text) - Grid.Columns("Amount").Value
'   TxtTotalQty.Caption = Val(TxtTotalQty.Caption) - Grid.Columns("Qty").Value
'   vTotDisc = vTotDisc - Grid.Columns("DiscVal").Value
'   TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) - Grid.Columns("TotalAmount").Value
'   SubCalculateFooter
'   FormStatus = ChangeMode
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Sub Grid_DblClick()
'   Call Grid_LostFocus
'End Sub
'
'Private Sub Grid_GotFocus()
'   Flag = True
'   TxtCode.Enabled = False
'   BtnProduct.Enabled = False
'   'TxtCode.BackColor = TxtProductName.BackColor
'   'TxtCode.TabStop = False
'End Sub
'
'Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
'End Sub
'
'Private Sub Grid_LostFocus()
'   Flag = False
'   LblCost.Visible = False
'   If Trim(Grid.Columns("ProductID").Text) = "" Then
'      TxtCode.Text = ""
'      TxtCode.Enabled = True
'      BtnProduct.Enabled = True
'      TxtCode.SetFocus
'   Else
'      TxtCode.Enabled = False
'      BtnProduct.Enabled = False
'      If TxtSQty.Enabled = True And TxtSQty.Visible Then TxtSQty.SetFocus
'      If BtnSave.Enabled = False Then FormStatus = ChangeMode
'   End If
'End Sub
'
'Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If Trim(Grid.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
'   If Button = 2 Then Me.PopupMenu MnuDelete
'End Sub
'
'Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
'   If Flag Then Call GetDataBackFromGridToTexBoxes
'End Sub
'
'Private Sub ImgExit_Click()
'   Unload Me
'End Sub
'
'Private Sub MniCostPrice_Click()
'   On Error GoTo ErrorHandler
'   If Trim(Grid.Columns("Cost").Text) = "" Then Exit Sub
'   LblCost.Caption = Grid.Columns("Cost").Value
'   LblCost.Visible = True
'Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Sub mniRemoveRow_Click()
'   On Error GoTo ErrorHandler
'   If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
'   RsBody.Filter = "Code='" & TxtCode.Text & "'"
'   If RsBody.RecordCount > 0 Then RsBody.Delete
'   Grid.SelBookmarks.RemoveAll
'   Grid.SelBookmarks.Add Grid.Bookmark
'   Grid.DeleteSelected
'   RsBody.Filter = 0
'   Grid.MoveLast
'   GetDataBackFromGridToTexBoxes
'Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Sub GetDataFromTexBoxesToGrid()
'   Dim vrowcounter As Integer
'   If Trim(TxtCode.Text) = "" Then
'      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
'      TxtCode.SetFocus
'      Exit Sub
'   End If
'   If Val(TxtSQty.Text) = 0 Then
'      'MsgBox "Enter Qty.", vbExclamation, "Alert"
'      TxtSQty.SetFocus
'      Exit Sub
'   End If
'On Error GoTo ErrorHandler
'RsBody.Filter = "ProductID='" & TxtPID.Text & "'"
'   If TxtCode.Enabled Then
'      If RsBody.RecordCount = 0 Then
''         If Trim(TxtSQty.Text) > Val(LblStock.Caption) Then
''            MsgBox "Insufficent Stock.", vbExclamation, "Alert"
''            TxtSQty.SetFocus
''            Exit Sub
''         End If
'         RsBody.AddNew
'         Grid.Columns("ProductID").Text = TxtPID.Text
'         Grid.Columns("Code").Text = TxtCode.Text
'         RsBody!ProductID = TxtPID.Text
'         RsBody!Code = TxtCode.Text
'      Else
'         Grid.Redraw = False
'         Grid.MoveFirst
'            For vrowcounter = 1 To Grid.Rows
'               If Grid.Columns("Productid").Text = TxtPID.Text Then
'                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
'                  'SubClearDetailArea
'                  With CN.Execute("select * from registry")
'                     If .RecordCount > 0 Then
'                        If !NegativeSale = False Then
'                           If DtpBillDate.Enabled = True Then
'                              If (Val(LblStock.Caption) - Val(TxtSQty.Text) - Val(Grid.Columns("Qty").Value)) < 0 Then
'                                 MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                                 Grid.MoveLast
'                                 Grid.Redraw = True
'                                 Exit Sub
'                              End If
'                           Else
'                              If (Val(LblStock.Caption) - Val(TxtSQty.Text) - Val(Grid.Columns("Qty").Value) + Val(Grid.Columns("QtyOrigional").Value)) < 0 Then
'                                 MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                                 Grid.MoveLast
'                                 Grid.Redraw = True
'                                 Exit Sub
'                              End If
'                           End If
'                        End If
'                     End If
'                     .Close
'                  End With
'                  TxtSQty.Text = Val(TxtSQty.Text) + Grid.Columns("Qty").Value
'                  TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtSQty.Text) - Val(Grid.Columns("Qty").Text)
'                  vTotDisc = vTotDisc + Val(TxtSDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
'                  'TxtSTotalDiscount.Caption = Val(TxtSTotalDiscount.Caption) + Val(TxtSDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
'                  TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text) - Val(Grid.Columns("TotalAmount").Text)
'                  TxtLastRate.Caption = Val(TxtPrice.Text) - Val(TxtSDiscPC.Text)
'                  TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtSAmount.Text) - Val(Grid.Columns("Amount").Text)
'                  Grid.Columns("ProductName").Text = TxtProductName.Text
'                  Grid.Columns("Qty").Value = Val(TxtSQty.Text)
'                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
'                  Grid.Columns("DiscPC").Value = Val(TxtSDiscPC.Text)
'                  Grid.Columns("DiscPer").Value = Val(TxtDiscPer.Text)
'                  Grid.Columns("DiscVal").Value = Val(TxtSDiscVal.Text)
'                  Grid.Columns("Amount").Value = Val(TxtSAmount.Text)
'                  Grid.Columns("Cost").Value = Val(TxtCost.Text)
'                  Grid.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
'                  Grid.Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
'                  RsBody!Qty = Val(TxtSQty.Text)
'                  RsBody!Price = Val(TxtPrice.Text)
'                  RsBody!DiscPC = Val(TxtSDiscPC.Text)
'                  RsBody!DiscPer = Val(TxtDiscPer.Text)
'                  RsBody!DiscVal = Val(TxtSDiscVal.Text)
'                  RsBody!Cost = Val(TxtCost.Text)
'                  RsBody!IsProduct = Abs(ChkIsProduct.Value)
'                  RsBody!Amount = Val(TxtSAmount.Text)
'                  Grid.MoveLast
'                  Call SubClearDetailArea
'                  TxtCode.SetFocus
'                  Grid.Redraw = True
'                  Exit Sub
'               End If
'               Grid.MoveNext
'            Next vrowcounter
'         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
'         SubClearDetailArea
'         Grid.MoveLast
'         TxtCode.SetFocus
'         Exit Sub
'      End If
'   End If
'   'Grid.Redraw = False
'   With Grid
'      With CN.Execute("select * from registry")
'         If .RecordCount > 0 Then
'            If !NegativeSale = False Then
'               If DtpBillDate.Enabled = True Then
'                  If (Val(LblStock.Caption) - Val(TxtSQty.Text)) < 0 Then
'                     MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                     Grid.Redraw = True
'                     Exit Sub
'                  End If
'               Else
'                  If (Val(LblStock.Caption) - Val(TxtSQty.Text) + Val(Grid.Columns("QtyOrigional").Value)) < 0 Then
'                     MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                     Grid.Redraw = True
'                     Exit Sub
'                  End If
'               End If
'            End If
'         End If
'         .Close
'      End With
'      If TxtCode.Enabled = True Then
'         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtSAmount.Text)
'         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtSQty.Text)
'         'TxtSTotalDiscount.Caption = Val(TxtSTotalDiscount.Caption) + Val(TxtSDiscVal.Text)
'         vTotDisc = vTotDisc + Val(TxtSDiscVal.Text)
'         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text)
'      Else
'         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtSAmount.Text) - Val(Grid.Columns("Amount").Text)
'         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtSQty.Text) - Val(.Columns("Qty").Text)
'         vTotDisc = vTotDisc + Val(TxtSDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
'         'TxtSTotalDiscount.Caption = Val(TxtSTotalDiscount.Caption) + Val(TxtSDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
'         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text) - Val(Grid.Columns("TotalAmount").Text)
'      End If
'      .Columns("ProductName").Text = TxtProductName.Text
'      .Columns("Qty").Value = Val(TxtSQty.Text)
'      .Columns("Price").Value = Val(TxtPrice.Text)
'      .Columns("DiscPC").Value = Val(TxtSDiscPC.Text)
'      .Columns("DiscPer").Value = Val(TxtDiscPer.Text)
'      .Columns("DiscVal").Value = Val(TxtSDiscVal.Text)
'      If Trim(TxtCost.Text) <> "" Then
'         .Columns("Cost").Value = Val(TxtCost.Text)
'      End If
'      .Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
'      .Columns("Amount").Value = Val(TxtSAmount.Text)
'      .Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
'      TxtLastRate.Caption = Val(TxtPrice.Text) - Val(TxtSDiscPC.Text)
'      RsBody!Qty = Val(TxtSQty.Text)
'      RsBody!Price = Val(TxtPrice.Text)
'      RsBody!DiscPC = Val(TxtSDiscPC.Text)
'      RsBody!DiscPer = Val(TxtDiscPer.Text)
'      RsBody!DiscVal = Val(TxtSDiscVal.Text)
'      If Trim(TxtCost.Text) <> "" Then
'         RsBody!Cost = Val(TxtCost.Text)
'      End If
'      If IsNull(RsBody!Cost) Then RsBody!Cost = 0
'      RsBody!IsProduct = Abs(ChkIsProduct.Value)
'      RsBody!Amount = Val(TxtSAmount.Text)
'      .MoveLast
'      If Trim(.Columns("Code").Text) <> "" Then
'         .AllowAddNew = True
'         .AddNew
'         .Columns("Code").Text = " "
'         .AllowAddNew = False
'      End If
'   End With
'   Call SubClearDetailArea
'   TxtCode.SetFocus
'   Grid.Redraw = True
'   Exit Sub
'ErrorHandler:
'   Grid.Redraw = True
'   Call ShowErrorMessage
'End Sub
'
'Private Sub SubClearDetailArea()
'   TxtCode.Enabled = True
'   BtnProduct.Enabled = True
'   TxtCode.Text = ""
'   TxtProductName.Text = ""
'   TxtSQty.Text = ""
'   TxtPrice.Text = ""
'   TxtSDiscPC.Text = ""
'   TxtDiscPer.Text = ""
'   TxtSDiscVal.Text = ""
'   TxtSAmount.Text = ""
'   TxtCost.Text = ""
'   TxtActualAmount.Text = ""
'   ChkIsProduct.Value = 1
'End Sub
'
'Private Sub GetDataBackFromGridToTexBoxes()
'   On Error GoTo ErrorHandler
'   With Grid
'      TxtPID.Text = .Columns("ProductID").Text
'      TxtCode.Text = .Columns("code").Text
'      TxtProductName.Text = .Columns("ProductName").Text
'      TxtSQty.Text = .Columns("Qty").Text
'      TxtPrice.Text = .Columns("Price").Text
'      TxtSDiscPC.Text = .Columns("DiscPC").Value
'      TxtDiscPer.Text = .Columns("DiscPer").Value
'      TxtSDiscVal.Text = .Columns("DiscVal").Value
'      TxtCost.Text = .Columns("Cost").Value
'      TxtSAmount.Text = .Columns("Amount").Text
'      TxtActualAmount.Text = .Columns("TotalAmount").Text
'      ChkIsProduct.Value = Abs(.Columns("IsProduct").Value)
'   End With
'   If Grid.Rows = 1 Then Grid.MoveLast
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Sub GetSale()
'   On Error GoTo ErrorHandler
'   sSql = "select h.*, p.partyname, BankMachineName, StoreName FROM SaleHeader h left outer join parties p on h.customerid=p.partyid left outer join BankMachines b on b.BankMachineid = h.BankMachineid inner join stores s on s.storeid = h.storeid where h.BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
'   With CN.Execute(sSql)
'      If Not .BOF Then
'         TxtStoreID.Text = !StoreID
'         TxtStoreName.Text = !StoreName
'         TxtTotalAmount.Caption = !TotalAmount
'         TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
'         TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
'         FrmPrint.OptBankCard.Value = !BankCard
'         FrmPrint.OptCash.Value = !Cash
'         FrmPrint.OptCredit.Value = !Credit
'         If FrmPrint.OptBankCard.Value = True Then
'            FrmPrint.TxtInvoiceNo.Text = !InvoiceNo
'            FrmPrint.TxtCommision.Text = !Commision
'            FrmPrint.TxtBankMachineID.Text = !BankMachineID
'            FrmPrint.TxtBankMachineName.Text = !BankMachineName
'            FrmPrint.TxtCashReceivedCash.Text = ""
'            FrmPrint.TxtCustomerID.Text = ""
'            FrmPrint.TxtCustomerName.Text = ""
'            FrmPrint.TxtCashCustomer.Text = ""
'            FrmPrint.TxtBankCustomer.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
'         End If
'         If FrmPrint.OptCash.Value = True Then
'            FrmPrint.TxtCommision.Text = ""
'            FrmPrint.TxtInvoiceNo.Text = ""
'            FrmPrint.TxtBankMachineID.Text = ""
'            FrmPrint.TxtBankMachineName.Text = ""
'            FrmPrint.TxtCashReceivedCash.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
'            FrmPrint.TxtCustomerID.Text = ""
'            FrmPrint.TxtCustomerName.Text = ""
'            FrmPrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
'            FrmPrint.TxtBankCustomer.Text = ""
'         End If
'         If FrmPrint.OptCredit.Value = True Then
'            FrmPrint.TxtCommision.Text = ""
'            FrmPrint.TxtInvoiceNo.Text = ""
'            FrmPrint.TxtBankMachineID.Text = ""
'            FrmPrint.TxtBankMachineName.Text = ""
'            FrmPrint.TxtCashReceivedCredit.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
'            FrmPrint.TxtCustomerID.Text = !CustomerID
'            FrmPrint.TxtCustomerName.Text = !PartyName
'            FrmPrint.TxtCashCustomer.Text = ""
'            FrmPrint.TxtBankCustomer.Text = ""
'         End If
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
'
'Private Sub TxtBillDisc_Change()
'   If ActiveControl.Name <> TxtBillDisc.Name Then Exit Sub
'   TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / Val(TxtTotalAmount.Caption), 2)
'   Call SubCalculateFooter
'End Sub
'
'Private Sub TxtBillDiscPer_Change()
'   If ActiveControl.Name <> TxtBillDiscPer.Name Then Exit Sub
'   TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtBillDiscPer.Text) / 100))
'   Call SubCalculateFooter
'End Sub
'
'Private Sub TxtSDiscPC_Change()
'   If ActiveControl.Name <> TxtSDiscPC.Name Then Exit Sub
'   If Val(TxtPrice.Text) = 0 Then Exit Sub
'   TxtDiscPer.Text = Round((Val(TxtSDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
'   Call SubCalculateBody
'End Sub
'
''Private Sub TxtSDiscPC_LostFocus()
''   Select Case Me.ActiveControl.Name
''   Case TxtCode.Name, TxtSQty.Name, TxtSDiscPC.Name
''      Exit Sub
''   End Select
''   Call GetDataFromTexBoxesToGrid
''End Sub
'
'Private Sub TxtDiscPer_Change()
'   If ActiveControl.Name <> TxtDiscPer.Name Then Exit Sub
'   TxtSDiscPC.Text = Round((Val(TxtPrice.Text) * Val(TxtDiscPer.Text) / 100), 2)
'   Call SubCalculateBody
'End Sub
'
'Private Sub TxtCode_Change()
'   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
'   If TxtProductName.Text <> "" Then
'      TxtProductName.Text = ""
'      TxtPrice.Text = ""
'      TxtSDiscPC.Text = ""
'   End If
'End Sub
'
'Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyDown Then Grid.SetFocus
'End Sub
'
''Private Sub TxtCode_LostFocus()
''   If Len(TxtCode.Text) > 7 Then
''      GetDataFromTexBoxesToGrid
''   End If
''End Sub
'
'Private Sub TxtCode_Validate(Cancel As Boolean)
'   On Error GoTo ErrorHandler
'   Dim vTemp As Boolean
'   If Trim(TxtCode.Text) = "" Then Exit Sub
'   vTemp = Not FunSelectProduct(ssValidate, True)
'   If vTemp = True Then
'      vTemp = Not FunSelectProduct(ssValidate, False)
'   End If
'   Cancel = vTemp
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Sub TxtSDiscVal_Change()
'   If TxtSDiscVal.Visible = False Then Exit Sub
'   If ActiveControl.Name <> TxtSDiscVal.Name Then Exit Sub
'   If Val(TxtPrice.Text) = 0 Then Exit Sub
'   If Val(TxtSQty.Text) = 0 Then Exit Sub
'   TxtSDiscPC.Text = Round(Val(TxtSDiscVal.Text) / (TxtSQty.Text), 3)
'   TxtDiscPer.Text = Round((Val(TxtSDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
'   TxtActualAmount.Text = Val(TxtSQty.Text) * Val(TxtPrice.Text)
'   TxtSAmount.Text = Val(TxtActualAmount.Text) - Val(TxtSDiscVal.Text)
'   TxtSTotalDiscount.Caption = vTotDisc
'   SubCalculateFooter
'End Sub
'
'Private Sub TxtPrice_Change()
'   Call SubCalculateBody
'End Sub
'
'Private Sub TxtSQty_Change()
'   Call SubCalculateBody
'End Sub
'
'Private Sub TxtStoreID_Change()
'   If TxtStoreID.Visible = False Then Exit Sub
'   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
'   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
'End Sub
'
'Private Sub TxtStoreID_Validate(Cancel As Boolean)
'   If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
'   On Error GoTo ErrorHandler
'   If TxtStoreName.Text <> "" Then Exit Sub
'   Dim vTemp As Boolean
'   vTemp = Not FunSelectStore(ssValidate, True)
'   If vTemp = True Then
'      vTemp = Not FunSelectStore(ssButton, False)
'   End If
'   Cancel = vTemp
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
