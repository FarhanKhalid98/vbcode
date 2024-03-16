VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmStockTransfer 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmStockTransfer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10755
      TabIndex        =   101
      Top             =   8640
      Width           =   1290
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5055
      MaxLength       =   10
      TabIndex        =   9
      Top             =   3990
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   6450
      MaxLength       =   100
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   3990
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   12045
      TabIndex        =   76
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   9795
      Width           =   2115
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmStockTransfer.frx":0ECA
      Left            =   10890
      List            =   "FrmStockTransfer.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   75
      Tag             =   "1"
      Top             =   10155
      Width           =   3276
   End
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   7116
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4620
      Width           =   2070
   End
   Begin SITextBox.Txt TxtTransferID 
      Height          =   315
      Left            =   2586
      TabIndex        =   0
      Top             =   2460
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
      Left            =   9060
      TabIndex        =   23
      Top             =   9480
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
      MICON           =   "FrmStockTransfer.frx":0ECE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7785
      TabIndex        =   19
      Top             =   9480
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
      MICON           =   "FrmStockTransfer.frx":0EEA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5115
      TabIndex        =   21
      Top             =   9480
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
      MICON           =   "FrmStockTransfer.frx":0F06
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10380
      TabIndex        =   24
      Top             =   9480
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
      MICON           =   "FrmStockTransfer.frx":0F22
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6435
      TabIndex        =   20
      Top             =   9480
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
      MICON           =   "FrmStockTransfer.frx":0F3E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   2520
      TabIndex        =   18
      Top             =   9120
      Width           =   9075
      _ExtentX        =   16007
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
      Left            =   2535
      TabIndex        =   10
      Top             =   4620
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
      Masked          =   1
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   4170
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4620
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
      MICON           =   "FrmStockTransfer.frx":0F5A
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   4530
      TabIndex        =   29
      Top             =   4620
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
      Height          =   3120
      Left            =   2535
      TabIndex        =   30
      Top             =   4935
      Width           =   9015
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   12
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
      stylesets(0).Picture=   "FrmStockTransfer.frx":0F76
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
      Columns.Count   =   12
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
      Columns(5).DataType=   4
      Columns(5).FieldLen=   256
      Columns(6).Width=   1429
      Columns(6).Caption=   "Qt.Loose"
      Columns(6).Name =   "QtyLoose"
      Columns(6).Alignment=   1
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "PackingID"
      Columns(7).Name =   "PackingID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "FromCost"
      Columns(8).Name =   "FromCost"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   4
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "ToCost"
      Columns(9).Name =   "ToCost"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   4
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "Price"
      Columns(10).Name=   "Price"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3201
      Columns(11).Caption=   "Amount"
      Columns(11).Name=   "Amount"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   15901
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpTransferDate 
      Height          =   315
      Left            =   3876
      TabIndex        =   1
      Top             =   2460
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   11016
      TabIndex        =   33
      Top             =   1785
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
   Begin SITextBox.Txt TxtFromStoreID 
      Height          =   315
      Left            =   2601
      TabIndex        =   5
      Top             =   3270
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
   Begin SITextBox.Txt TxtFromStoreName 
      Height          =   315
      Left            =   3636
      TabIndex        =   35
      Top             =   3270
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
   Begin JeweledBut.JeweledButton BtnFromStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   3276
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3270
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
      MICON           =   "FrmStockTransfer.frx":0F92
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtMultiplier 
      Height          =   315
      Left            =   9180
      TabIndex        =   12
      Top             =   4620
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
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
      Left            =   10470
      TabIndex        =   14
      Top             =   4620
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
      Left            =   9720
      TabIndex        =   13
      Top             =   4620
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
   Begin SITextBox.Txt TxtToStoreID 
      Height          =   315
      Left            =   6156
      TabIndex        =   6
      Top             =   3270
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
   Begin SITextBox.Txt TxtToStoreName 
      Height          =   315
      Left            =   7191
      TabIndex        =   43
      Top             =   3285
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
   Begin JeweledBut.JeweledButton BtnToStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6831
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3270
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
      MICON           =   "FrmStockTransfer.frx":0FAE
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtToCost 
      Height          =   315
      Left            =   8280
      TabIndex        =   48
      Top             =   1635
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
   Begin SITextBox.Txt TxtFromCost 
      Height          =   315
      Left            =   9225
      TabIndex        =   50
      Top             =   1635
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3825
      TabIndex        =   22
      Top             =   9450
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
      MICON           =   "FrmStockTransfer.frx":0FCA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPurID 
      Height          =   315
      Left            =   5526
      TabIndex        =   2
      Top             =   2460
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpPurchaseDate 
      Height          =   315
      Left            =   6441
      TabIndex        =   3
      Top             =   2460
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
   Begin JeweledBut.JeweledButton BtnPurchaseAll 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8106
      TabIndex        =   56
      Top             =   2445
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
      MICON           =   "FrmStockTransfer.frx":0FE6
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPurchase 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7746
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2445
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
      MICON           =   "FrmStockTransfer.frx":1002
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnToAdd 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   12645
      TabIndex        =   59
      Top             =   3300
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      TX              =   "To Add"
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
      MICON           =   "FrmStockTransfer.frx":101E
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnFromMinus 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   12645
      TabIndex        =   60
      Top             =   3750
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      TX              =   "From Minus"
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
      MICON           =   "FrmStockTransfer.frx":103A
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnBarCode 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2490
      TabIndex        =   61
      Top             =   9480
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Barcode"
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
      MICON           =   "FrmStockTransfer.frx":1056
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotalQtyLoose 
      Height          =   315
      Left            =   10500
      TabIndex        =   62
      Top             =   8070
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
   Begin SITextBox.Txt TxtTotalQtyPack 
      Height          =   315
      Left            =   9750
      TabIndex        =   63
      Top             =   8070
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6495
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   1635
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
      MICON           =   "FrmStockTransfer.frx":1072
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   6855
      TabIndex        =   65
      Tag             =   "NC"
      Top             =   1635
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
   Begin JeweledBut.JeweledButton BtnProductRange 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   3000
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   4245
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
      MICON           =   "FrmStockTransfer.frx":108E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   6000
      TabIndex        =   70
      Tag             =   "NC"
      Top             =   1635
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
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   2640
      TabIndex        =   71
      Tag             =   "NC"
      Top             =   1635
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
      Left            =   3945
      TabIndex        =   72
      Tag             =   "NC"
      Top             =   1635
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
      Left            =   3585
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   1635
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
      MICON           =   "FrmStockTransfer.frx":10AA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnUpdateStock 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   12495
      TabIndex        =   74
      Top             =   2865
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      TX              =   "Update Stock"
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
      MICON           =   "FrmStockTransfer.frx":10C6
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   9765
      TabIndex        =   7
      Top             =   3285
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
      Left            =   10875
      TabIndex        =   79
      Top             =   3285
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
      Left            =   10515
      TabIndex        =   80
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
      MICON           =   "FrmStockTransfer.frx":10E2
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtPrice 
      Height          =   315
      Left            =   11280
      TabIndex        =   15
      Top             =   4620
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
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
      DecimalPoint    =   3
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   12000
      TabIndex        =   83
      Top             =   4620
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
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
   Begin SITextBox.Txt TxtOtherChargesPer 
      Height          =   315
      Left            =   3600
      TabIndex        =   16
      Top             =   8595
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   6
      IntegralPoint   =   4
   End
   Begin SITextBox.Txt TxtOtherChargesVal 
      Height          =   315
      Left            =   4965
      TabIndex        =   17
      Top             =   8595
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
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
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   2475
      TabIndex        =   88
      Top             =   8595
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
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
   Begin SITextBox.Txt TxtNetAmount 
      Height          =   315
      Left            =   6480
      TabIndex        =   90
      Top             =   8595
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
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
   Begin JeweledBut.JeweledButton BtnSearch 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6075
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   4005
      Visible         =   0   'False
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
      MICON           =   "FrmStockTransfer.frx":10FE
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillNo 
      Height          =   315
      Left            =   4050
      TabIndex        =   8
      Top             =   3990
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtTotalPayable 
      Height          =   315
      Left            =   9405
      TabIndex        =   97
      Top             =   8595
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
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
   Begin SITextBox.Txt TxtPreviousPayable 
      Height          =   315
      Left            =   7875
      TabIndex        =   98
      Top             =   8595
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
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
   Begin VB.Label LblTtlPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payable"
      Height          =   195
      Left            =   9405
      TabIndex        =   100
      Top             =   8370
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Payable"
      Height          =   195
      Left            =   7875
      TabIndex        =   99
      Top             =   8370
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label LblBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      Height          =   195
      Left            =   4050
      TabIndex        =   96
      Top             =   3780
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LblName 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Name"
      Height          =   225
      Left            =   6450
      TabIndex        =   94
      Top             =   3780
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label LblID 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c No."
      Height          =   225
      Left            =   5055
      TabIndex        =   93
      Top             =   3780
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label LblNetAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      Height          =   195
      Left            =   6480
      TabIndex        =   91
      Top             =   8370
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label LblTotalAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   2475
      TabIndex        =   89
      Top             =   8370
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LblOtherChargesPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges %"
      Height          =   195
      Left            =   3600
      TabIndex        =   87
      Top             =   8370
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label LblOtherChargesVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges Val"
      Height          =   195
      Left            =   4965
      TabIndex        =   86
      Top             =   8370
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label LblPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   11340
      TabIndex        =   85
      Top             =   4425
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   12000
      TabIndex        =   84
      Top             =   4425
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   9765
      TabIndex        =   82
      Top             =   3060
      Width           =   525
   End
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   10875
      TabIndex        =   81
      Top             =   3060
      Width           =   780
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Print Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   12045
      TabIndex        =   78
      Top             =   9495
      Width           =   840
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
      Left            =   10215
      TabIndex        =   77
      Top             =   10200
      Width           =   570
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   2640
      TabIndex        =   69
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   3945
      TabIndex        =   68
      Top             =   1380
      Width           =   1350
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   6000
      TabIndex        =   66
      Top             =   1380
      Width           =   585
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchasel Date"
      Height          =   195
      Left            =   6486
      TabIndex        =   58
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase ID"
      Height          =   195
      Left            =   5526
      TabIndex        =   57
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label LblToStock 
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
      Left            =   11556
      TabIndex        =   55
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label LblToCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock To"
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
      Left            =   11556
      TabIndex        =   54
      Top             =   2220
      Width           =   1140
   End
   Begin VB.Label LblFromStock 
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
      Left            =   9441
      TabIndex        =   53
      Top             =   2490
      Width           =   975
   End
   Begin VB.Label LblFromCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock From"
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
      Left            =   9441
      TabIndex        =   52
      Top             =   2175
      Width           =   1470
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "From Cost"
      Height          =   195
      Left            =   9225
      TabIndex        =   51
      Top             =   1380
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "To Cost"
      Height          =   195
      Left            =   8280
      TabIndex        =   49
      Top             =   1380
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Transfer"
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
      TabIndex        =   47
      Top             =   270
      Width           =   2505
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "To Store ID"
      Height          =   195
      Left            =   6156
      TabIndex        =   46
      Top             =   3060
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "To Store Name"
      Height          =   195
      Left            =   7191
      TabIndex        =   45
      Top             =   3060
      Width           =   1080
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (Pack)"
      Height          =   195
      Left            =   9675
      TabIndex        =   42
      Top             =   4425
      Width           =   750
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (Loose)"
      Height          =   195
      Left            =   10470
      TabIndex        =   41
      Top             =   4425
      Width           =   810
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   7110
      TabIndex        =   40
      Top             =   4425
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   9180
      TabIndex        =   39
      Top             =   4425
      Width           =   375
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "From Store Name"
      Height          =   195
      Left            =   3681
      TabIndex        =   38
      Top             =   3075
      Width           =   1230
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "From Store ID"
      Height          =   195
      Left            =   2601
      TabIndex        =   37
      Top             =   3075
      Width           =   975
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   11016
      TabIndex        =   34
      Top             =   1590
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   2535
      TabIndex        =   32
      Top             =   4425
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   4530
      TabIndex        =   31
      Top             =   4425
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
      Left            =   2520
      TabIndex        =   27
      Top             =   8895
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Date"
      Height          =   195
      Left            =   3891
      TabIndex        =   26
      Top             =   2220
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer ID"
      Height          =   195
      Left            =   2601
      TabIndex        =   25
      Top             =   2220
      Width           =   795
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmStockTransfer"
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
Dim Flag As Boolean
Dim ssql As String
Dim i As Integer
Dim vStrSQL As String
Dim vTransferID  As Integer
Dim vTransferDate  As Date
Dim vUnitPrice, vQtyLoose As Double



'----------------------------------

Private Function FunSelectToStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectToStore = False: Exit Function
        TxtToStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Stores where StoreID=" & Val(TxtToStoreID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtToStoreName.Text = !StoreName
          FunSelectToStore = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectToStore = False
          .Close
          TxtToStoreID.Text = ""
          TxtToStoreName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectFromStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectFromStore = False: Exit Function
        TxtFromStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Stores where StoreID=" & Val(TxtFromStoreID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtFromStoreName.Text = !StoreName
          FunSelectFromStore = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectFromStore = False
          .Close
          TxtFromStoreID.Text = ""
          TxtFromStoreName.Text = ""
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
    If Trim(TxtCode.Text) = "" Then Exit Function
'    If Len(TxtCode.Text) <= 5 Then
'      TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
'    End If
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    vStrSQL = " SELECT p.productid, Code, ProductName, PurPrice, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where p.productid = " & TxtCode.Text & " or code='" & TxtCode.Text & "'"
  
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !PurPrice
         If IsNull(!PackingName) Then
            vUnitPrice = !PurPrice
         Else
            TxtMultiplier.Text = !Multiplier
            If !Multiplier <> 0 Then
               vUnitPrice = !PurPrice / !Multiplier
            Else
               vUnitPrice = !PurPrice
            End If
            CmbPackName.Text = !PackingName
         End If
         
         TxtMultiplier.Text = !Multiplier
         TxtQtyLoose.Text = 1
         Call SubCalculateBody
         vStrSQL = "select isnull(dbo.FunStock('" & TxtProductID.Text & "'," & Val(TxtFromStoreID.Text) & ",0,0,0,0,0,0,'" & DtpTransferDate.DateValue + 1 & "',0),0)"
         vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         LblFromStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblFromStock.Caption = LblFromStock.Caption & " " & CmbPackName.Text
         LblFromStock.Caption = LblFromStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblFromStock.Caption = LblFromStock.Caption & " " & "Loose"
         
'         If Not IsNull(!PackingName) Then CmbPackName.Text = !PackingName
'         With CN.Execute("select qtyloose from currentstockStore where productid ='" & TxtProductID.Text & "' and storeid='" & TxtFromStoreID.Text & "'")
'            If .RecordCount > 0 Then
'               LblFromStock.Caption = !QtyLoose
'            Else
'               LblFromStock.Caption = 0
'            End If
'         End With

         vStrSQL = "select isnull(dbo.FunStock('" & TxtProductID.Text & "'," & Val(TxtToStoreID.Text) & ",0,0,0,0,0,0,'" & DtpTransferDate.DateValue + 1 & "',0),0)"
         vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         LblToStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblToStock.Caption = LblToStock.Caption & " " & CmbPackName.Text
         LblToStock.Caption = LblToStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblToStock.Caption = LblToStock.Caption & " " & "Loose"
'         With CN.Execute("select qtyloose from currentstockStore where productid ='" & TxtProductID.Text & "' and storeid='" & TxtToStoreID.Text & "'")
'            If .RecordCount > 0 Then
'               LblToStock.Caption = !QtyLoose
'            Else
'               LblToStock.Caption = 0
'            End If
'         End With

         LblFromStock.Visible = True
         LblFromCaption.Visible = True
         LblToStock.Visible = True
         LblToCaption.Visible = True
'         Char.Speak TxtProductName.Text
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
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         LblFromStock.Visible = False
         LblFromCaption.Visible = False
         LblToStock.Visible = False
         LblToCaption.Visible = False
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnAdd_Click()

End Sub

Private Sub BtnBarCode_Click()
   On Error GoTo ErrorHandler
   If BtnSave.Enabled Then BtnSave_Click
   If vIsNewRecord = False Then
      vTransferID = TxtTransferID.Text
      vTransferDate = DtpTransferDate.DateValue
   End If
      ssql = " Select b.ProductID, ProductName, RetailPrice, isnull(Multiplier,1) * isnull(Qtypack,0)+ QtyLoose as QtyLoose" & vbCrLf _
         + " from StockTransferBody b inner join Products p on b.PRoductID = p.ProductID" & vbCrLf _
         + " where TransferID = " & vTransferID & " and TransferDate = '" & vTransferDate & "'"
'   sSql = "select b.ProductID, Code, ProductName from ProductBarcodes b inner join Products p on p.productid = b.ProductID where len(code) = 11 and code like '110%'"
   
   Dim i As Integer
   With cn.Execute(ssql)
      FrmMultiBarcodes.SubClearFields
      FrmMultiBarcodes.TxtTotQty.Text = "0"
      For i = 1 To .RecordCount
         FrmMultiBarcodes.Grid.Columns("ID").Text = !Productid
         FrmMultiBarcodes.Grid.Columns("Name").Text = !ProductName
         FrmMultiBarcodes.Grid.Columns("Qty").Value = !QtyLoose
         FrmMultiBarcodes.Grid.Update
         FrmMultiBarcodes.Grid.AddNew
         FrmMultiBarcodes.TxtTotQty.Text = Val(FrmMultiBarcodes.TxtTotQty.Text) + !QtyLoose
         .MoveNext
      Next i
   End With
'   FrmMultiBarcodes.Grid.FirstRow = 0
   FrmMultiBarcodes.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   ''''''''''''''''''''''''''''''''''''' User Activities '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   
    ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniStockTransferInvoice", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   cn.BeginTrans
   ''''''''''''''''''''''''''''''''''''' User Activities '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Grid.Redraw = False
   Grid.MoveFirst
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("ProductID").Text) <> "" Then
         cn.Execute "Delete from StockTransferBody where TransferID =" & Val(TxtTransferID.Text) & " and TransferDate='" & DtpTransferDate.DateValue & "' and ProductID='" & Grid.Columns("Productid").Text & "'"
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   cn.Execute "Delete from StockTransferHeader where TransferID = " & Val(TxtTransferID.Text) & " and TransferDate='" & DtpTransferDate.DateValue & "'"
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnFromMinus_Click()
   Call PopulateFromMinusToGrid
End Sub

Private Sub BtnFromStore_Click()
   If FunSelectFromStore(ssButton, False) = True Then
      TxtToStoreID.SetFocus
   Else
      TxtFromStoreID.SetFocus
   End If
End Sub

Private Sub BtnOpen_Click()
   SchStockTransfer.ParaInTransferDate = DtpTransferDate.DateValue
   SchStockTransfer.DtpTransfer.DateValue = DtpTransferDate.DateValue
   SchStockTransfer.Show vbModal
   If SchStockTransfer.ParaOutTransferID <> 0 Then
      TxtTransferID.Text = SchStockTransfer.ParaOutTransferID
      'Dim a
      'a = Split(SchStockTransfer.ParaOutTransferDate, "/")
      DtpTransferDate.DateValue = SchStockTransfer.ParaOutTransferDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetStockTransfer
   End If
End Sub

'Private Sub BtnPrint_Click()
'On Error GoTo ErrorHandler
'   vStrSql = "select u.username, h.TransferID, h.TransferDate, h.TotalAmount as tbill, isnull(h.discount,0) as discount, isnull(h.cashReceived,0) as cashReceived, p.productname, b.qty, b.price-b.discountvalue as price, b.amount" _
'            + " from StockTransferHeader h inner join StockTransferBody b on h.TransferID = b.TransferID and h.TransferDate = b.TransferDate" _
'            + " inner join products p on p.productid = b.productid" _
'            + " inner join users u on u.UserNo = h.UserNo" _
'            + " where h.TransferID= " & Val(TxtTransferID.Text) & " and h.TransferDate='" & DtpTransferDate.DateValue & "' order by SerialNo"
'
'    If RsReport.State = adStateOpen Then RsReport.Close
'    RsReport.Open vStrSql, CN, adOpenStatic, adLockReadOnly
'
'    Set RptReportViewer.Report = New CrpPurchaseInvoice
'    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
'    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
'    'RptReportViewer.Report.PaperSize = crPaperA4
'    'RptReportViewer.Report.PaperSize = crPaperUser
'    'RptReportViewer.Report.SetUserPaperSize 1400, 1200
'    'RptReportViewer.Report.PaperOrientation = crPortrait
'    'RptReportViewer.Show
'    RptReportViewer.Report.PrintOut False
'Exit Sub
'ErrorHandler:
'    Call ShowErrorMessage
'End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   vStrSQL = " Select H.TransferID,  H.TransferDate, PreviousAmount, H.FromStoreID,SF.StoreName FStoreName, H.ToStoreID, ST.StoreName TStoreName, H.EmpID, EmpName," & vbCrLf _
      + " B.PackingID,  B.Code, B.ProductID,  B.QtyPack,  B.QtyLoose ,  B.Multiplier,  " & IIf(ObjRegistry.ShowMultiBranches = True, "Price", "p.RetailPrice") & " RetailPrice  , " & IIf(ObjRegistry.ShowMultiBranches = True, "Amount", "Round(RetailPrice * (IsNull(QtyPack, 0) * IsNull(Multiplier, 0) + IsNull(Qtyloose, 0)), 2)") & " Amount,  P.ProductName, PK.PackingName, OtherChargesVal, OtherChargesPer, TotalAmount +  Isnull(OtherChargesVal,0) NetAmount " & vbCrLf _
      + " From StockTransferHeader H " & vbCrLf _
      + " inner Join StockTransferBody B on H.TransferID = B.TransferID and H.TransferDate = B.TransferDate" & vbCrLf _
      + " inner Join Stores SF on H.FromStoreID = SF. StoreID" & vbCrLf _
      + " inner Join Stores ST on H.ToStoreID = ST. StoreID" & vbCrLf _
      + " Left Outer Join Employees E on H.EmpID = E.EmpID" & vbCrLf _
      + " inner Join Products P on B.ProductID = P.ProductID" & vbCrLf _
      + " Left Outer Join Packings PK on  B. PackingID =  PK.PackingID " & vbCrLf _
      + " where H.TransferID = " & Val(TxtTransferID.Text) & " and H.TransferDate = '" & DtpTransferDate.DateValue & "' order by serialno"
        
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
  
'   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   
'   If InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
'      Set RptReportViewer.Report = New CrpStockTransInv
'      RptReportViewer.Report.PaperSize = crPaperA4
'      RptReportViewer.Report.PaperOrientation = crPortrait
'      RptReportViewer.Report.RightMargin = 0
'   ElseIf ObjRegistry.LaserPrintofSaleInvoice = True Then
'      Set RptReportViewer.Report = New CrpStockTransInv
'      RptReportViewer.Report.PaperSize = crPaperA4
'      RptReportViewer.Report.PaperOrientation = crLandscape
'      RptReportViewer.Report.TopMargin = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))
'      RptReportViewer.Report.LeftMargin = IIf(IsNull(ObjRegistry.X), 0, Val(ObjRegistry.X))
'      RptReportViewer.Report.RightMargin = 225
'   Else
'      Set RptReportViewer.Report = New CrpStockTransInvAurora
'      RptReportViewer.Report.TopMargin = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))
'      RptReportViewer.Report.LeftMargin = IIf(IsNull(ObjRegistry.X), 0, Val(ObjRegistry.X))
'      RptReportViewer.Report.RightMargin = 0
'   End If

   If cmbPrintType.Text = "Half Page" Then
      Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\CrpStockTransInvHalf1.rpt")
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.x
      RptReportViewer.Report.RightMargin = 225
   ElseIf cmbPrintType.Text = "Thermal" Then
      Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\CrpStockTransInvAurora.rpt")
      RptReportViewer.Report.TopMargin = 0
      RptReportViewer.Report.LeftMargin = 0
      RptReportViewer.Report.RightMargin = 0
   Else
      Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\CrpStockTransInv.rpt")
   
   End If
   
   
   RptReportViewer.Report.ReportTitle = "Stock Transfer Invoice"
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   'RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
   'RptReportViewer.Report.PrintOut False
   Dim vPrinter() As String
   vPrinter = Split(CmbPrinters.Text, ",")
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   If ChkIsPreview.Value = 1 Then
      RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False
   End If
'   RptReportViewer.Show vbModal, Me
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

Private Sub BtnProductRange_Click()
On Error GoTo ErrorHandler
   FrmProductRangeGrid.ParaInPartyID = ""
   FrmProductRangeGrid.Show vbModal, Me
   RsTemp.Filter = ""
   If RsTemp.RecordCount > 0 Then
      PopulateTempToGrid
   End If
'   If FrmProductRangeGrid.ParaOutFromID <> "" Then
'   Dim vPID As Long, vCounter As Long
'   vPID = SchProductRange.ParaOutFromID
'   For vCounter = CLng(SchProductRange.ParaOutFromID) To CLng(SchProductRange.ParaOutToID)
'      TxtCode.Text = vPID
'      FunSelectProduct ssValidate, False
'      TxtQtyLoose.Text = SchProductRange.ParaOutQty
'      Call SubCalculateBody
'      GetDataFromTexBoxesToGrid
'      vPID = vPID + 1
'      DoEvents
'   Next vCounter
'   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPurchase_Click()
 If FunSelectPurchase(ssButton, False) = True Then
      If TxtFromStoreID.Enabled Then TxtFromStoreID.SetFocus
   Else
      If TxtPurID.Enabled Then TxtPurID.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniStockTransferInvoice", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where touserno = " & vUser & " and Entrydate <='" & Date & "'"
    With cn.Execute(vStrSQL)
        If .Fields(0).Value >= DtpTransferDate.DateValue Then
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
    ''''''''''''''''''''''''Check Organization '''''''''''''''''''''''''''''''''
  If ObjRegistry.OrganizationMandatory = True And TxtOrganizationID.Text = "" Then
    MsgBox "Please Select Organization", vbInformation, Me.Caption
    If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
    Exit Sub
  End If
  
  ''''''''''''''''''''''''Check AccountNo '''''''''''''''''''''''''''''''''
  If ObjRegistry.ShowMultiBranches = True Then
   If Val(TxtID.Text) = 0 Then
      MsgBox "Please Select AccountNo", vbInformation, Me.Caption
      If TxtID.Visible = True Then TxtID.SetFocus
      Exit Sub
    End If
  End If
'  Header Validation
   If Trim(TxtFromStoreID.Text) = "" Then
      MsgBox "Enter From Store ID.", vbExclamation, Me.Caption
      TxtFromStoreID.SetFocus
      Exit Sub
   End If
   If Trim(TxtToStoreID.Text) = "" Then
      MsgBox "Enter To Store ID.", vbExclamation, Me.Caption
      TxtToStoreID.SetFocus
      Exit Sub
   End If
   If Val(TxtToStoreID.Text) = Val(TxtFromStoreID.Text) Then
      MsgBox "From Store ID not equal To Store ID.", vbExclamation, Me.Caption
      TxtToStoreID.SetFocus
      Exit Sub
   End If
   If DtpTransferDate.Enabled Then
      If cn.Execute("Select * from StockTransferHeader where TransferID = " & Val(TxtTransferID.Text) & " and TransferDate = '" & DtpTransferDate.DateValue & "'").RecordCount > 0 Then
         TxtTransferID.Text = FunGetMaxID
      End If
   End If
  RsBody.Filter = 0
  If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Transfer", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
  End If
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
  If Val(TxtTransferID.Text) = 0 Then
      TxtTransferID.Text = FunGetMaxID
   End If
   cn.BeginTrans
   
   Call UserActivities
   
   ssql = "select * from StockTransferHeader where TransferID=" & Val(TxtTransferID.Text) & " and TransferDate='" & DtpTransferDate.DateValue & "'"
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, cn, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !TransferID = Val(TxtTransferID.Text)
         !TransferDate = DtpTransferDate.DateValue
      End If
      !PurID = IIf(Trim(TxtPurID.Text) = "", Null, TxtPurID.Text)
      !PurchaseDate = IIf(Trim(TxtPurID.Text) = "", Null, DtpPurchaseDate.DateValue)
      !BillNo = IIf(Trim(TxtBillNo.Text) = "", Null, TxtBillNo.Text)
      !FromStoreID = TxtFromStoreID.Text
      !ToStoreID = TxtToStoreID.Text
      !AccountNo = IIf(TxtID.Text = "", Null, TxtID.Text)
      !Description = IIf(TxtDescription.Text = "", Null, TxtDescription.Text)
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !EmpID = IIf(Val(TxtEmployeeID.Text) = 0, Null, TxtEmployeeID.Text)
      !OtherChargesPer = IIf(Val(TxtOtherChargesPer.Text) = 0, Null, Val(TxtOtherChargesPer.Text))
      !OtherChargesVal = IIf(Val(TxtOtherChargesVal.Text) = 0, Null, Val(TxtOtherChargesVal.Text))
      !TotalAmount = IIf(Val(TxtTotalAmount.Text) = 0, Null, Val(TxtTotalAmount.Text))
      !PreviousAmount = IIf(lblPayable.Caption = "Previous Receivable", Val(TxtPreviousPayable.Text), Val(TxtPreviousPayable.Text) * -1)
      !UserNo = vUser
      !IsSync = 0
      .Update
      .Close
   End With
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !TransferID = Val(TxtTransferID.Text)
         !TransferDate = DtpTransferDate.DateValue
         !StoreID = Val(TxtStoreID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from StockTransferBody where TransferID=" & Val(TxtTransferID.Text) & " and TransferDate = '" & DtpTransferDate.DateValue & "'", cn, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.productname, code,b.* from StockTransferBody b join products p on p.productid = b.productid where TransferID=" & Val(TxtTransferID.Text) & " and TransferDate='" & DtpTransferDate.DateValue & "' order by serialno"
      With cn.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = ""
         TxtTotalQtyPack.Text = ""
         TxtTotalQtyLoose.Text = ""
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
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
               Grid.Columns("PackName").Text = cn.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
            End If
            Grid.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            Grid.Columns("QtyLoose").Value = !QtyLoose
            Grid.Columns("Price").Value = !Price
            Grid.Columns("Amount").Value = !Amount
            TxtTotalQtyPack.Text = Val(TxtTotalQtyPack.Text) + Grid.Columns("QtyPack").Value
            TxtTotalQtyLoose.Text = Val(TxtTotalQtyLoose.Text) + Grid.Columns("QtyLoose").Value
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Grid.Columns("Amount").Value
            
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
      LblFromStock.Visible = False
      LblFromCaption.Visible = False
      LblToStock.Visible = False
      LblToCaption.Visible = False
      TxtTransferID.Text = FunGetMaxID()
      DtpTransferDate.Enabled = True
      TxtFromStoreID.Enabled = True
      BtnFromStore.Enabled = True
      TxtToStoreID.Enabled = True
      BtnToStore.Enabled = True
      If DtpTransferDate.Enabled And DtpTransferDate.Visible Then DtpTransferDate.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      DtpTransferDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      TxtFromStoreID.Enabled = False
      BtnFromStore.Enabled = False
      TxtToStoreID.Enabled = False
      BtnToStore.Enabled = False
      LblFromStock.Visible = False
      LblFromCaption.Visible = False
      LblToStock.Visible = False
      LblToCaption.Visible = False
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

Private Sub BtnSearch_Click()
   On Error GoTo ErrorHandler
   If FunSelectAccount(ssButton, True) = True Then
      TxtCode.SetFocus
   Else
      TxtID.SetFocus
   End If
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage

End Sub

Private Sub BtnToAdd_Click()
   Call PopulateToAddToGrid
End Sub

Private Sub BtnToStore_Click()
   If FunSelectToStore(ssButton, False) = True Then
      TxtEmployeeID.SetFocus
   Else
      TxtToStoreID.SetFocus
   End If
End Sub

Private Sub BtnUpdateStock_Click()
 On Error GoTo ErrorHandler
    Me.MousePointer = vbHourglass
    cn.Execute ("ProdUpdatCurrentStockStore")
    Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub CmbPackName_Click()
   If CmbPackName.Text = "" Then
      TxtMultiplier.Enabled = False
      TxtQtyPack.Enabled = False
      TxtMultiplier.Text = ""
      TxtQtyPack.Text = ""
   Else
      TxtMultiplier.Enabled = True
      TxtQtyPack.Enabled = True
      If Trim(TxtCode.Text) <> "" Then
         With cn.Execute("select * from ProductPacking where productid='" & TxtProductID.Text & "' and packingid=" & CmbPackName.ItemData(CmbPackName.ListIndex))
            TxtMultiplier.Text = IIf(.RecordCount = 0, "", !Multiplier)
         .Close
         End With
      End If
   End If
End Sub

Private Sub DtpTransferDate_Change()
   TxtTransferID.Text = FunGetMaxID()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name
         Grid_DblClick
      Case TxtCode.Name
         If FunSelectProduct(ssValidate, False) = True Then keybd_event 9, 1, 1, 1: KeyCode = 0 'If ObjRegistry.AutoEnterBeforeQty = True Then GetDataFromTexBoxesToGrid Else keybd_event 9, 1, 1, 1: KeyCode = 0
      Case TxtQtyLoose.Name, TxtPrice.Name
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
'         Case vbKeyP
'            If BtnPrint.Enabled Then BtnPrint_Click
'            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtPurID.Name: If FunSelectPurchase(ssFunctionKey, False) = True Then TxtFromStoreID.SetFocus Else TxtPurID.SetFocus
         Case TxtFromStoreID.Name: If FunSelectFromStore(ssFunctionKey, False) = True Then TxtToStoreID.SetFocus Else TxtFromStoreID.SetFocus
         Case TxtToStoreID.Name: If FunSelectToStore(ssFunctionKey, False) = True Then TxtCode.SetFocus Else TxtToStoreID.SetFocus
         Case TxtID.Name: If FunSelectAccount(ssFunctionKey, False) = True Then TxtCode.SetFocus Else TxtToStoreID.SetFocus
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, False) = True Then CmbPackName.SetFocus Else TxtCode.SetFocus
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
   SetWindowText Me.hWnd, "Stock Transfer Invoice"
   
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
   
   DtpTransferDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))
   
   If ObjRegistry.ShowMultiBranches = True Then
      LblPrice.Visible = True
      TxtPrice.Visible = True
      LblAmount.Visible = True
      TxtAmount.Visible = True
      Grid.Columns("Price").Visible = True
      Grid.Columns("Amount").Visible = True
      LblTotalAmount.Visible = True
      TxtTotalAmount.Visible = True
      LblOtherChargesPer.Visible = True
      TxtOtherChargesPer.Visible = True
      LblOtherChargesVal.Visible = True
      TxtOtherChargesVal.Visible = True
            
      TxtPreviousPayable.Visible = True
      lblPayable.Visible = True
      TxtTotalPayable.Visible = True
      LblTtlPayable.Visible = True
      
      LblNetAmount.Visible = True
      TxtNetAmount.Visible = True
      LblBillNo.Visible = True
      TxtBillNo.Visible = True
      LblID.Visible = True
      TxtID.Visible = True
      BtnSearch.Visible = True
      LblName.Visible = True
      TxtName.Visible = True
      Grid.Columns("Price").Width = TxtPrice.Width
      Grid.Columns("Amount").Width = TxtAmount.Width
      Grid.Width = Grid.Width + TxtPrice.Width + TxtAmount.Width
   End If
   
   
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
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   With cn.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
      End If
      .Close
   End With
   With cn.Execute("Select * from Packings")
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
   If DtpTransferDate.IsDateValid = False Then Exit Function
   FunGetMaxID = cn.Execute("Select isnull(max(TransferID),0)+1 from StockTransferHeader Where TransferDate = '" & DtpTransferDate.DateValue & "'").Fields(0)
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
   TxtID.Text = ""
   TxtName.Text = ""
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
    Set FrmStockTransfer = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotalQtyPack.Text = Val(TxtTotalQtyPack.Text) - Grid.Columns("QtyPack").Value
   TxtTotalQtyLoose.Text = Val(TxtTotalQtyLoose.Text) - Grid.Columns("QtyLoose").Value
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
   RsBody.Filter = "Code='" & TxtCode.Text & "'"
   If RsBody.RecordCount > 0 Then RsBody.Delete
   cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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
      If TxtQtyPack.Enabled Then TxtQtyPack.SetFocus Else TxtQtyLoose.SetFocus
      Exit Sub
   End If
On Error GoTo ErrorHandler
   RsBody.Filter = "Code='" & TxtCode.Text & "'"
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ProductID").Text = TxtProductID.Text
         Grid.Columns("Code").Text = TxtCode.Text
         Grid.Columns("Price").Value = Val(TxtPrice.Text)
         RsBody!Productid = TxtProductID.Text
         RsBody!Code = TxtCode.Text
         RsBody!Price = Val(TxtPrice.Text)
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("Code").Text = TxtCode.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  TxtQtyLoose.Text = Val(TxtQtyLoose.Text) + Grid.Columns("QtyLoose").Value
                  TxtQtyPack.Text = Val(TxtQtyPack.Text) + Grid.Columns("QtyPack").Value
                  TxtTotalQtyPack.Text = Val(TxtTotalQtyPack.Text) + Val(TxtQtyPack.Text) - Val(Grid.Columns("QtyPack").Text)
                  TxtTotalQtyLoose.Text = Val(TxtTotalQtyLoose.Text) + Val(TxtQtyLoose.Text) - Val(Grid.Columns("QtyLoose").Text)
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
                  Grid.Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
                  Grid.Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, 0, Val(TxtMultiplier.Text))
                  Grid.Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
                  Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  RsBody!PackingID = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), Null)
                  RsBody!Multiplier = Val(TxtMultiplier.Text)
                  RsBody!QtyPack = Val(TxtQtyPack.Text)
                  RsBody!QtyLoose = Val(TxtQtyLoose.Text)
                  RsBody!Price = Val(TxtPrice.Text)
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
   Else
      TxtTotalQtyPack.Text = Val(TxtTotalQtyPack.Text) - Val(Grid.Columns("QtyPack").Text)
      TxtTotalQtyLoose.Text = Val(TxtTotalQtyLoose.Text) - Val(Grid.Columns("QtyLoose").Text)
   End If
   Grid.Redraw = False
   With Grid
      If TxtCode.Enabled = True Then
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text)
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
      End If
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
      .Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, 0, Val(TxtMultiplier.Text))
      .Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
      '.Columns("PackingID").Text = CmbPackName.ItemData(CmbPackName.ListIndex)
      .Columns("Pack").Value = Val(TxtMultiplier.Text)
      .Columns("QtyPack").Text = Val(TxtQtyPack.Text)
      .Columns("QtyLoose").Text = TxtQtyLoose.Text
      .Columns("Price").Value = Val(TxtPrice.Text)
      .Columns("Amount").Value = Val(TxtAmount.Text)
      TxtTotalQtyPack.Text = Val(TxtTotalQtyPack.Text) + Val(TxtQtyPack.Text)
      TxtTotalQtyLoose.Text = Val(TxtTotalQtyLoose.Text) + Val(TxtQtyLoose.Text)
      RsBody!PackingID = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), Null)
      RsBody!Multiplier = Val(TxtMultiplier.Text)
      RsBody!QtyPack = Val(TxtQtyPack.Text)
      RsBody!QtyLoose = Val(TxtQtyLoose.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!Amount = Val(TxtAmount.Text)
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
   CmbPackName.ListIndex = 0
   TxtMultiplier.Text = ""
   TxtQtyPack.Text = ""
   TxtQtyLoose.Text = ""
   TxtPrice.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("Code").Text
      TxtProductName.Text = .Columns("ProductName").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtAmount.Text = .Columns("Amount").Text
      If Trim(.Columns("PackName").Text) = "" Then
         CmbPackName.ListIndex = 0
      Else
         CmbPackName.Text = .Columns("PackName").Text
      End If
      TxtMultiplier.Text = .Columns("Pack").Text
      TxtQtyLoose.Text = .Columns("QtyLoose").Text
      TxtQtyPack.Text = .Columns("QtyPack").Text
      vUnitPrice = Val(.Columns("Price").Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))

   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   
         vStrSQL = "select isnull(dbo.FunStock('" & TxtProductID.Text & "'," & Val(TxtFromStoreID.Text) & ",0,0,0,0,0,0,'" & DtpTransferDate.DateValue + 1 & "',0),0)"
         vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         LblFromStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblFromStock.Caption = LblFromStock.Caption & " " & CmbPackName.Text
         LblFromStock.Caption = LblFromStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblFromStock.Caption = LblFromStock.Caption & " " & "Loose"
         LblFromStock.Visible = True
         LblFromCaption.Visible = True
         
         vStrSQL = "select isnull(dbo.FunStock('" & TxtProductID.Text & "'," & Val(TxtToStoreID.Text) & ",0,0,0,0,0,0,'" & DtpTransferDate.DateValue + 1 & "',0),0)"
         vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         LblToStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblToStock.Caption = LblToStock.Caption & " " & CmbPackName.Text
         LblToStock.Caption = LblToStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblToStock.Caption = LblToStock.Caption & " " & "Loose"
         LblToStock.Visible = True
         LblToCaption.Visible = True
         
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetStockTransfer()
   On Error GoTo ErrorHandler
   ssql = "select h.*,AccountName, EmpName, t.StoreName as ToStoreName, f.StoreName as FromStoreName FROM StockTransferHeader h join stores f on f.storeid = h.fromstoreid join stores t on t.storeid = h.ToStoreid Left Outer Join Employees E on H.EmpID = E.EmpID Left Outer Join ChartOfAccounts C on C.AccountNo = H.AccountNo where h.TransferID=" & Val(TxtTransferID.Text) & " and TransferDate='" & DtpTransferDate.DateValue & "'"
   With cn.Execute(ssql)
      If Not .BOF Then
          TxtPurID.Text = IIf(IsNull(!PurID), "", !PurID)
          DtpPurchaseDate.DateValue = IIf(IsNull(!PurchaseDate), "", !PurchaseDate)
          TxtBillNo.Text = IIf(IsNull(!BillNo), "", !BillNo)
          TxtToStoreID.Text = !ToStoreID
          TxtToStoreName.Text = !ToStoreName
          TxtFromStoreID.Text = !FromStoreID
          TxtFromStoreName.Text = !FromStoreName
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
          TxtEmployeeName.Text = IIf(IsNull(!EmpName), "", !EmpName)
          TxtID.Text = IIf(IsNull(!AccountNo), "", !AccountNo)
          TxtName.Text = IIf(IsNull(!AccountName), "", !AccountName)
          TxtTotalAmount.Text = IIf(IsNull(!TotalAmount), 0, !TotalAmount)
          TxtOtherChargesPer.Text = IIf(IsNull(!OtherChargesPer), 0, !OtherChargesPer)
          TxtOtherChargesVal.Text = IIf(IsNull(!OtherChargesVal), 0, !OtherChargesVal)
          TxtNetAmount.Text = IIf(IsNull(!TotalAmount), 0, !TotalAmount) + IIf(IsNull(!OtherChargesVal), 0, !OtherChargesVal)
          TxtPreviousPayable.Text = IIf(IsNull(!PreviousAmount), "", !PreviousAmount)
          lblPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Previous Receivable", "Previous Payable")
          LblTtlPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Total Receivable", "Total Payable")
          TxtPreviousPayable.Text = Abs(Val(TxtPreviousPayable.Text))
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

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
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

Private Sub TxtFromStoreID_Change()
   If TxtFromStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtFromStoreID.Name Then Exit Sub
   If TxtFromStoreName.Text <> "" Then
      TxtFromStoreName.Text = ""
   End If
End Sub

Private Sub TxtFromStoreID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtFromStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtFromStoreName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectFromStore(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectFromStore(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtMultiplier_Change()
If ActiveControl.Name <> TxtMultiplier.Name Then Exit Sub
    If Val(TxtMultiplier.Text) <> 0 Then
      TxtPrice.Text = Round(vUnitPrice * Val(TxtMultiplier.Text), 3)
   Else
      TxtPrice.Text = Round(vUnitPrice, 3)
   End If
   Call SubCalculateBody
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
      TxtFromStoreID.SetFocus
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


Private Sub TxtOtherChargesPer_Change()
   If TxtOtherChargesPer.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOtherChargesPer.Name Then Exit Sub
   TxtOtherChargesVal.Text = SelfRound((Val(TxtTotalAmount.Text) * Val(TxtOtherChargesPer.Text) / 100))
   Call SubCalculateFooter
End Sub

Private Sub TxtOtherChargesVAl_Change()
   If TxtOtherChargesVal.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOtherChargesVal.Name Then Exit Sub
   TxtOtherChargesPer.Text = Round((Val(TxtOtherChargesVal.Text) * 100) / IIf(Val(TxtTotalAmount.Text) = 0, 1, Val(TxtTotalAmount.Text)), 6)
   Call SubCalculateFooter
End Sub

Private Sub TxtPrice_Change()
If ActiveControl.Name <> TxtPrice.Name Then Exit Sub
If Val(TxtPrice.Text) = 0 Then Exit Sub
   If Val(TxtMultiplier.Text) <> 0 Then
      vUnitPrice = Val(TxtPrice.Text) / Val(TxtMultiplier.Text)
   Else
      vUnitPrice = Val(TxtPrice.Text)
   End If
   Call SubCalculateBody
End Sub

Private Sub TxtQtyLoose_Change()
   If ActiveControl.Name <> TxtQtyLoose.Name Then Exit Sub
   Call SubCalculateBody
End Sub

Private Sub TxtQtyLoose_LostFocus()
'   Select Case ActiveControl.Name
'   Case TxtCode.Name, CmbPackName.Name, TxtQtyPack.Name, TxtMultiplier.Name
'      Exit Sub
'   End Select
'   GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtQtyPack_Change()
If ActiveControl.Name <> TxtQtyPack.Name Then Exit Sub
   Call SubCalculateBody
End Sub

Private Sub TxtToStoreID_Change()
   If TxtToStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtToStoreID.Name Then Exit Sub
   If TxtToStoreName.Text <> "" Then
      TxtToStoreName.Text = ""
   End If
End Sub

Private Sub TxtToStoreID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtToStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtToStoreName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectToStore(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectToStore(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectPurchase(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchPurchase.Show vbModal, Me
        If SchPurchase.ParaOutPurchaseID = "" Then FunSelectPurchase = False: Exit Function
        TxtPurID.Text = SchPurchase.ParaOutPurchaseID
        DtpPurchaseDate.DateValue = SchPurchase.ParaOutPurchaseDate
    End If
    '---------------------------
    vStrSQL = "Select * from PurchaseHeader where PurID=" & Val(TxtPurID.Text) & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          If MsgBox("Do you want to Get Transfer Information from Purchase.", vbQuestion + vbYesNo, "Alert") = vbYes Then
            Call GetPurchase
          End If
          FunSelectPurchase = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          TxtPurID.Text = ""
          DtpPurchaseDate.DateValue = ""
          FunSelectPurchase = False
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub GetPurchase()
   Call PopulatePurchaseDataToGrid
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub PopulatePurchaseDataToGrid()
   ssql = "select h.*,AccountName FROM PurchaseHeader h Left Outer Join ChartOfAccounts C on C.AccountNo = H.VendorID where h.PurID=" & Val(TxtPurID.Text) & " and PurchaseDate='" & DtpTransferDate.DateValue & "'"
   With cn.Execute(ssql)
      If .RecordCount > 0 Then
'         TxtID.Text = !VendorID
         TxtOtherChargesVal.Text = IIf(IsNull(!OtherCharges), 0, !OtherCharges)
         TxtTotalAmount.Text = IIf(IsNull(!TotalAmount), 0, !TotalAmount)
      End If
   End With
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from StockTransferBody where TransferID=" & Val(TxtTransferID.Text) & " and TransferDate = '" & DtpTransferDate.DateValue & "'", cn, adOpenStatic, adLockBatchOptimistic
   ssql = "select p.productname, b.* from PurchaseBody b join products p on p.productid = b.productid where PurID=" & Val(TxtPurID.Text) & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
   With cn.Execute(ssql)
      If .RecordCount > 0 Then
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
               RsBody.AddNew
               RsBody!Code = !Productid
               RsBody!Productid = !Productid
               RsBody!PackingID = !PackingID
               RsBody!Multiplier = !Multiplier
               RsBody!QtyPack = !QtyPack
               RsBody!QtyLoose = !QtyLoose
               RsBody!Price = !Price
               RsBody!Amount = !Amount
               RsBody.Update
               Grid.AddNew
               Grid.Columns("Code").Text = !Productid
               Grid.Columns("ProductID").Text = !Productid
               Grid.Columns("ProductName").Text = !ProductName
               If !PackingID = 0 Or IsNull(!PackingID) Then
                  Grid.Columns("PackingID").Value = ""
               Else
                  Grid.Columns("PackingID").Value = !PackingID
               End If
               If !PackingID = 0 Or IsNull(!PackingID) Then
                  Grid.Columns("PackName").Text = ""
               Else
                  Grid.Columns("PackName").Text = cn.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
               End If
               Grid.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
               Grid.Columns("QtyPack").Value = IIf(IsNull(RsBody!QtyPack), "", RsBody!QtyPack)
               Grid.Columns("QtyLoose").Value = RsBody!QtyLoose
               Grid.Columns("Price").Value = !Price
               Grid.Columns("Amount").Value = !Amount
               TxtTotalQtyPack.Text = Val(TxtTotalQtyPack.Text) + Grid.Columns("QtyPack").Value
               TxtTotalQtyLoose.Text = Val(TxtTotalQtyLoose.Text) + Grid.Columns("QtyLoose").Value
            .MoveNext
         Wend
         .Close
         Grid.AddNew
         Grid.Columns("Code").Text = " "
         Grid.AllowAddNew = False
         Grid.Redraw = True
      End If
   End With
End Sub

Private Sub UserActivities()
     If vIsNewRecord = False Then
     
    With cn.Execute("Select  * from StockTransferHeader where TransferID =" & TxtTransferID.Text & " And TransferDate = '" & DtpTransferDate.DateValue & "'")
        If TxtFromStoreID.Text <> !FromStoreID Then
            cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtToStoreID.Text <> !ToStoreID Then
            cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.Rows - 1
        With cn.Execute("Select * from StockTransferBody Where TransferID = " & TxtTransferID.Text & " and TransferDate ='" & DtpTransferDate.DateValue & "' and Productid ='" & Grid.Columns("Productid").Text & "'")
        
             If .EOF = True Then
                cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("QtyPack").Text <> !QtyPack Or Grid.Columns("QtyLoose").Text <> !QtyLoose Then
                   cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Updated ProdcutID-" & Grid.Columns("Code").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
    
   Else
    cn.Execute ("Insert Into UserActivities values ('Stock Transfer'" & "," & TxtTransferID.Text & ",'" & DtpTransferDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub PopulateToAddToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from StockTransferBody where TransferID=" & Val(TxtTransferID.Text) & " and TransferDate = '" & DtpTransferDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
   'sSql = "select p.productname, b.* from PurchaseBody b join products p on p.productid = b.productid where PurID=" & Val(TxtPurID.Text) & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
    ssql = "Select c.*, ProductName" & vbCrLf _
      + " From CurrentStockStore c inner join Products p on p.ProductID = c.ProductID" & vbCrLf _
      + " where c.StoreID = " & TxtToStoreID.Text & " and QtyLoose < 0 " & vbCrLf _

   With cn.Execute(ssql)
      If .RecordCount > 0 Then
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            RsBody.AddNew
            RsBody!Code = !Productid
            RsBody!Productid = !Productid
            RsBody!PackingID = Null
            RsBody!Multiplier = Null
            RsBody!QtyPack = Null
            RsBody!QtyLoose = Abs(!QtyLoose)
            RsBody.Update
            Grid.AddNew
            Grid.Columns("Code").Text = !Productid
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("PackingID").Value = ""
            Grid.Columns("PackName").Text = ""
            Grid.Columns("Pack").Value = ""
            Grid.Columns("QtyPack").Value = ""
            Grid.Columns("QtyLoose").Value = Abs(!QtyLoose)
            .MoveNext
         Wend
         .Close
         Grid.AddNew
         Grid.Columns("Code").Text = " "
         Grid.AllowAddNew = False
         Grid.Redraw = True
      End If
   End With
End Sub

Private Sub PopulateFromMinusToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from StockTransferBody where TransferID=" & Val(TxtTransferID.Text) & " and TransferDate = '" & DtpTransferDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
   'sSql = "select p.productname, b.* from PurchaseBody b join products p on p.productid = b.productid where PurID=" & Val(TxtPurID.Text) & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
    ssql = "Select c.*, ProductName" & vbCrLf _
      + " From CurrentStockStore c inner join Products p on p.ProductID = c.ProductID" & vbCrLf _
      + " where c.StoreID = " & TxtFromStoreID.Text & " and QtyLoose > 0 " & vbCrLf _

   With cn.Execute(ssql)
      If .RecordCount > 0 Then
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            RsBody.AddNew
            RsBody!Code = !Productid
            RsBody!Productid = !Productid
            RsBody!PackingID = Null
            RsBody!Multiplier = Null
            RsBody!QtyPack = Null
            RsBody!QtyLoose = !QtyLoose
            RsBody.Update
            Grid.AddNew
            Grid.Columns("Code").Text = !Productid
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("PackingID").Value = ""
            Grid.Columns("PackName").Text = ""
            Grid.Columns("Pack").Value = ""
            Grid.Columns("QtyPack").Value = ""
            Grid.Columns("QtyLoose").Value = !QtyLoose
            .MoveNext
         Wend
         .Close
         Grid.AddNew
         Grid.Columns("Code").Text = " "
         Grid.AllowAddNew = False
         Grid.Redraw = True
      End If
   End With
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
    With cn.Execute(vStrSQL)
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
'   TxtVoucherNo.Text = FunGetMaxID
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub PopulateTempToGrid()
   On Error GoTo ErrorHandler
   With RsTemp
      RsTemp.MoveFirst
      Grid.Redraw = False
      'Grid.MoveLast
      'Grid.RemoveAll
      Grid.AllowAddNew = True
      'TxtTotalAmount.Text = 0
'      TxtVenderID.Text = !PartyID
''      Call FunSelectVender(1, True)
      While Not .EOF
         Grid.Columns("ProductID").Text = !Productid
         Grid.Columns("Code").Text = !Productid
         Grid.Columns("ProductName").Text = !ProductName
         Grid.Columns("PackName").Text = !PackingName
         Grid.Columns("QtyPack").Text = IIf(Val(!QtyPack) = 0, "", Val(!QtyPack))
         Grid.Columns("Pack").Text = !Multiplier
'         Grid.Columns("DiscPer").Value = !RetailPer
         
         
         
         
         RsBody.AddNew
         RsBody!Productid = !Productid
         RsBody!Code = !Productid
         
         Grid.Columns("QtyLoose").Value = !QtyLoose
'         Grid.Columns("Price").Value = !Price
         
''         vUnitPrice = 0
         If Val(!Multiplier) <> 0 Then
'            vUnitPrice = Val(!RetailPrice) / Val(!Multiplier)
         Else
'            vUnitPrice = Val(!RetailPrice)
         End If
'         If Val(!RetailPer) <> 0 Then
'            Grid.Columns("DiscPC").Value = Round((vUnitPrice * !RetailPer / 100), 4)
'            Grid.Columns("DiscVal").Value = Round((Val(vUnitPrice) * (Val(!QtyPack) * Val(!Multiplier) + Val(!QtyLoose))) * Val(!RetailPer) / 100, 2)
'         End If
         
'         Grid.Columns("RetailPrice").Value = 0
'         Grid.Columns("IsWSDiscb4ST").Value = 0
'         Grid.Columns("IsWSSaleTax").Value = 0
'         Grid.Columns("IsRetailSaleTax").Value = 0
'
'         Grid.Columns("Amount").Value = (!Price * !QtyLoose)

         '''''
'          RsBody!PackingID = !PackingID
          RsBody!Multiplier = !Multiplier
          RsBody!QtyPack = !QtyPack
          RsBody!QtyLoose = !QtyLoose
'         RsBody!Bonus = Null
'         RsBody!Price = !Price
         
'         RsBody!RetailPrice = 0
'         RsBody!IsWSDiscb4ST = 0
'         RsBody!IsWSSaleTax = 0
'         RsBody!IsRetailSaleTax = 0
'
'         RsBody!DiscPC = 0
'         RsBody!Offer = Null
'         RsBody!SaleTaxPer = Null
'         RsBody!SaleTaxval = Null
'         RsBody!DiscPer = 0
'         RsBody!DiscVal = 0
'         RsBody!Amount = (!Price * !QtyLoose)
         RsBody.Update
         ''''
         
'         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + (!Price * !QtyLoose)
'         TxtTotalItems.Text = Val(TxtTotalItems.Text) + !QtyLoose
         .MoveNext
         Grid.AddNew
      Wend
      .Close
   End With
'   Grid.AddNew
   Grid.Columns("Code").Text = " "
   Grid.AllowAddNew = False
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnEmployee_Click()
   On Error GoTo ErrorHandler
   If FunSelectEmployee(ssButton, False) = True Then
      If TxtBillNo.Visible = True Then TxtBillNo.SetFocus Else TxtCode.SetFocus
   Else
      TxtEmployeeID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub TxtEmployeeID_Change()
   If TxtEmployeeID.Visible = False Then Exit Sub
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
    ssql = "Select *" & vbCrLf _
            + " from Employees" & vbCrLf _
            + " where isLockEmployee = 0 and EmpID = " & Val(TxtEmployeeID.Text)
    With cn.Execute(ssql)
      If .RecordCount > 0 Then
        TxtEmployeeName.Text = !EmpName
        FunSelectEmployee = True
        .Close
        Exit Function
      Else
        FunSelectEmployee = False
        .Close
        TxtEmployeeID.Text = ""
        TxtEmployeeName.Text = ""
        Exit Function
      End If
    End With
Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function


Private Sub SubCalculateBody()
   On Error GoTo ErrorHandler
    
   TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))), IIf(ObjRegistry.IsRoundFigure, 0, 2))
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub SubCalculateFooter()
   On Error GoTo ErrorHandler
   
   TxtNetAmount.Text = SelfRound(Val(TxtTotalAmount.Text) + Val(TxtOtherChargesVal.Text))
   TxtTotalPayable.Text = Abs(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousPayable.Text, Val(TxtPreviousPayable.Text) * -1)))
   LblTtlPayable.Caption = IIf(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousPayable.Text, Val(TxtPreviousPayable.Text) * -1)) < 0, "Total Receivable", "Total Payable")
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub TxtTotalAmount_Change()
   TxtOtherChargesVal.Text = SelfRound((Val(TxtTotalAmount.Text) * Val(TxtOtherChargesPer.Text) / 100))
   Call SubCalculateFooter
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

    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtName.Text = !AccountName
          TxtPreviousPayable.Text = cn.Execute("SELECT isnull(dbo.FunCurrentDebit(" & Val(TxtID.Text) & ",'" & DtpTransferDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
          vStrSQL = " Select isnull(Sum(TotalAmount + isnull(OtherChargesval,0)),0) as Amount  " & vbCrLf _
                  + " FROM StockTransferHeader h INNER JOIN (Select TransferId, TransferDate, Sum(amount) TTLValue FROM StockTransferBody Group By TransferId, TransferDate)B " & vbCrLf _
                  + " ON h.Transferid = B.Transferid and h.TransferDate = b.TransferDate " & vbCrLf _
                  + " where Accountno = " & Val(TxtID.Text) & " and h.TransferDate = '" & DtpTransferDate.DateValue & "' and h.TransferID >= " & Val(TxtTransferID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
          TxtPreviousPayable.Text = TxtPreviousPayable.Text - cn.Execute(vStrSQL).Fields(0).Value
          lblPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Previous Receivable", "Previous Payable")
          TxtPreviousPayable.Text = Abs(TxtPreviousPayable.Text)
          FunSelectAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectAccount = False
          .Close
'          LblBalance.Visible = False
'          LblBalanceCaption.Visible = False
          TxtPreviousPayable.Text = ""
          lblPayable.Caption = "Previous Payable"
          LblTtlPayable.Caption = "Total Payable"
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


