VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmProductionOutOld 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmProductionOutOld.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742.268
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   4170
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3075
      Width           =   1845
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7365
      TabIndex        =   19
      Top             =   7388
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
      MICON           =   "FrmProductionOutOld.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6045
      TabIndex        =   15
      Top             =   7388
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
      MICON           =   "FrmProductionOutOld.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3405
      TabIndex        =   17
      Top             =   7388
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
      MICON           =   "FrmProductionOutOld.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8685
      TabIndex        =   20
      Top             =   7388
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
      MICON           =   "FrmProductionOutOld.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4725
      TabIndex        =   16
      Top             =   7388
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
      MICON           =   "FrmProductionOutOld.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtMultiplier1 
      Height          =   315
      Left            =   11355
      TabIndex        =   21
      Top             =   7680
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2070
      TabIndex        =   18
      Top             =   7388
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
      MICON           =   "FrmProductionOutOld.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCost 
      Height          =   315
      Left            =   10335
      TabIndex        =   23
      Top             =   7680
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
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
   Begin SITextBox.Txt TxtPOUTID 
      Height          =   315
      Left            =   803
      TabIndex        =   0
      Top             =   1320
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
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   7695
      TabIndex        =   5
      Top             =   2220
      Width           =   4230
      _ExtentX        =   7461
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
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpPOUTDate 
      Height          =   315
      Left            =   1853
      TabIndex        =   1
      Top             =   1320
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
      AllowEdit       =   0   'False
   End
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   8565
      TabIndex        =   30
      Top             =   6390
      Width           =   1530
      _ExtentX        =   2699
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
      DecimalPoint    =   3
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   8115
      TabIndex        =   13
      Top             =   3075
      Width           =   780
      _ExtentX        =   1376
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
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   8895
      TabIndex        =   14
      Top             =   3075
      Width           =   1215
      _ExtentX        =   2143
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   720
      TabIndex        =   6
      Top             =   3075
      Width           =   960
      _ExtentX        =   1693
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
      Left            =   1680
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3075
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
      MICON           =   "FrmProductionOutOld.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Top             =   3075
      Width           =   2130
      _ExtentX        =   3757
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
      Height          =   2865
      Left            =   720
      TabIndex        =   32
      Top             =   3390
      Width           =   9375
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   10
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
      stylesets(0).Picture=   "FrmProductionOutOld.frx":0F8E
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
      Columns.Count   =   10
      Columns(0).Width=   2328
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3757
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "ProductName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3254
      Columns(2).Caption=   "Pack Name"
      Columns(2).Name =   "PackName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   953
      Columns(3).Caption=   "Pack"
      Columns(3).Name =   "Pack"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   4
      Columns(3).FieldLen=   256
      Columns(4).Width=   1323
      Columns(4).Caption=   "Qt.Pack"
      Columns(4).Name =   "QtyPack"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   4
      Columns(4).FieldLen=   256
      Columns(5).Width=   1429
      Columns(5).Caption=   "Qt.Loose"
      Columns(5).Name =   "QtyLoose"
      Columns(5).Alignment=   1
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   4
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "PackingID"
      Columns(6).Name =   "PackingID"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1376
      Columns(7).Caption=   "Cost"
      Columns(7).Name =   "PurPrice"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   1508
      Columns(8).Caption=   "Amount"
      Columns(8).Name =   "Amount"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   5
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "ProductID"
      Columns(9).Name =   "ProductID"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   16536
      _ExtentY        =   5054
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
   Begin SITextBox.Txt TxtMultiplier 
      Height          =   315
      Left            =   6015
      TabIndex        =   10
      Top             =   3075
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
      Left            =   7305
      TabIndex        =   12
      Top             =   3075
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
      Left            =   6555
      TabIndex        =   11
      Top             =   3075
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   10680
      TabIndex        =   33
      Top             =   8250
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   1455
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
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   3443
      TabIndex        =   43
      Top             =   1320
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
      Left            =   4088
      TabIndex        =   44
      Top             =   1320
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
   Begin JeweledBut.JeweledButton BtnReturnAll 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5753
      TabIndex        =   45
      Top             =   1305
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      TX              =   "All Sale"
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
      MICON           =   "FrmProductionOutOld.frx":0FAA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSale 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5393
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1305
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
      MICON           =   "FrmProductionOutOld.frx":0FC6
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtFromStoreID 
      Height          =   315
      Left            =   855
      TabIndex        =   3
      Top             =   2220
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
      Left            =   1890
      TabIndex        =   49
      Top             =   2220
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
      Left            =   1530
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2220
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
      MICON           =   "FrmProductionOutOld.frx":0FE2
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtToStoreID 
      Height          =   315
      Left            =   4275
      TabIndex        =   4
      Top             =   2220
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
      Left            =   5310
      TabIndex        =   51
      Top             =   2220
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
      Left            =   4950
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2220
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
      MICON           =   "FrmProductionOutOld.frx":0FFE
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   7110
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   1305
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
      Left            =   8415
      TabIndex        =   59
      Tag             =   "NC"
      Top             =   1305
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
      Left            =   8055
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   1305
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
      MICON           =   "FrmProductionOutOld.frx":101A
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
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
      Left            =   7110
      TabIndex        =   62
      Top             =   1125
      Width           =   1335
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
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
      Left            =   8550
      TabIndex        =   61
      Top             =   1125
      Width           =   1620
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
      TabIndex        =   58
      Top             =   3060
      Width           =   720
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
      TabIndex        =   57
      Top             =   3375
      Width           =   975
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "From Store ID"
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
      Left            =   855
      TabIndex        =   56
      Top             =   2025
      Width           =   1185
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "From Store Name"
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
      Left            =   2250
      TabIndex        =   55
      Top             =   2025
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "To Store Name"
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
      Index           =   1
      Left            =   5310
      TabIndex        =   54
      Top             =   2025
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "To Store ID"
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
      Left            =   4275
      TabIndex        =   53
      Top             =   2025
      Width           =   1005
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      Left            =   4088
      TabIndex        =   48
      Top             =   1125
      Width           =   735
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill ID"
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
      Left            =   3443
      TabIndex        =   47
      Top             =   1125
      Width           =   525
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   8970
      TabIndex        =   42
      Top             =   2880
      Width           =   645
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
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
      Left            =   8100
      TabIndex        =   41
      Top             =   2880
      Width           =   390
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
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
      Left            =   6570
      TabIndex        =   40
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
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
      Left            =   7305
      TabIndex        =   39
      Top             =   2880
      Width           =   585
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
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
      Left            =   4185
      TabIndex        =   38
      Top             =   2880
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
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
      Index           =   0
      Left            =   6015
      TabIndex        =   37
      Top             =   2880
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   720
      TabIndex        =   36
      Top             =   2880
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
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
      Left            =   2040
      TabIndex        =   35
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   10680
      TabIndex        =   34
      Top             =   8055
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Left            =   7335
      TabIndex        =   31
      Top             =   6480
      Width           =   1140
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Production Out"
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
      Left            =   1890
      TabIndex        =   29
      Top             =   120
      Width           =   2625
   End
   Begin VB.Label LblUnit 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   9840
      TabIndex        =   28
      Top             =   6450
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "POut ID"
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
      Left            =   803
      TabIndex        =   27
      Top             =   1125
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POut Date"
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
      Left            =   1853
      TabIndex        =   26
      Top             =   1125
      Width           =   900
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   7695
      TabIndex        =   25
      Top             =   2025
      Width           =   975
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   10380
      TabIndex        =   24
      Top             =   7455
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   11355
      TabIndex        =   22
      Top             =   7515
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11520
      Top             =   0
      Width           =   330
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmProductionOutOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vIsNewRow As Boolean
Dim vCounter As Integer
Dim vUnitPrice As Double
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim sSql As String
Dim vStrSQL As String
Dim vQtyLoose As Double
Dim vOrder As Double
'----------------------------------

Private Sub SubCalculateBody()
   TxtAmount.Text = Val(TxtPurPrice.Text) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))
End Sub

Private Function FunSelectToStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectToStore = False: Exit Function
        TxtToStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Stores where islock = 0 and StoreID=" & Val(TxtToStoreID.Text)
    With CN.Execute(vStrSQL)
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
    vStrSQL = " Select * FROM Stores where islock = 0 and StoreID=" & Val(TxtFromStoreID.Text)
    With CN.Execute(vStrSQL)
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

Private Sub BtnFromStore_Click()
   If FunSelectFromStore(ssButton, False) = True Then
      TxtToStoreID.SetFocus
   Else
      TxtFromStoreID.SetFocus
   End If
End Sub

Private Sub BtnToStore_Click()
   If FunSelectToStore(ssButton, False) = True Then
      TxtCode.SetFocus
   Else
      TxtToStoreID.SetFocus
   End If
End Sub

Private Sub LblFromCaption_Click()

End Sub

Private Sub TxtFromStoreID_Change()
   If TxtFromStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtFromStoreID.Name Then Exit Sub
   If TxtFromStoreName.Text <> "" Then TxtFromStoreName.Text = ""
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

Private Sub TxtToStoreID_Change()
   If TxtToStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtToStoreID.Name Then Exit Sub
   If TxtToStoreName.Text <> "" Then TxtToStoreName.Text = ""
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

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.ParaInWhere = " and isLocked = 0 and isNoCostProduct = 0 and isRawProduct = 0 and (StoreID is Null or StoreID = " & TxtToStoreID.Text & ")"
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCode.Text) = "" Then Exit Function
    If Len(TxtCode.Text) <= 5 Then
      TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
    End If
    'If Len(TxtCode.Text) < 7 Then
    '  TxtCode.Text = "045" + Right("0000" + CStr(Val(TxtCode.Text)), 4)
    'End If
    
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    vStrSQL = " SELECT p.productid, Code, ProductName, PurPrice, isnull(Cost,0) as Cost, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join CurrentStock CS on CS.ProductID = P.ProductID " & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where (p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "')" & " and isLocked = 0 and isNoCostProduct = 0 and isRawProduct = 0  and (StoreID is Null or StoreID = " & TxtToStoreID.Text & ")"
  
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !ProductID
         TxtProductName.Text = !ProductName
         TxtPurPrice.Text = !Cost
         TxtMultiplier.Text = !Multiplier
         If Not IsNull(!PackingName) Then CmbPackName.Text = !PackingName
         With CN.Execute("select qtyloose from currentstockStore where productid ='" & TxtProductID.Text & "' and StoreID = " & TxtToStoreID.Text)
            If .RecordCount > 0 Then
               vQtyLoose = !QtyLoose
               LblStock.Caption = !QtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtProductID.Text & "')").Fields(0).Value
            Else
               vQtyLoose = 0
               LblStock.Caption = 0
            End If
         End With
         LblStock.Visible = True
         LblStockCaption.Visible = True
         With CN.Execute("select * from registry")
            If .RecordCount > 0 Then
               If !NegativeSale = False Then
                  If vQtyLoose <= 0 Then
                     MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                     FunSelectProduct = False
                     Exit Function
                  End If
               End If
            End If
            .Close
         End With
         SubCalculateBody
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         TxtCode.Text = ""
         TxtProductID.Text = ""
         If CmbPackName.ListCount > 0 Then CmbPackName.ListIndex = 0
         TxtProductName.Text = ""
         TxtMultiplier.Text = ""
         TxtPurPrice.Text = ""
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
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from ProductionOutBody where POutID = " & Val(TxtPOUTID.Text) & " and productid ='" & Grid.Columns("Productid").Text & "' " & IIf(Val(Grid.Columns("PackingID").Text) = 0, "And PackingID is Null", "And PackingID = " & Val(Grid.Columns("PackingID").Text))
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   CN.Execute "Delete from ProductionOutHeader where POutID = " & Val(TxtPOUTID.Text)
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchPOUT.Show vbModal
   If SchPOUT.ParaOutPOUTID <> 0 Then
      TxtPOUTID.Text = SchPOUT.ParaOutPOUTID
      GetProductionOut
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   BtnPrint.Enabled = False
   vStrSQL = " Select H.POutID, H.POutDate, H.FromStoreID, H.ToStoreID, H.Description," & vbCrLf _
            + " B.ProductID, B.PackingID, B.Multiplier, isnull(B.QtyPack,0) as QtyPack, B.QtyLoose, UnitName," & vbCrLf _
            + " Isnull(B.Price,0) Price, Isnull(B.Amount,0) Amount, f.StoreName as FromStoreName, t.StoreName as ToStoreName, ProductName," & vbCrLf _
            + " PackingName, UserName from ProductionOutHeader  H" & vbCrLf _
            + " join ProductionOutBody B on H.POutID = B.POutID" & vbCrLf _
            + " Join Stores F on H.FromStoreID = F.StoreID" & vbCrLf _
            + " Join Stores T on H.ToStoreID = T.StoreID" & vbCrLf _
            + " Join Products p on P.ProductID = B.ProductID" & vbCrLf _
            + " left outer join Packings PK on Pk.PackingID = B.PackingID " & vbCrLf _
            + " left outer join users u on u.userno = h.userno " & vbCrLf _
            + " Left Outer Join Units ut on ut.UnitId = p.unitid where h.POutID = " & Val(TxtPOUTID.Text)

      
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenDynamic, adLockReadOnly
  
   Set RptReportViewer.Report = New CrpProductionOUTInvoice
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    
   With CN.Execute("Select CompanyName,Address,City,PhoneNo,email from Company")
      If .RecordCount > 0 Then
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(IsNull(!CompanyName), "", CStr(!CompanyName))
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(IsNull(!Address), "", !Address) & IIf(IsNull(!City), "", ", " & !City & ".")
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(IsNull(!PhoneNo), "", CStr(!PhoneNo))
      End If
      .Close
   End With
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
    
'   RptReportViewer.Report.ParameterFields(1).AddCurrentValue objRegistry.CompanyName
'   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(objRegistry.CompanyAddress = "", "", objRegistry.CompanyAddress) & IIf(objRegistry.CompanyCity = "", "", ", " & objRegistry.CompanyCity)
'   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(objRegistry.CompanyPhoneNo = "", ".", " Phone # " & objRegistry.CompanyPhoneNo)
'   RptReportViewer.Report.ParameterFields(4).AddCurrentValue objRegistry.DevelopedBy
'   RptReportViewer.Report.SelectPrinter objRegistry.DriverName, objRegistry.DeviceName, objRegistry.Port
'   RptReportViewer.Report.PrintOut False
   RptReportViewer.Show vbModal, Me
   BtnPrint.Enabled = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
   BtnPrint.Enabled = True
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      If TxtQtyLoose.Enabled Then TxtQtyLoose.SetFocus
      'CmbPackName.SetFocus
   Else
      If TxtCode.Enabled Then TxtCode.SetFocus
   End If
End Sub

Private Sub BtnReturnAll_Click()
   PopulateSaleDataToGrid
End Sub

Private Sub BtnSale_Click()
   If FunSelectSale(ssButton, False) = True Then
      TxtDescription.SetFocus
   Else
      TxtBillID.SetFocus
   End If
End Sub

Private Function FunSelectSale(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSale.ParaInBillDate = DtpBillDate.DateValue
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
          If MsgBox("Do you want to Get Information from Sale.", vbQuestion + vbYesNo, "Alert") = vbYes Then
            Call GetSale
          End If
          FunSelectSale = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSale = False
          TxtBillID.Text = ""
          DtpBillDate.DateValue = Date
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub GetSale()
   Call PopulateSaleDataToGrid
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub PopulateSaleDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from ProductionOutBody where POUTID=" & Val(TxtPOUTID.Text), CN, adOpenDynamic, adLockBatchOptimistic
   sSql = "select p.productname, b.code, b.* from SaleBody b join products p on p.productid = b.productid where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   With CN.Execute(sSql)
      If .RecordCount > 0 Then
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            RsBody.AddNew
            RsBody!ProductID = !ProductID
            RsBody!PackingID = Null
            RsBody!Multiplier = Null
            RsBody!QtyPack = 0
            RsBody!QtyLoose = !Qty
            RsBody!Price = 0
            RsBody!Amount = 0
            RsBody.Update
            Grid.AddNew
            Grid.Columns("ProductID").Text = !ProductID
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("QtyLoose").Value = !Qty
            Grid.Columns("Cost").Value = 0
            Grid.Columns("Amount").Value = 0
            'TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
            .MoveNext
         Wend
         .Close
         Grid.AddNew
         Grid.Columns("productid").Text = " "
         Grid.AllowAddNew = False
         Grid.Redraw = True
      End If
   End With
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
     MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
     Exit Sub
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
   If vIsNewRecord = True Then
      If CN.Execute("Select * from ProductionOutHeader where POutID = " & Val(TxtPOUTID.Text)).RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtPOUTID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
  RsBody.Filter = 0
  If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product for Sale", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
  End If
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
   CN.BeginTrans
   sSql = "select * from ProductionOutHeader where POutID=" & Val(TxtPOUTID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, CN, adOpenDynamic, adLockOptimistic
      If .BOF Then
         .AddNew
         !POutID = Val(TxtPOUTID.Text)
      End If
      !BillID = IIf(Val(TxtBillID.Text) = 0, Null, (TxtBillID.Text))
      !BillDate = IIf(Val(TxtBillID.Text) = 0, Null, DtpBillDate.DateValue)
      !POutDate = DtpPOUTDate.DateValue
      !FromStoreID = IIf(TxtFromStoreID.Text = "", Null, (TxtFromStoreID.Text))
      !ToStoreID = IIf(TxtToStoreID.Text = "", Null, (TxtToStoreID.Text))
      !Description = IIf(TxtDescription.Text = "", Null, TxtDescription.Text)
      !TotalAmount = Val(TxtTotalAmount.Text)
      !UserNo = vUser
      .Update
      .Close
   End With
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !POutID = Val(TxtPOUTID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   CN.CommitTrans
'   If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
'      Call BtnPrint_Click
'   End If
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
   RsBody.Open "Select * from ProductionOutBody where POutID=" & Val(TxtPOUTID.Text), CN, adOpenDynamic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = "select p.productname, b.*, UnitName from ProductionOutBody b join products p on p.productid = b.productid Left Outer Join Units u on u.UnitID = p.UnitID where IsNoCostProduct = 0 and POutID=" & Val(TxtPOUTID.Text)
      With CN.Execute(sSql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !ProductID
            Grid.Columns("Code").Text = !ProductID
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
            Grid.Columns("PurPrice").Value = !Price
            Grid.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            Grid.Columns("QtyLoose").Value = IIf(IsNull(!QtyLoose), "", !QtyLoose)
            Grid.Columns("Amount").Value = !Amount
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("ProductID").Text = " "
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
      TxtProductID.Enabled = True
      TxtCode.Enabled = True
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtFromStoreID.Enabled = True
      BtnFromStore.Enabled = True
      TxtToStoreID.Enabled = True
      BtnToStore.Enabled = True
      BtnProduct.Enabled = True
      TxtPOUTID.Text = FunGetMaxID()
      If DtpPOUTDate.Enabled And DtpPOUTDate.Visible Then DtpPOUTDate.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      TxtFromStoreID.Enabled = False
      BtnFromStore.Enabled = False
      TxtToStoreID.Enabled = False
      BtnToStore.Enabled = False
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtProductID.Enabled = True
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
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
         With CN.Execute("select * from ProductPacking where productid='" & TxtProductID.Text & "' and packingid=" & CmbPackName.ItemData(CmbPackName.ListIndex))
            TxtMultiplier.Text = IIf(.RecordCount = 0, "", !Multiplier)
         .Close
         End With
      End If
   End If
End Sub

Private Sub CmbPackName_Validate(Cancel As Boolean)
'If ActiveControl.Name <> CmbPackName.Name Then Exit Sub
'If CmbPackName.Text = "" Then Exit Sub: CmbPackName.SetFocus
'   If CmbPackName.ItemData(CmbPackName.ListIndex) = 1 Then
'      TxtQtyPack.Text = 1
'      TxtQtyPack.Enabled = False
'      TxtTotal.Enabled = True
'      TxtMultiplier.Text = 1
'      TxtTotal.SetFocus
'   Else
'      TxtQtyPack.Enabled = True
'      TxtTotal.Enabled = False
'      If Trim(TxtProductID.Text) <> "" Then
'         With CN.Execute("select * from Packings where packingid=" & CmbPackName.ItemData(CmbPackName.ListIndex))
'            TxtMultiplier.Text = IIf(.RecordCount = 0, 1, !Multiplier)
'            .Close
'         End With
'      End If
'   End If
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
   ElseIf KeyCode = vbKeyDelete Then
        mniRemoveRow_Click
        KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
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
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then CmbPackName.SetFocus Else If TxtCode.Enabled Then TxtCode.SetFocus
         Case TxtFromStoreID.Name: If FunSelectFromStore(ssFunctionKey, False) = True Then TxtToStoreID.SetFocus Else TxtFromStoreID.SetFocus
         Case TxtToStoreID.Name: If FunSelectToStore(ssFunctionKey, False) = True Then TxtDescription.SetFocus Else TxtToStoreID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtFromStoreID.SetFocus Else TxtOrganizationID.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtProductID.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtProductID.Name Then
         KeyCode = 0
         If BtnSave.Enabled Then BtnSave.SetFocus
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
    ShowPicture Me
    AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
    SetWindowText Me.hWnd, "Production Out"
   
    DtpPOUTDate.DateValue = Date
    
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
   
   
'   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
'   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = CN.Execute("Select isnull(max(POutID),0)+1 from ProductionOutHeader").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub SubClearFields()
  On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is ComboBox Then
      End If
   Next
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ProductID").Text = " "
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
       'Set RptReportViewer.Report = Nothing
      Set RsBody = Nothing
      Set FrmProductionIn = Nothing
   End If
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

Private Sub Grid_DblClick()
   Call Grid_LostFocus
End Sub

Private Sub Grid_GotFocus()
   Flag = True
   TxtProductID.Enabled = False
   TxtCode.Enabled = False
   BtnProduct.Enabled = False
'   CmbPackName.Enabled = False
   'TxtProductID.BackColor = TxtProductName.BackColor
   'TxtProductID.TabStop = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      TxtProductID.Text = ""
      TxtProductID.Enabled = True
      CmbPackName.Enabled = True
      BtnProduct.Enabled = True
      TxtCode.Text = ""
      TxtCode.Enabled = True
      TxtCode.SetFocus
   Else
      TxtProductID.Enabled = False
      TxtCode.Enabled = False
      BtnProduct.Enabled = False
'      CmbPackName.Enabled = False
      If Me.ActiveControl.Name = Grid.Name Then CmbPackName.SetFocus
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
   If Trim(Grid.Columns("ProductID").Text) = "" Then Exit Sub
   RsBody.Filter = "ProductID='" & TxtProductID.Text & "' And PackingID = " & IIf(CmbPackName.ItemData(CmbPackName.ListIndex) = 0, "Null", CmbPackName.ItemData(CmbPackName.ListIndex))
   If RsBody.RecordCount > 0 Then RsBody.Delete
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

Private Sub GetDataFromTexBoxesToGrid()
   On Error GoTo ErrorHandler
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
   
   
   '************* Last Pur Price Check **************
   
   RsBody.Filter = "ProductID='" & TxtProductID.Text & "'"
   If TxtProductID.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ProductID").Text = TxtProductID.Text
         RsBody!ProductID = TxtProductID.Text
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
                  Grid.Columns("ProductID").Text = TxtProductID.Text
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("Code").Text = TxtProductID.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
                  Grid.Columns("PackingID").Text = IIf(CmbPackName.ListIndex = 0, "", CmbPackName.ItemData(CmbPackName.ListIndex))
                  Grid.Columns("Pack").Text = TxtMultiplier.Text
                  Grid.Columns("QtyPack").Text = TxtQtyPack.Text
                  Grid.Columns("QtyLoose").Text = Val(TxtQtyLoose.Text)
                  Grid.Columns("PurPrice").Text = TxtPurPrice.Text
                  Grid.Columns("Amount").Text = TxtAmount.Text


                  RsBody!ProductID = TxtProductID.Text
                  'RsBody!Code = TxtCode.Text
                  RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
                  RsBody!Multiplier = Val(TxtMultiplier.Text)
                  RsBody!QtyPack = Val(TxtQtyPack.Text)
                  RsBody!QtyLoose = Val(TxtQtyLoose.Text)
                  RsBody!Price = Val(TxtPurPrice.Text)
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
      If TxtProductID.Enabled = True Then
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text)
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         
      End If
      .Columns("ProductID").Text = TxtProductID.Text
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("Code").Text = TxtProductID.Text
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Text = IIf(CmbPackName.ListIndex = 0, "", CmbPackName.ItemData(CmbPackName.ListIndex))
      .Columns("Pack").Text = TxtMultiplier.Text
      .Columns("QtyPack").Text = TxtQtyPack.Text
      .Columns("QtyLoose").Text = Val(TxtQtyLoose.Text)
      .Columns("PurPrice").Text = TxtPurPrice.Text
      .Columns("Amount").Text = TxtAmount.Text


      RsBody!ProductID = TxtProductID.Text
      'RsBody!Code = TxtCode.Text
      RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
      RsBody!Multiplier = Val(TxtMultiplier.Text)
      RsBody!QtyPack = Val(TxtQtyPack.Text)
      RsBody!QtyLoose = Val(TxtQtyLoose.Text)
      RsBody!Price = Val(TxtPurPrice.Text)
      RsBody!Amount = Val(TxtAmount.Text)
      .MoveLast
      If Trim(.Columns("ProductID").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("ProductID").Text = " "
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
   TxtCode.Text = ""
   TxtProductName.Text = ""
   CmbPackName.ListIndex = 0
   TxtMultiplier.Text = ""
   TxtQtyPack.Text = ""
   TxtQtyLoose.Text = ""
   TxtPurPrice.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtCode.Text = .Columns("Code").Text
      TxtProductID.Text = .Columns("ProductID").Text
      TxtProductName.Text = .Columns("ProductName").Text
      If Trim(.Columns("PackName").Text) = "" Then
         CmbPackName.ListIndex = 0
      Else
         CmbPackName.Text = .Columns("PackName").Text
      End If
      TxtMultiplier.Text = .Columns("Pack").Text
      TxtQtyPack.Text = .Columns("QtyPack").Text
      TxtQtyLoose.Text = .Columns("QtyLoose").Text
      TxtPurPrice.Text = .Columns("PurPrice").Text
      TxtAmount.Text = .Columns("Amount").Text
      With CN.Execute("select QtyLoose from currentstockStore where productid ='" & TxtProductID.Text & "' and storeid = " & TxtToStoreID.Text)
         If .RecordCount > 0 Then
            vQtyLoose = !QtyLoose
            LblStock.Caption = !QtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtProductID.Text & "')").Fields(0).Value
            LblStock.Visible = True
            LblStockCaption.Visible = True
         Else
            vQtyLoose = 0
            LblStock.Caption = 0
            LblStock.Visible = True
            LblStockCaption.Visible = True
         End If
      End With
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetProductionOut()
   On Error GoTo ErrorHandler
   sSql = "Select h.*, t.StoreName as ToStoreName, f.StoreName as FromStoreName FROM ProductionOutHeader h left outer join stores f on f.storeid = h.fromstoreid join stores t on t.storeid = h.ToStoreid where h.POutID=" & Val(TxtPOUTID.Text)
   With CN.Execute(sSql)
      If Not .BOF Then
          DtpPOUTDate.Date = !POutDate
          TxtToStoreID.Text = !ToStoreID
          TxtToStoreName.Text = !ToStoreName
          TxtFromStoreID.Text = !FromStoreID
          TxtFromStoreName.Text = !FromStoreName
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtTotalAmount.Text = !TotalAmount
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
      TxtPurPrice.Text = ""
   End If
End Sub

Private Sub TxtCode_Validate(Cancel As Boolean)
 If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = FunSelectProduct(ssValidate, False)
   If vTemp = False Then
      vTemp = FunSelectProduct(ssButton, False)
      Cancel = False
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtMultiplier_Change()
   Call SubCalculateBody
End Sub

Private Sub TxtPurPrice_Change()
   Call SubCalculateBody
End Sub

Private Sub TxtPurPrice_LostFocus()
   On Error GoTo ErrorHandler
   Select Case ActiveControl.Name
   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtQtyPack.Name, TxtQtyLoose.Name, TxtAmount.Text
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtQtyLoose_Change()
   Call SubCalculateBody
End Sub

Private Sub TxtQtyPack_Change()
    Call SubCalculateBody
End Sub

Private Sub TxtProductID_Change()
   If ActiveControl.Name <> TxtProductID.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
      TxtPurPrice.Text = ""
   End If
End Sub

Private Sub TxtProductID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtProductID_Validate(Cancel As Boolean)
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtProductID.Text) = "" Then Exit Sub
   vTemp = FunSelectProduct(ssValidate, False)
   If vTemp = False Then   '   vTemp = FunSelectProduct(ssButton, False)
      Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOrganization_Click()
   If FunSelectOrganization(ssButton, False) = True Then
      If TxtFromStoreID.Enabled Then TxtFromStoreID.SetFocus
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

