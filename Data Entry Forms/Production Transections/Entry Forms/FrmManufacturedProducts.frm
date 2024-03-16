VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmManufacturedProducts 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmManufacturedProducts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3840
      Left            =   13680
      TabIndex        =   28
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
         Height          =   3435
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Tag             =   "NC"
         Text            =   "FrmManufacturedProducts.frx":0ECA
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
         TabIndex        =   30
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9088
      TabIndex        =   8
      Top             =   9075
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
      MICON           =   "FrmManufacturedProducts.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7750
      TabIndex        =   5
      Top             =   9075
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
      MICON           =   "FrmManufacturedProducts.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5100
      TabIndex        =   7
      Top             =   9075
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
      MICON           =   "FrmManufacturedProducts.frx":0FE2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10426
      TabIndex        =   9
      Top             =   9075
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
      MICON           =   "FrmManufacturedProducts.frx":0FFE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6412
      TabIndex        =   6
      Top             =   9075
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
      MICON           =   "FrmManufacturedProducts.frx":101A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnManufacturedProduct 
      Height          =   330
      Left            =   3428
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4215
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
      MICON           =   "FrmManufacturedProducts.frx":1036
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   3788
      TabIndex        =   11
      Top             =   4215
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   8153
      TabIndex        =   14
      Top             =   1830
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
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
      Left            =   6368
      TabIndex        =   4
      Top             =   4215
      Width           =   990
      _ExtentX        =   1746
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1793
      TabIndex        =   3
      Top             =   4215
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3120
      Left            =   1793
      TabIndex        =   18
      Top             =   4530
      Width           =   5820
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   5
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
      stylesets(0).Picture=   "FrmManufacturedProducts.frx":1052
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
      Columns.Count   =   5
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
      Columns(3).Width=   1746
      Columns(3).Caption=   "Qty"
      Columns(3).Name =   "Qty"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "PackingID"
      Columns(4).Name =   "PackingID"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   10266
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridUsed 
      Height          =   3390
      Left            =   7658
      TabIndex        =   19
      Top             =   4245
      Width           =   5910
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   7
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
      stylesets(0).Picture=   "FrmManufacturedProducts.frx":106E
      AllowUpdate     =   0   'False
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
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   7
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ProductID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5715
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "ProductName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1191
      Columns(2).Caption=   "Qty"
      Columns(2).Name =   "Qty"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   5
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "PackingID"
      Columns(3).Name =   "PackingID"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1244
      Columns(4).Caption=   "Rate"
      Columns(4).Name =   "Rate"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   5
      Columns(4).FieldLen=   256
      Columns(5).Width=   1773
      Columns(5).Caption=   "Amount"
      Columns(5).Name =   "Amount"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   5
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "FinishedProductId"
      Columns(6).Name =   "FinishedProductId"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   10425
      _ExtentY        =   5980
      _StockProps     =   79
      Caption         =   "Used Products"
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpManufacturedDate 
      Height          =   315
      Left            =   3653
      TabIndex        =   1
      Top             =   2775
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   7013
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   2775
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
      Left            =   8048
      TabIndex        =   20
      Tag             =   "NC"
      Top             =   2775
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
      Left            =   7688
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2775
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
      MICON           =   "FrmManufacturedProducts.frx":108A
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtManufacturedID 
      Height          =   315
      Left            =   2123
      TabIndex        =   0
      Top             =   2775
      Width           =   1200
      _ExtentX        =   2117
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
   Begin SITextBox.Txt TxtTotalQty 
      Height          =   315
      Left            =   11858
      TabIndex        =   26
      Top             =   8010
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3736
      TabIndex        =   32
      Top             =   9075
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
      MICON           =   "FrmManufacturedProducts.frx":10A6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnBarCode 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2393
      TabIndex        =   33
      Top             =   9075
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
      MICON           =   "FrmManufacturedProducts.frx":10C2
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblAllStock 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Store Stock"
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
      Left            =   10710
      TabIndex        =   36
      Top             =   3510
      Visible         =   0   'False
      Width           =   1905
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
      Left            =   11070
      TabIndex        =   35
      Top             =   3105
      Width           =   1065
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
      Left            =   11190
      TabIndex        =   34
      Top             =   2790
      Width           =   720
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
      Left            =   11385
      TabIndex        =   31
      Top             =   540
      Width           =   435
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantites"
      Height          =   195
      Left            =   11768
      TabIndex        =   27
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured ID"
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
      Left            =   2123
      TabIndex        =   25
      Top             =   2580
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured Date"
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
      Left            =   3668
      TabIndex        =   24
      Top             =   2580
      Width           =   1650
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
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
      Left            =   7013
      TabIndex        =   23
      Top             =   2580
      Width           =   720
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
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
      Left            =   8048
      TabIndex        =   22
      Top             =   2580
      Width           =   1005
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured Products"
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
      TabIndex        =   17
      Top             =   270
      Width           =   4020
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      Left            =   6578
      TabIndex        =   16
      Top             =   4020
      Width           =   300
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   8153
      TabIndex        =   15
      Top             =   1635
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured Code"
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
      Left            =   1793
      TabIndex        =   13
      Top             =   4020
      Width           =   1680
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured Product Name"
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
      Left            =   3788
      TabIndex        =   12
      Top             =   4020
      Width           =   2445
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
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
Attribute VB_Name = "FrmManufacturedProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim vCounter1 As Integer
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean, vNegativeSale As Boolean
Dim ssql As String, vQtyLoose As Double, vMultiplier As Double
Dim vStrSQL As String, vManufactureID As Long
'----------------------------------

Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectStore = False: Exit Function
        TxtStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Stores where islock = 0 and StoreID=" & Val(TxtStoreID.Text)
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

Private Function FunSelectFinishedProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchFinishedProduct.Show vbModal, Me
      If SchFinishedProduct.ParaOutID = "" Then FunSelectFinishedProduct = False: Exit Function
      TxtCode.Text = SchFinishedProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCode.Text) = "" Then Exit Function
    If TxtCode.Text = "" Then FunSelectFinishedProduct = False: Exit Function
    
    vStrSQL = " SELECT p.productid, Code, ProductName " & vbCrLf _
           + " from ProductProcessInfoHeader f inner join Products p on f.finishedproductid = p.productid" & vbCrLf _
           + " left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')"
           
  
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         
         If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock(" & TxtCode.Text & ",Null,0,0,0,0,0,0,'" & DtpManufacturedDate.DateValue + 1 & "',0),0)"
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
            LblAllStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & TxtCode.Text & ",(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(TxtCode.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & "Loose"
            LblAllStock.Visible = True
         Else
            LblAllStock.Visible = False
         End If
         
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & Val(TxtCode.Text)
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtCode.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpManufacturedDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & Val(TxtCode.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(TxtCode.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
         
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         FunSelectFinishedProduct = True
         
         .Close
         Exit Function
      Else
         TxtProductID.Text = ""
         TxtCode.Text = ""
         TxtProductName.Text = ""
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         FunSelectFinishedProduct = False
         .Close
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnBarCode_Click()
   On Error GoTo ErrorHandler
   If BtnSave.Enabled Then BtnSave_Click
   If vIsNewRecord = False Then
      vManufactureID = Val(TxtManufacturedID.Text)
   End If

   ssql = " Select b.ProductID, ProductName, QtyLoose, P.ProductID, GroupID" & vbCrLf _
      + " from ManufacturedProductsBody b inner join Products p on b.PRoductID = p.ProductID" & vbCrLf _
      + " where ManufacturedID = " & vManufactureID
'   sSql = "select b.ProductID, Code, ProductName from ProductBarcodes b inner join Products p on p.productid = b.ProductID where len(code) = 11 and code like '110%'"
   
   Dim i As Integer
   With cn.Execute(ssql)
      FrmMultiBarcodes.SubClearFields
      FrmMultiBarcodes.TxtTotQty.Text = "0"
      For i = 1 To .RecordCount
         FrmMultiBarcodes.Grid.Columns("ID").Text = !Productid
         FrmMultiBarcodes.Grid.Columns("Name").Text = !ProductName
         FrmMultiBarcodes.Grid.Columns("GroupID").Text = !GroupID
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
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniManufacturedProducts", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
   cn.BeginTrans   'Products Body
   Grid.Redraw = False
   Grid.MoveFirst
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("ProductID").Text) <> "" Then
         cn.Execute "Delete from ManufacturedProductsBody where ManufacturedID=" & Val(TxtManufacturedID.Text) & " and ProductID='" & Grid.Columns("Productid").Text & "'"
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
'   'Products Used
'   GridUsed.Redraw = False
'   GridUsed.MoveFirst
'   For vCounter = 1 To GridUsed.Rows
'      If Trim(GridUsed.Columns("Productid").Text) <> "" Then
'         CN.Execute "Delete from ManufacturedProductsUsed where ManufacturedID=" & Val(TxtManufacturedID.Text) & " and ProductID='" & GridUsed.Columns("Productid").Text & "'"
'      End If
'      GridUsed.MoveNext
'   Next vCounter
'   GridUsed.RemoveAll
'   GridUsed.Redraw = True
   'Header
   Call ActivityLog("Manufactured Products", eDelete, TxtManufacturedID.Text)
   cn.Execute "Delete from ManufacturedProductsHeader where ManufacturedID= " & (TxtManufacturedID.Text)
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchManufacturedProducts.Show vbModal
   If SchManufacturedProducts.ParaOutManufacturedID = 0 Then Exit Sub
   TxtManufacturedID.Text = SchManufacturedProducts.ParaOutManufacturedID
   GetManufacturedProduct
End Sub

Private Sub BtnPrint_Click()
On Error GoTo ErrorHandler
   vStrSQL = "Select ManufacturedDate, Code as FinishedProductID, P.ProductName as FinishedProductName, B.QtyLoose as FinishedQty, U.*, PU.ProductName from ManufacturedProductsHeader H" & vbCrLf _
            + " Inner Join ManufacturedProductsBody B ON B.ManufacturedID = H.ManufacturedID" & vbCrLf _
            + " Inner Join ManufacturedProductsUsed U ON U.ManufacturedID = H.ManufacturedID" & vbCrLf _
            + " Inner Join Products P on P.ProductID = Code" & vbCrLf _
            + " Inner Join Products PU on PU.ProductID = U.ProductID" & vbCrLf _
            + " Where H.ManufacturedID = " & Val(TxtManufacturedID.Text)

    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, cn, adOpenDynamic, adLockReadOnly
'
    Set RptReportViewer.Report = New CrpManufacturedProduct
    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
    
    With cn.Execute("Select CompanyName,Address,City,PhoneNo,email from Company")
      If .RecordCount > 0 Then
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(IsNull(!CompanyName), "", CStr(!CompanyName))
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(IsNull(!Address), "", !Address) & IIf(IsNull(!City), "", ", " & !City & ".")
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(IsNull(!PhoneNo), "", CStr(!PhoneNo))
      End If
      .Close
   End With
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue cn.Execute("Select Name from Manufacturer").Fields(0).Value
    
'   RptReportViewer.Report.ParameterFields(1).AddCurrentValue objRegistry.CompanyName
'   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(objRegistry.CompanyAddress = "", "", objRegistry.CompanyAddress) & IIf(objRegistry.CompanyCity = "", "", ", " & objRegistry.CompanyCity)
'   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(objRegistry.CompanyPhoneNo = "", ".", " Phone # " & objRegistry.CompanyPhoneNo)
'   RptReportViewer.Report.ParameterFields(4).AddCurrentValue objRegistry.DevelopedBy
'   RptReportViewer.Report.SelectPrinter objRegistry.DriverName, objRegistry.DeviceName, objRegistry.Port
'   RptReportViewer.Report.PrintOut False

    RptReportViewer.Show
    'RptReportViewer.Report.PrintOut False
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub BtnManufacturedProduct_Click()
   If FunSelectFinishedProduct(ssButton, True) = True Then
      TxtQty.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniManufacturedProducts", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
    MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
    Exit Sub
  End If
'  '/******** dummy retriction ***********/
'   If vIsNewRecord = False Then
'      MsgBox "You are not authorized to modify a record", vbCritical, "Error"
'   End If
'   '/************************************/
'  Header Validation
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      TxtStoreID.SetFocus
      Exit Sub
   End If
   If vIsNewRecord = True Then
      If cn.Execute("select * from ManufacturedProductsHeader where ManufacturedId=" & TxtManufacturedID.Text).RecordCount > 0 Then
         MsgBox "Manufactured ID Already Exist.", vbExclamation, Me.Caption
         TxtManufacturedID.SetFocus
         Exit Sub
      End If
   End If
   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one Manufactured Product", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
   End If
   cn.BeginTrans
'   ' delete from manufacture product used
'   With CN.Execute("Select * from ManufacturedProductsUsed where ManufacturedID=" & Val(TxtManufacturedID.Text))
'      While Not .EOF
'         CN.Execute "Delete from ManufacturedProductsUsed where ManufacturedID=" & !ManufacturedID & " and ProductID='" & !ProductID & "'"
'         .MoveNext
'      Wend
'      .Close
'   End With
   vManufactureID = Val(TxtManufacturedID.Text)
   If vIsNewRecord = False Then Call ActivityLog("Manufactured Products", eEdit, TxtManufacturedID.Text)
   ssql = "select * from ManufacturedProductsHeader where ManufacturedID=" & Val(TxtManufacturedID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, cn, adOpenDynamic, adLockOptimistic
      If .BOF Then
         .AddNew
         !ManufacturedID = Val(TxtManufacturedID.Text)
      End If
      !ManufacturedDate = DtpManufacturedDate.DateValue
      !StoreID = Val(TxtStoreID.Text)
      !UserNo = ObjUserSecurity.UserNo
      .Update
      .Close
   End With
   
  'Body Validation
  ' validation has been performed when a row is added to the grid
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ManufacturedID = TxtManufacturedID.Text
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   If vIsNewRecord = True Then Call ActivityLog("Manufactured Products", eAdd, TxtManufacturedID.Text)
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
   RsBody.Open "Select * from ManufacturedProductsBody where ManufacturedID=" & Val(TxtManufacturedID.Text), cn, adOpenDynamic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.productname, code, b.* from ManufacturedProductsBody b join products p on p.productid = b.productid where ManufacturedID=" & Val(TxtManufacturedID.Text)
      With cn.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = !Code
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !QtyLoose
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

Private Sub PopulateDataToUsed()
   ssql = "select ProductName, u.* from ManufacturedProductsUsed u join products p on p.productid = u.productid where ManufacturedID=" & Val(TxtManufacturedID.Text)
   With cn.Execute(ssql)
      GridUsed.Redraw = False
      GridUsed.MoveFirst
      GridUsed.RemoveAll
      GridUsed.AllowAddNew = True
      TxtTotalQty.Text = 0
      While Not .EOF
         GridUsed.AddNew
         GridUsed.Columns("ProductID").Text = !Productid
         GridUsed.Columns("ProductName").Text = !ProductName
         GridUsed.Columns("Qty").Value = !QtyLoose
         GridUsed.Columns("Rate").Value = !Rate
         GridUsed.Columns("Amount").Value = !Amount
         TxtTotalQty.Text = Val(TxtTotalQty.Text) + Val(!QtyLoose)
         .MoveNext
      Wend
      .Close
   End With
   GridUsed.AllowAddNew = False
   GridUsed.Redraw = True
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
      LblStock.Visible = False
      LblAllStock.Visible = False
      LblStockCaption.Visible = False
      TxtManufacturedID.Text = FunGetMaxID
      TxtCode.Enabled = True
      BtnManufacturedProduct.Enabled = True
      BtnStore.Enabled = True
      TxtStoreID.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnPrint.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtManufacturedID.Enabled = True
      If TxtManufacturedID.Enabled And TxtManufacturedID.Visible Then TxtManufacturedID.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnStore.Enabled = False
      BtnPrint.Enabled = True
      TxtStoreID.Enabled = False
      TxtCode.Enabled = True
      BtnManufacturedProduct.Enabled = True
      TxtManufacturedID.Enabled = False
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
      TxtCode.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
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
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
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
         Case vbKeyB
            If BtnBarCode.Enabled Then BtnBarCode_Click
            KeyCode = 0
'         Case vbKeyP
'            If BtnPrint.Enabled Then BtnPrint_Click
'            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtStoreID.SetFocus
         Case TxtCode.Name: If FunSelectFinishedProduct(ssFunctionKey, False) = True Then TxtQty.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
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

Private Sub GridUsed_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock(" & GridUsed.Columns("ProductID").Text & ",Null,0,0,0,0,0,0,'" & DtpManufacturedDate.DateValue + 1 & "',0),0)"
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
            LblAllStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & GridUsed.Columns("ProductID").Text & ",(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(TxtCode.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & "Loose"
            LblAllStock.Visible = True
         Else
            LblAllStock.Visible = False
         End If
         
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & Val(GridUsed.Columns("ProductID").Text)
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(GridUsed.Columns("ProductID").Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpManufacturedDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & Val(GridUsed.Columns("ProductID").Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(GridUsed.Columns("ProductID").Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
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
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Manufactured Products"
   HelpLocation Me
   TxtStoreID.Text = ObjRegistry.StoreID
   FunSelectStore ssValidate, True
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   vNegativeSale = ObjRegistry.NegativeSale
   
   FormStatus = NewMode
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Public Sub SubClearFields()
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
   Grid.Columns("Code").Text = " "
   Grid.Update
   GridUsed.CancelUpdate
   GridUsed.RemoveAll
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
    Set FrmManufacturedProducts = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
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
   BtnManufacturedProduct.Enabled = False
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
      BtnManufacturedProduct.Enabled = True
      TxtCode.SetFocus
   Else
      vBm = Grid.Bookmark
      TxtCode.Enabled = False
      BtnManufacturedProduct.Enabled = False
      TxtQty.SetFocus
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(Grid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub GridUsed_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotalQty.Text = Val(TxtTotalQty.Text) - Grid.Columns("Qty").Value
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub CalculateTotal()
On Error GoTo ErrorHandler
   Dim bm As Variant
   Dim i As Integer
   GridUsed.Redraw = False
   GridUsed.MoveFirst
   TxtTotalQty.Text = ""
   For i = 0 To GridUsed.Rows - 1
      bm = GridUsed.GetBookmark(i)
      TxtTotalQty.Text = Val(TxtTotalQty.Text) + GridUsed.Columns(2).CellValue(bm)
   Next i
   GridUsed.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
   RsBody.Filter = "Code='" & TxtCode.Text & "'"
   If RsBody.RecordCount > 0 Then RsBody.Delete
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   RemoveDataToGridUsed
   Grid.DeleteSelected
   Grid.Refresh
   RsBody.Filter = 0
   Grid.MoveLast
   Call CalculateTotal
   GetDataBackFromGridToTexBoxes
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Public Sub GetDataFromTexBoxesToGrid()
   On Error GoTo ErrorHandler
   Dim vRowCounter As Integer
   If Trim(TxtCode.Text) = "" Then
      TxtCode.SetFocus
      Exit Sub
   End If
   If Val(TxtQty.Text) = 0 Then
      TxtQty.SetFocus
      Exit Sub
   End If
   If AddDataToGridUsed = False Then
      Call SubClearDetailArea
      Grid.MoveLast
      Grid.MoveNext
      If TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
   End If
'   RsBody.Filter = "Code = '" & TxtCode.Text & "'"
'   If TxtCode.Enabled Then
'      If RsBody.RecordCount = 0 Then
'      ' Add new record
'            Grid.Columns("ProductID").Text = ""
'            Grid.Columns("Code").Text = ""
'            Call SubClearDetailArea
'            RsBody.CancelUpdate
'            If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
'            Exit Sub
'         End If
'      Else
'      ' update old record if add new record
'      End If
'   Else
'      'update old record
'   End If
   Grid.Bookmark = vBm

   RsBody.Filter = "ProductID = " & TxtProductID.Text
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ProductID").Text = TxtProductID.Text
         Grid.Columns("Code").Text = TxtCode.Text
         RsBody!Productid = TxtProductID.Text
         RsBody!Code = TxtCode.Text
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vRowCounter = 1 To Grid.Rows
               If Grid.Columns("ProductID").Text = TxtProductID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
'                  If AddDataToGridUsed = False Then
'                     Grid.Columns("ProductID").Text = ""
'                     Grid.Columns("Code").Text = ""
'                     Call SubClearDetailArea
'                     RsBody.CancelUpdate
'                     If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
'                     Exit Sub
'                  End If
                  TxtQty.Text = Val(TxtQty.Text) + Grid.Columns("Qty").Value
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("Qty").Value = Val(TxtQty.Text)
                  RsBody!QtyLoose = Val(TxtQty.Text)
                  Grid.MoveLast
                  Call SubClearDetailArea
                  TxtCode.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vRowCounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearDetailArea
         Grid.MoveLast
         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   Grid.Redraw = False
   With Grid
      'If TxtCode.Enabled = False Then If AddDataToGridUsed = False Then Exit Sub
      'If TxtCode.Enabled = True Then
'      If AddDataToGridUsed = False Then
'         Grid.Columns("ProductID").Text = ""
'         Grid.Columns("Code").Text = ""
'         Call SubClearDetailArea
'         RsBody.CancelUpdate
'         If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
'         Grid.Redraw = True
'         Exit Sub
'      End If
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("Qty").Text = Val(TxtQty.Text)
      RsBody!QtyLoose = Val(TxtQty.Text)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearDetailArea
   If TxtCode.Visible Then TxtCode.SetFocus
   Call CalculateTotal
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Function AddDataToGridUsed() As Boolean
   If Trim(TxtProductID.Text) = "" Then Exit Function
   'GridUsed.Redraw = False
   Dim Flag1 As Boolean
   With cn.Execute("exec spFinishedProducts " & Val(TxtProductID.Text))
      For vCounter1 = 1 To .RecordCount
         Flag1 = True
         GridUsed.MoveFirst
          
'         vStrSQL = " select QtyLoose from CurrentStockStore " & vbCrLf _
'         + " where ProductID = '" & !Productid & "' and StoreID = " & Val(TxtStoreID.Text)
'         vMultiplier = !QtyLoose
'         With cn.Execute(vStrSQL)
'            If .RecordCount > 0 Then
'               vQtyLoose = !QtyLoose
'            Else
'               vQtyLoose = 0
'            End If
'            .Close
'         End With

         
'          If ObjRegistry.ShowSavedStock = True Then
'            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & !Productid
'            With cn.Execute(vStrSQL)
'               If .RecordCount > 0 Then
'                  vQtyLoose = .Fields(0).Value
'               Else
'                  vQtyLoose = 0
'               End If
'            End With
'         Else
'            vStrSQL = "select isnull(dbo.FunStock(" & !Productid & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpManufacturedDate.DateValue + 1 & "',0),0)"
'            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
'         End If
         

         
         
         
'         vStrSQL = " Select QtyLoose from ManufacturedProductsUsed" & vbCrLf _
'                 + " Where ProductID = '" & !ProductID & "' and ManufacturedID = " & Val(TxtManufacturedID.Text)
'         With CN.Execute(vStrSQL)
'            If .RecordCount > 0 Then
'               vQtyLoose = vQtyLoose + !QtyLoose
'            End If
'            .Close
'         End With
                  
         For vCounter = 1 To GridUsed.Rows
            If GridUsed.Columns("ProductID").Text = !Productid Then
               If vNegativeSale = False Then
                  If (Val(vQtyLoose) - (Val(GridUsed.Columns("Qty").Value)) - (Val(TxtQty.Text) * vMultiplier) + IIf(TxtCode.Enabled = False, Val(Grid.Columns("Qty").Value) * vMultiplier, 0)) < 0 Then
                     MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                     'Grid.MoveLast
                     GridUsed.Redraw = True
                     AddDataToGridUsed = False
                     Exit Function
                  End If
               End If
               GridUsed.Columns("Rate").Value = !Rate
               If TxtCode.Enabled = True Then
                  GridUsed.Columns("Qty").Value = GridUsed.Columns("Qty").Value + (!QtyLoose * Val(TxtQty.Text))
               Else
                  GridUsed.Columns("Qty").Value = GridUsed.Columns("Qty").Value - (!QtyLoose * Grid.Columns("Qty").Value) + (!QtyLoose * Val(TxtQty.Text))
               End If
               GridUsed.Columns("Amount").Value = GridUsed.Columns("Rate").Value * GridUsed.Columns("Qty").Value
               Flag1 = False
            End If
            GridUsed.MoveNext
         Next vCounter
         GridUsed.MoveLast
          If Flag1 = True Then
            GridUsed.AllowAddNew = True
            If vNegativeSale = False Then
               If (Val(vQtyLoose) - (Val(TxtQty.Text) * vMultiplier)) < 0 Then
                  MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                  AddDataToGridUsed = False
                  GridUsed.Redraw = True
                  Exit Function
               End If
            End If
            GridUsed.AddNew
            GridUsed.Columns("ProductID").Text = !Productid
            GridUsed.Columns("ProductName").Text = !ProductName
            GridUsed.Columns("Rate").Value = !Rate
            GridUsed.Columns("Qty").Value = (!QtyLoose * Val(TxtQty.Text))
            GridUsed.Columns("Amount").Value = GridUsed.Columns("Rate").Value * GridUsed.Columns("Qty").Value
            GridUsed.Update
            GridUsed.AllowAddNew = False
         End If
         .MoveNext
      Next vCounter1
      .Close
      GridUsed.Redraw = True
   End With
   AddDataToGridUsed = True
   If BtnSave.Enabled Then FormStatus = SelectionMode
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub RemoveDataToGridUsed()
   If Trim(Grid.Columns("ProductID").Text) = "" Then Exit Sub
   GridUsed.Redraw = False
   GridUsed.MoveFirst
   With cn.Execute("exec spFinishedProducts " & Grid.Columns("ProductID").Text)
      For vCounter1 = 1 To .RecordCount
         GridUsed.MoveFirst
         For vCounter = 1 To GridUsed.Rows
            If GridUsed.Columns("ProductID").Text = !Productid Then
               GridUsed.Columns("Qty").Value = GridUsed.Columns("Qty").Value - (!QtyLoose * Grid.Columns("Qty").Value)
               GridUsed.Columns("Amount").Value = GridUsed.Columns("Rate").Value * GridUsed.Columns("Qty").Value
            End If
            If GridUsed.Columns("Qty").Value = 0 Then
               GridUsed.SelBookmarks.RemoveAll
               GridUsed.SelBookmarks.Add GridUsed.Bookmark
               GridUsed.DeleteSelected
               Grid.Refresh
            Else
               GridUsed.MoveNext
            End If
         Next vCounter
         .MoveNext
      Next vCounter1
   End With
   GridUsed.Redraw = True
   If BtnSave.Enabled Then FormStatus = SelectionMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtCode.Enabled = True
   BtnManufacturedProduct.Enabled = True
   TxtCode.Text = ""
   TxtProductName.Text = ""
   TxtQty.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("Code").Text
      TxtProductName.Text = .Columns("ProductName").Text
      TxtQty.Text = .Columns("Qty").Text
      
      If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock(" & TxtCode.Text & ",Null,0,0,0,0,0,0,'" & DtpManufacturedDate.DateValue + 1 & "',0),0)"
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
            LblAllStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & TxtCode.Text & ",(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(TxtCode.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & "Loose"
            LblAllStock.Visible = True
         Else
            LblAllStock.Visible = False
         End If
         
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & Val(TxtCode.Text)
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtCode.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpManufacturedDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & Val(TxtCode.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(TxtCode.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetManufacturedProduct()
   On Error GoTo ErrorHandler
   ssql = "select ManufacturedID, ManufacturedDate, h.StoreID, StoreName FROM ManufacturedProductsHeader h join Stores s on h.storeid = s.storeid where ManufacturedID=" & Val(TxtManufacturedID.Text)
   With cn.Execute(ssql)
      If Not .BOF Then
          TxtManufacturedID.Text = !ManufacturedID
          DtpManufacturedDate.DateValue = !ManufacturedDate
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
      End If
      .Close
   End With
   Call PopulateDataToGrid
   Call PopulateDataToUsed
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtCode_Change()
   If TxtCode.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then TxtProductName.Text = ""
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtCode_Validate(Cancel As Boolean)
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = FunSelectFinishedProduct(ssValidate, False)
   If vTemp = False Then   '   vTemp = FunSelectFinishedProduct(ssButton, False)
      Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtQty_LostFocus()
   Select Case ActiveControl.Name
   Case TxtCode.Name
      Exit Sub
   End Select
   GetDataFromTexBoxesToGrid
End Sub

Public Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select isnull(max(ManufacturedID),0)+1 from ManufacturedProductsHeader").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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
