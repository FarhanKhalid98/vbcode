VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmSetCurrentStock 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmSetCurrentStock.frx":0000
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
      Height          =   2085
      Left            =   13200
      TabIndex        =   33
      Top             =   1560
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
         Height          =   1815
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   34
         Tag             =   "NC"
         Text            =   "FrmSetCurrentStock.frx":0ECA
         Top             =   360
         Width           =   3930
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
         TabIndex        =   35
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.CheckBox ChkAdd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Opening Add"
      Height          =   255
      Left            =   2025
      TabIndex        =   32
      Top             =   2100
      Width           =   1335
   End
   Begin JeweledBut.JeweledButton BtnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5805
      TabIndex        =   8
      Top             =   9450
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
      MICON           =   "FrmSetCurrentStock.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7125
      TabIndex        =   9
      Top             =   9450
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
      MICON           =   "FrmSetCurrentStock.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8445
      TabIndex        =   11
      Top             =   9450
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
      MICON           =   "FrmSetCurrentStock.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   11745
      TabIndex        =   7
      Top             =   3135
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
      Left            =   12525
      TabIndex        =   12
      Top             =   3135
      Width           =   1110
      _ExtentX        =   1958
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
      Left            =   4260
      TabIndex        =   2
      Top             =   3135
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
      Left            =   5220
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3135
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
      MICON           =   "FrmSetCurrentStock.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   5580
      TabIndex        =   3
      Top             =   3135
      Width           =   3120
      _ExtentX        =   5503
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5175
      Left            =   1740
      TabIndex        =   10
      Top             =   3450
      Width           =   11895
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
      stylesets(0).Picture=   "FrmSetCurrentStock.frx":0F8E
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
      Columns.Count   =   14
      Columns(0).Width=   1693
      Columns(0).Caption=   "Store ID"
      Columns(0).Name =   "StoreID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2752
      Columns(1).Caption=   "Store Name"
      Columns(1).Name =   "StoreName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2328
      Columns(2).Caption=   "Code"
      Columns(2).Name =   "Code"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   5530
      Columns(3).Caption=   "Product Name"
      Columns(3).Name =   "ProductName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2619
      Columns(4).Caption=   "BarCodes"
      Columns(4).Name =   "BarCodes"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1296
      Columns(5).Caption=   "Qt.Loose"
      Columns(5).Name =   "QtyLoose"
      Columns(5).Alignment=   1
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "PackingID"
      Columns(6).Name =   "PackingID"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1429
      Columns(7).Caption=   "SalePrice"
      Columns(7).Name =   "SalePrice"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1376
      Columns(8).Caption=   "Pur Price"
      Columns(8).Name =   "PurPrice"
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1508
      Columns(9).Caption=   "Amount"
      Columns(9).Name =   "Amount"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "ProductID"
      Columns(10).Name=   "ProductID"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Caption=   "GroupID"
      Columns(11).Name=   "GroupID"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Caption=   "GroupName"
      Columns(12).Name=   "GroupName"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Caption=   "OpeningADD"
      Columns(13).Name=   "OpeningADD"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   11
      Columns(13).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20981
      _ExtentY        =   9128
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
   Begin SITextBox.Txt TxtQtyLoose 
      Height          =   315
      Left            =   10185
      TabIndex        =   5
      Top             =   3135
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   1740
      TabIndex        =   1
      Tag             =   "NC"
      Top             =   3135
      Width           =   585
      _ExtentX        =   1032
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
      Left            =   2685
      TabIndex        =   20
      Tag             =   "NC"
      Top             =   3135
      Width           =   1575
      _ExtentX        =   2778
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
      Left            =   2325
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3135
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
      MICON           =   "FrmSetCurrentStock.frx":0FAA
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   2235
      TabIndex        =   24
      Top             =   1485
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
   Begin SITextBox.Txt TxtSalePrice 
      Height          =   315
      Left            =   10935
      TabIndex        =   6
      Top             =   3135
      Width           =   810
      _ExtentX        =   1429
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
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3045
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2460
      Width           =   360
      _ExtentX        =   635
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
      MICON           =   "FrmSetCurrentStock.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   2265
      TabIndex        =   0
      Top             =   2460
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   3
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
      IntegralPoint   =   3
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtBarcodes 
      Height          =   315
      Left            =   8715
      TabIndex        =   4
      Tag             =   "NC"
      Top             =   3135
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   25
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
      IntegralPoint   =   3
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtSaleID 
      Height          =   315
      Left            =   9645
      TabIndex        =   37
      Tag             =   "NC"
      Top             =   2460
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
   Begin SITextBox.Txt TxtReturnID 
      Height          =   315
      Left            =   11625
      TabIndex        =   38
      Tag             =   "NC"
      Top             =   2460
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
   Begin SITextBox.Txt TxtPurID 
      Height          =   315
      Left            =   7665
      TabIndex        =   39
      Tag             =   "NC"
      Top             =   2460
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpSaleDate 
      Height          =   315
      Left            =   10320
      TabIndex        =   40
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpReturnDate 
      Height          =   315
      Left            =   12300
      TabIndex        =   41
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpPurDate 
      Height          =   315
      Left            =   8340
      TabIndex        =   42
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
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   3405
      TabIndex        =   43
      Top             =   2460
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur ID"
      Height          =   195
      Left            =   7665
      TabIndex        =   49
      Top             =   2250
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale ID"
      Height          =   195
      Left            =   9645
      TabIndex        =   48
      Top             =   2250
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Return ID"
      Height          =   195
      Left            =   11625
      TabIndex        =   47
      Top             =   2250
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Date"
      Height          =   195
      Left            =   10320
      TabIndex        =   46
      Top             =   2250
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      Height          =   195
      Left            =   12375
      TabIndex        =   45
      Top             =   2250
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Date"
      Height          =   195
      Left            =   8340
      TabIndex        =   44
      Top             =   2250
      Width           =   630
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
      Left            =   12990
      TabIndex        =   36
      Top             =   1650
      Width           =   435
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      Height          =   195
      Left            =   1725
      TabIndex        =   31
      Top             =   2475
      Width           =   435
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
      Left            =   5910
      TabIndex        =   29
      Top             =   2085
      Width           =   1035
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
      Left            =   5865
      TabIndex        =   28
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "BarCodes"
      Height          =   195
      Left            =   8685
      TabIndex        =   27
      Top             =   2910
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Price"
      Height          =   195
      Left            =   10965
      TabIndex        =   26
      Top             =   2910
      Width           =   720
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   2205
      TabIndex        =   25
      Top             =   1260
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   2685
      TabIndex        =   23
      Top             =   2940
      Width           =   840
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   1740
      TabIndex        =   22
      Top             =   2940
      Width           =   585
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   5580
      TabIndex        =   19
      Top             =   2940
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   4260
      TabIndex        =   18
      Top             =   2940
      Width           =   375
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (Loose)"
      Height          =   195
      Left            =   10095
      TabIndex        =   17
      Top             =   2910
      Width           =   810
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Current Stock"
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
      TabIndex        =   15
      Top             =   270
      Width           =   3075
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Price"
      Height          =   195
      Left            =   11820
      TabIndex        =   14
      Top             =   2940
      Width           =   645
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   12600
      TabIndex        =   13
      Top             =   2940
      Width           =   540
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   60
      Width           =   345
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmSetCurrentStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vCounter As Integer
Dim vIsNewRecord As Boolean
Dim RsBody As New ADODB.Recordset
Dim Flag As Boolean
Dim ssql As String
Dim vMax As String
Dim VStrSQL As String
Dim vIsNewRow As Boolean

Private Sub SubCalculateBody()
   TxtAmount.Text = Val(TxtPurPrice.Text) * Val(TxtQtyLoose.Text)
End Sub

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   'FormStatus = NewMode
   PopulateDataToGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtProductName.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnGroup_Click()
   If FunSelectGroup(ssButton, False) = True Then
      TxtStoreID.SetFocus
   Else
      TxtGroupID.SetFocus
   End If
End Sub

Private Sub TxtGroupID_Change()
   If TxtGroupID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   If TxtGroupName.Text <> "" Then TxtGroupName.Text = ""
End Sub

Private Sub TxtGroupID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtGroupID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectGroup(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectGroup(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchGroup.Show vbModal, Me
        If SchGroup.ParaOutGroupID = "" Then FunSelectGroup = False: Exit Function
        TxtGroupID.Text = SchGroup.ParaOutGroupID
    End If
    '---------------------------
    If Len(TxtGroupID.Text) < 3 Then
      TxtGroupID.Text = Right("000" + CStr(Val(TxtGroupID.Text)), 3)
    End If
    VStrSQL = " Select * FROM Groups where GroupID='" & TxtGroupID.Text & "'"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtGroupName.Text = !GroupName
          TxtProductID.Text = vMax
          TxtCode.Text = TxtProductID.Text
          FunSelectGroup = True
          .Close
          Exit Function
      Else
          FunSelectGroup = False
          .Close
          TxtGroupID.Text = ""
          TxtGroupName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxID() As String
   On Error GoTo ErrorHandler
   FunGetMaxID = CN.Execute("Select Right('00000' + Cast(isnull(max(cast(ProductId as smallint)),0) + 1 as varchar),5) From Products").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectStore = False: Exit Function
        TxtStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Stores where StoreID=" & Val(TxtStoreID.Text)
    With CN.Execute(VStrSQL)
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
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCode.Text) = "" Then Exit Function
    If Len(TxtCode.Text) <= 5 Then
      TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
    End If
    'If Len(TxtCode.Text) < 5 Then
    '  TxtCode.Text = "006" + Right("0000" + CStr(Val(TxtCode.Text)), 4)
    'End If
    If TxtGroupID.Text <> "" Then Exit Function
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    VStrSQL = " SELECT p.productid, Code, ProductName, PurPrice, RetailPrice " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
   
   With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPurPrice.Text = !PurPrice
         TxtSalePrice.Text = !RetailPrice
         LblStock.Visible = True
         LblStockCaption.Visible = True
         SubCalculateBody
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
'      Else
'         FunSelectProduct = False
'         .Close
'         TxtCode.Text = ""
'         TxtProductID.Text = ""
'         TxtProductName.Text = ""
'         TxtPurPrice.Text = ""
'         TxtSalePrice.Text = ""
'         LblStock.Visible = False
'         LblStockCaption.Visible = False
'         If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   With CN.Execute("select QtyLoose from openingstock where productid='" & TxtProductID.Text & "' and StoreID=" & TxtStoreID.Text)
      If .RecordCount > 0 Then
         LblStock.Caption = !QtyLoose + Val(CN.Execute("select dbo.FunStock('" & TxtProductID.Text & "'," & TxtStoreID.Text & "," & Val(TxtPurID.Text) & ",'" & DtpPurDate.DateValue & "'," & Val(TxtSaleID.Text) & ",'" & DtpSaleDate.DateValue & "'," & Val(TxtReturnID.Text) & ",'" & DtpReturnDate.DateValue & "') ").Fields(0).Value)
      Else
         LblStock.Caption = 0
      End If
   End With
    
    VStrSQL = "Select * from ProductBarcodes where ProductID = '" & TxtProductID.Text & "' and len(Code)>7"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtBarcodes.Text = !Code
      Else
         TxtBarcodes.Text = ""
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtCode.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
          Set frmObj = Nothing
      Next
      Set RsBody = Nothing
      Set FrmSetCurrentStock = Nothing
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
      TxtPurPrice.Text = ""
      TxtSalePrice.Text = ""
   End If
End Sub

Private Sub TxtCode_Validate(Cancel As Boolean)
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   If TxtGroupID.Text <> "" Then Exit Sub
   vTemp = FunSelectProduct(ssValidate, False)
   If vTemp = False Then
      vTemp = FunSelectProduct(ssButton, False)
      Cancel = False
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  'If VIsPosted And ObjUserSecurity.IsAdministrator = False Then
  '  MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
  '  Exit Sub
  'End If
'  Header Validation
   RsBody.Filter = ""
   If Grid.Rows = 1 Then
      MsgBox "Enter atleast one product to save", vbExclamation, "Alert"
      TxtProductID.SetFocus
      Exit Sub
   End If
  'Saving record
   CN.BeginTrans
   Dim vOpeningStock As Double, vDiff As Double
   Dim vCost As Double, vSetCost As Double
   Dim vQty As Double, vSetQty As Double
   Dim vCurrentCost As Double
   
   Grid.MoveFirst
   For vCounter = 1 To Grid.Rows - 1
      'Product
      With CN.Execute("select * from Products where productid='" & Grid.Columns("ProductID").Text & "'")
         If .RecordCount > 0 Then
            CN.Execute ("Update Products set ProductName='" & Replace(Grid.Columns("ProductName").Text, "'", "''") & "', PurPrice=" & Val(Grid.Columns("PurPrice").Text) & ", RetailPrice=" & Val(Grid.Columns("SalePrice").Text) & " where productid='" & Grid.Columns("ProductID").Text & "'")
         Else
            ssql = ("Insert into Products (GroupID, ProductID, ProductName, PurPrice, RetailPrice) values ('" & Grid.Columns("GroupID").Text & "','" & Grid.Columns("ProductID").Text & "','" & Replace(Grid.Columns("ProductName").Text, "'", "''") & "'," & Val(Grid.Columns("PurPrice").Text) & "," & Val(Grid.Columns("SalePrice").Text) & ")")
            CN.Execute (ssql)
            ssql = ("Insert into ProductBarCodes (ProductID, Code) values ('" & Grid.Columns("ProductID").Text & "','" & Grid.Columns("BarCodes").Text & "')")
            CN.Execute (ssql)
         End If
      End With
        
      'opening stock Store
      With CN.Execute("select * from openingstock where productid='" & Grid.Columns("ProductID").Text & "' and StoreID=" & Grid.Columns("StoreID").Text)
         If .RecordCount > 0 Then
            If Grid.Columns("OpeningAdd").Value = False Then
               vDiff = ((CN.Execute("select dbo.FunCurrentStock('" & Grid.Columns("ProductID").Text & "'," & Grid.Columns("StoreID").Text & ") ").Fields(0).Value) - (CN.Execute("select dbo.FunStock('" & Grid.Columns("ProductID").Text & "'," & Grid.Columns("StoreID").Text & "," & Val(TxtPurID.Text) & ",'" & DtpPurDate.DateValue & "'," & Val(TxtSaleID.Text) & ",'" & DtpSaleDate.DateValue & "'," & Val(TxtReturnID.Text) & ",'" & DtpReturnDate.DateValue & "') ").Fields(0).Value))
               vOpeningStock = Val(Grid.Columns("QtyLoose").Value) - (CN.Execute("select dbo.FunStock('" & Grid.Columns("ProductID").Text & "'," & Grid.Columns("StoreID").Text & "," & Val(TxtPurID.Text) & ",'" & DtpPurDate.DateValue & "'," & Val(TxtSaleID.Text) & ",'" & DtpSaleDate.DateValue & "'," & Val(TxtReturnID.Text) & ",'" & DtpReturnDate.DateValue & "') ").Fields(0).Value)
               CN.Execute ("Update OpeningStock set QtyLoose=" & vOpeningStock & ", PurPrice=" & Val(Grid.Columns("PurPrice").Text) & ",Amount=" & vOpeningStock * Val(Grid.Columns("Purprice").Text) & " where productid='" & Grid.Columns("ProductID").Text & "' and StoreID=" & Grid.Columns("StoreID").Text)
               CN.Execute ("Update CurrentStockStore set QtyLoose=" & Val(Grid.Columns("QtyLoose").Value) + vDiff & " where productid='" & Grid.Columns("ProductID").Text & "' and StoreID=" & Grid.Columns("StoreID").Text)
            Else
               vOpeningStock = Val(Grid.Columns("QtyLoose").Value) + !QtyLoose
               CN.Execute ("Update OpeningStock set QtyLoose=" & vOpeningStock & ", PurPrice=" & Val(Grid.Columns("PurPrice").Text) & ",Amount=" & vOpeningStock * Val(Grid.Columns("Purprice").Text) & " where productid='" & Grid.Columns("ProductID").Text & "' and StoreID=" & Grid.Columns("StoreID").Text)
            End If
         Else
            vOpeningStock = Val(Grid.Columns("QtyLoose").Value) '- (CN.Execute("select dbo.FunCurrentStock('" & Grid.Columns("ProductID").Text & "'," & Grid.Columns("StoreID").Text & "," & Val(TxtPurID.Text) & ",'" & DtpPurDate.DateValue & "'," & Val(TxtSaleID.Text) & ",'" & DtpSaleDate.DateValue & "'," & Val(TxtReturnID.Text) & ",'" & DtpReturnDate.DateValue & "') ").Fields(0).Value)
            ssql = ("Insert into OpeningStock (ProductID, StoreID, QtyLoose, PurPrice, Amount) values ('" & Grid.Columns("ProductID").Text & "'," & Grid.Columns("StoreID").Text & "," & vOpeningStock & "," & Val(Grid.Columns("PurPrice").Text) & "," & vOpeningStock * Val(Grid.Columns("PurPrice").Text) & ")")
            CN.Execute (ssql)
            CN.Execute ("Update CurrentStockStore set QtyLoose=" & Val(Grid.Columns("QtyLoose").Value) & " where productid='" & Grid.Columns("ProductID").Text & "' and StoreID=" & Grid.Columns("StoreID").Text)
         End If
         vOpeningStock = (CN.Execute("select isnull(sum(qtyloose),0) from openingstock where productid='" & Grid.Columns("ProductID").Text & "'").Fields(0).Value) + (CN.Execute("select dbo.FunCurrentStock('" & Grid.Columns("ProductID").Text & "',Null)").Fields(0).Value)
         CN.Execute ("Update CurrentStock set QtyLoose=" & vOpeningStock & " where productid='" & Grid.Columns("ProductID").Text & "'")
      End With
      ssql = "INSERT into ActivityLog(userno,FormType,EntryDate,Description,isnew,isedit,isdelete) values(" & ObjUserSecurity.UserNo & ",'Opening Stock', GetDate(),'ProductID = " & Grid.Columns("ProductID").Text & ", QtyLoose = " & Val(Grid.Columns("QtyLoose").Value) & ", Price = " & Val(Grid.Columns("PurPrice").Text) & "',1,0,0)"
      CN.Execute ssql
      Grid.MoveNext
   Next vCounter
   'Body Validation
   ' validation has been performed when a row is added to the grid
   '-------------------------------------------------------------------------
   CN.CommitTrans
   'CN.Execute "exec SPCurrentStock"
   Grid.Redraw = True
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
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
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Set Current Stock"
   HelpLocation Me
   LblStock.Visible = False
   LblStockCaption.Visible = False
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
      'If Me.ActiveControl.Name = Grid.Name And Grid.AddItemRowIndex(Grid.Bookmark) = Grid.Rows - 1 Then
      '   Grid.Update
      'End If
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtStoreID.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then TxtCode.SetFocus
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtProductName.SetFocus  'CmbPackName.SetFocus
      End Select
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtCode.Enabled Then TxtCode.SetFocus: Call SubClearDetailArea
   ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtProductID.Name Then
      KeyCode = 0
      BtnSave.SetFocus
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
      Case vbKeyS
          If BtnSave.Enabled Then BtnSave_Click
          KeyCode = 0
      Case vbKeyH
         FraHelp.ZOrder 0
         FraHelp.Visible = True
         KeyCode = 0
      Case vbKeyW
         If BtnClear.Enabled Then BtnClear_Click
         KeyCode = 0
      Case vbKeyQ
          If BtnClose.Enabled Then BtnClose_Click
          KeyCode = 0
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub ImgExit_Click()
   Unload Me
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
      'If RsBody.State = adStateOpen Then RsBody.Close
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      'PopulateDataToGrid
      vMax = FunGetMaxID
      TxtCode.Enabled = True
      vIsNewRow = True
   Case Is = OpenMode
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      vIsNewRow = True
   Case Is = ChangeMode
      BtnSave.Enabled = True
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

Private Sub PopulateDataToGrid()
   'If RsBody.State = adStateOpen Then RsBody.Close
   'RsBody.Open "select * from currentstockstore s inner join currentstock c on c.productid = s.productid Where c.qtyloose < 0 And storeid = 1", CN, adOpenStatic, adLockBatchOptimistic
   'If RsBody.RecordCount > 0 Then
      '================================================
      ssql = "select p.ProductID, Productname, abs(c.Qtyloose)  as QtyLoose, Cost * abs(c.Qtyloose) as Amount, Cost, StoreID from currentstockstore s inner join currentstock c on c.productid = s.productid inner join products p on p.productid = c.productid where c.qtyloose < 0 and storeid = 1"
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            Grid.AddNew
            Grid.Columns("GroupID").Text = ""
            Grid.Columns("OpeningAdd").Value = 0
            Grid.Columns("StoreID").Text = !StoreID
            Grid.Columns("Code").Text = !Productid
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("QtyLoose").Value = !QtyLoose
            Grid.Columns("PurPrice").Value = !Cost
            Grid.Columns("Amount").Value = !Amount
            Grid.Update
            .MoveNext
         Wend
         .Close
         'Grid.Row = 0
      End With
      'Grid.RemoveAll
      'Grid.AllowAddNew = True
      'Grid.AddNew
      'Grid.Columns("Code").Text = " "
      'Grid.AllowAddNew = False
      Grid.Redraw = True
'   End If
   'Grid.FirstRow = 0
End Sub

Private Sub GetDataFromTexBoxesToGrid()
On Error GoTo ErrorHandler
   If Trim(TxtStoreID.Text) = "" Then
      If TxtStoreID.Enabled = True Then TxtStoreID.SetFocus
      Exit Sub
   End If
   If Trim(TxtCode.Text) = "" Then
      'MsgBox "Enter Group ID.", vbExclamation, "Alert"
      If TxtCode.Enabled = True Then TxtCode.SetFocus
      Exit Sub
   End If
   If Trim(TxtQtyLoose.Text) = "" Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      TxtQtyLoose.SetFocus
      Exit Sub
   End If
      
   '-------------------------------------------------------------------

   'RsBody.Filter = "ProductID='" & TxtProductID.Text & "'"
   'If vIsNewRow Then
   '   If RsBody.RecordCount = 0 Then
   '      RsBody.AddNew
  '       RsBody!Productid = TxtProductID.Text
  '    Else
  '       MsgBox "The record already exist"
  '       SubClearDetailArea
  '       If TxtProductID.Enabled Then TxtProductID.SetFocus
  '       Exit Sub
  '    End If
  ' End If
   Grid.Redraw = False
   With Grid
      .Columns("StoreID").Text = TxtStoreID.Text
      .Columns("StoreName").Text = TxtStoreName.Text
      .Columns("GroupID").Text = TxtGroupID.Text
      .Columns("GroupName").Text = TxtGroupName.Text
      .Columns("ProductID").Text = TxtProductID.Text
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("Code").Text = TxtCode.Text
      .Columns("BarCodes").Text = TxtBarcodes.Text
      .Columns("QtyLoose").Text = TxtQtyLoose.Text
      .Columns("SalePrice").Text = TxtSalePrice.Text
      .Columns("PurPrice").Text = TxtPurPrice.Text
      .Columns("Amount").Text = TxtAmount.Text
      .Columns("OpeningAdd").Value = ChkAdd.Value
  '    RsBody!Qty = Val(TxtQty.Text)
  '    RsBody!PurPrice = Val(TxtPurPrice.Text)
  '    RsBody!Amount = Val(TxtAmount.Text)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   If Val(TxtGroupID.Text) <> 0 Then vMax = vMax + 1
   Call SubClearDetailArea
   TxtCode.SetFocus
 '  vIsNewRow = True
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtPurPrice_Change()
   Call SubCalculateBody
End Sub

Private Sub TxtPurPrice_LostFocus()
   Select Case ActiveControl.Name
      Case TxtGroupID.Name, TxtCode.Name, TxtBarcodes.Name, TxtProductName.Name, TxtQtyLoose.Name, TxtSalePrice.Name
   End Select
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtQtyLoose_Change()
   Call SubCalculateBody
End Sub

Private Sub SubClearDetailArea()
   TxtCode.Enabled = True
   TxtGroupID.Text = ""
   TxtGroupName.Text = ""
   TxtCode.Text = ""
   TxtProductName.Text = ""
   TxtBarcodes.Text = ""
   TxtQtyLoose.Text = ""
   TxtPurPrice.Text = ""
   TxtSalePrice.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtGroupID.Text = .Columns("GroupID").Text
      TxtGroupName.Text = .Columns("GroupName").Text
      TxtStoreID.Text = .Columns("StoreID").Text
      TxtStoreName.Text = .Columns("StoreName").Text
      TxtCode.Text = .Columns("Code").Text
      TxtProductID.Text = .Columns("ProductID").Text
      TxtProductName.Text = .Columns("ProductName").Text
      TxtBarcodes.Text = .Columns("BarCodes").Text
      TxtQtyLoose.Text = .Columns("QtyLoose").Text
      TxtSalePrice.Text = .Columns("SalePrice").Text
      TxtPurPrice.Text = .Columns("PurPrice").Text
      TxtAmount.Text = .Columns("Amount").Text
      ChkAdd.Value = .Columns("OpeningAdd").Text
   End With
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
   TxtStoreID.Enabled = False
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("Code").Text) = "" Then
      TxtCode.Enabled = True
      TxtStoreID.Enabled = True
      TxtCode.SetFocus
      BtnStore.Enabled = True
      BtnProduct.Enabled = True
      vIsNewRow = True
   Else
      TxtCode.Enabled = False
      TxtStoreID.Enabled = False
      BtnStore.Enabled = False
      BtnProduct.Enabled = False
      TxtQtyLoose.SetFocus
      vIsNewRow = False
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(Grid.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Code").Text) = "" Or Trim(Grid.Columns("StoreID").Text) = "" Then Exit Sub
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   'RsBody.Filter = "ProductID='" & Grid.Columns("ProductID").Text & "'"
   'If RsBody.RecordCount > 0 Then RsBody.Delete
   'RsBody.Filter = ""
   Grid.DeleteSelected
   Grid.SelBookmarks.RemoveAll
   Grid.Refresh
   Grid.MoveLast
   GetDataBackFromGridToTexBoxes
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtStoreID_Change()
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
End Sub

Private Sub TxtStoreID_GotFocus()
'   TxtCode.Text = ""
End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtStoreName.Text <> "" Then Exit Sub
   If TxtStoreID.Text = "" Then Exit Sub
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
