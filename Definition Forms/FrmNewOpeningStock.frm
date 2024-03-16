VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmOpeningStock 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmNewOpeningStock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbSession 
      Height          =   315
      Left            =   8070
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3330
      Width           =   1845
   End
   Begin VB.ComboBox CmbOrganization 
      Height          =   315
      Left            =   1733
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   2130
      Width           =   2145
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
      Height          =   2175
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
         Height          =   1725
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   32
         Tag             =   "NC"
         Text            =   "FrmNewOpeningStock.frx":0ECA
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
         TabIndex        =   33
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   6225
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3330
      Width           =   1845
   End
   Begin JeweledBut.JeweledButton BtnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5738
      TabIndex        =   9
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
      MICON           =   "FrmNewOpeningStock.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7058
      TabIndex        =   10
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
      MICON           =   "FrmNewOpeningStock.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8378
      TabIndex        =   11
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
      MICON           =   "FrmNewOpeningStock.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   12015
      TabIndex        =   13
      Top             =   3330
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
      Left            =   12795
      TabIndex        =   14
      Top             =   3330
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
      Left            =   2775
      TabIndex        =   1
      Top             =   3330
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
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   4095
      TabIndex        =   18
      Top             =   3330
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
      Height          =   5625
      Left            =   165
      TabIndex        =   12
      Top             =   3645
      Width           =   13740
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   16
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
      stylesets(0).Picture=   "FrmNewOpeningStock.frx":0F72
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
      Columns.Count   =   16
      Columns(0).Width=   1826
      Columns(0).Caption=   "Store ID"
      Columns(0).Name =   "StoreID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2778
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
      Columns(3).Width=   3757
      Columns(3).Caption=   "Product Name"
      Columns(3).Name =   "ProductName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3254
      Columns(4).Caption=   "Pack Name"
      Columns(4).Name =   "PackName"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3254
      Columns(5).Caption=   "Session Name"
      Columns(5).Name =   "SessionName"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   953
      Columns(6).Caption=   "Pack"
      Columns(6).Name =   "Pack"
      Columns(6).Alignment=   1
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   4
      Columns(6).FieldLen=   256
      Columns(7).Width=   1323
      Columns(7).Caption=   "Qt.Pack"
      Columns(7).Name =   "QtyPack"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   1429
      Columns(8).Caption=   "Qt.Loose"
      Columns(8).Name =   "QtyLoose"
      Columns(8).Alignment=   1
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   4
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "PackingID"
      Columns(9).Name =   "PackingID"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1376
      Columns(10).Caption=   "Pur Price"
      Columns(10).Name=   "PurPrice"
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   4
      Columns(10).FieldLen=   256
      Columns(11).Width=   1508
      Columns(11).Caption=   "Amount"
      Columns(11).Name=   "Amount"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   5
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "ProductID"
      Columns(12).Name=   "ProductID"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   1852
      Columns(13).Caption=   "BatchNo"
      Columns(13).Name=   "BatchNo"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   2090
      Columns(14).Caption=   "ExpiryDate"
      Columns(14).Name=   "ExpiryDate"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(15).Width=   3200
      Columns(15).Caption=   "SessionID"
      Columns(15).Name=   "SessionID"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   24236
      _ExtentY        =   9922
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
      Left            =   9915
      TabIndex        =   6
      Top             =   3330
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
      Left            =   11205
      TabIndex        =   8
      Top             =   3330
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
      DecimalPoint    =   3
      IntegralPoint   =   6
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtQtyPack 
      Height          =   315
      Left            =   10455
      TabIndex        =   7
      Top             =   3330
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
      Left            =   165
      TabIndex        =   0
      Top             =   3330
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
      Left            =   1200
      TabIndex        =   25
      Top             =   3330
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
      Left            =   840
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3330
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
      MICON           =   "FrmNewOpeningStock.frx":0F8E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   10118
      TabIndex        =   29
      Top             =   1425
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
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   3735
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3330
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
      MICON           =   "FrmNewOpeningStock.frx":0FAA
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBatchNo 
      Height          =   315
      Left            =   3180
      TabIndex        =   2
      Top             =   3015
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
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpExpiryDate 
      Height          =   315
      Left            =   3990
      TabIndex        =   3
      Top             =   3015
      Width           =   1215
      _Version        =   65543
      _ExtentX        =   2143
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
      EditMode        =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1740
      TabIndex        =   40
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label LblSession 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Session"
      Height          =   195
      Left            =   8070
      TabIndex        =   39
      Top             =   3135
      Width           =   555
   End
   Begin VB.Label LblBatchNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Batch No"
      Height          =   195
      Left            =   4748
      TabIndex        =   38
      Top             =   2580
      Width           =   675
   End
   Begin VB.Label LblExpiryDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date"
      Height          =   195
      Left            =   5558
      TabIndex        =   37
      Top             =   2580
      Width           =   810
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
      Left            =   11475
      TabIndex        =   34
      Top             =   495
      Width           =   435
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   10118
      TabIndex        =   30
      Top             =   1230
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   1200
      TabIndex        =   28
      Top             =   3135
      Width           =   840
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   165
      TabIndex        =   27
      Top             =   3135
      Width           =   585
   End
   Begin VB.Label LblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   5205
      TabIndex        =   24
      Top             =   3135
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   2775
      TabIndex        =   23
      Top             =   3135
      Width           =   375
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   9915
      TabIndex        =   22
      Top             =   3135
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   6315
      TabIndex        =   21
      Top             =   3135
      Width           =   840
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (Loose)"
      Height          =   195
      Left            =   11205
      TabIndex        =   20
      Top             =   3135
      Width           =   810
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (Pack)"
      Height          =   195
      Left            =   10410
      TabIndex        =   19
      Top             =   3135
      Width           =   750
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Stock"
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
      Width           =   2595
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Price"
      Height          =   195
      Left            =   12090
      TabIndex        =   16
      Top             =   3135
      Width           =   645
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   12870
      TabIndex        =   15
      Top             =   3135
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
Attribute VB_Name = "FrmOpeningStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vCounter As Integer
Dim vUnitPrice As Double
Dim vIsNewRecord As Boolean
Dim RsBody As New ADODB.Recordset
Dim Flag As Boolean
Dim ssql As String
Dim vStrSQL As String
Dim vIsNewRow As Boolean

Private Sub SubCalculateBody()
   TxtAmount.Text = Round((Val(vUnitPrice)) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 2)
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

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      If TxtBatchNo.Visible Then TxtBatchNo.SetFocus Else TxtQtyLoose.SetFocus
   Else
      If TxtCode.Enabled Then TxtCode.SetFocus
   End If
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
      SchProduct.ParaInWhere = ""
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
        vStrSQL = " SELECT p.productid, Code, ProductName, PurPrice, RetailPrice, PurDiscPC, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where p.productid = " & Val(TxtCode.Text) & " or code='" & TxtCode.Text & "'"
 
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPurPrice.Text = !PurPrice
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
         SubCalculateBody
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
         TxtPurPrice.Text = ""
         TxtAmount.Text = ""
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      If TxtStoreID.Enabled Then TxtStoreID.SetFocus
   End If
End Sub

Private Sub CmbOrganization_Change()
    If ActiveControl.Name <> CmbOrganization.Name Then Exit Sub
    Call PopulateDataToGrid
End Sub

Private Sub CmbOrganization_Click()
    If CmbOrganization.Visible = False Then Exit Sub
    If ActiveControl.Name <> CmbOrganization.Name Then Exit Sub
    Call PopulateDataToGrid
End Sub

Private Sub CmbPackName_Click()
   If CmbPackName.Text = "" Then
      TxtMultiplier.Enabled = False
      TxtQtyPack.Enabled = False
      TxtMultiplier.Text = ""
      TxtQtyPack.Text = ""
      TxtPurPrice.Text = Round(vUnitPrice, 3)
   Else
      TxtMultiplier.Enabled = True
      TxtQtyPack.Enabled = True
      If Trim(TxtCode.Text) <> "" Then
         With CN.Execute("select * from ProductPacking where ProductID = '" & TxtCode.Text & "' and packingid = " & CmbPackName.ItemData(CmbPackName.ListIndex))
            TxtMultiplier.Text = IIf(.RecordCount = 0, "", !Multiplier)
            If Val(TxtMultiplier.Text) <> 0 Then
               TxtPurPrice.Text = Round(vUnitPrice * !Multiplier, 3)
            Else
               TxtPurPrice.Text = Round(vUnitPrice, 3)
            End If
            .Close
         End With
      End If
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
      Set FrmOpeningStock = Nothing
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

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   'If VIsPosted And ObjUserSecurity.IsAdministrator = False Then
   '  MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
   '  Exit Sub
   'End If
   'Header Validation
   RsBody.Filter = ""
'   If Grid.Rows = 1 Then
'      MsgBox "Enter atleast one product to save", vbExclamation, "Alert"
'      TxtProductID.SetFocus
'      Exit Sub
'   End If
   
   'Body Validation
   ' validation has been performed when a row is added to the grid
  
   'Saving record
   CN.BeginTrans
   '-------------------------------------------------------------------------
   RsBody.UpdateBatch
   CN.CommitTrans
   'CN.Execute "exec SpcurrentStock"
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
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Opening Stock"
   HelpLocation Me
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
   
   TxtBatchNo.Visible = ObjRegistry.BatchNoVisible
   LblBatchNo.Visible = ObjRegistry.BatchNoVisible
   DtpExpiryDate.Visible = ObjRegistry.BatchNoVisible
   LblExpiryDate.Visible = ObjRegistry.BatchNoVisible

   If LblExpiryDate.Visible = False Then LblProductName.Left = TxtProductName.Left
   
   With CN.Execute("Select * from Organizations")
      While Not .EOF
         CmbOrganization.AddItem !OrganizationName
         CmbOrganization.ItemData(CmbOrganization.NewIndex) = !OrganizationID
         .MoveNext
      Wend
      .Close
   End With
   If CmbOrganization.ListCount > 0 Then CmbOrganization.ListIndex = 0
   
   CmbSession.Clear
   CmbSession.AddItem ""
   With CN.Execute("select * from Sessions")
      While Not .EOF
         CmbSession.AddItem !SessionName
         CmbSession.ItemData(CmbSession.NewIndex) = !SessionID
         .MoveNext
      Wend
      .Close
   End With
   CmbSession.ListIndex = vSessionID
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
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtCode.Enabled Then TxtCode.SetFocus: Call SubClearDetailArea
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then If TxtCode.Enabled Then TxtCode.SetFocus
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then If TxtBatchNo.Visible Then TxtBatchNo.SetFocus Else TxtQtyLoose.SetFocus
      End Select
   ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtProductID.Name Then
      KeyCode = 0
      If BtnSave.Enabled Then BtnSave.SetFocus
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
      If UCase(Me.ActiveControl.Name) Like "TXT*" Or UCase(Me.ActiveControl.Name) Like "DTP*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
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
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      PopulateDataToGrid
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
   
   DtpExpiryDate.DateValue = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   SubClearDetailArea
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from OpeningStock Where OrganizationID = " & CmbOrganization.ItemData(CmbOrganization.ListIndex) & " order by productid", CN, adOpenStatic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      '================================================
      ssql = "select pr.Productname, os.*, StoreName, PackingName, SessionName from OpeningStock os join Products pr on os.Productid = pr.Productid join stores s on s.storeid = os.storeid left outer join packings p on p.packingid = os.packingid left outer join sessions ss on ss.sessionid = os.sessionID Where os.OrganizationID = " & CmbOrganization.ItemData(CmbOrganization.ListIndex) & "  order by os.productid"
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
        
         While Not .EOF
            Grid.AddNew
            Grid.Columns("StoreID").Text = !StoreID
            Grid.Columns("StoreName").Text = !StoreName
            Grid.Columns("BatchNo").Text = IIf(IsNull(!BatchNo), "", !BatchNo)
            Grid.Columns("ExpiryDate").Text = IIf(IsNull(!ExpiryDate), "", !ExpiryDate)
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = !Productid
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("PackName").Text = IIf(IsNull(!PackingName), "", !PackingName)
            Grid.Columns("PackingID").Text = IIf(IsNull(!PackingID), "", !PackingID)
            Grid.Columns("SessionName").Text = IIf(IsNull(!SessionName), "", !SessionName)
            Grid.Columns("SessionID").Text = IIf(IsNull(!SessionID), "", !SessionID)
            Grid.Columns("Pack").Value = IIf(IsNull(!Multiplier), 0, !Multiplier)
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), 0, !QtyPack)
            Grid.Columns("QtyLoose").Value = !QtyLoose
            Grid.Columns("PurPrice").Value = !PurPrice
            Grid.Columns("Amount").Value = !Amount
            .MoveNext
         Wend
         .Close
         'Grid.Row = 0
      End With
      Grid.AllowAddNew = True
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True

'   End If
   Grid.FirstRow = 0
   Grid.MoveLast
   Grid.MoveNext
'   Call SubClearFields
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
   If Val(TxtQtyLoose.Text) = 0 And Val(TxtQtyPack.Text) = 0 Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      If TxtQtyPack.Enabled Then TxtQtyPack.SetFocus
      Exit Sub
   End If
      
   '-------------------------------------------------------------------
   If Trim(Grid.Columns("Productid").Text) = "" Then
      RsBody.Filter = "ProductID = " & TxtProductID.Text & " and BatchNo = " & IIf(Trim(TxtBatchNo.Text) = "", "null", "'" & Trim(TxtBatchNo.Text) & "'") & " and StoreID = " & Val(TxtStoreID.Text)
   Else
      RsBody.Filter = "ProductID = " & Grid.Columns("Productid").Text & " and BatchNo = " & IIf(Grid.Columns("BatchNo").Text = "", "null", "'" & Grid.Columns("BatchNo").Text & "'") & " and StoreID = " & Val(Grid.Columns("StoreID").Text)
   End If
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         RsBody!OrganizationID = CmbOrganization.ItemData(CmbOrganization.ListIndex)
         RsBody!Productid = TxtProductID.Text
         RsBody!StoreID = TxtStoreID.Text
         RsBody!BatchNo = Trim(TxtBatchNo.Text)
      Else
         MsgBox "The record already exist"
         SubClearDetailArea
            If TxtCode.Enabled = True Then TxtCode.SetFocus
         Exit Sub
      End If
   End If
   Grid.Redraw = False
   With Grid
      .Columns("BatchNo").Text = Trim(TxtBatchNo.Text)
      .Columns("ExpiryDate").Text = DtpExpiryDate.DateValue
      .Columns("StoreID").Text = TxtStoreID.Text
      .Columns("StoreName").Text = TxtStoreName.Text
      .Columns("ProductID").Text = TxtProductID.Text
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("Code").Text = TxtProductID.Text
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Text = IIf(CmbPackName.ListIndex = 0, "", CmbPackName.ItemData(CmbPackName.ListIndex))
      .Columns("SessionName").Text = CmbSession.Text
      .Columns("SessionID").Text = IIf(CmbSession.ListIndex = 0, "", CmbSession.ItemData(CmbSession.ListIndex))
      .Columns("Pack").Value = Val(TxtMultiplier.Text)
      .Columns("QtyPack").Value = Val(TxtQtyPack.Text)
      .Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
      .Columns("PurPrice").Value = Val(TxtPurPrice.Text)
      .Columns("Amount").Value = Val(TxtAmount.Text)
      
      
      RsBody!Productid = TxtProductID.Text
      'RsBody!Code = TxtCode.Text
      RsBody!BatchNo = IIf(Trim(TxtBatchNo.Text) = "", Null, Trim(TxtBatchNo.Text))
      RsBody!ExpiryDate = IIf(DtpExpiryDate.DateValue = "", Null, DtpExpiryDate.DateValue)
      RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
      RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
      RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
      RsBody!QtyLoose = Val(TxtQtyLoose.Text)
      RsBody!PurPrice = Val(TxtPurPrice.Text)
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
   If TxtCode.Enabled = True Then TxtCode.SetFocus
   vIsNewRow = True
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtPurPrice_Change()
   If TxtPurPrice.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtPurPrice.Name Then Exit Sub
   If Val(TxtMultiplier.Text) <> 0 Then
      vUnitPrice = Val(TxtPurPrice.Text) / Val(TxtMultiplier.Text)
   Else
      vUnitPrice = Val(TxtPurPrice.Text)
   End If
   Call SubCalculateBody
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

Private Sub TxtQtyLoose_Change()
   Call SubCalculateBody
End Sub

Private Sub TxtQtyLoose_LostFocus()
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtQtyPack_Change()
   Call SubCalculateBody
End Sub

Private Sub SubClearDetailArea()
   TxtCode.Enabled = True
   TxtCode.Text = ""
   TxtBatchNo.Text = ""
   DtpExpiryDate.DateValue = ""
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
      TxtStoreID.Text = .Columns("StoreID").Text
      TxtStoreName.Text = .Columns("StoreName").Text
      TxtBatchNo.Text = .Columns("BatchNo").Text
      DtpExpiryDate.DateValue = .Columns("ExpiryDate").Text
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
      If Val(TxtMultiplier.Text) = 0 Then
         vUnitPrice = IIf(.Columns("PurPrice").Text = "", 0, .Columns("PurPrice").Text)
      Else
         vUnitPrice = .Columns("PurPrice").Text / Val(.Columns("Pack").Text)
      End If
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
      If TxtCode.Enabled = True Then TxtCode.SetFocus
      BtnStore.Enabled = True
      BtnProduct.Enabled = True
      vIsNewRow = True
   Else
      TxtCode.Enabled = False
      TxtStoreID.Enabled = False
      BtnStore.Enabled = False
      BtnProduct.Enabled = False
      If CmbPackName.Enabled Then CmbPackName.SetFocus
      vIsNewRow = False
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
   RsBody.Filter = "ProductID = " & Grid.Columns("ProductID").Text & " and BatchNo = " & IIf(Trim(TxtBatchNo.Text) = "", "null", "'" & Trim(TxtBatchNo.Text) & "'") & " and StoreID = " & Grid.Columns("StoreID").Text
   If RsBody.RecordCount > 0 Then RsBody.Delete
   RsBody.Filter = ""
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
