VERSION 5.00
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmCustomOrderBooking 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
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
      Left            =   11970
      TabIndex        =   24
      Top             =   450
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
         TabIndex        =   25
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
         TabIndex        =   26
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6690
      TabIndex        =   13
      Top             =   7845
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
      MICON           =   "FrmCustomOrderBooking.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   5363
      TabIndex        =   10
      Top             =   7845
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
      MICON           =   "FrmCustomOrderBooking.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2723
      TabIndex        =   12
      Top             =   7845
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
      MICON           =   "FrmCustomOrderBooking.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8010
      TabIndex        =   14
      Top             =   7845
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
      MICON           =   "FrmCustomOrderBooking.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4043
      TabIndex        =   11
      Top             =   7845
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
      MICON           =   "FrmCustomOrderBooking.frx":0070
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnCustomProduct 
      Height          =   330
      Left            =   1515
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2970
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
      MICON           =   "FrmCustomOrderBooking.frx":008C
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomProductName 
      Height          =   315
      Left            =   1875
      TabIndex        =   16
      Top             =   2985
      Width           =   2700
      _ExtentX        =   4763
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
   Begin SITextBox.Txt TxtCustomProductCode 
      Height          =   315
      Left            =   75
      TabIndex        =   4
      Top             =   2985
      Width           =   1440
      _ExtentX        =   2540
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
      Left            =   60
      TabIndex        =   7
      Top             =   3300
      Width           =   7890
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   6
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
      stylesets(0).Picture=   "FrmCustomOrderBooking.frx":00A8
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
      Columns.Count   =   6
      Columns(0).Width=   3201
      Columns(0).Caption=   "Custom Product Code"
      Columns(0).Name =   "Code"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4763
      Columns(1).Caption=   "Custom Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1376
      Columns(2).Caption=   "Qty"
      Columns(2).Name =   "Qty"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "PackingID"
      Columns(3).Name =   "PackingID"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1693
      Columns(4).Caption=   "Price"
      Columns(4).Name =   "Price"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2408
      Columns(5).Caption=   "Amount"
      Columns(5).Name =   "Amount"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   13917
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridDetail 
      Height          =   3420
      Left            =   7950
      TabIndex        =   8
      Top             =   3000
      Width           =   3990
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   5
      stylesets.count =   3
      stylesets(0).Name=   "SelectedCol"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   12713983
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "FrmCustomOrderBooking.frx":00C4
      stylesets(1).Name=   "Select"
      stylesets(1).ForeColor=   16777215
      stylesets(1).BackColor=   8388608
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "FrmCustomOrderBooking.frx":00E0
      stylesets(2).Name=   "SelectedRow"
      stylesets(2).ForeColor=   16777215
      stylesets(2).BackColor=   8388608
      stylesets(2).HasFont=   -1  'True
      BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(2).Picture=   "FrmCustomOrderBooking.frx":00FC
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
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
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ID"
      Columns(0).Name =   "ID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2461
      Columns(1).Caption=   "Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3069
      Columns(2).Caption=   "Value"
      Columns(2).Name =   "Value"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1032
      Columns(3).Caption=   "Unit"
      Columns(3).Name =   "Unit"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Style=   3
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "UnitID"
      Columns(4).Name =   "UnitID"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   7038
      _ExtentY        =   6032
      _StockProps     =   79
      Caption         =   "Custom Product Mesurements"
      BackColor       =   15724527
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpOrderDate 
      Height          =   315
      Left            =   1905
      TabIndex        =   1
      Top             =   1335
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
   Begin SITextBox.Txt TxtOrderID 
      Height          =   315
      Left            =   405
      TabIndex        =   0
      Top             =   1335
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
   Begin SITextBox.Txt TxtTotAmount 
      Height          =   315
      Left            =   300
      TabIndex        =   22
      Top             =   6840
      Width           =   1035
      _ExtentX        =   1826
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpDueDate 
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Top             =   1335
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
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   390
      TabIndex        =   3
      Top             =   2025
      Width           =   8640
      _ExtentX        =   15240
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
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   4575
      TabIndex        =   5
      Top             =   2985
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
      Left            =   5355
      TabIndex        =   6
      Top             =   2985
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
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   6315
      TabIndex        =   30
      Top             =   2985
      Width           =   1620
      _ExtentX        =   2858
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
   Begin SITextBox.Txt TxtAdvance 
      Height          =   315
      Left            =   1485
      TabIndex        =   9
      Top             =   6840
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
      Masked          =   1
   End
   Begin SITextBox.Txt TxtBalance 
      Height          =   315
      Left            =   2670
      TabIndex        =   35
      Top             =   6840
      Width           =   1035
      _ExtentX        =   1826
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      Height          =   195
      Left            =   2670
      TabIndex        =   36
      Top             =   6600
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Advance"
      Height          =   195
      Left            =   1485
      TabIndex        =   34
      Top             =   6600
      Width           =   645
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   6315
      TabIndex        =   33
      Top             =   2790
      Width           =   540
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   4575
      TabIndex        =   32
      Top             =   2790
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   5355
      TabIndex        =   31
      Top             =   2790
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   390
      TabIndex        =   29
      Top             =   1830
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   195
      Left            =   3720
      TabIndex        =   28
      Top             =   1140
      Width           =   690
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
      TabIndex        =   27
      Top             =   540
      Width           =   435
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   300
      TabIndex        =   23
      Top             =   6600
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID"
      Height          =   195
      Left            =   405
      TabIndex        =   21
      Top             =   1140
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
      Height          =   195
      Left            =   1905
      TabIndex        =   20
      Top             =   1140
      Width           =   780
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Order Booking"
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
      TabIndex        =   19
      Top             =   180
      Width           =   4005
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Product Code"
      Height          =   195
      Left            =   75
      TabIndex        =   18
      Top             =   2790
      Width           =   1545
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Product Name"
      Height          =   195
      Left            =   1875
      TabIndex        =   17
      Top             =   2790
      Width           =   1590
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
Attribute VB_Name = "FrmCustomOrderBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vCounter As Integer
Dim vIsNewRecord As Boolean
Dim RsBody As New ADODB.Recordset
Dim RsDetails As New ADODB.Recordset
Dim DetailFlag As Boolean
Dim Flag As Boolean
Dim vIsNewRow As Boolean
Dim vStrSQL As String
Dim vBm As Variant
Dim i As Integer
Dim vMaxBinID As Integer
    
'-----------------------
Private Sub BtnCustomProduct_Click()
   If FunSelectCustomProduct(ssButton, True) = True Then
      TxtQty.SetFocus
   Else
      TxtCustomProductCode.SetFocus
   End If
End Sub

Private Function FunSelectCustomProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchCustomProduct.Show vbModal, Me
      If SchCustomProduct.ParaOutCode = "" Then FunSelectCustomProduct = False: Exit Function
      TxtCustomProductCode.Text = SchCustomProduct.ParaOutCode
   End If
    '---------------------------
    vStrSQL = "select * from CustomProductsMeasurements where depth = 1 and " & vbCrLf _
            + " ID = '" & TxtCustomProductCode.Text & "'"
            
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtCustomProductName.Text = !Name
         FunSelectCustomProduct = True
         .Close
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
      Else
         FunSelectCustomProduct = False
         .Close
         TxtCustomProductCode.Text = ""
         TxtCustomProductName.Text = ""
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
'   GridDetail.Visible = True
'   GridDetail.Redraw = False
'   GridDetail.MoveFirst
'   GridDetail.RemoveAll
'   If FunSelectCustomProduct = True Then
'      GridDetail.AllowAddNew = True
'      vStrSQL = "SELECT p.ProductID, ProductName, Qty" & vbCrLf _
'         + " from InsCustomerBody b inner join Products p on b.ProductID = p.ProductID" & vbCrLf _
'         + " WHERE CustomProductCode = '" & TxtAreaCode.Text & " " & TxtCustomProductCode.Text & "'"
'
'      With CN.Execute(vStrSQL)
'         If .RecordCount > 0 Then
'            While Not .EOF
'               GridDetail.AddNew
'               GridDetail.Columns("Code").Text = !ProductID
'               GridDetail.Columns("Name").Text = !ProductName
'               GridDetail.Columns("Qty").Value = !Qty
'               .MoveNext
'            Wend
'         End If
'      End With
'      GridDetail.AllowAddNew = False
'   End If
'   GridDetail.Redraw = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnOpen_Click()
   SchCustomOrder.Show vbModal, Me
   If SchCustomOrder.ParaOutOrderID <> 0 Then
      TxtOrderID.Text = SchCustomOrder.ParaOutOrderID
      GetPrevious
   End If
End Sub

Private Sub GetPrevious()
   On Error GoTo ErrorHandler
   vStrSQL = " Select h.*, p.partyname, BankMachineName from CustomOrderHeader h left outer join parties p on h.customerid = p.partyid left outer join BankMachines b on b.BankMachineid = h.BankMachineid " & _
          " where OrderId=" & Val(TxtOrderID.Text)
   With CN.Execute(vStrSQL)
      If Not .BOF Then
         TxtOrderID.Text = !OrderID
         DtpOrderDate.DateValue = !OrderDate
         DtpDueDate.DateValue = !DueDate
         TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
         TxtTotAmount.Text = IIf(IsNull(!TotalAmount), "", !TotalAmount)
         TxtAdvance.Text = IIf(IsNull(!Advance), "", !Advance)
         FrmCustomPrint.OptBankCard.Value = !BankCard
         FrmCustomPrint.OptCash.Value = !Cash
         FrmCustomPrint.OptCredit.Value = !Credit
         If FrmCustomPrint.OptBankCard.Value = True Then
            FrmCustomPrint.TxtInvoiceNo.Text = !InvoiceNo
            FrmCustomPrint.TxtCommision.Text = !Commision
            FrmCustomPrint.TxtBankMachineID.Text = !BankMachineID
            FrmCustomPrint.TxtBankMachineName.Text = !BankMachineName
            FrmCustomPrint.TxtCustomerID.Text = ""
            FrmCustomPrint.TxtCustomerName.Text = ""
            FrmCustomPrint.TxtCashCustomer.Text = ""
            FrmCustomPrint.TxtBankCustomer.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
         End If
         If FrmCustomPrint.OptCash.Value = True Then
            FrmCustomPrint.TxtCommision.Text = ""
            FrmCustomPrint.TxtInvoiceNo.Text = ""
            FrmCustomPrint.TxtBankMachineID.Text = ""
            FrmCustomPrint.TxtBankMachineName.Text = ""
            FrmCustomPrint.TxtCustomerID.Text = ""
            FrmCustomPrint.TxtCustomerName.Text = ""
            FrmCustomPrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
            FrmCustomPrint.TxtBankCustomer.Text = ""
         End If
         If FrmCustomPrint.OptCredit.Value = True Then
            FrmCustomPrint.TxtCommision.Text = ""
            FrmCustomPrint.TxtInvoiceNo.Text = ""
            FrmCustomPrint.TxtBankMachineID.Text = ""
            FrmCustomPrint.TxtBankMachineName.Text = ""
            FrmCustomPrint.TxtCustomerID.Text = !CustomerID
            FrmCustomPrint.TxtCustomerName.Text = !PartyName
            FrmCustomPrint.TxtCashCustomer.Text = ""
            FrmCustomPrint.TxtBankCustomer.Text = ""
         End If
      End If
      .Close
   End With
   PopulateDataToGrid
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If ObjUserSecurity.IsAdministrator = False And TxtOrderID.Enabled = False Then
        MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
        Exit Sub
   End If
'  Header Validation
'   If Trim(TxtAreaCode.Text) = "" Then
'      MsgBox "Enter Area Code.", vbExclamation, Me.Caption
'      TxtAreaCode.SetFocus
'      Exit Sub
'   End If
   FrmCustomPrint.ParaInChoice = "Cash"
   FrmCustomPrint.Show vbModal, Me
   If FrmCustomPrint.ParaOutSelection = False Then Exit Sub

   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
      MsgBox "Please enter atleast one Customer Order", vbInformation + vbOKOnly, "Aletr"
      TxtCustomProductCode.SetFocus
      Exit Sub
   End If
   
   If TxtOrderID.Enabled Then
      If CN.Execute("Select * from CustomOrderHeader where OrderId = " & Val(TxtOrderID.Text)).RecordCount > 0 Then
         MsgBox "This Order ID already exists. A new Order ID. has been generated. Please try again", vbCritical, "Alert"
         TxtOrderID.Text = FunGetMaxID
         Exit Sub
      End If
   End If
  
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
   Dim Rs As New ADODB.Recordset
   CN.BeginTrans
   vStrSQL = "select * from CustomOrderHeader where OrderId=" & Val(TxtOrderID.Text)
   With Rs
      .Open vStrSQL, CN, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !OrderID = Val(TxtOrderID.Text)
      End If
         !OrderDate = DtpOrderDate.DateValue
         !DueDate = DtpDueDate.DateValue
         !Description = IIf(Trim(TxtDescription.Text) = "", Null, TxtDescription.Text)
         !TotalAmount = TxtTotAmount.Text
         !Advance = IIf(Trim(TxtAdvance.Text) = "", Null, TxtAdvance.Text)
         If FrmCustomPrint.OptBankCard.Value = True Then
            !InvoiceNo = FrmCustomPrint.TxtInvoiceNo.Text
            !Commision = FrmCustomPrint.TxtCommision.Text
            !BankMachineID = FrmCustomPrint.TxtBankMachineID.Text
            !CustomerID = "621"
            !CustomerName = IIf(Trim(FrmCustomPrint.TxtBankCustomer.Text) = "", Null, FrmCustomPrint.TxtBankCustomer.Text)
         End If
         If FrmCustomPrint.OptCash.Value = True Then
            !Commision = Null
            !InvoiceNo = Null
            !BankMachineID = Null
            !CustomerID = "621"
            !CustomerName = IIf(Trim(FrmCustomPrint.TxtCashCustomer.Text) = "", Null, FrmCustomPrint.TxtCashCustomer.Text)
         End If
         If FrmCustomPrint.OptCredit.Value = True Then
            !Commision = Null
            !InvoiceNo = Null
            !BankMachineID = Null
            !CustomerID = FrmCustomPrint.TxtCustomerID.Text
            !CustomerName = Null
         End If
         !BankCard = FrmCustomPrint.OptBankCard.Value
         !Cash = FrmCustomPrint.OptCash.Value
         !Credit = FrmCustomPrint.OptCredit.Value
         !UserNo = vUser
      .Update
      .Close
   End With
   
   RsBody.Filter = 0
   RsBody.MoveFirst
   For vCounter = 1 To RsBody.RecordCount
      RsBody!OrderID = Val(TxtOrderID.Text)
      RsBody.MoveNext
   Next vCounter
   RsBody.UpdateBatch
   
   RsDetails.Filter = 0
   RsDetails.MoveFirst
   For vCounter = 1 To RsDetails.RecordCount
      RsDetails!OrderID = Val(TxtOrderID.Text)
      RsDetails.MoveNext
   Next vCounter
   RsDetails.UpdateBatch
   
   CN.CommitTrans
   Grid.Redraw = True
   TxtOrderID.Enabled = False
   'MsgBox "Record has been saved", vbOKOnly + vbInformation, "Alert"
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub GridDetail_Change()
   DetailFlag = True
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
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
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Customer Order Booking"
   HelpLocation Me
   DtpOrderDate.DateValue = Date
   With CN.Execute("Select * FROM Units")
      GridDetail.Columns("Unit").AddItem ""
      For vCounter = 1 To .RecordCount
         GridDetail.Columns("Unit").AddItem !UnitName
         GridDetail.Columns("Unit").ItemData(vCounter) = !UnitID
         .MoveNext
      Next vCounter
   End With
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Property Get FormStatus() As FormMode
  'Nothing
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
   'Based upon the value of vNewValue, we shall deCustomProductCodee what controls to enable/disable
   On Error GoTo ErrorHandler
   vMode = vNewValue
   Select Case vNewValue
   Case Is = NewMode
      TxtCustomProductCode.Text = ""
      Call SubClearFields
      'If RsBody.State = adStateOpen Then RsBody.Close
      TxtOrderID.Text = FunGetMaxID
      'DtpDate.Value = Date
      TxtCustomProductCode.Enabled = True
      BtnCustomProduct.Enabled = True
      'GridDetail.Visible = False
      If DtpOrderDate.Enabled And DtpOrderDate.Visible Then DtpOrderDate.SetFocus
      'If TxtAreaCode.Enabled And TxtAreaCode.Visible Then TxtAreaCode.SetFocus
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtOrderID.Enabled = True
      PopulateDataToGrid
      vIsNewRecord = True
      vIsNewRow = True
   Case Is = OpenMode
      TxtOrderID.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      If TxtCustomProductCode.Enabled And TxtCustomProductCode.Visible Then TxtCustomProductCode.SetFocus
      vIsNewRecord = False
      vIsNewRow = True
   Case Is = ChangeMode
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      Select Case ActiveControl.Name
         Case TxtCustomProductCode.Name, TxtQty.Name, TxtPrice.Name
         If TxtCustomProductCode.Enabled Then TxtCustomProductCode.SetFocus: Call SubClearDetailArea
      End Select
   ElseIf Shift = vbCtrlMask Then
      If ActiveControl.Name = Grid.Name Then
         If KeyCode = vbKeyDelete Then
            If Trim(Grid.Columns("Code").Text <> "") Then Call mniRemoveRow_Click
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
'         Case vbKeyP
'            If BtnPrint.Enabled Then BtnPrint_Click
'            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name
         Grid_DblClick
'      Case TxtCode.Name
'         If FunSelectProduct(ssValidate, False) = True Then GetDataFromTexBoxesToGrid
'      Case TxtQty.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name, TxtPrice.Name
'         GetDataFromTexBoxesToGrid
      Case Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCustomProductCode.Name: If FunSelectCustomProduct(ssFunctionKey, False) = True Then TxtQty.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtCustomProductCode.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCustomProductCode.Name Then
         KeyCode = 0
         TxtAdvance.SetFocus
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

Private Sub PopulateGridDetail()
   On Error GoTo ErrorHandler
   Dim vSQL As String
   vSQL = "select p.*, unitname from CustomProductsMeasurements p left outer join units u on u.unitid = p.unitid where ParentID = '" & Grid.Columns("Code").Text & "'"
   GridDetail.Redraw = False
   GridDetail.MoveFirst
   GridDetail.RemoveAll
   With CN.Execute(vSQL)
      While Not .EOF
         GridDetail.AddNew
         GridDetail.Columns("ID").Text = !id
         GridDetail.Columns("Name").Text = !Name
         RsDetails.Filter = " CustomProductCode = '" & Grid.Columns("Code").Text & "' and ID = '" & !id & "'"
         If RsDetails.RecordCount > 0 Then
            GridDetail.Columns("Value").Text = RsDetails!Value
         Else
            GridDetail.Columns("Value").Text = ""
         End If
         If RsDetails.RecordCount > 0 Then
            If IsNull(RsDetails!UnitID) = False Then
               GridDetail.Columns("Unit").Text = CN.Execute("select UnitName from Units where unitid= " & RsDetails!UnitID).Fields(0).Value
            Else
               GridDetail.Columns("Unit").Text = ""
            End If
         Else
            GridDetail.Columns("Unit").Text = IIf(IsNull(!UnitName) = True, "", !UnitName)
         End If
         GridDetail.Update
         .MoveNext
      Wend
   End With
   If GridDetail.Rows > 0 Then GridDetail.FirstRow = 0
   GridDetail.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from CustomOrderBody where OrderId=" & Val(TxtOrderID.Text), CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      vStrSQL = "select b.*, Name from CustomOrderBody b join CustomProductsMeasurements p on p.ID = b.CustomProductCode where OrderId=" & Val(TxtOrderID.Text)
      With CN.Execute(vStrSQL)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotAmount.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("Code").Text = !CustomProductCode
            Grid.Columns("Name").Text = !Name
            Grid.Columns("Qty").Value = !Qty
            Grid.Columns("Price").Value = !Price
            Grid.Columns("Amount").Value = !Amount
            TxtTotAmount.Text = Val(TxtTotAmount.Text) + Val(!Amount)
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
   RsDetails.Filter = 0
   If RsDetails.State = adStateOpen Then RsDetails.Close
   RsDetails.Open "Select * from CustomOrderDetail where OrderId=" & Val(TxtOrderID.Text), CN, adOpenStatic, adLockBatchOptimistic
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   'GridDetail.Visible = False
   If Trim(TxtCustomProductCode.Text) = "" Then
      MsgBox "Enter Order Product Code.", vbExclamation, "Alert"
      TxtCustomProductCode.SetFocus
      Exit Sub
   End If
   If Trim(TxtQty.Text) = "" Then
      MsgBox "Enter Qty.", vbExclamation, "Alert"
      TxtQty.SetFocus
      Exit Sub
   End If
On Error GoTo ErrorHandler
   RsBody.Filter = "CustomProductCode='" & TxtCustomProductCode.Text & "'"
   If vIsNewRow Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("Code").Text = TxtCustomProductCode.Text
         RsBody!CustomProductCode = TxtCustomProductCode.Text
      Else
         MsgBox "This Product Already Exists.", vbOKOnly + vbInformation, "Alert"
         If TxtCustomProductCode.Visible And TxtCustomProductCode.Enabled Then TxtCustomProductCode.SetFocus
         Exit Sub
      End If
   End If
   
'   GridDetail.Redraw = False
'   vBm = GridDetail.Bookmark
'   GridDetail.MoveFirst
'   For i = 0 To GridDetail.Rows - 1
'      RsDetails.Filter = "CustomProductCode = '" & TxtCustomProductCode.Text & "' and ProductID = " & GridDetail.Columns("ID").CellValue(GridDetail.GetBookmark(i))
'      If RsDetails.RecordCount = 0 Then RsDetails.AddNew
'      RsDetails!CustomProductCode = TxtCustomProductCode.Text
'      RsDetails!ProductID = GridDetail.Columns("ID").CellValue(GridDetail.GetBookmark(i))
'      RsDetails!Qty = GridDetail.Columns("Qty").CellValue(GridDetail.GetBookmark(i))
'      RsDetails.Update
'   Next i
'   RsDetails.Filter = 0
'   GridDetail.Bookmark = vBm
'   GridDetail.Redraw = True
'
   Grid.Redraw = False
   With Grid
      If TxtCustomProductCode.Enabled = True Then
         TxtTotAmount.Text = Val(TxtTotAmount.Text) + Val(TxtAmount.Text)
      Else
         TxtTotAmount.Text = Val(TxtTotAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
      End If
      .Columns("Name").Text = TxtCustomProductName.Text
      .Columns("Qty").Text = Val(TxtQty.Text)
      .Columns("Price").Text = Val(TxtPrice.Text)
      .Columns("Amount").Text = Val(TxtAmount.Text)
      RsBody!Qty = Val(TxtQty.Text)
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
   TxtCustomProductCode.SetFocus
   vIsNewRow = True
   RsBody.Filter = 0
   Grid.Redraw = True
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub GridDetail_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   DispPromptMsg = 0
End Sub

Private Sub SubDetailUpdate()
   On Error GoTo ErrorHandler
   RsDetails.Filter = "CustomProductCode='" & IIf(Grid.Rows = 1, "", RsBody!CustomProductCode) & "' and ID='" & GridDetail.Columns("ID").Text & "'"
   If RsDetails.RecordCount = 0 And GridDetail.Columns("Value").Text <> "" Then
      RsDetails.AddNew
      RsDetails!CustomProductCode = RsBody!CustomProductCode
      RsDetails!id = GridDetail.Columns("ID").Text
      RsDetails!Value = GridDetail.Columns("Value").Text
      If GridDetail.Columns("Unit").Text = "" Then
         RsDetails!UnitID = Null
      Else
         RsDetails!UnitID = GridDetail.Columns("Unit").ItemData(GridDetail.Columns("Unit").ListIndex)
      End If
   ElseIf RsDetails.RecordCount = 1 And GridDetail.Columns("Value").Text = "" Then
      RsDetails.Delete
   ElseIf RsDetails.RecordCount = 1 Then
      RsDetails!Value = GridDetail.Columns("Value").Text
      If GridDetail.Columns("Unit").Text = "" Then
         RsDetails!UnitID = Null
      Else
         RsDetails!UnitID = GridDetail.Columns("Unit").ItemData(GridDetail.Columns("Unit").ListIndex)
      End If
      RsDetails.Update
   End If
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   DetailFlag = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridDetail_BeforeUpdate(Cancel As Integer)
   'If GridDetail.Visible = False Then Exit Sub
   If ActiveControl.Name <> GridDetail.Name Then Exit Sub
   On Error GoTo ErrorHandler
   SubDetailUpdate
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridDetail_GotFocus()
   RsBody.Filter = "CustomProductCode='" & TxtCustomProductCode.Text & "'"
   GridDetail.Row = 0
   GridDetail.Col = 0
   SendKeys "{Right}"
End Sub

Private Sub GridDetail_LostFocus()
   On Error GoTo ErrorHandler
   If DetailFlag = True Then SubDetailUpdate
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(GridDetail.Columns("ID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub TxtAdvance_Change()
   SubCalculateFooter
End Sub

Private Sub TxtPrice_LostFocus()
   Select Case ActiveControl.Name
   Case TxtCustomProductCode.Name, TxtPrice.Name, TxtQty.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = CN.Execute("Select isnull(max(OrderID),0) from CustomOrderHeader").Fields(0) + 1
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is txt Then
         If ctl.Tag = "" Then ctl.Text = ""
      End If
   Next
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("Code").Text = " "
   Grid.Update
   GridDetail.RemoveAll
   Unload FrmCustomPrint
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Are you sure to close without save?", vbQuestion + vbApplicationModal + vbYesNo, "Alert") = vbNo Then
         Cancel = 1
      End If
   Else
      If RsBody.State = adStateOpen Then RsBody.Close
      Set RsBody = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotAmount.Text = Val(TxtTotAmount.Text) - Grid.Columns("Amount").Value
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   Call Grid_LostFocus
End Sub

Private Sub Grid_GotFocus()
   Flag = True
   TxtCustomProductCode.Enabled = False
   BtnCustomProduct.Enabled = False
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("Code").Text) = "" Then
      TxtCustomProductCode.Text = ""
      TxtCustomProductCode.Enabled = True
      BtnCustomProduct.Enabled = True
      If TxtCustomProductCode.Visible And TxtCustomProductCode.Enabled Then TxtCustomProductCode.SetFocus
      vIsNewRow = True
   Else
      TxtCustomProductCode.Enabled = False
      BtnCustomProduct.Enabled = False
      GridDetail.SetFocus
      vIsNewRow = False
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(Grid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
'Flag = True
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
'   If ActiveControl.Name = GridDetail.Name Then
'      RsDetails.Filter = "CustomProductCode='" & Grid.Columns("Code").Text & "'"
'      If RsDetails.RecordCount > 0 Then RsDetails.Delete
'      'GridDetail.SelBookmarks.RemoveAll
'      'GridDetail.SelBookmarks.Add GridDetail.Bookmark
'      'GridDetail.DeleteSelected
'      'GridDetail.Refresh
'   End If
   If ActiveControl.Name = Grid.Name Then
      If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
      RsBody.Filter = "CustomProductCode='" & TxtCustomProductCode.Text & "'"
      If RsBody.RecordCount > 0 Then
         RsBody.Delete
         Grid.SelBookmarks.RemoveAll
         Grid.SelBookmarks.Add Grid.Bookmark
         Grid.DeleteSelected
         Grid.Refresh
         RsBody.Filter = 0
         RsDetails.Filter = "CustomProductCode='" & TxtCustomProductCode.Text & "'"
         If RsDetails.RecordCount > 0 Then RsDetails.Delete
      End If
   End If
   If vIsNewRow = False Then GetDataBackFromGridToTexBoxes
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtCustomProductCode.Text = .Columns("Code").Text
      TxtCustomProductName.Text = .Columns("Name").Text
      TxtQty.Text = .Columns("Qty").Value
      TxtPrice.Text = .Columns("Price").Value
      TxtAmount.Text = .Columns("Amount").Value
   End With
   PopulateGridDetail
   If Grid.Rows = 1 Then Grid.MoveLast
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
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   If ObjUserSecurity.IsAdministrator = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   CN.BeginTrans
   
   vMaxBinID = FunGetMaxBinID
   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
   CN.Execute ("Insert Into Bin_CustomOrderHeader Select " & vMaxBinID & ",'" & Date & "',* from CustomOrderHeader Where OrderID = " & TxtOrderID.Text & " And OrderDate ='" & DtpOrderDate.DateValue & "'")
    '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into Bin_CustomOrderBody Select " & vMaxBinID & ",'" & Date & "', * from CustomOrderBody Where OrderID = " & TxtOrderID.Text)
    '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into Bin_CustomOrderDetail Select " & vMaxBinID & ",'" & Date & "', * from CustomOrderDetail Where OrderID = " & TxtOrderID.Text)

   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Custom Order Booking'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'   Grid.Redraw = False
'   Grid.MoveFirst
'   For vCounter = 1 To Grid.Rows
'      If Trim(Grid.Columns("CustomProductCode").Text) <> "" Then
'         CN.Execute "Update inscustomerheader set TotalAmount=" & Grid.Columns("TotalAmount").Value & " where CustomProductCode='" & Grid.Columns("CustomProductCode").Text & "'"
'      End If
'      Grid.MoveNext
'   Next vCounter
'   Grid.RemoveAll
'   Grid.Redraw = True
   CN.Execute "Delete from CustomOrderDetail where OrderId = " & Val(TxtOrderID.Text)
   CN.Execute "Delete from CustomOrderBody where OrderId = " & Val(TxtOrderID.Text)
   CN.Execute "Delete from CustomOrderHeader where OrderId = " & Val(TxtOrderID.Text)
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtCustomProductCode.Enabled = True
   BtnCustomProduct.Enabled = True
   TxtCustomProductCode.Text = ""
   TxtCustomProductName.Text = ""
   TxtQty.Text = ""
   TxtPrice.Text = ""
   TxtAmount.Text = ""
   End Sub

Private Sub TxtCustomProductCode_Change()
   If TxtCustomProductCode.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomProductCode.Name Then Exit Sub
   If TxtCustomProductName.Text <> "" Then TxtCustomProductName.Text = ""
End Sub

Private Sub TxtCustomProductCode_LostFocus()
   TxtCustomProductCode.Text = StrConv(TxtCustomProductCode.Text, vbUpperCase)
End Sub

Private Sub TxtCustomProductCode_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If Me.ActiveControl.Name <> TxtCustomProductCode.Name Then Exit Sub
    If Trim(TxtCustomProductCode.Text) = "" Then Exit Sub
    Dim vTemp As Boolean
    vTemp = Not FunSelectCustomProduct(ssValidate, True)
    If vTemp = True Then
        vTemp = Not FunSelectCustomProduct(ssButton, False)
    End If
    Cancel = vTemp
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub TxtCustomProductCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtPrice_Change()
   Call SubCalculateBody
End Sub

Private Sub TxtQty_Change()
   Call SubCalculateBody
End Sub

Private Sub SubCalculateBody()
    TxtAmount.Text = Val(TxtQty.Text) * Val(TxtPrice.Text)
End Sub

Private Sub SubCalculateFooter()
    TxtBalance.Text = Val(TxtTotAmount.Text) - Val(TxtAdvance.Text)
End Sub

Private Sub TxtTotAmount_Change()
   SubCalculateFooter
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpOrderDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = CN.Execute("Select isnull(max(BinID),0)+1 from Bin_CustomOrderHeader ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
