VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmCustomOrderReturn 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmCustomOrderReturn.frx":0000
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
      Left            =   11970
      TabIndex        =   18
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
         TabIndex        =   19
         Tag             =   "NC"
         Text            =   "FrmCustomOrderReturn.frx":0ECA
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
         TabIndex        =   20
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9038
      TabIndex        =   11
      Top             =   8708
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
      MICON           =   "FrmCustomOrderReturn.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7711
      TabIndex        =   8
      Top             =   8708
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
      MICON           =   "FrmCustomOrderReturn.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5071
      TabIndex        =   10
      Top             =   8708
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
      MICON           =   "FrmCustomOrderReturn.frx":0FE2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10358
      TabIndex        =   12
      Top             =   8708
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
      MICON           =   "FrmCustomOrderReturn.frx":0FFE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6391
      TabIndex        =   9
      Top             =   8708
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
      MICON           =   "FrmCustomOrderReturn.frx":101A
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3420
      Left            =   1740
      TabIndex        =   5
      Top             =   3863
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
      stylesets(0).Picture=   "FrmCustomOrderReturn.frx":1036
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
      _ExtentY        =   6032
      _StockProps     =   79
      Caption         =   "Booking"
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
      Left            =   9630
      TabIndex        =   6
      Top             =   3863
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
      stylesets(0).Picture=   "FrmCustomOrderReturn.frx":1052
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
      stylesets(1).Picture=   "FrmCustomOrderReturn.frx":106E
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
      stylesets(2).Picture=   "FrmCustomOrderReturn.frx":108A
      AllowUpdate     =   0   'False
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
      Left            =   6690
      TabIndex        =   1
      Top             =   2198
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
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
      Left            =   5190
      TabIndex        =   0
      Top             =   2198
      Width           =   1140
      _ExtentX        =   2011
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
   End
   Begin SITextBox.Txt TxtPreviousBalance 
      Height          =   315
      Left            =   4560
      TabIndex        =   16
      Top             =   7793
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
      Left            =   8085
      TabIndex        =   2
      Top             =   2198
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
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
      Left            =   2430
      TabIndex        =   4
      Top             =   2888
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
   Begin SITextBox.Txt TxtPayment 
      Height          =   315
      Left            =   5865
      TabIndex        =   7
      Top             =   7793
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpReturnDate 
      Height          =   315
      Left            =   3705
      TabIndex        =   25
      Top             =   2198
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
   Begin SITextBox.Txt TxtReturnID 
      Height          =   315
      Left            =   2430
      TabIndex        =   26
      Top             =   2198
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
   Begin JeweledBut.JeweledButton BtnOrder 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6330
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2183
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
      MICON           =   "FrmCustomOrderReturn.frx":10A6
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotAmount 
      Height          =   315
      Left            =   1950
      TabIndex        =   30
      Top             =   7808
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
   Begin SITextBox.Txt TxtAdvance 
      Height          =   315
      Left            =   3255
      TabIndex        =   31
      Top             =   7808
      Width           =   1035
      _ExtentX        =   1826
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
      Masked          =   1
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3728
      TabIndex        =   34
      Top             =   8708
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
      MICON           =   "FrmCustomOrderReturn.frx":10C2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   9690
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   2198
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
      Left            =   10755
      TabIndex        =   35
      Tag             =   "NC"
      Top             =   2198
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
      Left            =   10395
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2198
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
      MICON           =   "FrmCustomOrderReturn.frx":10DE
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   10875
      TabIndex        =   38
      Top             =   2003
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   9690
      TabIndex        =   37
      Top             =   2003
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   1950
      TabIndex        =   33
      Top             =   7568
      Width           =   945
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Advance"
      Height          =   195
      Left            =   3255
      TabIndex        =   32
      Top             =   7568
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      Height          =   195
      Left            =   3705
      TabIndex        =   28
      Top             =   2003
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Retrun ID"
      Height          =   195
      Left            =   2430
      TabIndex        =   27
      Top             =   2003
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment"
      Height          =   195
      Left            =   5865
      TabIndex        =   24
      Top             =   7553
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   2430
      TabIndex        =   23
      Top             =   2693
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   195
      Left            =   8085
      TabIndex        =   22
      Top             =   2003
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
      TabIndex        =   21
      Top             =   540
      Width           =   435
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Balance"
      Height          =   195
      Left            =   4560
      TabIndex        =   17
      Top             =   7553
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID"
      Height          =   195
      Left            =   5190
      TabIndex        =   15
      Top             =   2003
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
      Height          =   195
      Left            =   6690
      TabIndex        =   14
      Top             =   2003
      Width           =   780
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Order Retrun"
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
      TabIndex        =   13
      Top             =   270
      Width           =   3720
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
Attribute VB_Name = "FrmCustomOrderReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vDate As Date, vHDiff As Integer, vSystemDate As Boolean
Dim vCounter As Integer
Dim vIsNewRecord As Boolean
Dim RsBody As New ADODB.Recordset
Dim RsDetails As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim DetailFlag As Boolean
Dim Flag As Boolean
Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String, vTotDisc As Double
Dim vIsNewRow As Boolean
Dim vStrSQL As String
Dim ssql As String
Dim vBm As Variant
Dim i As Integer
Dim vMaxBinID As Integer
    
'-----------------------
Private Sub BtnOpen_Click()
   SchCustomOrderReturn.Show vbModal, Me
   If SchCustomOrderReturn.ParaOutReturnID <> 0 Then
      TxtReturnID.Text = SchCustomOrderReturn.ParaOutReturnID
      cn.Execute ("Insert Into UserActivities values ('Custom Order Return'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetPrevious
   End If
End Sub

Private Function FunSelectBooking(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCustomOrder.ParaInShowOrder = False
        SchCustomOrder.Show vbModal, Me
        If SchCustomOrder.ParaOutOrderID = 0 Then FunSelectBooking = False: Exit Function
        TxtOrderID.Text = SchCustomOrder.ParaOutOrderID
    End If
    '---------------------------
    If Trim(TxtOrderID.Text) = "" Then Exit Function
'    ssql = "Select h.*, p.partyname, BankMachineName from CustomOrderHeader h left outer join parties p on h.customerid = p.partyid left outer join BankMachines b on b.BankMachineid = h.BankMachineid " & _
           "where ReturnID = " & Val(TxtReturnID.Text)
   
    ssql = " Select h.*, OrganizationName " & vbCrLf _
            + " from CustomOrderHeader h left outer join Organizations o on o.OrganizationID = h.OrganizationID" & vbCrLf _
            + " where OrderID = " & Val(TxtOrderID.Text)
    With cn.Execute(ssql)
      If .RecordCount > 0 Then
        DtpOrderDate.DateValue = !OrderDate
        DtpDueDate.DateValue = !DueDate
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
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
            FrmCustomPrint.TxtCustomerName.Text = !partyname
            FrmCustomPrint.TxtCashCustomer.Text = ""
            FrmCustomPrint.TxtBankCustomer.Text = ""
         End If
         PopulateDataToGrid
         FunSelectBooking = True
         .Close
         Exit Function
      Else
         FunSelectBooking = False
         .Close
         TxtOrderID.Text = ""
         DtpOrderDate.DateValue = ""
         DtpDueDate.DateValue = ""
         Exit Function
      End If
    End With
Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub GetPrevious()
   On Error GoTo ErrorHandler
   vStrSQL = " Select h.*, co.*, p.partyname, BankMachineName, OrganizationName  from CustomOrderReturn h inner join CustomOrderHeader co on co.OrderID = h.OrderID left outer join Organizations o on o.OrganizationID = h.OrganizationID left outer join parties p on h.customerid = p.partyid left outer join BankMachines b on b.BankMachineid = h.BankMachineid " & _
          " where ReturnID = " & Val(TxtReturnID.Text)
   With cn.Execute(vStrSQL)
      If Not .BOF Then
         TxtReturnID.Text = !ReturnID
         DtpReturnDate.DateValue = !ReturnDate
         TxtOrderID.Text = !OrderID
         DtpOrderDate.DateValue = !OrderDate
         DtpDueDate.DateValue = !DueDate
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
         TxtTotAmount.Text = IIf(IsNull(!TotalAmount), "", !TotalAmount)
         TxtAdvance.Text = IIf(IsNull(!Advance), "", !Advance)
         TxtPayment.Text = IIf(IsNull(!Payment), "", !Payment)
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
            FrmCustomPrint.TxtCustomerName.Text = !partyname
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

'Private Sub TxtOrderID_Change()
'   If TxtOrderID.Visible = False Then Exit Sub
'   If ActiveControl.Name <> TxtOrderID.Name Then Exit Sub
''   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
'End Sub

'Private Sub TxtOrderID_Validate(Cancel As Boolean)
'   If Me.ActiveControl.Name <> TxtOrderID.Name Then Exit Sub
'   On Error GoTo ErrorHandler
'   'If TxtOrderID.Text <> "" Then Exit Sub
'   Dim vTemp As Boolean
'   vTemp = Not FunSelectBooking(ssValidate, True)
'   If vTemp = True Then
'      vTemp = Not FunSelectBooking(ssButton, False)
'   End If
'   Cancel = vTemp
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub

Private Sub BtnOrder_Click()
   If FunSelectBooking(ssButton, False) = True Then
      TxtDescription.SetFocus
   Else
      If TxtOrderID.Enabled = True Then TxtOrderID.SetFocus
   End If
End Sub


Private Sub BtnPrint_Click()
'On Error GoTo ErrorHandler
'   vStrSQL = " select u.username, d.ReturnID, d.ReturnDate, h.OrderID, h.OrderDate, h.TotalAmount, h.Advance, d.payment, Name as productname, b.qty, b.price as price, b.amount, " & vbCrLf _
'      + " d.Cash, d.Credit, d.BankCard" & vbCrLf _
'      + " from CustomOrderHeader h inner join CustomOrderBody b on h.OrderID = b.OrderID " & vbCrLf _
'      + " inner Join CustomOrderReturn D on d.orderID = h.OrderID" & vbCrLf _
'      + " Inner join CustomProductsMeasurements CPM on CPM.id  = b.CustomProductCode" & vbCrLf _
'      + " inner join users u on u.UserNo = h.UserNo" & vbCrLf _
'      + " inner join Parties Pr on Pr.PartyID = h.CustomerID" & vbCrLf _
'      + "  where d.ReturnID = " & Val(TxtReturnID.Text) & " and d.ReturnDate ='" & DtpReturnDate.DateValue & "'" & vbCrLf _
'      + " Order By SerialNo"
'
'   If RsReport.State = adStateOpen Then RsReport.Close
'   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
'
'   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
'
'      If InStr(1, Printer.DeviceName, "CBM1000") > 0 Then
'         Set RptReportViewer.Report = New CrptCustomOrderReturnCBM
'      ElseIf InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
'         Set RptReportViewer.Report = New CrptCustomOrderReturnAurora
'         RptReportViewer.Report.LeftMargin = 225
'         RptReportViewer.Report.RightMargin = 0
'         RptReportViewer.Report.TopMargin = 255
'      Else 'InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
'         Set RptReportViewer.Report = New CrptCustomOrderReturnAurora
'         RptReportViewer.Report.LeftMargin = 0
'         RptReportViewer.Report.RightMargin = 0
'         RptReportViewer.Report.TopMargin = 0
'      End If
'
'
'    RptReportViewer.Report.DiscardSavedData
'    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
'
'    'RptReportViewer.Report.LeftMargin = 0
'    'RptReportViewer.Report.RightMargin = 0
'    vStrComp = "Select CompanyName,Address,City,PhoneNo,email from Company"
'    With CN.Execute(vStrComp)
'      If .RecordCount > 0 Then
'         vCompanyName = !CompanyName
'         vAddress = !Address
'         vAddress = !Address & IIf(!City = "", "", IIf(!Address = "", "", ", ") & !City)
'         vPhone = IIf(!PhoneNo = "", "", !PhoneNo)
'            RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
'            RptReportViewer.Report.ParameterFields(2).AddCurrentValue vAddress
'            RptReportViewer.Report.ParameterFields(4).AddCurrentValue vPhone
'      End If
'   End With
'   RptReportViewer.Report.ParameterFields(3).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
'   With CN.Execute("select * from registry")
'      If .RecordCount > 0 Then
'         RptReportViewer.Report.ParameterFields(5).AddCurrentValue "" 'IIf(!AddSpace = True, ".", "")
'         RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(False) 'CBool(!CashReceived)
'         RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(!OrderStatement)
'      End If
'      .Close
'   End With
    cn.Execute ("Insert Into UserActivities values ('Custom Order Return'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
'   'RptReportViewer.Show
'   RptReportViewer.Report.PrintOut False
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniCustomOrderReturn", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If

   'If ObjUserSecurity.IsAdministrator = False And TxtOrderID.Enabled = False Then
   '     MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
   '     Exit Sub
   'End If
   
   FrmCustomPrint.ParaInChoice = "Cash"
   FrmCustomPrint.Show vbModal, Me
   If FrmCustomPrint.ParaOutSelection = False Then Exit Sub

   If TxtReturnID.Enabled Then
      If cn.Execute("Select * from CustomOrderReturn where ReturnID = " & Val(TxtReturnID.Text)).RecordCount > 0 Then
         MsgBox "This Return ID already exists. A new Return ID. has been generated. Please try again", vbCritical, "Alert"
         TxtReturnID.Text = FunGetMaxID
         Exit Sub
      End If
   End If
   
    '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where touserno = " & vUser & " and Entrydate <='" & Date & "'"
    With cn.Execute(vStrSQL)
        If .Fields(0).Value >= DtpReturnDate.DateValue Then
            MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
    
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
   Dim Rs As New ADODB.Recordset
   cn.BeginTrans
   
   Call UserActivities
   
   vStrSQL = "select * from CustomOrderReturn where ReturnID=" & Val(TxtReturnID.Text)
   With Rs
      .Open vStrSQL, cn, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !ReturnID = Val(TxtReturnID.Text)
      End If
         !OrganizationID = IIf(Trim(TxtOrganizationID.Text) = "", Null, TxtOrganizationID.Text)
         !OrderID = Val(TxtOrderID.Text)
         !ReturnDate = DtpReturnDate.DateValue
         !Description = IIf(Trim(TxtDescription.Text) = "", Null, TxtDescription.Text)
         '!TotalAmount = TxtTotAmount.Text
         '!Advance = IIf(Trim(TxtAdvance.Text) = "", Null, TxtAdvance.Text)
         !Payment = IIf(Trim(TxtPayment.Text) = "", Null, TxtPayment.Text)
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
   
   cn.CommitTrans
   'Grid.Redraw = True
   'TxtOrderID.Enabled = False
   'MsgBox "Record has been saved", vbOKOnly + vbInformation, "Alert"
   'If FrmCustomPrint.ChkPrint.Value = 1 Then Call BtnPrint_Click
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub GridDetail_Change()
   DetailFlag = True
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
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
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Customer Order Return"
   HelpLocation Me
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   With cn.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
      End If
      .Close
   End With

   vSystemDate = Abs(ObjRegistry.SystemDate)
   vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
   
   With cn.Execute("Select * FROM Units")
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
      Call SubClearFields
      'If RsBody.State = adStateOpen Then RsBody.Close
      TxtReturnID.Text = FunGetMaxID
      'DtpDate.Value = Date
      'GridDetail.Visible = False
      vDate = IIf(Abs(vSystemDate) = 1, cn.Execute("Select SystemDate From SystemDate").Fields(0).Value, Date)
      DtpReturnDate.DateValue = IIf(Abs(vSystemDate) = 1, vDate, IIf(Format(Now, "hh") >= vHDiff, vDate, DateAdd("d", -1, vDate)))
      If DtpReturnDate.Enabled And DtpReturnDate.Visible Then DtpReturnDate.SetFocus
      'If TxtAreaCode.Enabled And TxtAreaCode.Visible Then TxtAreaCode.SetFocus
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      TxtReturnID.Enabled = True
      PopulateDataToGrid
      vIsNewRecord = True
   Case Is = OpenMode
      TxtReturnID.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      'If TxtCustomProductCode.Enabled And TxtCustomProductCode.Visible Then TxtCustomProductCode.SetFocus
      vIsNewRecord = False
   Case Is = ChangeMode
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
      BtnPrint.Enabled = False
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
   ElseIf KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
'      Case Grid.Name
'         Grid_DblClick
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
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtDescription.SetFocus Else TxtOrganizationID.SetFocus
      End Select
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
   vSQL = " select Name, p.ID, isnull(u.UnitName,u1.UnitName) as UnitName, isnull(Value,0) as Value" & vbCrLf _
      + " from CustomProductsMeasurements p " & vbCrLf _
      + " left outer Join (select * from CustomOrderDetail where OrderID = " & Val(TxtOrderID.Text) & ")d on p.ID = d.ID " & vbCrLf _
      + " left outer join units u on u.unitid = d.unitid " & vbCrLf _
      + " Left Outer join units u1 on u1.unitid = p.unitid " & vbCrLf _
      + " where ParentID = '" & Grid.Columns("Code").Text & "'"
   
   GridDetail.Redraw = False
   GridDetail.MoveFirst
   GridDetail.RemoveAll
   With cn.Execute(vSQL)
      While Not .EOF
         GridDetail.AddNew
         GridDetail.Columns("ID").Text = !ID
         GridDetail.Columns("Name").Text = !Name
         GridDetail.Columns("Value").Text = !Value
         GridDetail.Columns("Unit").Text = IIf(IsNull(!UnitName), "", !UnitName)
         GridDetail.Update
         .MoveNext
      Wend
   End With
   GridDetail.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   'RsBody.Filter = 0
   'If RsBody.State = adStateOpen Then RsBody.Close
   'RsBody.Open "Select * from CustomOrderBody where OrderId=" & Val(TxtOrderID.Text), CN, adOpenStatic, adLockBatchOptimistic
   Grid.Redraw = False
   Grid.MoveFirst
   Grid.RemoveAll
   vStrSQL = " select b.*, Name from CustomOrderBody b join CustomProductsMeasurements p on p.ID = b.CustomProductCode where OrderId=" & Val(TxtOrderID.Text)
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
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
      End If
      .Close
   End With
'   Grid.AddNew
'   Grid.Columns("Code").Text = " "
'   Grid.AllowAddNew = False
   Grid.Redraw = True
   'RsDetails.Filter = 0
   'If RsDetails.State = adStateOpen Then RsDetails.Close
   'RsDetails.Open "Select * from CustomOrderDetail where OrderId=" & Val(TxtOrderID.Text), CN, adOpenStatic, adLockBatchOptimistic
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtAdvance_Change()
   SubCalculateFooter
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select isnull(max(ReturnID),0) from CustomOrderReturn").Fields(0) + 1
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
'   Grid.AddNew
'   Grid.Columns("Code").Text = " "
'   Grid.Update
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
'   Else
'      If RsBody.State = adStateOpen Then RsBody.Close
'      Set RsBody = Nothing
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

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   'Flag = True
   PopulateGridDetail
   'If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   cn.Execute ("Insert Into UserActivities values ('Custom Order Return'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   cn.Execute ("Insert Into UserActivities values ('Custom Order Return'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
    ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniCustomOrderReturn", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   If ObjUserSecurity.IsAdministrator = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   cn.BeginTrans
   
'   vMaxBinID = FunGetMaxBinID
'   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_CustomOrderReturn Select " & vMaxBinID & ",'" & Date & "',* from CustomOrderReturn Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")

   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   cn.Execute ("Insert Into UserActivities values ('Custom Order Return'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   cn.Execute "Delete from CustomOrderReturn where ReturnID = " & Val(TxtReturnID.Text)
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateFooter()
   TxtPreviousBalance.Text = Val(TxtTotAmount.Text) - Val(TxtAdvance.Text)
End Sub

Private Sub TxtPayment_Change()
   SubCalculateFooter
End Sub

Private Sub TxtTotAmount_Change()
   SubCalculateFooter
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpOrderDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_CustomOrderReturn ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
    If vIsNewRecord = False Then
    Grid.MoveFirst
    For i = 1 To Grid.rows - 1
        With cn.Execute("Select * from CustomOrderBody Where DeliveryID = " & TxtReturnID.Text & " and CustomProductCode ='" & Grid.Columns("Code").Text & "'")
             If .EOF = True Then
                cn.Execute ("Insert Into UserActivities values ('Custom Order Return'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("Qty").Text <> !Qty Or Grid.Columns("Price").Text <> !Price Then
                   cn.Execute ("Insert Into UserActivities values ('Custom Order Return'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & !Qty & " Price-" & !Price & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
   Else
    cn.Execute ("Insert Into UserActivities values ('Custom Order Return'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
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
    vStrSQL = "Select * FROM Organizations where OrganizationID = " & Val(TxtOrganizationID.Text)
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
      TxtDescription.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

