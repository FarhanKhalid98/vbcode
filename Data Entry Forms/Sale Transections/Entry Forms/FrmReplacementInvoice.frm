VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmReplacementInvoice 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmReplacementInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame RFrame 
      Height          =   2175
      Left            =   90
      TabIndex        =   142
      Top             =   4500
      Visible         =   0   'False
      Width           =   2295
      Begin VB.TextBox TxtRSerial 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   20
         TabIndex        =   143
         Top             =   180
         Width           =   2025
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridRSerial 
         Height          =   1500
         Left            =   120
         TabIndex        =   144
         Top             =   555
         Width           =   2040
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   2
         stylesets.count =   1
         stylesets(0).Name=   "SelectedRow"
         stylesets(0).ForeColor=   -2147483634
         stylesets(0).BackColor=   -2147483635
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
         stylesets(0).Picture=   "FrmReplacementInvoice.frx":0ECA
         AllowDelete     =   -1  'True
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
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
         ForeColorEven   =   0
         BackColorOdd    =   15724527
         RowHeight       =   423
         ExtraHeight     =   26
         ActiveRowStyleSet=   "SelectedRow"
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "ProductID"
         Columns(0).Name =   "ProductID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3096
         Columns(1).Caption=   "Serial In"
         Columns(1).Name =   "Serial"
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   3598
         _ExtentY        =   2646
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
   End
   Begin VB.Frame PFrame 
      Caption         =   "New Serial"
      Height          =   1185
      Left            =   135
      TabIndex        =   141
      Top             =   1350
      Visible         =   0   'False
      Width           =   8100
      Begin VB.TextBox TxtPSerial 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5805
         MaxLength       =   20
         TabIndex        =   149
         Top             =   510
         Width           =   2025
      End
      Begin SITextBox.Txt TxtPCode 
         Height          =   315
         Left            =   90
         TabIndex        =   145
         Top             =   510
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         Appearance      =   0
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
         IntegralPoint   =   15
         Mandatory       =   1
      End
      Begin JeweledBut.JeweledButton BtnPProduct 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   1950
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   510
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
         MICON           =   "FrmReplacementInvoice.frx":0EE6
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtPProductName 
         Height          =   315
         Left            =   2310
         TabIndex        =   147
         Top             =   510
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Serial"
         Height          =   195
         Left            =   5805
         TabIndex        =   151
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   195
         Left            =   2310
         TabIndex        =   150
         Top             =   315
         Width           =   1020
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   195
         Left            =   90
         TabIndex        =   148
         Top             =   315
         Width           =   375
      End
   End
   Begin VB.Frame SFrame 
      Height          =   2175
      Left            =   90
      TabIndex        =   138
      Top             =   7605
      Visible         =   0   'False
      Width           =   2295
      Begin VB.TextBox TxtSSerial 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   20
         TabIndex        =   139
         Top             =   180
         Width           =   2025
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridSSerial 
         Height          =   1500
         Left            =   120
         TabIndex        =   140
         Top             =   555
         Width           =   2040
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   2
         stylesets.count =   1
         stylesets(0).Name=   "SelectedRow"
         stylesets(0).ForeColor=   -2147483634
         stylesets(0).BackColor=   -2147483635
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
         stylesets(0).Picture=   "FrmReplacementInvoice.frx":0F02
         AllowDelete     =   -1  'True
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
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
         ForeColorEven   =   0
         BackColorOdd    =   15724527
         RowHeight       =   423
         ExtraHeight     =   26
         ActiveRowStyleSet=   "SelectedRow"
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "ProductID"
         Columns(0).Name =   "ProductID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3096
         Columns(1).Caption=   "Serial Out"
         Columns(1).Name =   "Serial"
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   3598
         _ExtentY        =   2646
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
   End
   Begin VB.ComboBox CmbSColourName 
      Height          =   315
      Left            =   6930
      Style           =   2  'Dropdown List
      TabIndex        =   128
      Top             =   7170
      Width           =   1200
   End
   Begin VB.ComboBox cmbSSizeName 
      Height          =   315
      Left            =   8085
      Style           =   2  'Dropdown List
      TabIndex        =   127
      Top             =   7170
      Width           =   840
   End
   Begin VB.ComboBox CmbRColourName 
      Height          =   315
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   126
      Top             =   4125
      Width           =   1200
   End
   Begin VB.ComboBox cmbRSizeName 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   125
      Top             =   4125
      Width           =   840
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1365
      MaxLength       =   50
      TabIndex        =   94
      Top             =   795
      Visible         =   0   'False
      Width           =   4125
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
      Height          =   4650
      Left            =   13860
      TabIndex        =   74
      Top             =   1035
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
         Height          =   4200
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   75
         Tag             =   "NC"
         Text            =   "FrmReplacementInvoice.frx":0F1E
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
         TabIndex        =   76
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.CheckBox ChkIsProduct 
      Caption         =   "Is Product"
      Height          =   255
      Left            =   10890
      TabIndex        =   48
      Top             =   1395
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1365
   End
   Begin SITextBox.Txt TxtReturnID 
      Height          =   315
      Left            =   1215
      TabIndex        =   33
      Top             =   1545
      Visible         =   0   'False
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
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10140
      TabIndex        =   27
      Top             =   10215
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
      MICON           =   "FrmReplacementInvoice.frx":1057
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8820
      TabIndex        =   23
      Top             =   10215
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
      MICON           =   "FrmReplacementInvoice.frx":1073
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6180
      TabIndex        =   25
      Top             =   10215
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
      MICON           =   "FrmReplacementInvoice.frx":108F
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11460
      TabIndex        =   28
      Top             =   10215
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
      MICON           =   "FrmReplacementInvoice.frx":10AB
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7500
      TabIndex        =   24
      Top             =   10215
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
      MICON           =   "FrmReplacementInvoice.frx":10C7
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtRBillDisc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5190
      TabIndex        =   9
      Top             =   6360
      Width           =   705
      _ExtentX        =   1244
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4860
      TabIndex        =   26
      Top             =   10215
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
      MICON           =   "FrmReplacementInvoice.frx":10E3
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtActualAmount 
      Height          =   315
      Left            =   8850
      TabIndex        =   37
      Top             =   750
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   1455
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
      Left            =   6150
      TabIndex        =   116
      Tag             =   "NC"
      Top             =   2805
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
      Left            =   7185
      TabIndex        =   39
      Tag             =   "NC"
      Top             =   2805
      Width           =   1080
      _ExtentX        =   1905
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
      Left            =   6825
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2805
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
      MICON           =   "FrmReplacementInvoice.frx":10FF
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpReturnDate 
      Height          =   315
      Left            =   2130
      TabIndex        =   32
      Top             =   1545
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtPID 
      Height          =   315
      Left            =   9960
      TabIndex        =   43
      Top             =   750
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Left            =   10845
      TabIndex        =   49
      Top             =   750
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
   Begin SITextBox.Txt TxtRBillID 
      Height          =   315
      Left            =   3345
      TabIndex        =   114
      Top             =   2805
      Width           =   510
      _ExtentX        =   900
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpRBillDate 
      Height          =   315
      Left            =   3855
      TabIndex        =   115
      Top             =   2805
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
   Begin SITextBox.Txt TxtRBillDiscPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6855
      TabIndex        =   10
      Top             =   6360
      Width           =   615
      _ExtentX        =   1085
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
      Height          =   465
      Left            =   5520
      TabIndex        =   31
      Top             =   2700
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   820
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
      MICON           =   "FrmReplacementInvoice.frx":111B
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSale 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5160
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2790
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
      MICON           =   "FrmReplacementInvoice.frx":1137
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtRNetAmount 
      Height          =   315
      Left            =   11265
      TabIndex        =   54
      Top             =   6360
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
   Begin SITextBox.Txt TxtReplaceID 
      Height          =   315
      Left            =   1185
      TabIndex        =   112
      Top             =   2820
      Width           =   825
      _ExtentX        =   1455
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpReplaceDate 
      Height          =   315
      Left            =   2010
      TabIndex        =   113
      Top             =   2820
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
   Begin SITextBox.Txt TxtTotReturnQty 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9465
      TabIndex        =   60
      Top             =   6360
      Width           =   615
      _ExtentX        =   1085
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
      Masked          =   1
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   3795
      TabIndex        =   62
      Top             =   1515
      Visible         =   0   'False
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
      Left            =   4710
      TabIndex        =   63
      Top             =   1515
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtSBillDisc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5055
      TabIndex        =   20
      Top             =   9720
      Width           =   705
      _ExtentX        =   1244
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
   Begin SITextBox.Txt TxtSBillDiscPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6810
      TabIndex        =   21
      Top             =   9720
      Width           =   615
      _ExtentX        =   1085
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
   Begin SITextBox.Txt TxtSNetAmount 
      Height          =   315
      Left            =   11265
      TabIndex        =   66
      Top             =   9720
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
   Begin SITextBox.Txt TxtTotSaleQty 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9540
      TabIndex        =   29
      Top             =   9720
      Width           =   615
      _ExtentX        =   1085
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
      Masked          =   1
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtMemberName 
      Height          =   315
      Left            =   7980
      TabIndex        =   79
      Top             =   1515
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
   Begin JeweledBut.JeweledButton BtnMember 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7620
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   1515
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
      MICON           =   "FrmReplacementInvoice.frx":1153
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   11235
      TabIndex        =   118
      Top             =   2805
      Width           =   615
      _ExtentX        =   1085
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
      Left            =   12210
      TabIndex        =   83
      Top             =   2805
      Width           =   1260
      _ExtentX        =   2223
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
      Left            =   11850
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   2805
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
      MICON           =   "FrmReplacementInvoice.frx":116F
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCommission 
      Height          =   315
      Left            =   11670
      TabIndex        =   87
      Top             =   780
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
   Begin SITextBox.Txt TxtManualBillNo 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2355
      TabIndex        =   8
      Top             =   6360
      Width           =   1515
      _ExtentX        =   2672
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
      Masked          =   1
      IntegralPoint   =   4
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   8295
      TabIndex        =   117
      Tag             =   "NC"
      Top             =   2805
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
      Left            =   9360
      TabIndex        =   90
      Tag             =   "NC"
      Top             =   2805
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
      Left            =   9000
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   2805
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
      MICON           =   "FrmReplacementInvoice.frx":118B
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmpComm 
      Height          =   315
      Left            =   420
      TabIndex        =   95
      Top             =   780
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
   Begin SITextBox.Txt TxtRCode 
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   4125
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Appearance      =   0
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
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtRQty 
      Height          =   315
      Left            =   8880
      TabIndex        =   1
      Top             =   4125
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
   Begin SITextBox.Txt TxtRPrice 
      Height          =   315
      Left            =   9660
      TabIndex        =   2
      Top             =   4125
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
   Begin SITextBox.Txt TxtRAmount 
      Height          =   315
      Left            =   13245
      TabIndex        =   98
      Top             =   4125
      Width           =   1545
      _ExtentX        =   2725
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
   Begin JeweledBut.JeweledButton BtnRProduct 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   3000
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   4125
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
      MICON           =   "FrmReplacementInvoice.frx":11A7
      BC              =   12632256
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridReturn 
      CausesValidation=   0   'False
      Height          =   1920
      Left            =   1140
      TabIndex        =   7
      Top             =   4440
      Width           =   13740
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   21
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "FrmReplacementInvoice.frx":11C3
      AllowUpdate     =   0   'False
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
      ExtraHeight     =   873
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   21
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
      Columns(4).Width=   2117
      Columns(4).Caption=   "Colour"
      Columns(4).Name =   "ColourName"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1455
      Columns(5).Caption=   "Size"
      Columns(5).Name =   "SizeName"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1376
      Columns(6).Caption=   "Qty"
      Columns(6).Name =   "Qty"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   4
      Columns(6).FieldLen=   256
      Columns(7).Width=   1693
      Columns(7).Caption=   "Price"
      Columns(7).Name =   "Price"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   1217
      Columns(8).Caption=   "Disc/Pc"
      Columns(8).Name =   "DiscPC"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   979
      Columns(9).Caption=   "Disc%"
      Columns(9).Name =   "DiscPer"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1217
      Columns(10).Caption=   "Dis. Val"
      Columns(10).Name=   "DiscVal"
      Columns(10).Alignment=   1
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   4
      Columns(10).FieldLen=   256
      Columns(11).Width=   1217
      Columns(11).Caption=   "SC"
      Columns(11).Name=   "SC"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   2328
      Columns(12).Caption=   "Amount"
      Columns(12).Name=   "Amount"
      Columns(12).Alignment=   1
      Columns(12).CaptionAlignment=   2
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   5
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "TotalAmount"
      Columns(13).Name=   "TotalAmount"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   3200
      Columns(14).Visible=   0   'False
      Columns(14).Caption=   "Cost"
      Columns(14).Name=   "Cost"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   4
      Columns(14).FieldLen=   256
      Columns(15).Width=   3200
      Columns(15).Visible=   0   'False
      Columns(15).Caption=   "IsProduct"
      Columns(15).Name=   "IsProduct"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   11
      Columns(15).FieldLen=   256
      Columns(15).Style=   2
      Columns(16).Width=   3200
      Columns(16).Visible=   0   'False
      Columns(16).Caption=   "ColourID"
      Columns(16).Name=   "ColourID"
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      Columns(17).Width=   3200
      Columns(17).Visible=   0   'False
      Columns(17).Caption=   "SizeID"
      Columns(17).Name=   "SizeID"
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      Columns(18).Width=   3200
      Columns(18).Caption=   "StoreID"
      Columns(18).Name=   "StoreID"
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(19).Width=   3200
      Columns(19).Visible=   0   'False
      Columns(19).Caption=   "IsSerial"
      Columns(19).Name=   "IsSerial"
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(20).Width=   3200
      Columns(20).Visible=   0   'False
      Columns(20).Caption=   "EmpComm"
      Columns(20).Name=   "EmpComm"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   24236
      _ExtentY        =   3387
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
   Begin SITextBox.Txt TxtRDiscVal 
      Height          =   315
      Left            =   11865
      TabIndex        =   5
      Top             =   4125
      Width           =   690
      _ExtentX        =   1217
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
   Begin SITextBox.Txt TxtRDiscPC 
      Height          =   315
      Left            =   10620
      TabIndex        =   3
      Top             =   4125
      Width           =   690
      _ExtentX        =   1217
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
   Begin SITextBox.Txt TxtRDiscPer 
      Height          =   315
      Left            =   11310
      TabIndex        =   4
      Top             =   4125
      Width           =   555
      _ExtentX        =   979
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
   Begin SITextBox.Txt TxtRSC 
      Height          =   315
      Left            =   12555
      TabIndex        =   6
      Top             =   4125
      Width           =   690
      _ExtentX        =   1217
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
   Begin SITextBox.Txt TxtSCode 
      Height          =   315
      Left            =   1185
      TabIndex        =   12
      Top             =   7170
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Appearance      =   0
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
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtSQty 
      Height          =   315
      Left            =   8925
      TabIndex        =   13
      Top             =   7170
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
      Left            =   9705
      TabIndex        =   14
      Top             =   7170
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
      Left            =   13290
      TabIndex        =   109
      Top             =   7170
      Width           =   1515
      _ExtentX        =   2672
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
      Left            =   3045
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   7170
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
      MICON           =   "FrmReplacementInvoice.frx":11DF
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtSProductName 
      Height          =   315
      Left            =   3405
      TabIndex        =   111
      Top             =   7170
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
      Height          =   2235
      Left            =   1185
      TabIndex        =   19
      Top             =   7485
      Width           =   13635
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   21
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "FrmReplacementInvoice.frx":11FB
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
      ExtraHeight     =   344
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   21
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
      Columns(4).Width=   2117
      Columns(4).Caption=   "Colour"
      Columns(4).Name =   "ColourName"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1455
      Columns(5).Caption=   "Size"
      Columns(5).Name =   "SizeName"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1376
      Columns(6).Caption=   "Qty"
      Columns(6).Name =   "Qty"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   4
      Columns(6).FieldLen=   256
      Columns(7).Width=   1693
      Columns(7).Caption=   "Price"
      Columns(7).Name =   "Price"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   1217
      Columns(8).Caption=   "Disc/Pc"
      Columns(8).Name =   "DiscPC"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   979
      Columns(9).Caption=   "Disc%"
      Columns(9).Name =   "DiscPer"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1217
      Columns(10).Caption=   "Dis. Val"
      Columns(10).Name=   "DiscVal"
      Columns(10).Alignment=   1
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   4
      Columns(10).FieldLen=   256
      Columns(11).Width=   1217
      Columns(11).Caption=   "SC"
      Columns(11).Name=   "Sc"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   4
      Columns(11).FieldLen=   256
      Columns(12).Width=   2302
      Columns(12).Caption=   "Amount"
      Columns(12).Name=   "Amount"
      Columns(12).Alignment=   1
      Columns(12).CaptionAlignment=   2
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   5
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "TotalAmount"
      Columns(13).Name=   "TotalAmount"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   3200
      Columns(14).Visible=   0   'False
      Columns(14).Caption=   "Cost"
      Columns(14).Name=   "Cost"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   4
      Columns(14).FieldLen=   256
      Columns(15).Width=   3200
      Columns(15).Visible=   0   'False
      Columns(15).Caption=   "QtyOrigional"
      Columns(15).Name=   "QtyOrigional"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   4
      Columns(15).FieldLen=   256
      Columns(16).Width=   3200
      Columns(16).Visible=   0   'False
      Columns(16).Caption=   "IsProduct"
      Columns(16).Name=   "IsProduct"
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   11
      Columns(16).FieldLen=   256
      Columns(16).Style=   2
      Columns(17).Width=   3200
      Columns(17).Visible=   0   'False
      Columns(17).Caption=   "EmpComm"
      Columns(17).Name=   "EmpComm"
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      Columns(18).Width=   3200
      Columns(18).Visible=   0   'False
      Columns(18).Caption=   "ColourID"
      Columns(18).Name=   "ColourID"
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(19).Width=   3200
      Columns(19).Visible=   0   'False
      Columns(19).Caption=   "SizeID"
      Columns(19).Name=   "SizeID"
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(20).Width=   3200
      Columns(20).Caption=   "StoreID"
      Columns(20).Name=   "StoreID"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   24051
      _ExtentY        =   3942
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
   Begin SITextBox.Txt TxtSDiscVal 
      Height          =   315
      Left            =   11910
      TabIndex        =   17
      Top             =   7170
      Width           =   690
      _ExtentX        =   1217
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
   Begin SITextBox.Txt TxtSDiscPC 
      Height          =   315
      Left            =   10665
      TabIndex        =   15
      Top             =   7170
      Width           =   690
      _ExtentX        =   1217
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
   Begin SITextBox.Txt TxtSDiscPer 
      Height          =   315
      Left            =   11340
      TabIndex        =   16
      Top             =   7170
      Width           =   555
      _ExtentX        =   979
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
   Begin SITextBox.Txt TxtSSC 
      Height          =   315
      Left            =   12600
      TabIndex        =   18
      Top             =   7170
      Width           =   690
      _ExtentX        =   1217
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
   Begin SITextBox.Txt TxtRServiceCharges 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7980
      TabIndex        =   11
      Top             =   6360
      Width           =   705
      _ExtentX        =   1244
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
   Begin SITextBox.Txt TxtSServiceCharges 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8070
      TabIndex        =   22
      Top             =   9720
      Width           =   705
      _ExtentX        =   1244
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
   Begin SITextBox.Txt TxtRProductName 
      Height          =   315
      Left            =   3360
      TabIndex        =   124
      Top             =   4125
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
   Begin SITextBox.Txt TxtMemberID 
      Height          =   315
      Left            =   6180
      TabIndex        =   130
      Top             =   1515
      Width           =   1440
      _ExtentX        =   2540
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
      IntegralPoint   =   10
   End
   Begin SITextBox.Txt TxtMemberBarCode 
      Height          =   315
      Left            =   9360
      TabIndex        =   131
      Top             =   1515
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
   Begin SITextBox.Txt TxtRSID 
      Height          =   315
      Left            =   300
      TabIndex        =   132
      Top             =   2805
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
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
   Begin SITextBox.Txt TxtSSID 
      Height          =   315
      Left            =   360
      TabIndex        =   134
      Top             =   7155
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
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
   Begin SITextBox.Txt TxtSID 
      Height          =   315
      Left            =   480
      TabIndex        =   136
      Top             =   1515
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "SID"
      Height          =   195
      Left            =   480
      TabIndex        =   137
      Top             =   1320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "SID Out"
      Height          =   195
      Left            =   360
      TabIndex        =   135
      Top             =   6960
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "SID In"
      Height          =   195
      Left            =   300
      TabIndex        =   133
      Top             =   2610
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label LblMemberBarCode 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member BarCode"
      Height          =   195
      Left            =   9405
      TabIndex        =   129
      Top             =   1275
      Width           =   1230
   End
   Begin VB.Label LblRColour 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
      Height          =   195
      Left            =   6885
      TabIndex        =   123
      Top             =   3930
      Width           =   450
   End
   Begin VB.Label LblRSize 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   8085
      TabIndex        =   122
      Top             =   3930
      Width           =   300
   End
   Begin VB.Label LblRCost 
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
      Left            =   1905
      TabIndex        =   121
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "S.C."
      Height          =   195
      Left            =   7665
      TabIndex        =   120
      Top             =   9780
      Width           =   300
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "S.C."
      Height          =   195
      Left            =   7575
      TabIndex        =   119
      Top             =   6405
      Width           =   300
   End
   Begin VB.Label LblRAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   13245
      TabIndex        =   108
      Top             =   3930
      Width           =   540
   End
   Begin VB.Label LblRQty 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   8880
      TabIndex        =   107
      Top             =   3930
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   1140
      TabIndex        =   106
      Top             =   3930
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3360
      TabIndex        =   105
      Top             =   3930
      Width           =   1020
   End
   Begin VB.Label LblRProdPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   9660
      TabIndex        =   104
      Top             =   3930
      Width           =   360
   End
   Begin VB.Label LblRSC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "S.C."
      Height          =   195
      Left            =   12585
      TabIndex        =   103
      Top             =   3930
      Width           =   300
   End
   Begin VB.Label LblRDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. %"
      Height          =   195
      Left            =   11295
      TabIndex        =   102
      Top             =   3930
      Width           =   525
   End
   Begin VB.Label LblRDiscPC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc / PC"
      Height          =   195
      Left            =   10560
      TabIndex        =   101
      Top             =   3930
      Width           =   690
   End
   Begin VB.Label LblRDiscVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. Val"
      Height          =   195
      Left            =   11865
      TabIndex        =   100
      Top             =   3930
      Width           =   630
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
      Height          =   225
      Index           =   1
      Left            =   1380
      TabIndex        =   97
      Top             =   555
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "EmpComm"
      Height          =   195
      Left            =   420
      TabIndex        =   96
      Top             =   555
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   9480
      TabIndex        =   93
      Top             =   2580
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   8295
      TabIndex        =   92
      Top             =   2580
      Width           =   1095
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill No"
      Height          =   195
      Left            =   1275
      TabIndex        =   89
      Top             =   6405
      Width           =   1020
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      Height          =   195
      Left            =   11535
      TabIndex        =   88
      Top             =   555
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   12240
      TabIndex        =   86
      Top             =   2580
      Width           =   780
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   11265
      TabIndex        =   85
      Top             =   2580
      Width           =   525
   End
   Begin VB.Label LblMemberID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      Height          =   195
      Left            =   6180
      TabIndex        =   82
      Top             =   1275
      Width           =   780
   End
   Begin VB.Label LblMemberName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      Height          =   195
      Left            =   8025
      TabIndex        =   81
      Top             =   1275
      Width           =   1035
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
      Height          =   225
      Index           =   0
      Left            =   8310
      TabIndex        =   78
      Top             =   690
      Visible         =   0   'False
      Width           =   900
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
      Left            =   12300
      TabIndex        =   77
      Top             =   1950
      Width           =   435
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "In"
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
      Left            =   1215
      TabIndex        =   73
      Top             =   3465
      Width           =   330
   End
   Begin VB.Label LblAmount 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Received"
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
      Height          =   285
      Left            =   1950
      TabIndex        =   72
      Top             =   9735
      Width           =   1740
   End
   Begin VB.Label TxtTotalAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   1635
      TabIndex        =   71
      Top             =   10005
      Width           =   2370
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Disc."
      Height          =   195
      Left            =   4380
      TabIndex        =   70
      Top             =   9780
      Width           =   600
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Disc. (%)"
      Height          =   195
      Left            =   5865
      TabIndex        =   69
      Top             =   9780
      Width           =   855
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Amount"
      Height          =   195
      Left            =   10230
      TabIndex        =   68
      Top             =   9780
      Width           =   900
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
      Height          =   195
      Left            =   8790
      TabIndex        =   67
      Top             =   9780
      Width           =   645
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   4710
      TabIndex        =   65
      Top             =   1320
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " Bill ID"
      Height          =   195
      Left            =   3795
      TabIndex        =   64
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
      Height          =   195
      Left            =   8745
      TabIndex        =   61
      Top             =   6405
      Width           =   645
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
      Left            =   2250
      TabIndex        =   59
      Top             =   6900
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Replace ID"
      Height          =   195
      Left            =   1185
      TabIndex        =   58
      Top             =   2625
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace Date"
      Height          =   195
      Left            =   2010
      TabIndex        =   57
      Top             =   2625
      Width           =   990
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Out"
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
      Left            =   1230
      TabIndex        =   56
      Top             =   6765
      Width           =   615
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Return Amount"
      Height          =   195
      Left            =   10140
      TabIndex        =   55
      Top             =   6405
      Width           =   1065
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Disc. (%)"
      Height          =   195
      Left            =   5955
      TabIndex        =   53
      Top             =   6405
      Width           =   855
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " Bill ID"
      Height          =   195
      Left            =   3345
      TabIndex        =   52
      Top             =   2610
      Width           =   450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   3855
      TabIndex        =   51
      Top             =   2610
      Width           =   585
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   10845
      TabIndex        =   50
      Top             =   525
      Visible         =   0   'False
      Width           =   315
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
      Left            =   10605
      TabIndex        =   47
      Top             =   2175
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
      Left            =   10590
      TabIndex        =   46
      Top             =   1905
      Width           =   720
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
      Left            =   2700
      TabIndex        =   45
      Top             =   270
      Width           =   3570
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   9960
      TabIndex        =   44
      Top             =   555
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   6150
      TabIndex        =   42
      Top             =   2610
      Width           =   585
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   7185
      TabIndex        =   41
      Top             =   2610
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Amount"
      Height          =   195
      Left            =   8835
      TabIndex        =   38
      Top             =   540
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   12660
      Top             =   1380
      Width           =   345
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Disc."
      Height          =   195
      Left            =   4515
      TabIndex        =   36
      Top             =   6405
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      Height          =   195
      Left            =   2130
      TabIndex        =   35
      Top             =   1350
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Return ID"
      Height          =   195
      Left            =   1215
      TabIndex        =   34
      Top             =   1350
      Visible         =   0   'False
      Width           =   690
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
Attribute VB_Name = "FrmReplacementInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim Application1 As New CRAXDRT.Application
Dim vDate, vServerDate As Date, vHDiff As Integer, vSystemDate As Boolean
Dim vCounter, vGridRows As Integer
Dim RsReturnBody As New ADODB.Recordset
Dim RsSaleBody As New ADODB.Recordset
Dim RsBodySerial As New ADODB.Recordset
Dim RsPurchaseSerial As New ADODB.Recordset
Dim RsReturnSerial As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vIsNewRecord As Boolean
Dim Flag As Boolean
Dim DateFlag As Boolean
Dim ssql As String
Dim i As Integer
Dim vStrSQL As String, vX As Integer, vY As Integer
Dim vQtyLoose As Double, vTotalAmount As Double
Dim vTotDisc As Double, vTotal As Double, vNoofPrints As Byte
Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String, vTotReturnDisc As Double, vTotalReturn As Double
Dim vMobileNo() As String, vMobile As String
Dim vStrDetailin, vStrDetailOut, vRandomID  As String
Dim vStrPara, vConnStr, vProducts As String
Dim vColour, vSerialAdd, vIsSerial, vIsNewSerial, vShowStock As Boolean
Dim vMasterID As Long
Public objFSO As New Scripting.FileSystemObject
Dim Cnn As New ADODB.Connection
Dim vPOSID As String, vFBRInvoiceNo As String, vUSIN As Long
Dim vStrDetail, vSamePid, vSQL As String
'----------------------------------

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

Private Sub BtnPProduct_Click()
If FunSelectPProduct(ssButton, True) = True Then
      TxtPSerial.SetFocus
   Else
      TxtPCode.SetFocus
   End If
End Sub

Private Sub TxtOrganizationID_Change()
   On Error GoTo ErrorHandler
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizationName.Text <> "" Then TxtOrganizationName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
'      If TxtCustomerID.Enabled Then TxtCustomerID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub SubCalculateRBody()
   TxtRDiscVal.Text = Val(TxtRQty.Text) * Val(TxtRDiscPC.Text)
   TxtActualAmount.Text = Val(TxtRQty.Text) * (Val(TxtRPrice.Text) + Val(TxtRSC.Text))
   TxtRAmount.Text = Val(TxtActualAmount.Text) - Val(TxtRDiscVal.Text)
   SubCalculateRFooter
End Sub

Private Sub SubPaidReceived()
   TxtTotalAmount.Caption = Abs(Val(TxtSNetAmount.Text) - Val(TxtRNetAmount.Text))
   LblAmount.Caption = IIf((Val(TxtSNetAmount.Text) - Val(TxtRNetAmount.Text)) >= 0, "Cash Received", "Cash Paid")
End Sub

Private Sub SubCalculateRFooter()
   TxtRNetAmount.Text = SelfRound(Val(vTotalReturn) - (vTotReturnDisc) - Val(TxtRBillDisc.Text) + Val(TxtRServiceCharges.Text))
End Sub

Private Sub SubDestroyMember()
   On Error GoTo ErrorHandler
   GridSale.MoveFirst
   ssql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With CN.Execute(ssql)
      While Trim(GridSale.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = " & Val(GridSale.Columns("ProductID").Text)
         If .RecordCount > 0 Then
            'GetDataBackFromGridSaleToTexBoxes
            
            RsSaleBody.Filter = "ProductID = " & Val(!Productid)
            GridSale.Columns("DiscPer").Value = 0 'IIf(IsNull(!DiscPer), 0, !DiscPer)
            GridSale.Columns("DiscPC").Value = 0 'Round((Val(RsSaleBody!Price) * Val(GridSale.Columns("DiscPer").Value) / 100), 2)
            GridSale.Columns("DiscVal").Value = 0 'Val(GridSale.Columns("DiscPC").Value) * Val(GridSale.Columns("Qty").Value)
            GridSale.Columns("Amount").Value = (Val(GridSale.Columns("Price").Value) * Val(GridSale.Columns("Qty").Value)) - Val(GridSale.Columns("DiscVal").Value)
            
            'TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - RsSaleBody!Amount + Val(GridSale.Columns("Amount").Text)
            vTotDisc = vTotDisc - RsSaleBody!DiscVal + Val(GridSale.Columns("DiscVal").Text)
            vTotalAmount = vTotalAmount - RsSaleBody!Amount + Val(GridSale.Columns("Amount").Text)
            
            RsSaleBody!DiscPC = Val(GridSale.Columns("DiscPC").Value)
            RsSaleBody!DiscPer = Val(GridSale.Columns("DiscPer").Value)
            RsSaleBody!DiscVal = Val(GridSale.Columns("DiscVal").Value)
            RsSaleBody!Amount = Val(GridSale.Columns("Amount").Value)
         End If
         GridSale.MoveNext
      Wend
      .Close
   End With
   SubCalculateSFooter
   'SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubApplyMember()
   On Error GoTo ErrorHandler
   GridSale.MoveFirst
   ssql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With CN.Execute(ssql)
      While Trim(GridSale.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = " & Val(GridSale.Columns("ProductID").Text)
         If .RecordCount > 0 Then
            'GetDataBackFromGridSaleToTexBoxes
            RsSaleBody.Filter = "ProductID = " & Val(!Productid)
            GridSale.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
            GridSale.Columns("DiscPC").Value = Round((Val(RsSaleBody!Price) * Val(GridSale.Columns("DiscPer").Value) / 100), 2)
            GridSale.Columns("DiscVal").Value = Val(GridSale.Columns("DiscPC").Value) * Val(GridSale.Columns("Qty").Value)
            GridSale.Columns("Amount").Value = (Val(GridSale.Columns("Price").Value) * Val(GridSale.Columns("Qty").Value)) - Val(GridSale.Columns("DiscVal").Value)
            
            'TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - RsSaleBody!Amount + Val(GridSale.Columns("Amount").Text)
            vTotDisc = vTotDisc - RsSaleBody!DiscVal + Val(GridSale.Columns("DiscVal").Text)
            vTotalAmount = vTotalAmount - RsSaleBody!Amount + Val(GridSale.Columns("Amount").Text)
            
            RsSaleBody!DiscPC = Val(GridSale.Columns("DiscPC").Value)
            RsSaleBody!DiscPer = Val(GridSale.Columns("DiscPer").Value)
            RsSaleBody!DiscVal = Val(GridSale.Columns("DiscVal").Value)
            RsSaleBody!Amount = Val(GridSale.Columns("Amount").Value)
     
         End If
         GridSale.MoveNext
      Wend
      .Close
   End With
   GridSale.MoveLast
   SubCalculateSFooter
   'SubCalculateFooter
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
            + " where isLockEmployee = 0 and EmpID=" & Val(TxtEmployeeID.Text)
    With CN.Execute(ssql)
      If .RecordCount > 0 Then
        TxtEmployeeName.Text = !empname
        TxtCommission.Text = !Commission
        FunSelectEmployee = True
        .Close
        Exit Function
      Else
        FunSelectEmployee = False
        .Close
        TxtEmployeeID.Text = ""
        TxtEmployeeName.Text = ""
        TxtCommission.Text = ""
        Exit Function
      End If
    End With
Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub BtnEmployee_Click()
   On Error GoTo ErrorHandler
   If FunSelectEmployee(ssButton, False) = True Then
'      If TxtMemberID.Visible = True Then TxtMemberID.SetFocus Else If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      TxtEmployeeID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtMemberID_Change()
   If ActiveControl.Name <> TxtMemberID.Name Then Exit Sub
   If TxtMemberName.Text <> "" Then TxtMemberName.Text = "": TxtMemberBarCode.Text = "": Call SubDestroyMember
End Sub

Private Sub TxtMemberID_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If TxtMemberName.Text <> "" Then Exit Sub
    If TxtMemberID.Text = "" Then Exit Sub
    Dim vTemp As Boolean
    vTemp = Not FunSelectMember(ssValidate, True)
    If vTemp = True Then
        vTemp = Not FunSelectMember(ssButton, False)
    End If
    Cancel = vTemp
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub BtnMember_Click()
   On Error GoTo ErrorHandler
   If FunSelectMember(ssButton, False) = True Then
      If TxtRCode.Enabled Then TxtRCode.SetFocus
   Else
      TxtMemberID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectMember(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchMember.Show vbModal, Me
        If SchMember.ParaOutMemberID = "" Then FunSelectMember = False: Exit Function
        TxtMemberID.Text = SchMember.ParaOutMemberID
    End If
    '---------------------------
    If Trim(TxtMemberID.Text) = "" Then Exit Function
    ssql = "Select * " & vbCrLf _
            + " from Members" & vbCrLf _
            + " where IsLockMember = 0 and ( MemberID = case when isnumeric('" & Trim(TxtMemberID.Text) & " ')=1 then '" & Trim(TxtMemberID.Text) & " ' else '' end or BarCode = '" & Trim(TxtMemberID.Text) & "')"
    With CN.Execute(ssql)
      If .RecordCount > 0 Then
        TxtMemberID.Text = !MemberID
        TxtMemberName.Text = !MemberName
        TxtMemberBarCode.Text = IIf(IsNull(!BarCode), "", !BarCode)
        Call SubApplyMember
        FunSelectMember = True
        .Close
        Exit Function
      Else
        FunSelectMember = False
        .Close
        TxtMemberID.Text = ""
        TxtMemberName.Text = ""
        Exit Function
      End If
    End With
Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Function FunSelectSale(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSale.ParaInBillDate = DtpRBillDate.DateValue
        SchSale.Show vbModal, Me
        If SchSale.ParaOutBillID = "" Then FunSelectSale = False: Exit Function
        TxtSSID.Text = SchSale.ParaOutSID
        TxtRBillID.Text = SchSale.ParaOutBillID
        DtpRBillDate.DateValue = SchSale.ParaOutBillDate
    End If
    '---------------------------
    vStrSQL = " Select * FROM SaleHeader where SID = " & Val(TxtSSID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          DtpRBillDate.DateValue = !BillDate
          FunSelectSale = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSale = False
          .Close
          TxtRBillID.Text = ""
          DtpRBillDate.DateValue = Date
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

Private Sub FindRRebate()
   Dim Rebate
   On Error GoTo ErrorHandler
    With CN.Execute("Select * from ProductOffers where Rebate <> 0 and ProductID = " & Val(TxtPID.Text))
        If .RecordCount > 0 Then
            Rebate = Val(TxtRQty.Text)
            Rebate = Rebate \ !Qty
            Rebate = Rebate * !Rebate
            TxtRDiscVal.Text = Rebate
            If Val(TxtRPrice.Text) = 0 Then Exit Sub
            If Val(TxtRQty.Text) = 0 Then Exit Sub
            TxtRDiscPC.Text = Round(Val(TxtRDiscVal.Text) / (TxtRQty.Text), 3)
            TxtRDiscPer.Text = Round((Val(TxtRDiscPC.Text) * 100) / Val(TxtRPrice.Text), 2)
            TxtActualAmount.Text = Val(TxtRQty.Text) * Val(TxtRPrice.Text)
            TxtRAmount.Text = Val(TxtActualAmount.Text) - Val(TxtRDiscVal.Text)
'            vTotReturnDisc = vTotReturnDisc + Val(TxtRDiscVal.Text)
'            vTotalReturn = vTotalReturn + TxtRAmount.Text
            SubCalculateRFooter
        End If
    End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectRProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      If vColour = True Then
         SchItemCode.ParaInWhere = " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ") "
         SchItemCode.Show vbModal, Me
         TxtRCode.Text = SchItemCode.ParaOutItemCode
      Else
         SchProduct.ParaInWhere = " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ") "
         SchProduct.ParainShowStock = vShowStock
         SchProduct.Show vbModal, Me
         TxtRCode.Text = SchProduct.ParaOutID
      End If
   End If
    '---------------------------
   If TxtRCode.Enabled = False Then FunSelectRProduct = False: Exit Function
   If Trim(TxtRCode.Text) = "" Then FunSelectRProduct = False: Exit Function
   If TxtRCode.Text = "" Then FunSelectRProduct = False: Exit Function
   
   If vColour = True Then
      ssql = "select c.ColourID, ColourName from productcolours pc inner join Colours c on pc.colourid = c.colourid " & vbCrLf _
             & "inner join products p on p.productid = pc.productid " & vbCrLf _
             & "where ItemCode = '" & IIf(Len(TxtRCode.Text) = 9, TxtRCode.Text & "'", Mid(TxtRCode.Text, 1, 9) & "' and c.colourid = " & Val(Mid(TxtRCode.Text, 10, 2)))
      With CN.Execute(ssql)
         If .RecordCount > 0 Then
            CmbRColourName.AddItem !ColourName
            CmbRColourName.ItemData(CmbRColourName.NewIndex) = !ColourID
            CmbRColourName.ListIndex = 0
         End If
      End With
      
      ssql = "select s.SizeID, SizeName from productSizes pz inner join Sizes s on pz.Sizeid = s.Sizeid " & vbCrLf _
      & "inner join products p on p.productid = pz.productid " & vbCrLf _
      & "where ItemCode = '" & IIf(Len(TxtRCode.Text) = 13, Mid(TxtRCode.Text, 1, 9) & "' and s.sizeid = " & Val(Mid(TxtRCode.Text, 12, 2)), TxtRCode.Text & "'")
      With CN.Execute(ssql)
         If .RecordCount > 0 Then
            cmbRSizeName.AddItem !SizeName
            cmbRSizeName.ItemData(cmbRSizeName.NewIndex) = !SizeID
            cmbRSizeName.ListIndex = 0
         End If
      End With
      TxtRCode.Text = CStr(Left(TxtRCode.Text, 9))
   End If
   
   ''''''''''''' Serail '''''''''''''''''''''''''''''''''
   vSerialAdd = False
   vStrSQL = "Select ProductID, Serial, SerialAdd from vuPurchaseSerial where Serial = '" & Trim(TxtRCode.Text) & "' or ProductID = " & Val(TxtRCode.Text)
   With CN.Execute(vStrSQL)
      If .EOF = False Then
            If RFrame.Visible = False Then
               RFrame.Visible = True
               RFrame.ZOrder 0
            End If
            TxtRSerial.Text = IIf(Trim(TxtPSerial.Text) = "", Trim(TxtRCode.Text), Trim(TxtPSerial.Text))
            TxtRCode.Text = !Productid
            GetDataFromTexBoxesToGridRSerial
            If vSerialAdd = False Then
               TxtRCode.Text = ""
               FunSelectRProduct = False
               Exit Function
            End If
      End If
   End With
 '''''''''''''''''''''''''''''''''''''''''''''

    ''''''''***********   Checking Union   ***********''''''''
    vStrSQL = " SELECT p.productid, Code, ProductName, IsSerial, ServiceCharges, RetailPrice, DiscPer, DiscPC" & vbCrLf _
         + " from PackageDealInfoHeader un inner join Products p on un.PackageDealID = p.productid" & vbCrLf _
         + " left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
         + " where ( " & IIf(IsNumeric(TxtRCode.Text) = False, "", "p.productid = " & (TxtRCode.Text) & " or ") & " code = '" & TxtRCode.Text & "')" & " and isLocked = 0 "
         

   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtRProductName.Text = !ProductName
         vIsSerial = !isSerial
         TxtRPrice.Text = !RetailPrice
         TxtRQty.Text = IIf(Val(TxtRQty.Text) = 0, 1, TxtRQty.Text)
         vStrSQL = " select sum(isnull(Cost,PurPrice)* b.QtyLoose) as Cost from PackageDealInfoHeader h inner join PackageDealInfoBody b on h.id = b.id" & vbCrLf _
               + " inner join Products p on p.productid = b.productid" & vbCrLf _
               + " left outer join CurrentStock cs on cs.productid = p.productid " & vbCrLf _
               + " where h.PackageDealID = '" & TxtPID.Text & "'"
         With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
'        VStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = '" & TxtPID.Text & "'"
'         With CN.Execute(VStrSQL)
'            If .RecordCount > 0 Then
'               vQtyLoose = .Fields(0).Value
'            Else
'               vQtyLoose = 0
'            End If
'         End With
'         LblStock.Caption = vQtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'         LblStock.Visible = True
'         LblStockCaption.Visible = True
'         LblStock.Visible = True
'         LblStockCaption.Visible = True
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpReplaceDate.DateValue + 1 & "',0),0)"
         vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = vQtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit(" & Val(TxtPID.Text) & ")").Fields(0).Value
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
        
'         VStrSQL = " select Floor(min(css.QtyLoose/b.QtyLoose)) as QtyLoose " & vbCrLf _
'                  + " from PackageDealInfoHeader h inner join PackageDealInfoBody b on h.id = b.id" & vbCrLf _
'                  + " inner join Products p on p.productid = b.productid" & vbCrLf _
'                  + " left outer join CurrentStockStore css on css.productid = p.productid " & vbCrLf _
'                  + " where h.PackageDealID ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text
'         With CN.Execute(VStrSQL)
'            If .RecordCount > 0 Then
'               vQtyLoose = !QtyLoose
'               LblStock.Caption = IIf(IsNull(!QtyLoose), 0, !QtyLoose) & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'            Else
'               vQtyLoose = 0
'               LblStock.Caption = 0
'            End If
'         End With
         
'         If ObjRegistry.NegativeSale = False Then
'            If LblStock.Caption <= 0 Then
'               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'               FunSelectRProduct = False
'               Exit Function
'            End If
'         End If
         TxtRSC.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtRDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtRDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtRDiscPC.Text) <> 0 Then
            TxtRDiscPer.Text = Round((Val(TxtRDiscPC.Text) * 100) / Val(TxtRPrice.Text), 2)
         End If
'         ChkIsProduct.Value = 0
         SubCalculateRBody
'         Char.Speak TxtRProductName.Text
         FunSelectRProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      End If
   End With

   ''''''''***********   Checking Product  ***********''''''''
    vStrSQL = " SELECT p.productid, Qty, code, ProductName, IsSerial, ServiceCharges, RetailPrice, DiscPer, DiscPC, EmpComm" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtRCode.Text) = False, "", "p.productid = " & (TxtRCode.Text) & " or ") & " code = '" & TxtRCode.Text & "')" & " and isLocked = 0 "

   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtRProductName.Text = !ProductName
         vIsSerial = !isSerial
         TxtRPrice.Text = !RetailPrice
         TxtRQty.Text = IIf(Len(TxtRCode.Text) <= 5 And IsNumeric(TxtRCode.Text), 1, IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty))  'IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text) 'IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty)
         With CN.Execute("select cost from currentstock where productid = " & Val(TxtPID.Text))
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
         TxtEmpComm.Text = IIf(IsNull(!EmpComm), 0, !EmpComm)
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpReplaceDate.DateValue + 1 & "',0),0)"
         vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = vQtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit(" & Val(TxtPID.Text) & ")").Fields(0).Value
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         
'         With CN.Execute("select QtyLoose from currentstockStore where productid ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text)
'            If .RecordCount > 0 Then
'               vQtyLoose = !QtyLoose
'               LblStock.Caption = !QtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'            Else
'               vQtyLoose = 0
'               LblStock.Caption = 0
'            End If
'         End With
'         LblStock.Visible = True
'         LblStockCaption.Visible = True

'         If ObjRegistry.NegativeSale = False Then
'            If Val(LblStock.Caption) <= 0 Then
'               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'               FunSelectRProduct = False
'               Exit Function
'            End If
'         End If
         
         TxtRSC.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtRDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtRDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtRDiscPC.Text) <> 0 Then
            TxtRDiscPer.Text = Round((Val(TxtRDiscPC.Text) * 100) / Val(TxtRPrice.Text), 2)
         End If
         ChkIsProduct.Value = 1
         If Val(TxtRQty.Text) > 1 Then FindRRebate
         SubCalculateRBody
'         Char.Speak TxtRProductName.Text
         FunSelectRProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectRProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtPID.Text = ""
         TxtRCode.Text = ""
         TxtRProductName.Text = ""
         TxtRPrice.Text = ""
         TxtRSC.Text = ""
         TxtRDiscPC.Text = ""
         TxtRDiscPer.Text = ""
         TxtRAmount.Text = ""
         TxtCost.Text = ""
         TxtEmpComm.Text = ""
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

Private Function FunSelectPProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
    '---------------------------
   SchProduct.Show vbModal, Me
   TxtPCode.Text = SchProduct.ParaOutID
   If TxtPCode.Enabled = False Then FunSelectPProduct = False: Exit Function
   If Trim(TxtPCode.Text) = "" Then FunSelectPProduct = False: Exit Function
   If TxtPCode.Text = "" Then FunSelectPProduct = False: Exit Function
   TxtRCode.Text = TxtPCode.Text
''''''''''''' Serail '''''''''''''''''''''''''''''''''
   If TxtPCode.Text = "" Then FunSelectPProduct = False: Exit Function
     
   
    vStrSQL = " SELECT p.productid, code, ProductName" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where isSerial = 1 and p.productid = " & Val(TxtPCode.Text) & " or code = '" & TxtPCode.Text & "'"
  
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPCode.Text = !Productid
         TxtPProductName.Text = !ProductName
         FunSelectPProduct = True
         vIsNewSerial = True
         .Close
         Exit Function
      Else
         FunSelectPProduct = False
         vIsNewSerial = False
         .Close
         MsgBox "Invalid Serial Product ID.", vbOKOnly, "Alert"
         TxtPCode.Text = ""
         TxtPProductName.Text = ""
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
'   cn.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
'    Call DeleteTempActivityLogBin(vRandomID)
      ''''''''''''''''''''''''''' Cleared Replacement In Products
      vGridRows = 0
      GridReturn.Redraw = False
      GridReturn.MoveFirst
      For vCounter = 2 To GridReturn.rows
         vGridRows = vGridRows + 1
         If Trim(GridReturn.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From saleReturnbody where SID=" & Val(TxtRSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(GridReturn.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmReplacementInvoice, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Cleared Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         GridReturn.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmReplacementInvoice, eClearSavedRecord, TxtReplaceID.Text, DtpReplaceDate.DateValue, vGridRows & " Replace In Product/s Cleared ")
      GridReturn.Redraw = True
      ''''''''''''''''''''''''''''''''''''''''
      
      ''''''''''''''''''''''''''' Cleared Replacement Out Products
      vGridRows = 0
      GridSale.Redraw = False
      GridSale.MoveFirst
      For vCounter = 2 To GridSale.rows
         vGridRows = vGridRows + 1
         If Trim(GridSale.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From salebody where SID=" & Val(TxtSSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(GridSale.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmReplacementInvoice, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Cleared Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         GridSale.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmReplacementInvoice, eClearSavedRecord, TxtReplaceID.Text, DtpReplaceDate.DateValue, vGridRows & " Replace Out Product/s Cleared ")
      GridSale.Redraw = True
      ''''''''''''''''''''''''''''''''''''''''
      
      
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
     '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
    ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniReplacementInvoice", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
   CN.BeginTrans
   
   Call BinData
   Call ActivityLogBin("", eFrmReplacementInvoice, eDelete, TxtReplaceID.Text, DtpReplaceDate.DateValue, GridReturn.rows - 1 & " Relplace In Product/s Deleted Amount: " & Val(TxtRNetAmount.Text))
   Call ActivityLogBin("", eFrmReplacementInvoice, eDelete, TxtReplaceID.Text, DtpReplaceDate.DateValue, GridSale.rows - 1 & " Relplace Out Product/s Deleted Amount: " & Val(TxtSNetAmount.Text))
   Call ActivityLogBin("", eFrmReplacementInvoice, eDelete, TxtReplaceID.Text, DtpReplaceDate.DateValue, " Relplacement Invoice Amount: " & (Val(TxtSNetAmount.Text) - Val(TxtRNetAmount.Text)))
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  ''''''''''''''''''''''''''Delete Serials'''''''''''''''''''''
    If RsReturnSerial.RecordCount > 0 Then
        RsReturnSerial.MoveFirst
        For vCounter = 1 To RsReturnSerial.RecordCount
            CN.Execute "Delete from SaleReturnSerial where ReturnID = " & Val(TxtRSID.Text) & " And ReturnDate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(RsReturnSerial!Productid) & " and Serial ='" & RsReturnSerial!Serial & "'"
            RsReturnSerial.MoveNext
        Next vCounter
    End If
    
   vStrDetailin = " In:-"
   ' Return body
   GridReturn.Redraw = False
   GridReturn.MoveFirst
'   Call ActivityLog("Replacement Invoice", eDelete, TxtReplaceID.Text, DtpReplaceDate.DateValue)
   For vCounter = 1 To GridReturn.rows
      If Trim(GridReturn.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from SaleReturnBody where SID = " & Val(TxtRSID.Text) & "And ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "' and productid = " & Val(GridReturn.Columns("Productid").Text) & " and StoreID = " & Val(TxtStoreID.Text)
         vStrDetailin = vStrDetailin & " (P" & GridReturn.Columns("Productid").Text & " Q" & GridReturn.Columns("Qty").Value & " A" & GridReturn.Columns("Amount").Value & ")"
      End If
      GridReturn.MoveNext
   Next vCounter
   GridReturn.RemoveAll
   GridReturn.Redraw = True
   
   ''''''''''''''''''''''''''Delete Serials'''''''''''''''''''''
    If RsBodySerial.RecordCount > 0 Then
        RsBodySerial.MoveFirst
        For vCounter = 1 To RsBodySerial.RecordCount
            CN.Execute "Delete from SaleBodySerial where BillID = " & Val(TxtSSID.Text) & " And BillDate ='" & DtpBillDate.DateValue & "' and productid = " & Val(RsBodySerial!Productid) & " and Serial ='" & RsBodySerial!Serial & "'"
            RsBodySerial.MoveNext
        Next vCounter
    End If
    
   vStrDetailOut = " Out:-"
   ' Sale body
   GridSale.Redraw = False
   GridSale.MoveFirst
'   Call ActivityLog("Replacement Invoice", eDelete, TxtReplaceID.Text, DtpReplaceDate.DateValue)
   For vCounter = 1 To GridSale.rows
      If Trim(GridSale.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from SaleBody where SID = " & Val(TxtSSID.Text) & " And BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "' and productid = " & Val(GridSale.Columns("Productid").Text) & " and StoreID = " & Val(TxtStoreID.Text)
         CN.Execute "Exec UpdateStockPlus " & TxtStoreID.Text & "," & Val(GridSale.Columns("ProductID").Text) & "," & GridSale.Columns("Qty").Value & "," & Val(TxtReplaceID.Text) & ",'" & DtpReplaceDate.DateValue & "'"
         vStrDetailOut = vStrDetailOut & " (P" & GridSale.Columns("Productid").Text & " Q" & GridSale.Columns("Qty").Value & " A" & GridSale.Columns("Amount").Value & ")"
      End If
      GridSale.MoveNext
   Next vCounter
   GridSale.RemoveAll
   GridSale.Redraw = True
   
   
   CN.Execute "Delete from SaleReturnHeader where SID = " & Val(TxtRSID.Text)
   CN.Execute "Delete from SaleHeader where SID = " & Val(TxtSSID.Text)
   CN.Execute "Delete from ReplacementHeader where SID = " & Val(TxtSID.Text)
   
   '''-------------- Mobile SMS -------------------
    If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = "Deleted Replacement ID:" & TxtReplaceID.Text & vbCrLf & " Date:" & Format(DtpReplaceDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtSBillDisc.Text) = 0, "", " SDisc:" & TxtSBillDisc.Text) & IIf(Val(TxtRBillDisc.Text) = 0, "", " RDisc:" & TxtRBillDisc.Text) & vbCrLf & " NetAmt:" & TxtTotalAmount.Caption
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetailin & vStrDetailOut, "") & "','')"
               CN.Execute ssql
            End If
         Next
   End If
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   GridReturn.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchReplacement.ParaInReplaceDate = DtpReplaceDate.DateValue
   SchReplacement.Show vbModal
   If SchReplacement.ParaOutReplaceID <> -1 Then
      TxtSID.Text = SchReplacement.ParaOutSID
      TxtSSID.Text = SchReplacement.ParaOutSSID
      TxtRSID.Text = SchReplacement.ParaOutRSID
      TxtReplaceID.Text = SchReplacement.ParaOutReplaceID
      DtpReplaceDate.DateValue = SchReplacement.ParaOutReplaceDate
      TxtReturnID.Text = SchReplacement.ParaOutReturnID
      DtpReturnDate.DateValue = SchReplacement.ParaOutReturnDate
      TxtBillID.Text = SchReplacement.ParaOutBillID
      DtpBillDate.DateValue = SchReplacement.ParaOutBillDate
'      cn.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetSaleReturn
      GetSale
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   
'   vStrSQL = " select username, ReplaceID, ReplaceDate, 'Sale' as Type, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, " & vbCrLf _
      + " isnull(r.cashReceived,0) as cashReceived, p.productname, unitname, b.qty, b.price as price, b.amount, b.DiscVal" & vbCrLf _
      + " , Case when CustomerID = '621' then isnull(CustomerName,PartyName) Else PartyName End as Customer, Cash, Credit, BankCard, 1 as Rec, SaleAmount, ReturnAmount" & vbCrLf _
      + " from ReplacementHeader r " & vbCrLf _
      + " inner join SaleHeader h on r.billid = h.billid and r.billdate = h.billdate" & vbCrLf _
      + " inner join salebody b on b.billid = h.billid and b.billdate = h.billdate " & vbCrLf _
      + " inner join Products p on p.Productid = b.ProductID" & vbCrLf _
      + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
      + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
      + " left outer join Parties Pr on Pr.PartyID = h.CustomerID" & vbCrLf _
      + " where ReplaceID= " & Val(TxtReplaceID.Text) & " and ReplaceDate='" & DtpReplaceDate.DateValue & "'" & vbCrLf _
      + " union all" & vbCrLf _
      + " select username, ReplaceID, ReplaceDate, 'Return', h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, " & vbCrLf _
      + " isnull(r.cashReceived,0) as cashReceived, p.productname, unitname, b.qty, b.price as price, b.amount, b.DiscVal" & vbCrLf _
      + " , Case when CustomerID = '621' then isnull(CustomerName,PartyName) Else PartyName End as Customer, Cash, Credit, cast(0 as bit) as BankCard, 1 as Rec, SaleAmount, ReturnAmount" & vbCrLf _
      + " from ReplacementHeader r " & vbCrLf _
      + " inner join SaleReturnHeader h on r.Returnid = h.Returnid and r.Returndate = h.ReturnDate" & vbCrLf _
      + " inner join saleReturnbody b on b.ReturnID = h.ReturnID and b.ReturnDate = h.ReturnDate " & vbCrLf _
      + " inner join Products p on p.Productid = b.ProductID" & vbCrLf _
      + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
      + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
      + " left outer join Parties Pr on Pr.PartyID = h.CustomerID" & vbCrLf _
      + " where ReplaceID= " & Val(TxtReplaceID.Text) & " and ReplaceDate='" & DtpReplaceDate.DateValue & "'"

   vStrSQL = " select username, ssid, rsid, ReplaceID, ReplaceDate, 'Sale' as Type, h.TotalAmount as tbill, Billtime, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges," & vbCrLf _
      + " isnull(r.cashReceived,0) as cashReceived, isnull(h.BankAmount,0) as BankAmount, p.ProductID, p.productname, unitname, b.qty, b.price as price, b.amount, b.DiscVal, isnull(b.SC,0) as SC" & vbCrLf _
      + " , Case when CustomerID = '621' then isnull(CustomerName,PartyName) Else PartyName End as Customer, h.empid, empname, Cash, Credit, BankCard, 1 as Rec, SaleAmount, ReturnAmount, p.ProductID, p.ItemCode, right('00'+ cast(b.ColourID as varchar(2)),2) as ColourID, right('00'+ cast(b.SizeID as varchar(2)),2) as SizeID, ColourName, SizeName" & vbCrLf _
      + " from ReplacementHeader r " & vbCrLf _
      + " inner join SaleHeader h on r.SSid = h.Sid" & vbCrLf _
      + " inner join salebody b on b.sid = h.sid" & vbCrLf _
      + " inner join Products p on p.Productid = b.ProductID" & vbCrLf _
      + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
      + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
      + " Left outer join employees e on e.empid = h.empid" & vbCrLf _
      + " left outer join Parties Pr on Pr.PartyID = h.CustomerID" & vbCrLf _
      + " Left outer join Colours Col on Col.Colourid = b.ColourID" & vbCrLf _
      + " Left Outer join Sizes Sz on Sz.SizeID = b.SizeID " & vbCrLf _
      + " where r.SID= " & Val(TxtSID.Text) & vbCrLf _
      + " union all"
    vStrSQL = vStrSQL & vbCrLf _
      + " select username, ssid, rsid, ReplaceID, ReplaceDate, 'Return', h.TotalAmount as tbill, ReturnTime as Billtime, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, " & vbCrLf _
      + " isnull(r.cashReceived,0) as cashReceived, isnull(h.BankAmount,0) as BankAmount, p.ProductID, p.productname, unitname, b.qty, b.price as price, b.amount, b.DiscVal, isnull(b.SC,0) as SC" & vbCrLf _
      + " , Case when CustomerID = '621' then isnull(CustomerName,PartyName) Else PartyName End as Customer, h.empid, empname, Cash, Credit, BankCard, 1 as Rec, SaleAmount, ReturnAmount, p.ProductID, p.ItemCode, right('00'+ cast(b.ColourID as varchar(2)),2) as ColourID, right('00'+ cast(b.SizeID as varchar(2)),2) as SizeID, ColourName, SizeName" & vbCrLf _
      + " from ReplacementHeader r " & vbCrLf _
      + " inner join SaleReturnHeader h on r.RSID = h.SID" & vbCrLf _
      + " inner join saleReturnbody b on b.SID = h.SID" & vbCrLf _
      + " inner join Products p on p.Productid = b.ProductID" & vbCrLf _
      + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
      + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
      + " Left outer join employees e on e.empid = h.empid" & vbCrLf _
      + " left outer join Parties Pr on Pr.PartyID = h.CustomerID" & vbCrLf _
      + " Left outer join Colours Col on Col.Colourid = b.ColourID" & vbCrLf _
      + " Left Outer join Sizes Sz on Sz.SizeID = b.SizeID " & vbCrLf _
      + " where r.SID= " & Val(TxtSID.Text)


   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   
   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
        
   If ObjRegistry.LaserPrintofSaleInvoice = True Or InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
'      Set RptReportViewer.Report = New CrpReplacementHalf
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpReplacementHalf.rpt")
'      RptReportViewer.Report.PaperSize = crPaperA4
'      RptReportViewer.Report.PaperOrientation = crLandscape
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.x
      RptReportViewer.Report.RightMargin = 225
   ElseIf InStr(1, Printer.DeviceName, "CBM1000") > 0 Then
      Set RptReportViewer.Report = New CrpReplacementCBM
   ElseIf InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
      Set RptReportViewer.Report = New CrpReplacementAurora
      RptReportViewer.Report.LeftMargin = 225
      RptReportViewer.Report.RightMargin = 0
      RptReportViewer.Report.TopMargin = 255
   Else 'InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpReplacementAurora.rpt")
      RptReportViewer.Report.TopMargin = 0
      RptReportViewer.Report.LeftMargin = 0
      RptReportViewer.Report.RightMargin = 0
'      Set RptReportViewer.Report = New CrpReplacementAurora
      'RptReportViewer.Report.LeftMargin = 0
      'RptReportViewer.Report.RightMargin = 0
   End If
   

   
   'RptReportViewer.Report.PaperOrientation = crPortrait
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ReportTitle = "Replacement Invoice"
    
    'RptReportViewer.Report.LeftMargin = 0
    'RptReportViewer.Report.RightMargin = 0
   
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & ObjRegistry.CompanyCity
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.CompanyPhoneNo
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.AddSpace = True, ".", "")
   RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(ObjRegistry.CashReceived)
   RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(ObjRegistry.Statement)
   'CN.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
   'If vIsNewRecord = True Then Call ActivityLog("Replacement Invoice", eAdd, TxtReplaceID.Text, DtpReplaceDate.DateValue)
   'RptReportViewer.Report.SelectPrinter "RASDD.DLL", "CBM1000 Partial Cut", "Com1" 'RptReportViewer.Report.SelectPrinter  "RASDD.DLL", "CBM1000 Partial Cut", "Com1"
   'RptReportViewer.Show
    If ObjRegistry.IsPortrait = False Then RptReportViewer.Report.PaperOrientation = crLandscape
   RptReportViewer.Report.PrintOut False
   
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmReplacementInvoice, eRePrint, TxtReturnID.Text, DtpReturnDate.DateValue, "RePrinted Amount: " & Val(TxtSAmount.Text) - Val(TxtRAmount.Text))
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnRProduct_Click()
   If FunSelectRProduct(ssButton, True) = True Then
      TxtRQty.SetFocus
   Else
      TxtRCode.SetFocus
   End If
End Sub

Private Sub BtnReturnAll_Click()
   PopulateSaleDataToGridReturn
End Sub

Private Sub BtnSale_Click()
   If FunSelectSale(ssButton, False) = True Then
      BtnReturnAll.SetFocus
   Else
      TxtRBillID.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Discount ''''''''''''''
   If Val(ObjUserSecurity.AllowMaximmDiscPer) <> 0 Then
      If Val(TxtSBillDiscPer.Text) > Val(ObjUserSecurity.AllowMaximmDiscPer) Or Val(TxtRBillDiscPer.Text) > Val(ObjUserSecurity.AllowMaximmDiscPer) Then
         MsgBox "Discount greater than Fixed Discount", vbCritical, "Error"
         Exit Sub
      End If
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniReplacementInvoice", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
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
   If CN.Execute("Select * From AdminClosing where ToUserNo = " & vUser & " and EntryDate = '" & DtpReplaceDate.DateValue & "'").RecordCount > 0 Then
      MsgBox "You are not authorized to Add Record in Closing Dates.", vbCritical, "Alert"
      Exit Sub
   End If
   
   '''''''''''''''''''''''Check Employee '''''''''''''''''''''''''''''''''
   If ObjRegistry.EmployeeMandatory = True And TxtEmployeeID.Text = "" Then
      MsgBox "Please Select Employee", vbInformation, Me.Caption
      If TxtEmployeeID.Visible = True Then TxtEmployeeID.SetFocus
      Exit Sub
   End If
   
   ''''''''''''''''''''''''Check Organization '''''''''''''''''''''''''''''''''
  If ObjRegistry.OrganizationMandatory = True And TxtOrganizationID.Text = "" Then
    MsgBox "Please Select Organization", vbInformation, Me.Caption
    If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
    Exit Sub
  End If
  
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
'    ssql = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where ToUserNo = " & vUser & " and Entrydate <='" & Date & "'"
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where ToUserNo = " & vUser
    With CN.Execute(vStrSQL)
        If .Fields(0).Value >= DtpReplaceDate.DateValue Then
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
    
    '''''''''''''''''''''''Check Current Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.CurrentDateDataEntry = True Then
       If DtpBillDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
   ''''''''''''''''''''''''''''''''''''''
   '   'Body Validation

   'validation has been performed when a row is added to the GridReturn
   RsReturnBody.Filter = 0
   If RsReturnBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to In", vbExclamation, "Alert"
      If TxtRCode.Visible And TxtRCode.Enabled Then TxtRCode.SetFocus
      Exit Sub
   End If
   
   'validation has been performed when a row is added to the GridSale
   RsSaleBody.Filter = 0
   If RsSaleBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Out", vbExclamation, "Alert"
      If TxtSCode.Visible And TxtSCode.Enabled Then TxtSCode.SetFocus
      Exit Sub
   End If
   
   If Val(TxtRBillID.Text) <> 0 Then
      RsReturnBody.Filter = 0
      RsReturnBody.MoveFirst
      For vCounter = 1 To RsReturnBody.RecordCount
         With CN.Execute("select isnull(Qty,0) as Qty from salebody where BillID = " & Val(TxtRBillID.Text) & " and BillDate = '" & DtpRBillDate.DateValue & "' and ProductID = " & Val(RsReturnBody!Productid))
            If .RecordCount > 0 Then
               If !Qty < RsReturnBody!Qty Then
                  MsgBox "Sale Quantity is less than Sale Return Quantity of Code = " & RsReturnBody!Code, vbExclamation, "Alert"
                  If TxtRCode.Visible And TxtRCode.Enabled Then TxtRCode.SetFocus
                  Exit Sub
               End If
            Else
               MsgBox "Sale of This Bill ID can not be found", vbExclamation, "Alert"
               If TxtRBillID.Visible And TxtRBillID.Enabled Then TxtRBillID.SetFocus
               Exit Sub
            End If
         End With
         RsReturnBody.MoveNext
      Next vCounter
   End If
   
'   FrmReplacePrint.SubClearFields
   If LblAmount.Caption = "Cash Paid" Then FrmReplacePrint.TxtCashReceivedBank.Enabled = False Else FrmReplacePrint.TxtCashReceivedBank.Enabled = True
   If FrmReplacePrint.OptCash.Visible Then FrmReplacePrint.OptCash.SetFocus
   FrmReplacePrint.TxtNetAmount.Text = TxtTotalAmount.Caption
   FrmReplacePrint.LblCaption.Caption = LblAmount.Caption
   FrmReplacePrint.LblCreditCaption.Caption = LblAmount.Caption
   FrmReplacePrint.LblBankCaption.Caption = LblAmount.Caption
   
   FrmReplacePrint.ParaInPrint = True
   FrmReplacePrint.ParaInChoice = "Cash"
   
   FrmReplacePrint.ParaInOgtanizationID = Val(TxtOrganizationID.Text)
   FrmReplacePrint.ParaInID = Val(TxtBillID.Text)
   FrmReplacePrint.ParaInDate = DtpBillDate.DateValue
  
   FrmReplacePrint.Show vbModal, Me

   If FrmReplacePrint.ParaOutSelection = False Then Exit Sub
   
   If DtpReplaceDate.Enabled And DtpReplaceDate.Date <> Date And DateFlag = True Then
      If MsgBox("Are you sure to Change Replace Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
         DtpReplaceDate.DateValue = Date
         DtpReturnDate.DateValue = Date
         DtpBillDate.DateValue = Date
      End If
      DateFlag = False
   End If
   
   If DtpReplaceDate.Enabled Then
      If CN.Execute("Select * from ReplacementHeader where ReplaceID = " & Val(TxtReplaceID.Text) & " and ReplaceDate = '" & DtpReplaceDate.DateValue & "'").RecordCount > 0 Then
         TxtReplaceID.Text = FunGetMaxID
      End If
   End If
   
   
  'Saving record
   CN.BeginTrans
      
   If DtpReplaceDate.Enabled Then
      DtpReturnDate.DateValue = DtpReplaceDate.DateValue
      DtpBillDate.DateValue = DtpReplaceDate.DateValue
      TxtBillID.Text = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where BillDate = '" & DtpBillDate.DateValue & "'").Fields(0)
      TxtReturnID.Text = CN.Execute("Select isnull(max(ReturnID),0)+1 from SaleReturnHeader where ReturnDate = '" & DtpReturnDate.DateValue & "'").Fields(0)
   End If

   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmReplacementInvoice, eEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Amount: " & Val(FrmReplacePrint.TxtNetAmount.Text))
'   If vIsNewRecord = False Then Call ActivityLog("Replacement Invoice", eEdit, TxtReplaceID.Text, DtpReplaceDate.DateValue)
   
'   Call UserActivities
   If vIsNewRecord = True Then
      TxtRSID.Text = ""
      TxtBillID.Text = FunGetMaxIDOut
      TxtReturnID.Text = FunGetMaxIDIn
   End If
   Call SaveReturnIn
   
   
'   Dim Rs As New ADODB.Recordset
'   'Return

'
'   ssql = "select * from SaleReturnHeader where SID=" & Val(TxtRSID.Text)
'   With Rs
'      .Open ssql, cn, adOpenDynamic, adLockPessimistic
'      If .BOF Then
'         .AddNew
'         !ReturnID = Val(TxtReturnID.Text)
'         !ReturnDate = DtpReturnDate.DateValue
'         !ReturnTime = Now
'         !UserNo = vUser
'      End If
'      !isReplace = 1
'      !isPosted = 0
'      !isTransfer = 0
'      !IsSync = 0
'      !ServerEntry = Now
'      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
'      !BillID = IIf(Val(TxtRBillID.Text) = 0, Null, Val(TxtRBillID.Text))
'      !BillDate = IIf(IsNull(!BillID), Null, DtpRBillDate.DateValue)
'      !StoreID = TxtStoreID.Text
'      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
'      !EmpComm = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text))
'      !TotalAmount = Round(vTotalReturn - vTotReturnDisc)
'      !BillDisc = IIf(TxtRBillDisc.Text = "", Null, Val(TxtRBillDisc.Text))
'      !BillDiscPer = IIf(TxtRBillDiscPer.Text = "", Null, Val(TxtRBillDiscPer.Text))
'      !ServiceCharges = IIf(TxtRServiceCharges.Text = "", Null, Val(TxtRServiceCharges.Text))
'      If FrmReplacePrint.OptBankCard.Value = True Then
'         !InvoiceNo = FrmReplacePrint.TxtInvoiceNo.Text
'         !Commision = FrmReplacePrint.TxtCommision.Text
'         !BankMachineID = FrmReplacePrint.TxtBankMachineID.Text
'         !CashPaid = 0
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(FrmReplacePrint.TxtBankCustomer.Text) = "", Null, FrmReplacePrint.TxtBankCustomer.Text)
'      End If
'      If FrmReplacePrint.OptCash.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !CashPaid = TxtRNetAmount.Text
'         !CustomerID = "621"
'         !BankMachineID = Null
'         !CustomerName = IIf(Trim(FrmReplacePrint.TxtCashCustomer.Text) = "", Null, FrmReplacePrint.TxtCashCustomer.Text)
'      End If
'      If FrmReplacePrint.OptCredit.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         If LblAmount.Caption = "Cash Paid" Then
'            !CashPaid = Val(FrmReplacePrint.TxtCashReceivedCredit.Text)
'            !BankAmount = Val(FrmReplacePrint.TxtBankAmount.Text)
'         Else
'            !CashPaid = 0
'            !BankAmount = 0
'         End If
'         !BankMachineID = Null
'         !CustomerID = FrmReplacePrint.TxtCustomerID.Text
'         !BankMachineID = IIf(Trim(FrmReplacePrint.TxtBankMachineCreditID.Text) = "", Null, FrmReplacePrint.TxtBankMachineCreditID.Text)
''         !BankAmount = Val(FrmReplacePrint.TxtBankAmount.Text)
'         If Val(FrmReplacePrint.TxtBankMachineCreditID.Text) > 0 Then
'            !Commision = Val(FrmReplacePrint.TxtCommision.Text)
'         Else
'            !Commision = Null
'         End If
'         !CustomerName = Null
'      End If
'      !BankCard = FrmReplacePrint.OptBankCard.Value
'      !Cash = FrmReplacePrint.OptCash.Value
'      !Credit = FrmReplacePrint.OptCredit.Value
'      '!Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
'      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
'      !ManualBillNo = IIf(Trim(TxtManualBillNo.Text) = "", "", TxtManualBillNo.Text)
'      .Update
'      .Close
'      If vIsNewRecord = True Then TxtRSID.Text = cn.Execute("select @@identity").Fields(0).Value
'   End With
'   vStrDetailin = " In:-"
'   With RsReturnBody
'      .Filter = 0
'      .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !SID = Val(TxtRSID.Text)
'         !ReturnID = Val(TxtReturnID.Text)
'         !ReturnDate = DtpReturnDate.DateValue
'         !StoreID = Val(TxtStoreID.Text)
'         vStrDetailin = vStrDetailin & " (P" & !Productid & " Q" & !Qty & " A" & !Amount & ")"
'         .MoveNext
'      Next vCounter
'      .UpdateBatch
'   End With
'   RsReturnSerial.Filter = 0
'   If RsReturnSerial.RecordCount > 0 Then
'     With RsReturnSerial
''      .Filter = 0
'      .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !ReturnID = Val(TxtRSID.Text)
'         !ReturnDate = DtpReturnDate.DateValue
'
'         RsPurchaseSerial.Filter = "Serial = " & RsReturnSerial!Serial
'         If RsPurchaseSerial.RecordCount > 0 Then
'            RsPurchaseSerial!SerialAdd = 1
'            RsPurchaseSerial.Update
'         End If
'         .MoveNext
'      Next vCounter
'      .UpdateBatch
'     End With
'   End If
   'Sale
   If vIsNewRecord = True Then TxtSSID.Text = ""
   With CN.Execute("select * from Salebody where SID = " & Val(TxtSSID.Text))
    GridSale.Redraw = False
    GridSale.MoveFirst
    While Not .EOF
        CN.Execute "Exec UpdateStockPlus " & TxtStoreID.Text & "," & Val(!Productid) & "," & !Qty & "," & Val(TxtReplaceID.Text) & ",'" & DtpReplaceDate.DateValue & "'"
        GridSale.MoveNext
        .MoveNext
    Wend
   End With
   GridSale.Redraw = True
   GridSale.MoveLast
   
'   ssql = "select * from SaleHeader where SID = " & Val(TxtSSID.Text)
'   With Rs
'      .Open ssql, cn, adOpenStatic, adLockOptimistic
'      If .BOF Then
'         .AddNew
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         !BillTime = Now
'         !UserNo = vUser
'      End If
'      !isReplace = 1
'      !isTransfer = 0
'      !isPosted = 0
'      !IsSync = 0
'      !ServerEntry = Now
'      !StoreID = TxtStoreID.Text
'      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
'      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
'      !EmpComm = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text))
'      !TotalAmount = Round(vTotal - vTotDisc)
'      !BillDisc = IIf(TxtSBillDisc.Text = "", Null, Val(TxtSBillDisc.Text))
'      !BillDiscPer = IIf(TxtSBillDiscPer.Text = "", Null, Val(TxtSBillDiscPer.Text))
'      !ServiceCharges = IIf(TxtSServiceCharges.Text = "", Null, Val(TxtSServiceCharges.Text))
'      If FrmReplacePrint.OptBankCard.Value = True Then
'         !InvoiceNo = FrmReplacePrint.TxtInvoiceNo.Text
'         !Commision = FrmReplacePrint.TxtCommision.Text
'         !BankMachineID = FrmReplacePrint.TxtBankMachineID.Text
'         !CashReceived = 0
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(FrmReplacePrint.TxtBankCustomer.Text) = "", Null, FrmReplacePrint.TxtBankCustomer.Text)
'         !CashReceived = Val(FrmReplacePrint.TxtCashReceivedBank.Text)
'      End If
'      If FrmReplacePrint.OptCash.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !BankMachineID = Null
'         !CashReceived = Val(FrmReplacePrint.TxtCashReceivedCash.Text)
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(FrmReplacePrint.TxtCashCustomer.Text) = "", Null, FrmReplacePrint.TxtCashCustomer.Text)
'      End If
'      If FrmReplacePrint.OptCredit.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !BankMachineID = Null
'         If LblAmount.Caption = "Cash Paid" Then
'            !CashReceived = 0
'            !BankAmount = 0
'         Else
'            !CashReceived = Val(FrmReplacePrint.TxtCashReceivedCredit.Text)
'            !BankAmount = Val(FrmReplacePrint.TxtBankAmount.Text)
'         End If
'         !CustomerID = FrmReplacePrint.TxtCustomerID.Text
'         !BankMachineID = IIf(Trim(FrmReplacePrint.TxtBankMachineCreditID.Text) = "", Null, FrmReplacePrint.TxtBankMachineCreditID.Text)
''         !BankAmount = Val(FrmReplacePrint.TxtBankAmount.Text)
'         If Val(FrmReplacePrint.TxtBankMachineCreditID.Text) > 0 Then
'            !Commision = Val(FrmReplacePrint.TxtCommision.Text)
'         Else
'            !Commision = Null
'         End If
'         !CustomerName = Null
'      End If
'      !BankCard = FrmReplacePrint.OptBankCard.Value
'      !Cash = FrmReplacePrint.OptCash.Value
'      !Credit = FrmReplacePrint.OptCredit.Value
'
'      '!Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
'      '!Remarks = IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text)
'      !ManualBillNo = IIf(Trim(TxtManualBillNo.Text) = "", "", TxtManualBillNo.Text)
'      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
'      .Update
'      .Close
'      If vIsNewRecord = True Then TxtSSID.Text = cn.Execute("select @@identity").Fields(0).Value
'   End With
'   vStrDetailOut = " Out:-"
'   With RsSaleBody
'      .Filter = 0
'      .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !SID = Val(TxtSSID.Text)
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         !StoreID = Val(TxtStoreID.Text)
'         vStrDetailOut = vStrDetailOut & " (P" & Val(!Productid) & " Q" & !Qty & " A" & !Amount & ")"
'         ssql = "exec UpdateStockMinus " & Val(TxtStoreID.Text) & "," & Val(!Productid) & "," & !Qty & "," & Val(TxtReplaceID.Text) & ",'" & DtpReplaceDate.DateValue & "'"
'         cn.Execute (ssql)
'         .MoveNext
'      Next vCounter
'      .UpdateBatch
'   End With
   Call SaveSaleOut
   
   RsBodySerial.Filter = 0
   If RsBodySerial.RecordCount > 0 Then
     With RsBodySerial
'      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !BillID = Val(TxtSSID.Text)
         !BillDate = DtpBillDate.DateValue
         
         RsPurchaseSerial.Filter = "Serial = " & RsBodySerial!Serial
         If RsPurchaseSerial.RecordCount > 0 Then
            RsPurchaseSerial!SerialAdd = 0
            RsPurchaseSerial.Update
         End If
         .MoveNext
      Next vCounter
      .UpdateBatch
     End With
   End If
   
   RsPurchaseSerial.Filter = ""
   If RsPurchaseSerial.RecordCount > 0 Then RsPurchaseSerial.UpdateBatch
   RsReturnSerial.Filter = ""
   If RsReturnSerial.RecordCount > 0 Then RsReturnSerial.UpdateBatch
   
   'Replace
   If vIsNewRecord = True Then TxtSID.Text = ""
'   ssql = "select * from ReplacementHeader where SID = " & Val(TxtSID.Text) & " AND StoreID = " & Val(TxtStoreID.Text)
'
'   With Rs
'      .Open ssql, cn, adOpenStatic, adLockOptimistic
'      If .BOF Then
'         .AddNew
'         !ReplaceId = Val(TxtReplaceID.Text)
'         !ReplaceDate = DtpReplaceDate.DateValue
'         !UserNo = vUser
'         vIsNewRecord = True
'      End If
'      !ServerEntry = Now
'      !isTransfer = 0
'      !IsSync = 0
'      !SSID = Val(TxtSSID.Text)
'      !RSID = Val(TxtRSID.Text)
'      !BillID = Val(TxtBillID.Text)
'      !BillDate = DtpBillDate.DateValue
'      !ReturnID = Val(TxtReturnID.Text)
'      !ReturnDate = DtpReturnDate.DateValue
'      !StoreID = Val(TxtStoreID.Text)
'      !SaleAmount = Val(TxtSNetAmount.Text)
'      !ReturnAmount = Val(TxtRNetAmount.Text)
'      If FrmReplacePrint.OptCredit = True Then FrmReplacePrint.TxtCashReceivedCash.Text = FrmReplacePrint.TxtCashReceivedCredit.Text
'      !CashReceived = Val(FrmReplacePrint.TxtCashReceivedCash.Text)
'      !Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
'      !isPosted = 0
'
''      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
'      .Update
'      .Close
'      If vIsNewRecord = True Then TxtSID.Text = cn.Execute("select @@identity").Fields(0).Value
'   End With
   
   Call SaveReplacement
'   If vIsNewRecord = True Then Call ActivityLog("Sale Return Invoice", eAdd, TxtReturnID.Text, DtpReturnDate.DateValue)
   
   '''-------------- Mobile SMS -------------------
    If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = "Saved Replacement ID:" & TxtReplaceID.Text & vbCrLf & " Date:" & Format(DtpReplaceDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtSBillDisc.Text) = 0, "", " SDisc:" & TxtSBillDisc.Text) & IIf(Val(TxtRBillDisc.Text) = 0, "", " RDisc:" & TxtRBillDisc.Text) & vbCrLf & " NetAmt:" & TxtTotalAmount.Caption
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetailin & vStrDetailOut, "") & "','')"
               CN.Execute ssql
            End If
         Next
   End If
   
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmReplacementInvoice, eAdd, TxtReplaceID.Text, DtpReplaceDate.DateValue, GridReturn.rows - 1 & " New In Product/s Added Amount: " & Val(TxtRNetAmount.Text))
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmReplacementInvoice, eAdd, TxtReplaceID.Text, DtpReplaceDate.DateValue, GridSale.rows - 1 & " New Out Product/s Added Amount: " & Val(TxtSNetAmount.Text))
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmReplacementInvoice, eAdd, TxtReturnID.Text, DtpReturnDate.DateValue, "Net Amount: " & Val(FrmReplacePrint.TxtNetAmount.Text))
   
   CN.CommitTrans
'   Char.Speak "Thank you for comming"
   'If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
   If FrmReplacePrint.ChkPrint.Value = 1 Then Call BtnPrint_Click
   'End If
   FrmReplacePrint.SubClearFields
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   GridReturn.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGridReturn()
   RsReturnBody.Filter = 0
   If RsReturnBody.State = adStateOpen Then RsReturnBody.Close
   RsReturnBody.Open "Select * from SaleReturnBody where SID=" & Val(TxtRSID.Text), CN, adOpenDynamic, adLockBatchOptimistic
   If RsReturnBody.RecordCount > 0 Then
      ssql = "select p.productname, b.code, b.ColourID, b.SizeID, ColourName, SizeName, b.* from SaleReturnBody b join products p on p.productid = b.productid Left outer join Colours Col on Col.Colourid = b.ColourID Left Outer join Sizes Sz on Sz.SizeID = b.SizeID where SID=" & Val(TxtRSID.Text)
      With CN.Execute(ssql)
         GridReturn.Redraw = False
         GridReturn.MoveFirst
         GridReturn.RemoveAll
         GridReturn.AllowAddNew = True
         vTotReturnDisc = 0
         vTotalReturn = 0
         TxtTotReturnQty.Text = 0
         While Not .EOF
            GridReturn.AddNew
            GridReturn.Columns("ProductID").Text = !Productid
            GridReturn.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            GridReturn.Columns("ProductName").Text = !ProductName
            vIsSerial = !isSerial
            GridReturn.Columns("ColourID").Value = IIf(IsNull(!ColourID), "", !ColourID)
            GridReturn.Columns("ColourName").Value = IIf(IsNull(!ColourName), "", !ColourName)
            GridReturn.Columns("SizeID").Value = IIf(IsNull(!SizeID), "", !SizeID)
            GridReturn.Columns("SizeName").Value = IIf(IsNull(!SizeName), "", !SizeName)
            GridSale.Columns("StoreID").Value = IIf(IsNull(!StoreID), "", !StoreID)
            GridReturn.Columns("Qty").Value = !Qty
            GridReturn.Columns("Price").Value = !Price
            GridReturn.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            GridReturn.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            GridReturn.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            GridReturn.Columns("SC").Value = IIf(IsNull(!SC), "", !SC)
            GridReturn.Columns("Amount").Value = !Amount
            GridReturn.Columns("TotalAmount").Value = Val(!Qty) * (Val(!Price) + Val(IIf(IsNull(!SC), "", !SC)))
            vTotReturnDisc = vTotReturnDisc + Val(!DiscVal)
            vTotalReturn = vTotalReturn + GridReturn.Columns("TotalAmount").Value
            TxtTotReturnQty.Text = Val(TxtTotReturnQty.Text) + GridReturn.Columns("Qty").Value
            .MoveNext
         Wend
         .Close
      End With
      Call SubCalculateRBody
      GridReturn.AddNew
      GridReturn.Columns("productid").Text = " "
      GridReturn.AllowAddNew = False
      GridReturn.Redraw = True
   End If
   RsReturnSerial.Filter = 0
   If RsReturnSerial.State = adStateOpen Then RsReturnSerial.Close
   vStrSQL = "select * from SaleReturnSerial  where ReturnID=" & Val(TxtRSID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
   RsReturnSerial.Open vStrSQL, CN, adOpenDynamic, adLockBatchOptimistic
   
'   PopulateDataToGridRSerial
End Sub

Private Sub PopulateSaleDataToGridReturn()
   If Val(TxtSSID.Text) = 0 And Val(TxtRBillID.Text) <> 0 Then
      vStrSQL = "Select sid from saleheader where billid = " & Val(TxtRBillID.Text) & " and billdate = '" & DtpRBillDate.DateValue & "'"
      With CN.Execute(vStrSQL)
         If Not .EOF Then
            TxtSSID.Text = .Fields("sid").Value
         End If
      End With
   End If
   RsReturnBody.Filter = 0
   If RsReturnBody.State = adStateOpen Then RsReturnBody.Close
   RsReturnBody.Open "Select * from SaleReturnBody where SID = " & Val(TxtRSID.Text), CN, adOpenDynamic, adLockBatchOptimistic
   ssql = "select p.productname, b.code, b.* from SaleBody b join products p on p.productid = b.productid left outer join units u on p.unitid = u.unitid where SID=" & Val(TxtSSID.Text)
   With CN.Execute(ssql)
      If .RecordCount > 0 Then
         GridReturn.Redraw = False
         GridReturn.MoveFirst
         GridReturn.RemoveAll
         GridReturn.AllowAddNew = True
         TxtTotReturnQty.Text = 0
         vTotReturnDisc = 0
         vTotalReturn = 0
         While Not .EOF
            RsReturnBody.AddNew
            RsReturnBody!Productid = !Productid
            RsReturnBody!Code = !Code
            RsReturnBody!Qty = !Qty
            RsReturnBody!Price = !Price
            RsReturnBody!DiscPC = !DiscPC
            RsReturnBody!DiscPer = !DiscPer
            RsReturnBody!DiscVal = !DiscVal
            RsReturnBody!Cost = !Cost
            RsReturnBody!isProduct = !isProduct
            RsReturnBody!Amount = !Amount
            RsReturnBody.Update
            GridReturn.AddNew
            GridReturn.Columns("ProductID").Text = !Productid
            GridReturn.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            GridReturn.Columns("ProductName").Text = !ProductName
            GridReturn.Columns("Qty").Value = !Qty
            GridReturn.Columns("Price").Value = !Price
            GridReturn.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            GridReturn.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            GridReturn.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            GridReturn.Columns("Amount").Value = !Amount
            GridReturn.Columns("IsProduct").Value = Abs(!isProduct)
            GridReturn.Columns("TotalAmount").Value = Val(!Price) * Val(!Qty)
            GridReturn.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            vTotReturnDisc = vTotReturnDisc + Val(!DiscVal)
            vTotalReturn = vTotalReturn + GridReturn.Columns("TotalAmount").Value
            TxtTotReturnQty.Text = Val(TxtTotReturnQty.Text) + GridReturn.Columns("Qty").Value
            .MoveNext
         Wend
         .Close
         Call SubCalculateRBody
         GridReturn.AddNew
         GridReturn.Columns("productid").Text = " "
         GridReturn.AllowAddNew = False
         GridReturn.Redraw = True
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
      vRandomID = Rnd() * 11111 & " " & Format(Now, "dd/mm hh:mm:ss")
      vDate = IIf(vSystemDate = True, CN.Execute("Select SystemDate From SystemDate").Fields(0).Value, vServerDate)
      DtpReplaceDate.DateValue = IIf(vSystemDate = True, IIf(IsNull(vDate), Date, vDate), IIf(Format(Now, "hh") >= vHDiff, vDate, DateAdd("d", -1, vDate)))
      DtpBillDate.DateValue = DtpReplaceDate.DateValue
      DtpRBillDate.DateValue = DtpReplaceDate.DateValue
      DtpReturnDate.DateValue = DtpReplaceDate.DateValue
      
      TxtReplaceID.Text = FunGetMaxID()
      TxtBillID.Text = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where BillDate = '" & DtpBillDate.DateValue & "'").Fields(0)
      TxtReturnID.Text = CN.Execute("Select isnull(max(ReturnID),0)+1 from SaleReturnHeader where ReturnDate = '" & DtpReturnDate.DateValue & "'").Fields(0)
      
      Call PopulateDataToGridSale
      Call PopulateDataPurchaseSerial
      Call PopulateDataToGridReturn
      TxtRBillID.Enabled = True
      DtpRBillDate.Enabled = True
      BtnReturnAll.Enabled = True
      TxtRCode.Enabled = True
      BtnRProduct.Enabled = True
      BtnSale.Enabled = True
      TxtSCode.Enabled = True
      BtnSProduct.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      LblStock.Visible = False
      LblStockCaption.Visible = False
'      DtpReplaceDate.Enabled = True
      If TxtRCode.Visible And TxtRCode.Enabled Then TxtRCode.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      BtnSale.Enabled = False
      TxtRBillID.Enabled = False
      DtpRBillDate.Enabled = False
      BtnReturnAll.Enabled = False
      
      DtpReplaceDate.Enabled = False
      
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      TxtRCode.Enabled = True
      BtnRProduct.Enabled = True
      TxtRCode.SetFocus
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
      If TxtRCode.Enabled Then TxtRCode.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub GetSaleReturn()
   On Error GoTo ErrorHandler
   ssql = " select h.*, OrganizationName, BankMachineName, c.AccountName, StoreName FROM SaleReturnHeader h left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join BankMachines b on b.BankMachineid = h.BankMachineid left outer join Organizations o on o.OrganizationID = h.OrganizationID inner join stores s on s.storeid = h.storeid where h.SID=" & Val(TxtRSID.Text) & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With CN.Execute(ssql)
      If Not .BOF Then
         TxtRBillID.Text = IIf(IsNull(!BillID), "", !BillID)
         DtpRBillDate.DateValue = IIf(IsNull(!BillDate), "", !BillDate)
         TxtStoreID.Text = !StoreID
         TxtStoreName.Text = !StoreName
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         TxtRBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
         TxtRBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
         TxtRServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
         TxtManualBillNo.Text = IIf(IsNull(!ManualBillNo), "", !ManualBillNo)
         FrmReplacePrint.TxtCashReceivedCredit.Text = !CashPaid
         FrmReplacePrint.TxtCommision.Text = IIf(IsNull(!Commision), "", !Commision)
         FrmReplacePrint.TxtBankMachineCreditID.Text = IIf(IsNull(!BankMachineID), "", !BankMachineID)
         FrmReplacePrint.TxtBankMachineCreditName.Text = IIf(IsNull(!BankMachineName), "", !BankMachineName)
         FrmReplacePrint.TxtBankAmount.Text = IIf(IsNull(!BankAmount), "", !BankAmount)
      End If
      .Close
   End With
      
   PopulateDataToGridReturn
'   RsReturnBody.Filter = 0
'   If RsReturnBody.State = adStateOpen Then RsReturnBody.Close
'   RsReturnBody.Open "Select * from SaleReturnBody where ReturnId=" & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
'   If RsReturnBody.RecordCount > 0 Then
'      sSql = "select p.productname, b.code,b.* from SaleReturnBody b join products p on p.productid = b.productid where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
'      With CN.Execute(sSql)
'         GridReturn.Redraw = False
'         GridReturn.MoveFirst
'         GridReturn.RemoveAll
'         GridReturn.AllowAddNew = True
'         vTotReturnDisc = 0
'         vTotalReturn = 0
'         While Not .EOF
'            GridReturn.AddNew
'            GridReturn.Columns("ProductID").Text = !ProductID
'            GridReturn.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
'            GridReturn.Columns("ProductName").Text = !ProductName
'            GridReturn.Columns("Qty").Value = !Qty
'            GridReturn.Columns("Price").Value = !Price
'            GridReturn.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
'            GridReturn.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
'            GridReturn.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
'            GridReturn.Columns("Amount").Value = !Amount
'            GridReturn.Columns("IsProduct").Value = Abs(!IsProduct)
'            GridReturn.Columns("TotalAmount").Value = Val(!Price) * Val(!Qty)
'            GridReturn.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
'            vTotReturnDisc = vTotReturnDisc + Val(!DiscVal)
'            vTotalReturn = vTotalReturn + GridReturn.Columns("TotalAmount").Value
'            .MoveNext
'         Wend
'         .Close
'      End With
'      Call SubCalculateRBody
'      GridReturn.AddNew
'      GridReturn.Columns("ProductID").Text = " "
'      GridReturn.AllowAddNew = False
'      GridReturn.Redraw = True
'   End If
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   GridReturn.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn And Shift = vbShiftMask Then
      Select Case ActiveControl.Name
      Case TxtPCode.Name
         If FunSelectPProduct(ssValidate, False) = True Then TxtPSerial.SetFocus
      'Return
      Case TxtRCode.Name
         If FunSelectRProduct(ssValidate, False) = True Then TxtRQty.SetFocus
      Case TxtRQty.Name
         If TxtRPrice.Enabled Then TxtRPrice.SetFocus Else TxtRDiscPC.SetFocus
      Case TxtRPrice.Name
         TxtRDiscPC.SetFocus
      Case TxtRDiscPC.Name
         TxtRDiscPer.SetFocus
      Case TxtRDiscPer.Name
         If TxtRDiscVal.Enabled Then TxtRDiscVal.SetFocus
      'Sale
      Case TxtPCode.Name
         If FunSelectPProduct(ssValidate, False) = True Then TxtPSerial.SetFocus
      Case TxtSCode.Name
         If FunSelectSProduct(ssValidate, False) = True Then TxtSQty.SetFocus
      Case TxtSQty.Name
         If TxtSPrice.Enabled Then TxtSPrice.SetFocus Else TxtSDiscPC.SetFocus
      Case TxtSPrice.Name
         TxtSDiscPC.SetFocus
      Case TxtSDiscPC.Name
         TxtSDiscPer.SetFocus
      Case TxtSDiscPer.Name
         If TxtSDiscVal.Enabled Then TxtSDiscVal.SetFocus
      End Select
      KeyCode = 0
      Shift = 0
   ElseIf KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case GridReturn.Name
         GridReturn_DblClick
      Case TxtRCode.Name
         FunSelectRProduct ssValidate, False
            GetDataFromTexBoxesToGridReturn
      Case TxtRQty.Name, TxtRDiscPC.Name, TxtRDiscPer.Name, TxtRPrice.Name, TxtRDiscVal.Name, TxtRSC.Name
             GetDataFromTexBoxesToGridReturn
      Case GridSale.Name
         GridSale_DblClick
      Case TxtSCode.Name
         FunSelectSProduct ssValidate, False
         GetDataFromTexBoxesToGridSale
      Case TxtSCode.Name
         FunSelectSProduct ssValidate, False
         GetDataFromTexBoxesToGridSale
      Case TxtSQty.Name, TxtSDiscPC.Name, TxtSDiscPer.Name, TxtSPrice.Name, TxtSDiscVal.Name, TxtSSC.Name
         GetDataFromTexBoxesToGridSale
      Case TxtPSerial.Name
         If Trim(TxtPSerial.Text) = "" Or (TxtPCode.Text) = "" Then Exit Sub
         TxtRCode.Text = Trim(TxtPCode.Text)
         If FunSelectRProduct(ssValidate, False) = True Then
            GetDataFromTexBoxesToGridReturn
            TxtRSerial.Text = ""
            TxtRSerial.SetFocus
         End If
      Case TxtRSerial.Name
         If Trim(TxtRSerial.Text) = "" Or TxtRCode.Enabled = False Then Exit Sub
         TxtRCode.Text = Trim(TxtRSerial.Text)
         If FunSelectRProduct(ssValidate, False) = True Then
            GetDataFromTexBoxesToGridReturn
            TxtRSerial.Text = ""
            TxtRSerial.SetFocus
         Else
            keybd_event 9, 1, 1, 1
            KeyCode = 0
         End If
      Case TxtSSerial.Name
         If Trim(TxtSSerial.Text) = "" Or TxtSCode.Enabled = False Then Exit Sub
         TxtSCode.Text = Trim(TxtSSerial.Text)
         If FunSelectSProduct(ssValidate, False) = True Then
            GetDataFromTexBoxesToGridSale
            TxtSSerial.Text = ""
            TxtSSerial.SetFocus
         Else
            keybd_event 9, 1, 1, 1
            KeyCode = 0
         End If
       Case TxtPSerial.Name
         If Trim(TxtPSerial.Text) = "" Or TxtPCode.Enabled = False Then Exit Sub
'         TxtPCode.Text = Trim(TxtPSerial.Text)
            If vIsSerial = True Then
               GetDataFromTexBoxesToGridPSerial
               GetDataFromTexBoxesToGridReturn
               TxtPSerial.Text = ""
               TxtPSerial.SetFocus
            End If
         
      Case Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      Select Case ActiveControl.Name
      Case TxtRCode.Name, TxtRQty.Name, TxtRDiscPC.Name, TxtRDiscPer.Name, TxtRPrice.Name, TxtRDiscVal.Name, TxtRSC.Name
         If TxtRCode.Enabled Then TxtRCode.SetFocus: Call SubClearRDetailArea
      Case TxtSCode.Name, TxtSQty.Name, TxtSDiscPC.Name, TxtSDiscPer.Name, TxtSPrice.Name, TxtSDiscVal.Name, TxtSSC.Name
         If TxtSCode.Enabled Then TxtSCode.SetFocus: Call SubClearSDetailArea
      End Select
   ElseIf Shift = vbCtrlMask Then
      Select Case ActiveControl.Name
      Case GridReturn.Name
         If KeyCode = vbKeyDelete Then
            If Trim(GridReturn.Columns("ProductID").Text <> "") Then
               Call mniRemoveRow_Click
               KeyCode = 0
            Else
               KeyCode = 0: Exit Sub
            End If
         End If
      Case GridSale.Name
         If KeyCode = vbKeyDelete Then
            If Trim(GridSale.Columns("ProductID").Text <> "") Then
               Call mniRemoveRow_Click
               KeyCode = 0
            Else
               KeyCode = 0: Exit Sub
            End If
         End If
      End Select
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
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
               KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled And BtnDelete.Visible Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
         Case vbKeyD
            If TxtSCode.Enabled = True And TxtRCode.Enabled = True Then TxtSCode.SetFocus
         Case vbKeyU
            If TxtSCode.Enabled = True And TxtRCode.Enabled = True Then TxtRCode.SetFocus
         Case vbKeyDelete
            Select Case ActiveControl.Name
            Case GridReturn.Name
               If Trim(GridReturn.Columns("ProductID").Text <> "") Then Call mniRemoveRow_Click
            Case GridSale.Name
               If Trim(GridSale.Columns("ProductID").Text <> "") Then Call mniRemoveRow_Click
            End Select
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtRBillID.Name: If FunSelectSale(ssFunctionKey, False) = True Then BtnReturnAll.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtRCode.Enabled Then TxtRCode.SetFocus
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then If TxtEmployeeID.Visible = True Then If TxtEmployeeID.Enabled Then TxtEmployeeID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then If TxtEmployeeID.Visible And TxtEmployeeID.Enabled Then TxtEmployeeID.SetFocus Else TxtOrganizationID.SetFocus
         Case TxtMemberID.Name: If FunSelectMember(ssFunctionKey, True) = True Then If TxtMemberID.Enabled Then TxtMemberID.SetFocus
         Case TxtPCode.Name: If FunSelectPProduct(ssFunctionKey, True) = True Then TxtPSerial.SetFocus
         Case TxtRCode.Name: If FunSelectRProduct(ssFunctionKey, True) = True Then TxtRQty.SetFocus
         Case TxtSCode.Name: If FunSelectSProduct(ssFunctionKey, True) = True Then TxtSQty.SetFocus
         
      End Select
   ElseIf KeyCode = vbKeyF2 Then
         If RFrame.Visible = True Then
            RFrame.Visible = False
            SFrame.Visible = False
            PFrame.Visible = False
'            If ActiveControl.Name = TxtRCode.Name And TxtRCode.Enabled = True Then TxtRCode.SetFocus
'            If ActiveControl.Name = TxtSCode.Name And TxtSCode.Enabled = True Then TxtSCode.SetFocus
        Else
            RFrame.Visible = True
            RFrame.ZOrder 0
            KeyCode = 0
            SFrame.Visible = True
            SFrame.ZOrder 0
            KeyCode = 0
            PFrame.Visible = True
            PFrame.ZOrder 0
            KeyCode = 0
'            If ActiveControl.Name = TxtRCode.Name And TxtRSerial.Enabled = True And TxtRSerial.Visible = True Then TxtRSerial.SetFocus
'            If ActiveControl.Name = TxtSCode.Name And TxtSSerial.Enabled = True And TxtSSerial.Visible = True Then TxtSSerial.SetFocus
        End If
     ElseIf ActiveControl.Name = TxtRCode.Name Then
      If KeyCode = vbKeyDown Then
         GridReturn.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtRCode.Name Then
         KeyCode = 0
         TxtRBillDisc.SetFocus
      End If
   ElseIf ActiveControl.Name = TxtSCode.Name Then
      If KeyCode = vbKeyDown Then
         GridSale.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtSCode.Name Then
         KeyCode = 0
         TxtSBillDisc.SetFocus
      End If
   ElseIf ActiveControl.Name = GridSale.Name And KeyCode = vbKeyF4 Then
      If Trim(GridSale.Columns("ProductID").Text <> "") Then
'         If MniCostPrice.Visible = True Then
'            Call MniCostPrice_Click
'         End If
      End If
   ElseIf KeyCode = vbKeyF5 And ObjUserSecurity.ShowPrice = True Then
      Select Case ActiveControl.Name
      Case TxtRCode.Name, TxtRQty.Name, TxtRPrice.Name, TxtRDiscPC.Name, GridReturn.Name, TxtSCode.Name, TxtSQty.Name, TxtSPrice.Name, TxtSDiscPC.Name, GridSale.Name
         LblRCost.Caption = CN.Execute("select dbo.FunPurPrice('" & TxtRCode.Text & "')").Fields(0).Value
         LblCost.Caption = CN.Execute("select dbo.FunPurPrice('" & TxtSCode.Text & "')").Fields(0).Value
         Call MniCostPrice_Click
'         LblCost.Visible = True
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
   SetWindowText Me.hWnd, "Replacement Invoice"
   HelpLocation Me
   
   Dim vConnString As String
   
   If objFSO.FileExists(vTmp & "\backup.ini") Then
      Open vTmp & "\backup.ini" For Input As #2
      Line Input #2, vConnStr
      Close #2
      vConnString = "Provider=SQLOLEDB.1;User ID=sa;Initial Catalog=" & vConnStr
      If Cnn.State = adStateOpen Then Cnn.Close
      Cnn.Open vConnString
   Else
      vConnStr = ""
   End If
   vColour = ObjRegistry.ShowColourSize
   
   '''''''''''' Return setting
   LblRColour.Visible = vColour
   CmbRColourName.Visible = vColour
   LblRSize.Visible = vColour
   cmbRSizeName.Visible = vColour
   GridReturn.Columns("ColourName").Visible = vColour
   GridReturn.Columns("SizeName").Visible = vColour
   
   '''''''''''' sale setting
   CmbSColourName.Visible = vColour
   cmbSSizeName.Visible = vColour
   GridSale.Columns("ColourName").Visible = vColour
   GridSale.Columns("SizeName").Visible = vColour
   
   If vColour = False Then
       '''''''''''' Return setting
      LblRQty.Left = LblRQty.Left - CmbRColourName.Width - cmbRSizeName.Width
      TxtRQty.Left = TxtRQty.Left - CmbRColourName.Width - cmbRSizeName.Width
      LblRProdPrice.Left = LblRProdPrice.Left - CmbRColourName.Width - cmbRSizeName.Width
      TxtRPrice.Left = TxtRPrice.Left - CmbRColourName.Width - cmbRSizeName.Width
      LblRDiscPC.Left = LblRDiscPC.Left - CmbRColourName.Width - cmbRSizeName.Width
      TxtRDiscPC.Left = TxtRDiscPC.Left - CmbRColourName.Width - cmbRSizeName.Width
      LblRDiscPer.Left = LblRDiscPer.Left - CmbRColourName.Width - cmbRSizeName.Width
      TxtRDiscPer.Left = TxtRDiscPer.Left - CmbRColourName.Width - cmbRSizeName.Width
      LblRDiscVal.Left = LblRDiscVal.Left - CmbRColourName.Width - cmbRSizeName.Width
      TxtRDiscVal.Left = TxtRDiscVal.Left - CmbRColourName.Width - cmbRSizeName.Width
      LblRSC.Left = LblRSC.Left - CmbRColourName.Width - cmbRSizeName.Width
      TxtRSC.Left = TxtRSC.Left - CmbRColourName.Width - cmbRSizeName.Width
      LblRAmount.Left = LblRAmount.Left - CmbRColourName.Width - cmbRSizeName.Width
      TxtRAmount.Left = TxtRAmount.Left - CmbRColourName.Width - cmbRSizeName.Width
      GridReturn.Width = GridReturn.Width - CmbRColourName.Width - cmbRSizeName.Width
      
      '''''''''''' sale setting
      
      TxtSQty.Left = TxtSQty.Left - CmbSColourName.Width - cmbSSizeName.Width
      TxtSPrice.Left = TxtSPrice.Left - CmbSColourName.Width - cmbSSizeName.Width
      TxtSDiscPC.Left = TxtSDiscPC.Left - CmbSColourName.Width - cmbSSizeName.Width
      TxtSDiscPer.Left = TxtSDiscPer.Left - CmbSColourName.Width - cmbSSizeName.Width
      TxtSDiscVal.Left = TxtSDiscVal.Left - CmbSColourName.Width - cmbSSizeName.Width
      TxtSSC.Left = TxtSSC.Left - CmbSColourName.Width - cmbSSizeName.Width
      TxtSAmount.Left = TxtSAmount.Left - CmbSColourName.Width - cmbSSizeName.Width
      GridSale.Width = GridSale.Width - CmbSColourName.Width - cmbSSizeName.Width
   End If

   vServerDate = CN.Execute("Select GetDate()").Fields(0).Value
   vSystemDate = Abs(ObjRegistry.SystemDate)
   vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
   
'   DtpReplaceDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))
'   DtpBillDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))
'   DtpRBillDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))
'   DtpReturnDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))
'
   
   If ObjUserSecurity.ShowStock = True Or ObjUserSecurity.IsAdministrator Then
      vShowStock = True
   Else
      vShowStock = False
   End If
   
   TxtStoreID.Text = IIf((ObjRegistry.StoreID = ""), "", ObjRegistry.StoreID)
   FunSelectStore ssValidate, True
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   LblEmpID.Visible = ObjRegistry.EmpVisible
   LblEmpName.Visible = ObjRegistry.EmpVisible
   TxtEmployeeID.Visible = ObjRegistry.EmpVisible
   TxtEmployeeName.Visible = ObjRegistry.EmpVisible
   BtnEmployee.Visible = ObjRegistry.EmpVisible
         
   TxtManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   LblManualBillNo.Visible = ObjRegistry.ManualBillNoVisible

   vX = IIf(IsNull(ObjRegistry.x), 0, Val(ObjRegistry.x))
   vY = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))

   If ObjUserSecurity.IsAdministrator = False Then
      TxtRDiscPC.Enabled = ObjRegistry.DiscAllowed
      TxtRDiscPer.Enabled = ObjRegistry.DiscAllowed
      TxtRDiscVal.Enabled = ObjRegistry.DiscAllowed
      TxtRBillDisc.Enabled = ObjRegistry.DiscAllowed
      TxtRBillDiscPer.Enabled = ObjRegistry.DiscAllowed
      TxtRServiceCharges.Enabled = ObjRegistry.DiscAllowed
      
      TxtSDiscPC.Enabled = ObjRegistry.DiscAllowed
      TxtSDiscPer.Enabled = ObjRegistry.DiscAllowed
      TxtSDiscVal.Enabled = ObjRegistry.DiscAllowed
      TxtSBillDisc.Enabled = ObjRegistry.DiscAllowed
      TxtSBillDiscPer.Enabled = ObjRegistry.DiscAllowed
      TxtSServiceCharges.Enabled = ObjRegistry.DiscAllowed
      If ObjRegistry.DiscAllowed = False Then
         TxtRDiscPC.Tag = "NC"
         TxtRDiscPer.Tag = "NC"
         TxtRDiscVal.Tag = "NC"
         TxtRBillDisc.Tag = "NC"
         TxtRBillDiscPer.Tag = "NC"
      
         TxtSDiscPC.Tag = "NC"
         TxtSDiscPer.Tag = "NC"
         TxtSDiscVal.Tag = "NC"
         TxtSBillDisc.Tag = "NC"
         TxtSBillDiscPer.Tag = "NC"
      End If
   End If
''         MniCostPrice.Visible = !CostVisible
   If ObjUserSecurity.IsAdministrator = True Then
      TxtSPrice.Enabled = True
      TxtRPrice.Enabled = True
      TxtSPrice.Tag = ""
      TxtRPrice.Tag = ""
   Else
      TxtSPrice.Enabled = ObjUserSecurity.IsChangeRetail
      TxtRPrice.Enabled = ObjUserSecurity.IsChangeRetail
      TxtSPrice.Tag = IIf(TxtSPrice.Enabled = True, "", "D")
      TxtRPrice.Tag = IIf(TxtRPrice.Enabled = True, "", "D")
   End If

   With CN.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
         'vNoofPrints = IIf(IsNull(!NoofPrints) Or !NoofPrints = 0, 1, !NoofPrints)
         If ObjRegistry.ChangePrice = True Then TxtSPrice.Enabled = True
         If ObjRegistry.ChangePrice = True Then TxtRPrice.Enabled = True
      End If
      .Close
   End With
   
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus

   DateFlag = True
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpReplaceDate.IsDateValid = False Then Exit Function
   FunGetMaxID = CN.Execute("Select isnull(max(ReplaceID),0)+1 from ReplacementHeader where ReplaceDate = '" & DtpReplaceDate.DateValue & "'").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Function FunGetMaxIDOut() As Long
   On Error GoTo ErrorHandler
   If DtpBillDate.IsDateValid = False Then FunGetMaxIDOut = 1: Exit Function
   If ObjRegistry.AllowContinuousBillNo = True Then
      FunGetMaxIDOut = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader").Fields(0)
   ElseIf ObjRegistry.AllowMonthlyBillNo = True Then
      FunGetMaxIDOut = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where Month(BillDate) = '" & Month(DtpBillDate.DateValue) & "' and  year(BillDate) ='" & Year(DtpBillDate.DateValue) & "'").Fields(0)
   ElseIf ObjRegistry.AllowDailyBillNo = True Then
      FunGetMaxIDOut = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where BillDate = '" & DtpBillDate.DateValue & "'").Fields(0)
   Else
      FunGetMaxIDOut = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where BillDate = '" & DtpBillDate.DateValue & "' and StoreID = " & TxtStoreID.Text).Fields(0)
   End If
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Function FunGetMaxIDIn() As Long
   On Error GoTo ErrorHandler
   If DtpReturnDate.IsDateValid = False Then Exit Function
   FunGetMaxIDIn = CN.Execute("Select isnull(max(ReturnID),0)+1 from SaleReturnHeader where ReturnDate = '" & DtpReturnDate.DateValue & "'").Fields(0)
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
   
   For Each ctl In FrmReplacePrint.Controls
      If TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then
            ctl.Text = ""
         End If
      End If
      FrmReplacePrint.OptCash.Value = True
   Next
   FrmReplacePrint.TxtCashReceivedBank.Text = ""
   vTotReturnDisc = 0
   vTotalReturn = 0
   TxtRNetAmount.Text = ""
   GridReturn.CancelUpdate
   GridReturn.RemoveAll
   GridReturn.AddNew
   GridReturn.Columns("ProductID").Text = " "
   GridReturn.Update
      
   vTotDisc = 0
   vTotal = 0
   TxtSNetAmount.Text = ""
   GridSale.CancelUpdate
   GridSale.RemoveAll
   GridSale.AddNew
   GridSale.Columns("ProductID").Text = " "
   GridSale.Update
   
   
   Call SubClearSSerialFields
   Call SubClearRSerialFields
'   SFrame.Visible = False
'   RFrame.Visible = False
   
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
      Dim frmObj As Object
      For Each frmObj In Forms
          Set frmObj = Nothing
      Next
      Set RsReturnBody = Nothing
      Set RsReport = Nothing
      Set FrmReplacementInvoice = Nothing
   End If
'   Call DeleteTempActivityLogBin(vRandomID)
      ''''''''''''''''''''''''''' Closed Replacement In Products
      vGridRows = 0
      GridReturn.Redraw = False
      GridReturn.MoveFirst
      For vCounter = 2 To GridReturn.rows
         vGridRows = vGridRows + 1
         If Trim(GridReturn.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From saleReturnbody where SID=" & Val(TxtRSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(GridReturn.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmReplacementInvoice, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Closed Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         GridReturn.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmReplacementInvoice, eCloseSavedRecord, TxtReplaceID.Text, DtpReplaceDate.DateValue, vGridRows & " Replace In Product/s Closed ")
      GridReturn.Redraw = True
      ''''''''''''''''''''''''''''''''''''''''
      
      ''''''''''''''''''''''''''' Closed Replacement Out Products
      vGridRows = 0
      GridSale.Redraw = False
      GridSale.MoveFirst
      For vCounter = 2 To GridSale.rows
         vGridRows = vGridRows + 1
         If Trim(GridSale.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From salebody where SID=" & Val(TxtSSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(GridSale.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmReplacementInvoice, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Closed Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         GridSale.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmReplacementInvoice, eCloseSavedRecord, TxtReplaceID.Text, DtpReplaceDate.DateValue, vGridRows & " Replace Out Product/s Closed ")
      GridSale.Redraw = True
      ''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridReturn_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotReturnQty.Text = Val(TxtTotReturnQty.Text) - GridReturn.Columns("Qty").Value
   vTotReturnDisc = vTotReturnDisc - Val(GridReturn.Columns("DiscVal").Text)
   vTotalReturn = vTotalReturn - GridReturn.Columns("TotalAmount").Value
   SubCalculateRFooter
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridReturn_DblClick()
   Call GridReturn_LostFocus
End Sub

Private Sub GridReturn_GotFocus()
   Flag = True
   TxtRCode.Enabled = False
   BtnRProduct.Enabled = False
   'TxtRCode.BackColor = TxtRProductName.BackColor
   'TxtRCode.TabStop = False
End Sub

Private Sub GridReturn_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub GridReturn_LostFocus()
   Flag = False
   If Trim(GridReturn.Columns("ProductID").Text) = "" Then
      TxtRCode.Text = ""
      TxtRCode.Enabled = True
      BtnRProduct.Enabled = True
      TxtRCode.SetFocus
   Else
      TxtRCode.Enabled = False
      BtnRProduct.Enabled = False
      If TxtRQty.Enabled = True And TxtRQty.Visible Then TxtRQty.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
End Sub

Private Sub GridReturn_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(GridReturn.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub GridReturn_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridReturnToTexBoxes
   Call PopulateDataToGridRSerial
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If ActiveControl.Name = GridReturn.Name Then
      If Trim(GridReturn.Columns("Code").Text) = "" Then Exit Sub
       ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtRSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(GridReturn.Columns("Code").Text)
      With CN.Execute(ssql)
         If .EOF Then
            Call ActivityLogBin("", eFrmReplacementInvoice, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Removed Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
         Else
            Call ActivityLogBin("", eFrmReplacementInvoice, eRemoveRow, TxtReplaceID.Text, DtpReplaceDate.DateValue, "Removed Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
            Call ActivityLogBin(vRandomID, eFrmReplacementInvoice, eAddTempRecord, TxtReplaceID.Text, DtpReplaceDate.DateValue, "Pending Remove Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
         End If
      End With
      RsReturnBody.Filter = "Code='" & TxtRCode.Text & "'"
      If RsReturnBody.RecordCount > 0 Then RsReturnBody.Delete
'      cn.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','In Removed Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
      GridReturn.SelBookmarks.RemoveAll
      GridReturn.SelBookmarks.Add GridReturn.Bookmark
      GridReturn.DeleteSelected
      GridReturn.Refresh
      RsReturnBody.Filter = 0
      GridReturn.MoveLast
      GetDataBackFromGridReturnToTexBoxes
      
      RsReturnSerial.Filter = "ProductID = " & Val(TxtRCode.Text)
      While Not RsReturnSerial.EOF
         RsPurchaseSerial.Filter = "Serial = " & RsReturnSerial!Serial
         If RsPurchaseSerial.RecordCount > 0 Then
            RsPurchaseSerial!SerialAdd = 1
            RsPurchaseSerial.Update
         End If
         RsReturnSerial.Delete
         RsReturnSerial.MoveNext
      Wend
      
      Call SubClearRSerialFields
   ElseIf ActiveControl.Name = GridSale.Name Then
      If Trim(GridSale.Columns("Code").Text) = "" Then Exit Sub
      ssql = "Select Productid From salebody where sid=" & Val(TxtSSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(GridSale.Columns("Code").Text)
      With CN.Execute(ssql)
         If .EOF Then
            Call ActivityLogBin("", eFrmReplacementInvoice, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Removed Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
         Else
            Call ActivityLogBin("", eFrmReplacementInvoice, eRemoveRow, TxtReplaceID.Text, DtpReplaceDate.DateValue, "Removed Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
            Call ActivityLogBin(vRandomID, eFrmReplacementInvoice, eAddTempRecord, TxtReplaceID.Text, DtpReplaceDate.DateValue, "Pending Remove Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
         End If
      End With
      RsSaleBody.Filter = "Code='" & TxtSCode.Text & "'"
      If RsSaleBody.RecordCount > 0 Then RsSaleBody.Delete
'      cn.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Out Removed Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
      GridSale.SelBookmarks.RemoveAll
      GridSale.SelBookmarks.Add GridSale.Bookmark
      GridSale.DeleteSelected
      GridSale.Refresh
      RsSaleBody.Filter = 0
      GridSale.MoveLast
      GetDataBackFromGridSaleToTexBoxes
      
      RsBodySerial.Filter = "ProductID = " & Val(TxtSCode.Text)
      While Not RsBodySerial.EOF
         RsPurchaseSerial.Filter = "Serial = " & RsBodySerial!Serial
         If RsPurchaseSerial.RecordCount > 0 Then
            RsPurchaseSerial!SerialAdd = 1
            RsPurchaseSerial.Update
         End If
         RsBodySerial.Delete
      Wend
      Call SubClearRSerialFields
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGridReturn()
   Dim vrowcounter As Integer
   If Trim(TxtRCode.Text) = "" Then
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
      TxtRCode.SetFocus
      Exit Sub
   End If
   If Val(TxtRQty.Text) = 0 Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      TxtRQty.SetFocus
      Exit Sub
   End If
   If Round(Val(TxtRDiscPer.Text), 2) <> Round((Val(TxtRDiscPC.Text) * 100) / (IIf(Val(TxtRPrice.Text) = 0, 1, Val(TxtRPrice.Text))), 2) Then
      MsgBox "Please update the Discount for change Price.", vbExclamation, "Alert"
      If TxtRDiscPer.Enabled And TxtRDiscPer.Visible Then TxtRDiscPer.SetFocus
      Exit Sub
   End If
   
    If (CmbRColourName.Text = "" Or cmbRSizeName.Text = "") And vColour = True Then
      MsgBox "Please Select Colour and Size", vbInformation + vbOKOnly, "Error"
      Exit Sub
   End If
   '''''''''   check Serial
   RsReturnSerial.Filter = "ProductID =" & Val(TxtRCode.Text)
   If (TxtRCode.Enabled = False And RsReturnSerial.RecordCount <> 0) And RsReturnSerial.RecordCount <> TxtRQty.Text Then
      MsgBox "Qty Should be equal to Serial", vbInformation + vbOKOnly, "Error"
      Call SubClearRSerialFields
      If TxtRCode.Enabled And TxtRCode.Visible Then TxtRCode.SetFocus
      Exit Sub
   End If
   RsBodySerial.Filter = ""
''''''''
   Call SubClearRSerialFields
On Error GoTo ErrorHandler
   RsReturnBody.Filter = "ProductID = " & Val(TxtPID.Text)
   If TxtRCode.Enabled Then
      If RsReturnBody.RecordCount = 0 Then
         RsReturnBody.AddNew
         GridReturn.Columns("ProductID").Text = TxtPID.Text
         GridReturn.Columns("Code").Text = TxtRCode.Text
         RsReturnBody!Productid = TxtPID.Text
         RsReturnBody!Code = TxtRCode.Text
         RsReturnBody!StoreID = TxtStoreID.Text
      Else
         GridReturn.Redraw = False
         GridReturn.MoveFirst
            For vrowcounter = 1 To GridReturn.rows
               If GridReturn.Columns("Productid").Text = TxtPID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearRDetailArea
                  ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtRSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(GridReturn.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmReplacementInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Effected Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmReplacementInvoice, eEdit, TxtReplaceID.Text, DtpReplaceDate.DateValue, "Effected Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
                     End If
                  End With
                  
                  TxtRQty.Text = Val(TxtRQty.Text) + GridReturn.Columns("Qty").Value
                  TxtTotReturnQty.Text = Val(TxtTotReturnQty.Text) + Val(TxtRQty.Text) - GridReturn.Columns("Qty").Value
                  vTotReturnDisc = vTotReturnDisc + Val(TxtRDiscVal.Text) - Val(GridReturn.Columns("DiscVal").Text)
                  vTotalReturn = vTotalReturn + Val(TxtActualAmount.Text) - GridReturn.Columns("TotalAmount").Value
                  GridReturn.Columns("ProductName").Text = TxtRProductName.Text
                  GridReturn.Columns("IsSerial").Value = vIsSerial
                  GridReturn.Columns("Qty").Value = Val(TxtRQty.Text)
                  GridReturn.Columns("Price").Value = Val(TxtRPrice.Text)
                  GridReturn.Columns("DiscPC").Value = Val(TxtRDiscPC.Text)
                  GridReturn.Columns("DiscPer").Value = Val(TxtRDiscPer.Text)
                  GridReturn.Columns("DiscVal").Value = Val(TxtRDiscVal.Text)
                  GridReturn.Columns("SC").Value = Val(TxtRSC.Text)
                  GridReturn.Columns("Cost").Value = Val(TxtCost.Text)
                  GridReturn.Columns("EmpComm").Value = Val(TxtEmpComm.Text)
                  GridReturn.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
                  GridReturn.Columns("Amount").Value = Val(TxtRAmount.Text)
                  GridReturn.Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
                  RsReturnBody!Qty = Val(TxtRQty.Text)
                  RsReturnBody!Price = Val(TxtRPrice.Text)
                  RsReturnBody!DiscPC = Val(TxtRDiscPC.Text)
                  RsReturnBody!DiscPer = Val(TxtRDiscPer.Text)
                  RsReturnBody!DiscVal = Val(TxtRDiscVal.Text)
                  RsReturnBody!SC = IIf(Val(TxtRSC.Text) = 0, Null, Val(TxtRSC.Text))
                  RsReturnBody!Cost = Val(TxtCost.Text)
                  RsReturnBody!EmpComm = Val(TxtEmpComm.Text)
                  RsReturnBody!isProduct = Abs(ChkIsProduct.Value)
                  RsReturnBody!Amount = Val(TxtRAmount.Text)
                  ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtRSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(GridReturn.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmReplacementInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Updated Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmReplacementInvoice, eEdit, TxtReplaceID.Text, DtpReplaceDate.DateValue, "Updated Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmReplacementInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Pending Update Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
                  GridReturn.MoveLast
                  Call SubClearRDetailArea
                  TxtRCode.SetFocus
                  GridReturn.Redraw = True
                  Exit Sub
               End If
               GridReturn.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearRDetailArea
         GridReturn.MoveLast
         TxtRCode.SetFocus
         Exit Sub
      End If
   End If
   GridReturn.Redraw = False
   With GridReturn
      If TxtRCode.Enabled = True Then
         TxtTotReturnQty.Text = Val(TxtTotReturnQty.Text) + Val(TxtRQty.Text)
         vTotReturnDisc = vTotReturnDisc + Val(TxtRDiscVal.Text)
         vTotalReturn = vTotalReturn + Val(TxtActualAmount.Text)
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmReplacementInvoice, eAddNewRowByEdit, TxtReplaceID.Text, DtpReplaceDate.DateValue, "Add New Replace In Code-" & TxtRCode.Text & " Qty-" & TxtRQty.Text & " Price-" & TxtRPrice.Text & " Disc-" & TxtRDiscPer.Text & " Amount-" & TxtRAmount.Text)
         Call ActivityLogBin(vRandomID, eFrmReplacementInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Pending Add New Replace In Code-" & TxtRCode.Text & " Qty-" & TxtRQty.Text & " Price-" & TxtRPrice.Text & " Disc-" & TxtRDiscPer.Text & " Amount-" & TxtRAmount.Text)
      Else
         TxtTotReturnQty.Text = Val(TxtTotReturnQty.Text) + Val(TxtRQty.Text) - GridReturn.Columns("Qty").Value
         vTotReturnDisc = vTotReturnDisc + Val(TxtRDiscVal.Text) - Val(GridReturn.Columns("DiscVal").Text)
         vTotalReturn = vTotalReturn + Val(TxtActualAmount.Text) - GridReturn.Columns("TotalAmount").Value
         ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtRSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(GridReturn.Columns("Code").Text)
         With CN.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmReplacementInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Effected Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmReplacementInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Updated Replace In Code-" & TxtRCode.Text & " Qty-" & TxtRQty.Text & " Price-" & TxtRPrice.Text & " Disc-" & TxtRDiscPer.Text & " Amount-" & TxtRAmount.Text)
            Else
               Call ActivityLogBin("", eFrmReplacementInvoice, eEdit, TxtReplaceID.Text, DtpReplaceDate.Date, "Effected Replace In Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmReplacementInvoice, eEdit, TxtReplaceID.Text, DtpReplaceDate.Date, "Updated Replace In Code-" & TxtSCode.Text & " Qty-" & TxtSQty.Text & " Price-" & TxtSPrice.Text & " Disc-" & Val(TxtSDiscPer.Text) & " Amount-" & TxtSAmount.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmReplacementInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Pending Update Replace In Code-" & TxtRCode.Text & " Qty-" & TxtRQty.Text & " Price-" & TxtRPrice.Text & " Disc-" & TxtRDiscPer.Text & " Amount-" & TxtRAmount.Text)
      End If
      .Columns("ProductName").Text = TxtRProductName.Text
      
      GridReturn.Columns("ColourName").Text = CmbRColourName.Text
      If CmbRColourName.Text <> "" Then GridReturn.Columns("ColourID").Value = CmbRColourName.ItemData(CmbRColourName.ListIndex)
      GridReturn.Columns("SizeName").Text = cmbRSizeName.Text
      If cmbRSizeName.Text <> "" Then GridReturn.Columns("SizeID").Value = cmbRSizeName.ItemData(cmbRSizeName.ListIndex)
      
      If vColour = True And GridReturn.Columns("ColourID").Text <> "" Then
         RsReturnBody!ColourID = GridReturn.Columns("ColourID").Text
         RsReturnBody!SizeID = GridReturn.Columns("SizeID").Text
      End If
            
      .Columns("Qty").Value = Val(TxtRQty.Text)
      .Columns("Price").Value = Val(TxtRPrice.Text)
      .Columns("DiscPC").Value = Val(TxtRDiscPC.Text)
      .Columns("DiscPer").Value = Val(TxtRDiscPer.Text)
      .Columns("DiscVal").Value = Val(TxtRDiscVal.Text)
      .Columns("SC").Value = Val(TxtRSC.Text)
      If Trim(TxtCost.Text) <> "" Then
         .Columns("Cost").Value = Val(TxtCost.Text)
      End If
      .Columns("EmpComm").Value = Val(TxtEmpComm.Text)
      .Columns("IsSerial").Value = Abs(vIsSerial)
      .Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
      .Columns("Amount").Value = Val(TxtRAmount.Text)
      .Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
      RsReturnBody!isSerial = vIsSerial
      RsReturnBody!Qty = Val(TxtRQty.Text)
      RsReturnBody!Price = Val(TxtRPrice.Text)
      RsReturnBody!DiscPC = Val(TxtRDiscPC.Text)
      RsReturnBody!DiscPer = Val(TxtRDiscPer.Text)
      RsReturnBody!DiscVal = Val(TxtRDiscVal.Text)
      RsReturnBody!SC = IIf(Val(TxtRSC.Text) = 0, Null, Val(TxtRSC.Text))
      If Trim(TxtCost.Text) <> "" Then
         RsReturnBody!Cost = Val(TxtCost.Text)
      End If
      RsReturnBody!EmpComm = Val(TxtEmpComm.Text)
      If IsNull(RsReturnBody!Cost) Then RsReturnBody!Cost = 0
      RsReturnBody!Amount = Val(TxtRAmount.Text)
      RsReturnBody!isProduct = Abs(ChkIsProduct.Value)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearRDetailArea
   TxtRCode.SetFocus
   GridReturn.Redraw = True
   Exit Sub
ErrorHandler:
   GridReturn.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SubClearRDetailArea()
   CmbRColourName.Clear
   cmbRSizeName.Clear
   TxtRCode.Enabled = True
   BtnRProduct.Enabled = True
   TxtRCode.Text = ""
   TxtRProductName.Text = ""
   TxtRQty.Text = ""
   TxtRPrice.Text = ""
   TxtRDiscPC.Text = ""
   TxtRDiscPer.Text = ""
   TxtRDiscVal.Text = ""
   TxtRSC.Text = ""
   TxtEmpComm.Text = ""
   TxtRAmount.Text = ""
   TxtActualAmount.Text = ""
   ChkIsProduct.Value = 1
   TxtCost.Text = ""
End Sub

Private Sub GetDataBackFromGridReturnToTexBoxes()
   On Error GoTo ErrorHandler
   With GridReturn
      TxtPID.Text = .Columns("ProductID").Text
      TxtRCode.Text = .Columns("code").Text
      TxtRProductName.Text = .Columns("ProductName").Text
      TxtRQty.Text = .Columns("Qty").Text
      TxtRPrice.Text = .Columns("Price").Text
      TxtRDiscPC.Text = .Columns("DiscPC").Value
      TxtRDiscPer.Text = .Columns("DiscPer").Value
      TxtRDiscVal.Text = .Columns("DiscVal").Value
      TxtCost.Text = .Columns("Cost").Value
      TxtEmpComm.Text = .Columns("EmpComm").Value
      TxtRAmount.Text = .Columns("Amount").Text
      TxtActualAmount.Text = .Columns("TotalAmount").Text
      ChkIsProduct.Value = Abs(.Columns("IsProduct").Value)
      vIsSerial = Val(.Columns("IsSerial").Value)
      If Trim(.Columns("ColourName").Text) <> "" Then
         CmbRColourName.AddItem .Columns("ColourName").Text
         CmbRColourName.ItemData(CmbRColourName.NewIndex) = .Columns("ColourID").Text
         CmbRColourName.ListIndex = 0
      End If
      
      If Trim(.Columns("SizeName").Text) <> "" Then
         cmbRSizeName.AddItem .Columns("ColourName").Text
         cmbRSizeName.ItemData(cmbRSizeName.NewIndex) = .Columns("SizeID").Text
         cmbRSizeName.ListIndex = 0
      End If
      
'      With CN.Execute("select QtyLoose from currentstockStore where productid ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text)
'         If .RecordCount > 0 Then
'            vQtyLoose = !QtyLoose
'            LblStock.Caption = !QtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'         Else
'            vQtyLoose = 0
'            LblStock.Caption = 0
'         End If
'      End With
   End With
   
'        With CN.Execute("select QtyLoose from currentstockStore where productid ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text)
'         If .RecordCount > 0 Then
'            vQtyLoose = !QtyLoose
'            LblStock.Caption = !QtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'         Else
'            vQtyLoose = 0
'            LblStock.Caption = 0
'         End If
'      End With
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpReplaceDate.DateValue + 1 & "',0),0)"
         vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = vQtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit(" & Val(TxtPID.Text) & ")").Fields(0).Value
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         
   
   If GridReturn.rows = 1 Then GridReturn.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtPCode_Change()
 If ActiveControl.Name <> TxtPCode.Name Then Exit Sub
   If TxtPProductName.Text <> "" Then
      TxtPCode.Text = ""
      TxtPID.Text = ""
      TxtPProductName.Text = ""
      TxtRPrice.Text = ""
      TxtRDiscPC.Text = ""
   End If
End Sub

Private Sub TxtPCode_Validate(Cancel As Boolean)
 On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtPCode.Text) = "" Then Exit Sub
   vTemp = Not FunSelectPProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectPProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtRBillDisc_Change()
   If ActiveControl.Name <> TxtRBillDisc.Name Then Exit Sub
   TxtRBillDiscPer.Text = Round((Val(TxtRBillDisc.Text) * 100) / Val(vTotalReturn - vTotReturnDisc), 2)
   Call SubCalculateRFooter
End Sub

Private Sub TxtRBillDiscPer_Change()
   If ActiveControl.Name <> TxtRBillDiscPer.Name Then Exit Sub
   TxtRBillDisc.Text = SelfRound((Val(vTotalReturn - vTotReturnDisc) * Val(TxtRBillDiscPer.Text) / 100))
   Call SubCalculateRFooter
End Sub

Private Sub TxtRBillID_Change()
   If Trim(TxtRBillID.Text) = "" Then DtpRBillDate.Enabled = False Else DtpRBillDate.Enabled = True
End Sub

Private Sub TxtRDiscPC_Change()
   If ActiveControl.Name <> TxtRDiscPC.Name Then Exit Sub
   If Val(TxtRPrice.Text) = 0 Then Exit Sub
   TxtRDiscPer.Text = Round((Val(TxtRDiscPC.Text) * 100) / Val(TxtRPrice.Text), 2)
   Call SubCalculateRBody
End Sub

'Private Sub TxtRDiscPC_LostFocus()
'   Select Case Me.ActiveControl.Name
'   Case TxtRCode.Name, TxtRQty.Name, TxtRDiscPC.Name
'      Exit Sub
'   End Select
'   Call GetDataFromTexBoxesToGridReturn
'End Sub

Private Sub TxtRDiscPer_Change()
   If ActiveControl.Name <> TxtRDiscPer.Name Then Exit Sub
   TxtRDiscPC.Text = Round((Val(TxtRPrice.Text) * Val(TxtRDiscPer.Text) / 100), 2)
   Call SubCalculateRBody
End Sub

Private Sub TxtRCode_Change()
   If ActiveControl.Name <> TxtRCode.Name Then Exit Sub
   If TxtRProductName.Text <> "" Then
      TxtRCode.Text = ""
      TxtPID.Text = ""
      TxtRProductName.Text = ""
      TxtRPrice.Text = ""
      TxtRDiscPC.Text = ""
   End If
End Sub

Private Sub TxtRCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridReturn.SetFocus
End Sub

'Private Sub TxtRCode_LostFocus()
'   If Len(TxtRCode.Text) > 7 Then
'      GetDataFromTexBoxesToGridReturn
'   End If
'End Sub

Private Sub TxtRCode_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtRCode.Text) = "" Then Exit Sub
   vTemp = Not FunSelectRProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectRProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtRDiscVal_Change()
   If TxtRDiscVal.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtRDiscVal.Name Then Exit Sub
   If Val(TxtRPrice.Text) = 0 Then Exit Sub
   If Val(TxtRQty.Text) = 0 Then Exit Sub
   TxtActualAmount.Text = Val(TxtRQty.Text) * (Val(TxtRPrice.Text) + Val(TxtRSC.Text))
   TxtRDiscPC.Text = Round(Val(TxtRDiscVal.Text) / (TxtRQty.Text), 3)
   TxtRDiscPer.Text = Round((Val(TxtRDiscPC.Text) * 100) / Val(TxtRPrice.Text), 2)
   TxtRAmount.Text = Val(TxtActualAmount.Text) - Val(TxtRDiscVal.Text)
   SubCalculateRFooter
End Sub

Private Sub TxtRNetAmount_Change()
   Call SubPaidReceived
End Sub

Private Sub TxtRPrice_Change()
   Call SubCalculateRBody
End Sub

Private Sub TxtRQty_Change()
   Call SubCalculateRBody
   Call FindRRebate
End Sub

Private Sub TxtRSC_Change()
 On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtRSC.Name Then Exit Sub
   Call SubCalculateRBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtRServiceCharges_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtRServiceCharges.Name Then Exit Sub
   Call SubCalculateRFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSNetAmount_Change()
   Call SubPaidReceived
End Sub

Private Sub TxtSSC_Change()
 On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSSC.Name Then Exit Sub
   Call SubCalculateSBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSServiceCharges_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSServiceCharges.Name Then Exit Sub
   Call SubCalculateSFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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

Private Sub SubCalculateSBody()
    TxtSDiscVal.Text = Val(TxtSQty.Text) * Val(TxtSDiscPC.Text)
    TxtActualAmount.Text = Val(TxtSQty.Text) * (Val(TxtSPrice.Text) + Val(TxtSSC.Text))
    TxtSAmount.Text = Val(TxtActualAmount.Text) - Val(TxtSDiscVal.Text)
''    TxtTotalDiscount.Caption = vTotDisc
    SubCalculateSFooter
End Sub

Private Sub SubCalculateSFooter()
   TxtSNetAmount.Text = SelfRound(Val(vTotal - vTotDisc) - Val(TxtSBillDisc.Text) + Val(TxtSServiceCharges.Text))
End Sub

Private Sub FindSRebate()
   Dim Rebate
   On Error GoTo ErrorHandler
    With CN.Execute("Select * from ProductOffers where Rebate <> 0 and ProductID = " & Val(TxtPID.Text))
        If .RecordCount > 0 Then
            Rebate = Val(TxtSQty.Text)
            Rebate = Rebate \ !Qty
            Rebate = Rebate * !Rebate
            TxtSDiscVal.Text = Rebate
            If Val(TxtSPrice.Text) = 0 Then Exit Sub
            If Val(TxtSQty.Text) = 0 Then Exit Sub
            TxtSDiscPC.Text = Round(Val(TxtSDiscVal.Text) / (TxtSQty.Text), 3)
            TxtSDiscPer.Text = Round((Val(TxtSDiscPC.Text) * 100) / Val(TxtSPrice.Text), 2)
            TxtActualAmount.Text = Val(TxtSQty.Text) * Val(TxtSPrice.Text)
            TxtSAmount.Text = Val(TxtActualAmount.Text) - Val(TxtSDiscVal.Text)
'            vTotReturnDisc = vTotReturnDisc + Val(TxtRDiscVal.Text)
'            vTotalReturn = vTotalReturn + TxtRAmount.Text
            SubCalculateSFooter
        End If
    End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      If vColour = True Then
         SchItemCode.ParaInWhere = " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ") "
         SchItemCode.Show vbModal, Me
         
         TxtSCode.Text = SchItemCode.ParaOutItemCode
      Else
         SchProduct.ParaInWhere = " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ") "
         SchProduct.ParainShowStock = vShowStock
         SchProduct.Show vbModal, Me
         TxtSCode.Text = SchProduct.ParaOutID
      End If
   End If
    '---------------------------
   If TxtSCode.Enabled = False Then FunSelectSProduct = False: Exit Function
   If Trim(TxtSCode.Text) = "" Then FunSelectSProduct = False: Exit Function
   If TxtSCode.Text = "" Then FunSelectSProduct = False: Exit Function
   
   ''''''''''''' Serail '''''''''''''''''''''''''''''''''
   vSerialAdd = False
   vStrSQL = "Select ProductID, Serial, SerialAdd from vuPurchaseSerial where Serial = '" & Trim(TxtSCode.Text) & "' or ProductID = " & Val(TxtSCode.Text)
   With CN.Execute(vStrSQL)
      If .EOF = False Then
            If SFrame.Visible = False Then
               SFrame.Visible = True
               SFrame.ZOrder 0
            End If
            TxtSSerial.Text = TxtSCode.Text
            TxtSCode.Text = !Productid
            GetDataFromTexBoxesToGridSSerial
            If vSerialAdd = False Then
               TxtSCode.Text = ""
               FunSelectSProduct = False
               Exit Function
            End If
      End If
   End With
 '''''''''''''''''''''''''''''''''''''''''''''

   
    If vColour = True Then
      ssql = "select c.ColourID, ColourName from productcolours pc inner join Colours c on pc.colourid = c.colourid " & vbCrLf _
             & "inner join products p on p.productid = pc.productid " & vbCrLf _
             & "where ItemCode = '" & IIf(Len(TxtSCode.Text) = 9, TxtSCode.Text & "'", Mid(TxtSCode.Text, 1, 9) & "' and c.colourid = " & Val(Mid(TxtSCode.Text, 10, 2)))
      With CN.Execute(ssql)
         If .RecordCount > 0 Then
            CmbSColourName.AddItem !ColourName
            CmbSColourName.ItemData(CmbSColourName.NewIndex) = !ColourID
            CmbSColourName.ListIndex = 0
         End If
      End With
      
      ssql = "select s.SizeID, SizeName from productSizes pz inner join Sizes s on pz.Sizeid = s.Sizeid " & vbCrLf _
      & "inner join products p on p.productid = pz.productid " & vbCrLf _
      & "where ItemCode = '" & IIf(Len(TxtSCode.Text) = 13, Mid(TxtSCode.Text, 1, 9) & "' and s.sizeid = " & Val(Mid(TxtSCode.Text, 12, 2)), TxtSCode.Text & "'")
      With CN.Execute(ssql)
         If .RecordCount > 0 Then
            cmbSSizeName.AddItem !SizeName
            cmbSSizeName.ItemData(cmbSSizeName.NewIndex) = !SizeID
            cmbSSizeName.ListIndex = 0
         End If
      End With
      TxtSCode.Text = CStr(Left(TxtSCode.Text, 9))
   End If
   
    ''''''''***********   Checking Union   ***********''''''''
   vStrSQL = " SELECT p.productid, Code, ProductName, ServiceCharges, RetailPrice, DiscPer, DiscPC, EmpComm" & vbCrLf _
         + " from PackageDealInfoHeader un inner join Products p on un.PackageDealID = p.productid" & vbCrLf _
         + " left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
         + " where ( " & IIf(IsNumeric(TxtSCode.Text) = False, "", "p.productid = " & (TxtSCode.Text) & " or ") & " code = '" & TxtSCode.Text & "')" & " and isLocked = 0 "
         
         
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtSProductName.Text = !ProductName
         TxtSPrice.Text = !RetailPrice
         TxtEmpComm.Text = !EmpComm
         TxtSQty.Text = IIf(Val(TxtSQty.Text) = 0, 1, TxtSQty.Text)
         vStrSQL = " select sum(isnull(Cost,PurPrice)* b.QtyLoose) as Cost from PackageDealInfoHeader h inner join PackageDealInfoBody b on h.id = b.id" & vbCrLf _
               + " inner join Products p on p.productid = b.productid" & vbCrLf _
               + " left outer join CurrentStock cs on cs.productid = p.productid " & vbCrLf _
               + " where h.PackageDealID ='" & TxtPID.Text & "'"
         With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
                 
         vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & Val(TxtPID.Text)
         With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               vQtyLoose = .Fields(0).Value
            Else
               vQtyLoose = 0
            End If
         End With
         LblStock.Caption = vQtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit(" & Val(TxtPID.Text) & ")").Fields(0).Value
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         
         
'         VStrSQL = " select Floor(min(css.QtyLoose/b.QtyLoose)) as QtyLoose " & vbCrLf _
'                  + " from PackageDealInfoHeader h inner join PackageDealInfoBody b on h.id = b.id" & vbCrLf _
'                  + " inner join Products p on p.productid = b.productid" & vbCrLf _
'                  + " left outer join CurrentStockStore css on css.productid = p.productid " & vbCrLf _
'                  + " where h.PackageDealID ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text
'         With CN.Execute(VStrSQL)
'            If .RecordCount > 0 Then
'               vQtyLoose = !QtyLoose
'               LblStock.Caption = IIf(IsNull(!QtyLoose), 0, !QtyLoose) & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'            Else
'               vQtyLoose = 0
'               LblStock.Caption = 0
'            End If
'         End With
'         LblStock.Visible = True
'         LblStockCaption.Visible = True
         
   
'               If !LastRateVisible = True Then
'                  If FrmReplacePrint.TxtCustomerID.Text <> "" Then
'                     LblPrice = CN.Execute("Select dbo.FunLastPrice('S','" & DtpBillDate.DateValue & "','" & TxtPID.Text & "','" & FrmReplacePrint.TxtCustomerID.Text & "')").Fields(0).Value
'                     LblCaptionPrice.Visible = True
'                     LblPrice.Visible = True
'                  End If
'               End If
         TxtRSC.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtSDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtSDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         TxtEmpComm.Text = IIf(IsNull(!EmpComm), 0, !EmpComm)
         If Val(TxtSDiscPC.Text) <> 0 Then
            TxtSDiscPer.Text = Round((Val(TxtSDiscPC.Text) * 100) / Val(TxtSPrice.Text), 2)
         End If
'         ChkIsProduct.Value = 0
         SubCalculateSBody
'         Char.Speak TxtSProductName.Text
         FunSelectSProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      End If
   End With
    
''''''''***********   Checking Product  ***********''''''''
    vStrSQL = " SELECT p.productid, Qty, code, ProductName, ServiceCharges, RetailPrice, DiscPer, DiscPC, EmpComm" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtSCode.Text) = False, "", "p.productid = " & (TxtSCode.Text) & " or ") & " code = '" & TxtSCode.Text & "')" & " and isLocked = 0 "

   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtSProductName.Text = !ProductName
         TxtSPrice.Text = !RetailPrice
         TxtEmpComm.Text = IIf(IsNull(!EmpComm), 0, !EmpComm)
         TxtSQty.Text = IIf(Len(TxtSCode.Text) <= 5 And IsNumeric(TxtSCode.Text), 1, IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty)) ' IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty)
         With CN.Execute("select cost from currentstock where productid = " & Val(TxtPID.Text))
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
         
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & Val(TxtPID.Text)
            With CN.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpReplaceDate.DateValue + 1 & "',0),0)"
            vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = vQtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit(" & Val(TxtPID.Text) & ")").Fields(0).Value
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         
         If ObjRegistry.NegativeSale = False Then
            If vQtyLoose <= 0 Then
               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
               FunSelectSProduct = False
               Exit Function
            End If
         End If
'         With CN.Execute("select qtyloose from currentstockStore where productid ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text)
'            If .RecordCount > 0 Then
'               vQtyLoose = !QtyLoose
'               LblStock.Caption = !QtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'            Else
'               vQtyLoose = 0
'               LblStock.Caption = 0
'            End If
'         End With
'         LblStock.Visible = True
'         LblStockCaption.Visible = True
'         If ObjRegistry.NegativeSale = False Then
'            If vQtyLoose <= 0 Then
'                MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                FunSelectSProduct = False
'                Exit Function
'            End If
'         End If

'            If !LastRateVisible = True Then
'               If FrmReplacePrint.TxtCustomerID.Text <> "" Then
'                  LblPrice = CN.Execute("Select dbo.FunLastPrice('S','" & DtpBillDate.DateValue & "','" & TxtPID.Text & "','" & FrmReplacePrint.TxtCustomerID.Text & "')").Fields(0).Value
'                  LblCaptionPrice.Visible = True
'                  LblPrice.Visible = True
'               End If
'            End If
         TxtSSC.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtSDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtSDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtSDiscPC.Text) <> 0 Then
            TxtSDiscPer.Text = Round((Val(TxtSDiscPC.Text) * 100) / Val(TxtSPrice.Text), 2)
         End If
         ChkIsProduct.Value = 1
         If Val(TxtSQty.Text) > 1 Then FindSRebate
         SubCalculateSBody
'         Char.Speak TxtSProductName.Text
         FunSelectSProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectSProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtPID.Text = ""
         TxtSCode.Text = ""
         TxtSProductName.Text = ""
         TxtSPrice.Text = ""
         TxtSDiscPC.Text = ""
         TxtSDiscPer.Text = ""
         TxtSSC.Text = ""
         TxtSAmount.Text = ""
         TxtEmpComm.Text = ""
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

Private Sub BtnSProduct_Click()
   If FunSelectSProduct(ssButton, True) = True Then
      TxtSQty.SetFocus
   Else
      TxtSCode.SetFocus
   End If
End Sub

Private Sub PopulateDataToGridSale()
   RsSaleBody.Filter = 0
   If RsSaleBody.State = adStateOpen Then RsSaleBody.Close
   RsSaleBody.Open "Select * from SaleBody where SID = " & Val(TxtSSID.Text), CN, adOpenDynamic, adLockBatchOptimistic
   If RsSaleBody.RecordCount > 0 Then
      ssql = "select p.productname, b.code, b.ColourID, b.SizeID, ColourName, SizeName, b.* from salebody b join products p on p.productid = b.productid Left outer join Colours Col on Col.Colourid = b.ColourID Left Outer join Sizes Sz on Sz.SizeID = b.SizeID where SID = " & Val(TxtSSID.Text)
      With CN.Execute(ssql)
         GridSale.Redraw = False
         GridSale.MoveFirst
         GridSale.RemoveAll
         GridSale.AllowAddNew = True
         vTotal = 0
         vTotDisc = 0
         While Not .EOF
            GridSale.AddNew
            GridSale.Columns("ProductID").Text = !Productid
            GridSale.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            GridSale.Columns("ProductName").Text = !ProductName
            
            GridSale.Columns("ColourID").Value = IIf(IsNull(!ColourID), "", !ColourID)
            GridSale.Columns("ColourName").Value = IIf(IsNull(!ColourName), "", !ColourName)
            GridSale.Columns("SizeID").Value = IIf(IsNull(!SizeID), "", !SizeID)
            GridSale.Columns("SizeName").Value = IIf(IsNull(!SizeName), "", !SizeName)
            GridSale.Columns("StoreID").Value = IIf(IsNull(!StoreID), "", !StoreID)
            GridSale.Columns("Qty").Value = !Qty
            GridSale.Columns("QtyOrigional").Value = !Qty
            GridSale.Columns("Price").Value = !Price
            GridSale.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            GridSale.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            GridSale.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            GridSale.Columns("SC").Value = IIf(IsNull(!SC), "", !SC)
            GridSale.Columns("Amount").Value = !Amount
            GridSale.Columns("IsProduct").Value = Abs(!isProduct)
            GridSale.Columns("TotalAmount").Value = Val(!Qty) * (Val(!Price) + Val(IIf(IsNull(!SC), "", !SC)))
            GridSale.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            GridSale.Columns("EmpComm").Value = IIf(IsNull(!EmpComm), 0, !EmpComm)
            TxtTotSaleQty.Text = Val(TxtTotSaleQty.Text) + Val(!Qty)
            vTotDisc = vTotDisc + Val(!DiscVal)
            vTotal = vTotal + GridSale.Columns("TotalAmount").Value
            .MoveNext
         Wend
         .Close
      End With
      Call SubCalculateSBody
      GridSale.AddNew
      GridSale.Columns("ProductID").Text = " "
      GridSale.AllowAddNew = False
      GridSale.Redraw = True
      
   End If
    RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   vStrSQL = "select * from SaleBodySerial  where BillID=" & Val(TxtSSID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   RsBodySerial.Open vStrSQL, CN, adOpenDynamic, adLockBatchOptimistic
   
   PopulateDataToGridSSerial
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      LblCost.Visible = False
      LblRCost.Visible = False
   End If
End Sub

Private Sub GridSale_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   vTotal = vTotal - Val(GridSale.Columns("TotalAmount").Text)
   TxtTotSaleQty.Text = Val(TxtTotSaleQty.Text) - Val(GridSale.Columns("Qty").Text)
   vTotDisc = vTotDisc - Val(GridSale.Columns("DiscVal").Text)
   SubCalculateSFooter
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridSale_DblClick()
   Call GridSale_LostFocus
End Sub

Private Sub GridSale_GotFocus()
   Flag = True
   TxtSCode.Enabled = False
   BtnSProduct.Enabled = False
   'TxtSCode.BackColor = TxtSProductName.BackColor
   'TxtSCode.TabStop = False
End Sub

Private Sub GridSale_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub GridSale_LostFocus()
   Flag = False
   LblCost.Visible = False
   If Trim(GridSale.Columns("ProductID").Text) = "" Then
      TxtSCode.Text = ""
      TxtSCode.Enabled = True
      BtnSProduct.Enabled = True
'      If TxtSCode.Enabled And TxtSCode.Visible Then TxtSCode.SetFocus
   Else
      TxtSCode.Enabled = False
      BtnSProduct.Enabled = False
      If TxtSQty.Enabled = True And TxtSQty.Visible Then TxtSQty.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
End Sub

Private Sub GridSale_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(GridSale.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub GridSale_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then
      Call GetDataBackFromGridSaleToTexBoxes
      Call PopulateDataToGridSSerial
   End If
End Sub

Private Sub MniCostPrice_Click()
   On Error GoTo ErrorHandler
'   If Trim(GridSale.Columns("Cost").Text) = "" Then Exit Sub
   If ObjUserSecurity.ShowPurchasePriceInInvoice = True Or ObjUserSecurity.IsAdministrator = True Then
'   LblCost.Caption = GridSale.Columns("Cost").Value
      LblCost.Visible = True
      LblRCost.Visible = True
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub mniRemoveRow_Click()
'   On Error GoTo ErrorHandler
'   If Trim(GridSale.Columns("Code").Text) = "" Then Exit Sub
'   RsSaleBody.Filter = "Code='" & TxtSCode.Text & "'"
'   If RsSaleBody.RecordCount > 0 Then RsSaleBody.Delete
'   GridSale.SelBookmarks.RemoveAll
'   GridSale.SelBookmarks.Add GridSale.Bookmark
'   GridSale.DeleteSelected
'   RsSaleBody.Filter = 0
'   GridSale.MoveLast
'   GetDataBackFromGridSaleToTexBoxes
'Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub

Private Sub GetDataFromTexBoxesToGridSale()
   Dim vrowcounter As Integer
   If Trim(TxtSCode.Text) = "" Then
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
      TxtSCode.SetFocus
      Exit Sub
   End If
   If Val(TxtSQty.Text) = 0 Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      TxtSQty.SetFocus
      Exit Sub
   End If
   If Round(Val(TxtSDiscPer.Text), 2) <> Round((Val(TxtSDiscPC.Text) * 100) / (IIf(Val(TxtSPrice.Text) = 0, 1, Val(TxtSPrice.Text))), 2) Then
      MsgBox "Please update the Discount for change Price.", vbExclamation, "Alert"
      If TxtSDiscPer.Enabled And TxtSDiscPer.Visible Then TxtSDiscPer.SetFocus
      Exit Sub
   End If
    If (CmbSColourName.Text = "" Or cmbSSizeName.Text = "") And vColour = True Then
      MsgBox "Please Select Colour and Size", vbInformation + vbOKOnly, "Error"
      Exit Sub
   End If
   
   '''''''''   check Serial
   RsBodySerial.Filter = "ProductID =" & Val(TxtSCode.Text)
   If (TxtSCode.Enabled = False And RsBodySerial.RecordCount <> 0) And RsBodySerial.RecordCount <> TxtSQty.Text Then
      MsgBox "Qty Should be equal to Serial", vbInformation + vbOKOnly, "Error"
      Call SubClearSDetailArea
      If TxtSCode.Enabled And TxtSCode.Visible Then TxtSCode.SetFocus
      Exit Sub
   End If
   RsBodySerial.Filter = ""
''''''''

   If ObjRegistry.NegativeSale = False Then
      If vIsNewRecord = True Then
         If (Val(vQtyLoose) - Val(TxtSQty.Text)) < 0 Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            GridSale.Redraw = True
            Call SubClearSDetailArea
            If TxtSCode.Enabled And TxtSCode.Visible Then TxtSCode.SetFocus
            Exit Sub
         End If
      Else
         If (Val(vQtyLoose) - Val(TxtSQty.Text) + Val(GridSale.Columns("Qty").Value)) < 0 Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            GridSale.Redraw = True
            Call SubClearSDetailArea
            If TxtSCode.Enabled And TxtSCode.Visible Then TxtSCode.SetFocus
            Exit Sub
         End If
      End If
   End If
On Error GoTo ErrorHandler

RsSaleBody.Filter = "ProductID = " & Val(TxtPID.Text)
   If TxtSCode.Enabled Then
      If RsSaleBody.RecordCount = 0 Then
'         If Trim(TxtSQty.Text) > Val(LblStock.Caption) Then
'            MsgBox "Insufficent Stock.", vbExclamation, "Alert"
'            TxtSQty.SetFocus
'            Exit Sub
'         End If
         RsSaleBody.AddNew
         GridSale.Columns("ProductID").Text = TxtPID.Text
         GridSale.Columns("Code").Text = TxtSCode.Text
         RsSaleBody!Productid = TxtPID.Text
         RsSaleBody!Code = TxtSCode.Text
         RsSaleBody!StoreID = TxtStoreID.Text
      Else
         GridSale.Redraw = False
         GridSale.MoveFirst
            For vrowcounter = 1 To GridSale.rows
               If GridSale.Columns("Productid").Text = TxtPID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearSDetailArea
                  If ObjRegistry.NegativeSale = False Then
                     If vIsNewRecord = True Then
                       If (Val(vQtyLoose) - Val(TxtSQty.Text) - Val(GridSale.Columns("Qty").Value)) < 0 Then
                         MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                         GridSale.MoveLast
                         GridSale.Redraw = True
                         Exit Sub
                       End If
                     Else
                       If (Val(vQtyLoose) - Val(TxtSQty.Text) - Val(GridSale.Columns("Qty").Value) + Val(GridSale.Columns("QtyOrigional").Value)) < 0 Then
                         MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                         GridSale.MoveLast
                         GridSale.Redraw = True
                         Exit Sub
                       End If
                    End If
                  End If
                  ssql = "Select Productid From salebody where sid=" & Val(TxtSSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(GridSale.Columns("ProductID").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmReplacementInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Effected Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmReplacementInvoice, eEdit, TxtReplaceID.Text, DtpReplaceDate.DateValue, "Effected Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
                     End If
                  End With
                  
                  TxtSQty.Text = Val(TxtSQty.Text) + GridSale.Columns("Qty").Value
                  vTotal = vTotal + Val(TxtActualAmount.Text) - Val(GridSale.Columns("TotalAmount").Text)
                  TxtTotSaleQty.Text = Val(TxtTotSaleQty.Text) + Val(TxtSQty.Text) - Val(GridSale.Columns("Qty").Text)
                  vTotDisc = vTotDisc + Val(TxtSDiscVal.Text) - Val(GridSale.Columns("DiscVal").Text)
                  GridSale.Columns("ProductName").Text = TxtSProductName.Text
                  GridSale.Columns("Qty").Value = Val(TxtSQty.Text)
                  GridSale.Columns("Price").Value = Val(TxtSPrice.Text)
                  GridSale.Columns("DiscPC").Value = Val(TxtSDiscPC.Text)
                  GridSale.Columns("DiscPer").Value = Val(TxtSDiscPer.Text)
                  GridSale.Columns("DiscVal").Value = Val(TxtSDiscVal.Text)
                  GridSale.Columns("SC").Value = Val(TxtSSC.Text)
                  GridSale.Columns("Amount").Value = Val(TxtSAmount.Text)
                  GridSale.Columns("EmpComm").Value = Val(TxtEmpComm.Text)
                  GridSale.Columns("Cost").Value = Val(TxtCost.Text)
                  GridSale.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
                  GridSale.Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
                  RsSaleBody!Qty = Val(TxtSQty.Text)
                  RsSaleBody!Price = Val(TxtSPrice.Text)
                  RsSaleBody!DiscPC = Val(TxtSDiscPC.Text)
                  RsSaleBody!DiscPer = Val(TxtSDiscPer.Text)
                  RsSaleBody!DiscVal = Val(TxtSDiscVal.Text)
                  RsSaleBody!SC = IIf(Val(TxtSSC.Text) = 0, Null, Val(TxtSSC.Text))
                  RsSaleBody!Cost = Val(TxtCost.Text)
                  RsSaleBody!EmpComm = Val(TxtEmpComm.Text)
                  RsSaleBody!isProduct = Abs(ChkIsProduct.Value)
                  RsSaleBody!Amount = Val(TxtSAmount.Text)
                  ssql = "Select Productid From salebody where sid = " & Val(TxtSSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid='" & GridSale.Columns("ProductID").Text & "'"
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmReplacementInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Updated Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmReplacementInvoice, eEdit, TxtReplaceID.Text, DtpReplaceDate.DateValue, "Updated Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmReplacementInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Pending Update Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
                 
                  GridSale.MoveLast
                  Call SubClearSDetailArea
                  TxtSCode.SetFocus
                  GridSale.Redraw = True
                  Exit Sub
               End If
               GridSale.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearSDetailArea
         GridSale.MoveLast
         TxtSCode.SetFocus
         Exit Sub
      End If
   End If
   'GridSale.Redraw = False
   With GridSale
'        If ObjRegistry.NegativeSale = False Then
'           If vIsNewRecord = True Then
'              If (Val(vQtyLoose) - Val(TxtSQty.Text)) < 0 Then
'                 MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                 GridSale.Redraw = True
'                 Exit Sub
'              End If
'           Else
'              If (Val(vQtyLoose) - Val(TxtSQty.Text) + Val(GridSale.Columns("QtyOrigional").Value)) < 0 Then
'                 MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                 GridSale.Redraw = True
'                 Exit Sub
'              End If
'           End If
'        End If
      If TxtSCode.Enabled = True Then
         vTotal = vTotal + Val(TxtActualAmount.Text)
         TxtTotSaleQty.Text = Val(TxtTotSaleQty.Text) + Val(TxtSQty.Text)
         vTotDisc = vTotDisc + Val(TxtSDiscVal.Text)
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmReplacementInvoice, eAddNewRowByEdit, TxtReplaceID.Text, DtpReplaceDate.DateValue, "Add New Replace Out Code-" & TxtSCode.Text & " Qty-" & TxtSQty.Text & " Price-" & TxtSPrice.Text & " Disc-" & TxtSDiscPer.Text & " Amount-" & TxtSAmount.Text)
         Call ActivityLogBin(vRandomID, eFrmReplacementInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Pending Add New Replace Out Code-" & TxtSCode.Text & " Qty-" & TxtSQty.Text & " Price-" & TxtSPrice.Text & " Disc-" & TxtSDiscPer.Text & " Amount-" & TxtSAmount.Text)
      Else
         vTotal = vTotal + Val(TxtActualAmount.Text) - Val(GridSale.Columns("TotalAmount").Text)
         TxtTotSaleQty.Text = Val(TxtTotSaleQty.Text) + Val(TxtSQty.Text) - Val(GridSale.Columns("Qty").Text)
         vTotDisc = vTotDisc + Val(TxtSDiscVal.Text) - Val(GridSale.Columns("DiscVal").Text)
         ssql = "Select Productid From salebody where sid=" & Val(TxtSSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(GridSale.Columns("Code").Text)
         With CN.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmReplacementInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Effected Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmReplacementInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Updated Replace Out Code-" & TxtSCode.Text & " Qty-" & TxtSQty.Text & " Price-" & TxtSPrice.Text & " Disc-" & Val(TxtSDiscPer.Text) & " Amount-" & TxtSAmount.Text)
            Else
               Call ActivityLogBin("", eFrmReplacementInvoice, eEdit, TxtReplaceID.Text, DtpReplaceDate.Date, "Effected Replace Out Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmReplacementInvoice, eEdit, TxtReplaceID.Text, DtpReplaceDate.Date, "Updated Replace Out Code-" & TxtSCode.Text & " Qty-" & TxtSQty.Text & " Price-" & TxtSPrice.Text & " Disc-" & Val(TxtSDiscPer.Text) & " Amount-" & TxtSAmount.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmReplacementInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReplaceID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReplaceDate.Date), "Pending Update Replace Out Code-" & TxtSCode.Text & " Qty-" & TxtSQty.Text & " Price-" & TxtSPrice.Text & " Disc-" & TxtSDiscPer.Text & " Amount-" & TxtSAmount.Text)
      End If
      
      GridSale.Columns("ColourName").Text = CmbSColourName.Text
      If CmbSColourName.Text <> "" Then GridSale.Columns("ColourID").Value = CmbSColourName.ItemData(CmbSColourName.ListIndex)
      GridSale.Columns("SizeName").Text = cmbSSizeName.Text
      If cmbSSizeName.Text <> "" Then GridSale.Columns("SizeID").Value = cmbSSizeName.ItemData(cmbSSizeName.ListIndex)
      
      
      If vColour = True And GridSale.Columns("ColourID").Text <> "" Then
         RsSaleBody!ColourID = GridSale.Columns("ColourID").Text
         RsSaleBody!SizeID = GridSale.Columns("SizeID").Text
      End If
      
      .Columns("ProductName").Text = TxtSProductName.Text
      .Columns("Qty").Value = Val(TxtSQty.Text)
      .Columns("Price").Value = Val(TxtSPrice.Text)
      .Columns("DiscPC").Value = Val(TxtSDiscPC.Text)
      .Columns("DiscPer").Value = Val(TxtSDiscPer.Text)
      .Columns("DiscVal").Value = Val(TxtSDiscVal.Text)
      .Columns("SC").Value = Val(TxtSSC.Text)
      If Trim(TxtCost.Text) <> "" Then
         .Columns("Cost").Value = Val(TxtCost.Text)
      End If
      .Columns("EmpComm").Value = Val(TxtEmpComm.Text)
      .Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
      .Columns("Amount").Value = Val(TxtSAmount.Text)
      .Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
      RsSaleBody!Qty = Val(TxtSQty.Text)
      RsSaleBody!Price = Val(TxtSPrice.Text)
      RsSaleBody!DiscPC = Val(TxtSDiscPC.Text)
      RsSaleBody!DiscPer = Val(TxtSDiscPer.Text)
      RsSaleBody!DiscVal = Val(TxtSDiscVal.Text)
      RsSaleBody!SC = IIf(Val(TxtSSC.Text) = 0, Null, Val(TxtSSC.Text))
      If Trim(TxtCost.Text) <> "" Then
         RsSaleBody!Cost = Val(TxtCost.Text)
      End If
      If IsNull(RsSaleBody!Cost) Then RsSaleBody!Cost = 0
      RsSaleBody!EmpComm = Val(TxtEmpComm.Text)
      RsSaleBody!isProduct = Abs(ChkIsProduct.Value)
      RsSaleBody!Amount = Val(TxtSAmount.Text)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearSDetailArea
   TxtSCode.SetFocus
   GridSale.Redraw = True
   Exit Sub
ErrorHandler:
   GridSale.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SubClearSDetailArea()
   CmbSColourName.Clear
   cmbSSizeName.Clear
   TxtSCode.Enabled = True
   BtnSProduct.Enabled = True
   TxtSCode.Text = ""
   TxtSProductName.Text = ""
   TxtSQty.Text = ""
   TxtSPrice.Text = ""
   TxtSDiscPC.Text = ""
   TxtSDiscPer.Text = ""
   TxtSDiscVal.Text = ""
   TxtSSC.Text = ""
   TxtSAmount.Text = ""
   TxtCost.Text = ""
   TxtEmpComm.Text = ""
   TxtActualAmount.Text = ""
   ChkIsProduct.Value = 1
End Sub

Private Sub GetDataBackFromGridSaleToTexBoxes()
   On Error GoTo ErrorHandler
   With GridSale
      TxtPID.Text = .Columns("ProductID").Text
      TxtSCode.Text = .Columns("code").Text
      TxtSProductName.Text = .Columns("ProductName").Text
      TxtSQty.Text = .Columns("Qty").Text
      TxtSPrice.Text = .Columns("Price").Text
      TxtSDiscPC.Text = .Columns("DiscPC").Value
      TxtSDiscPer.Text = .Columns("DiscPer").Value
      TxtSDiscVal.Text = .Columns("DiscVal").Value
      TxtCost.Text = .Columns("Cost").Value
      TxtEmpComm.Text = .Columns("EmpComm").Value
      TxtSAmount.Text = .Columns("Amount").Text
      TxtActualAmount.Text = .Columns("TotalAmount").Text
      ChkIsProduct.Value = Abs(.Columns("IsProduct").Value)
      
      If Trim(.Columns("ColourName").Text) <> "" Then
         CmbSColourName.AddItem .Columns("ColourName").Text
         CmbSColourName.ItemData(CmbSColourName.NewIndex) = .Columns("ColourID").Text
         CmbSColourName.ListIndex = 0
      End If
      
      If Trim(.Columns("SizeName").Text) <> "" Then
         cmbSSizeName.AddItem .Columns("ColourName").Text
         cmbSSizeName.ItemData(cmbSSizeName.NewIndex) = .Columns("SizeID").Text
         cmbSSizeName.ListIndex = 0
      End If
      
'      With CN.Execute("select QtyLoose from currentstockStore where productid ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text)
'         If .RecordCount > 0 Then
'            vQtyLoose = !QtyLoose
'            LblStock.Caption = !QtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'         Else
'            vQtyLoose = 0
'            LblStock.Caption = 0
'         End If
'      End With
   End With
'        With CN.Execute("select QtyLoose from currentstockStore where productid ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text)
'         If .RecordCount > 0 Then
'            vQtyLoose = !QtyLoose
'            LblStock.Caption = !QtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
'         Else
'            vQtyLoose = 0
'            LblStock.Caption = 0
'         End If
'      End With
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpReplaceDate.DateValue + 1 & "',0),0)"
         vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         
         LblStock.Caption = vQtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit(" & Val(TxtPID.Text) & ")").Fields(0).Value
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         
   If GridSale.rows = 1 Then GridSale.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetSale()
   On Error GoTo ErrorHandler
   ssql = "select h.*, c.AccountName, BankMachineName, StoreName, EmpName FROM SaleHeader h left outer join chartofaccounts c on h.customerid = c.AccountNo left outer join BankMachines b on b.BankMachineid = h.BankMachineid inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID where h.SID=" & Val(TxtSSID.Text) & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With CN.Execute(ssql)
      If Not .BOF Then
         TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
         TxtEmployeeName.Text = IIf(IsNull(!empname), "", !empname)
         TxtSBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
         TxtSBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
         TxtSServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
         FrmReplacePrint.OptCash.Value = !Cash
         FrmReplacePrint.OptCredit.Value = !Credit
         FrmReplacePrint.OptBankCard.Value = !BankCard
         If FrmReplacePrint.OptBankCard.Value = True Then
            FrmReplacePrint.TxtInvoiceNo.Text = IIf(IsNull(!InvoiceNo), "", !InvoiceNo)
            FrmReplacePrint.TxtCommision.Text = IIf(IsNull(!Commision), "", !Commision)
            FrmReplacePrint.TxtBankMachineID.Text = !BankMachineID
            FrmReplacePrint.TxtBankMachineName.Text = !BankMachineName
            FrmReplacePrint.TxtCashReceivedCash.Text = ""
            FrmReplacePrint.TxtCashReceivedCredit.Text = ""
            FrmReplacePrint.TxtCashReceivedBank.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            FrmReplacePrint.TxtCustomerID.Text = ""
            FrmReplacePrint.TxtCustomerName.Text = ""
            FrmReplacePrint.TxtCashCustomer.Text = ""
            FrmReplacePrint.TxtBankCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
         End If
         If FrmReplacePrint.OptCash.Value = True Then
            FrmReplacePrint.TxtCommision.Text = ""
            FrmReplacePrint.TxtInvoiceNo.Text = ""
            FrmReplacePrint.TxtBankMachineID.Text = ""
            FrmReplacePrint.TxtBankMachineName.Text = ""
            FrmReplacePrint.TxtCashReceivedCash.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            FrmReplacePrint.TxtCashReceivedCredit.Text = ""
            'FrmReplacePrint.TxtCashReceivedBank.Text = ""
            FrmReplacePrint.TxtCustomerID.Text = ""
            FrmReplacePrint.TxtCustomerName.Text = ""
            FrmReplacePrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
            FrmReplacePrint.TxtBankCustomer.Text = ""
         End If
         If FrmReplacePrint.OptCredit.Value = True Then
'            FrmReplacePrint.TxtCommision.Text = ""
'            FrmReplacePrint.TxtInvoiceNo.Text = ""
'            FrmReplacePrint.TxtBankMachineID.Text = ""
'            FrmReplacePrint.TxtBankMachineName.Text = ""
'            FrmReplacePrint.TxtCashReceivedCash.Text = ""
            If Val(FrmReplacePrint.TxtCashReceivedCredit.Text) = 0 Then
               FrmReplacePrint.TxtCashReceivedCredit.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
               FrmReplacePrint.TxtCommision.Text = IIf(IsNull(!Commision), "", !Commision)
               FrmReplacePrint.TxtBankMachineCreditID.Text = IIf(IsNull(!BankMachineID), "", !BankMachineID)
               FrmReplacePrint.TxtBankMachineCreditName.Text = IIf(IsNull(!BankMachineName), "", !BankMachineName)
               FrmReplacePrint.TxtBankAmount.Text = IIf(IsNull(!BankAmount), "", !BankAmount)
            End If
            'FrmReplacePrint.TxtCashReceivedBank.Text = ""
            FrmReplacePrint.TxtCustomerID.Text = !CustomerID
            FrmReplacePrint.TxtCustomerName.Text = !AccountName
            FrmReplacePrint.TxtCashCustomer.Text = ""
            FrmReplacePrint.TxtBankCustomer.Text = ""
         End If
      End If
      .Close
   End With
   Call PopulateDataToGridSale
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   GridSale.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtSBillDisc_Change()
   If ActiveControl.Name <> TxtSBillDisc.Name Then Exit Sub
   TxtSBillDiscPer.Text = Round((Val(TxtSBillDisc.Text) * 100) / Val(vTotal - vTotDisc), 2)
   Call SubCalculateSFooter
End Sub

Private Sub TxtSBillDiscPer_Change()
   If ActiveControl.Name <> TxtSBillDiscPer.Name Then Exit Sub
   TxtSBillDisc.Text = SelfRound((Val(vTotal - vTotDisc) * Val(TxtSBillDiscPer.Text) / 100))
   Call SubCalculateSFooter
End Sub

Private Sub TxtSDiscPC_Change()
   If ActiveControl.Name <> TxtSDiscPC.Name Then Exit Sub
   If Val(TxtSPrice.Text) = 0 Then Exit Sub
   TxtSDiscPer.Text = Round((Val(TxtSDiscPC.Text) * 100) / Val(TxtSPrice.Text), 2)
   Call SubCalculateSBody
End Sub

'Private Sub TxtSDiscPC_LostFocus()
'   Select Case Me.ActiveControl.Name
'   Case TxtSCode.Name, TxtSQty.Name, TxtSDiscPC.Name
'      Exit Sub
'   End Select
'   Call GetDataFromTexBoxesToGridSale
'End Sub

Private Sub TxtSDiscPer_Change()
   If ActiveControl.Name <> TxtSDiscPer.Name Then Exit Sub
   TxtSDiscPC.Text = Round((Val(TxtSPrice.Text) * Val(TxtSDiscPer.Text) / 100), 2)
   Call SubCalculateSBody
End Sub

Private Sub TxtSCode_Change()
   If ActiveControl.Name <> TxtSCode.Name Then Exit Sub
   If TxtSProductName.Text <> "" Then
      TxtSProductName.Text = ""
      TxtSPrice.Text = ""
      TxtSDiscPC.Text = ""
   End If
End Sub

Private Sub TxtSCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridSale.SetFocus
End Sub

'Private Sub TxtSCode_LostFocus()
'   If Len(TxtSCode.Text) > 7 Then
'      GetDataFromTexBoxesToGridSale
'   End If
'End Sub

Private Sub TxtSCode_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtSCode.Text) = "" Then Exit Sub
   vTemp = Not FunSelectSProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSDiscVal_Change()
   If TxtSDiscVal.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSDiscVal.Name Then Exit Sub
   If Val(TxtSPrice.Text) = 0 Then Exit Sub
   If Val(TxtSQty.Text) = 0 Then Exit Sub
   TxtSDiscPC.Text = Round(Val(TxtSDiscVal.Text) / (TxtSQty.Text), 3)
   TxtSDiscPer.Text = Round((Val(TxtSDiscPC.Text) * 100) / Val(TxtSPrice.Text), 2)
   TxtActualAmount.Text = Val(TxtSQty.Text) * (Val(TxtSPrice.Text) + Val(TxtSSC.Text))
   TxtSAmount.Text = Val(TxtActualAmount.Text) - Val(TxtSDiscVal.Text)
   SubCalculateSFooter
End Sub

Private Sub TxtSPrice_Change()
   Call SubCalculateSBody
End Sub

Private Sub TxtSQty_Change()
   Call SubCalculateSBody
   Call FindSRebate
End Sub

Private Sub TxtEmployeeID_Change()
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

Private Sub UserActivities()
     If vIsNewRecord = False Then
    With CN.Execute("Select  * from SaleHeader where BillID =" & TxtBillID.Text & " And BillDate = '" & DtpBillDate.DateValue & "'")
        If Val(TxtEmployeeID.Text) <> IIf(IsNull(!EmpID), 0, !EmpID) Then
            CN.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Updated EmpID-" & !EmpID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtMemberID.Text <> !MemberID Then
            CN.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Updated MemberID-" & !MemberID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtStoreID.Text <> !StoreID Then
            CN.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    
    ''''''''''''''''''''''''''''''''''''''''IN''''''''''''''''''''''''''''
    GridReturn.MoveFirst
    For i = 1 To GridReturn.rows - 1
        With CN.Execute("Select * from SaleReturnBody Where ReturnID = " & TxtReturnID.Text & " and ReturnDate ='" & DtpReturnDate.DateValue & "' and Productid = " & Val(GridReturn.Columns("Productid").Text))
             If .EOF = True Then
                CN.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','In Inserted New Code-" & GridReturn.Columns("Code").Text & " Qty-" & GridReturn.Columns("Qty").Text & " Price-" & GridReturn.Columns("Price").Text & " Disc-" & GridReturn.Columns("DiscPer").Text & " Amount-" & GridReturn.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If GridReturn.Columns("Qty").Text <> !Qty Or GridReturn.Columns("Price").Text <> !Price Or GridReturn.Columns("discper").Text <> !DiscPer Then
                   CN.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','In Updated Code-" & GridReturn.Columns("Code").Text & " Qty-" & !Qty & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    GridReturn.MoveNext
    Next
        
    ''''''''''''''''''''''''''''''''''''''''Out''''''''''''''''''''''''''''
    GridSale.MoveFirst
    For i = 1 To GridSale.rows - 1
        With CN.Execute("Select * from SaleBody Where billID = " & TxtBillID.Text & " and billdate ='" & DtpBillDate.DateValue & "' and Productid = " & Val(GridSale.Columns("Productid").Text))
             If .EOF = True Then
                CN.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Out Inserted New Code-" & GridSale.Columns("Code").Text & " Qty-" & GridSale.Columns("Qty").Text & " Price-" & GridSale.Columns("Price").Text & " Disc-" & GridSale.Columns("DiscPer").Text & " Amount-" & GridSale.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If GridSale.Columns("Qty").Text <> !Qty Or GridSale.Columns("Price").Text <> !Price Or GridSale.Columns("discper").Text <> !DiscPer Then
                   CN.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Out Updated Code-" & GridSale.Columns("Code").Text & " Qty-" & !Qty & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    GridSale.MoveNext
    Next
    
   Else
    CN.Execute ("Insert Into UserActivities values ('Replacement Invoice'" & "," & TxtReplaceID.Text & ",'" & DtpReplaceDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
   
End Sub
Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      ''''''''''''''''''''' Sale Invoice ''''''''''''''''''''
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmSaleInvoicePOS) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmSaleInvoicePOS & ", " & vUser & "," & TableHeaderFields(eFrmSaleInvoicePOS) & " from SaleHeader " & vbCrLf _
             & "Where SID = " & TxtSSID.Text
      CN.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleBodyBin (" & TableBodyFields(eFrmSaleInvoicePOS) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmSaleInvoicePOS) & " from SaleBody " & vbCrLf _
             & "Where SID = " & TxtSSID.Text
      CN.Execute vStrSQL
      ''''''''''''''''''''' Sale Return Invoice ''''''''''''''''''''
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleReturnHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmSaleReturnInvoiceDIS) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmSaleReturnInvoiceDIS & ", " & vUser & "," & TableHeaderFields(eFrmSaleReturnInvoiceDIS) & " from SaleReturnHeader " & vbCrLf _
             & "Where SID = " & TxtRSID.Text
      CN.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleReturnBodyBin (" & TableBodyFields(eFrmSaleReturnInvoiceDIS) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmSaleReturnInvoiceDIS) & " from SaleReturnBody " & vbCrLf _
             & "Where SID = " & TxtRSID.Text
      CN.Execute vStrSQL
      ''''''''''''''''''''' Replacement Invoice ''''''''''''''''''''
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.ReplacementHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmReplacementInvoice) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmReplacementInvoice & ", " & vUser & "," & TableHeaderFields(eFrmReplacementInvoice) & " from ReplacementHeader " & vbCrLf _
             & "Where SID = " & TxtSID.Text
      CN.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub TxtSSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridSSerial.SetFocus
End Sub

Private Sub PopulateDataToGridSSerial()
      If Trim(GridSale.Columns("ProductID").Text) = "" Then
      RsBodySerial.Filter = 0
   Else
      RsBodySerial.Filter = "ProductID = '" & GridSale.Columns("ProductID").Text & "'"
   End If
   GridSSerial.Redraw = False
   GridSSerial.MoveFirst
   GridSSerial.RemoveAll
   GridSSerial.AllowAddNew = True
   If RsBodySerial.RecordCount > 0 Then
      With RsBodySerial
         .MoveFirst
         While Not .EOF
            GridSSerial.AddNew
            GridSSerial.Columns("ProductID").Text = !Productid
            GridSSerial.Columns("Serial").Text = !Serial
            .MoveNext
         Wend
'      .Close
      GridSSerial.MoveLast
      End With
   End If
   GridSSerial.AddNew
   GridSSerial.Columns("ProductID").Text = " "
   GridSSerial.AllowAddNew = False
   GridSSerial.Redraw = True
   RsBodySerial.Filter = 0

End Sub

Private Sub SubClearSSerialFields()
   TxtSSerial.Text = ""
'   TxtSSerial.Enabled = False
   GridSSerial.CancelUpdate
   GridSSerial.RemoveAll
   GridSSerial.AddNew
   GridSSerial.Columns("Serial").Text = " "
   GridSSerial.Update
End Sub

Private Sub TxtRSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridRSerial.SetFocus
End Sub

Private Sub PopulateDataToGridRSerial()
   If Trim(GridReturn.Columns("ProductID").Text) = "" Then
      RsReturnSerial.Filter = 0
   Else
      RsReturnSerial.Filter = "ProductID = '" & GridReturn.Columns("ProductID").Text & "'"
   End If
   
   If RsReturnSerial.RecordCount > 0 Then
       ssql = "select d.* from SaleReturnSerial d  where ReturnID=" & Val(TxtRSID.Text) & " and ReturnDate='" & DtpBillDate.DateValue & "' and ProductID = " & Val(GridReturn.Columns("ProductID").Text)
'      With cn.Execute(ssql)
       With RsReturnSerial
         GridRSerial.Redraw = False
         GridRSerial.MoveFirst
         GridRSerial.RemoveAll
         GridRSerial.AllowAddNew = True
         .MoveFirst
         While Not .EOF
            GridRSerial.AddNew
            GridRSerial.Columns("ProductID").Text = !Productid
            GridRSerial.Columns("Serial").Text = !Serial
            .MoveNext
         Wend
'      .Close
      End With
      GridRSerial.AddNew
      GridRSerial.Columns("Serial").Text = " "
      GridRSerial.AllowAddNew = False
      GridRSerial.Redraw = True
   Else
    Call SubClearRSerialFields
   End If
   RsReturnSerial.Filter = 0
End Sub

Private Sub SubClearRSerialFields()
   TxtRSerial.Text = ""
'   TxtPSerial.Enabled = False
   GridRSerial.CancelUpdate
   GridRSerial.RemoveAll
   GridRSerial.AddNew
   GridRSerial.Columns("Serial").Text = " "
   GridRSerial.Update
End Sub

Private Sub GetDataFromTexBoxesToGridSSerial()
   
  vStrSQL = "Select ProductID, Serial, SerialAdd from vuPurchaseSerial where Serial = '" & Trim(TxtSSerial.Text) & "'"
   
   With CN.Execute(vStrSQL)
      If .EOF Then
         MsgBox "The Serail cannot be inserted because it is not Exist", vbInformation + vbOKOnly, "Error"
         TxtRSerial.Text = ""
         Exit Sub
      ElseIf !SerialAdd = False Then
         MsgBox "The Serial Already Sold", vbInformation + vbOKOnly, "Error"
         TxtRSerial.Text = ""
         Exit Sub
      End If
   End With
   
   GridSSerial.MoveLast
   
   RsBodySerial.Filter = ""
   RsBodySerial.Filter = "ProductID =" & TxtSCode.Text & " And Serial='" & TxtSSerial.Text & "'"
   If RsBodySerial.RecordCount > 0 Then
      MsgBox "The Serail cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
      TxtSSerial.Text = ""
      Exit Sub
   End If
   RsBodySerial.Filter = ""
   If TxtSSerial.Enabled Then
        
         GridSSerial.MoveLast
         GridSSerial.Columns("ProductID").Text = TxtSCode.Text
         GridSSerial.Columns("Serial").Text = TxtSSerial.Text
         
         RsBodySerial.AddNew
         RsBodySerial!Productid = TxtSCode.Text
         RsBodySerial!Serial = TxtSSerial.Text
         vSerialAdd = True
         TxtSSerial.Text = ""
  End If
  
   With GridSSerial
      If Trim(.Columns("Serial").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Serial").Text = " "
         .AllowAddNew = False
      End If
   End With
   
   Exit Sub
ErrorHandler:
   GridSSerial.Redraw = True
   Call ShowErrorMessage
End Sub


Private Sub GetDataFromTexBoxesToGridPSerial()
On Error GoTo ErrorHandler

   GridRSerial.MoveLast
   
   RsReturnSerial.Filter = ""
   RsReturnSerial.Filter = "ProductID =" & TxtPCode.Text & " And Serial='" & TxtRSerial.Text & "'"
   If RsReturnSerial.RecordCount > 0 Then
      MsgBox "The Serail cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
      TxtRSerial.Text = ""
      Exit Sub
   End If
   
   If TxtRSerial.Enabled Then
         
         GridRSerial.MoveLast
         GridRSerial.Columns("ProductID").Text = TxtPCode.Text
         GridRSerial.Columns("Serial").Text = TxtPSerial.Text
         
         RsReturnSerial.AddNew
         RsReturnSerial!Productid = TxtPCode.Text
         RsReturnSerial!Serial = TxtPSerial.Text
         RsReturnSerial.Update
         vSerialAdd = True
         TxtPSerial.Text = ""
  End If
  
   With GridRSerial
      If Trim(.Columns("Serial").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Serial").Text = " "
         .AllowAddNew = False
      End If
   End With
   
   Exit Sub
ErrorHandler:
   GridRSerial.Redraw = True
   Call ShowErrorMessage
   End Sub

Private Sub PopulateDataPurchaseSerial()
   If RsPurchaseSerial.State = adStateOpen Then RsPurchaseSerial.Close
   vStrSQL = "select * from PurchaseBodySerial  "
   RsPurchaseSerial.Open vStrSQL, CN, adOpenDynamic, adLockBatchOptimistic
   RsPurchaseSerial.Filter = 0
End Sub

Private Sub GetDataFromTexBoxesToGridRSerial()
   On Error GoTo ErrorHandler
   
   If vIsNewSerial = False Then
      vStrSQL = "Select ProductID, Serial, SerialAdd from vuPurchaseSerial where Serial = '" & Trim(TxtRSerial.Text) & "'"
      
      With CN.Execute(vStrSQL)
         If .EOF = True Then
            MsgBox "The Serail cannot be inserted because it is not Exist", vbInformation + vbOKOnly, "Error"
            TxtRSerial.Text = ""
            Exit Sub
         ElseIf !SerialAdd = True Then
            MsgBox "The Serail Not Sold", vbInformation + vbOKOnly, "Error"
            TxtRSerial.Text = ""
            Exit Sub
         End If
      End With
   Else
      If CN.Execute("Select ProductID, Serial, SerialAdd from vuPurchaseSerial where Serial = '" & Trim(TxtPSerial.Text) & "'").EOF = False Then
      MsgBox "This is not new Serial.", vbOKOnly, "Alert"
      Exit Sub
    End If
   End If
  
   vIsNewSerial = True
   GridRSerial.MoveLast
   
   RsReturnSerial.Filter = ""
   RsReturnSerial.Filter = "ProductID =" & TxtRCode.Text & " And Serial='" & TxtRSerial.Text & "'"
   If RsReturnSerial.RecordCount > 0 Then
      MsgBox "The Serail cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
      TxtRSerial.Text = ""
      Exit Sub
   End If
   
   If TxtRSerial.Enabled Then
         
         GridRSerial.MoveLast
         GridRSerial.Columns("ProductID").Text = TxtRCode.Text
         GridRSerial.Columns("Serial").Text = TxtRSerial.Text
         
         RsReturnSerial.AddNew
         RsReturnSerial!Productid = TxtRCode.Text
         RsReturnSerial!Serial = TxtRSerial.Text
         RsReturnSerial.Update
         vSerialAdd = True
         TxtRSerial.Text = ""
         TxtPSerial.Text = ""
  End If
  
   With GridRSerial
      If Trim(.Columns("Serial").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Serial").Text = " "
         .AllowAddNew = False
      End If
   End With
   
   Exit Sub
ErrorHandler:
   GridRSerial.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SaveReturnIn()

Dim vInvoiceNo, vComission, vBankMachineID, vCashReceived, vBankAmount, vCashPaid, vCustomerID, vCustomerName As String

      If FrmReplacePrint.OptBankCard.Value = True Then
         vInvoiceNo = FrmReplacePrint.TxtInvoiceNo.Text
         vComission = FrmReplacePrint.TxtCommision.Text
         vBankMachineID = FrmReplacePrint.TxtBankMachineID.Text
         vCashPaid = 0
         vCustomerID = "621"
         vCustomerName = IIf(Trim(FrmReplacePrint.TxtBankCustomer.Text) = "", Null, FrmReplacePrint.TxtBankCustomer.Text)
      End If
      If FrmReplacePrint.OptCash.Value = True Then
         vComission = "''"
         vInvoiceNo = Null
         vCashPaid = TxtRNetAmount.Text
         vCustomerID = "621"
         vBankMachineID = Null
         vCustomerName = IIf(Trim(FrmReplacePrint.TxtCashCustomer.Text) = "", Null, FrmReplacePrint.TxtCashCustomer.Text)
      End If
      If FrmReplacePrint.OptCredit.Value = True Then
         vComission = "''"
         vInvoiceNo = Null
         If LblAmount.Caption = "Cash Paid" Then
            vCashPaid = Val(FrmReplacePrint.TxtCashReceivedCredit.Text)
            vBankAmount = Val(FrmReplacePrint.TxtBankAmount.Text)
         Else
            vCashPaid = 0
            vBankAmount = 0
         End If
         vCustomerID = FrmReplacePrint.TxtCustomerID.Text
         vBankMachineID = IIf(Trim(FrmReplacePrint.TxtBankMachineCreditID.Text) = "", Null, FrmReplacePrint.TxtBankMachineCreditID.Text)
'         !BankAmount = Val(FrmReplacePrint.TxtBankAmount.Text)
         If Val(FrmReplacePrint.TxtBankMachineCreditID.Text) > 0 Then
            vComission = Val(FrmReplacePrint.TxtCommision.Text)
         Else
            vComission = "''"
         End If
         vCustomerName = "''"
      End If


vStrPara = ""
vStrPara = Abs(ObjRegistry.AllowContinuousBillNo) & ","
vStrPara = vStrPara & Abs(ObjRegistry.AllowMonthlyBillNo) & ","
vStrPara = vStrPara & Abs(ObjRegistry.AllowDailyBillNo) & "," 'AllowDailyBillNo
vStrPara = vStrPara & Val(TxtRSID.Text) & "," 'RSID
vStrPara = vStrPara & TxtReturnID.Text & ","
vStrPara = vStrPara & "'" & DtpReturnDate.DateValue & "',"
vStrPara = vStrPara & "'" & vCustomerID & "'," 'CustomerID
'vStrPara = vStrPara & SelfRound(vTotalAmount) & "," ' Total Amount
vStrPara = vStrPara & SelfRound(TxtRNetAmount.Text + Val(TxtRBillDisc.Text) - Val(TxtRServiceCharges.Text)) & ","     ' Total Amount
vStrPara = vStrPara & Val(TxtRBillDisc.Text) & "," 'BillDisc
vStrPara = vStrPara & IIf(IsNull(vCashPaid), 0, vCashPaid) & "," ' 'CashPaid
vStrPara = vStrPara & vUser & "," 'UserNo
vStrPara = vStrPara & TxtStoreID.Text & "," 'StoreID
vStrPara = vStrPara & IIf(FrmReplacePrint.OptBankCard.Value = True, 1, 0) & "," 'BankCard
vStrPara = vStrPara & IIf(FrmReplacePrint.OptCredit.Value = True, 1, 0) & "," 'Credit
vStrPara = vStrPara & IIf(FrmReplacePrint.OptCash.Value = True, 1, 0) & "," 'Cash
vStrPara = vStrPara & "'" & vBankMachineID & "'," 'BankMachineID
vStrPara = vStrPara & "'" & vInvoiceNo & "',"  'InvoiceNo
vStrPara = vStrPara & "'" & vCustomerName & "'," 'CustomerName
vStrPara = vStrPara & Val(TxtRBillDiscPer.Text) & "," 'BillDiscPer
vStrPara = vStrPara & vComission & ","     'Commision
vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtCommission.Text)) & "," 'EmpComm
vStrPara = vStrPara & "'" & IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text)) & "'," 'EmpID
vStrPara = vStrPara & 1 & "," 'isReplace
vStrPara = vStrPara & 0 & "," 'isPosted
vStrPara = vStrPara & IIf(Trim(TxtMemberID.Text) = "", "''", TxtMemberID.Text) & "," 'MemberID
'vStrPara = vStrPara & "'" & vNow & "'," 'BillTime
vStrPara = vStrPara & "'" & vIsNewRecord & "'," 'Tag
vStrPara = vStrPara & "'" & IIf(Trim(TxtManualBillNo.Text) = "", Null, TxtManualBillNo.Text) & "'," 'ManualBillNo
vStrPara = vStrPara & "'" & Null & "',"  'Remarks
vStrPara = vStrPara & IIf(Trim(TxtOrganizationID.Text) = "", "''", TxtOrganizationID.Text) & ","  'OrganizationID
vStrPara = vStrPara & "'" & Null & "'," ' BillNo
vStrPara = vStrPara & "'" & Null & "'," ' Bilty No
vStrPara = vStrPara & "'" & Null & "'," 'Description
vStrPara = vStrPara & "''" & "," 'PAIDAMOUNT
vStrPara = vStrPara & "'" & Null & "',"  'EntryDate
vStrPara = vStrPara & IIf(FrmReplacePrint.OptCredit.Value = True, 0, 0) & "," 'PreviousAmount
vStrPara = vStrPara & 0 & "," 'OtherCharges
vStrPara = vStrPara & "'" & Null & "'," 'SaleManID
vStrPara = vStrPara & 0 & "," 'TotalExpense
'vStrPara = vStrPara & IIf(Val(TxtOrderID.Text) = 0, "''", TxtOrderID.Text) & "," 'OrderID
'vStrPara = vStrPara & "'" & DtpOrderDate.DateValue & "'," 'OrderDate
'vStrPara = vStrPara & 0 & "," 'Freight
'vStrPara = vStrPara & 0 & "," 'IsCustomerFreight
vStrPara = vStrPara & "'" & Null & "'," 'VechicleNo
vStrPara = vStrPara & IIf(TxtRServiceCharges.Text = "", "''", Val(TxtRServiceCharges.Text)) & "," 'ServiceCharges
vStrPara = vStrPara & "''" & "," 'ServiceChargesPer
vStrPara = vStrPara & "''" & "," 'STax
vStrPara = vStrPara & "''" & ","  'STaxPer
vStrPara = vStrPara & "'" & Null & "',"  'TableID
vStrPara = vStrPara & "'" & Now & "'," 'ServerEntry
'vStrPara = vStrPara & "'" & IIf(CmbType.Visible = False, Null, CmbType.Text) & "'," 'InvType
'vStrPara = vStrPara & "'" & DtpDeliveryDate.DateValue & "'," 'DeliveryDate
'vStrPara = vStrPara & "'" & DTPDeliveryTime.Value & "'," 'DeliveryTime
'vStrPara = vStrPara & "'" & Null & "'," 'isPrinted
'vStrPara = vStrPara & "'" & Null & "'," 'RemarksUrdu
'vStrPara = vStrPara & "Default" & ","  'StampID
vStrPara = vStrPara & 0 & "," 'isTransfer
'vStrPara = vStrPara & IIf(DtpPromiseDate.DateValue = Empty, "Null", "'" & DtpPromiseDate.DateValue & "'") & "," 'PromiseDate
'vStrPara = vStrPara & "Null," 'Expiry Invoice
'vStrPara = vStrPara & "Null," 'Syllabus
vStrPara = vStrPara & "'" & IIf(Trim(vSessionID) = 0, Null, Val(vSessionID)) & "',"  'vSessionID
'vStrPara = vStrPara & IIf(TxtAdvTaxVal.Text = "", "''", Val(TxtAdvTaxVal.Text)) & "," 'AdvTaxVal
'vStrPara = vStrPara & IIf(TxtAdvTaxPer.Text = "", "''", Val(TxtAdvTaxPer.Text)) & "," 'AdvTaxPer
'vStrPara = vStrPara & IIf(TxtExtraTaxVal.Text = "", "''", Val(TxtExtraTaxVal.Text)) & "," 'ExtraTaxVal
'vStrPara = vStrPara & IIf(TxtExtraTaxPer.Text = "", "''", Val(TxtExtraTaxPer.Text)) & "," 'ExtraTaxPer
'vStrPara = vStrPara & "'" & IIf(Trim(TxtCNIC.Text) = "", Null, TxtCNIC.Text) & "',"  'CNIC
'vStrPara = vStrPara & "'" & IIf(Trim(TxtCellNo.Text) = "", Null, TxtCellNo.Text) & "',"  'CellNo
'vStrPara = vStrPara & Val(TxtSumDiscAmount.Text) & "," 'Sum Disc Amount
'vStrPara = vStrPara & "Null," 'DispatchDate
'vStrPara = vStrPara & "Null," 'Terms
'vStrPara = vStrPara & "'" & IIf(Trim(TxtRefID.Text) = "", Null, TxtRefID.Text) & "',"  'RefID
'vStrPara = vStrPara & "'" & IIf(Trim(TxtRefComm.Text) = "", Null, TxtRefComm.Text) & "',"  'Refcomm
vStrPara = vStrPara & IIf(vBankAmount = 0, "''", Val(vBankAmount)) 'Bank Amount in Credit Option
vStrPara = Replace(vStrPara, "''", "Null")

vStrPara = "DECLARE @returnvalue INT EXEC @returnvalue = SaleReturnHeaderInsert " & vStrPara & " Select @returnvalue"
   vMasterID = CN.Execute(vStrPara).Fields(0).Value
   TxtRSID.Text = vMasterID
   '/******* FBR Integeration*************/
'   If vPOSID <> "" Then
'      If ObjRegistry.AllowFBRContinuousNo Then
'         vUSIN = cn.Execute("select isnull(max(USIN),0) + 1 as USIN from SaleReturnHeader").Fields(0).Value
'      Else
'         vUSIN = TxtSID.Text
'      End If
'      vHeader = "{InvoiceNumber:'',POSID:'" & vPOSID & "',DateTime:'" & Replace(DtpBillDate.Date, "/", "-") & "',BuyerName:'" & FrmReplacePrint.TxtCustomerName.Text & "',TotalQuantity:" & Val(TxtTotalQty.Caption) & ",TotalSaleValue:" & Val(TxtNetAmount.Caption) - Val(TxtTotalSaleTaxValue.Text) + Val(TxtTotalDiscount.Caption) & ",Totaltaxcharged:" & Val(TxtTotalSaleTaxValue.Text) & ",Discount:" & Val(TxtTotalDiscount.Caption) & ",TotalBillAmount:" & Val(TxtNetAmount.Caption) + Val(TxtTotalDiscount.Caption) & ",PaymentMode:1,InvoiceType:1,USIN:'" & vUSIN & "', items : ["
'   End If
   ''''''''''''''''''''''''''''

''' insert Sale Return Body
vStrDetail = ""
vGridRows = 0
vProducts = ""
i = 0
vSamePid = ""
With GridReturn
 .Redraw = False
 .MoveFirst
   For vCounter = 1 To .rows
      If Trim(.Columns("Productid").Text) <> "" Then
        '''''' ActivityLogBin Follwoin lines check the same product id which was enter seperate row or new new row
        If (InStr(1, vSamePid, .Columns("Productid").Text)) = 0 Then vGridRows = vGridRows + 1
        vSamePid = vSamePid & " , " & .Columns("Productid").Text
        '''''''''''''''''''''''''''''''''''''

      ''''''''''''''''''''''''''''
        vStrPara = ""
        TxtReturnID.Text = CN.Execute("Select ReturnID from SaleReturnheader where SID = " & vMasterID).Fields(0).Value
        vStrPara = vStrPara & "'" & 0 & "'," 'check stock update or not
        vStrPara = vStrPara & vMasterID & ","
        vStrPara = vStrPara & TxtReturnID.Text & ","
        vStrPara = vStrPara & "'" & DtpReturnDate.DateValue & "',"
        'vStrPara = vStrPara & .Columns("SerialNo").Text & ","
        'vStrPara = vStrPara & .Columns("BillID").Text & ","
        'vStrPara = vStrPara & .Columns("BillDate").Text & ","
        vStrPara = vStrPara & .Columns("ProductID").Text & ","
        vStrPara = vStrPara & .Columns("Qty").Text & ","
        vStrPara = vStrPara & .Columns("Price").Text & ","
        vStrPara = vStrPara & .Columns("DiscPC").Text & ","
        vStrPara = vStrPara & .Columns("Amount").Text & ","
        vStrPara = vStrPara & "'" & .Columns("Code").Text & "',"
        vStrPara = vStrPara & .Columns("DiscPer").Text & ","
        vStrPara = vStrPara & .Columns("DiscVal").Text & ","

        vStrPara = vStrPara & 0 & "," ' isDiscB4TradeOffer
        vStrPara = vStrPara & 0 & ","   'isDiscB4ExtraScheme
        vStrPara = vStrPara & 0 & "," 'isDiscB4SaleTax
        vStrPara = vStrPara & "''" & ","  'TradeOffer1
        vStrPara = vStrPara & "''" & ","   'TradeOffer2
        vStrPara = vStrPara & "''" & ","   'ExtraSchemePer
        vStrPara = vStrPara & "''" & ","   'TradeValue
        vStrPara = vStrPara & "''" & ","   'ExtraSchemeValue

        vStrPara = vStrPara & Val(.Columns("Cost").Text) & ","
        vStrPara = vStrPara & Val(.Columns("isProduct").Text) & ","
        vStrPara = vStrPara & "''" & "," ' Pack Name
        vStrPara = vStrPara & "''" & "," ' Qty Pack
        vStrPara = vStrPara & "''" & "," ' Pack
        vStrPara = vStrPara & "''" & "," ' Bonus
        vStrPara = vStrPara & "''" & "," 'Offer
        vStrPara = vStrPara & 0 & ","  'SaleTaxPer
        vStrPara = vStrPara & 0 & ","  ' SaleTaxVal
'        vStrPara = vStrPara & Val(.Columns("TokenVal").Text) & ","
'        vStrPara = vStrPara & Val(TxtPrice.Text) & "," 'RetailPrice
        vStrPara = vStrPara & 0 & "," 'IsWSSaleTax
        vStrPara = vStrPara & 0 & ","  'IsRetailSaleTax
        vStrPara = vStrPara & 0 & ","  'IsWSDiscb4ST
        vStrPara = vStrPara & Val(.Columns("SC").Text) & "," 'SC
        vStrPara = vStrPara & Val(.Columns("EmpComm").Value & ",") & ","  'EmpComm
        vStrPara = vStrPara & "''" & "," 'BatchNo
        'vStrPara = vStrPara & "''" & "," 'StampID
        vStrPara = vStrPara & TxtStoreID.Text & ","                  'StoreID
        If ObjRegistry.AllowEmployeProductWise Then
           vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtEmployeeID.Text)) & "," 'EmpID
        Else
           vStrPara = vStrPara & "''" & "," 'EmpID
        End If
        vStrPara = vStrPara & "'" & IIf(Trim(.Columns("ColourID").Text) = "", Null, Val(.Columns("ColourID").Text)) & "'," ' ColourID
        vStrPara = vStrPara & "'" & IIf(Trim(.Columns("SizeID").Text) = "", Null, Val(.Columns("SizeID").Text)) & "'," ' SizeID
        vStrPara = vStrPara & "null" & ","  'Gross Qty
        vStrPara = vStrPara & "null" & ","  'Gross Unit
        If ObjRegistry.AllowStoreProductWise Then
           vStrPara = vStrPara & "'" & IIf(Trim(.Columns("ColourID").Text) = "", Null, Val(.Columns("ColourID").Text)) & "'," 'HeaderStoreID
        Else
           vStrPara = vStrPara & "''," 'HeaderStoreID
        End If
        vStrPara = vStrPara & 0 & "," ' Disc Amount
        vStrPara = vStrPara & "Null" & "," ' isLastPrice
        vStrPara = vStrPara & "Null" & ","   'Re SPrice
        vStrPara = vStrPara & "Null" & ""   'Re SAmount
        vStrPara = Replace(vStrPara, "''", "Null")
        vStrPara = "Exec SaleReturnBodyInsert " & vStrPara
        CN.Execute vStrPara
      End If
      .MoveNext
   Next vCounter
   .RemoveAll
   .Redraw = True
End With

End Sub
Private Sub SaveSaleOut()
   Dim vInvoiceNo, vComission, vBankMachineID, vCashReceived, vBankAmount, vCustomerID, vCustomerName As String
   If FrmReplacePrint.OptBankCard.Value = True Then
         vInvoiceNo = FrmReplacePrint.TxtInvoiceNo.Text
         vComission = FrmReplacePrint.TxtCommision.Text
         vBankMachineID = FrmReplacePrint.TxtBankMachineID.Text
         vCashReceived = 0
         vCustomerID = "621"
         vCustomerName = IIf(Trim(FrmReplacePrint.TxtBankCustomer.Text) = "", Null, FrmReplacePrint.TxtBankCustomer.Text)
         vCashReceived = Val(FrmReplacePrint.TxtCashReceivedBank.Text)
      End If
      If FrmReplacePrint.OptCash.Value = True Then
         vComission = "''"
         vInvoiceNo = Null
         vBankMachineID = Null
         vCashReceived = Val(FrmReplacePrint.TxtCashReceivedCash.Text)
         vCustomerID = "621"
         vCustomerName = IIf(Trim(FrmReplacePrint.TxtCashCustomer.Text) = "", Null, FrmReplacePrint.TxtCashCustomer.Text)
      End If
      If FrmReplacePrint.OptCredit.Value = True Then
         vComission = "''"
         vInvoiceNo = Null
         vBankMachineID = Null
         If LblAmount.Caption = "Cash Paid" Then
            vCashReceived = 0
            vBankAmount = 0
         Else
            vCashReceived = Val(FrmReplacePrint.TxtCashReceivedCredit.Text)
            vBankAmount = Val(FrmReplacePrint.TxtBankAmount.Text)
         End If
         vCustomerID = FrmReplacePrint.TxtCustomerID.Text
         vBankMachineID = IIf(Trim(FrmReplacePrint.TxtBankMachineCreditID.Text) = "", Null, FrmReplacePrint.TxtBankMachineCreditID.Text)
'         !BankAmount = Val(FrmReplacePrint.TxtBankAmount.Text)
         If Val(FrmReplacePrint.TxtBankMachineCreditID.Text) > 0 Then
            vComission = Val(FrmReplacePrint.TxtCommision.Text)
         Else
            vComission = "''"
         End If
         vCustomerName = "''"
      End If
   
   vStrPara = ""
   vStrPara = Abs(ObjRegistry.AllowContinuousBillNo) & ","
   vStrPara = vStrPara & Abs(ObjRegistry.AllowMonthlyBillNo) & ","
   vStrPara = vStrPara & Abs(ObjRegistry.AllowDailyBillNo) & "," 'AllowDailyBillNo
   vStrPara = vStrPara & Val(TxtSSID.Text) & "," 'SID
   vStrPara = vStrPara & TxtBillID.Text & ","
   vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
   vStrPara = vStrPara & "'" & vCustomerID & "'," 'CustomerID
   'vStrPara = vStrPara & SelfRound(vTotalAmount) & "," ' Total Amount
   vStrPara = vStrPara & SelfRound(TxtSNetAmount.Text + Val(TxtSBillDisc.Text) - Val(TxtSServiceCharges.Text)) & ","     ' Total Amount
   vStrPara = vStrPara & Val(TxtSBillDisc.Text) & "," 'BillDisc
   vStrPara = vStrPara & vCashReceived & "," ' 'CashReceived
   vStrPara = vStrPara & vUser & "," 'UserNo
   vStrPara = vStrPara & TxtStoreID.Text & "," 'StoreID
   vStrPara = vStrPara & IIf(FrmReplacePrint.OptBankCard.Value = True, 1, 0) & "," 'BankCard
   vStrPara = vStrPara & IIf(FrmReplacePrint.OptCredit.Value = True, 1, 0) & "," 'Credit
   vStrPara = vStrPara & IIf(FrmReplacePrint.OptCash.Value = True, 1, 0) & "," 'Cash
   vStrPara = vStrPara & "'" & vBankMachineID & "'," 'BankMachineID
   vStrPara = vStrPara & "'" & vInvoiceNo & "',"  'InvoiceNo
   vStrPara = vStrPara & "'" & vCustomerName & "'," 'CustomerName
   vStrPara = vStrPara & Val(TxtSBillDiscPer.Text) & "," 'BillDiscPer
   vStrPara = vStrPara & vComission & ","   'Commision
   vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtCommission.Text)) & "," 'EmpComm
   vStrPara = vStrPara & "'" & IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text)) & "'," 'EmpID
   vStrPara = vStrPara & 1 & "," 'isReplace
   vStrPara = vStrPara & 0 & "," 'isPosted
   vStrPara = vStrPara & IIf(Trim(TxtMemberID.Text) = "", "''", TxtMemberID.Text) & "," 'MemberID
   vStrPara = vStrPara & "'" & Now & "'," 'BillTime
   vStrPara = vStrPara & "'" & vIsNewRecord & "'," 'Tag
   vStrPara = vStrPara & "'" & IIf(Trim(TxtManualBillNo.Text) = "", Null, TxtManualBillNo.Text) & "'," 'ManualBillNo
   vStrPara = vStrPara & "'" & Null & "',"   'Remarks
   vStrPara = vStrPara & IIf(Trim(TxtOrganizationID.Text) = "", "''", TxtOrganizationID.Text) & ","  'OrganizationID
   vStrPara = vStrPara & "'" & Null & "'," ' BillNo
   vStrPara = vStrPara & "'" & Null & "'," ' Bilty No
   vStrPara = vStrPara & "'" & Null & "'," 'Description
   vStrPara = vStrPara & "''" & "," 'PAIDAMOUNT
   vStrPara = vStrPara & "'" & Null & "',"  'EntryDate
   vStrPara = vStrPara & IIf(FrmReplacePrint.OptCredit.Value = True, 0, 0) & "," 'PreviousAmount
   vStrPara = vStrPara & 0 & "," 'OtherCharges
   vStrPara = vStrPara & "'" & Null & "'," 'SaleManID
   vStrPara = vStrPara & 0 & "," 'TotalExpense
   vStrPara = vStrPara & "''" & "," 'OrderID
   vStrPara = vStrPara & "'" & "'," 'OrderDate
   vStrPara = vStrPara & 0 & "," 'Freight
   vStrPara = vStrPara & 0 & "," 'IsCustomerFreight
   vStrPara = vStrPara & "'" & Null & "'," 'VechicleNo
   vStrPara = vStrPara & IIf(TxtSServiceCharges.Text = "", "''", Val(TxtSServiceCharges.Text)) & "," 'ServiceCharges
   vStrPara = vStrPara & "''" & ","   'ServiceChargesPer
   vStrPara = vStrPara & "''" & "," 'STax
   vStrPara = vStrPara & "''" & "," 'STaxPer
   vStrPara = vStrPara & "'" & Null & "'," 'TableID
   vStrPara = vStrPara & "'" & Now & "'," 'ServerEntry
   vStrPara = vStrPara & "'" & Null & "'," 'InvType
   vStrPara = vStrPara & "'" & Null & "'," 'DeliveryDate
   vStrPara = vStrPara & "'" & Null & "'," 'DeliveryTime
   vStrPara = vStrPara & "'" & Null & "'," 'isPrinted
   vStrPara = vStrPara & "'" & Null & "'," 'RemarksUrdu
   'vStrPara = vStrPara & "Default" & ","  'StampID
   vStrPara = vStrPara & 0 & "," 'isTransfer
   vStrPara = vStrPara & "Null," 'PromiseDate
   vStrPara = vStrPara & "Null," 'Expiry Invoice
   vStrPara = vStrPara & "Null," 'Syllabus
   vStrPara = vStrPara & "'" & IIf(Trim(vSessionID) = 0, Null, Val(vSessionID)) & "',"  'vSessionID
   vStrPara = vStrPara & "''" & "," 'AdvTaxVal
   vStrPara = vStrPara & "Null," 'AdvTaxPer
   vStrPara = vStrPara & "Null," 'ExtraTaxVal
   vStrPara = vStrPara & "Null," 'ExtraTaxPer
   vStrPara = vStrPara & "Null," 'CNIC
   vStrPara = vStrPara & "Null," 'CellNo
   vStrPara = vStrPara & "0," 'Sum Disc Amount
   vStrPara = vStrPara & "Null," 'DispatchDate
   vStrPara = vStrPara & "Null," 'Terms
   vStrPara = vStrPara & "Null," 'RefID
   vStrPara = vStrPara & "Null,"  'Refcomm
   vStrPara = vStrPara & Val(vBankAmount) 'Bank Amount in Credit Option
   vStrPara = Replace(vStrPara, "''", "Null")
   
   vStrPara = "DECLARE @returnvalue INT EXEC @returnvalue = saleheaderinsert " & vStrPara & " Select @returnvalue"
      vMasterID = CN.Execute(vStrPara).Fields(0).Value
      TxtSSID.Text = vMasterID
   '   MsgBox vMasterID
      
   '============================================
   '   Start backup entry Master
   '============================================
      Dim vMasterID1 As Long
            
'      If vConnStr <> "" Then
'         vMasterID1 = Cnn.Execute(vStrPara).Fields(0).Value
'      End If
   '============================================
   '   End backup entry Master
   '============================================
      
   
   '/******* FBR Integeration*************/
'      If vPOSID <> "" Then
'         If ObjRegistry.AllowFBRContinuousNo Then
'            vUSIN = cn.Execute("select isnull(max(USIN),0) + 1 as USIN from SaleHeader").Fields(0).Value
'         Else
'            vUSIN = TxtSID.Text
'         End If
'         vHeader = "{InvoiceNumber:'',POSID:'" & vPOSID & "',DateTime:'" & Replace(DtpBillDate.Date, "/", "-") & "',BuyerName:'" & FrmReplacePrint.TxtCashCustomer.Text & "',TotalQuantity:" & Val(TxtTotSaleQty.Text) & ",TotalSaleValue:" & Val(TxtSNetAmount.Text) - Val(TxtTotalSaleTaxValue.Text) + Val(TxtTotalDiscount.Caption) & ",Totaltaxcharged:" & Val(TxtTotalSaleTaxValue.Text) & ",Discount:" & Val(TxtTotalDiscount.Caption) & ",TotalBillAmount:" & Val(TxtNetAmount.Caption) + Val(TxtTotalDiscount.Caption) & ",PaymentMode:1,InvoiceType:1,USIN:'" & vUSIN & "', items : ["
'      End If
   ''''''''''''''''''''''''''''
    
       
   ''' insert Sale Body
   vStrDetail = ""
   vGridRows = 0
   vProducts = ""
   i = 0
   vSamePid = ""
   With GridSale
    .Redraw = False
    .MoveFirst
      For vCounter = 1 To .rows
         If Trim(.Columns("Productid").Text) <> "" Then
         
         '''''' ActivityLogBin Follwoin lines check the same product id which was enter seperate row or new new row
         If (InStr(1, vSamePid, .Columns("Productid").Text)) = 0 Then vGridRows = vGridRows + 1
         vSamePid = vSamePid & " , " & .Columns("Productid").Text
         '''''''''''''''''''''''''''''''''''''
         
       ''''''''''''''''''''''''''''
    vStrPara = ""
   TxtBillID.Text = CN.Execute("Select billID from Saleheader where SID = " & vMasterID).Fields(0).Value
   vStrPara = vStrPara & "'" & 0 & "'," 'check stock update or not
   vStrPara = vStrPara & vMasterID & ","
   vStrPara = vStrPara & TxtBillID.Text & ","
   vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
   'vStrPara = vStrPara & .Columns("SerialNo").Text & ","
   'vStrPara = vStrPara & .Columns("BillID").Text & ","
   'vStrPara = vStrPara & .Columns("BillDate").Text & ","
   vStrPara = vStrPara & .Columns("ProductID").Text & ","
   vStrPara = vStrPara & .Columns("Qty").Text & ","
   vStrPara = vStrPara & .Columns("Price").Text & ","
   vStrPara = vStrPara & .Columns("DiscPC").Text & ","
   vStrPara = vStrPara & .Columns("Amount").Text & ","
   vStrPara = vStrPara & "'" & .Columns("Code").Text & "',"
   vStrPara = vStrPara & .Columns("DiscPer").Text & ","
   vStrPara = vStrPara & .Columns("DiscVal").Text & ","
   
   vStrPara = vStrPara & 0 & "," ' isDiscB4TradeOffer
   vStrPara = vStrPara & 0 & ","   'isDiscB4ExtraScheme
   vStrPara = vStrPara & 0 & "," 'isDiscB4SaleTax
   vStrPara = vStrPara & "''" & ","  'TradeOffer1
   vStrPara = vStrPara & "''" & ","   'TradeOffer2
   vStrPara = vStrPara & "''" & ","   'ExtraSchemePer
   vStrPara = vStrPara & "''" & ","   'TradeValue
   vStrPara = vStrPara & "''" & ","   'ExtraSchemeValue
   
   vStrPara = vStrPara & .Columns("Cost").Text & ","
   vStrPara = vStrPara & .Columns("isProduct").Text & ","
   vStrPara = vStrPara & "''" & "," ' Pack Name
   vStrPara = vStrPara & "''" & "," ' Qty Pack
   vStrPara = vStrPara & "''" & "," ' Pack
   vStrPara = vStrPara & "''" & "," ' Bonus
   vStrPara = vStrPara & "''" & "," 'Offer
   vStrPara = vStrPara & 0 & ","  'SaleTaxPer
   vStrPara = vStrPara & 0 & ","  ' SaleTaxVal
   vStrPara = vStrPara & 0 & ","
   vStrPara = vStrPara & Val(TxtSPrice.Text) & "," 'RetailPrice
   vStrPara = vStrPara & 0 & "," 'IsWSSaleTax
   vStrPara = vStrPara & 0 & ","  'IsRetailSaleTax
   vStrPara = vStrPara & 0 & ","  'IsWSDiscb4ST
   vStrPara = vStrPara & Val(.Columns("SC").Text) & "," 'SC
   vStrPara = vStrPara & Val(.Columns("EmpComm").Value & ",") & ","  'EmpComm
   vStrPara = vStrPara & "''" & "," 'BatchNo
   'vStrPara = vStrPara & "''" & "," 'StampID
   vStrPara = vStrPara & TxtStoreID.Text & ","                  'StoreID
   If ObjRegistry.AllowEmployeProductWise Then
      vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtEmployeeID.Text)) & "," 'EmpID
   Else
      vStrPara = vStrPara & "''" & "," 'EmpID
   End If
   vStrPara = vStrPara & "'" & IIf(Trim(.Columns("ColourID").Text) = "", Null, Val(.Columns("ColourID").Text)) & "'," ' ColourID
   vStrPara = vStrPara & "'" & IIf(Trim(.Columns("SizeID").Text) = "", Null, Val(.Columns("SizeID").Text)) & "'," ' SizeID
   vStrPara = vStrPara & "null" & ","  'Gross Qty
   vStrPara = vStrPara & "null" & ","  'Gross Unit
   If ObjRegistry.AllowStoreProductWise Then
      vStrPara = vStrPara & "'" & IIf(Trim(.Columns("ColourID").Text) = "", Null, Val(.Columns("ColourID").Text)) & "'," 'HeaderStoreID
   Else
      vStrPara = vStrPara & "''," 'HeaderStoreID
   End If
   vStrPara = vStrPara & 0 & "," ' Disc Amount
   vStrPara = vStrPara & "Null" & "," ' isLastPrice
   vStrPara = vStrPara & "Null" & ","   'Re SPrice
   vStrPara = vStrPara & "Null" & ""   'Re SAmount
   vStrPara = Replace(vStrPara, "''", "Null")
   CN.Execute ("Exec SaleBodyInsert " & vStrPara)

   
'============================================
'   end backup entry Body
'============================================



   vStrDetail = vStrDetail & " (P" & .Columns("ProductID").Text & " Q" & .Columns("Qty").Text & " A" & .Columns("Amount").Text & ")"

   '/******* FBR Integeration*************/
   
'   If vPOSID <> "" Then
'      vStrSQL = "Select isnull(p.PCTCode,isnull(g.PCTCode,'01011000')) from Groups g inner join products p on p.groupid = g.groupid where p.productid = " & Val(.Columns("ProductID").Text)
'      vSQL = "Select case when isnull(Is3rdScheduleItem,0) = 1 then 11 else 1 end from products where productid = " & Val(.Columns("ProductID").Text)
'      vProducts = vProducts + "{itemcode:'" & .Columns("ProductID").Text & "', itemname:'" & Replace(.Columns("ProductName").Text, "'", "") & "',PCTCODE:'" & cn.Execute(vStrSQL).Fields(0).Value & "',quantity:" & .Columns("Qty").Text & ",taxrate:" & .Columns("SaleTaxPer").Value & ",SaleValue:" & Val(.Columns("Amount").Value) - Val(.Columns("SaleTaxVal").Value) + Val(.Columns("DiscVal").Value) & ",Discount:" & Val(.Columns("DiscVal").Value) & ",taxcharged:" & Val(.Columns("SaleTaxVal").Value) & ",totalamount:" & Val(.Columns("Amount").Value) + Val(.Columns("DiscVal").Value) & ",InvoiceType:" & cn.Execute(vSQL).Fields(0).Value & "},"
'   End If
   ''''''''''''''''''''''''''''
      
      End If
      .MoveNext
   Next vCounter
   .RemoveAll
   .Redraw = True
   End With
End Sub
Private Sub SaveReplacement()
   vStrPara = ""
   vStrPara = Abs(ObjRegistry.AllowContinuousBillNo) & ","
   vStrPara = vStrPara & Abs(ObjRegistry.AllowMonthlyBillNo) & ","
   vStrPara = vStrPara & Abs(ObjRegistry.AllowDailyBillNo) & "," 'AllowDailyBillNo
    vStrPara = vStrPara & Val(TxtSID.Text) & "," 'SID
   vStrPara = vStrPara & Val(TxtSSID.Text) & "," 'SSID
   vStrPara = vStrPara & Val(TxtRSID.Text) & "," 'RSID
   vStrPara = vStrPara & TxtReplaceID.Text & "," 'ReplaceID
   vStrPara = vStrPara & "'" & DtpReplaceDate.DateValue & "'," 'ReplaceDate
   vStrPara = vStrPara & TxtBillID.Text & "," 'BillID
   vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "'," 'BillDate
   vStrPara = vStrPara & TxtReturnID.Text & "," 'ReturnID
   vStrPara = vStrPara & "'" & DtpReturnDate.DateValue & "'," 'ReturnDate
   vStrPara = vStrPara & "'" & Val(TxtSNetAmount.Text) & "'," 'SaleAmount
   vStrPara = vStrPara & "'" & Val(TxtRNetAmount.Text) & "'," 'ReturnAmount
  
   If FrmReplacePrint.OptCredit = True Then FrmReplacePrint.TxtCashReceivedCash.Text = FrmReplacePrint.TxtCashReceivedCredit.Text
   vStrPara = vStrPara & "'" & Val(FrmReplacePrint.TxtCashReceivedCash.Text) & "'," 'Cash Received
   
   vStrPara = vStrPara & "'" & vUser & "',"  'user No
   vStrPara = vStrPara & "0,"  'isPosted
   vStrPara = vStrPara & "'" & vIsNewRecord & "',"   'Tag
   vStrPara = vStrPara & "0,"  'isTransfer
   vStrPara = vStrPara & TxtStoreID.Text & "," 'StoreID
   vStrPara = vStrPara & vSessionID & "," 'SessionID
   vStrPara = vStrPara & "0"  'isSync
   
   
   vStrPara = Replace(vStrPara, "''", "Null")
   
   vStrPara = "DECLARE @returnvalue INT EXEC @returnvalue = ReplacementHeaderInsert " & vStrPara & " Select @returnvalue"
   vMasterID = CN.Execute(vStrPara).Fields(0).Value
   TxtSID.Text = vMasterID
End Sub
