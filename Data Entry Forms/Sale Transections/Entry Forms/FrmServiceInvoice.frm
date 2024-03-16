VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmServiceInvoice 
   BorderStyle     =   0  'None
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmServiceInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1808
      MaxLength       =   50
      TabIndex        =   75
      Top             =   9728
      Visible         =   0   'False
      Width           =   4125
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   12833
      Top             =   8993
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
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
      Height          =   4665
      Left            =   13920
      TabIndex        =   66
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
         Height          =   4200
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   67
         Tag             =   "NC"
         Text            =   "FrmServiceInvoice.frx":0ECA
         Top             =   300
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
         TabIndex        =   68
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.CheckBox ChkIsProduct 
      Caption         =   "Is Product"
      Height          =   255
      Left            =   7118
      TabIndex        =   51
      Top             =   1388
      Visible         =   0   'False
      Width           =   1050
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   1883
      TabIndex        =   18
      Top             =   2198
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
      Left            =   10898
      TabIndex        =   9
      Top             =   2918
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   1928
      TabIndex        =   4
      Top             =   2918
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
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   7643
      TabIndex        =   5
      Top             =   2918
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
      Left            =   8423
      TabIndex        =   6
      Top             =   2918
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
      Left            =   11888
      TabIndex        =   19
      Top             =   2918
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
      Left            =   3788
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2918
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
      MICON           =   "FrmServiceInvoice.frx":1041
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9098
      TabIndex        =   16
      Top             =   9113
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
      MICON           =   "FrmServiceInvoice.frx":105D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7778
      TabIndex        =   12
      Top             =   9113
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
      MICON           =   "FrmServiceInvoice.frx":1079
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5138
      TabIndex        =   14
      Top             =   9113
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
      MICON           =   "FrmServiceInvoice.frx":1095
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10418
      TabIndex        =   17
      Top             =   9113
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
      MICON           =   "FrmServiceInvoice.frx":10B1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6458
      TabIndex        =   13
      Top             =   9113
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
      MICON           =   "FrmServiceInvoice.frx":10CD
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillDisc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1913
      TabIndex        =   10
      Top             =   7463
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
      Left            =   4148
      TabIndex        =   30
      Top             =   2918
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
      Left            =   1928
      TabIndex        =   31
      Top             =   3233
      Width           =   11625
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
      stylesets(0).Picture=   "FrmServiceInvoice.frx":10E9
      AllowUpdate     =   0   'False
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
      RowHeight       =   503
      ExtraHeight     =   397
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   10
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Sr"
      Columns(0).Name =   "Sr"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3916
      Columns(1).Caption=   "Product ID"
      Columns(1).Name =   "ProductID"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   6165
      Columns(2).Caption=   "Product Name"
      Columns(2).Name =   "ProductName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1376
      Columns(3).Caption=   "Qty"
      Columns(3).Name =   "Qty"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   4
      Columns(3).FieldLen=   256
      Columns(4).Width=   1693
      Columns(4).Caption=   "Price"
      Columns(4).Name =   "Price"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   4
      Columns(4).FieldLen=   256
      Columns(5).Width=   1455
      Columns(5).Caption=   "Disc / Pc"
      Columns(5).Name =   "DiscPC"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1217
      Columns(6).Caption=   "Disc%"
      Columns(6).Name =   "DiscPer"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1746
      Columns(7).Caption=   "Disc. Val"
      Columns(7).Name =   "DiscVal"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   2461
      Columns(8).Caption=   "Amount"
      Columns(8).Name =   "Amount"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   5
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "TotalAmount"
      Columns(9).Name =   "TotalAmount"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
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
      Left            =   3818
      TabIndex        =   15
      Top             =   9113
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
      MICON           =   "FrmServiceInvoice.frx":1105
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   9383
      TabIndex        =   7
      Top             =   2918
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
      Left            =   9938
      TabIndex        =   37
      Top             =   1313
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
      Left            =   4283
      TabIndex        =   1
      Tag             =   "NC"
      Top             =   2198
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
      Left            =   5318
      TabIndex        =   39
      Tag             =   "NC"
      Top             =   2198
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   4958
      TabIndex        =   40
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
      MICON           =   "FrmServiceInvoice.frx":1121
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpBillDate 
      Height          =   315
      Left            =   2843
      TabIndex        =   0
      Top             =   2198
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
      Left            =   10208
      TabIndex        =   8
      Top             =   2918
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
      Left            =   11528
      TabIndex        =   44
      Top             =   1313
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtCost 
      Height          =   315
      Left            =   9113
      TabIndex        =   47
      Top             =   1298
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
   Begin SITextBox.Txt TxtBillDiscPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1913
      TabIndex        =   11
      Top             =   8168
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
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   6773
      TabIndex        =   2
      Top             =   2198
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
      Left            =   7883
      TabIndex        =   58
      Top             =   2198
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
      Left            =   7523
      TabIndex        =   59
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
      MICON           =   "FrmServiceInvoice.frx":113D
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCommission 
      Height          =   315
      Left            =   8288
      TabIndex        =   62
      Top             =   1298
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
   Begin SITextBox.Txt TxtMemberID 
      Height          =   315
      Left            =   9458
      TabIndex        =   3
      Top             =   2198
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
   End
   Begin SITextBox.Txt TxtMemberName 
      Height          =   315
      Left            =   10493
      TabIndex        =   70
      Top             =   2198
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
      Left            =   10133
      TabIndex        =   71
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
      MICON           =   "FrmServiceInvoice.frx":1159
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
      Height          =   225
      Left            =   1823
      TabIndex        =   76
      Top             =   9488
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label LblNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1853
      TabIndex        =   74
      Top             =   9218
      Width           =   165
   End
   Begin VB.Label LblMemberName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      Height          =   195
      Left            =   10493
      TabIndex        =   73
      Top             =   1973
      Width           =   1035
   End
   Begin VB.Label LblMemberID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      Height          =   195
      Left            =   9458
      TabIndex        =   72
      Top             =   1973
      Width           =   780
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
      Left            =   13058
      TabIndex        =   69
      Top             =   1568
      Width           =   435
   End
   Begin VB.Label LblCaptionPrice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   11753
      TabIndex        =   65
      Top             =   1343
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label LblPrice 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   11978
      TabIndex        =   64
      Top             =   1658
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      Height          =   195
      Left            =   8153
      TabIndex        =   63
      Top             =   1073
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   7883
      TabIndex        =   61
      Top             =   1973
      Width           =   780
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   6773
      TabIndex        =   60
      Top             =   1973
      Width           =   525
   End
   Begin VB.Label TxtTotalQty 
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
      ForeColor       =   &H0000FFFF&
      Height          =   915
      Left            =   5558
      TabIndex        =   57
      Top             =   7643
      Width           =   1380
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
      Height          =   915
      Left            =   6968
      TabIndex        =   56
      Top             =   7643
      Width           =   2370
   End
   Begin VB.Label TxtTotalDiscount 
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
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   9368
      TabIndex        =   55
      Top             =   7643
      Width           =   1740
   End
   Begin VB.Label TxtNetAmount 
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
      Height          =   915
      Left            =   11138
      TabIndex        =   54
      Top             =   7643
      Width           =   2370
   End
   Begin VB.Label TxtLastRate 
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
      ForeColor       =   &H0000FF00&
      Height          =   915
      Left            =   3743
      TabIndex        =   53
      Top             =   7643
      Width           =   1785
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Discount (%)"
      Height          =   195
      Left            =   1913
      TabIndex        =   52
      Top             =   7943
      Width           =   1125
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Invoice"
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
      Width           =   2610
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
      Left            =   2903
      TabIndex        =   49
      Top             =   2663
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   9143
      TabIndex        =   48
      Top             =   1073
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
      Left            =   3743
      TabIndex        =   46
      Top             =   7313
      Width           =   1830
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   11573
      TabIndex        =   45
      Top             =   1118
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. %"
      Height          =   195
      Left            =   10193
      TabIndex        =   43
      Top             =   2723
      Width           =   525
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   4283
      TabIndex        =   42
      Top             =   1973
      Width           =   585
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   5318
      TabIndex        =   41
      Top             =   1973
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Amount"
      Height          =   195
      Left            =   9953
      TabIndex        =   38
      Top             =   1103
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
      Left            =   9368
      TabIndex        =   36
      Top             =   7343
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
      Left            =   7178
      TabIndex        =   35
      Top             =   7343
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   8423
      TabIndex        =   34
      Top             =   2723
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
      Left            =   5603
      TabIndex        =   33
      Top             =   7313
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
      Left            =   4148
      TabIndex        =   32
      Top             =   2723
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
      Left            =   11663
      TabIndex        =   29
      Top             =   7343
      Width           =   1440
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Discount"
      Height          =   195
      Left            =   1913
      TabIndex        =   28
      Top             =   7238
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   2843
      TabIndex        =   27
      Top             =   1973
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " Bill ID"
      Height          =   195
      Left            =   1883
      TabIndex        =   26
      Top             =   1973
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      Height          =   195
      Left            =   1928
      TabIndex        =   25
      Top             =   2723
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc / PC"
      Height          =   195
      Left            =   9383
      TabIndex        =   24
      Top             =   2723
      Width           =   690
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   7643
      TabIndex        =   23
      Top             =   2723
      Width           =   240
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   11888
      TabIndex        =   22
      Top             =   2723
      Width           =   540
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. Val"
      Height          =   195
      Left            =   10898
      TabIndex        =   21
      Top             =   2723
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
Attribute VB_Name = "FrmServiceInvoice"
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
Dim vFlag As Boolean
Dim vBm As Variant
Dim DateFlag As Boolean
'Dim vCurrentDate As Date
Dim ssql As String
Dim VStrSQL As String
Dim vQtyLoose As Double, vTotalAmount As Double
Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String, vTotDisc As Double
Dim i As Integer, vCashDrawer As Boolean, vLaserInvoice As Boolean, vPrintHeader  As Boolean, vNoofPrints As Byte, vX As Integer, vY As Integer
'----------------------------------

Private Sub SubCalculateBody()
    TxtDiscVal.Text = Val(TxtQty.Text) * Val(TxtDiscPC.Text)
    TxtActualAmount.Text = Val(TxtQty.Text) * Val(TxtPrice.Text)
    TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
    TxtTotalDiscount.Caption = vTotDisc
    SubCalculateFooter
End Sub

Private Sub SubApplyMember()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   ssql = " select DiscInPer as DiscPer " & vbCrLf _
         + " from MemberTypes t inner join Members m on t.MemberTypeID = m.MemberTypeID  where MemberID = " & TxtMemberID.Text
   With CN.Execute(ssql)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         RsBody.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
         If RsBody.RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
            'RsBody.Filter = "ProductID='" & !ProductID & "'"
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
            Grid.Columns("DiscPC").Value = Round((Val(RsBody!Price) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) - Val(Grid.Columns("DiscVal").Value)
            
            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - RsBody!Amount + Val(Grid.Columns("Amount").Text)
            vTotDisc = vTotDisc - RsBody!DiscVal + Val(Grid.Columns("DiscVal").Text)
            vTotalAmount = vTotalAmount - RsBody!Amount + Val(Grid.Columns("Amount").Text)
            
            RsBody!DiscPC = Val(Grid.Columns("DiscPC").Value)
            RsBody!DiscPer = Val(Grid.Columns("DiscPer").Value)
            RsBody!DiscVal = Val(Grid.Columns("DiscVal").Value)
            RsBody!Amount = Val(Grid.Columns("Amount").Value)
     
         End If
         Grid.MoveNext
      Wend
      .Close
   End With
   Grid.MoveLast
   SubCalculateFooter
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubDestroyMember()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
      
      While Trim(Grid.Columns("ProductID").Text) <> ""
         RsBody.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
         If RsBody.RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
            
            'RsBody.Filter = "ProductID='" & !ProductID & "'"
            Grid.Columns("DiscPer").Value = 0 'IIf(IsNull(!DiscPer), 0, !DiscPer)
            Grid.Columns("DiscPC").Value = 0 'Round((Val(RsBody!Price) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = 0 'Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) - Val(Grid.Columns("DiscVal").Value)
            
            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - RsBody!Amount + Val(Grid.Columns("Amount").Text)
            vTotDisc = vTotDisc - RsBody!DiscVal + Val(Grid.Columns("DiscVal").Text)
            vTotalAmount = vTotalAmount - RsBody!Amount + Val(Grid.Columns("Amount").Text)
            
            RsBody!DiscPC = Val(Grid.Columns("DiscPC").Value)
            RsBody!DiscPer = Val(Grid.Columns("DiscPer").Value)
            RsBody!DiscVal = Val(Grid.Columns("DiscVal").Value)
            RsBody!Amount = Val(Grid.Columns("Amount").Value)
         End If
         Grid.MoveNext
      Wend
   SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateFooter()
   TxtTotalDiscount.Caption = Val(TxtBillDisc.Text) + vTotDisc
   TxtNetAmount.Caption = SelfRound(Val(TxtTotalAmount.Caption) - Val(TxtTotalDiscount.Caption))
   'If TxtGrossAmount.Text = "" Then Exit Sub
   'TxtNetAmount.Caption = Round(Val(TxtGrossAmount.Text)) - Val(TxtBillDisc.Text)
   'TxtCashReturn.Text = IIf(Val(TxtCashReceived.Text) > 0, Val(TxtCashReceived.Text) - Val(TxtNetAmount.Caption), "")
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
        TxtEmployeeName.Text = !EmpName
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
            + " where IsLockMember = 0 and MemberID = " & Val(TxtMemberID.Text)
    With CN.Execute(ssql)
      If .RecordCount > 0 Then
        TxtMemberName.Text = !MemberName
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
      SchSProduct.ParaInWhere = " and isLocked = 0"
      SchSProduct.Show vbModal, Me
      vFlag = False
      If SchSProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtProductID.Text = SchSProduct.ParaOutID
   End If
    '---------------------------
    If TxtProductID.Enabled = False Then FunSelectProduct = False: Exit Function
    If Trim(TxtProductID.Text) = "" Then FunSelectProduct = False: Exit Function
    If Len(TxtProductID.Text) < 5 Then
      TxtProductID.Text = Right("00000" + CStr(Val(TxtProductID.Text)), 5)
    End If
    If TxtProductID.Text = "" Then FunSelectProduct = False: Exit Function
    
   ''''''''***********   Checking Product  ***********''''''''
    VStrSQL = " SELECT p.productid, ProductName, RetailPrice, DiscPer, DiscPC" & vbCrLf _
           + " from SProducts p " & vbCrLf _
           + " where isLocked = 0 and p.productid = '" & TxtProductID.Text & "'"

   With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !RetailPrice
         TxtQty.Text = IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtDiscPC.Text) <> 0 Then
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
         End If
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
         TxtProductID.Text = ""
         TxtProductName.Text = ""
         TxtPrice.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtAmount.Text = ""
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
   If MsgBox("Are you sure to Clear the Data?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Exit Sub
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
   Call ActivityLog("Service Invoice", eDelete, TxtBillID.Text, DtpBillDate.DateValue)
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("ProductID").Text) <> "" Then
         CN.Execute "Delete from ServiceBody where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "' and ProductID ='" & Grid.Columns("Productid").Text & "'"
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   CN.Execute "Delete from ServiceHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnEmployee_Click()
   On Error GoTo ErrorHandler
   If FunSelectEmployee(ssButton, False) = True Then
      If TxtMemberID.Visible = True Then TxtMemberID.SetFocus Else If TxtProductID.Enabled Then TxtProductID.SetFocus
   Else
      TxtEmployeeID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnMember_Click()
   On Error GoTo ErrorHandler
   If FunSelectMember(ssButton, False) = True Then
      If TxtProductID.Enabled Then TxtProductID.SetFocus
   Else
      TxtMemberID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchService.ParaInBillDate = DtpBillDate.DateValue
   SchService.Show vbModal
   If SchService.ParaOutBillID <> -1 Then
      TxtBillID.Text = SchService.ParaOutBillID
      'Dim a
      'a = Split(SchSale.ParaOutBillDate, "/")
      DtpBillDate.DateValue = SchService.ParaOutBillDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      GetSale
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   
   VStrSQL = " select username, h.billid, h.BillDate, isnull(h.BillTime,0) as BillTime, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.cashReceived,0) as CashReceived, p.ProductName as ProductName, '' as unitname, b.qty, b.price as price, b.amount, b.DiscVal, InvoiceNo" & vbCrLf _
            + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName End as Customer, Cash, Credit, BankCard, b.ProductID" & vbCrLf _
            + " from ServiceHeader h inner join ServiceBody b on h.billid = b.billid and h.BillDate = b.BillDate" & vbCrLf _
            + " inner join SProducts p on p.productid = b.productid" & vbCrLf _
            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
            + " left outer join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate ='" & DtpBillDate.DateValue & "' Order By SerialNo"

   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
   
   
    If vLaserInvoice = True Then
'      If vLaserInvoice = False Then
'         Set RptReportViewer.Report = New CrpSaleInvoice
'         RptReportViewer.Report.TopMargin = 225
'         RptReportViewer.Report.LeftMargin = 225
'         RptReportViewer.Report.RightMargin = 225
'         'RptReportViewer.Report.PaperOrientation = crPortrait
'      Else
         Set RptReportViewer.Report = New CrpSaleInvoiceHalf
         RptReportViewer.Report.TopMargin = vY
         RptReportViewer.Report.LeftMargin = vX
         RptReportViewer.Report.RightMargin = 225
'      End If
   Else
      If InStr(1, Printer.DeviceName, "CBM1000") > 0 Then
         Set RptReportViewer.Report = New CrpSaleInvoiceCBM
   '   ElseIf InStr(1, Printer.DeviceName, "Generic") > 0 Then
   '      Set RptReportViewer.Report = New CrpSaleInvoice
   '      RptReportViewer.Report.PaperSize = crPaperA4
   '
   '   ElseIf InStr(1, Printer.DeviceName, "Generic") > 0 Then
   '      Set RptReportViewer.Report = New CrpSaleInvoiceGeneric
   '      RptReportViewer.Report.PaperSize = crPaperEnvelope14
      ElseIf InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
         Set RptReportViewer.Report = New CrpSaleInvoiceAurora
         RptReportViewer.Report.LeftMargin = 225
         RptReportViewer.Report.RightMargin = 0
         RptReportViewer.Report.TopMargin = 255
      Else 'InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
         Set RptReportViewer.Report = New CrpSaleInvoiceAurora
         'RptReportViewer.Report.LeftMargin = 0
         'RptReportViewer.Report.RightMargin = 0
      End If
      'RptReportViewer.Report.PaperOrientation = crPortrait
    End If
    RptReportViewer.Report.DiscardSavedData
    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
    RptReportViewer.Report.ReportTitle = "Services Invoice"
    
    'RptReportViewer.Report.LeftMargin = 0
    'RptReportViewer.Report.RightMargin = 0
    vStrComp = "Select CompanyName,Address,City,PhoneNo,email from Company"
    With CN.Execute(vStrComp)
      If .RecordCount > 0 Then
         vCompanyName = !CompanyName
         vAddress = !Address
         vAddress = !Address & IIf(!City = "", "", IIf(!Address = "", "", ", ") & !City)
         vPhone = IIf(!PhoneNo = "", "", !PhoneNo)
         If vPrintHeader = True Then
            RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
            RptReportViewer.Report.ParameterFields(2).AddCurrentValue vAddress
            RptReportViewer.Report.ParameterFields(4).AddCurrentValue vPhone
         Else
            RptReportViewer.Report.ParameterFields(1).AddCurrentValue ""
            RptReportViewer.Report.ParameterFields(2).AddCurrentValue ""
            RptReportViewer.Report.ParameterFields(4).AddCurrentValue ""
         End If
      End If
   End With
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
   With CN.Execute("select * from registry")
      If .RecordCount > 0 Then
         RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(!AddSpace = True, ".", "")
         RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(!CashReceived)
         RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(!Statement)
         If InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
            RptReportViewer.Report.ParameterFields(8).AddCurrentValue IIf(!AddSpace = True, ".", "")
         Else
            RptReportViewer.Report.ParameterFields(8).AddCurrentValue ""
         End If
      End If
      .Close
   End With
   
   'RptReportViewer.Report.SelectPrinter "RASDD.DLL", "CBM1000 Partial Cut", "Com1" 'RptReportViewer.Report.SelectPrinter  "RASDD.DLL", "CBM1000 Partial Cut", "Com1"
   'RptReportViewer.Show
   RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProduct_Click()
   On Error GoTo ErrorHandler
   If FunSelectProduct(ssButton, True) = True Then
      TxtQty.SetFocus
   Else
      TxtProductID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
'   If DtpBillDate.Enabled = True Then
'      If FrmPrint.OptCash.Visible Then FrmPrint.OptCash.SetFocus
'      FrmPrint.SubClearFields
'   End If
   FrmPrint.TxtNetAmount.Text = TxtNetAmount.Caption
   With CN.Execute("select * from registry")
      If .RecordCount > 0 Then
         If !CashReceived = False Then
            FrmPrint.TxtCashReceivedCash.Text = TxtNetAmount.Caption
         End If
      End If
      .Close
   End With
   FrmPrint.ParaInPrint = True
   FrmPrint.ParaInChoice = "Cash"
   FrmPrint.ParaInDate = DtpBillDate.DateValue
   FrmPrint.Show vbModal, Me
   If FrmPrint.ParaOutSelection = False Then Exit Sub
   If DtpBillDate.Enabled And DtpBillDate.Date <> IIf(Format(Now, "hh") > 2, Date, DateAdd("d", -1, Date)) And DateFlag = True Then
      If MsgBox("Are you sure to Change Bill Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
         DtpBillDate.DateValue = IIf(Format(Now, "hh") > 2, Date, DateAdd("d", -1, Date))
         TxtBillID.Text = FunGetMaxID()
      End If
      DateFlag = False
   End If
   If DtpBillDate.Enabled Then
      If CN.Execute("Select * from ServiceHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'").RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtBillID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Service", vbExclamation, "Alert"
      If TxtProductID.Visible And TxtProductID.Enabled Then TxtProductID.SetFocus
      Exit Sub
   End If
  'Body Validation
  ' validation has been performed when a row is added to the grid

  'Saving record
    CN.BeginTrans
    If vIsNewRecord = False Then Call ActivityLog("Service Invoice", eEdit, TxtBillID.Text, DtpBillDate.DateValue)
    ssql = "select * from ServiceHeader where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
    Dim Rs As New ADODB.Recordset
    With Rs
      .Open ssql, CN, adOpenStatic, adLockOptimistic
      If .BOF Then
         .AddNew
         !BillID = Val(TxtBillID.Text)
         !BillDate = DtpBillDate.DateValue
         !BillTime = Now
      End If
      !isPosted = 0
      !StoreID = TxtStoreID.Text
      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
      !EmpComm = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text))
      !MemberID = IIf(Trim(TxtMemberID.Text) = "", Null, TxtMemberID.Text)
      !TotalAmount = SelfRound(vTotalAmount)
      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
      If FrmPrint.OptBankCard.Value = True Then
         !InvoiceNo = FrmPrint.TxtInvoiceNo.Text
         !Commision = FrmPrint.TxtCommision.Text
         !BankMachineID = FrmPrint.TxtBankMachineID.Text
         !CashReceived = Null
         !CustomerID = "621"
         !CustomerName = IIf(Trim(FrmPrint.TxtBankCustomer.Text) = "", Null, FrmPrint.TxtBankCustomer.Text)
      End If
      If FrmPrint.OptCash.Value = True Then
         !Commision = Null
         !InvoiceNo = Null
         !BankMachineID = Null
         !CashReceived = Val(FrmPrint.TxtCashReceivedCash.Text)
         !CustomerID = "621"
         !CustomerName = IIf(Trim(FrmPrint.TxtCashCustomer.Text) = "", Null, FrmPrint.TxtCashCustomer.Text)
      End If
      If FrmPrint.OptCredit.Value = True Then
         !Commision = Null
         !InvoiceNo = Null
         !BankMachineID = Null
         !CashReceived = Val(FrmPrint.TxtCashReceivedCredit.Text)
         !CustomerID = FrmPrint.TxtCustomerID.Text
         !CustomerName = Null
      End If
      !BankCard = FrmPrint.OptBankCard.Value
      !Cash = FrmPrint.OptCash.Value
      !Credit = FrmPrint.OptCredit.Value
      !Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
      !UserNo = vUser
      .Update
      .Close
   End With
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !BillID = Val(TxtBillID.Text)
         !BillDate = DtpBillDate.DateValue
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   If vIsNewRecord = True Then Call ActivityLog("Service Invoice", eAdd, TxtBillID.Text, DtpBillDate.DateValue)
   CN.CommitTrans
   Char.Speak "Thank you for comming"
   'If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
   If FrmPrint.ChkPrint.Value = 1 Then Call BtnPrint_Click
'   If vCashDrawer = True Then
'      'Shell "mode com1 9600,n,8,1", vbNormalFocus
'      'Shell "echo ^G>com1", vbNormalFocus
'      MSComm1.Output = "O"
'   End If
   FormStatus = NewMode
   'End If
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from ServiceBody where BillId=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.productname, b.* from ServiceBody b join SProducts p on p.productid = b.productid where billid=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         'TxtGrossAmount.Text = 0
         TxtTotalQty.Caption = 0
         'TxtTotalDiscount.Caption = 0
         vTotDisc = 0
         vTotalAmount = 0
         TxtTotalAmount.Caption = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
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
            vTotalAmount = vTotalAmount + !Amount
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
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
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
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      TxtBillID.Text = FunGetMaxID()
      Call PopulateDataToGrid
      'TxtCustomerID.Text = "621"
      'TxtCustomerName.Text = "Counter Sale"
      'DtpBillDate.DateValue = Date
      LblCaptionPrice.Visible = False
      LblPrice.Visible = False
      TxtProductID.Enabled = True
      BtnProduct.Enabled = True
      DtpBillDate.Enabled = True
      'If DtpBillDate.Enabled And DtpBillDate.Visible Then DtpBillDate.SetFocus
      TxtProductID.Enabled = True
      If TxtProductID.Visible And TxtProductID.Enabled Then TxtProductID.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      DtpBillDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      TxtProductID.Enabled = True
      BtnProduct.Enabled = True
      TxtProductID.SetFocus
      LblCaptionPrice.Visible = False
      LblPrice.Visible = False
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
      TxtEmployeeID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub DtpBillDate_LostFocus()
   On Error GoTo ErrorHandler
   TxtBillID.Text = FunGetMaxID()
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Activate()
   On Error GoTo ErrorHandler
   If vFlag = True Then
      If TxtProductID.Enabled = True And TxtProductID.Visible = True Then TxtProductID.SetFocus
   Else
      vFlag = True
   End If
   Exit Sub
ErrorHandler:
   If Err.Number = 5 Then Resume Next
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      Select Case ActiveControl.Name
         Case TxtProductID.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name
         If TxtProductID.Enabled Then TxtProductID.SetFocus: Call SubClearDetailArea
      End Select
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
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
         Case vbKeyM
               If TxtMemberID.Visible = True And TxtMemberID.Enabled = True Then TxtMemberID.SetFocus
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
   ElseIf KeyCode = vbKeyC And Shift = vbAltMask Then
      FrmPrint.ParaInChoice = "Credit"
      FrmPrint.Show vbModal, Me
   ElseIf KeyCode = vbKeyReturn And Shift = vbShiftMask Then
      Select Case ActiveControl.Name
      Case TxtProductID.Name
         If FunSelectProduct(ssValidate, False) = True Then TxtQty.SetFocus
      Case TxtQty.Name
         If TxtPrice.Visible = False Then TxtDiscPC.SetFocus Else TxtPrice.SetFocus
      Case TxtPrice.Name
         TxtDiscPC.SetFocus
      Case TxtDiscPC.Name
         TxtDiscPer.SetFocus
      Case TxtDiscPer.Name
         TxtDiscVal.SetFocus
      End Select
      KeyCode = 0
      Shift = 0
   ElseIf KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name
         Grid_DblClick
      Case TxtProductID.Name
         If FunSelectProduct(ssValidate, False) = True Then GetDataFromTexBoxesToGrid
      Case TxtQty.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name, TxtPrice.Name
         GetDataFromTexBoxesToGrid
      Case Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtEmployeeID.SetFocus
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then If TxtMemberID.Visible = True Then If TxtProductID.Enabled Then TxtProductID.SetFocus
         Case TxtProductID.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtQty.SetFocus
         Case TxtMemberID.Name: If FunSelectMember(ssFunctionKey, True) = True Then If TxtProductID.Enabled Then TxtProductID.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtProductID.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtProductID.Name Then
         KeyCode = 0
         TxtBillDisc.SetFocus
      End If
   ElseIf ActiveControl.Name = Grid.Name And KeyCode = vbKeyF4 Then
      If Trim(Grid.Columns("ProductID").Text <> "") Then
         If MniCostPrice.Visible = True Then
            Call MniCostPrice_Click
         End If
      End If
   ElseIf KeyCode = vbKeyF5 Then
      If TxtPID.Text <> "" Then
         Select Case ActiveControl.Name
         Case TxtProductID.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, Grid.Name
            LblCost.Caption = CN.Execute("select dbo.FunPurPrice('" & TxtPID.Text & "')").Fields(0).Value
            LblCost.Visible = True
         End Select
      End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrorHandler
   If KeyAscii = vbKeyReturn Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If ActiveControl.Name = Grid.Name And (KeyCode = vbKeyF4 Or KeyCode = vbKeyF5) Then
      LblCost.Visible = False
   End If
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
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   vFlag = True
   Call InvoiceNo
   SetWindowText Me.hWnd, "Services Invoice (" & LblNo & ")"
   HelpLocation Me
   DtpBillDate.DateValue = IIf(Format(Now, "hh") > 3, Date, DateAdd("d", -1, Date))
   With CN.Execute("select * from registry")
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         TxtStoreID.Visible = !StoreVisible
         BtnStore.Visible = !StoreVisible
         TxtStoreName.Visible = !StoreVisible
         LblStoreID.Visible = !StoreVisible
         LblStoreName.Visible = !StoreVisible
         LblEmpID.Visible = !EmpVisible
         LblEmpName.Visible = !EmpVisible
         TxtEmployeeID.Visible = !EmpVisible
         TxtEmployeeName.Visible = !EmpVisible
         BtnEmployee.Visible = !EmpVisible
         vX = IIf(IsNull(!X), 0, !X)
         vY = IIf(IsNull(!Y), 0, !Y)
         LblMemberID.Visible = !MemberVisible
         LblMemberName.Visible = !MemberVisible
         TxtMemberID.Visible = !MemberVisible
         TxtMemberName.Visible = !MemberVisible
         BtnMember.Visible = !MemberVisible
         'vCashDrawer = !CashDrawer
         vLaserInvoice = !LaserPrintofSaleInvoice
         vPrintHeader = !PrintHeadersSaleInvoice
         vNoofPrints = IIf(IsNull(!NoofPrints) Or !NoofPrints = 0, 1, !NoofPrints)
         MniCostPrice.Visible = !CostVisible
         If !ChangePrice = True Then
            If ObjUserSecurity.IsAdministrator = True Then
               TxtPrice.Enabled = True
            End If
         End If
      End If
      .Close
   End With
  
   DateFlag = True
   FormStatus = NewMode
   'If TxtProductID.Visible And TxtProductID.Enabled Then TxtProductID.SetFocus
'   If vCashDrawer = True Then
'      MSComm1.CommPort = 1             'Use com1 port
'      MSComm1.Settings = "9600,N,8,1" 'Port Settings
'      If MSComm1.PortOpen = False Then MSComm1.PortOpen = True         'open port
'   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub InvoiceNo()
   On Error GoTo ErrorHandler
   Dim vC As Byte, LoopFlag As Boolean
   vC = 1: LoopFlag = True
   With CN.Execute("Select * from TempNo where UserNo = " & vUser & " order by tempno")
      While (Not .EOF) And LoopFlag = True
         If vC <> !TempNo And Not .EOF Then
            LoopFlag = False
         Else
            vC = vC + 1
         End If
         .MoveNext
      Wend
      LblNo.Caption = " Inv. Open # " & CStr(vC)
      CN.Execute "INSERT INTO TempNo(TempNo,UserNo) VALUES (" & vC & "," & vUser & ")"
      .Close
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpBillDate.IsDateValid = False Then FunGetMaxID = 1: Exit Function
   FunGetMaxID = CN.Execute("Select isnull(max(BillID),0)+1 from ServiceHeader where BillDate = '" & DtpBillDate.DateValue & "'").Fields(0)
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
   vTotalAmount = 0
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
      If MsgBox("Are you sure to close without save?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then
         Cancel = 1
      End If
   Else
   CN.Execute "delete from tempno where tempno = " & Val(Right(LblNo.Caption, 1))
    'CN.Execute ("exec spcurrentstock")
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set RsBody = Nothing
    Set RsReport = Nothing
    Set FrmServiceInvoice = Nothing
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
   vTotalAmount = vTotalAmount - Grid.Columns("Amount").Value
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
   TxtProductID.Enabled = False
   BtnProduct.Enabled = False
   'TxtProductID.BackColor = TxtProductName.BackColor
   'TxtProductID.TabStop = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   LblCost.Visible = False
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      TxtProductID.Text = ""
      TxtProductID.Enabled = True
      BtnProduct.Enabled = True
      If TxtProductID.Enabled Then TxtProductID.SetFocus
   Else
      TxtProductID.Enabled = False
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

Private Sub MniCostPrice_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Cost").Text) = "" Then Exit Sub
   LblCost.Caption = Grid.Columns("Cost").Value
   LblCost.Visible = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ProductID").Text) = "" Then Exit Sub
   RsBody.Filter = "ProductID='" & TxtProductID.Text & "'"
   If RsBody.RecordCount > 0 Then RsBody.Delete
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   RsBody.Filter = 0
   Grid.MoveLast
   GetDataBackFromGridToTexBoxes
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   If Trim(TxtProductID.Text) = "" Then
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
      TxtProductID.SetFocus
      Exit Sub
   End If
   If Val(TxtQty.Text) = 0 Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      TxtQty.SetFocus
      Exit Sub
   End If
   RsBody.Filter = "ProductID='" & TxtPID.Text & "'"
   If TxtProductID.Enabled Then
      If RsBody.RecordCount = 0 Then
'         If Trim(TxtQty.Text) > Val(LblStock.Caption) Then
'            MsgBox "Insufficent Stock.", vbExclamation, "Alert"
'            TxtQty.SetFocus
'            Exit Sub
'         End If
         RsBody.AddNew
         Grid.Columns("ProductID").Text = TxtProductID.Text
         RsBody!Productid = TxtProductID.Text
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("Productid").Text = TxtProductID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  TxtQty.Text = Val(TxtQty.Text) + Grid.Columns("Qty").Value
                  TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text) - Val(Grid.Columns("Qty").Text)
                  vTotDisc = vTotDisc + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
                  vTotalAmount = vTotalAmount + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  
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
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
                  RsBody!Qty = Val(TxtQty.Text)
                  RsBody!Price = Val(TxtPrice.Text)
                  RsBody!DiscPC = Val(TxtDiscPC.Text)
                  RsBody!DiscPer = Val(TxtDiscPer.Text)
                  RsBody!DiscVal = Val(TxtDiscVal.Text)
                  RsBody!Amount = Val(TxtAmount.Text)
                  Grid.MoveLast
                  Call SubClearDetailArea
                  TxtProductID.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearDetailArea
         Grid.MoveLast
         TxtProductID.SetFocus
         Exit Sub
      End If
   End If
   'Grid.Redraw = False
   With Grid
      If TxtProductID.Enabled = True Then
         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text)
         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text)
         'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(TxtDiscVal.Text)
         vTotDisc = vTotDisc + Val(TxtDiscVal.Text)
         vTotalAmount = vTotalAmount + Val(TxtAmount.Text)
         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text)
      Else
         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text) - Val(.Columns("Qty").Text)
         vTotDisc = vTotDisc + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
         vTotalAmount = vTotalAmount + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text) - Val(Grid.Columns("TotalAmount").Text)
      End If
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("Qty").Value = Val(TxtQty.Text)
      .Columns("Price").Value = Val(TxtPrice.Text)
      .Columns("DiscPC").Value = Val(TxtDiscPC.Text)
      .Columns("DiscPer").Value = Val(TxtDiscPer.Text)
      .Columns("DiscVal").Value = Val(TxtDiscVal.Text)
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
      TxtLastRate.Caption = Val(TxtPrice.Text) - Val(TxtDiscPC.Text)
      RsBody!Qty = Val(TxtQty.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!DiscPC = Val(TxtDiscPC.Text)
      RsBody!DiscPer = Val(TxtDiscPer.Text)
      RsBody!DiscVal = Val(TxtDiscVal.Text)
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
   TxtProductID.SetFocus
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtProductID.Enabled = True
   BtnProduct.Enabled = True
   TxtProductID.Text = ""
   TxtProductName.Text = ""
   TxtQty.Text = ""
   TxtPrice.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtDiscVal.Text = ""
   TxtAmount.Text = ""
   TxtCost.Text = ""
   TxtActualAmount.Text = ""
   ChkIsProduct.Value = 1
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtPID.Text = .Columns("ProductID").Text
      TxtProductID.Text = .Columns("ProductID").Text
      TxtProductName.Text = .Columns("ProductName").Text
      TxtQty.Text = .Columns("Qty").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      'TxtCost.Text = .Columns("Cost").Value
      TxtAmount.Text = .Columns("Amount").Text
      TxtActualAmount.Text = .Columns("TotalAmount").Text
      'ChkIsProduct.Value = Abs(.Columns("IsProduct").Value)
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetSale()
   On Error GoTo ErrorHandler
   ssql = "select h.*, c.AccountName, BankMachineName, StoreName, EmpName, MemberName FROM ServiceHeader h left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join BankMachines b on b.BankMachineid = h.BankMachineid left outer join Members m on m.MemberID = h.MemberID inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID where h.BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   With CN.Execute(ssql)
      If Not .BOF Then
         TxtStoreID.Text = !StoreID
         TxtStoreName.Text = !StoreName
         TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
         TxtEmployeeName.Text = IIf(IsNull(!EmpName), "", !EmpName)
         TxtMemberID.Text = IIf(IsNull(!MemberID), "", !MemberID)
         TxtMemberName.Text = IIf(IsNull(!MemberName), "", !MemberName)
         TxtTotalAmount.Caption = !TotalAmount
         TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
         TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
         TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
         FrmPrint.OptBankCard.Value = !BankCard
         FrmPrint.OptCash.Value = !Cash
         FrmPrint.OptCredit.Value = !Credit
         If FrmPrint.OptBankCard.Value = True Then
            FrmPrint.TxtInvoiceNo.Text = !InvoiceNo
            FrmPrint.TxtCommision.Text = !Commision
            FrmPrint.TxtBankMachineID.Text = !BankMachineID
            FrmPrint.TxtBankMachineName.Text = !BankMachineName
            FrmPrint.TxtCashReceivedCash.Text = ""
            FrmPrint.TxtCustomerID.Text = ""
            FrmPrint.TxtCustomerName.Text = ""
            FrmPrint.TxtCashCustomer.Text = ""
            FrmPrint.TxtBankCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
         End If
         If FrmPrint.OptCash.Value = True Then
            FrmPrint.TxtCommision.Text = ""
            FrmPrint.TxtInvoiceNo.Text = ""
            FrmPrint.TxtBankMachineID.Text = ""
            FrmPrint.TxtBankMachineName.Text = ""
            FrmPrint.TxtCashReceivedCash.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            FrmPrint.TxtCustomerID.Text = ""
            FrmPrint.TxtCustomerName.Text = ""
            FrmPrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
            FrmPrint.TxtBankCustomer.Text = ""
         End If
         If FrmPrint.OptCredit.Value = True Then
            FrmPrint.TxtCommision.Text = ""
            FrmPrint.TxtInvoiceNo.Text = ""
            FrmPrint.TxtBankMachineID.Text = ""
            FrmPrint.TxtBankMachineName.Text = ""
            FrmPrint.TxtCashReceivedCredit.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            FrmPrint.TxtCustomerID.Text = !CustomerID
            FrmPrint.TxtCustomerName.Text = !AccountName
            FrmPrint.TxtCashCustomer.Text = ""
            FrmPrint.TxtBankCustomer.Text = ""
         End If
          TxtNetAmount.Caption = !TotalAmount
         Call PopulateDataToGrid
      End If
      .Close
   End With
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

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

Private Sub TxtDiscPC_Change()
   If ActiveControl.Name <> TxtDiscPC.Name Then Exit Sub
   If Val(TxtPrice.Text) = 0 Then Exit Sub
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
   Call SubCalculateBody
End Sub

'Private Sub TxtDiscPC_LostFocus()
'   Select Case Me.ActiveControl.Name
'   Case TxtProductID.Name, TxtQty.Name, TxtDiscPC.Name
'      Exit Sub
'   End Select
'   Call GetDataFromTexBoxesToGrid
'End Sub

Private Sub TxtDiscPer_Change()
   If ActiveControl.Name <> TxtDiscPer.Name Then Exit Sub
   TxtDiscPC.Text = Round((Val(TxtPrice.Text) * Val(TxtDiscPer.Text) / 100), 2)
   Call SubCalculateBody
End Sub

Private Sub TxtProductID_Change()
   If ActiveControl.Name <> TxtProductID.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
      TxtPrice.Text = ""
      TxtDiscPC.Text = ""
   End If
End Sub

Private Sub TxtProductID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

'Private Sub TxtProductID_LostFocus()
'   If Len(TxtProductID.Text) > 7 Then
'      GetDataFromTexBoxesToGrid
'   End If
'End Sub

Private Sub TxtProductID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtProductID.Text) = "" Then Exit Sub
   vTemp = Not FunSelectProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDiscVal_Change()
   If TxtDiscVal.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtDiscVal.Name Then Exit Sub
   If Val(TxtPrice.Text) = 0 Then Exit Sub
   If Val(TxtQty.Text) = 0 Then Exit Sub
   TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (TxtQty.Text), 3)
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
   TxtActualAmount.Text = Val(TxtQty.Text) * Val(TxtPrice.Text)
   TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
   TxtTotalDiscount.Caption = vTotDisc
   SubCalculateFooter
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

Private Sub TxtMemberID_Change()
   If ActiveControl.Name <> TxtMemberID.Name Then Exit Sub
   If TxtMemberName.Text <> "" Then TxtMemberName.Text = "": Call SubDestroyMember
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
