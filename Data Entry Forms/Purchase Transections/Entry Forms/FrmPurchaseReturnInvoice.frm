VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmPurchaseReturnInvoice 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkIsPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Print"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12825
      TabIndex        =   136
      Top             =   10110
      Width           =   1290
   End
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12825
      TabIndex        =   133
      Top             =   9855
      Width           =   1290
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   495
      TabIndex        =   130
      Top             =   8325
      Visible         =   0   'False
      Width           =   2295
      Begin VB.TextBox TxtSerial 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   20
         TabIndex        =   131
         Top             =   180
         Width           =   2025
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridSerial 
         Height          =   1500
         Left            =   120
         TabIndex        =   132
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
         stylesets(0).Picture=   "FrmPurchaseReturnInvoice.frx":0000
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
         Columns(1).Caption=   "Serial No"
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
   Begin VB.Frame FrmHistory 
      Height          =   1635
      Left            =   3060
      TabIndex        =   128
      Top             =   6165
      Visible         =   0   'False
      Width           =   11295
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridHistory 
         Height          =   1455
         Left            =   -750
         TabIndex        =   129
         Top             =   0
         Width           =   11805
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   14
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
         stylesets(0).Picture=   "FrmPurchaseReturnInvoice.frx":001C
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
         stylesets(1).Picture=   "FrmPurchaseReturnInvoice.frx":0038
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
         stylesets(2).Picture=   "FrmPurchaseReturnInvoice.frx":0054
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
         Columns.Count   =   14
         Columns(0).Width=   1402
         Columns(0).Caption=   "VID"
         Columns(0).Name =   "ID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3281
         Columns(1).Caption=   "Vendor Name"
         Columns(1).Name =   "Name"
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   1270
         Columns(2).Caption=   "PurID"
         Columns(2).Name =   "PurID"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1799
         Columns(3).Caption=   "Date"
         Columns(3).Name =   "Date"
         Columns(3).CaptionAlignment=   2
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).NumberFormat=   "dd/MM/yyyy"
         Columns(3).FieldLen=   256
         Columns(4).Width=   1852
         Columns(4).Caption=   "Expiry Date"
         Columns(4).Name =   "ExpiryDate"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).NumberFormat=   "dd/MM/yyyy"
         Columns(4).FieldLen=   256
         Columns(5).Width=   847
         Columns(5).Caption=   "Pack"
         Columns(5).Name =   "Pack"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   1005
         Columns(6).Caption=   "Qty(P)"
         Columns(6).Name =   "QtyPack"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   1032
         Columns(7).Caption=   "Qty(L)"
         Columns(7).Name =   "QtyLoose"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   688
         Columns(8).Caption=   "Bns"
         Columns(8).Name =   "Bonus"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   1005
         Columns(9).Caption=   "Price"
         Columns(9).Name =   "Price"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   1111
         Columns(10).Caption=   "D(PC)"
         Columns(10).Name=   "DiscPC"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(11).Width=   979
         Columns(11).Caption=   "D(Per)"
         Columns(11).Name=   "DiscPer"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(12).Width=   1085
         Columns(12).Caption=   "D(Val)"
         Columns(12).Name=   "DiscVal"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
         Columns(13).Width=   1588
         Columns(13).Caption=   "Amount"
         Columns(13).Name=   "Value"
         Columns(13).Alignment=   1
         Columns(13).CaptionAlignment=   2
         Columns(13).DataField=   "Column 13"
         Columns(13).DataType=   8
         Columns(13).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   20823
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "History"
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
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmPurchaseReturnInvoice.frx":0070
      Left            =   9540
      List            =   "FrmPurchaseReturnInvoice.frx":0072
      Style           =   2  'Dropdown List
      TabIndex        =   119
      Tag             =   "1"
      Top             =   9930
      Width           =   3276
   End
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   11640
      TabIndex        =   118
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   9480
      Width           =   1170
   End
   Begin VB.Frame FrmProductPrices 
      Height          =   1095
      Left            =   6510
      TabIndex        =   116
      Top             =   780
      Visible         =   0   'False
      Width           =   6270
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridProductPrices 
         Height          =   885
         Left            =   60
         TabIndex        =   117
         Top             =   150
         Width           =   6135
         ScrollBars      =   0
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
         stylesets(0).Picture=   "FrmPurchaseReturnInvoice.frx":0074
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
         stylesets(1).Picture=   "FrmPurchaseReturnInvoice.frx":0090
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
         stylesets(2).Picture=   "FrmPurchaseReturnInvoice.frx":00AC
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
         ForeColorOdd    =   8388736
         BackColorEven   =   16776960
         RowHeight       =   714
         Columns.Count   =   5
         Columns(0).Width=   1402
         Columns(0).Caption=   "Pur"
         Columns(0).Name =   "Pur"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   5
         Columns(0).FieldLen=   256
         Columns(1).Width=   1402
         Columns(1).Caption=   "List"
         Columns(1).Name =   "List"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   5
         Columns(1).FieldLen=   256
         Columns(2).Width=   1402
         Columns(2).Caption=   "WS"
         Columns(2).Name =   "WS"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   5
         Columns(2).FieldLen=   256
         Columns(3).Width=   1402
         Columns(3).Caption=   "Retail"
         Columns(3).Name =   "Retail"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   5
         Columns(3).FieldLen=   256
         Columns(4).Width=   5239
         Columns(4).Caption=   "Description"
         Columns(4).Name =   "Description"
         Columns(4).CaptionAlignment=   2
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
         TabNavigation   =   1
         _ExtentX        =   10821
         _ExtentY        =   1561
         _StockProps     =   79
         Caption         =   "Product Prices"
         ForeColor       =   0
         BackColor       =   16776960
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
   End
   Begin VB.ComboBox cmbSizeName 
      Height          =   315
      Left            =   5745
      Style           =   2  'Dropdown List
      TabIndex        =   113
      Top             =   4500
      Width           =   840
   End
   Begin VB.ComboBox CmbColourName 
      Height          =   315
      Left            =   4545
      Style           =   2  'Dropdown List
      TabIndex        =   112
      Top             =   4500
      Width           =   1200
   End
   Begin VB.Frame FramExpense 
      Height          =   2415
      Left            =   9000
      TabIndex        =   72
      Top             =   5438
      Visible         =   0   'False
      Width           =   4215
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridExpense 
         Height          =   1860
         Left            =   120
         TabIndex        =   73
         Top             =   120
         Width           =   3990
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   3
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
         stylesets(0).Picture=   "FrmPurchaseReturnInvoice.frx":00C8
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
         stylesets(1).Picture=   "FrmPurchaseReturnInvoice.frx":00E4
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
         stylesets(2).Picture=   "FrmPurchaseReturnInvoice.frx":0100
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
         Columns.Count   =   3
         Columns(0).Width=   1058
         Columns(0).Caption=   "ID"
         Columns(0).Name =   "ID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3969
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   1429
         Columns(2).Caption=   "Amount"
         Columns(2).Name =   "Value"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   2
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   7038
         _ExtentY        =   3281
         _StockProps     =   79
         Caption         =   "Expenses"
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
      Begin SITextBox.Txt TxtTotalExpense 
         Height          =   315
         Left            =   2880
         TabIndex        =   74
         Top             =   2040
         Width           =   915
         _ExtentX        =   1614
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
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Expense"
         Height          =   195
         Left            =   1680
         TabIndex        =   75
         Top             =   2100
         Width           =   1020
      End
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
      Height          =   4110
      Left            =   13800
      TabIndex        =   62
      Top             =   1440
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
         Height          =   3750
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   63
         Tag             =   "NC"
         Text            =   "FrmPurchaseReturnInvoice.frx":011C
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
         TabIndex        =   64
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   6555
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4493
      Width           =   1425
   End
   Begin SITextBox.Txt TxtReceivedAmount 
      Height          =   315
      Left            =   11145
      TabIndex        =   27
      Top             =   8408
      Width           =   1395
      _ExtentX        =   2461
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
      IntegralPoint   =   9
   End
   Begin SITextBox.Txt TxtVenderID 
      Height          =   315
      Left            =   1905
      TabIndex        =   6
      Top             =   3105
      Width           =   930
      _ExtentX        =   1640
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtVenderName 
      Height          =   315
      Left            =   3195
      TabIndex        =   37
      Top             =   3098
      Width           =   3645
      _ExtentX        =   6429
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
   Begin SITextBox.Txt TxtAddress 
      Height          =   315
      Left            =   6840
      TabIndex        =   36
      Top             =   3098
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin SITextBox.Txt TxtReturnID 
      Height          =   315
      Left            =   1890
      TabIndex        =   0
      Top             =   2303
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
   Begin SITextBox.Txt TxtCity 
      Height          =   315
      Left            =   11370
      TabIndex        =   35
      Top             =   3098
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin JeweledBut.JeweledButton BtnVender 
      Height          =   330
      Left            =   2835
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3098
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":0233
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9027
      TabIndex        =   32
      Top             =   9413
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":024F
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7785
      TabIndex        =   28
      Top             =   9420
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":026B
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5067
      TabIndex        =   30
      Top             =   9413
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":0287
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10347
      TabIndex        =   33
      Top             =   9413
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":02A3
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6387
      TabIndex        =   29
      Top             =   9413
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":02BF
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpReturnDate 
      Height          =   315
      Left            =   3180
      TabIndex        =   1
      Top             =   2303
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3739
      TabIndex        =   31
      Top             =   9413
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":02DB
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   11130
      TabIndex        =   47
      Top             =   1538
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   750
      TabIndex        =   11
      Top             =   4500
      Width           =   855
      _ExtentX        =   1508
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
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   1605
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4500
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":02F7
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   1965
      TabIndex        =   14
      Top             =   4500
      Width           =   2565
      _ExtentX        =   4524
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
      Left            =   180
      TabIndex        =   50
      Top             =   4815
      Width           =   14505
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   27
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
      stylesets(0).Picture=   "FrmPurchaseReturnInvoice.frx":0313
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
      Columns.Count   =   27
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ProductID"
      Columns(0).Name =   "ProductID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1005
      Columns(1).Caption=   "Serial"
      Columns(1).Name =   "Serial"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2143
      Columns(2).Caption=   "Code"
      Columns(2).Name =   "Code"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   4498
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
      Columns(6).Width=   2514
      Columns(6).Caption=   "Pack Name"
      Columns(6).Name =   "PackName"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   900
      Columns(7).Caption=   "Pack"
      Columns(7).Name =   "Pack"
      Columns(7).Alignment=   1
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   900
      Columns(8).Caption=   "Q(P)"
      Columns(8).Name =   "QtyPack"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   4
      Columns(8).FieldLen=   256
      Columns(9).Width=   953
      Columns(9).Caption=   "Q(L)"
      Columns(9).Name =   "QtyLoose"
      Columns(9).Alignment=   1
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   4
      Columns(9).FieldLen=   256
      Columns(10).Width=   953
      Columns(10).Caption=   "Bns"
      Columns(10).Name=   "Bonus"
      Columns(10).Alignment=   1
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   4
      Columns(10).FieldLen=   256
      Columns(11).Width=   847
      Columns(11).Caption=   "Offer"
      Columns(11).Name=   "Offer"
      Columns(11).Alignment=   1
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   1138
      Columns(12).Caption=   "Price"
      Columns(12).Name=   "Price"
      Columns(12).Alignment=   1
      Columns(12).CaptionAlignment=   2
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   4
      Columns(12).FieldLen=   256
      Columns(13).Width=   1191
      Columns(13).Caption=   "DiscPC"
      Columns(13).Name=   "DiscPC"
      Columns(13).Alignment=   1
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   926
      Columns(14).Caption=   "Tax%"
      Columns(14).Name=   "SaleTaxPer"
      Columns(14).Alignment=   1
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(15).Width=   847
      Columns(15).Caption=   "Dis%"
      Columns(15).Name=   "DiscPer"
      Columns(15).Alignment=   1
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   1191
      Columns(16).Caption=   "Dis.Val"
      Columns(16).Name=   "DiscVal"
      Columns(16).Alignment=   1
      Columns(16).CaptionAlignment=   2
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   4
      Columns(16).FieldLen=   256
      Columns(17).Width=   1508
      Columns(17).Caption=   "Amount"
      Columns(17).Name=   "Amount"
      Columns(17).Alignment=   1
      Columns(17).CaptionAlignment=   2
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   5
      Columns(17).FieldLen=   256
      Columns(18).Width=   3200
      Columns(18).Visible=   0   'False
      Columns(18).Caption=   "PackingID"
      Columns(18).Name=   "PackingID"
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(19).Width=   3200
      Columns(19).Visible=   0   'False
      Columns(19).Caption=   "SaleTaxVal"
      Columns(19).Name=   "SaleTaxVal"
      Columns(19).Alignment=   1
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(20).Width=   1984
      Columns(20).Caption=   "BatchNo"
      Columns(20).Name=   "BatchNo"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      Columns(21).Width=   3200
      Columns(21).Visible=   0   'False
      Columns(21).Caption=   "ColourID"
      Columns(21).Name=   "ColourID"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   8
      Columns(21).FieldLen=   256
      Columns(22).Width=   3200
      Columns(22).Visible=   0   'False
      Columns(22).Caption=   "SizeID"
      Columns(22).Name=   "SizeID"
      Columns(22).DataField=   "Column 22"
      Columns(22).DataType=   8
      Columns(22).FieldLen=   256
      Columns(23).Width=   3200
      Columns(23).Caption=   "RetailPrice"
      Columns(23).Name=   "RetailPrice"
      Columns(23).DataField=   "Column 23"
      Columns(23).DataType=   8
      Columns(23).FieldLen=   256
      Columns(24).Width=   3200
      Columns(24).Caption=   "RetailAmount"
      Columns(24).Name=   "RetailAmount"
      Columns(24).DataField=   "Column 24"
      Columns(24).DataType=   8
      Columns(24).FieldLen=   256
      Columns(25).Width=   3200
      Columns(25).Caption=   "ProfitAmount"
      Columns(25).Name=   "ProfitAmount"
      Columns(25).DataField=   "Column 25"
      Columns(25).DataType=   8
      Columns(25).FieldLen=   256
      Columns(26).Width=   3200
      Columns(26).Caption=   "SaleDiscPer"
      Columns(26).Name=   "SaleDiscPer"
      Columns(26).DataField=   "Column 26"
      Columns(26).DataType=   8
      Columns(26).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   25585
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
   Begin SITextBox.Txt TxtMultiplier 
      Height          =   315
      Left            =   7980
      TabIndex        =   16
      Top             =   4500
      Width           =   510
      _ExtentX        =   900
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
   End
   Begin SITextBox.Txt TxtQtyLoose 
      Height          =   315
      Left            =   9000
      TabIndex        =   18
      Top             =   4500
      Width           =   540
      _ExtentX        =   953
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
      Left            =   8490
      TabIndex        =   17
      Top             =   4500
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
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
      IntegralPoint   =   4
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtDiscVal 
      Height          =   315
      Left            =   12870
      TabIndex        =   25
      Top             =   4500
      Width           =   675
      _ExtentX        =   1191
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
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtPrice 
      Height          =   315
      Left            =   10560
      TabIndex        =   21
      Top             =   4500
      Width           =   645
      _ExtentX        =   1138
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
   Begin SITextBox.Txt TxtBonus 
      Height          =   315
      Left            =   9540
      TabIndex        =   19
      Top             =   4500
      Width           =   540
      _ExtentX        =   953
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
   End
   Begin SITextBox.Txt TxtOffer 
      Height          =   315
      Left            =   10080
      TabIndex        =   20
      Top             =   4500
      Width           =   480
      _ExtentX        =   847
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
   End
   Begin SITextBox.Txt TxtSaleTaxVal 
      Height          =   315
      Left            =   10485
      TabIndex        =   66
      Top             =   1508
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   12390
      TabIndex        =   24
      Top             =   4500
      Width           =   480
      _ExtentX        =   847
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   13545
      TabIndex        =   26
      Top             =   4500
      Width           =   1125
      _ExtentX        =   1984
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
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   11205
      TabIndex        =   22
      Top             =   4500
      Width           =   675
      _ExtentX        =   1191
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
   Begin SITextBox.Txt TxtSaleTaxPer 
      Height          =   315
      Left            =   11880
      TabIndex        =   23
      Top             =   4500
      Width           =   510
      _ExtentX        =   900
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   2
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridOffer 
      Height          =   1365
      Left            =   1740
      TabIndex        =   71
      Top             =   6758
      Width           =   3705
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   4
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
      stylesets(0).Picture=   "FrmPurchaseReturnInvoice.frx":032F
      UseGroups       =   -1  'True
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowColumnMoving=   2
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Groups(0).Width =   6059
      Groups(0).Caption=   "Product Offer"
      Groups(0).Columns.Count=   4
      Groups(0).Columns(0).Width=   2090
      Groups(0).Columns(0).Visible=   0   'False
      Groups(0).Columns(0).Caption=   "Product ID"
      Groups(0).Columns(0).Name=   "ProductID"
      Groups(0).Columns(0).CaptionAlignment=   2
      Groups(0).Columns(0).DataField=   "Column 0"
      Groups(0).Columns(0).DataType=   8
      Groups(0).Columns(0).FieldLen=   256
      Groups(0).Columns(0).Locked=   -1  'True
      Groups(0).Columns(1).Width=   1693
      Groups(0).Columns(1).Caption=   "Product ID"
      Groups(0).Columns(1).Name=   "ProductOfferID"
      Groups(0).Columns(1).DataField=   "Column 1"
      Groups(0).Columns(1).DataType=   8
      Groups(0).Columns(1).FieldLen=   256
      Groups(0).Columns(1).Locked=   -1  'True
      Groups(0).Columns(2).Width=   3440
      Groups(0).Columns(2).Caption=   "Product Name"
      Groups(0).Columns(2).Name=   "ProductName"
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   8
      Groups(0).Columns(2).FieldLen=   256
      Groups(0).Columns(2).Locked=   -1  'True
      Groups(0).Columns(3).Width=   926
      Groups(0).Columns(3).Caption=   "Qty"
      Groups(0).Columns(3).Name=   "Qty"
      Groups(0).Columns(3).Alignment=   1
      Groups(0).Columns(3).CaptionAlignment=   2
      Groups(0).Columns(3).DataField=   "Column 3"
      Groups(0).Columns(3).DataType=   2
      Groups(0).Columns(3).FieldLen=   256
      _ExtentX        =   6535
      _ExtentY        =   2408
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
   Begin SITextBox.Txt TxtBillNo 
      Height          =   315
      Left            =   1950
      TabIndex        =   7
      Top             =   3803
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
   Begin SITextBox.Txt TxtBiltyNo 
      Height          =   315
      Left            =   2700
      TabIndex        =   8
      Top             =   3803
      Width           =   750
      _ExtentX        =   1323
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
   End
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   4965
      TabIndex        =   10
      Top             =   3803
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   200
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
   Begin SITextBox.Txt TxtVehicleNo 
      Height          =   315
      Left            =   3450
      TabIndex        =   9
      Top             =   3803
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   30
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
   Begin SITextBox.Txt TxtRetailPrice 
      Height          =   315
      Left            =   10395
      TabIndex        =   82
      Top             =   3803
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
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
   End
   Begin SITextBox.Txt TxtBatchNo 
      Height          =   315
      Left            =   1605
      TabIndex        =   12
      Top             =   4185
      Width           =   855
      _ExtentX        =   1508
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
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   3780
      TabIndex        =   84
      Top             =   8408
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
   Begin SITextBox.Txt TxtBillDiscPer 
      Height          =   315
      Left            =   4770
      TabIndex        =   85
      Top             =   8408
      Width           =   660
      _ExtentX        =   1164
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
      DecimalPoint    =   6
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtBillDisc 
      Height          =   315
      Left            =   5430
      TabIndex        =   86
      Top             =   8408
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtTotalItems 
      Height          =   315
      Left            =   2895
      TabIndex        =   87
      Top             =   8408
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
   Begin SITextBox.Txt TxtOtherCharges 
      Height          =   315
      Left            =   6270
      TabIndex        =   88
      Top             =   8408
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtNetAmount 
      Height          =   315
      Left            =   7365
      TabIndex        =   89
      Top             =   8408
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
   Begin SITextBox.Txt TxtTotalPayable 
      Height          =   315
      Left            =   9885
      TabIndex        =   90
      Top             =   8408
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
      Left            =   8625
      TabIndex        =   91
      Top             =   8408
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   7635
      TabIndex        =   4
      Tag             =   "NC"
      Top             =   2258
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
      Left            =   8670
      TabIndex        =   100
      Tag             =   "NC"
      Top             =   2258
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8310
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   2258
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":034B
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   10110
      TabIndex        =   5
      Tag             =   "NC"
      Top             =   2258
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
      Left            =   11415
      TabIndex        =   102
      Tag             =   "NC"
      Top             =   2258
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
      Left            =   11055
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   2258
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":0367
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPurchase 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6915
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   2303
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":0383
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtPurchaseID 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   2303
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpPurchaseDate 
      Height          =   315
      Left            =   5610
      TabIndex        =   3
      Top             =   2303
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
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin JeweledBut.JeweledButton BtnProductRange 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   1230
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   4185
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
      MICON           =   "FrmPurchaseReturnInvoice.frx":039F
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtRetailAmount 
      Height          =   315
      Left            =   135
      TabIndex        =   122
      Top             =   2355
      Width           =   855
      _ExtentX        =   1508
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
   Begin SITextBox.Txt TxtProfitAmount 
      Height          =   315
      Left            =   135
      TabIndex        =   123
      Top             =   2940
      Width           =   855
      _ExtentX        =   1508
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
   Begin SITextBox.Txt TxtSaleDiscPer 
      Height          =   315
      Left            =   135
      TabIndex        =   124
      Top             =   1725
      Width           =   855
      _ExtentX        =   1508
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
   Begin SITextBox.Txt TxtSID 
      Height          =   315
      Left            =   1980
      TabIndex        =   134
      Top             =   1650
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "SID"
      Height          =   195
      Left            =   1980
      TabIndex        =   135
      Top             =   1440
      Width           =   270
   End
   Begin VB.Label LblRetailAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Retail Amount"
      Height          =   195
      Left            =   135
      TabIndex        =   127
      Top             =   2160
      Width           =   990
   End
   Begin VB.Label LblProfitAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Profit Amount"
      Height          =   195
      Left            =   135
      TabIndex        =   126
      Top             =   2745
      Width           =   945
   End
   Begin VB.Label LblSaleDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Disc Per"
      Height          =   195
      Left            =   135
      TabIndex        =   125
      Top             =   1530
      Width           =   960
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
      Left            =   8865
      TabIndex        =   121
      Top             =   9975
      Width           =   570
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
      Left            =   11655
      TabIndex        =   120
      Top             =   9225
      Width           =   840
   End
   Begin VB.Label LblSize 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   5745
      TabIndex        =   115
      Top             =   4305
      Width           =   300
   End
   Begin VB.Label LblColour 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
      Height          =   195
      Left            =   4545
      TabIndex        =   114
      Top             =   4305
      Width           =   450
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   7635
      TabIndex        =   110
      Top             =   2063
      Width           =   585
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   8670
      TabIndex        =   109
      Top             =   2063
      Width           =   840
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   10110
      TabIndex        =   108
      Top             =   2063
      Width           =   1095
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   11415
      TabIndex        =   107
      Top             =   2063
      Width           =   1350
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase ID"
      Height          =   195
      Left            =   4560
      TabIndex        =   106
      Top             =   2108
      Width           =   885
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Date"
      Height          =   195
      Left            =   5610
      TabIndex        =   105
      Top             =   2108
      Width           =   1065
   End
   Begin VB.Label LblTotalAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amount"
      Height          =   195
      Left            =   3735
      TabIndex        =   99
      Top             =   8183
      Width           =   990
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. (%)"
      Height          =   195
      Left            =   4815
      TabIndex        =   98
      Top             =   8183
      Width           =   615
   End
   Begin VB.Label LblNetAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      Height          =   195
      Left            =   7425
      TabIndex        =   97
      Top             =   8183
      Width           =   840
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   5535
      TabIndex        =   96
      Top             =   8183
      Width           =   630
   End
   Begin VB.Label LblTtlPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payable"
      Height          =   195
      Left            =   9930
      TabIndex        =   95
      Top             =   8183
      Width           =   975
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items"
      Height          =   195
      Left            =   2895
      TabIndex        =   94
      Top             =   8183
      Width           =   780
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      Height          =   195
      Left            =   6285
      TabIndex        =   93
      Top             =   8183
      Width           =   1020
   End
   Begin VB.Label lblPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Payable"
      Height          =   195
      Left            =   8580
      TabIndex        =   92
      Top             =   8183
      Width           =   1260
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ret Price"
      Height          =   195
      Left            =   10425
      TabIndex        =   83
      Top             =   3548
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label LblCaptionRetailPrice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retail Price"
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
      Left            =   12165
      TabIndex        =   81
      Top             =   3503
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LblRetailPrice 
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
      Left            =   12570
      TabIndex        =   80
      Top             =   3818
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   4965
      TabIndex        =   79
      Top             =   3593
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bilty No."
      Height          =   195
      Left            =   2685
      TabIndex        =   78
      Top             =   3593
      Width           =   585
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      Height          =   195
      Left            =   1950
      TabIndex        =   77
      Top             =   3593
      Width           =   495
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      Height          =   195
      Left            =   3435
      TabIndex        =   76
      Top             =   3593
      Width           =   780
   End
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   13545
      TabIndex        =   70
      Top             =   4305
      Width           =   540
   End
   Begin VB.Label LblOffer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Offer"
      Height          =   195
      Left            =   10080
      TabIndex        =   69
      Top             =   4305
      Width           =   345
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Val"
      Height          =   195
      Left            =   10485
      TabIndex        =   68
      Top             =   1298
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label LblSaleTaxPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax%"
      Height          =   195
      Left            =   11880
      TabIndex        =   67
      Top             =   4305
      Width           =   390
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
      Left            =   13065
      TabIndex        =   65
      Top             =   1613
      Width           =   435
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   2910
      TabIndex        =   61
      Top             =   4305
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   750
      TabIndex        =   60
      Top             =   4305
      Width           =   375
   End
   Begin VB.Label LblPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   10560
      TabIndex        =   59
      Top             =   4305
      Width           =   405
   End
   Begin VB.Label LblMultiplier 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   8010
      TabIndex        =   58
      Top             =   4305
      Width           =   375
   End
   Begin VB.Label LblPackName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   6555
      TabIndex        =   57
      Top             =   4305
      Width           =   840
   End
   Begin VB.Label LblQtyLoose 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   9000
      TabIndex        =   56
      Top             =   4305
      Width           =   465
   End
   Begin VB.Label LblQtyPack 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   8490
      TabIndex        =   55
      Top             =   4305
      Width           =   480
   End
   Begin VB.Label LblDiscVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. Val"
      Height          =   195
      Left            =   12870
      TabIndex        =   54
      Top             =   4305
      Width           =   630
   End
   Begin VB.Label LblDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Dis%"
      Height          =   195
      Left            =   12390
      TabIndex        =   53
      Top             =   4305
      Width           =   345
   End
   Begin VB.Label LblDiscPC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc/PC"
      Height          =   195
      Left            =   11205
      TabIndex        =   52
      Top             =   4305
      Width           =   600
   End
   Begin VB.Label LblBonus 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bns(L)"
      Height          =   195
      Left            =   9540
      TabIndex        =   51
      Top             =   4305
      Width           =   450
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Return Invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   49
      Top             =   270
      Width           =   4260
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   11130
      TabIndex        =   48
      Top             =   1343
      Visible         =   0   'False
      Width           =   720
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
      Left            =   11235
      TabIndex        =   46
      Top             =   3503
      Width           =   720
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
      Left            =   11130
      TabIndex        =   45
      Top             =   3818
      Width           =   1035
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Received Amount"
      Height          =   195
      Left            =   11145
      TabIndex        =   44
      Top             =   8198
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   11370
      TabIndex        =   43
      Top             =   2918
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   6840
      TabIndex        =   42
      Top             =   2888
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Name"
      Height          =   195
      Left            =   3195
      TabIndex        =   41
      Top             =   2888
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender ID"
      Height          =   195
      Left            =   1890
      TabIndex        =   40
      Top             =   2888
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      Height          =   195
      Left            =   3195
      TabIndex        =   39
      Top             =   2108
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Return ID"
      Height          =   195
      Left            =   1905
      TabIndex        =   38
      Top             =   2108
      Width           =   690
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmPurchaseReturnInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim vDate, vServerDate As Date, vHDiff As Integer, vSystemDate As Boolean
Dim vMode As FormMode
Dim vUnitPrice, vNegativeSale As Double, vAutoEnterBeforeQty As Boolean
Dim vUnitRetailPrice As Double
Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim vMaxBinID, vGridRows As Integer
Dim RsBody As New ADODB.Recordset
Dim RsProductOffer As New ADODB.Recordset
Dim RsBodySerial As New ADODB.Recordset
Dim RsPurchaseSerial As New ADODB.Recordset
Dim RsExpense As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim QtyOffer As Integer
Dim Rebate As Integer
Dim Flag As Boolean
Dim ssql As String
Dim vStrSQL, vRandomID As String
Dim ExpenseFlag As Boolean
Dim vExpAmount As Double
Dim vQtyLoose As Double
Dim vColour, vSerialAdd, vShowStock, isPrice As Boolean
Dim vMobileNo() As String, vMobile As String
Dim i As Integer
Dim vNoofPrints As Byte
Dim vPrinter() As String
'----------------------------------

Private Sub SubCalculateBody()
   TxtDiscVal.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) * Val(TxtDiscPer.Text) / 100, 2)
   If Val(TxtDiscVal.Text) = 0 Then TxtDiscVal.Text = ""
   TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), 2)
'    If vIsWSSaleTax = True And vIsWSDiscb4ST = True Then
'        TxtSaleTaxVal.Text = Round(Val(TxtAmount.Text) * Val(TxtSaleTaxPer.Text) / 100, 3)
'    ElseIf vIsWSSaleTax = True And vIsWSDiscb4ST = False Then
'        TxtSaleTaxVal.Text = Round(Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
'    ElseIf vIsRetailSaleTax = True Then
'        TxtSaleTaxVal.Text = Round(Val(vUnitRetailPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
'    Else
'        TxtSaleTaxVal.Text = Round(Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
'    End If
    TxtAmount.Text = Val(TxtAmount.Text) + Val(TxtSaleTaxVal.Text) - Val(TxtOffer.Text)
    TxtRetailAmount.Text = Round((Val(vUnitRetailPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitRetailPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleDiscPer.Text) / 100), 2)
    TxtProfitAmount.Text = Val(TxtRetailAmount.Text) - Val(TxtAmount.Text)
End Sub

Private Sub SubCalculateFooter()
   If TxtTotalAmount.Text = "" Then Exit Sub
   TxtNetAmount.Text = SelfRound(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtTotalExpense.Text) + Val(TxtOtherCharges.Text)
   TxtTotalPayable.Text = Abs(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", Val(TxtPreviousPayable.Text) * -1, TxtPreviousPayable.Text)))
   LblTtlPayable.Caption = IIf(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", Val(TxtPreviousPayable.Text) * -1, TxtPreviousPayable.Text)) < 0, "Total Payable", "Total Receivable")
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

Private Function FunSelectVender(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
'        SchAccounts.CmbFilter = "Vendors"
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectVender = False: Exit Function
        TxtVenderID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select c.AccountNo, c.AccountName as AccountName, Address, City" & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " where c.AccountNo = " & Val(TxtVenderID.Text) & " and (c.AccountNo like '6%') and isDetailed = 1 and isLocked = 0"
    
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtVenderName.Text = !AccountName
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtPreviousPayable.Text = CN.Execute("SELECT isnull(dbo.FunCurrentDebit(" & Val(TxtVenderID.Text) & ",'" & DtpReturnDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
          lblPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Previous Receivable", "Previous Payable")
          TxtPreviousPayable.Text = Abs(TxtPreviousPayable.Text)
          FunSelectVender = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectVender = False
          .Close
          TxtVenderID.Text = ""
          TxtVenderName.Text = ""
          TxtAddress.Text = ""
          TxtCity.Text = ""
          TxtPreviousPayable.Text = ""
          lblPayable.Caption = "Previous Payable"
          LblTtlPayable.Caption = "Total Payable"
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
       If vColour = True Then
         SchProduct.ParaInWhere = " and isLocked = 0 and isNoCostProduct = 0 and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
         SchItemCode.Show vbModal, Me
         TxtCode.Text = SchItemCode.ParaOutItemCode
      Else
         SchProduct.ParaInWhere = " and isLocked = 0 and isNoCostProduct = 0 and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
         SchProduct.ParainShowStock = vShowStock
         SchProduct.Show vbModal, Me
         If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
         TxtCode.Text = SchProduct.ParaOutID
      End If
   End If
    '---------------------------
   If TxtCode.Enabled = False Then FunSelectProduct = False: Exit Function
   If Trim(TxtCode.Text) = "" Then FunSelectProduct = False: Exit Function
   
  If vColour = True Then
      ssql = "select c.ColourID, ColourName from productcolours pc inner join Colours c on pc.colourid = c.colourid " & vbCrLf _
             & "inner join products p on p.productid = pc.productid " & vbCrLf _
             & "where ItemCode = '" & IIf(Len(TxtCode.Text) = 9, TxtCode.Text & "'", Mid(TxtCode.Text, 1, 9) & "' and c.colourid = " & Val(Mid(TxtCode.Text, 10, 2)))
      With CN.Execute(ssql)
         If .RecordCount > 0 Then
            CmbColourName.AddItem !ColourName
            CmbColourName.ItemData(CmbColourName.NewIndex) = !ColourID
            CmbColourName.ListIndex = 0
         End If
      End With
      
      ssql = "select s.SizeID, SizeName from productSizes pz inner join Sizes s on pz.Sizeid = s.Sizeid " & vbCrLf _
      & "inner join products p on p.productid = pz.productid " & vbCrLf _
      & "where ItemCode = '" & IIf(Len(TxtCode.Text) = 13, Mid(TxtCode.Text, 1, 9) & "' and s.sizeid = " & Val(Mid(TxtCode.Text, 12, 2)), TxtCode.Text & "'")
      With CN.Execute(ssql)
         If .RecordCount > 0 Then
            cmbSizeName.AddItem !SizeName
            cmbSizeName.ItemData(cmbSizeName.NewIndex) = !SizeID
            cmbSizeName.ListIndex = 0
         End If
      End With
      TxtCode.Text = CStr(Left(TxtCode.Text, 9))
   End If
   
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
   
   ''''''''''''' Serail '''''''''''''''''''''''''''''''''
   vSerialAdd = False
   vStrSQL = "Select ProductID, Serial, SerialAdd from vuPurchaseSerial where Serial = '" & Trim(TxtCode.Text) & "' or ProductID = " & Val(TxtCode.Text)
   With CN.Execute(vStrSQL)
      If .EOF = False Then
            If Frame3.Visible = False Then
               Frame3.Visible = True
               Frame3.ZOrder 0
            End If
            TxtSerial.Text = TxtCode.Text
            TxtCode.Text = !Productid
            GetDataFromTexBoxesToGridSerial
            If vSerialAdd = False Then
               TxtCode.Text = ""
               FunSelectProduct = False
               Exit Function
            End If
      End If
   End With
 '''''''''''''''''''''''''''''''''''''''''''''
        vStrSQL = " SELECT p.productid, Code, ProductName, PurPrice, RetailPrice, DiscPer, PurDiscPC, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.Purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')" & " and isLocked = 0 "
 
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !PurPrice
         TxtRetailPrice.Text = !RetailPrice
         TxtSaleDiscPer.Text = IIf(IsNull(!DiscPer), "", !DiscPer)
         If Len(TxtCode.Text) > 5 Then
            TxtQtyLoose.Text = IIf(vAutoEnterBeforeQty, "1", TxtQtyLoose.Text)
         End If
         LblRetailPrice.Caption = !RetailPrice
         If IsNull(!PackingName) Then
            vUnitPrice = !PurPrice
            vUnitRetailPrice = !RetailPrice
            TxtMultiplier.Text = ""
            CmbPackName.ListIndex = 0
         Else
            TxtMultiplier.Text = !Multiplier
            If !Multiplier <> 0 Then
               vUnitPrice = !PurPrice / !Multiplier
               vUnitRetailPrice = !RetailPrice / !Multiplier

            Else
               vUnitPrice = !PurPrice
               vUnitRetailPrice = !RetailPrice
            End If
            CmbPackName.Text = !PackingName
         End If
         TxtDiscPC.Text = IIf(IsNull(!PurDiscPC), "", !PurDiscPC)
         If vUnitPrice = 0 Then
            TxtDiscPer.Text = "0"
         Else
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
         End If
         If ObjRegistry.AutoApplyPartyLastPrice Then
            If Trim(TxtVenderID.Text) <> "" Then
               vStrSQL = "Select Top 1 * " & vbCrLf _
                     + " from PurchaseHeader h inner join PurchaseBody b on h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate" & vbCrLf _
                     + " left outer join ProductPacking pp on pp.packingid = b.packingid and pp.productid = b.productid" & vbCrLf _
                     + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
                     + " where h.VendorID = " & Val(TxtVenderID.Text) & " and b.ProductID = " & Val(TxtProductID.Text) & vbCrLf _
                     + " Order by h.PurchaseDate Desc, h.PurID desc"
               With CN.Execute(vStrSQL)
                  If .RecordCount > 0 Then
                     If IsNull(!PackingName) Then
                        vUnitPrice = !Price
                        TxtMultiplier.Text = ""
                        CmbPackName.ListIndex = 0
                     Else
                        If !Multiplier <> 0 Then
                           vUnitPrice = !Price / !Multiplier
                        End If
                        CmbPackName.Text = !PackingName
                     End If
                     TxtMultiplier.Text = IIf(IsNull(!Multiplier), "", !Multiplier)
                     TxtPrice.Text = !Price
                     TxtDiscPC.Text = !DiscPC
                     TxtDiscPer.Text = !DiscPer
                  End If
               End With
            End If
         End If
         
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpPurchaseDate.DateValue + 1 & "',0),0)"
          With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               vQtyLoose = .Fields(0).Value
            Else
               vQtyLoose = 0
            End If
         End With
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         
'         With CN.Execute("select Productid, QtyLoose from CurrentStockStore where ProductID ='" & TxtProductID.Text & "' and StoreID = " & TxtStoreID.Text)
'            If .RecordCount > 0 Then
'               LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack('" & !Productid & "',Floor(" & !QtyLoose & "))").Fields(0).Value
'               'LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'               LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose('" & !Productid & "',Floor(" & !QtyLoose & "))").Fields(0).Value
'               'LblStock.Caption = LblStock.Caption & " " & "Loose"
'            Else
'               LblStock.Caption = 0
'            End If
'         End With
         If ObjRegistry.NegativeSale = False Then
            If Val(LblStock.Caption) <= 0 Then
               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
               FunSelectProduct = False
               Exit Function
            End If
         End If
         
         PopulateDataToHistoryGrid
         FrmHistory.Left = Grid.Left
         FrmHistory.Visible = True
         FrmHistory.ZOrder 0
         GridHistory.Visible = True
         GridHistory.ZOrder 0
         
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
         
         If ObjRegistry.ShowAllPrices Then
            PopulateDataToPriceGrid
            FrmProductPrices.Visible = True
         Else
            FrmProductPrices.Visible = False
         End If
         If vSerialAdd = True Then TxtQtyLoose.Text = 1
         SubCalculateBody
         If vSerialAdd = True Then GetDataFromTexBoxesToGrid
         'Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         FrmProductPrices.Visible = False
         TxtProductID.Text = ""
         TxtCode.Text = ""
         If CmbPackName.ListCount > 0 Then CmbPackName.ListIndex = 0
         TxtProductName.Text = ""
         TxtMultiplier.Text = ""
         TxtPrice.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtAmount.Text = ""
         LblStock.Visible = False
         LblStockCaption.Visible = False
         LblCaptionRetailPrice.Visible = False
         LblRetailPrice.Visible = False
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
'   Call DeleteTempActivityLogBin(vRandomID)
    vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.Rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From PurchaseReturnbody where ReturnID=" & Val(TxtReturnID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Cleared Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
      
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eClearSavedRecord, TxtReturnID.Text, DtpReturnDate.DateValue, vGridRows & " Product/s Cleared ")
      Grid.Redraw = True
'   cn.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()

 '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
      On Error GoTo ErrorHandler
        
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniPurchaseReturnInvoice", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   
    '''''''''''''''''''''''Check Import / Export'''''''''''''''''''''''''''''''''
    If ObjRegistry.ShowMultiBranches = True Then
      vStrSQL = "select * from PurchaseReturnHeader where Tag is not null And SID=" & Val(TxtReturnID.Text) & " and Returndate='" & DtpReturnDate.DateValue & "'"
      With CN.Execute(vStrSQL)
          If Not .EOF Then
              MsgBox "Import/Export Record Cannot be Deleted", vbInformation, Me.Caption
              Exit Sub
          End If
      End With
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   CN.BeginTrans
   Call BinData
   Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eDelete, TxtReturnID.Text, DtpReturnDate.DateValue, Grid.Rows - 1 & " Product/s Deleted Amount: " & Val(TxtNetAmount.Text))
'
'    vMaxBinID = FunGetMaxBinID
'
'   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_PurchaseReturnHeader Select " & vMaxBinID & ",'" & Date & "',* from PurchaseReturnHeader Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
'    '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_PurchaseReturnBody Select " & vMaxBinID & ",'" & Date & "', * from PurchaseReturnBody Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Detail''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_PurchaseReturnSerial Select " & vMaxBinID & ",'" & Date & "', * from PurchaseReturnSerial Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin ProductOffer''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_PurchaseBodyOffer Select " & vMaxBinID & ",'" & Date & "', * from PurchaseReturnOffer Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
'
'   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
'   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
   ''''''''''''''''''''''''''Delete Product OFfer'''''''''''''''''''''
   GridOffer.Redraw = False
   GridOffer.MoveFirst
   For vCounter = 1 To GridOffer.Rows
      If Trim(GridOffer.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from PurchaseReturnOffer where ReturnID = " & Val(TxtReturnID.Text) & " And ReturnDate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(GridOffer.Columns("Productid").Text)
      End If
      GridOffer.MoveNext
   Next vCounter
   GridOffer.RemoveAll
   GridOffer.Redraw = True
   GridOffer.ZOrder 1
   
   ''''''''''''''''''' Delete salebodyserial '''''''''''''''''''''''
   RsBodySerial.Filter = ""
   
   While Not RsBodySerial.EOF
      RsPurchaseSerial.Filter = ""
      RsPurchaseSerial.Filter = "Serial = " & RsBodySerial!serial
      If RsPurchaseSerial.RecordCount > 0 Then RsPurchaseSerial!SerialAdd = 1
      
      CN.Execute "Delete from PurchaseReturnSerial where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Productid").Text)
      RsBodySerial.MoveNext
   Wend
      
   
'   RsBodySerial.Filter = ""
'   If RsBodySerial.RecordCount > 0 Then RsBodySerial.UpdateBatch
'
   RsPurchaseSerial.Filter = ""
   If RsPurchaseSerial.RecordCount > 0 Then RsPurchaseSerial.UpdateBatch
   
   Grid.Redraw = False
   Grid.MoveFirst
   Call ActivityLog("Purchase Return Invoice", eDelete, TxtReturnID.Text, DtpReturnDate.DateValue)
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from PurchaseReturnBody where SID = " & Val(TxtSID.Text) & " and Returndate='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("ProductID").Text) & " and BatchNo " & IIf(Trim(Grid.Columns("BatchNo").Text) = "", " is null", " = '" & Trim(Grid.Columns("BatchNo").Text) & "'") & " and Price = " & Val(Grid.Columns("Price").Text)
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
    '''''''''''''''''''''''''''''''''''''''Delete Expense'''''''''''''''''''''''''''''''''''''''
   CN.Execute "Delete from PurchaseReturnExpense where ReturnID = " & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"

   CN.Execute "Delete from PurchaseReturnHeader where SID = " & Val(TxtSID.Text) & " and Returndate='" & DtpReturnDate.DateValue & "'"
   
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnDelete Then
   vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = ObjRegistry.PrefixPhoneNo + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = ObjUserSecurity.UserName & " " & FrmPurchaseReturnInvoice.Caption & " Deleted ID:" & TxtReturnID.Text & vbCrLf & " Date:" & Format(DtpReturnDate.DateValue, "dd-MMM-yyyy") & " Time: " & Time & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt" & TxtNetAmount.Text
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & "','')"
               CN.Execute ssql
            End If
         Next
   End If
   
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchPurchaseReturn.ParaInReturnDate = DtpReturnDate.DateValue
   SchPurchaseReturn.Show vbModal
   If SchPurchaseReturn.ParaOutReturnID <> 0 Then
      TxtSID.Text = SchPurchaseReturn.ParaOutSID
      TxtReturnID.Text = SchPurchaseReturn.ParaOutReturnID
      DtpReturnDate.DateValue = SchPurchaseReturn.ParaOutReturnDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetReturn
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

Private Sub PopulateTempToGrid()
   On Error GoTo ErrorHandler
   With RsTemp
      RsTemp.MoveFirst
      Grid.Redraw = False
      'Grid.MoveLast
      'Grid.RemoveAll
      Grid.AllowAddNew = True
      'TxtTotalAmount.Text = 0
      While Not .EOF
         Grid.Columns("ProductID").Text = !Productid
         Grid.Columns("Code").Text = !Productid
         Grid.Columns("ProductName").Text = !ProductName
         
         RsBody.AddNew
         RsBody!Productid = Val(!Productid)
         RsBody!Code = !Productid
         
         Grid.Columns("QtyLoose").Value = !QtyLoose
         Grid.Columns("Price").Value = !Price
         
'         Grid.Columns("RetailPrice").Value = 0
'         Grid.Columns("IsWSDiscb4ST").Value = 0
'         Grid.Columns("IsWSSaleTax").Value = 0
'         Grid.Columns("IsRetailSaleTax").Value = 0
         
         Grid.Columns("Amount").Value = (!Price * !QtyLoose)

         '''''
         RsBody!Multiplier = Null
         RsBody!QtyPack = Null
         RsBody!QtyLoose = !QtyLoose
         RsBody!Bonus = Null
         RsBody!Price = !Price
         
 '        RsBody!RetailPrice = 0
'         RsBody!IsWSDiscb4ST = 0
'         RsBody!IsWSSaleTax = 0
'         RsBody!IsRetailSaleTax = 0
         
         RsBody!DiscPC = 0
         RsBody!Offer = Null
         RsBody!SaleTaxPer = Null
         RsBody!SaleTaxval = Null
         RsBody!DiscPer = 0
         RsBody!DiscVal = 0
         RsBody!Amount = (!Price * !QtyLoose)
         RsBody.Update
         ''''
         
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + (!Price * !QtyLoose)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + !QtyLoose
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

Private Sub BtnProductRange_Click()
   On Error GoTo ErrorHandler
   FrmProductRangeGrid.Show vbModal, Me
   RsTemp.Filter = ""
   If RsTemp.RecordCount > 0 Then
      PopulateTempToGrid
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPurchase_Click()
   On Error GoTo ErrorHandler
   SchPurchase.ParaInPurchasedate = DtpPurchaseDate.DateValue
   SchPurchase.Show vbModal
   If SchPurchase.ParaOutPurchaseID <> "" Then
      TxtPurchaseID.Text = SchPurchase.ParaOutPurchaseID
      'Dim a
      'a = Split(SchSale.ParaOutBillDate, "/")
      DtpPurchaseDate.DateValue = SchPurchase.ParaOutPurchaseDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
'      CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetPurchase
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetPurchase()
   On Error GoTo ErrorHandler
   ssql = "Select h.*, OrganizationName, p.partyname, Address, City, StoreName FROM PurchaseHeader h join parties p on h.Vendorid = p.partyid inner join stores s on s.storeid = h.storeid left outer join Organizations o on o.OrganizationID = h.OrganizationID where h.PurID=" & Val(TxtPurchaseID.Text) & " and PurchaseDate='" & DtpPurchaseDate.DateValue & "'"
   With CN.Execute(ssql)
      If Not .BOF Then
          TxtVenderID.Text = !vendorID
          TxtVenderName.Text = !PartyName
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
'          TxtBillNo.Text = IIf(IsNull(!BillNo), "", !BillNo)
'          TxtBiltyNo.Text = IIf(IsNull(!BiltyNo), "", !BiltyNo)
'          TxtVehicleNo.Text = IIf(IsNull(!VehicleNo), "", !VehicleNo)
          TxtTotalAmount.Text = !TotalAmount
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtTotalExpense.Text = IIf(IsNull(!TotalExpense), "", !TotalExpense)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
      End If
      .Close
   End With
   Call PopulatePurchaseToGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulatePurchaseToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from PurchaseReturnBody where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      ssql = " select p.ProductID, ProductName, sob.*" & vbCrLf _
      + " from Purchasebody sob inner join Products p on p.ProductID = sob.productid" & vbCrLf _
      + " where sob.PurID = " & Val(TxtPurchaseID.Text) & " and sob.PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"

      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = !Productid
            Grid.Columns("ProductName").Text = !ProductName
            
            RsBody.AddNew
            RsBody!Productid = !Productid
            RsBody!Code = !Productid
            If vColour = True Then
            If !ColourID = 0 Or IsNull(!ColourID) Then
               Grid.Columns("ColourID").Value = ""
            Else
               Grid.Columns("ColourID").Value = !ColourID
            End If
            If !ColourID = 0 Or IsNull(!ColourID) Then
               Grid.Columns("ColourName").Text = ""
            Else
               Grid.Columns("ColourName").Text = CN.Execute("Select ColourName from Colours where ColourID=" & !ColourID).Fields(0).Value
            End If
            If !SizeID = 0 Or IsNull(!SizeID) Then
               Grid.Columns("SizeID").Value = ""
            Else
               Grid.Columns("SizeID").Value = !SizeID
            End If
            If !SizeID = 0 Or IsNull(!SizeID) Then
               Grid.Columns("SizeName").Text = ""
            Else
               Grid.Columns("SizeName").Text = CN.Execute("Select SizeName from Sizes where SizeID=" & !SizeID).Fields(0).Value
            End If
            RsBody!ColourID = !ColourID
            RsBody!SizeID = !SizeID
            End If
            
            If !PackingID = 0 Or IsNull(!PackingID) Then
               Grid.Columns("PackingID").Value = ""
            Else
               Grid.Columns("PackingID").Value = !PackingID
               RsBody!PackingID = !PackingID
            End If
            If !PackingID = 0 Or IsNull(!PackingID) Then
               Grid.Columns("PackName").Text = ""
            Else
               Grid.Columns("PackName").Text = CN.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
            End If
            Grid.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            Grid.Columns("QtyLoose").Value = !QtyLoose
            Grid.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            Grid.Columns("Price").Value = !Price
            
'            Grid.Columns("RetailPrice").Value = !RetailPrice
'            Grid.Columns("IsWSDiscb4ST").Value = !IsWSDiscb4ST
'            Grid.Columns("IsWSSaleTax").Value = !IsWSSaleTax
'            Grid.Columns("IsRetailSaleTax").Value = !IsRetailSaleTax
            
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("RetailPrice").Value = !RetailPrice
            Grid.Columns("SaleDiscPer").Value = IIf(IsNull(!SaleDiscPer), "", !SaleDiscPer)
            Grid.Columns("RetailAmount").Value = IIf(IsNull(!RetailAmount), "", !RetailAmount)
            Grid.Columns("ProfitAmount").Value = IIf(IsNull(!ProfitAmount), "", !ProfitAmount)

            '''''
            RsBody!Multiplier = IIf(IsNull(!Multiplier), Null, !Multiplier)
            RsBody!QtyPack = IIf(IsNull(!QtyPack), Null, !QtyPack)
            RsBody!QtyLoose = !QtyLoose
            RsBody!Bonus = IIf(IsNull(!Bonus), Null, !Bonus)
            RsBody!Price = !Price
            
'            RsBody!RetailPrice = !RetailPrice
'            RsBody!IsWSDiscb4ST = !IsWSDiscb4ST
'            RsBody!IsWSSaleTax = !IsWSSaleTax
'            RsBody!IsRetailSaleTax = !IsRetailSaleTax
            
            RsBody!DiscPC = IIf(IsNull(!DiscPC), Null, !DiscPC)
            RsBody!Offer = IIf(IsNull(!Offer), Null, !Offer)
            RsBody!SaleTaxPer = IIf(IsNull(!SaleTaxPer), Null, !SaleTaxPer)
            RsBody!SaleTaxval = IIf(IsNull(!SaleTaxval), Null, !SaleTaxval)
            RsBody!DiscPer = IIf(IsNull(!DiscPer), Null, !DiscPer)
            RsBody!DiscVal = IIf(IsNull(!DiscVal), "", !DiscVal)
            RsBody!Amount = !Amount
            RsBody!RetailPrice = !RetailPrice
            RsBody!SaleDiscPer = IIf(IsNull(!SaleDiscPer), Null, !SaleDiscPer)
            RsBody!RetailAmount = IIf(IsNull(!RetailAmount), Null, !RetailAmount)
            RsBody!ProfitAmount = !ProfitAmount
            RsBody.Update
            ''''
            
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + !Amount
            TxtTotalItems.Text = Val(TxtTotalItems.Text) + !QtyLoose + IIf(IsNull(!Bonus), "0", !Bonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack))
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
'   End If
   
'   RsBodySerial.Filter = 0
'   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
'   RsBodySerial.Open "Select * from PurchaseBodySerial where PurID=" & Val(TxtPurchaseID.Text) & " and Purchasedate = '" & DtpPurchaseDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
'   Call PopulatePOToGridOffer
'   Call PopulatePOToGridSerial
'   Call PopulatePOToGridExpense
End Sub

Private Sub GridExpense_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   DispPromptMsg = 0
End Sub

Private Sub SubDetailUpdate()
   On Error GoTo ErrorHandler
   RsExpense.Filter = "ExpenseID='" & GridExpense.Columns("ID").Text & "'"
   If RsExpense.RecordCount = 0 Then
      RsExpense.AddNew
      RsExpense!ExpenseID = GridExpense.Columns("ID").Text
      RsExpense!ExpAmount = GridExpense.Columns("Value").Value
   ElseIf RsExpense.RecordCount = 1 And GridExpense.Columns("Value").Value = 0 Then
      RsExpense.Delete
   ElseIf RsExpense.RecordCount = 1 Then
      RsExpense!ExpAmount = GridExpense.Columns("Value").Text
      RsExpense.Update
   End If
   TxtTotalExpense.Text = Val(TxtTotalExpense.Text) + GridExpense.Columns("Value").Value - vExpAmount
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   ExpenseFlag = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridExpense_BeforeUpdate(Cancel As Integer)
   If GridExpense.Visible = False Then Exit Sub
   If ActiveControl.Name <> GridExpense.Name Then Exit Sub
   On Error GoTo ErrorHandler
   SubDetailUpdate
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridExpense_Change()
 ExpenseFlag = True
 If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub GridExpense_GotFocus()
'   RsBody.Filter = "ExpenseID='" & GridExpense.Columns("ID").Text & "'"
   GridExpense.Row = 0
   GridExpense.Col = 0
   SendKeys "{Right}"
End Sub

Private Sub GridExpense_LostFocus()
   On Error GoTo ErrorHandler
   If ExpenseFlag = True Then SubDetailUpdate
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridExpense_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(GridExpense.Columns("ID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub GridExpense_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If GridExpense.Visible = False Then Exit Sub
   If ActiveControl.Name <> GridExpense.Name Then Exit Sub
   vExpAmount = Val(GridExpense.Columns("Value").Value)
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
      If TxtVenderID.Enabled Then TxtVenderID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   
'   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
      
   If MsgBox("Do you want to print this invoice With Amount", vbQuestion + vbYesNo, "Alert") = vbYes Then
      isPrice = True
   Else
      isPrice = False
   End If
   
'   If InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
'
'      vStrSQL = " select UserName, h.ReturnID, h.ReturnDate, Pr.PartyName + ' - ' + H.VendorID as Vend_Name_ID, Pr.Address, pr.City, H.StoreID, StoreName, ServerEntry as EntryTime, H.BillNo, h.Description, " & vbCrLf _
'           + " Isnull(H.TotalAmount,0) TotalAmount, b.ProductID as code, ProductName, companyName, dbo.FunPurchaseReturnSerial(b.ReturnID,b.ReturnDate,b.ProductId) Serial," & vbCrLf _
'           + " dbo.FunPurchaseReturnOffer(b.ReturnID,b.ReturnDate, b.ProductId) ProductOffer, QtyPack, Multiplier, QtyLoose," & vbCrLf _
'           + " P.PurPrice, Isnull(B.Price,0) Price, Bonus, b.DiscPc, b.DiscPer, DiscVal, Offer, b.SaleTaxPer, SaleTaxval,Isnull(B.Amount,0) Amount, RetailPrice, isnull(BillDisc,0) as BillDisc, " & vbCrLf _
'           + " isnull(OtherCharges,0) as OtherCharges, BiltyNo, p.ItemCode, right('00'+ cast(b.ColourID as varchar(2)),2) as ColourID, right('00'+ cast(b.SizeID as varchar(2)),2) as SizeID, ColourName, SizeName, " & vbCrLf _
'           + " Isnull(H.ReceivedAmount,0) ReceivedAmount, Isnull(H.PreviousAmount,0) PreviousAmount, isnull(PK.PackingName,'PC') as PackingName," & vbCrLf _
'           + " Sec.SectorName , z.ZoneID, ZoneName, pr.PartyName, isnull(pr.Address,'') + isnull(' - ' + pr.City,'') + isnull(',' + pr.Phone1,'') + isnull(',' + pr.Phone2, '') + isnull(',' + pr.Mobile, '') as AddressFull" & vbCrLf _
'           + " from PurchaseReturnBody b inner join products p on b.productid = p.productid" & vbCrLf _
'           + " inner join PurchaseReturnHeader h on h.ReturnID = b.ReturnID and h.ReturnDate = b.ReturnDate" & vbCrLf _
'           + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
'           + " inner join parties pr on pr.partyid = h.VendorID" & vbCrLf _
'           + " left outer join companies c on c.companyid = p.companyid " & vbCrLf _
'           + " Left Outer jOin Groups g on g.Groupid = p.Groupid" & vbCrLf _
'           + " Left Outer jOin SubGroups sg on sg.subGroupid = p.subGroupid" & vbCrLf _
'           + " Left Outer jOin Brands bd on bd.brandid = p.brandid" & vbCrLf _
'           + " Left Outer jOin Seasons se on se.Seasonid = p.Seasonid " & vbCrLf _
'           + " Left Join Sectors Sec on Sec.SectorID = pr.SectorID" & vbCrLf _
'           + " Left Join Zones z on z.ZoneID = Sec.ZoneID" & vbCrLf _
'           + " left outer Join Packings PK on Pk.PackingID = B.PackingID" & vbCrLf _
'           + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
'           + " Left outer join Colours Col on Col.Colourid = b.ColourID" & vbCrLf _
'           + " Left Outer join Sizes Sz on Sz.SizeID = b.SizeID " & vbCrLf _
'           + " where h.ReturnID =" & TxtReturnID.Text & " and h.ReturnDate='" & DtpReturnDate.DateValue & "'"
'   Else
'      vStrSQL = "  select UserName, h.ReturnID as billid, h.ReturnDate as BillDate, isnull(h.Description,'') as Remarks, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount," & vbCrLf _
'            + " isnull(h.ReceivedAmount,0) as CashReceived, p.ProductName, companyName, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + QtyLoose as Qty, b.price/isnull(multiplier,1) as price, b.amount, b.DiscPC, b.DiscPer, b.DiscVal, AccountName as Customer, b.ProductID, p.ItemCode, right('00'+ cast(b.ColourID as varchar(2)),2) as ColourID, right('00'+ cast(b.SizeID as varchar(2)),2) as SizeID, ColourName, SizeName" & vbCrLf _
'            + " from PurchaseReturnHeader h inner join PurchaseReturnBody b on h.ReturnID = b.ReturnID and h.ReturnDate = b.ReturnDate" & vbCrLf _
'            + " inner join products p on p.productid = b.productid" & vbCrLf _
'            + " left outer join companies c on c.companyid = p.companyid" & vbCrLf _
'            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
'            + " left outer join ChartofAccounts ca on ca.AccountNo = h.VendorID" & vbCrLf _
'            + " Left outer join Colours Col on Col.Colourid = b.ColourID" & vbCrLf _
'            + " Left Outer join Sizes Sz on Sz.SizeID = b.SizeID " & vbCrLf _
'            + " where h.ReturnID =" & TxtReturnID.Text & " and h.ReturnDate = '" & DtpReturnDate.DateValue & "' Order By SerialNo"
'   End If
'   + " P.PurPrice, Isnull(B.Price,0) Price, Bonus, b.DiscPc, b.DiscPer, DiscVal, Offer, b.SaleTaxPer, SaleTaxval,Isnull(B.Amount,0) Amount, b.RetailPrice, isnull(BillDisc,0) as BillDisc, "
    vStrSQL = " select UserName, h.ReturnID, h.ReturnDate, Pr.PartyName + ' - ' + cast(H.VendorID as varchar(10)) as Vend_Name_ID, Pr.Address, pr.City, isnull( pr.Phone1  + ', ','') + isnull( pr.Phone2 + ', ','')  + isnull( pr.mobile + ', ','') +  isnull( pr.mobile2 + ', ','') as Moblie, H.StoreID, StoreName, ServerEntry as EntryTime, H.BillNo, h.Description, " & vbCrLf _
           + " Isnull(H.TotalAmount,0) TotalAmount, b.ProductID, b.code, ProductName, companyName, dbo.FunPurchaseReturnSerial(b.ReturnID,b.ReturnDate,b.ProductId) Serial," & vbCrLf _
           + " dbo.FunPurchaseReturnOffer(b.ReturnID,b.ReturnDate, b.ProductId) ProductOffer, QtyPack, Multiplier, QtyLoose," & vbCrLf _
           + IIf(isPrice = True, " b.DiscPc, b.DiscPer, DiscVal, Offer, b.SaleTaxPer, SaleTaxval, ", " 0 DiscPc, 0 DiscPer, 0 DiscVal, 0 Offer, 0 SaleTaxPer, 0 SaleTaxval, ") & vbCrLf _
           + IIf(isPrice = True, " P.PurPrice, b.RetailPrice, isnull(BillDisc,0) as BillDisc,  Isnull(B.Price,0) price, ", " 0 PurPrice, 0 RetailPrice, 0 as BillDisc, 0 price,") & vbCrLf _
           + IIf(isPrice = True, " Isnull(B.Amount,0) Amount,", " 0 amount,") & vbCrLf _
           + " isnull(OtherCharges,0) as OtherCharges, BiltyNo, p.ItemCode, right('00'+ cast(b.ColourID as varchar(2)),2) as ColourID, right('00'+ cast(b.SizeID as varchar(2)),2) as SizeID, ColourName, SizeName, " & vbCrLf _
           + " Isnull(H.ReceivedAmount,0) ReceivedAmount, Isnull(H.PreviousAmount,0) PreviousAmount, isnull(PK.PackingName,'PC') as PackingName," & vbCrLf _
           + " Sec.SectorName , z.ZoneID, ZoneName, pr.PartyName, isnull(pr.Address,'') + isnull(' - ' + pr.City,'') + isnull(',' + pr.Phone1,'') + isnull(',' + pr.Phone2, '') + isnull(',' + pr.Mobile, '') as AddressFull" & vbCrLf _
           + " from PurchaseReturnBody b inner join products p on b.productid = p.productid" & vbCrLf _
           + " inner join PurchaseReturnHeader h on h.ReturnID = b.ReturnID and h.ReturnDate = b.ReturnDate"
           vStrSQL = vStrSQL & " inner join stores s on s.storeid = h.storeid" & vbCrLf _
           + " inner join parties pr on pr.partyid = h.VendorID" & vbCrLf _
           + " left outer join companies c on c.companyid = p.companyid " & vbCrLf _
           + " Left Outer jOin Groups g on g.Groupid = p.Groupid" & vbCrLf _
           + " Left Outer jOin SubGroups sg on sg.subGroupid = p.subGroupid" & vbCrLf _
           + " Left Outer jOin Brands bd on bd.brandid = p.brandid" & vbCrLf _
           + " Left Outer jOin Seasons se on se.Seasonid = p.Seasonid " & vbCrLf _
           + " Left Join Sectors Sec on Sec.SectorID = pr.SectorID" & vbCrLf _
           + " Left Join Zones z on z.ZoneID = Sec.ZoneID" & vbCrLf _
           + " left outer Join Packings PK on Pk.PackingID = B.PackingID" & vbCrLf _
           + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
           + " Left outer join Colours Col on Col.Colourid = b.ColourID" & vbCrLf _
           + " Left Outer join Sizes Sz on Sz.SizeID = b.SizeID " & vbCrLf _
           + " where h.ReturnID =" & TxtReturnID.Text & " and h.ReturnDate='" & DtpReturnDate.DateValue & "'"
   
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   
   If cmbPrintType.Text = "Half Page" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseReturnInvoiceeHalf1.rpt")
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.x
      RptReportViewer.Report.RightMargin = 225
   ElseIf cmbPrintType.Text = "Thermal" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpPurchaseReturnInvoiceAurora.rpt")
      RptReportViewer.Report.TopMargin = 0
      RptReportViewer.Report.LeftMargin = 0
      RptReportViewer.Report.RightMargin = 0
   Else
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseReturnInvoice.rpt")
   End If
   
   
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port

   
'   If InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
'      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseReturnInvoice.rpt")
'      RptReportViewer.Report.PaperSize = crPaperA4
'      RptReportViewer.Report.PaperOrientation = crLandscape
'      RptReportViewer.Report.TopMargin = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))
'      RptReportViewer.Report.LeftMargin = IIf(IsNull(ObjRegistry.x), 0, Val(ObjRegistry.x))
'      RptReportViewer.Report.RightMargin = 225
'   Else
''      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpPurchaseReturnInvoiceAurora.rpt")
'      RptReportViewer.Report.LeftMargin = 225
'      RptReportViewer.Report.RightMargin = 0
'      RptReportViewer.Report.TopMargin = 255
'   End If
   
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ReportTitle = "Purchase Return Invoice"

   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
  
'   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port

   
   vPrinter = Split(CmbPrinters.Text, ",")
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   
  
'   RptReportViewer.Report.ParameterFields(4).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
   'RptReportViewer.Report.PrintOut False
   CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
   If ObjRegistry.PreviewSaleInoice = True Or ChkIsPreview.Value = 1 Then
      If ChkIsPrint.Value = 1 Then
         RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
      End If
       RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      If TxtBatchNo.Visible Then TxtBatchNo.SetFocus Else GetDataFromTexBoxesToGrid
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniPurchaseReturnInvoice", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
   
    ''''''''''''''''''''''''Check Organization '''''''''''''''''''''''''''''''''
  If ObjRegistry.OrganizationMandatory = True And TxtOrganizationID.Text = "" Then
    MsgBox "Please Select Organization", vbInformation, Me.Caption
    If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
    Exit Sub
  End If
  
'  Header Validation
   If Trim(TxtVenderID.Text) = "" Then
      MsgBox "Enter Vender ID.", vbExclamation, Me.Caption
      TxtVenderID.SetFocus
      Exit Sub
   End If
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      If TxtStoreID.Visible And TxtStoreID.Enabled Then TxtStoreID.SetFocus
      Exit Sub
   End If
    '''''''''''''''''''''''Check Entry Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.isEntryDate = True Then
       If ObjRegistry.FromDate > Date Or ObjRegistry.ToDate < Date Then
         MsgBox "Data can not be saved Because Date is not set according to the Software's Entry date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
   '''''''''''''''''''''''Check Current Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.CurrentDateDataEntry = True And ObjUserSecurity.IsAdministrator = False Then
       If DtpReturnDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
   Dim vBm As Variant
   Dim i As Integer
   Grid.Redraw = False
   vBm = Grid.Bookmark
   TxtTotalAmount.Text = "0"
   'vTotalAmount = 0
   Grid.MoveFirst
   For i = 0 To Grid.Rows - 1
      TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(Grid.Columns("Amount").CellValue(Grid.GetBookmark(i)))
'      vTotalAmount = vTotalAmount + Val(Grid.Columns("Amount").CellValue(Grid.GetBookmark(i)))
   Next i
   Grid.Bookmark = vBm
   Grid.Redraw = True

   With CN.Execute("select dbo.DefaultValue('Counter Purchase')")
      If TxtVenderID.Text = .Fields(0).Value Then
         TxtReceivedAmount.Text = TxtNetAmount.Text
      End If
   End With
   
   Call SubCalculateFooter

   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
   vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where touserno = " & vUser & " and Entrydate <='" & Date & "'"
   With CN.Execute(vStrSQL)
      If .Fields(0).Value >= DtpReturnDate.DateValue Then
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
   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Return", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
   End If
   
   '''''''''''''''''''''''Check Import / Export'''''''''''''''''''''''''''''''''
    If ObjRegistry.ShowMultiBranches = True Then
      vStrSQL = "select * from PurchaseReturnHeader where Tag is not null And ReturnID=" & Val(TxtReturnID.Text) & " and Returndate='" & DtpReturnDate.DateValue & "'"
      With CN.Execute(vStrSQL)
          If Not .EOF Then
              MsgBox "Import/Export Record Cannot be Updated", vbInformation, Me.Caption
              Exit Sub
          End If
      End With
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
  
   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'Purchase Return Invoice' and LocalComputerName = '" & LocalComputerName & "'"
   If CN.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview, IsPrint ) Values ('" & LocalComputerName & "', 'Purchase Return Invoice','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & "," & ChkIsPrint.Value & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & ", IsPrint = " & ChkIsPrint.Value & " Where FormType = 'Purchase Return Invoice' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   CN.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   
   CN.BeginTrans

   If vIsNewRecord = True Then
      If CN.Execute("Select * from PurchaseReturnHeader where ReturnID = " & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'").RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtReturnID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
   
   ''''''''''''''''' Following Code used for Debug When Make a new Purid in PurcahseBody and New SID in PurchaseHeader
   If Val(TxtReturnID.Text) = 0 Then
      TxtReturnID.Text = FunGetMaxID
      If vIsNewRecord = False Then
         MsgBox "Please take a screen shot and send to SoftInn. PurchaseReturnHeader ReturnID = 0 and vIsNewRecord = False", vbExclamation, "Alert"
         Unload Me
      End If
   End If
   
   ''''''''''''''''' Following Code used for Debug When Make a new Purid in PurcahseBody and New SID in PurchaseHeader
   If vIsNewRecord = False Then
      ssql = "select ReturnID from PurchaseReturnHeader where SID=" & Val(TxtSID.Text)
      If CN.Execute(ssql).Fields(0).Value <> Val(TxtReturnID.Text) Then
         MsgBox "Please take a screen shot and send to SoftInn. PurchaseReturnHeader ReturnID = " & Val(TxtReturnID.Text) & " and SID = " & Val(TxtSID.Text) & " and vIsNewRecord = False", vbExclamation, "Alert"
         Unload Me
      End If
   End If
   
   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Amount: " & Val(TxtNetAmount.Text))
   
'   If vIsNewRecord = False Then Call ActivityLog("Purchase Return Invoice", eEdit, TxtReturnID.Text, DtpReturnDate.DateValue)
   
'   Call UserActivities

   '''' Thsi Statement was writen to remove the bug for ReturnID. Becuase Some Time New ReturnID Saved when edit the record
   If vIsNewRecord = False Then TxtReturnID.Text = CN.Execute("select ReturnID From PurchaseReturnHeader Where SID =" & TxtSID.Text).Fields(0).Value
   
   ssql = "Select * From PurchaseReturnHeader where ReturnID=" & Val(TxtReturnID.Text) & " and Returndate='" & DtpReturnDate.DateValue & "'"
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, CN, adOpenDynamic, adLockPessimistic
      If .BOF Then
         .AddNew
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
         !PurID = IIf(Val(TxtPurchaseID.Text) = 0, Null, TxtPurchaseID.Text)
         !PurchaseDate = DtpPurchaseDate.DateValue
         !UserNo = vUser
      End If
      !vendorID = TxtVenderID.Text
      !StoreID = TxtStoreID.Text
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !BillNo = IIf(TxtBillNo.Text = "", Null, TxtBillNo.Text)
      !BiltyNo = IIf(TxtBiltyNo.Text = "", Null, TxtBiltyNo.Text)
      !VehicleNo = IIf(TxtVehicleNo.Text = "", Null, TxtVehicleNo.Text)
      !TotalAmount = Round(Val(TxtTotalAmount.Text))
      !TotalExpense = IIf(Val(TxtTotalExpense.Text) = 0, Null, Val(TxtTotalExpense.Text))
      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
      !OtherCharges = IIf(Val(TxtOtherCharges.Text) = 0, Null, Val(TxtOtherCharges.Text))
      !ReceivedAmount = IIf(TxtReceivedAmount.Text = "", Null, Val(TxtReceivedAmount.Text))
      !Description = IIf(TxtDescription.Text = "", Null, TxtDescription.Text)
      !PreviousAmount = IIf(lblPayable.Caption = "Previous Receivable", Val(TxtPreviousPayable.Text), Val(TxtPreviousPayable.Text) * -1)
'      !UserNo = vUser
      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
      .Update
      .Close
      If vIsNewRecord = True Then TxtSID.Text = CN.Execute("select @@identity").Fields(0).Value
   End With
   
   
   

   
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !SID = Val(TxtSID.Text)
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   
   RsBodySerial.Filter = 0
   If RsBodySerial.RecordCount > 0 Then
     With RsBodySerial
'      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
                  
         RsPurchaseSerial.Filter = "Serial = " & RsBodySerial!serial
         If RsPurchaseSerial.RecordCount > 0 Then RsPurchaseSerial!SerialAdd = 0
            
         .Update
         .MoveNext
      Next vCounter
      .UpdateBatch
     End With
   End If
   RsPurchaseSerial.Filter = ""
   If RsPurchaseSerial.RecordCount > 0 Then RsPurchaseSerial.UpdateBatch
   
   With RsProductOffer
      .Filter = 0
      If Not .EOF Then
        .MoveFirst
        For vCounter = 1 To .RecordCount
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
         .MoveNext
        Next vCounter
      End If
      .UpdateBatch
   End With
   
   With RsExpense
      .Filter = 0
      If Not .EOF Then
        .MoveFirst
        For vCounter = 1 To .RecordCount
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
         .MoveNext
        Next vCounter
      End If
      .UpdateBatch
   End With
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eAdd, TxtReturnID.Text, DtpReturnDate.DateValue, Grid.Rows - 1 & " New Product/s Added Amount: " & Val(TxtNetAmount.Text))
   
   If vIsNewRecord = True Then Call ActivityLog("Purchase Return Invoice", eAdd, TxtReturnID.Text, DtpReturnDate.DateValue)
   CN.CommitTrans
   
'   If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo + vbDefaultButton1, "Alert") = vbYes Then BtnPrint_Click
   
   If ChkIsPreview.Value = 1 Or ChkIsPrint.Value = 1 Then
      Call BtnPrint_Click
   End If
   
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToHistoryGrid()
      ssql = "select top 3 h.purID, pt.PartyName, VendorID, code, b.* " & vbCrLf & _
      " from PurchaseHeader h inner join Purchasebody b on h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate" & vbCrLf & _
      " inner join Parties pt on pt.PartyID = h.VendorID " & vbCrLf & _
      " where b.productid = " & Val(TxtProductID.Text) & " and h.purchasedate < '" & DtpPurchaseDate.DateValue & "' order by b.PurchaseDate Desc, b.purid Desc"
      
      With CN.Execute(ssql)
         GridHistory.Redraw = False
         GridHistory.MoveFirst
         GridHistory.RemoveAll
         GridHistory.AllowAddNew = True
         
         If ObjRegistry.isShowListPrice Then
            If Not .EOF Then
               LblCaptionRetailPrice = "Last Price"
               LblRetailPrice.Caption = !Price & ", DiscPack = " & Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * Val(IIf(IsNull(!Multiplier), "1", !Multiplier))
            End If
         End If
         
         While Not .EOF
            GridHistory.AddNew
            GridHistory.Columns("ID").Text = !vendorID
            GridHistory.Columns("Name").Text = !PartyName
            GridHistory.Columns("PurID").Text = !PurID
            GridHistory.Columns("Date").Text = !PurchaseDate
'            If !PackingID = 0 Or IsNull(!PackingID) Then
'               GridHistory.Columns("PackingID").Value = ""
'            Else
'               GridHistory.Columns("PackingID").Value = !PackingID
'            End If
'            If !PackingID = 0 Or IsNull(!PackingID) Then
'               GridHistory.Columns("PackName").Text = ""
'            Else
'               GridHistory.Columns("PackName").Text = CN.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
'            End If
            GridHistory.Columns("ExpiryDate").Value = IIf(IsNull(!ExpiryDate), "", !ExpiryDate)
            GridHistory.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
            GridHistory.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            GridHistory.Columns("QtyLoose").Value = !QtyLoose
            GridHistory.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            GridHistory.Columns("Price").Value = !Price
            GridHistory.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
'            GridHistory.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
'            GridHistory.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
'            GridHistory.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            GridHistory.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            GridHistory.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            GridHistory.Columns("Amount").Value = !Amount
            .MoveNext
         Wend
         .Close
         GridHistory.MoveFirst
         GridHistory.Redraw = True
      End With
End Sub

Private Sub PopulateDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from PurchaseReturnBody where SID = " & Val(TxtSID.Text) & " and Returndate = '" & DtpReturnDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.productname,code,b.* from PurchaseReturnBody b join products p on p.productid = b.productid where SID=" & Val(TxtSID.Text) & " and Returndate='" & DtpReturnDate.DateValue & "' Order by SerialNo asc "
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("Serial").Text = Grid.Rows
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = !Code
            Grid.Columns("BatchNo").Text = IIf(IsNull(!BatchNo), "", !BatchNo)
            Grid.Columns("ProductName").Text = !ProductName
            
            If vColour = True Then
            If !ColourID = 0 Or IsNull(!ColourID) Then
               Grid.Columns("ColourID").Value = ""
            Else
               Grid.Columns("ColourID").Value = !ColourID
            End If
            If !ColourID = 0 Or IsNull(!ColourID) Then
               Grid.Columns("ColourName").Text = ""
            Else
               Grid.Columns("ColourName").Text = CN.Execute("Select ColourName from Colours where ColourID=" & !ColourID).Fields(0).Value
            End If
            If !SizeID = 0 Or IsNull(!SizeID) Then
               Grid.Columns("SizeID").Value = ""
            Else
               Grid.Columns("SizeID").Value = !SizeID
            End If
            If !SizeID = 0 Or IsNull(!SizeID) Then
               Grid.Columns("SizeName").Text = ""
            Else
               Grid.Columns("SizeName").Text = CN.Execute("Select SizeName from Sizes where SizeID=" & !SizeID).Fields(0).Value
            End If
            
            End If
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
            Grid.Columns("QtyLoose").Value = !QtyLoose
            Grid.Columns("Price").Value = !Price
            Grid.Columns("Bonus").Value = !Bonus
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("RetailPrice").Value = IIf(IsNull(!RetailPrice), "", !RetailPrice)
            Grid.Columns("SaleDiscPer").Value = IIf(IsNull(!SaleDiscPer), "", !SaleDiscPer)
            Grid.Columns("RetailAmount").Value = IIf(IsNull(!RetailAmount), "", !RetailAmount)
            Grid.Columns("ProfitAmount").Value = IIf(IsNull(!ProfitAmount), "", !ProfitAmount)
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(!Amount)
            TxtTotalItems.Text = Val(TxtTotalItems.Text) + !QtyLoose + IIf(IsNull(!Bonus), "0", !Bonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack))
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
   RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   RsBodySerial.Open "Select * from PurchaseReturnSerial where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
   Call PopulateDataToGridserial
   Call PopulateDataToGridOffer
   Call PopulateDataToGridExpense
End Sub

Private Sub PopulateDataToGridExpense()
    If RsExpense.State = adStateOpen Then RsExpense.Close
    RsExpense.Open "Select * from PurchaseReturnExpense where ReturnID =" & Val(TxtReturnID.Text) & " And ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
'    GridExpense.Visible = True
    ssql = "select EA.AccountNo, Accountname, PRE.ExpAmount from ExpenseAccounts EA Left Outer join ChartofAccounts C on C.AccountNo = EA.AccountNo Left Outer Join (Select * from PurchaseReturnExpense where ReturnID =" & Val(TxtReturnID.Text) & " And ReturnDate = '" & DtpReturnDate.DateValue & "') PRE On PRE.ExpenseID = EA.AccountNo"
      With CN.Execute(ssql)
         GridExpense.Redraw = False
         GridExpense.MoveFirst
         GridExpense.RemoveAll
         While Not .EOF
            GridExpense.AddNew
            GridExpense.Columns("ID").Text = !AccountNo
            GridExpense.Columns("Name").Text = IIf(IsNull(!AccountName), "", !AccountName)
            GridExpense.Columns("Value").Value = IIf(IsNull(!ExpAmount), 0, !ExpAmount)
'            TxtTotalExpense.Text = Val(TxtTotalExpense.Text) + Val(!expAmount)
            GridExpense.Update
            .MoveNext
         Wend
      End With

     If GridExpense.Rows > 0 Then GridExpense.FirstRow = 0
     GridExpense.Redraw = True
'      GridExpense.Visible = False
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
      FrmProductPrices.Visible = False
      vDate = IIf(vSystemDate = True, CN.Execute("Select SystemDate From SystemDate").Fields(0).Value, vServerDate)
      DtpReturnDate.DateValue = IIf(vSystemDate = True, IIf(IsNull(vDate), Date, vDate), IIf(Format(Now, "hh") >= vHDiff, vDate, DateAdd("d", -1, vDate)))
      TxtReturnID.Text = FunGetMaxID()
      Call PopulateDataToGrid
      Call PopulateDataPurchaseSerial
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      TxtCode.Enabled = True
      TxtStoreID.Enabled = True
      BtnStore.Enabled = True
      LblStock.Visible = False
      LblStockCaption.Visible = False
      LblCaptionRetailPrice.Visible = False
      LblRetailPrice.Visible = False
      BtnProduct.Enabled = True
      DtpReturnDate.Enabled = True
      If DtpReturnDate.Enabled And DtpReturnDate.Visible Then DtpReturnDate.SetFocus
      GridOffer.Visible = False
      FramExpense.ZOrder 0
      vIsNewRecord = True
   Case Is = OpenMode
      'TxtReturnID.Enabled = False
      DtpReturnDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      TxtStoreID.Enabled = False
      BtnStore.Enabled = False
      LblStock.Visible = False
      LblStockCaption.Visible = False
      LblCaptionRetailPrice.Visible = False
      LblRetailPrice.Visible = False
      TxtVenderID.SetFocus
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      TxtStoreID.Enabled = True
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
      TxtVenderID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
      TxtBiltyNo.SetFocus
   Else
      TxtVenderID.SetFocus
   End If
End Sub

Private Sub CmbPackName_Click()
   On Error GoTo ErrorHandler
   If CmbPackName.Text = "" Then
      TxtMultiplier.Text = ""
      TxtQtyPack.Text = ""
      TxtMultiplier.Enabled = False
      TxtQtyPack.Enabled = False
      TxtPrice.Text = Round(vUnitPrice, 3)
      TxtQtyLoose.Enabled = True
   Else
      If ObjRegistry.ChangeQtyPack = True Then TxtMultiplier.Enabled = True Else TxtMultiplier.Enabled = False
      TxtQtyPack.Enabled = True
      TxtQtyLoose.Enabled = Not ObjRegistry.EitherPackORLooseEnter
      If TxtQtyLoose.Enabled = False Then TxtQtyLoose.Text = ""
      If Trim(TxtCode.Text) <> "" Then
         With CN.Execute("select * from ProductPacking where ProductID = " & TxtProductID.Text & " and packingid=" & CmbPackName.ItemData(CmbPackName.ListIndex))
            TxtMultiplier.Text = IIf(.RecordCount = 0, "", !Multiplier)
            If Val(TxtMultiplier.Text) <> 0 Then
               TxtPrice.Text = Round(vUnitPrice * !Multiplier, 3)
            Else
               TxtPrice.Text = Round(vUnitPrice, 3)
            End If
            Call SubCalculateBody
         .Close
         End With
      End If
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub DtpReturnDate_Validate(Cancel As Boolean)
   TxtReturnID.Text = FunGetMaxID()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = "Grid" Then
         Grid_DblClick
      ElseIf ActiveControl.Name = "GridSerial" Then
         GridSerial_DblClick
      ElseIf ActiveControl.Name = TxtCode.Name Then
         If FunSelectProduct(ssValidate, False) = True Then If TxtBatchNo.Visible Then TxtBatchNo.SetFocus Else GetDataFromTexBoxesToGrid
      ElseIf ActiveControl.Name = TxtSerial.Name Then
         If Trim(TxtSerial.Text) = "" Or TxtCode.Enabled = False Then Exit Sub
         TxtCode.Text = Trim(TxtSerial.Text)
         If FunSelectProduct(ssValidate, False) = True Then
               GetDataFromTexBoxesToGrid
               TxtSerial.Text = ""
               TxtSerial.SetFocus
         End If
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
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then If TxtBatchNo.Visible Then TxtBatchNo.SetFocus Else GetDataFromTexBoxesToGrid
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, False) = True Then TxtBiltyNo.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtOrganizationID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then If TxtVenderID.Enabled Then TxtVenderID.SetFocus Else TxtStoreID.SetFocus
      End Select
    ElseIf KeyCode = vbKeyF2 Then
         If Frame3.Visible = True Then
            Frame3.Visible = False
            If TxtCode.Enabled = True Then TxtCode.SetFocus Else Grid.SetFocus
        Else
            Frame3.Visible = True
            Frame3.ZOrder 0
            KeyCode = 0
            If TxtSerial.Enabled = True And TxtSerial.Visible = True Then TxtSerial.SetFocus
        End If
   ElseIf KeyCode = vbKeyF3 Then
         If FramExpense.Visible = True Then
            FramExpense.Visible = False
'            If TxtExpenseID.Enabled = True Then TxtExpenseID.SetFocus Else Grid.SetFocus
        Else
            FramExpense.Visible = True
'            If TxtExpenseID.Enabled = True Then TxtExpenseID.SetFocus
        End If
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
         KeyCode = 0
         TxtBillDiscPer.SetFocus
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

Private Sub Grid_Change()
If GridOffer.Visible = False Then Exit Sub
    If Me.ActiveControl.Name <> GridOffer.Name Then Exit Sub
    If GridOffer.Columns("Qty").Value = Val("") Then GridOffer.Columns("qty").Value = 0
End Sub

Private Sub GridOffer_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
    DispPromptMsg = 0
End Sub

Private Sub GridOffer_BeforeUpdate(Cancel As Integer)
    If GridOffer.Visible = False Then Exit Sub
'    If Me.ActiveControl.Name <> GridOffer.Name Then Exit Sub
    
    If GridOffer.Columns("ProductName").Text = "" Then Exit Sub
    If RsProductOffer.RecordCount = 0 Then Exit Sub
    RsProductOffer.Filter = 0
    RsProductOffer.Filter = "ProductID = " & Val(GridOffer.Columns("ProductID").Text)
    If RsProductOffer.RecordCount = 0 Then Exit Sub
    RsProductOffer!QtyOffer = Val(GridOffer.Columns("Qty").Value)
    RsProductOffer.Update
    RsProductOffer.Filter = 0
    If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub GridOffer_Click()
'    If Me.ActiveControl.Name <> GridOffer.Name Then Exit Sub
'    If GridOffer.Columns("ProductName").Text = "" Then GridOffer.Columns("Qty").Locked = True Else GridOffer.Columns("Qty").Locked = False: Exit Sub
End Sub

Private Sub GridOffer_DblClick()
    If Me.ActiveControl.Name <> GridOffer.Name Then Exit Sub
'    If GridOffer.Top = 5700 Then
'        GridOffer.Top = 1750
'        GridOffer.Left = 8250
'    Else
'        GridOffer.Top = 5700
'        GridOffer.Left = 45
'    End If
End Sub

Private Sub GridOffer_GotFocus()
    GridOffer.Row = 0
   GridOffer.Col = 0
   SendKeys "{Right}"
End Sub

Private Sub GridOffer_LostFocus()
'    If Me.ActiveControl.Name <> GridOffer.Name Then Exit Sub
    If Trim(Grid.Columns("Code").Text) <> "" Then Exit Sub
    GridOffer.MoveLast
End Sub

Private Sub GridOffer_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(GridOffer.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub GetDataFromTextBoxesToGridOffer()
On Error GoTo ErrorHandler
    With CN.Execute("Select * from ProductOffers where Rebate = 0 and ProductID = " & Val(TxtProductID.Text))
            
        If .RecordCount > 0 Then
            QtyOffer = QtyOffer + Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)
            QtyOffer = QtyOffer \ !Qty * !QtyOffer
            If QtyOffer > 0 Then
                
                RsProductOffer.Filter = "ProductID = " & Val(TxtProductID.Text)
                If TxtProductID.Enabled Then
                    If RsProductOffer.RecordCount = 0 Then
                        RsProductOffer.AddNew
                        GridOffer.Columns("ProductID").Text = TxtProductID.Text
                        GridOffer.Columns("ProductOfferID").Text = !ProductOfferID
                        RsProductOffer!Productid = TxtProductID.Text
                        RsProductOffer!ProductOfferID = !ProductOfferID
                    Else
                        GridOffer.Redraw = False
                        GridOffer.MoveFirst
                        For vCounter = 1 To GridOffer.Rows
                        If GridOffer.Columns("ProductID").Text = TxtProductID.Text Then
                            GridOffer.Columns("ProductName").Text = CN.Execute("Select ProductName from products where productid = " & Val(GridOffer.Columns("ProductOfferID").Text)).Fields(0)
                            GridOffer.Columns("Qty").Value = QtyOffer
                            RsProductOffer!QtyOffer = QtyOffer
                            GridOffer.MoveLast
                            If TxtCode.Enabled = True Then TxtCode.SetFocus
                            GridOffer.Redraw = True
                            Exit Sub
                        End If
                            GridOffer.MoveNext
                        Next vCounter
                        GridOffer.MoveLast
                        Exit Sub
                    End If
                End If
                GridOffer.Redraw = False
                If RsProductOffer.RecordCount = 0 Then
                        RsProductOffer.AddNew
                        GridOffer.Columns("ProductID").Text = TxtProductID.Text
                        GridOffer.Columns("ProductOfferID").Text = !ProductOfferID
                        RsProductOffer!Productid = TxtProductID.Text
                        RsProductOffer!ProductOfferID = !ProductOfferID
                End If
                    GridOffer.Columns("ProductName").Text = CN.Execute("Select ProductName from products where productid = " & Val(GridOffer.Columns("ProductOfferID").Text)).Fields(0)
                    GridOffer.Columns("Qty").Value = QtyOffer
                    RsProductOffer!QtyOffer = QtyOffer
                    GridOffer.MoveLast
                If Trim(GridOffer.Columns("ProductID").Text) <> "" Then
                    GridOffer.AllowAddNew = True
                    GridOffer.AddNew
                    GridOffer.Columns("ProductID").Text = " "
                    GridOffer.AllowAddNew = False
                End If
            GridOffer.Redraw = True
            
            ''''''''''*******QtyOffer Less than total qty than Delete Row in GridOffer
        Else
            RsProductOffer.Filter = "ProductID = " & Val(TxtProductID.Text)
            If RsProductOffer.RecordCount > 0 Then
                GridOffer.MoveFirst
                For vCounter = 1 To GridOffer.Rows
                If GridOffer.Columns("ProductID").Text = TxtProductID.Text Then
                    If RsProductOffer.RecordCount > 0 Then RsProductOffer.Delete
                    GridOffer.SelBookmarks.RemoveAll
                    GridOffer.SelBookmarks.Add GridOffer.Bookmark
                    GridOffer.DeleteSelected
                    GridOffer.Refresh
                    GridOffer.MoveLast
                    Exit Sub
                End If
                    GridOffer.MoveNext
                    Next vCounter
                End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
        
    End If
    End With
   Exit Sub
ErrorHandler:
   GridOffer.Redraw = True
   Call ShowErrorMessage
End Sub


Private Sub PopulateDataToGridOffer()
    If RsProductOffer.State = adStateOpen Then RsProductOffer.Close
    RsProductOffer.Open "Select * from PurchaseReturnOffer where ReturnID =" & Val(TxtReturnID.Text) & " And ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
    If RsProductOffer.RecordCount > 0 Then
    GridOffer.Visible = True
    ssql = "select p.productname, D.* from PurchaseReturnOffer D Inner join products p on p.productid = D.productOfferid where ReturnID =" & Val(TxtReturnID.Text) & " And ReturnDate = '" & DtpReturnDate.DateValue & "'"
      With CN.Execute(ssql)
         GridOffer.Redraw = False
         GridOffer.MoveFirst
         GridOffer.RemoveAll
         GridOffer.AllowAddNew = True
         While Not .EOF
            GridOffer.AddNew
            GridOffer.Columns("ProductID").Text = !Productid
            GridOffer.Columns("ProductOfferID").Text = !ProductOfferID
            GridOffer.Columns("ProductName").Text = !ProductName
            GridOffer.Columns("Qty").Value = !QtyOffer
            .MoveNext
         Wend
      End With
      GridOffer.AddNew
      GridOffer.Columns("ProductID").Text = " "
      GridOffer.AllowAddNew = False
      GridOffer.Redraw = True
    Else
      GridOffer.CancelUpdate
      GridOffer.RemoveAll
      GridOffer.AddNew
      GridOffer.Columns("ProductID").Text = " "
      GridOffer.Update
      GridOffer.Visible = False
    End If
End Sub
Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   LblHelp.ForeColor = &H800000
   FraHelp.ZOrder 0
   FraHelp.Visible = True
End Sub


Private Sub LblClose_Click()
   FraHelp.Visible = False
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
   SetWindowText Me.hWnd, "Purchase Return Invoice"
   HelpLocation Me
   
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
   
   If ObjUserSecurity.ShowStock = True Or ObjUserSecurity.IsAdministrator Then
      vShowStock = True
   Else
      vShowStock = False
   End If
   
   '''''''''''''''' Form Default Setting  ''''''''''''''''''''''
   ssql = "select * from FormDefaultSetting Where FormType = 'Purchase Return Invoice' and LocalComputerName = '" & LocalComputerName & "'"
   With CN.Execute(ssql)
     If .RecordCount > 0 Then
        cmbPrintType.Text = !Size
        ChkIsPreview.Value = Abs(!IsPreview)
        ChkIsPrint.Value = Abs(!IsPrint)
        If Not IsNull(!DeviceName) Then
            CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
        Else
            CmbPrinters.ListIndex = 0
        End If
     End If
     .Close
   End With
   ''''''''''''''''''''''''''''''''''''''''''''''
   
   vNoofPrints = IIf(IsNull(ObjRegistry.NoofPrints) Or Val(ObjRegistry.NoofPrints) = 0, 1, ObjRegistry.NoofPrints)
   vServerDate = CN.Execute("Select CONVERT(datetime, CONVERT(varchar, GETDATE(), 110)) ServerDate").Fields(0).Value
   vSystemDate = Abs(ObjRegistry.SystemDate)
   vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
   
   vNegativeSale = ObjRegistry.NegativeSale
   vColour = ObjRegistry.ShowColourSize
   
   LblColour.Visible = vColour
   CmbColourName.Visible = vColour
   LblSize.Visible = vColour
   cmbSizeName.Visible = vColour
   Grid.Columns("ColourName").Visible = vColour
   Grid.Columns("SizeName").Visible = vColour
   
   LblRetailAmount.Visible = ObjRegistry.ShowPurchaseProfit
   LblProfitAmount.Visible = ObjRegistry.ShowPurchaseProfit
   LblSaleDiscPer.Visible = ObjRegistry.ShowPurchaseProfit
   TxtRetailAmount.Visible = ObjRegistry.ShowPurchaseProfit
   TxtProfitAmount.Visible = ObjRegistry.ShowPurchaseProfit
   TxtSaleDiscPer.Visible = ObjRegistry.ShowPurchaseProfit
   
   If vColour = False Then
      LblPackName.Left = LblPackName.Left - CmbColourName.Width - cmbSizeName.Width
      CmbPackName.Left = CmbPackName.Left - CmbColourName.Width - cmbSizeName.Width
      LblMultiplier.Left = LblMultiplier.Left - CmbColourName.Width - cmbSizeName.Width
      TxtMultiplier.Left = TxtMultiplier.Left - CmbColourName.Width - cmbSizeName.Width
      LblQtyPack.Left = LblQtyPack.Left - CmbColourName.Width - cmbSizeName.Width
      TxtQtyPack.Left = TxtQtyPack.Left - CmbColourName.Width - cmbSizeName.Width
      LblQtyLoose.Left = LblQtyLoose.Left - CmbColourName.Width - cmbSizeName.Width
      TxtQtyLoose.Left = TxtQtyLoose.Left - CmbColourName.Width - cmbSizeName.Width
      LblBonus.Left = LblBonus.Left - CmbColourName.Width - cmbSizeName.Width
      TxtBonus.Left = TxtBonus.Left - CmbColourName.Width - cmbSizeName.Width
      LblOffer.Left = LblOffer.Left - CmbColourName.Width - cmbSizeName.Width
      TxtOffer.Left = TxtOffer.Left - CmbColourName.Width - cmbSizeName.Width
      LblPrice.Left = LblPrice.Left - CmbColourName.Width - cmbSizeName.Width
      TxtPrice.Left = TxtPrice.Left - CmbColourName.Width - cmbSizeName.Width
      LblDiscPC.Left = LblDiscPC.Left - CmbColourName.Width - cmbSizeName.Width
      TxtDiscPC.Left = TxtDiscPC.Left - CmbColourName.Width - cmbSizeName.Width
      LblSaleTaxPer.Left = LblSaleTaxPer.Left - CmbColourName.Width - cmbSizeName.Width
      TxtSaleTaxPer.Left = TxtSaleTaxPer.Left - CmbColourName.Width - cmbSizeName.Width
      LblDiscPer.Left = LblDiscPer.Left - CmbColourName.Width - cmbSizeName.Width
      TxtDiscPer.Left = TxtDiscPer.Left - CmbColourName.Width - cmbSizeName.Width
      LblDiscVal.Left = LblDiscVal.Left - CmbColourName.Width - cmbSizeName.Width
      TxtDiscVal.Left = TxtDiscVal.Left - CmbColourName.Width - cmbSizeName.Width
      LblAmount.Left = LblAmount.Left - CmbColourName.Width - cmbSizeName.Width
      TxtAmount.Left = TxtAmount.Left - CmbColourName.Width - cmbSizeName.Width
      Grid.Width = Grid.Width - CmbColourName.Width - cmbSizeName.Width
   End If

   With CN.Execute("Select * from Packings")
      CmbPackName.AddItem ""
      While Not .EOF
         CmbPackName.AddItem !PackingName
         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
         .MoveNext
      Wend
      .Close
   End With
   
   TxtStoreID.Text = IIf((ObjRegistry.StoreID = ""), "", ObjRegistry.StoreID)
   FunSelectStore ssValidate, True
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   TxtBatchNo.Visible = ObjRegistry.BatchNoVisible

   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   If ObjUserSecurity.IsAdministrator = False Then
      LblPrice.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtPrice.Visible = Not ObjRegistry.HidePurchaseAmount
      LblAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      LblTotalAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtTotalAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      LblNetAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtNetAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      lblPayable.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtPreviousPayable.Visible = Not ObjRegistry.HidePurchaseAmount
      LblTtlPayable.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtTotalPayable.Visible = Not ObjRegistry.HidePurchaseAmount
      Grid.Columns("Price").Visible = Not ObjRegistry.HidePurchaseAmount
      Grid.Columns("Amount").Visible = Not ObjRegistry.HidePurchaseAmount
   End If
   With CN.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
      End If
      .Close
   End With
   vAutoEnterBeforeQty = ObjRegistry.AutoEnterBeforeQty
   
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpReturnDate.IsDateValid = False Then Exit Function
   If ObjRegistry.AllowContinuousBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(ReturnID),0)+1 from PurchaseReturnHeader").Fields(0)
   ElseIf ObjRegistry.AllowMonthlyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(ReturnID),0)+1 from PurchaseReturnHeader where Month(Returndate) = '" & Month(DtpReturnDate.DateValue) & "' and  year(Returndate) ='" & Year(DtpReturnDate.DateValue) & "'").Fields(0)
   ElseIf ObjRegistry.AllowDailyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(ReturnID),0)+1 from PurchaseReturnHeader where Returndate = '" & DtpReturnDate.DateValue & "'").Fields(0)
   Else
      FunGetMaxID = CN.Execute("Select isnull(max(ReturnID),0)+1 from PurchaseReturnHeader where Returndate = '" & DtpReturnDate.DateValue & "' and StoreID = " & TxtStoreID.Text).Fields(0)
   End If
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
   CmbColourName.Clear
   cmbSizeName.Clear
   TxtNetAmount.Text = 0
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("Code").Text = " "
   Grid.Update
   
   GridOffer.Visible = False
   Call SubClearSerialFields
   Frame3.Visible = False
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
    Set RsBodySerial = Nothing
    Set RsProductOffer = Nothing
    Set FrmPurchaseReturnInvoice = Nothing
   End If
   
   ''''''''''''''''' ActivityLogBin For Close Action
'      Call DeleteTempActivityLogBin(vRandomID)
      If Grid.Rows > 1 And Cancel = 0 Then
         vGridRows = 0
         Grid.Redraw = False
         Grid.MoveFirst
         For vCounter = 2 To Grid.Rows
            vGridRows = vGridRows + 1
            If Trim(Grid.Columns("Code").Text) <> "" Then
               ssql = "Select Productid From PurchaseReturnbody where ReturnID=" & Val(TxtReturnID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
               With CN.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Closed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eCloseSavedRecord, TxtReturnID.Text, DtpReturnDate.DateValue, vGridRows & " Product/s Closed")
         Grid.Redraw = True
      End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("Amount").Value
   TxtTotalItems.Text = Val(TxtTotalItems.Text) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Grid.Columns("QtyPack").Value)))
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
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(Grid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
   Call PopulateDataToGridserial
    GridOffer.MoveFirst
    For vCounter = 1 To GridOffer.Rows
    If GridOffer.Columns("ProductID").Text = Grid.Columns("ProductID").Text Then Exit Sub
        GridOffer.MoveNext
    Next vCounter
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name = "Grid" Then
   If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
   
   ssql = "Select Productid From PurchaseReturnbody where ReturnID=" & Val(TxtReturnID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
   With CN.Execute(ssql)
      If .EOF Then
         Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      Else
         Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eRemoveRow, TxtReturnID.Text, DtpReturnDate.DateValue, "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
         Call ActivityLogBin(vRandomID, eFrmPurchaseReturnInvoice, eAddTempRecord, TxtReturnID.Text, DtpReturnDate.DateValue, "Pending Remove Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      End If
   End With
   
   RsBody.Filter = "ProductID = " & Val(TxtProductID.Text) & " and BatchNo = " & IIf(Trim(TxtBatchNo.Text) = "", "null", "'" & Trim(TxtBatchNo.Text) & "'") & " and Price = " & Val(TxtPrice.Text)
   If RsBody.RecordCount > 0 Then RsBody.Delete
'    RsProductOffer.Filter = "ProductID='" & GridOffer.Columns("code").Text & "'"
'    If RsProductOffer.RecordCount > 0 Then
'        RsProductOffer.Delete
'        GridOffer.SelBookmarks.RemoveAll
'        GridOffer.SelBookmarks.Add GridOffer.Bookmark
'        GridOffer.DeleteSelected
'        GridOffer.Refresh
'        RsProductOffer.Filter = 0
'    End If
    
   RsBodySerial.Filter = ""
   RsBodySerial.Filter = "ProductID = " & Val(TxtCode.Text)
   
   While Not RsBodySerial.EOF
      
      RsPurchaseSerial.Filter = "Serial = " & RsBodySerial!serial
      If RsPurchaseSerial.RecordCount > 0 Then
         RsPurchaseSerial!SerialAdd = 1
         RsPurchaseSerial.Update
       End If
            
      RsBodySerial.Delete
      RsBodySerial.MoveNext
   Wend
   
   
    SubClearSerialFields
    CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
    Grid.SelBookmarks.RemoveAll
    Grid.SelBookmarks.Add Grid.Bookmark
    Grid.DeleteSelected
    Grid.Refresh
    RsBody.Filter = 0
    Grid.MoveLast
    GetDataBackFromGridToTexBoxes
   ElseIf Me.ActiveControl.Name = "GridSerial" Then
    If Trim(GridSerial.Columns("Serial").Text) = "" Then Exit Sub
    RsBodySerial.Filter = "Serial = '" & TxtSerial.Text & "'"
    If RsBodySerial.RecordCount > 0 Then RsBodySerial.Delete
    CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed ProdcutID-" & GridSerial.Columns("ProductID").Text & " Serial-" & GridSerial.Columns("Serial").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
    GridSerial.SelBookmarks.RemoveAll
    GridSerial.SelBookmarks.Add GridSerial.Bookmark
    GridSerial.DeleteSelected
    GridSerial.Refresh
    RsBodySerial.Filter = 0
'    GridSerial.MoveLast
    GetDataBackFromGridSerialToTexBoxes
   ElseIf Me.ActiveControl.Name = GridOffer.Name Then
    RsProductOffer.Filter = "ProductID = " & Val(GridOffer.Columns("ProductID").Text)
    If RsProductOffer.RecordCount > 0 Then
        RsProductOffer.Delete
        GridOffer.SelBookmarks.RemoveAll
        GridOffer.SelBookmarks.Add GridOffer.Bookmark
        GridOffer.DeleteSelected
        GridOffer.Refresh
        RsProductOffer.Filter = 0
    End If
   End If
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
   If Val(TxtPrice.Text) <> 0 Then
      If Round(Val(TxtDiscPer.Text), 2) <> Round((Val(TxtDiscPC.Text) * 100) / (Val(TxtPrice.Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))), 2) Then
         MsgBox "Please update the Discount for change Price.", vbExclamation, "Alert"
         If TxtDiscPer.Enabled And TxtDiscPer.Visible Then TxtDiscPer.SetFocus
         Exit Sub
      End If
   End If
   If (CmbColourName.Text = "" Or cmbSizeName.Text = "") And vColour = True Then
      MsgBox "Please Select Colour and Size", vbInformation + vbOKOnly, "Error"
      Exit Sub
   End If
   '''''''''   check Serial
   RsBodySerial.Filter = "ProductID =" & Val(TxtCode.Text)
   If (TxtCode.Enabled = False And RsBodySerial.RecordCount <> 0) And RsBodySerial.RecordCount <> Val(Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) Then
      MsgBox "Qty Should be equal to Serial", vbInformation + vbOKOnly, "Error"
      Call SubClearDetailArea
      If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
      Exit Sub
   End If
   RsBodySerial.Filter = ""
''''''''
   FrmProductPrices.Visible = False
'    If vNegativeSale = False Then
'      If vIsNewRecord = True Then
'         If (Val(vQtyLoose) - ((Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text)) + (TxtQtyLoose.Text))) < 0 Then
'            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'            Grid.Redraw = True
'            Call SubClearDetailArea
'            If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
'            Exit Sub
'         End If
'      Else
'         If (Val(vQtyLoose) - ((Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text)) + (TxtQtyLoose.Text)) + Val(Grid.Columns("QtyOrigional").Value)) < 0 Then
'            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'            Grid.Redraw = True
'            Call SubClearDetailArea
'            If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
'            Exit Sub
'         End If
'      End If
'   End If
   
   
   If Val(Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) <> Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value) + Val(Grid.Columns("QtyLoose").Value) Then
      If Val(Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) > Val(vQtyLoose) And ObjRegistry.NegativeSale = False Then
         MsgBox "Return Stock can not greater than available stock", vbInformation + vbOKOnly, "Error"
         TxtQtyLoose.SetFocus
         Exit Sub
      End If
   End If
   FrmHistory.Visible = False
   
   RsBody.Filter = ""
   
   If Trim(Grid.Columns("Productid").Text) = "" Then
      RsBody.Filter = "ProductID = " & Val(TxtProductID.Text) & " and BatchNo = " & IIf(Trim(TxtBatchNo.Text) = "", "null", "'" & Trim(TxtBatchNo.Text) & "'") & " and Price = " & Val(TxtPrice.Text)
   Else
      RsBody.Filter = "ProductID = " & Val(Grid.Columns("Productid").Text) & " and BatchNo = " & IIf(Grid.Columns("BatchNo").Text = "", "null", "'" & Grid.Columns("BatchNo").Text & "'") & " and Price = " & Val(Grid.Columns("Price").Text)
   End If
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("Serial").Text = Grid.Rows
         Grid.Columns("ProductID").Text = TxtProductID.Text
         Grid.Columns("Code").Text = TxtCode.Text
         Grid.Columns("Price").Value = Val(TxtPrice.Text)
         Grid.Columns("BatchNo").Text = Trim(TxtBatchNo.Text)
         RsBody!Productid = TxtProductID.Text
         RsBody!Code = TxtCode.Text
         RsBody!Price = Val(TxtPrice.Text)
         RsBody!BatchNo = Trim(TxtBatchNo.Text)
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("Productid").Text = TxtProductID.Text And Grid.Columns("BatchNo").Text = Trim(TxtBatchNo.Text) And Val(Grid.Columns("Price").Text) = Round(Val(TxtPrice.Text), 2) Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  
                  ssql = "Select Productid From PurchaseReturnbody where ReturnID=" & Val(TxtReturnID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Grid.Columns("ProductID").Text
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  '''''''''''''''''''''''''This QtyOffer Is used for DetailGrid
                  QtyOffer = Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value) + Val(Grid.Columns("QtyLoose").Value)
                  GetDataFromTextBoxesToGridOffer
                  TxtOffer.Text = Val(TxtOffer.Text) + Val(Grid.Columns("Offer").Text)
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  
                  TxtQtyLoose.Text = Val(TxtQtyLoose.Text) + Grid.Columns("QtyLoose").Value
                  TxtQtyPack.Text = Val(TxtQtyPack.Text) + Grid.Columns("QtyPack").Value
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Val(Grid.Columns("QtyLoose").Value) + Val(Grid.Columns("Bonus").Value) + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
                  Grid.Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
                  Grid.Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, 0, Val(TxtMultiplier.Text))
                  Grid.Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
                  Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("Bonus").Value = Val(TxtBonus.Text)
                  Grid.Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  Grid.Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  Grid.Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  Grid.Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  Grid.Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  Grid.Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("RetailPrice").Value = Val(TxtRetailPrice.Text)
                  Grid.Columns("SaleDiscPer").Value = IIf(Val(TxtSaleDiscPer.Text) = 0, 0, Val(TxtSaleDiscPer.Text))
                  Grid.Columns("RetailAmount").Value = Val(TxtRetailAmount.Text)
                  Grid.Columns("ProfitAmount").Value = Val(TxtProfitAmount.Text)
                  
                  RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
                  RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
                  RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
                  RsBody!QtyLoose = Val(TxtQtyLoose.Text)
                  RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  RsBody!Bonus = Val(TxtBonus.Text)
                  RsBody!Price = Val(TxtPrice.Text)
                  RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  RsBody!Amount = Val(TxtAmount.Text)
                  RsBody!RetailPrice = Val(TxtRetailPrice.Text)
                  RsBody!SaleDiscPer = IIf(Val(TxtSaleDiscPer.Text) = 0, 0, Val(TxtSaleDiscPer.Text))
                  RsBody!RetailAmount = Val(TxtRetailAmount.Text)
                  RsBody!ProfitAmount = Val(TxtProfitAmount.Text)
                  ssql = "Select Productid From PurchaseReturnbody where ReturnID=" & Val(TxtReturnID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmPurchaseReturnInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Pending Update Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
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
'      If Trim(Grid.Columns("Productid").Text) = "" Then
         If TxtCode.Enabled = True Then
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text)))
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eAddNewRowByEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
         Call ActivityLogBin(vRandomID, eFrmPurchaseReturnInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Pending Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Val(Grid.Columns("Pack").Value)) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
         ssql = "Select Productid From PurchaseReturnbody where ReturnID=" & Val(TxtReturnID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
         With CN.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            Else
               Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eEdit, TxtReturnID.Text, DtpReturnDate.Date, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmPurchaseReturnInvoice, eEdit, TxtReturnID.Text, DtpReturnDate.Date, "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmPurchaseReturnInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Pending Update Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
      End If
      
      .Columns("ProductID").Text = TxtProductID.Text
      .Columns("Code").Text = TxtCode.Text
      .Columns("BatchNo").Text = Trim(TxtBatchNo.Text)
      .Columns("ProductName").Text = TxtProductName.Text
      
      .Columns("ColourName").Text = CmbColourName.Text
      If CmbColourName.Text <> "" Then .Columns("ColourID").Value = CmbColourName.ItemData(CmbColourName.ListIndex)
      .Columns("SizeName").Text = cmbSizeName.Text
      If cmbSizeName.Text <> "" Then .Columns("SizeID").Value = cmbSizeName.ItemData(cmbSizeName.ListIndex)
      
      If vColour = True And .Columns("ColourID").Text <> "" Then
         RsBody!ColourID = .Columns("ColourID").Text
         RsBody!SizeID = .Columns("SizeID").Text
      End If
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
      .Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, 0, Val(TxtMultiplier.Text))
      .Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
      .Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
      .Columns("Bonus").Value = Val(TxtBonus.Text)
      .Columns("Price").Value = Val(TxtPrice.Text)
      .Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      .Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
      .Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
      .Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      .Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      .Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("RetailPrice").Value = Val(TxtRetailPrice.Text)
      .Columns("SaleDiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      .Columns("RetailAmount").Value = Val(TxtRetailAmount.Text)
      .Columns("ProfitAmount").Value = Val(TxtProfitAmount.Text)
      RsBody!BatchNo = IIf(Trim(TxtBatchNo.Text) = "", Null, Trim(TxtBatchNo.Text))
      RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
      RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
      RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
      RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      RsBody!QtyLoose = Val(TxtQtyLoose.Text)
      RsBody!Bonus = Val(TxtBonus.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      RsBody!Amount = Val(TxtAmount.Text)
      RsBody!RetailPrice = Val(TxtRetailPrice.Text)
      RsBody!SaleDiscPer = IIf(Val(TxtSaleDiscPer.Text) = 0, 0, Val(TxtSaleDiscPer.Text))
      RsBody!RetailAmount = Val(TxtRetailAmount.Text)
      RsBody!ProfitAmount = Val(TxtProfitAmount.Text)
      .MoveLast
      If TxtCode.Enabled = True Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   QtyOffer = 0
   GetDataFromTextBoxesToGridOffer
   Call SubClearDetailArea
   GetDataBackFromGridToTexBoxes
   Grid_LostFocus
   
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   CmbColourName.Clear
   cmbSizeName.Clear
   TxtCode.Enabled = True
   BtnProduct.Enabled = True
   TxtCode.Text = ""
   TxtProductName.Text = ""
   CmbPackName.ListIndex = 0
   TxtBatchNo.Text = ""
   TxtMultiplier.Text = ""
   TxtQtyPack.Text = ""
   TxtQtyLoose.Text = ""
   TxtBonus.Text = ""
   TxtPrice.Text = ""
   TxtOffer.Text = ""
   TxtSaleTaxPer.Text = ""
   TxtSaleTaxVal.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtDiscVal.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("Code").Text
      TxtBatchNo.Text = .Columns("BatchNo").Text
      TxtProductName.Text = .Columns("ProductName").Text
      If Trim(.Columns("PackName").Text) = "" Then
         CmbPackName.ListIndex = 0
      Else
         CmbPackName.Text = .Columns("PackName").Text
      End If
      TxtMultiplier.Text = .Columns("Pack").Text
      TxtQtyLoose.Text = .Columns("QtyLoose").Text
      TxtQtyPack.Text = .Columns("QtyPack").Text
      TxtBonus.Text = .Columns("Bonus").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtOffer.Text = .Columns("Offer").Value
      TxtSaleTaxPer.Text = .Columns("SaleTaxPer").Value
      TxtSaleTaxVal.Text = .Columns("SaleTaxVal").Value
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      TxtAmount.Text = .Columns("Amount").Value
      TxtRetailPrice.Text = .Columns("RetailPrice").Text
      TxtSaleDiscPer.Text = .Columns("SaleDiscPer").Value
      TxtRetailAmount.Text = Abs(Val(.Columns("RetailAmount").Value))
      TxtProfitAmount.Text = Abs(Val(.Columns("ProfitAmount").Value))
      
      If ObjRegistry.ShowAllPrices Then
         PopulateDataToPriceGrid
         FrmProductPrices.Visible = True
      End If
      
      If LblStock.Visible = False Then
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
      End If
       vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpPurchaseDate.DateValue + 1 & "',0),0)"
          With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               vQtyLoose = .Fields(0).Value
            Else
               vQtyLoose = 0
            End If
         End With
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
'      With CN.Execute("select QtyLoose from currentstockStore where productid ='" & TxtProductID.Text & "' and storeid = " & TxtStoreID.Text)
'         If .RecordCount > 0 Then
'            'vQtyLoose = !QtyLoose
'            LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack('" & TxtProductID.Text & "',Floor(" & !QtyLoose & "))").Fields(0).Value
'            LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'            LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & !QtyLoose & "))").Fields(0).Value
'            LblStock.Caption = LblStock.Caption & " " & "Loose"
'         Else
'            'vQtyLoose = 0
'            LblStock.Caption = 0
'         End If
'      End With
      vUnitPrice = Val(.Columns("Price").Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))
      'vUnitRetailPrice = Val(.Columns("RetailPrice").Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))
      If Trim(TxtProductID.Text) <> "" Then
         LblRetailPrice.Caption = CN.Execute("Select RetailPrice from Products where ProductID = " & Val(TxtProductID.Text)).Fields(0).Value
      End If

   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetReturn()
   On Error GoTo ErrorHandler
   ssql = "Select h.*, OrganizationName, p.partyname, Address, City, StoreName FROM PurchaseReturnHeader h join parties p on h.Vendorid = p.partyid inner join stores s on s.storeid = h.storeid left outer join Organizations o on o.OrganizationID = h.OrganizationID where h.SID=" & Val(TxtSID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'" & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With CN.Execute(ssql)
      If Not .BOF Then
          DtpReturnDate.DateValue = !ReturnDate
          TxtPurchaseID.Text = IIf(IsNull(!PurID), "", !PurID)
          DtpPurchaseDate.DateValue = IIf(IsNull(!PurchaseDate), "01/01/1990", !PurchaseDate)
          TxtVenderID.Text = !vendorID
          TxtVenderName.Text = !PartyName
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
          TxtBillNo.Text = IIf(IsNull(!BillNo), "", !BillNo)
          TxtBiltyNo.Text = IIf(IsNull(!BiltyNo), "", !BiltyNo)
          TxtVehicleNo.Text = IIf(IsNull(!VehicleNo), "", !VehicleNo)
          TxtTotalAmount.Text = !TotalAmount
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtTotalExpense.Text = IIf(IsNull(!TotalExpense), "", !TotalExpense)
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          TxtReceivedAmount.Text = IIf(IsNull(!ReceivedAmount), "", !ReceivedAmount)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtReceivedAmount.Text = IIf(IsNull(!ReceivedAmount), "", !ReceivedAmount)
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

Private Sub GridSerial_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub GridSerial_DblClick()
   If Grid.Columns("Code").Text <> " " And GridSerial.Columns("Serial").Text = " " Then
        TxtSerial.Enabled = True
        TxtSerial.SetFocus
    Else
'        TxtSerial.Enabled = False
    End If
End Sub

Private Sub GridSerial_GotFocus()
   If Grid.Columns("Code").Text <> " " And GridSerial.Columns("Serial").Text = " " Then
        TxtSerial.Enabled = True
    Else
'        TxtSerial.Enabled = False
    End If
End Sub

Private Sub GridSerial_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(GridSerial.Columns("Serial").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub GridSerial_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    GetDataBackFromGridSerialToTexBoxes
End Sub

Private Sub GridSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub TxtBillDisc_Change()
   If ActiveControl.Name <> TxtBillDisc.Name Then Exit Sub
   TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / Val(TxtTotalAmount.Text), 6)
   Call SubCalculateFooter
End Sub

Private Sub TxtBillDiscPer_Change()
   If ActiveControl.Name <> TxtBillDiscPer.Name Then Exit Sub
   TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Text) * Val(TxtBillDiscPer.Text) / 100))
   Call SubCalculateFooter
End Sub

Private Sub TxtDiscPC_Change()
   If ActiveControl.Name <> TxtDiscPC.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
   If Val(TxtDiscPer.Text) = 0 Then TxtDiscPer.Text = ""
   Call SubCalculateBody
End Sub

Private Sub TxtDiscPer_Change()
   If ActiveControl.Name <> TxtDiscPer.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   TxtDiscPC.Text = Round((vUnitPrice * Val(TxtDiscPer.Text) / 100), 4)
   If Val(TxtDiscPC.Text) = 0 Then TxtDiscPC.Text = ""
   Call SubCalculateBody
End Sub

Private Sub TxtDiscPer_LostFocus()
   Select Case ActiveControl.Name
   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtOffer.Name, TxtSaleTaxPer.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
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

Private Sub TxtOffer_Change()
   Call SubCalculateBody
End Sub

Private Sub TxtOtherCharges_Change()
   Call SubCalculateFooter
End Sub

Private Sub TxtPrice_Change()
   If ActiveControl.Name <> TxtPrice.Name Then Exit Sub
   If Val(TxtMultiplier.Text) <> 0 Then
      vUnitPrice = Val(TxtPrice.Text) / Val(TxtMultiplier.Text)
   Else
      vUnitPrice = Val(TxtPrice.Text)
   End If
   Call SubCalculateBody
End Sub

Private Sub TxtQtyLoose_Change()
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub TxtQtyPack_Change()
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub FindRebate()
   On Error GoTo ErrorHandler
    With CN.Execute("Select * from ProductOffers where Rebate <> 0 and ProductID = " & Val(TxtProductID.Text))
        If .RecordCount > 0 Then
            Rebate = Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)
            Rebate = Rebate \ !Qty
            Rebate = Rebate * !Rebate
            TxtOffer.Text = Rebate
        End If
    End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSaleTaxPer_Change()
   If ActiveControl.Name <> TxtSaleTaxPer.Name Then Exit Sub
   Call SubCalculateBody
End Sub

Private Sub TxtSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridSerial.SetFocus
End Sub


Private Sub TxtTotalAmount_Change()
   TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Text) * Val(TxtBillDiscPer.Text) / 100))
   Call SubCalculateFooter
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
      TxtPrice.Text = ""
   End If
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

'Private Sub TxtCode_LostFocus()
'   If Len(TxtCode.Text) > 7 Then
'      GetDataFromTexBoxesToGrid
'   End If
'End Sub

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

Private Sub TxtTotalExpense_Change()
  Call SubCalculateFooter
End Sub

Private Sub TxtVenderID_Change()
   If TxtVenderID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   If TxtVenderName.Text <> "" Then
      TxtVenderName.Text = ""
   End If
End Sub

Private Sub TxtVenderID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtVenderName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVender(ssValidate, True)
'   If vTemp = True Then
'      vTemp = Not FunSelectVender(ssButton, False)
'   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

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

Private Sub GetDataFromTexBoxesToGridSerial()
   On Error GoTo ErrorHandler
   
     
   vStrSQL = "Select ProductID, Serial, SerialAdd from vuPurchaseSerial where Serial = '" & Trim(TxtSerial.Text) & "'"
      
      With CN.Execute(vStrSQL)
         If .EOF = True Then
            MsgBox "The Serail cannot be inserted because it is not Exist", vbInformation + vbOKOnly, "Error"
            TxtSerial.Text = ""
            Exit Sub
         ElseIf !SerialAdd = False Then
            MsgBox "There is not stock of this Serail", vbInformation + vbOKOnly, "Error"
            TxtSerial.Text = ""
            Exit Sub
         End If
      End With

   GridSerial.MoveLast
   
   RsBodySerial.Filter = ""
   RsBodySerial.Filter = "ProductID =" & TxtCode.Text & " And Serial='" & TxtSerial.Text & "'"
   If RsBodySerial.RecordCount > 0 Then
      MsgBox "The Serail cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
      TxtSerial.Text = ""
      Exit Sub
   End If
   
   If TxtSerial.Enabled Then
         
         GridSerial.MoveLast
         GridSerial.Columns("ProductID").Text = TxtCode.Text
         GridSerial.Columns("Serial").Text = TxtSerial.Text
         
         RsBodySerial.AddNew
         RsBodySerial!Productid = TxtCode.Text
         RsBodySerial!serial = TxtSerial.Text
         RsBodySerial.Update
         vSerialAdd = True
         TxtSerial.Text = ""
  End If
  
   With GridSerial
      If Trim(.Columns("Serial").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Serial").Text = " "
         .AllowAddNew = False
      End If
   End With
   
   Exit Sub
ErrorHandler:
   GridSerial.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGridSerial2()
   On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   If Trim(TxtSerial.Text) = "" Then
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
'      TxtSerial.SetFocus
      Exit Sub
   End If
   RsBodySerial.Filter = "ProductID = " & Val(Grid.Columns("ProductID").Text) & " And Serial='" & TxtSerial.Text & "'"
   If TxtSerial.Enabled Then
      If RsBodySerial.RecordCount = 0 Then
         RsBodySerial.AddNew
         GridSerial.Columns("ProductID").Text = TxtCode.Text
         GridSerial.Columns("Serial").Text = TxtSerial.Text
         RsBodySerial!Productid = TxtCode.Text
         RsBodySerial!serial = TxtSerial.Text
         TxtSerial.Text = ""
      Else
         GridSerial.Redraw = False
         GridSerial.MoveFirst
            For vrowcounter = 1 To GridSerial.Rows
               If GridSerial.Columns("Serial").Text = TxtSerial.Text Then
                  MsgBox "The Product cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  GridSerial.MoveLast
                  TxtSerial.SetFocus
                  GridSerial.Redraw = True
                  Exit Sub
               End If
               GridSerial.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         
         GridSerial.MoveLast
         TxtSerial.SetFocus
         Exit Sub
      End If
   End If
   'GridSerial.Redraw = False
   With GridSerial
      If Trim(.Columns("Serial").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Serial").Text = " "
         .AllowAddNew = False
      End If
   End With
   TxtSerial.SetFocus
   GridSerial.Redraw = True
   Exit Sub
ErrorHandler:
   GridSerial.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub GetDataBackFromGridSerialToTexBoxes()
   On Error GoTo ErrorHandler
   With GridSerial
      TxtSerial.Text = .Columns("Serial").Text
   End With
   If GridSerial.Rows = 1 Then GridSerial.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearSerialFields()
   TxtSerial.Text = ""
'   TxtSerial.Enabled = False
   GridSerial.CancelUpdate
   GridSerial.RemoveAll
   GridSerial.AddNew
   GridSerial.Columns("Serial").Text = " "
   GridSerial.Update
End Sub

Private Sub PopulateDataToGridserial()
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      RsBodySerial.Filter = 0
   Else
      RsBodySerial.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
   End If
   GridSerial.Redraw = False
   GridSerial.MoveFirst
   GridSerial.RemoveAll
   GridSerial.AllowAddNew = True
   If RsBodySerial.RecordCount > 0 Then
      With RsBodySerial
         .MoveFirst
         While Not .EOF
            GridSerial.AddNew
            GridSerial.Columns("ProductID").Text = !Productid
            GridSerial.Columns("Serial").Text = !serial
            .MoveNext
         Wend
'      .Close
      GridSerial.MoveLast
      End With
   End If
   GridSerial.AddNew
   GridSerial.Columns("ProductID").Text = " "
   GridSerial.AllowAddNew = False
   GridSerial.Redraw = True
   RsBodySerial.Filter = 0
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpReturnDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = CN.Execute("Select isnull(max(BinID),0)+1 from Bin_PurchaseReturnHeader ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
     If vIsNewRecord = False Then
    With CN.Execute("Select  * from PurchaseReturnHeader where ReturnID =" & TxtReturnID.Text & " And ReturnDate = '" & DtpReturnDate.DateValue & "'")
        If Val(TxtVenderID.Text) <> IIf(IsNull(!vendorID), 0, !vendorID) Then
            CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated VenderID-" & !vendorID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtStoreID.Text <> !StoreID Then
            CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.Rows - 1
        With CN.Execute("Select * from PurchaseReturnBody Where ReturnID = " & TxtReturnID.Text & " and ReturnDate ='" & DtpReturnDate.DateValue & "' and Productid = " & Val(Grid.Columns("Productid").Text))
        
             If .EOF = True Then
                ssql = "Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")"
                CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("QtyLoose").Text <> !QtyLoose Or Grid.Columns("Price").Text <> !Price Or Grid.Columns("discper").Text <> !DiscPer Then
                   CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
    
   Else
    CN.Execute ("Insert Into UserActivities values ('Purchase Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub PopulateDataToPriceGrid()
    
      ssql = "select desc1, Listprice, WsPrice, RetailPrice, PurPrice-isnull(purDiscPC*isnull(multiplier,1),0) as PurPrice " & vbCrLf & _
            " from products p" & vbCrLf & _
            " left outer join productpacking pp on p.productid = pp.productid and p.PurchasePackingID = pp.packingid" & vbCrLf & _
            " where p.productID = " & Val(TxtProductID.Text)
      
      
      With CN.Execute(ssql)
         GridProductPrices.Redraw = False
         GridProductPrices.MoveFirst
         GridProductPrices.RemoveAll
         GridProductPrices.AllowAddNew = True
         While Not .EOF
            GridProductPrices.AddNew
            GridProductPrices.Columns("Description").Text = IIf(IsNull(!Desc1), "", !Desc1)
            GridProductPrices.Columns("Pur").Value = !PurPrice
            GridProductPrices.Columns("List").Value = IIf(IsNull(!ListPrice), "0", !ListPrice)
            GridProductPrices.Columns("WS").Value = !WSPrice
            GridProductPrices.Columns("Retail").Value = !RetailPrice
            .MoveNext
         Wend
         .Close
         GridProductPrices.MoveFirst
         GridProductPrices.Redraw = True
      End With
End Sub

Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.PurchaseReturnHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmPurchaseReturnInvoice) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmPurchaseReturnInvoice & ", " & vUser & "," & TableHeaderFields(eFrmPurchaseReturnInvoice) & " from PurchaseReturnHeader " & vbCrLf _
             & "Where ReturnID = " & TxtReturnID.Text & " and ReturnDate = '" & DtpReturnDate.DateValue & "'"
      CN.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.PurchaseReturnBodyBin (" & TableBodyFields(eFrmPurchaseReturnInvoice) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmPurchaseReturnInvoice) & " from PurchaseReturnBody " & vbCrLf _
             & "Where ReturnID = " & TxtReturnID.Text & " and ReturnDate = '" & DtpReturnDate.DateValue & "'"
      CN.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataPurchaseSerial()
   If RsPurchaseSerial.State = adStateOpen Then RsPurchaseSerial.Close
   vStrSQL = "select * from PurchaseBodySerial  "
   RsPurchaseSerial.Open vStrSQL, CN, adOpenDynamic, adLockBatchOptimistic
   RsPurchaseSerial.Filter = 0
End Sub
