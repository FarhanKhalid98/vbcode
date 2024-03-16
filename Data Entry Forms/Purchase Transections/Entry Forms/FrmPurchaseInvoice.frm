VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmPurchaseInvoice 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15255
   Icon            =   "FrmPurchaseInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1017
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkIsPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Print"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9540
      TabIndex        =   201
      Top             =   10080
      Width           =   1290
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmPurchaseInvoice.frx":000C
      Left            =   8235
      List            =   "FrmPurchaseInvoice.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   196
      Tag             =   "1"
      Top             =   9675
      Width           =   3276
   End
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   195
      Top             =   10080
      Width           =   2100
   End
   Begin VB.CheckBox ChkDiscB4SaleTax 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13950
      TabIndex        =   192
      Top             =   9720
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   10695
      TabIndex        =   172
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   9225
      Width           =   1170
   End
   Begin VB.Frame FrmProductPrices 
      Height          =   1095
      Left            =   6750
      TabIndex        =   170
      Top             =   390
      Visible         =   0   'False
      Width           =   6270
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridProductPrices 
         Height          =   885
         Left            =   60
         TabIndex        =   171
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
         stylesets(0).Picture=   "FrmPurchaseInvoice.frx":0010
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
         stylesets(1).Picture=   "FrmPurchaseInvoice.frx":002C
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
         stylesets(2).Picture=   "FrmPurchaseInvoice.frx":0048
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
   Begin VB.CheckBox ChkPurchaseReplacement 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Purchase Replacement"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   630
      TabIndex        =   163
      Top             =   4140
      Width           =   2100
   End
   Begin VB.CheckBox ChkDiscB4ExtraScheme 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13995
      TabIndex        =   159
      Top             =   8865
      Width           =   1290
   End
   Begin VB.CheckBox ChkDiscB4TradeOffer 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4 "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13995
      TabIndex        =   158
      Top             =   8145
      Width           =   1290
   End
   Begin VB.ComboBox CmbColourName 
      Height          =   315
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   148
      Top             =   375
      Width           =   1200
   End
   Begin VB.ComboBox cmbSizeName 
      Height          =   315
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   147
      Top             =   375
      Width           =   840
   End
   Begin VB.Frame FrmHistory 
      Height          =   1635
      Left            =   2445
      TabIndex        =   143
      Top             =   6465
      Visible         =   0   'False
      Width           =   11295
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridHistory 
         Height          =   1455
         Left            =   -750
         TabIndex        =   144
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
         stylesets(0).Picture=   "FrmPurchaseInvoice.frx":0064
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
         stylesets(1).Picture=   "FrmPurchaseInvoice.frx":0080
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
         stylesets(2).Picture=   "FrmPurchaseInvoice.frx":009C
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
   Begin VB.Frame FramExpense 
      Height          =   2415
      Left            =   8070
      TabIndex        =   120
      Top             =   5655
      Visible         =   0   'False
      Width           =   4215
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridExpense 
         Height          =   1860
         Left            =   120
         TabIndex        =   121
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
         stylesets(0).Picture=   "FrmPurchaseInvoice.frx":00B8
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
         stylesets(1).Picture=   "FrmPurchaseInvoice.frx":00D4
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
         stylesets(2).Picture=   "FrmPurchaseInvoice.frx":00F0
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
         TabIndex        =   122
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
         TabIndex        =   123
         Top             =   2100
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EFC09E&
      Caption         =   "Freight"
      Height          =   990
      Left            =   12300
      TabIndex        =   129
      Top             =   8213
      Width           =   1335
      Begin VB.OptionButton OptMe 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Me@Vender"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   137
         Top             =   225
         Width           =   1500
      End
      Begin VB.OptionButton OptExpense 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Expense"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   131
         Top             =   705
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.OptionButton OptVender 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Vender@Me"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   130
         Top             =   465
         Width           =   1500
      End
   End
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   4410
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   4680
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   645
      TabIndex        =   98
      Top             =   6015
      Visible         =   0   'False
      Width           =   2295
      Begin VB.TextBox TxtSerial 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   20
         TabIndex        =   99
         Top             =   180
         Width           =   2025
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridSerial 
         Height          =   1500
         Left            =   120
         TabIndex        =   100
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
         stylesets(0).Picture=   "FrmPurchaseInvoice.frx":010C
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
         Columns(1).Caption=   "Serial Purchase"
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
      Height          =   4380
      Left            =   15300
      TabIndex        =   92
      Top             =   135
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
         Height          =   3930
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   93
         Tag             =   "NC"
         Text            =   "FrmPurchaseInvoice.frx":0128
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
         TabIndex        =   94
         Top             =   90
         Width           =   135
      End
   End
   Begin SITextBox.Txt TxtVenderID 
      Height          =   315
      Left            =   1905
      TabIndex        =   7
      Top             =   2400
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
      TabIndex        =   63
      Top             =   2400
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
      TabIndex        =   62
      Top             =   2400
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
   Begin SITextBox.Txt TxtPurchaseID 
      Height          =   315
      Left            =   1875
      TabIndex        =   0
      Top             =   1665
      Width           =   555
      _ExtentX        =   979
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
      TabIndex        =   61
      Top             =   2400
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
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2835
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2400
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
      MICON           =   "FrmPurchaseInvoice.frx":023F
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8100
      TabIndex        =   58
      Top             =   9180
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
      MICON           =   "FrmPurchaseInvoice.frx":025B
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6750
      TabIndex        =   52
      Top             =   9180
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
      MICON           =   "FrmPurchaseInvoice.frx":0277
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4125
      TabIndex        =   54
      Top             =   9180
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
      MICON           =   "FrmPurchaseInvoice.frx":0293
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9405
      TabIndex        =   59
      Top             =   9180
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
      MICON           =   "FrmPurchaseInvoice.frx":02AF
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5445
      TabIndex        =   53
      Top             =   9180
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
      MICON           =   "FrmPurchaseInvoice.frx":02CB
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGrossAmount 
      Height          =   315
      Left            =   3090
      TabIndex        =   71
      Top             =   8595
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
      Left            =   4080
      TabIndex        =   43
      Top             =   8595
      Width           =   660
      _ExtentX        =   1164
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   645
      TabIndex        =   13
      Top             =   4680
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
      Left            =   1500
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4680
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
      MICON           =   "FrmPurchaseInvoice.frx":02E7
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpPurchaseDate 
      Height          =   315
      Left            =   2430
      TabIndex        =   1
      Top             =   1665
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   7755
      TabIndex        =   5
      Tag             =   "NC"
      Top             =   1665
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
      Left            =   8790
      TabIndex        =   78
      Tag             =   "NC"
      Top             =   1665
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
      Left            =   8430
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   1665
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
      MICON           =   "FrmPurchaseInvoice.frx":0303
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2835
      TabIndex        =   55
      Top             =   9180
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
      MICON           =   "FrmPurchaseInvoice.frx":031F
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillDisc 
      Height          =   315
      Left            =   4740
      TabIndex        =   44
      Top             =   8595
      Width           =   840
      _ExtentX        =   1482
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   7470
      TabIndex        =   86
      Top             =   450
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
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
   Begin SITextBox.Txt TxtTotalItems 
      Height          =   315
      Left            =   330
      TabIndex        =   89
      Top             =   8595
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
      Left            =   5580
      TabIndex        =   45
      Top             =   8595
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
   Begin JeweledBut.JeweledButton BtnBarCode 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   1485
      TabIndex        =   56
      Top             =   9180
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
      MICON           =   "FrmPurchaseInvoice.frx":033B
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnChangePrice 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   120
      TabIndex        =   57
      Top             =   9180
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   741
      TX              =   "Change Price"
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
      MICON           =   "FrmPurchaseInvoice.frx":0357
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtMultiplier 
      Height          =   315
      Left            =   5850
      TabIndex        =   19
      Top             =   4680
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
      Masked          =   1
   End
   Begin SITextBox.Txt TxtQtyLoose 
      Height          =   315
      Left            =   6870
      TabIndex        =   23
      Top             =   4680
      Width           =   540
      _ExtentX        =   953
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
      DecimalPoint    =   3
      IntegralPoint   =   7
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtQtyPack 
      Height          =   315
      Left            =   6360
      TabIndex        =   22
      Top             =   4680
      Width           =   510
      _ExtentX        =   900
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
      Left            =   8430
      TabIndex        =   26
      Top             =   4680
      Width           =   645
      _ExtentX        =   1138
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
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   11220
      TabIndex        =   30
      Top             =   4680
      Width           =   480
      _ExtentX        =   847
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
      DecimalPoint    =   5
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   9750
      TabIndex        =   28
      Top             =   4680
      Width           =   675
      _ExtentX        =   1191
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
      DecimalPoint    =   3
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtBonus 
      Height          =   315
      Left            =   7410
      TabIndex        =   24
      Top             =   4680
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
      Left            =   7950
      TabIndex        =   25
      Top             =   4680
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
   Begin SITextBox.Txt TxtSaleTaxPer 
      Height          =   315
      Left            =   12375
      TabIndex        =   32
      Top             =   4680
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   7
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
      IntegralPoint   =   4
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   13605
      TabIndex        =   41
      Top             =   4680
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridOffer 
      Height          =   1365
      Left            =   645
      TabIndex        =   114
      Top             =   6855
      Visible         =   0   'False
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
      stylesets(0).Picture=   "FrmPurchaseInvoice.frx":0373
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
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   10230
      TabIndex        =   6
      Tag             =   "NC"
      Top             =   1665
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
      Left            =   11535
      TabIndex        =   117
      Tag             =   "NC"
      Top             =   1665
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
      Left            =   11175
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   1665
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
      MICON           =   "FrmPurchaseInvoice.frx":038F
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtRetailPrice 
      Height          =   315
      Left            =   9075
      TabIndex        =   27
      Top             =   4680
      Width           =   675
      _ExtentX        =   1191
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
   Begin JeweledBut.JeweledButton BtnPurchaseOrder 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6090
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   1665
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
      MICON           =   "FrmPurchaseInvoice.frx":03AB
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrderID 
      Height          =   315
      Left            =   3735
      TabIndex        =   2
      Top             =   1665
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpOrderDate 
      Height          =   315
      Left            =   4785
      TabIndex        =   3
      Top             =   1665
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
   Begin SITextBox.Txt TxtFreight 
      Height          =   315
      Left            =   11265
      TabIndex        =   47
      Top             =   8595
      Width           =   840
      _ExtentX        =   1482
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpEntryDate 
      Height          =   315
      Left            =   6450
      TabIndex        =   4
      Top             =   1665
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
   Begin SITextBox.Txt TxtBillNo 
      Height          =   315
      Left            =   1905
      TabIndex        =   8
      Top             =   3120
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
      Left            =   2655
      TabIndex        =   9
      Top             =   3120
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
      Left            =   4920
      TabIndex        =   11
      Top             =   3120
      Width           =   4800
      _ExtentX        =   8467
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
   Begin SITextBox.Txt TxtVehicleNo 
      Height          =   315
      Left            =   3405
      TabIndex        =   10
      Top             =   3120
      Width           =   1515
      _ExtentX        =   2672
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
   Begin SITextBox.Txt TxtPaidAmount 
      Height          =   315
      Left            =   10275
      TabIndex        =   46
      Top             =   8595
      Width           =   990
      _ExtentX        =   1746
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
   Begin SITextBox.Txt TxtNetAmount 
      Height          =   315
      Left            =   6675
      TabIndex        =   138
      Top             =   8595
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
      Left            =   9195
      TabIndex        =   139
      Top             =   8595
      Width           =   1080
      _ExtentX        =   1905
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
      Left            =   7935
      TabIndex        =   140
      Top             =   8595
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
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   1860
      TabIndex        =   141
      Top             =   4680
      Width           =   2550
      _ExtentX        =   4498
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
   Begin SITextBox.Txt TxtBatchNo 
      Height          =   315
      Left            =   1410
      TabIndex        =   14
      Top             =   4365
      Width           =   675
      _ExtentX        =   1191
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
      Left            =   2085
      TabIndex        =   15
      Top             =   4365
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
   Begin JeweledBut.JeweledButton BtnProductRange 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   1050
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   4350
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
      MICON           =   "FrmPurchaseInvoice.frx":03C7
      BC              =   12632256
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3120
      Left            =   150
      TabIndex        =   75
      Top             =   4995
      Width           =   14595
      ScrollBars      =   3
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   49
      stylesets.count =   4
      stylesets(0).Name=   "Red"
      stylesets(0).ForeColor=   255
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "FrmPurchaseInvoice.frx":03E3
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
      stylesets(1).Picture=   "FrmPurchaseInvoice.frx":03FF
      stylesets(2).Name=   "Orange"
      stylesets(2).ForeColor=   33023
      stylesets(2).HasFont=   -1  'True
      BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(2).Picture=   "FrmPurchaseInvoice.frx":041B
      stylesets(3).Name=   "Green"
      stylesets(3).ForeColor=   4227072
      stylesets(3).HasFont=   -1  'True
      BeginProperty stylesets(3).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(3).Picture=   "FrmPurchaseInvoice.frx":0437
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
      Columns.Count   =   49
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ProductID"
      Columns(0).Name =   "ProductID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   873
      Columns(1).Caption=   "SrNo"
      Columns(1).Name =   "SrNo"
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
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "Colour"
      Columns(4).Name =   "ColourName"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "Size"
      Columns(5).Name =   "SizeName"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2540
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
      Columns(12).Caption=   "P Price"
      Columns(12).Name=   "Price"
      Columns(12).Alignment=   1
      Columns(12).CaptionAlignment=   2
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   5
      Columns(12).FieldLen=   256
      Columns(13).Width=   1191
      Columns(13).Caption=   "R Price"
      Columns(13).Name=   "RetailPrice"
      Columns(13).Alignment=   1
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   1191
      Columns(14).Caption=   "DiscPC"
      Columns(14).Name=   "DiscPC"
      Columns(14).Alignment=   1
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(15).Width=   1429
      Columns(15).Caption=   "DiscPack"
      Columns(15).Name=   "DiscPack"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   847
      Columns(16).Caption=   "Dis%"
      Columns(16).Name=   "DiscPer"
      Columns(16).Alignment=   1
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      Columns(17).Width=   1191
      Columns(17).Caption=   "Dis.Val"
      Columns(17).Name=   "DiscVal"
      Columns(17).Alignment=   1
      Columns(17).CaptionAlignment=   2
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   4
      Columns(17).FieldLen=   256
      Columns(18).Width=   926
      Columns(18).Caption=   "Tax%"
      Columns(18).Name=   "SaleTaxPer"
      Columns(18).Alignment=   1
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(19).Width=   1217
      Columns(19).Caption=   "SaleTaxVal"
      Columns(19).Name=   "SaleTaxVal"
      Columns(19).Alignment=   1
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(20).Width=   3200
      Columns(20).Visible=   0   'False
      Columns(20).Caption=   "SC"
      Columns(20).Name=   "SC"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      Columns(21).Width=   1508
      Columns(21).Caption=   "Amount"
      Columns(21).Name=   "Amount"
      Columns(21).Alignment=   1
      Columns(21).CaptionAlignment=   2
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   5
      Columns(21).FieldLen=   256
      Columns(22).Width=   3200
      Columns(22).Visible=   0   'False
      Columns(22).Caption=   "PackingID"
      Columns(22).Name=   "PackingID"
      Columns(22).DataField=   "Column 22"
      Columns(22).DataType=   8
      Columns(22).FieldLen=   256
      Columns(23).Width=   1058
      Columns(23).Caption=   "Qty(G)"
      Columns(23).Name=   "GrossQty"
      Columns(23).DataField=   "Column 23"
      Columns(23).DataType=   8
      Columns(23).FieldLen=   256
      Columns(24).Width=   1058
      Columns(24).Caption=   "Qty(U)"
      Columns(24).Name=   "GrossUnit"
      Columns(24).DataField=   "Column 24"
      Columns(24).DataType=   8
      Columns(24).FieldLen=   256
      Columns(25).Width=   2249
      Columns(25).Caption=   "IsWSDiscb4ST"
      Columns(25).Name=   "IsWSDiscb4ST"
      Columns(25).DataField=   "Column 25"
      Columns(25).DataType=   11
      Columns(25).FieldLen=   256
      Columns(26).Width=   2249
      Columns(26).Caption=   "IsWSSaleTax"
      Columns(26).Name=   "IsWSSaleTax"
      Columns(26).DataField=   "Column 26"
      Columns(26).DataType=   11
      Columns(26).FieldLen=   256
      Columns(27).Width=   2461
      Columns(27).Caption=   "IsRetailSaleTax"
      Columns(27).Name=   "IsRetailSaleTax"
      Columns(27).DataField=   "Column 27"
      Columns(27).DataType=   11
      Columns(27).FieldLen=   256
      Columns(28).Width=   1931
      Columns(28).Caption=   "BatchNo"
      Columns(28).Name=   "BatchNo"
      Columns(28).DataField=   "Column 28"
      Columns(28).DataType=   8
      Columns(28).FieldLen=   256
      Columns(29).Width=   2090
      Columns(29).Caption=   "ExpiryDate"
      Columns(29).Name=   "ExpiryDate"
      Columns(29).DataField=   "Column 29"
      Columns(29).DataType=   8
      Columns(29).FieldLen=   256
      Columns(30).Width=   3200
      Columns(30).Visible=   0   'False
      Columns(30).Caption=   "ExpiryTime"
      Columns(30).Name=   "ExpiryTime"
      Columns(30).DataField=   "Column 30"
      Columns(30).DataType=   8
      Columns(30).FieldLen=   256
      Columns(31).Width=   3200
      Columns(31).Visible=   0   'False
      Columns(31).Caption=   "ColourID"
      Columns(31).Name=   "ColourID"
      Columns(31).DataField=   "Column 31"
      Columns(31).DataType=   8
      Columns(31).FieldLen=   256
      Columns(32).Width=   3200
      Columns(32).Visible=   0   'False
      Columns(32).Caption=   "SizeID"
      Columns(32).Name=   "SizeID"
      Columns(32).DataField=   "Column 32"
      Columns(32).DataType=   8
      Columns(32).FieldLen=   256
      Columns(33).Width=   3200
      Columns(33).Caption=   "isDiscB4TradeOffer"
      Columns(33).Name=   "isDiscB4TradeOffer"
      Columns(33).DataField=   "Column 33"
      Columns(33).DataType=   8
      Columns(33).FieldLen=   256
      Columns(34).Width=   3200
      Columns(34).Caption=   "isDiscB4ExtraScheme"
      Columns(34).Name=   "isDiscB4ExtraScheme"
      Columns(34).DataField=   "Column 34"
      Columns(34).DataType=   8
      Columns(34).FieldLen=   256
      Columns(35).Width=   3200
      Columns(35).Caption=   "isDiscB4SaleTax"
      Columns(35).Name=   "isDiscB4SaleTax"
      Columns(35).DataField=   "Column 35"
      Columns(35).DataType=   8
      Columns(35).FieldLen=   256
      Columns(36).Width=   3200
      Columns(36).Caption=   "TradeOffer1"
      Columns(36).Name=   "TradeOffer1"
      Columns(36).DataField=   "Column 36"
      Columns(36).DataType=   8
      Columns(36).FieldLen=   256
      Columns(37).Width=   3200
      Columns(37).Caption=   "TradeOffer2"
      Columns(37).Name=   "TradeOffer2"
      Columns(37).DataField=   "Column 37"
      Columns(37).DataType=   8
      Columns(37).FieldLen=   256
      Columns(38).Width=   3200
      Columns(38).Caption=   "ExtraSchemePer"
      Columns(38).Name=   "ExtraSchemePer"
      Columns(38).DataField=   "Column 38"
      Columns(38).DataType=   8
      Columns(38).FieldLen=   256
      Columns(39).Width=   3200
      Columns(39).Caption=   "TradeValue"
      Columns(39).Name=   "TradeValue"
      Columns(39).DataField=   "Column 39"
      Columns(39).DataType=   8
      Columns(39).FieldLen=   256
      Columns(40).Width=   3200
      Columns(40).Caption=   "ExtraSchemeValue"
      Columns(40).Name=   "ExtraSchemeValue"
      Columns(40).DataField=   "Column 40"
      Columns(40).DataType=   8
      Columns(40).FieldLen=   256
      Columns(41).Width=   3200
      Columns(41).Caption=   "DiscAmount"
      Columns(41).Name=   "DiscAmount"
      Columns(41).DataField=   "Column 41"
      Columns(41).DataType=   8
      Columns(41).FieldLen=   256
      Columns(42).Width=   3200
      Columns(42).Caption=   "OldPrice"
      Columns(42).Name=   "OldPrice"
      Columns(42).DataField=   "Column 42"
      Columns(42).DataType=   8
      Columns(42).FieldLen=   256
      Columns(43).Width=   1958
      Columns(43).Caption=   "RetailAmount"
      Columns(43).Name=   "RetailAmount"
      Columns(43).DataField=   "Column 43"
      Columns(43).DataType=   8
      Columns(43).FieldLen=   256
      Columns(44).Width=   3200
      Columns(44).Caption=   "ProfitAmount"
      Columns(44).Name=   "ProfitAmount"
      Columns(44).DataField=   "Column 44"
      Columns(44).DataType=   8
      Columns(44).FieldLen=   256
      Columns(45).Width=   2037
      Columns(45).Caption=   "SaleDiscPer"
      Columns(45).Name=   "SaleDiscPer"
      Columns(45).DataField=   "Column 45"
      Columns(45).DataType=   8
      Columns(45).FieldLen=   256
      Columns(46).Width=   3200
      Columns(46).Visible=   0   'False
      Columns(46).Caption=   "IsSerial"
      Columns(46).Name=   "IsSerial"
      Columns(46).DataField=   "Column 46"
      Columns(46).DataType=   11
      Columns(46).FieldLen=   256
      Columns(47).Width=   3200
      Columns(47).Caption=   "DiscPer2"
      Columns(47).Name=   "DiscPer2"
      Columns(47).DataField=   "Column 47"
      Columns(47).DataType=   8
      Columns(47).FieldLen=   256
      Columns(48).Width=   3200
      Columns(48).Caption=   "DiscVal2"
      Columns(48).Name=   "DiscVal2"
      Columns(48).DataField=   "Column 48"
      Columns(48).DataType=   8
      Columns(48).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   25744
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpPromiseDate 
      Height          =   315
      Left            =   4770
      TabIndex        =   145
      Top             =   1095
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
   Begin SITextBox.Txt TxtRemarks 
      Height          =   315
      Left            =   1905
      TabIndex        =   12
      Top             =   3720
      Width           =   7860
      _ExtentX        =   13864
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
   End
   Begin SITextBox.Txt TxtGrossQty 
      Height          =   315
      Left            =   6675
      TabIndex        =   20
      Top             =   4095
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
   Begin SITextBox.Txt TxtDiscVal 
      Height          =   315
      Left            =   11700
      TabIndex        =   31
      Top             =   4680
      Width           =   675
      _ExtentX        =   1191
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
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtSC 
      Height          =   315
      Left            =   14265
      TabIndex        =   42
      Top             =   3240
      Width           =   690
      _ExtentX        =   1217
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
   Begin SITextBox.Txt TxtGrossUnit 
      Height          =   315
      Left            =   8115
      TabIndex        =   21
      Top             =   4095
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
   Begin SITextBox.Txt TxtTradeOffer2 
      Height          =   315
      Left            =   13425
      TabIndex        =   37
      Top             =   4140
      Width           =   630
      _ExtentX        =   1111
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
   Begin SITextBox.Txt TxtTradeOffer1 
      Height          =   315
      Left            =   12660
      TabIndex        =   36
      Top             =   4140
      Width           =   585
      _ExtentX        =   1032
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
   Begin SITextBox.Txt TxtExtraSchemePer 
      Height          =   315
      Left            =   14175
      TabIndex        =   38
      Top             =   4140
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
      Masked          =   2
      DecimalPoint    =   3
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtTradeOfferValue 
      Height          =   315
      Left            =   13995
      TabIndex        =   39
      Top             =   8505
      Width           =   870
      _ExtentX        =   1535
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
      DecimalPoint    =   3
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtExtraSchemeValue 
      Height          =   315
      Left            =   13995
      TabIndex        =   40
      Top             =   9315
      Width           =   870
      _ExtentX        =   1535
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
      DecimalPoint    =   3
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtSaleTaxVal 
      Height          =   315
      Left            =   12915
      TabIndex        =   33
      Top             =   4680
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
      DecimalPoint    =   3
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtExtraTaxVal 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11880
      TabIndex        =   50
      Top             =   10020
      Width           =   1200
      _ExtentX        =   2117
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
   Begin SITextBox.Txt TxtExtraTaxPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   13080
      TabIndex        =   51
      Top             =   10020
      Width           =   570
      _ExtentX        =   1005
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
   Begin SITextBox.Txt TxtAdvTaxVal 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11880
      TabIndex        =   48
      Top             =   9465
      Width           =   1200
      _ExtentX        =   2117
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
   Begin SITextBox.Txt TxtAdvTaxPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   13080
      TabIndex        =   49
      Top             =   9465
      Width           =   570
      _ExtentX        =   1005
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
   Begin SITextBox.Txt TxtSumDiscAmount 
      Height          =   315
      Left            =   2205
      TabIndex        =   168
      Top             =   8595
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
   Begin SITextBox.Txt TxtBarCode 
      Height          =   315
      Left            =   3735
      TabIndex        =   16
      Top             =   4095
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   20
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
   Begin JeweledBut.JeweledButton BtnAddBarCode 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5355
      TabIndex        =   175
      TabStop         =   0   'False
      Top             =   4095
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "Add"
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
      MICON           =   "FrmPurchaseInvoice.frx":0453
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtTerms 
      Height          =   315
      Left            =   3015
      TabIndex        =   177
      Tag             =   "NC"
      Top             =   1095
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   556
      Alignment       =   1
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
      DecimalPoint    =   3
      IntegralPoint   =   3
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtProductName2 
      Height          =   315
      Left            =   135
      TabIndex        =   179
      Top             =   9675
      Width           =   7200
      _ExtentX        =   12700
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
      Masked          =   5
   End
   Begin SITextBox.Txt TxtRetailAmount 
      Height          =   315
      Left            =   90
      TabIndex        =   181
      Top             =   1770
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
      Left            =   90
      TabIndex        =   183
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
   Begin SITextBox.Txt TxtSaleDiscPer 
      Height          =   315
      Left            =   90
      TabIndex        =   185
      Top             =   1140
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
   Begin SITextBox.Txt TxtDiscAmount 
      Height          =   315
      Left            =   135
      TabIndex        =   187
      Top             =   3075
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
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
      Masked          =   1
   End
   Begin SITextBox.Txt TxtPhoneNo 
      Height          =   315
      Left            =   13140
      TabIndex        =   189
      Top             =   2400
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
   Begin SITextBox.Txt TxtDiscPack 
      Height          =   315
      Left            =   10425
      TabIndex        =   29
      Top             =   4680
      Width           =   810
      _ExtentX        =   1429
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
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   1215
      TabIndex        =   193
      Top             =   8595
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
   Begin SITextBox.Txt TxtSID 
      Height          =   315
      Left            =   1800
      TabIndex        =   197
      Top             =   1020
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
   Begin SITextBox.Txt TxtDiscPer2 
      Height          =   315
      Left            =   10500
      TabIndex        =   34
      Top             =   4110
      Width           =   480
      _ExtentX        =   847
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
      DecimalPoint    =   5
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtDiscVal2 
      Height          =   315
      Left            =   10980
      TabIndex        =   35
      Top             =   4110
      Width           =   675
      _ExtentX        =   1191
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
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin VB.Label LblDiscPer2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Dis2%"
      Height          =   195
      Left            =   10485
      TabIndex        =   200
      Top             =   3915
      Width           =   435
   End
   Begin VB.Label LblDiscVal2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.Val 2"
      Height          =   195
      Left            =   10965
      TabIndex        =   199
      Top             =   3915
      Width           =   720
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "SID"
      Height          =   195
      Left            =   1800
      TabIndex        =   198
      Top             =   810
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label LblTotalAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   1215
      TabIndex        =   194
      Top             =   8370
      Width           =   945
   End
   Begin VB.Label LblDiscPack 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Pack"
      Height          =   195
      Left            =   10440
      TabIndex        =   191
      Top             =   4485
      Width           =   735
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No"
      Height          =   195
      Left            =   13140
      TabIndex        =   190
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label LblDiscAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Amount"
      Height          =   195
      Left            =   135
      TabIndex        =   188
      Top             =   2835
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label LblSaleDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Disc Per"
      Height          =   195
      Left            =   90
      TabIndex        =   186
      Top             =   945
      Width           =   960
   End
   Begin VB.Label LblProfitAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Profit Amount"
      Height          =   195
      Left            =   90
      TabIndex        =   184
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label LblRetailAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Retail Amount"
      Height          =   195
      Left            =   90
      TabIndex        =   182
      Top             =   1575
      Width           =   990
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
      Left            =   10035
      TabIndex        =   180
      Top             =   3645
      Width           =   1905
   End
   Begin VB.Label LblTerms 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Terms"
      Height          =   195
      Left            =   2475
      TabIndex        =   178
      Top             =   1125
      Width           =   435
   End
   Begin VB.Label LblBarCode 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Add BarCode"
      Height          =   195
      Left            =   2745
      TabIndex        =   176
      Top             =   4140
      Width           =   945
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
      Left            =   10710
      TabIndex        =   174
      Top             =   8970
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
      Left            =   7515
      TabIndex        =   173
      Top             =   9720
      Width           =   570
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Amt"
      Height          =   195
      Left            =   2205
      TabIndex        =   169
      Top             =   8370
      Width           =   630
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   13080
      TabIndex        =   167
      Top             =   9795
      Width           =   210
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Tax"
      Height          =   195
      Left            =   11880
      TabIndex        =   166
      Top             =   9795
      Width           =   675
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Tax"
      Height          =   195
      Left            =   11880
      TabIndex        =   165
      Top             =   9240
      Width           =   960
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   13080
      TabIndex        =   164
      Top             =   9240
      Width           =   210
   End
   Begin VB.Label LblTradeValue 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Value"
      Height          =   195
      Left            =   13995
      TabIndex        =   162
      Top             =   8325
      Width           =   870
   End
   Begin VB.Label LblExtraSchemeValue 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ex Scheme Value"
      Height          =   195
      Left            =   13995
      TabIndex        =   161
      Top             =   9090
      Width           =   1260
   End
   Begin VB.Label LblSaleTaxVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Val"
      Height          =   195
      Left            =   12915
      TabIndex        =   160
      Top             =   4485
      Width           =   540
   End
   Begin VB.Label LblPlusSign 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Left            =   13290
      TabIndex        =   157
      Top             =   4185
      Width           =   120
   End
   Begin VB.Label LblTradeOffer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Offer"
      Height          =   195
      Left            =   12600
      TabIndex        =   156
      Top             =   3915
      Width           =   810
   End
   Begin VB.Label LblExtraSchemePer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ex. Scheme %"
      Height          =   195
      Left            =   13815
      TabIndex        =   155
      Top             =   3915
      Width           =   1020
   End
   Begin VB.Label LblGrossUnit 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Unit"
      Height          =   195
      Left            =   7335
      TabIndex        =   154
      Top             =   4140
      Width           =   735
   End
   Begin VB.Label LblSC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "S.C."
      Height          =   195
      Left            =   13860
      TabIndex        =   153
      Top             =   3240
      Width           =   300
   End
   Begin VB.Label LblGrossQty 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Qty"
      Height          =   195
      Left            =   5895
      TabIndex        =   152
      Top             =   4140
      Width           =   690
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   1905
      TabIndex        =   151
      Top             =   3510
      Width           =   630
   End
   Begin VB.Label LblColour 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
      Height          =   195
      Left            =   135
      TabIndex        =   150
      Top             =   180
      Width           =   450
   End
   Begin VB.Label LblSize 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   1335
      TabIndex        =   149
      Top             =   180
      Width           =   300
   End
   Begin VB.Label LblPromiseDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Promise Date"
      Height          =   195
      Left            =   3780
      TabIndex        =   146
      Top             =   1125
      Width           =   945
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   4920
      TabIndex        =   136
      Top             =   2910
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bilty No."
      Height          =   195
      Left            =   2640
      TabIndex        =   135
      Top             =   2910
      Width           =   585
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      Height          =   195
      Left            =   1905
      TabIndex        =   134
      Top             =   2910
      Width           =   495
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      Height          =   195
      Left            =   3390
      TabIndex        =   133
      Top             =   2910
      Width           =   780
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date"
      Height          =   195
      Left            =   6450
      TabIndex        =   132
      Top             =   1470
      Width           =   750
   End
   Begin VB.Label LblFreight 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      Height          =   195
      Left            =   11265
      TabIndex        =   128
      Top             =   8370
      Width           =   480
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
      Height          =   195
      Left            =   4785
      TabIndex        =   127
      Top             =   1470
      Width           =   780
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID"
      Height          =   195
      Left            =   3735
      TabIndex        =   126
      Top             =   1470
      Width           =   600
   End
   Begin VB.Label lblPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Payable"
      Height          =   195
      Left            =   7935
      TabIndex        =   124
      Top             =   8363
      Width           =   1260
   End
   Begin VB.Label LblRetail 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "R Price"
      Height          =   195
      Left            =   9090
      TabIndex        =   119
      Top             =   4485
      Width           =   525
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   11535
      TabIndex        =   116
      Top             =   1470
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   10230
      TabIndex        =   115
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   13575
      TabIndex        =   113
      Top             =   4485
      Width           =   540
   End
   Begin VB.Label LblSaleTaxPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax%"
      Height          =   195
      Left            =   12360
      TabIndex        =   112
      Top             =   4485
      Width           =   390
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Val"
      Height          =   195
      Left            =   8835
      TabIndex        =   111
      Top             =   150
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label LblOffer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Offer"
      Height          =   195
      Left            =   7950
      TabIndex        =   110
      Top             =   4485
      Width           =   345
   End
   Begin VB.Label LblPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   8415
      TabIndex        =   109
      Top             =   4485
      Width           =   405
   End
   Begin VB.Label LblMultiplier 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   5880
      TabIndex        =   108
      Top             =   4485
      Width           =   375
   End
   Begin VB.Label LblPackName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   4455
      TabIndex        =   107
      Top             =   4485
      Width           =   840
   End
   Begin VB.Label LblQtyLoose 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   6870
      TabIndex        =   106
      Top             =   4485
      Width           =   465
   End
   Begin VB.Label LblQtyPack 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   6360
      TabIndex        =   105
      Top             =   4485
      Width           =   480
   End
   Begin VB.Label LblDiscVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.Val"
      Height          =   195
      Left            =   11685
      TabIndex        =   104
      Top             =   4485
      Width           =   585
   End
   Begin VB.Label LblDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Dis%"
      Height          =   195
      Left            =   11235
      TabIndex        =   103
      Top             =   4485
      Width           =   345
   End
   Begin VB.Label LblDiscPC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc/PC"
      Height          =   195
      Left            =   9750
      TabIndex        =   102
      Top             =   4485
      Width           =   600
   End
   Begin VB.Label LblBonus 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bns(L)"
      Height          =   195
      Left            =   7410
      TabIndex        =   101
      Top             =   4485
      Width           =   450
   End
   Begin VB.Label LblRetailPrice 
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
      Height          =   1140
      Left            =   12120
      TabIndex        =   97
      Top             =   3135
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   12120
      TabIndex        =   96
      Top             =   2775
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   12885
      TabIndex        =   95
      Top             =   1650
      Width           =   435
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      Height          =   195
      Left            =   5580
      TabIndex        =   91
      Top             =   8363
      Width           =   1020
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items"
      Height          =   195
      Left            =   330
      TabIndex        =   90
      Top             =   8370
      Width           =   780
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Invoice"
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
      TabIndex        =   88
      Top             =   270
      Width           =   3000
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   7470
      TabIndex        =   87
      Top             =   135
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LblTtlPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payable"
      Height          =   195
      Left            =   9195
      TabIndex        =   85
      Top             =   8363
      Width           =   975
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   4815
      TabIndex        =   84
      Top             =   8370
      Width           =   630
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
      Left            =   10530
      TabIndex        =   83
      Top             =   3045
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
      Left            =   10410
      TabIndex        =   82
      Top             =   3360
      Width           =   1065
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   8790
      TabIndex        =   81
      Top             =   1470
      Width           =   840
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   7755
      TabIndex        =   80
      Top             =   1470
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   645
      TabIndex        =   77
      Top             =   4485
      Width           =   375
   End
   Begin VB.Label LblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3345
      TabIndex        =   76
      Top             =   4485
      Width           =   1020
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label LblNetAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      Height          =   195
      Left            =   6675
      TabIndex        =   74
      Top             =   8363
      Width           =   840
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. (%)"
      Height          =   195
      Left            =   4080
      TabIndex        =   73
      Top             =   8363
      Width           =   615
   End
   Begin VB.Label LblGrossAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amount"
      Height          =   195
      Left            =   3045
      TabIndex        =   72
      Top             =   8370
      Width           =   990
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Amount"
      Height          =   195
      Left            =   10275
      TabIndex        =   70
      Top             =   8363
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   11370
      TabIndex        =   69
      Top             =   2190
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   6840
      TabIndex        =   68
      Top             =   2190
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Name"
      Height          =   195
      Left            =   3195
      TabIndex        =   67
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender ID"
      Height          =   195
      Left            =   1890
      TabIndex        =   66
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Date"
      Height          =   195
      Left            =   2430
      TabIndex        =   65
      Top             =   1470
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur ID"
      Height          =   195
      Left            =   1890
      TabIndex        =   64
      Top             =   1470
      Width           =   450
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmPurchaseInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim vDate, vNow, vServerDate As Date, vHDiff As Integer, vSystemDate As Boolean
Dim vMode As FormMode, vNewRow As Boolean
Dim vUnitPrice, vOldPrice As Double
Dim vUnitWSPrice As Double
Dim vUnitRetailPrice, vSaleDiscPer As Double
Dim vIsWSDiscb4ST As Boolean
Dim vIsRetailSaleTax As Boolean
Dim vIsWSSaleTax As Boolean
Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim vMaxBinID, vGridRows As Integer
Dim RsBody As New ADODB.Recordset
Dim RsBodySerial As New ADODB.Recordset
Dim RsProductOffer As New ADODB.Recordset
Dim RsExpense As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim QtyOffer As Integer
Dim Rebate As Integer
Dim Flag As Boolean
Dim ssql As String, vExpiryTime, vExpiryColor As String
Dim vStrSQL, vRandomID  As String
Dim vPurchaseID  As Integer
Dim vPurchaseDate  As Date
Dim ExpenseFlag As Boolean
Dim vExpAmount As Double
Dim vQtyLoose As Double
Dim vMargin As String
Dim vColour, vTradeOffer, vIsSerial As Boolean
Dim vMobileNo() As String, vMobile As String
Dim i As Integer, vDiscPackFlag, vAdvDiscPerFlag As Boolean
Dim vDelProductID, vDel() As String
Dim vNoofPrints As Byte
Dim vPrinter() As String
Dim vBarcode As String
Dim vShowStock As Boolean

'----------------------------------
Private Sub SubCalculateBody()
    On Error GoTo ErrorHandler
    If vDiscPackFlag = False Then
      If ActiveControl.Name <> TxtDiscVal.Name Then
         TxtDiscVal.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) * Val(TxtDiscPer.Text) / 100, IIf(ObjRegistry.IsRoundFigure, 0, 2))
      End If
      If ActiveControl.Name <> TxtDiscPC.Name Then
         TxtDiscPC.Text = Round((vUnitPrice * Val(TxtDiscPer.Text) / 100), IIf(ObjRegistry.IsRoundFigure, 1, 4))
         If Val(TxtDiscPC.Text) = 0 Then TxtDiscPC.Text = ""
      End If
      If TxtDiscPack.Visible = True Then
         If ActiveControl.Name <> TxtDiscPack.Name Then
            TxtDiscPack.Text = Round(Val(TxtDiscPC.Text) * Val(TxtMultiplier.Text), 0)
         End If
      End If
      TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), IIf(ObjRegistry.IsRoundFigure, 0, 2))
   Else
      If ActiveControl.Name <> TxtDiscPC.Name Then
         TxtDiscPC.Text = Round(Val(TxtDiscPack.Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text)), IIf(ObjRegistry.IsRoundFigure, 1, 4))
      End If
      If Val(TxtDiscPack.Text) <> 0 And TxtDiscPack.Visible And vUnitPrice <> 0 Then
         TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, IIf(ObjRegistry.IsRoundFigure, 3, 5))
      End If
      If ActiveControl.Name <> TxtDiscVal.Name Then
         TxtDiscVal.Text = Round(Val(TxtDiscPC.Text) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), IIf(ObjRegistry.IsRoundFigure, 0, 2))
      End If
      TxtAmount.Text = Round(((Val(vUnitPrice) - Val(TxtDiscPC.Text)) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))), IIf(ObjRegistry.IsRoundFigure, 0, 2))
   End If
   If Val(TxtDiscVal.Text) = 0 Then TxtDiscVal.Text = ""
   
'   If ActiveControl.Name <> TxtSaleTaxVal.Name Then
'
'   End If
   
   If vTradeOffer = True Then CalculateValue
   
   TxtAmount.Text = (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)))
   TxtAmount.Text = Round(Val(TxtAmount.Text) - Val(TxtDiscVal.Text) - Val(TxtTradeOfferValue.Text) - Val(TxtExtraSchemeValue.Text) + Val(TxtSaleTaxVal.Text) + Val(TxtSC.Text) - Val(TxtOffer.Text), 2)
'   TxtAmount.Text = Val(TxtAmount.Text) + Val(TxtSaleTaxVal.Text) - Val(TxtOffer.Text) + Val(TxtSC.Text)
   If ObjRegistry.IsRoundFigure = True Then TxtAmount.Text = SelfRound(TxtAmount.Text)
   TxtDiscAmount.Text = Val(TxtDiscVal.Text)
   
'   TxtDiscVal2.Text = SelfRound((Val(TxtAmount.Text) * Val(TxtDiscPer2.Text) / 100))
'   TxtDiscPer2.Text = Round((Val(TxtDiscVal2.Text) * 100) / IIf(Val(TxtAmount.Text) = 0, 1, Val(TxtAmount.Text)), 2)
   TxtAmount.Text = TxtAmount.Text - Val(TxtDiscVal2.Text)
'   Call CalculateValue
    
     ''''''Calculate Profit Amount
   TxtRetailAmount.Text = Round((Val(vUnitRetailPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitRetailPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleDiscPer.Text) / 100), 2)
   TxtProfitAmount.Text = Val(TxtRetailAmount.Text) - Val(TxtAmount.Text)
   '''''''''''
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateFooter()
   On Error GoTo ErrorHandler
   If TxtGrossAmount.Text = "" Then Exit Sub
   TxtNetAmount.Text = SelfRound(Val(TxtGrossAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text)
   If Val(TxtAdvTaxVal.Text) <> 0 And Val(TxtNetAmount.Text) <> 0 Then
      If vAdvDiscPerFlag = False Then
         TxtAdvTaxPer.Text = Round((Val(TxtAdvTaxVal.Text) * 100) / IIf(Val(TxtNetAmount.Text) = 0, 1, Val(TxtNetAmount.Text)), 2)
      Else
         TxtAdvTaxVal.Text = SelfRound((Val(TxtNetAmount.Text) * Val(TxtAdvTaxPer.Text) / 100))
      End If
   End If
   TxtNetAmount.Text = SelfRound(Val(TxtGrossAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text) + Val(TxtAdvTaxVal.Text) + Val(TxtExtraTaxVal.Text)
   TxtTotalPayable.Text = Abs(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousPayable.Text, Val(TxtPreviousPayable.Text) * -1)))
   LblTtlPayable.Caption = IIf(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousPayable.Text, Val(TxtPreviousPayable.Text) * -1)) < 0, "Total Receivable", "Total Payable")
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearSerialFields()
   On Error GoTo ErrorHandler
   TxtSerial.Text = ""
'   TxtSerial.Enabled = False
   TxtSerial.Enabled = True
   GridSerial.CancelUpdate
   GridSerial.RemoveAll
   GridSerial.AddNew
   GridSerial.Columns("Serial").Text = " "
   GridSerial.Update
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
    vStrSQL = " Select * FROM Stores where StoreID = " & Val(TxtStoreID.Text)
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
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%') and c.isLocked = 0 and c.isDetailed = 1"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectVender = False: Exit Function
        TxtVenderID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select c.AccountNo, c.AccountName as AccountName, Address, City, Description, isnull(p.phone1,'') + isnull(' ' + p.phone2,'') + isnull(' ' + p.Mobile,'') + isnull(' '+p.Mobile2,'') as PhoneNo " & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " where c.AccountNo = " & Val(TxtVenderID.Text) & " and (c.AccountNo like '6%') and isDetailed = 1 and isLocked = 0"
    
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtVenderName.Text = !AccountName
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtPhoneNo.Text = IIf(IsNull(!PhoneNo), "", !PhoneNo)
          TxtPreviousPayable.Text = CN.Execute("SELECT isnull(dbo.FunCurrentDebit(" & Val(TxtVenderID.Text) & ",'" & DtpPurchaseDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
          vStrSQL = " Select isnull(Sum(TotalAmount - isnull(BillDisc,0) + isnull(OtherCharges,0)),0) as Amount " & vbCrLf _
                  + " FROM PurchaseHeader h INNER JOIN (Select PurId, PurchaseDate, Sum(amount) TTLValue FROM PurchaseBody Group By PurId, PurchaseDate)B " & vbCrLf _
                  + " ON h.PurId = B.PurId and h.PurchaseDate = B.PurchaseDate " & vbCrLf _
                  + " where VendorID = " & Val(TxtVenderID.Text) & " and h.PurchaseDate = '" & DtpPurchaseDate.DateValue & "' and h.PurID >= " & Val(TxtPurchaseID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
          TxtPreviousPayable.Text = TxtPreviousPayable.Text - CN.Execute(vStrSQL).Fields(0).Value
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
          TxtPhoneNo.Text = ""
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
'         SchProduct.ParaInWhere = " and isLocked = 0 and isNoCostProduct = 0 and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
         SchProduct.ParaInWhere = " and isLocked = 0 and isNoCostProduct = 0 " & IIf(ObjRegistry.ProductSearchWithStore = True, " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")", "")
         SchProduct.ParaInPurchase = True
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
             & "where ItemCode = '" & IIf(Len(TxtCode.Text) = 9, TxtCode.Text & "'", Mid(TxtCode.Text, 1, 9) & "' and c.colourid = " & Val(Mid(TxtCode.Text, 10, 2))) & " or P.ProductID = " & TxtCode.Text
      With CN.Execute(ssql)
         If .RecordCount > 0 Then
            CmbColourName.AddItem !ColourName
            CmbColourName.ItemData(CmbColourName.NewIndex) = !ColourID
            CmbColourName.ListIndex = 0
         End If
      End With
      
      ssql = "select s.SizeID, SizeName from productSizes pz inner join Sizes s on pz.Sizeid = s.Sizeid " & vbCrLf _
      & "inner join products p on p.productid = pz.productid " & vbCrLf _
      & "where ItemCode = '" & IIf(Len(TxtCode.Text) = 13, Mid(TxtCode.Text, 1, 9) & "' and s.sizeid = " & Val(Mid(TxtCode.Text, 12, 2)), TxtCode.Text & "'") & " or P.ProductID = " & TxtCode.Text
      
      With CN.Execute(ssql)
         If .RecordCount > 0 Then
            cmbSizeName.AddItem !SizeName
            cmbSizeName.ItemData(cmbSizeName.NewIndex) = !SizeID
            cmbSizeName.ListIndex = 0
         End If
      End With
      TxtCode.Text = CStr(Left(TxtCode.Text, 9))
   End If
   CmbPackName.Clear
   vStrSQL = "select distinct pp.PackingID, Packingname from ProductPacking pp inner join packings p on p.packingid = pp.packingid" & vbCrLf _
           + "left outer join ProductBarcodes b on b.productid = pp.productid" & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "pp.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')"
           
   With CN.Execute(vStrSQL)
      CmbPackName.AddItem ""
      While Not .EOF
         CmbPackName.AddItem !PackingName
         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
         .MoveNext
      Wend
      .Close
   End With
   
   ''''''''***********   Prefix BarCode For Label Weight Machine   ***********''''''''
   If ObjRegistry.BarCodePrefix <> 0 Then
      vBarcode = TxtCode.Text
      If ObjRegistry.BarCodePrefix = Mid(vBarcode, 1, 2) And Len(vBarcode) > 5 Then
         TxtCode.Text = Mid(vBarcode, 3, 5)
      End If
   End If
   '''''''''''''''''''''''''''''''
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
        vStrSQL = " SELECT p.productid, Code, Qty, ProductName, PurPrice, WSPrice, RetailPrice, DiscPer, " & vbCrLf _
           + " IsSerial, IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST, IsDiscB4TradeOffer, IsDiscB4ExtraScheme, isDiscB4SaleTax, TradeOffer1, TradeOffer2, ExtraSchemePer,  " & vbCrLf _
           + " SaleTaxPer, PurDiscPC, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')" & " and isLocked = 0 "
           
 
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !PurPrice
         TxtRetailPrice.Text = !RetailPrice
         TxtSaleDiscPer.Text = IIf(IsNull(!DiscPer), "", !DiscPer)
         vIsSerial = !IsSerial
         vIsWSDiscb4ST = !IsWSDiscb4ST
         vIsWSSaleTax = !IsWSSaleTax
         vIsRetailSaleTax = !IsRetailSaleTax
         ChkDiscB4TradeOffer.Value = Abs(!isDiscB4TradeOffer)
         ChkDiscB4ExtraScheme.Value = Abs(!IsDiscB4ExtraScheme)
         ChkDiscB4SaleTax.Value = Abs(!isDiscB4SaleTax)
         TxtTradeOffer1.Text = IIf(IsNull(!TradeOffer1), 0, !TradeOffer1)
         TxtTradeOffer2.Text = IIf(IsNull(!TradeOffer2), 0, !TradeOffer2)
         TxtExtraSchemePer.Text = IIf(IsNull(!ExtraSchemePer), 0, !ExtraSchemePer)
         TxtSaleTaxPer.Text = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
         TxtDiscPC.Text = IIf(IsNull(!PurDiscPC), "", !PurDiscPC)
         LblRetailPrice.Caption = !RetailPrice
         If ObjRegistry.IsAllowBarCodeQtyInpurchaseQty = True Then
            If ObjRegistry.BarCodePrefix = Mid(vBarcode, 1, 2) And Len(vBarcode) > 5 Then
               TxtQtyLoose.Text = Round(Val(Mid(vBarcode, 8, 5)) / 1000, 3)
            Else
               TxtQtyLoose.Text = IIf(Len(TxtCode.Text) <= 5 And IsNumeric(TxtCode.Text), 1, IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty))  'IIf(Val(TxtQtyLoose.Text) = 0, 1, TxtQtyLoose.Text)
            End If
         End If
         If IsNull(!PackingName) Then
            vUnitPrice = !PurPrice
            vUnitWSPrice = !WSPrice
            vUnitRetailPrice = !RetailPrice
            TxtMultiplier.Text = ""
            CmbPackName.ListIndex = 0
         Else
            TxtMultiplier.Text = !Multiplier
            If !Multiplier <> 0 Then
               vUnitPrice = !PurPrice / !Multiplier
               vUnitWSPrice = !WSPrice / !Multiplier
               vUnitRetailPrice = !RetailPrice / !Multiplier
            Else
               vUnitPrice = !PurPrice
               vUnitWSPrice = !WSPrice
               vUnitRetailPrice = !RetailPrice
            End If
            CmbPackName.Text = !PackingName
         End If
         TxtDiscPC.Text = IIf(IsNull(!PurDiscPC), "", !PurDiscPC)
         If vUnitPrice = 0 Then
            TxtDiscPer.Text = "0"
         Else
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, IIf(ObjRegistry.IsRoundFigure, 2, 3))
         End If
         If TxtDiscPack.Visible Then
            vDiscPackFlag = IIf(Val(TxtMultiplier.Text) = 0, False, True)
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
                     If ObjRegistry.IsRoundFigure Then
                        TxtDiscPC.Text = Round(!DiscPC, 0)
                        TxtDiscPer.Text = Round(!DiscPer, 2)
                     Else
                        TxtDiscPC.Text = !DiscPC
                        TxtDiscPer.Text = !DiscPer
                     End If
                     TxtSaleTaxPer.Text = IIf(IsNull(!SaleTaxPer), "0", !SaleTaxPer)
                  End If
               End With
            End If
         End If
         
         If ObjRegistry.AutoApplyPartyLastDiscount Then
            If Trim(TxtVenderID.Text) <> "" Then
               vStrSQL = "Select Top 1 * " & vbCrLf _
                     + " from PurchaseHeader h inner join PurchaseBody b on h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate" & vbCrLf _
                     + " left outer join ProductPacking pp on pp.packingid = b.packingid and pp.productid = b.productid" & vbCrLf _
                     + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
                     + " where h.VendorID = " & Val(TxtVenderID.Text) & " and b.ProductID = " & Val(TxtProductID.Text) & vbCrLf _
                     + " Order by h.PurchaseDate Desc, h.PurID desc"
               With CN.Execute(vStrSQL)
                  If .RecordCount > 0 Then
                     If ObjRegistry.IsRoundFigure Then
                        TxtDiscPC.Text = Round(!DiscPC, 0)
                        TxtDiscPer.Text = Round(!DiscPer, 2)
                     Else
                        TxtDiscPC.Text = !DiscPC
                        TxtDiscPer.Text = !DiscPer
                     End If
                     TxtSaleTaxPer.Text = IIf(IsNull(!SaleTaxPer), "0", !SaleTaxPer)
                  End If
               End With
            End If
         End If

            If ObjRegistry.AlertAllocateProduct = True Then
               vStrSQL = "select ListPrice, Multiplier, isnull(vp.DiscPer, isnull(vc.DiscPer,0)) as DiscPer, isnull(vp.DiscPack, isnull(vc.DiscPack,0)) as DiscPack" & vbCrLf _
                     + " from products p" & vbCrLf _
                     + " left outer join ProductPacking pp on pp.packingid = p.PurchasePackingid and pp.productid = p.productid" & vbCrLf _
                     + " left outer join PurchaseProductDisc vp on vp.productid = p.productid" & vbCrLf _
                     + " left outer join PurchaseCompanyDisc vc on vc.companyid = p.companyid" & vbCrLf _
                     + " where p.ProductID = " & Val(TxtProductID.Text)
               
               Dim vMultiplier As Double
               With CN.Execute(vStrSQL)
                  If .RecordCount > 0 Then
                     If Not (!DiscPer = 0 And !DiscPack = 0) Then
                        If !ListPrice <> 0 Then
                           TxtPrice.Text = !ListPrice
                        End If
                     End If
                     If !Multiplier <> 0 Then
                        vUnitPrice = Val(TxtPrice.Text) / !Multiplier
                        vMultiplier = !Multiplier
                     Else
                        vUnitPrice = Val(TxtPrice.Text)
                        vMultiplier = 1
                     End If
                     TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
                     TxtDiscPack.Text = IIf(IsNull(!DiscPack), 0, !DiscPack)
                     If TxtDiscPer.Text <> "0" Then
                        TxtDiscPC.Text = Round((vUnitPrice * Val(TxtDiscPer.Text) / 100), IIf(ObjRegistry.IsRoundFigure, 0, 4))
                        'vDiscPackFlag = False
                     Else
                        TxtDiscPC.Text = Round(TxtDiscPack.Text / IIf(vMultiplier = 0, 1, vMultiplier), IIf(ObjRegistry.IsRoundFigure, 0, 4))
                        TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / IIf(vUnitPrice = 0, 1, vUnitPrice), IIf(ObjRegistry.IsRoundFigure, 2, 5))
                        'vDiscPackFlag = True
                     End If
                  End If
               End With
            End If

         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpPurchaseDate.DateValue + 1 & "',0),0)"
          With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               vQtyLoose = .Fields(0).Value
            Else
               vQtyLoose = 0
            End If
         End With
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         
         With CN.Execute("Select isnull(abbreviation,'') from packings where packingname = '" & CmbPackName.Text & "'")
            If .RecordCount > 0 Then
               LblStock.Caption = LblStock.Caption & " " & .Fields(0).Value
            Else
               LblStock.Caption = LblStock.Caption & " "
            End If
         End With
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
                  
'         If ObjRegistry.NegativeSale = False Then
'            If Val(LblStock.Caption) <= 0 Then
'               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'               FunSelectProduct = False
'               Exit Function
'            End If
'         End If
         If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & ",Null,0,0,0,0,0,0,'" & DtpPurchaseDate.DateValue + 1 & "',0),0)"
            With CN.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
            LblAllStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
            With CN.Execute("Select isnull(abbreviation,'') from packings where packingname = '" & CmbPackName.Text & "'")
               If .RecordCount > 0 Then
                  LblAllStock.Caption = LblAllStock.Caption & " " & .Fields(0).Value
               Else
                  LblAllStock.Caption = LblAllStock.Caption & " "
               End If
            End With
            LblAllStock.Caption = LblAllStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & "Loose"
            LblAllStock.Visible = vShowStock
            LblStock.Visible = Not LblAllStock.Visible
         Else
            LblAllStock.Visible = False
            LblStock.Visible = True
         End If
         
         PopulateDataToHistoryGrid
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

         SubCalculateBody
         'Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         FrmHistory.Visible = False
         FrmProductPrices.Visible = False
         TxtProductID.Text = ""
         TxtCode.Text = ""
         If CmbPackName.ListCount > 0 Then CmbPackName.ListIndex = 0
         TxtProductName.Text = ""
         TxtMultiplier.Text = ""
         TxtPrice.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtTradeOffer1.Text = ""
         TxtTradeOffer2.Text = ""
         TxtExtraSchemePer.Text = ""
         TxtTradeOfferValue.Text = ""
         TxtExtraSchemeValue.Text = ""
         ChkDiscB4TradeOffer.Value = 0
         ChkDiscB4ExtraScheme.Value = 0
         ChkDiscB4SaleTax.Value = 0
         TxtAmount.Text = ""
         LblStock.Visible = False
         LblAllStock.Visible = False
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

Private Sub BtnAddBarCode_Click()
   On Error GoTo ErrorHandler
      If Trim(TxtBarcode.Text) = "" Then Exit Sub
      With CN.Execute("Select code from productbarcodes Where Code = '" & TxtBarcode.Text & "' and ProductID = " & Val(TxtCode.Text))
         If Not .EOF Then Exit Sub
      End With
      CN.Execute ("Insert productbarcodes (ProductID, Code) Values  (" & Val(TxtCode.Text) & ",'" & TxtBarcode.Text & "')")
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnBarCode_Click()
   On Error GoTo ErrorHandler
   If BtnSave.Enabled Then BtnSave_Click
   Dim i As Integer
   If vIsNewRecord = False Then
      vPurchaseID = TxtPurchaseID.Text
      vPurchaseDate = DtpPurchaseDate.DateValue
   End If
   If vColour = True Then
            ssql = " Select ItemCode, b.ProductID, ProductName, p.RetailPrice, isnull(Multiplier,1) * isnull(Qtypack,0)+ isnull(Bonus,0) + QtyLoose as QtyLoose, ColourID, SizeID, GroupID" & vbCrLf _
         + " from PurchaseBody b inner join Products p on b.PRoductID = p.ProductID" & vbCrLf _
         + " where PurID = " & vPurchaseID & " and PurchaseDate = '" & vPurchaseDate & "' order by SerialNo"
   
      With CN.Execute(ssql)
         FrmMultiBarcodesDetail.SubClearFields
         FrmMultiBarcodesDetail.TxtTotQty.Text = "0"
         For i = 1 To .RecordCount
            FrmMultiBarcodesDetail.Grid.Columns("ID").Text = !ItemCode
            FrmMultiBarcodesDetail.Grid.Columns("ProductID").Text = !Productid
            FrmMultiBarcodesDetail.Grid.Columns("GroupID").Text = !GroupID
            FrmMultiBarcodesDetail.Grid.Columns("ColourID").Text = !ColourID
            FrmMultiBarcodesDetail.Grid.Columns("ColourName").Text = CN.Execute("Select ColourName from Colours where ColourID=" & !ColourID).Fields(0).Value
            FrmMultiBarcodesDetail.Grid.Columns("SizeID").Text = !SizeID
            FrmMultiBarcodesDetail.Grid.Columns("SizeName").Text = CN.Execute("Select SizeName from Sizes where SizeID=" & !SizeID).Fields(0).Value
            FrmMultiBarcodesDetail.Grid.Columns("Name").Text = !ProductName
            FrmMultiBarcodesDetail.Grid.Columns("Qty").Value = !QtyLoose
            FrmMultiBarcodesDetail.Grid.Update
            FrmMultiBarcodesDetail.Grid.AddNew
            FrmMultiBarcodesDetail.TxtTotQty.Text = Val(FrmMultiBarcodesDetail.TxtTotQty.Text) + !QtyLoose
            .MoveNext
         Next i
      End With
      'FrmMultiBarcodesDetail.Grid.FirstRow = 0
      FrmMultiBarcodesDetail.Show

   Else
      ssql = " Select b.ProductID, ProductName, Price, isnull(Multiplier,1) * isnull(Qtypack,0)+ isnull(Bonus,0) + QtyLoose as QtyLoose, GroupID" & vbCrLf _
            + " from PurchaseBody b inner join Products p on b.PRoductID = p.ProductID" & vbCrLf _
            + " where PurID = " & vPurchaseID & " and PurchaseDate = '" & vPurchaseDate & "' order by SerialNo"
   '   sSql = "select b.ProductID, Code, ProductName from ProductBarcodes b inner join Products p on p.productid = b.ProductID where len(code) = 11 and code like '110%'"
      
      With CN.Execute(ssql)
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
'            FrmMultiBarcodes.Grid.Row =
            FrmMultiBarcodes.Show
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnChangePrice_Click()
   On Error GoTo ErrorHandler
   Dim vPurID As Integer
   Dim vPurchaseDate As Date
   vPurID = Val(TxtPurchaseID.Text)
   vPurchaseDate = DtpPurchaseDate.DateValue
   
   Dim PurchaseFlag As Boolean
   'PurchaseFlag = False
   'If MsgBox("Do you want to Show Purchase Price of Current Invoice?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbYes Then PurchaseFlag = True
   'If vIsNewRecord = False Then
   '   vPurchaseID = TxtPurchaseID.Text
   'End If
   If Grid.Rows > 1 Then
    FrmChangePrice.vParaSave = True
    FrmChangePrice.vParaProductID = ""
    Grid.Redraw = False
    Grid.MoveFirst
    For vCounter = 1 To Grid.Rows - 1
        FrmChangePrice.vParaProductID = FrmChangePrice.vParaProductID & Val(Grid.Columns("Code").Text) & ","
        Grid.MoveNext
    Next vCounter
    Grid.MoveLast
    Grid.Redraw = True
    FrmChangePrice.vParaProductID = FrmChangePrice.vParaProductID & "''"
   Else
      FrmChangePrice.vParaSave = False
   End If

   If ActiveControl.Name = BtnChangePrice.Name Then If BtnSave.Enabled Then BtnSave_Click

   FrmChangePrice.Grid.Redraw = False
   If FrmChangePrice.Rs.State = adStateOpen Then
      FrmChangePrice.Rs.CancelBatch
      FrmChangePrice.Rs.Close
   End If
  
  If ObjRegistry.ShowWholeSaleMargin = True Then
      vMargin = "round(case when P.WSPrice = 0 then 0 else ((isnull(P.WSPrice,0)/isnull(sp.multiplier,1)) - (P.PurPrice/isnull(pp.multiplier,1))) / ((isnull(P.WSPrice,1)/isnull(sp.multiplier,1)) )*100 end,3) as margin "
  Else
      vMargin = "round(case when P.RetailPrice = 0 then 0 else (isnull(P.RetailPrice,0) - (P.PurPrice/isnull(pp.multiplier,1)))*100/ isnull(P.RetailPrice,1) end,3) as margin"
  End If
  
  
'   ssql = " SELECT distinct p.*, isnull(PackingName,'') as PackingName, isnull(PP.Multiplier,0) as Multiplier,  PurPrice as PurchasePrice, pb.SerialNo, " & vMargin & " from Products p" & vbCrLf _
       + " inner join (Select * from PurchaseBody where PurID = " & vPurID & " and PurchaseDate = '" & vPurchaseDate & "')pb on p.productid = pb.productid" & vbCrLf _
       + " left outer join ProductBarCodes b on p.productid = b.productid " & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = '" & Right("00000" + CStr(Val(TxtProductID.Text)), 5) & "' or b.Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & " Order by Pb.SerialNo"
       
       ssql = " SELECT distinct p.*, isnull(PackingName,'') as PackingName, isnull(PP.Multiplier,0) as Multiplier, PurPrice, round(amount/(isnull(qtypack,0)+qtyloose),2) as PurchasePrice," & vbCrLf _
       + " isnull(pb.DiscPC,0) PbDiscPC, isnull(pb.DiscPer,0) PbDiscPer, isnull(pb.DiscVal,0) PbDiscval, isnull(pb.multiplier,0) PbMultiplier,  Pb.SerialNo, " & vMargin & " from Products p" & vbCrLf _
       + " inner join (Select * from PurchaseBody where PurID = " & vPurID & " and PurchaseDate = '" & vPurchaseDate & "')pb on p.productid = pb.productid" & vbCrLf _
       + " left outer join ProductBarCodes b on p.productid = b.productid " & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text) & " or b.Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & " Order by Pb.SerialNo"
      
      
'      ssql = "Select p.ProductID, ProductName, p.IsRawProduct, p.IsNoCostProduct, p.IsLocked, PurPrice, p.RetailPrice, p.WSPrice, p.DiscPC, p.DiscPer, MinStockLimit, MaxStockLimit from Products p " & vbCrLf _
      inner join (Select * from PurchaseBody where PurID = " & vPurID & " and PurchaseDate = '" & vPurchaseDate & "') b " & vbCrLf _
         + " on p.productid = b.productid "
   FrmChangePrice.vParaSQL = ssql
   vStrSQL = "Select * from Products where ProductID in (select ProductID from PurchaseBody where PurID = " & vPurID & " and PurchaseDate = '" & vPurchaseDate & "')"
   FrmChangePrice.Rs.Open vStrSQL, CN, adOpenDynamic, adLockBatchOptimistic
   FrmChangePrice.Grid.CancelUpdate
   FrmChangePrice.Grid.RemoveAll
   FrmChangePrice.vSuppressUpdateEvent = True
   
   With CN.Execute(ssql)
      Do Until .EOF
         FrmChangePrice.Rs.Find "ProductID = " & Val(!Productid), , adSearchForward, 1
         'If FrmChangePrice.Rs.EOF Then MsgBox "Cannot Locate Record for updation. Please Try again", vbCritical, "Error": Cancel = True: Exit Sub
         'FrmChangePrice.Rs!ProductName = !ProductName
         FrmChangePrice.Rs!PurPrice = !PurPrice
         FrmChangePrice.Rs!RetailPrice = !RetailPrice
         FrmChangePrice.Rs!WSPrice = !WSPrice
         FrmChangePrice.Rs!ListPrice = !ListPrice
         FrmChangePrice.Rs!DiscPC = !DiscPC
         'FrmChangePrice.Rs!StockLimit = !StockLimit
         FrmChangePrice.Rs!SaleTaxPer = !SaleTaxPer
         FrmChangePrice.Rs!PCTCode = !PCTCode
         FrmChangePrice.Rs!is3rdScheduleItem = !is3rdScheduleItem
                  
         FrmChangePrice.Rs.Update
  
         FrmChangePrice.Grid.AddNew
         FrmChangePrice.Grid.Columns("ID").Text = !Productid
         FrmChangePrice.Grid.Columns("Name").Text = !ProductName
         FrmChangePrice.Grid.Columns("Packing").Text = !PackingName
         FrmChangePrice.Grid.Columns("Multiplier").Value = !Multiplier
         FrmChangePrice.Grid.Columns("PurPrice").Value = IIf(ObjRegistry.ShowDiscPurPrice = True, !PurchasePrice, !PurPrice)
         'FrmChangePrice.Grid.Columns("PurPrice").Value = !PurPrice - IIf(ObjRegistry.isShowListPrice = True, (!PbDiscPC * IIf(!multiplier = 0, 1, !multiplier)), 0)
         FrmChangePrice.Grid.Columns("RetailPrice").Value = !RetailPrice
         FrmChangePrice.Grid.Columns("WSPrice").Value = !WSPrice
         FrmChangePrice.Grid.Columns("ListPrice").Value = !ListPrice
         FrmChangePrice.Grid.Columns("Margin").Value = !Margin
         FrmChangePrice.Grid.Columns("DiscPC").Value = !DiscPC
         FrmChangePrice.Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
         FrmChangePrice.Grid.Columns("DiscVal").Value = IIf(ObjRegistry.ShowWholeSaleMargin = True, !WSPrice, !RetailPrice) * IIf(IsNull(!DiscPer), 0, !DiscPer) / 100
         FrmChangePrice.Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), 0, !SaleTaxPer)
         FrmChangePrice.Grid.Columns("PCTCode").Value = IIf(IsNull(!PCTCode), "", !PCTCode)
         FrmChangePrice.Grid.Columns("3rdSchedule").Value = Abs(IIf(IsNull(!is3rdScheduleItem), 0, !is3rdScheduleItem))
         FrmChangePrice.Grid.Columns("PurDiscPC").Value = !PbDiscPC
         FrmChangePrice.Grid.Columns("PurDiscPer").Value = IIf(IsNull(!PbDiscPer), 0, !PbDiscPer)
         FrmChangePrice.Grid.Columns("MinStockLimit").Value = IIf(IsNull(!MinStockLimit), 0, !MinStockLimit)
         FrmChangePrice.Grid.Columns("MaxStockLimit").Value = IIf(IsNull(!MinStockLimit), 0, !MinStockLimit)
         FrmChangePrice.Grid.Columns("Lock").Value = Abs(!IsLocked)
         FrmChangePrice.Grid.Columns("NoCost").Value = Abs(!IsNoCostProduct)
         FrmChangePrice.Grid.Columns("Raw").Value = Abs(!IsRawProduct)
         FrmChangePrice.Grid.Update
         .MoveNext
      Loop
   End With
   FrmChangePrice.vSuppressUpdateEvent = False
   FrmChangePrice.Grid.Redraw = True
   FrmChangePrice.Grid.MoveFirst
   FrmChangePrice.Grid.FirstRow = 0
   FrmChangePrice.Show , Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   If MsgBox("Are you sure to Clear the Data?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Exit Sub
   
    '''''''''''''''''' ActivityLogBin For Clear Action
'      Call DeleteTempActivityLogBin(vRandomID)
      vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.Rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From purchasebody where PurID = " & Val(TxtPurchaseID.Text) & " and PurchaseDate='" & DtpPurchaseDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmPurchaseInvoice, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtPurchaseID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPurchaseDate.Date), "Cleared Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmPurchaseInvoice, eClearSavedRecord, TxtPurchaseID.Text, DtpPurchaseDate.DateValue, vGridRows & " Product/s Cleared")
      Grid.Redraw = True
  ''''''''''''''''''
     FormStatus = NewMode
'   cn.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & Val(TxtPurchaseID.Text) & ",'" & DtpPurchaseDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniPurchaseInvoice", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
      vStrSQL = "select * from PurchaseHeader where Tag is not null And SID=" & Val(TxtSID.Text) & " and Purchasedate='" & DtpPurchaseDate.DateValue & "'"
      With CN.Execute(vStrSQL)
          If Not .EOF Then
              MsgBox "Import/Export Record Cannot be Deleted", vbInformation, Me.Caption
              Exit Sub
          End If
      End With
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   ''''''''''''''''''''''' Check Low Stock  '''''''''''''''''''''''''''''''''
   If ObjRegistry.NegativeSale = False Then
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 1 To Grid.Rows
         If Trim(Grid.Columns("Productid").Text) <> "" Then
            vStrSQL = "select isnull(dbo.FunStock(" & Val(Grid.Columns("Productid").Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & Date + 1 & "',0),0)"
                With CN.Execute(vStrSQL)
                  If .RecordCount > 0 Then
                     vQtyLoose = .Fields(0).Value
                  Else
                     vQtyLoose = 0
                  End If
               End With
            If (Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value)) + Val(Grid.Columns("QtyLoose").Value) > Val(vQtyLoose) Then
               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
               Exit Sub
            End If
         End If
         Grid.MoveNext
      Next vCounter
      Grid.Redraw = True
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   CN.BeginTrans
   
   Call BinData
   Call ActivityLogBin("", eFrmPurchaseInvoice, eDelete, TxtPurchaseID.Text, DtpPurchaseDate.DateValue, Grid.Rows - 1 & " Product/s Deleted Amount: " & Val(TxtNetAmount.Text))
'   vMaxBinID = FunGetMaxBinID
'   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_PurchaseHeader Select " & vMaxBinID & ",'" & Date & "',* from PurchaseHeader Where PurID = " & TxtPurchaseID.Text & " And PurchaseDate ='" & DtpPurchaseDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_PurchaseBody Select " & vMaxBinID & ",'" & Date & "', * from PurchaseBody Where PurID = " & TxtPurchaseID.Text & " And PurchaseDate ='" & DtpPurchaseDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Serial''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_PurchaseBodySerial Select " & vMaxBinID & ",'" & Date & "', * from PurchaseBodySerial Where PurID = " & TxtPurchaseID.Text & " And PurchaseDate ='" & DtpPurchaseDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin ProductOffer''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_PurchaseBodyOffer Select " & vMaxBinID & ",'" & Date & "', * from PurchaseBodyOffer Where PurID = " & TxtPurchaseID.Text & " And PurchaseDate ='" & DtpPurchaseDate.DateValue & "'")
'
'  '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  ''''''''''''''''''''''''''Delete Product Offer'''''''''''''''''''''
 
   Dim vBm As Variant
   GridOffer.Redraw = False
   vBm = GridOffer.Bookmark
   GridOffer.MoveFirst
   
   For vCounter = 0 To GridOffer.Rows - 1
      If Trim(GridOffer.Columns("Productid").CellValue(GridOffer.GetBookmark(i))) <> "" Then
         CN.Execute "Delete from PurchaseBodyOffer where PurID = " & Val(TxtPurchaseID.Text) & " And PurchaseDate ='" & DtpPurchaseDate.DateValue & "' and productid = " & Val(GridOffer.Columns("Productid").CellValue(GridOffer.GetBookmark(i)))
      End If
   Next vCounter
   
   GridOffer.Bookmark = vBm
   
   GridOffer.RemoveAll
   GridOffer.Redraw = True

  ''''''''''''''''''''''''''Delete Serials'''''''''''''''''''''
    If RsBodySerial.RecordCount > 0 Then
        RsBodySerial.MoveFirst
        For vCounter = 1 To RsBodySerial.RecordCount
            CN.Execute "Delete from PurchaseBodySerial where PurID = " & Val(TxtPurchaseID.Text) & " And PurchaseDate ='" & DtpPurchaseDate.DateValue & "' and productid = " & Val(RsBodySerial!Productid) & " and Serial ='" & RsBodySerial!serial & "'"
            RsBodySerial.MoveNext
        Next vCounter
    End If
      
  ''''''''''''''''''''''''''Delete Purchase Body'''''''''''''''''''''
   Grid.Redraw = False
   Grid.MoveFirst
   Call ActivityLog("Purchase Invoice", eDelete, Val(TxtPurchaseID.Text), DtpPurchaseDate.DateValue)
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from PurchaseBody where SID = " & Val(TxtSID.Text) & " and Purchasedate='" & DtpPurchaseDate.DateValue & "' and productid = " & Val(Grid.Columns("ProductID").Text) & " and BatchNo " & IIf(Trim(Grid.Columns("BatchNo").Text) = "", " is null", " = '" & Trim(Grid.Columns("BatchNo").Text) & "'") & " and Price = " & Val(Grid.Columns("Price").Text) & IIf(vColour = True, " and ColourID = " & Val(Grid.Columns("ColourID").Text) & " and SizeID = " & Val(Grid.Columns("SizeID").Text), "")
         '''
         ssql = "if exists (Select top 1 Price from PurchaseBody where productid = '" & Grid.Columns("ProductID").Text & "') Begin Update products Set IsSync = 0, PurPrice = (Select top 1 Price from PurchaseBody where productid = '" & Grid.Columns("ProductID").Text & "' Order by purchasedate desc) , PurDiscPC = (Select top 1 DiscPC from PurchaseBody where productid = '" & Grid.Columns("ProductID").Text & "' Order by purchasedate desc) Where productid = " & Val(Grid.Columns("ProductID").Text) & " End"
         CN.Execute (ssql)
'          CN.Execute ("Insert Into Bin_PurchaseBody Select " & FunGetMaxBinID & ", * from PurchaseBody Where PurID = " & TxtPurchaseID.Text & " And PurchaseDate ='" & DtpPurchaseDate.DateValue & "' and productid ='" & Grid.Columns("Productid").Text & "'")    RsBody.Filter = "ProductID = '" & TxtProductID.Text & "' and BatchNo = " & IIf(Trim(TxtBatchNo.Text) = "", "null", "'" & Trim(TxtBatchNo.Text) & "'") & " and Price = " & Val(TxtPrice.Text)
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
    '''''''''''''''''''''''''''''''''''''''Delete Expense'''''''''''''''''''''''''''''''''''''''
   CN.Execute "Delete from PurchaseExpense where PurID = " & Val(TxtPurchaseID.Text) & " and PurchaseDate='" & DtpPurchaseDate.DateValue & "'"
   
   CN.Execute "Delete from PurchaseHeader where SID = " & Val(TxtPurchaseID.Text) & " and Purchasedate='" & DtpPurchaseDate.DateValue & "'"
   
   CN.Execute ("Update PurchaseOrderHeader set isPurchase = 0 Where OrderID = " & Val(TxtOrderID.Text) & "And Orderdate ='" & DtpOrderDate.DateValue & "'")
   
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnDelete Then
   vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = ObjRegistry.PrefixPhoneNo + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = ObjUserSecurity.UserName & " " & FrmPurchaseInvoice.Caption & " Deleted ID:" & TxtPurchaseID.Text & vbCrLf & " Date:" & Format(DtpPurchaseDate.DateValue, "dd-MMM-yyyy") & " Time: " & Time & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt" & TxtNetAmount.Text
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & "','')"
               CN.Execute ssql
            End If
         Next
   End If
   
   CN.CommitTrans

   FormStatus = NewMode
   GridOffer.ZOrder 1
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchPurchase.ParaInPurchasedate = DtpPurchaseDate.DateValue
   SchPurchase.Show vbModal
   If SchPurchase.ParaOutPurchaseID <> "" Then
      TxtSID.Text = SchPurchase.ParaOutSID
      TxtPurchaseID.Text = SchPurchase.ParaOutPurchaseID
      DtpPurchaseDate.DateValue = SchPurchase.ParaOutPurchaseDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetPurchase
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
      Grid.AllowAddNew = True
      TxtVenderID.Text = !PartyID
      Call FunSelectVender(1, True)
      
      While Not .EOF
         TxtCode.Text = !Productid
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !Price
         TxtQtyPack.Text = IIf(Val(!QtyPack) = 0, "", Val(!QtyPack))
         TxtMultiplier.Text = IIf(Val(!QtyPack) = 0, "", Val(!Multiplier))
         TxtQtyLoose.Text = !QtyLoose
         TxtRetailPrice.Text = !RetailPrice
         vIsWSDiscb4ST = 0
         vIsWSSaleTax = 0
         vIsRetailSaleTax = 0
         TxtSaleTaxPer.Text = ""
         ChkDiscB4TradeOffer.Value = 0
         ChkDiscB4ExtraScheme.Value = 0
         ChkDiscB4SaleTax.Value = 0
         TxtTradeOffer1.Text = 0
         TxtTradeOffer2.Text = 0
         TxtExtraSchemePer.Text = 0
         TxtSaleTaxPer.Text = 0
         TxtDiscPer.Text = IIf(Val(!DiscPer) = 0, "", Val(!DiscPer))
         
         CmbPackName.Clear
         With CN.Execute("Select * from Packings")
            CmbPackName.AddItem ""
            While Not .EOF
            CmbPackName.AddItem !PackingName
            CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
            .MoveNext
            Wend
         .Close
        End With
            
         If !PackingName = "" Then
            vUnitPrice = Val(TxtPrice.Text)
            vUnitRetailPrice = vUnitPrice
            TxtMultiplier.Text = ""
            CmbPackName.ListIndex = 0
         Else
            TxtMultiplier.Text = IIf(Val(!QtyPack) = 0, "", Val(!Multiplier))
            If Val(!Multiplier) <> 0 Then
               vUnitPrice = Val(TxtPrice.Text) / Val(!Multiplier)
               'If !QtyPack = 0 Then TxtPrice.Text = vUnitPrice
            Else
               vUnitPrice = Val(TxtPrice.Text)
            End If
            vUnitRetailPrice = vUnitPrice
            If !QtyPack = 0 Then
               CmbPackName.ListIndex = 0
            Else
               CmbPackName.Text = !PackingName
            End If
         End If
           
'         If vUnitPrice = 0 Then Exit Sub
         TxtDiscPC.Text = Round((vUnitPrice * Val(TxtDiscPer.Text) / 100), 4)
         If Val(TxtDiscPC.Text) = 0 Then TxtDiscPC.Text = ""
         Call SubCalculateBody
            
         GetDataFromTexBoxesToGrid
         
         .MoveNext
      Wend
      .Close
   End With
   Grid.AllowAddNew = False
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProductRange_Click()
   On Error GoTo ErrorHandler
   FrmProductRangeGrid.ParaInPartyID = TxtVenderID.Text
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

Private Sub BtnPurchaseOrder_Click()
   On Error GoTo ErrorHandler
   SchPurchaseOrder.ParaInOrderDate = DtpOrderDate.DateValue
   SchPurchaseOrder.Show vbModal
   If SchPurchaseOrder.ParaOutOrderID <> "" Then
      TxtOrderID.Text = SchPurchaseOrder.ParaOutOrderID
      'Dim a
      'a = Split(SchSale.ParaOutBillDate, "/")
      DtpOrderDate.DateValue = SchPurchaseOrder.ParaOutOrderDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
'      CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetPurchaseOrder
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_RowLoaded(ByVal Bookmark As Variant)
   With Grid
'      If Val(.Columns("ExpiryTime").Value) = 0 Then
'         .Columns("ProductName").CellStyleSet ""
'      ElseIf Val(.Columns("ExpiryTime").Value) <= 90 And Val(.Columns("ExpiryTime").Value) > 30 Then
'         .Columns("ProductName").CellStyleSet "Green"
'      ElseIf Val(.Columns("ExpiryTime").Value) <= 30 And Val(.Columns("ExpiryTime").Value) > 0 Then
'         .Columns("ProductName").CellStyleSet "Orange"
'      ElseIf Val(.Columns("ExpiryTime").Value) < 0 Then
'         .Columns("ProductName").CellStyleSet "Red"
'      End If
      '''''' Get ExpiryColor
      If Val(.Columns("ExpiryTime").Value) < 0 Then
         .Columns("ProductName").CellStyleSet "Red"
      Else
         ssql = "Select * from ExpiryDayColor Where " & Val(.Columns("ExpiryTime").Value) & " >= DayFrom and " & Val(.Columns("ExpiryTime").Value) & " <= DayTo"
         With CN.Execute(ssql)
            If .RecordCount <> 0 Then vExpiryColor = !ExpiryColor Else vExpiryColor = ""
         End With
      .Columns("ProductName").CellStyleSet vExpiryColor
      End If
   End With
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

Private Sub TxtAmount_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtAmount.Name Then Exit Sub
   If (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) = 0 Then Exit Sub
   TxtPrice.Text = ((Val(TxtAmount.Text) + Val(TxtDiscVal.Text)) / (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) * IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))
   If Val(TxtMultiplier.Text) <> 0 Then
      vUnitPrice = Val(TxtPrice.Text) / Val(TxtMultiplier.Text)
   Else
      vUnitPrice = Val(TxtPrice.Text)
   End If
   'Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub TxtBarCode_Validate(Cancel As Boolean)
'   If ActiveControl.Name <> TxtBarCode.Name Then Exit Sub
'   Call BtnAddBarCode_Click
End Sub

Private Sub TxtBillNo_LostFocus()
   On Error GoTo ErrorHandler
   With CN.Execute("select * from PurchaseHeader where BillNo = '" & TxtBillNo.Text & "' and VendorID = " & Val(TxtVenderID.Text))
      If .RecordCount > 0 Then
         MsgBox "This Bill No. is alrady exist.", vbExclamation, Me.Caption
      End If
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDiscPer2_Change()
   If ActiveControl.Name <> TxtDiscPer2.Name Then Exit Sub
   TxtDiscVal2.Text = SelfRound((Val(TxtAmount.Text) * Val(TxtDiscPer2.Text) / 100))
   SubCalculateBody
End Sub

Private Sub TxtDiscVal2_Change()
  On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDiscVal2.Name Then Exit Sub
   TxtDiscPer2.Text = Round((Val(TxtDiscVal2.Text) * 100) / IIf(Val(TxtAmount.Text) = 0, 1, Val(TxtAmount.Text)), 2)
   SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtExtraTaxPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtExtraTaxPer.Name Then Exit Sub
   TxtExtraTaxVal.Text = SelfRound((Val(TxtSumDiscAmount.Text) * Val(TxtExtraTaxPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtExtraTaxVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtExtraTaxVal.Name Then Exit Sub
   TxtExtraTaxPer.Text = Round((Val(TxtExtraTaxVal.Text) * 100) / IIf(Val(TxtSumDiscAmount.Text) = 0, 1, Val(TxtSumDiscAmount.Text)), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage

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
      TxtVenderID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.PurchaseRePrint = False Then
   vStrSQL = "Select Purid from PurchaseHeader  where PurID = " & Val(TxtPurchaseID.Text) & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
    With CN.Execute(vStrSQL)
        If .EOF = False Then
            MsgBox "You are not Allowed to Reprint", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
   End If
   
   vStrSQL = " Select h.PurID, h.PurchaseDate, OrderID, OrderDate, EntryDate, Pr.PartyName + ' - ' + cast(H.VendorID as varchar(10)) as Vend_Name_ID, isnull( pr.Phone1  + ', ','') + isnull( pr.Phone2 + ', ','')  + isnull( pr.mobile + ', ','') +  isnull( pr.mobile2 + ', ','') as Moblie, " & vbCrLf _
      + " StoreName, h.BillNo, h.Description, h.Remarks, isnull(PreviousAmount,0) PreviousAmount, isnull(PaidAmount,0) PaidAmount, Totalamount, " & vbCrLf _
      + " b.code, b.productID, dbo.FunPurchaseBodySerial(b.purID,b.purchaseDate,b.ProductId) Serial, AdvTaxVal, AdvTaxPer, ExtraTaxVal, ExtraTaxPer" & vbCrLf _
      + " ,dbo.FunPurchaseBodyOffer(b.purID,b.purchaseDate,b.ProductId) ProductOffer, ProductName, batchno, b.expirydate, C.CompanyID, Companyname, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + QtyLoose as Qty, QtyPack, Multiplier, GrossQty, isnull(b.GrossUnit,0)GrossUnit, QtyLoose, " & vbCrLf _
      + " Bonus,b.DiscPc, b.DiscPer, DiscVal, Offer, b.SaleTaxPer, SaleTaxval, SC, h.BillDisc, isnull(OtherCharges,0) as OtherCharges, " & vbCrLf _
      + " (Amount-(Amount* (isnull(BillDiscPer,0)/100)) + (Amount* isnull(OtherCharges,0) /TotalAmount)) /((isnull(QtyPack,0)*isnull(Multiplier,0))+QtyLoose+isnull(Bonus,0)+(Amount* isnull(TotalExpense,0) /TotalAmount))  as cost, " & vbCrLf _
      + " Amount, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + QtyLoose as Qty, b.price/isnull(multiplier,1) as Qtyprice," & vbCrLf _
      + " dbo.FunLastPurPrice(h.PurID,h.PurchaseDate,b.ProductID) as LastPrice, isnull(PK.PackingName,'PC') as PackingName, z.ZoneID, ZoneName, Sec.SectorName, UserName, isnull(pr.Address,'') + ' - ' + isnull(pr.City,'') as Address," & vbCrLf _
      + " OldPrice, b.price, p.RetailPrice, biltyNo, p.ItemCode, right('00'+ cast(b.ColourID as varchar(2)),2) as ColourID, right('00'+ cast(b.SizeID as varchar(2)),2) as SizeID, ColourName, SizeName," & vbCrLf _
      + " p.RetailPrice - (Amount-(Amount* (isnull(BillDiscPer,0)/100)) + (Amount* isnull(OtherCharges,0) /TotalAmount)) /((isnull(QtyPack,0)*isnull(Multiplier,0))+QtyLoose+isnull(Bonus,0))  as Profit" & vbCrLf _
      + " from purchasebody b inner join products p on b.productid = p.productid" & vbCrLf _
      + " inner join PurchaseHeader h on h.purid = b.purid and h.purchaseDate = b.PurchaseDate" & vbCrLf _
      + " left outer join companies C on C.companyid = p.companyid" & vbCrLf _
      + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
      + " inner join parties pr on pr.partyid = h.VendorID" & vbCrLf _
      + " left outer join Packings PK on Pk.PackingID = B.PackingID" & vbCrLf _
      + " left outer join sectors sec on sec.sectorid = pr.sectorid" & vbCrLf _
      + " left outer join zones z on z.zoneid = sec.zoneid" & vbCrLf _
      + " left outer join users u on u.userno = h.userno" & vbCrLf _
      + " Left outer join Colours Col on Col.Colourid = b.ColourID" & vbCrLf _
      + " Left Outer join Sizes Sz on Sz.SizeID = b.SizeID " & vbCrLf _
      + " where h.SID = " & Val(TxtSID.Text) & " and h.PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
      
    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
    
    If cmbPrintType.Text = "Half Page" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseInvoiceHalf1.rpt")
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.x
      RptReportViewer.Report.RightMargin = 225
   ElseIf cmbPrintType.Text = "Thermal" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseInvoiceAurora.rpt")
      RptReportViewer.Report.TopMargin = 0
      RptReportViewer.Report.LeftMargin = 0
      RptReportViewer.Report.RightMargin = 0
   Else
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseInvoice.rpt")
   
   End If
  
  

   
'   If ObjRegistry.LaserPrintofSaleInvoice = True Then
'      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseInvoice.rpt")
'      RptReportViewer.Report.PaperSize = crPaperA4
'      RptReportViewer.Report.PaperOrientation = crLandscape
'      RptReportViewer.Report.TopMargin = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))
'      RptReportViewer.Report.LeftMargin = IIf(IsNull(ObjRegistry.X), 0, Val(ObjRegistry.X))
'      RptReportViewer.Report.RightMargin = 225
'   ElseIf InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
'      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseInvoice.rpt")
'   Else
'      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseInvoice.rpt")
'      RptReportViewer.Report.TopMargin = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))
'      RptReportViewer.Report.LeftMargin = IIf(IsNull(ObjRegistry.X), 0, Val(ObjRegistry.X))
'      RptReportViewer.Report.RightMargin = 0
'   End If

    
    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
'   RptReportViewer.Report.ParameterFields(6).AddCurrentValue CStr(ObjRegistry.Statement)
'   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port

   'RptReportViewer.Report.ParameterFields(4).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
   'RptReportViewer.Report.PrintOut False
   
   
   vPrinter = Split(CmbPrinters.Text, ",")
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   
'    RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   If ObjRegistry.PreviewSaleInoice = True Or ChkIsPreview.Value = 1 Then
      If ChkIsPrint.Value = 1 Then
         RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
      End If
       RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
   End If
   CN.Execute "update PurchaseHeader set isPrinted = 1 where isnull(isPrinted,0) = 0 And PurID = " & Val(TxtPurchaseID.Text) & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
   CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
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
  If TxtCode.Text <> "" Then
      MsgBox "Code is not clear", vbCritical, "Error"
    Exit Sub
  End If
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniPurchaseInvoice", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   
   ''''''''''''' Check Terms ''''''''''''''
   
   
   
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
  
  If ObjRegistry.TermAllowZero = False And Val(TxtTerms.Text) = 0 Then
      MsgBox "Please Enter Terms", vbInformation, Me.Caption
      If TxtTerms.Visible = True Then TxtTerms.SetFocus
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
'  Header Validation
   If Trim(TxtVenderID.Text) = "" Then
      MsgBox "Enter Vender ID.", vbExclamation, Me.Caption
      TxtVenderID.SetFocus
      Exit Sub
   Else
'      With cn.Execute("SELECT isnull(dbo.DefaultValue('Counter Purchase'),0)")
'         If .RecordCount > 0 Then
'
'         End If
'      End With
   End If
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      If TxtStoreID.Visible And TxtStoreID.Enabled Then TxtStoreID.SetFocus
      Exit Sub
   End If
   
   Dim vBm As Variant
   Dim i As Integer
   Grid.Redraw = False
   vBm = Grid.Bookmark
   Dim vTotalAmount
   vTotalAmount = 0
   Grid.MoveFirst
   For i = 0 To Grid.Rows - 1
      'TxtGrossAmount.Text = Val(TxtGrossAmount.Text) + Val(Grid.Columns("Amount").CellValue(Grid.GetBookmark(i)))
      vTotalAmount = vTotalAmount + Val(Grid.Columns("Amount").CellValue(Grid.GetBookmark(i)))
   Next i
   TxtGrossAmount.Text = vTotalAmount
   Grid.Bookmark = vBm
   Grid.Redraw = True

   With CN.Execute("select dbo.DefaultValue('Counter Purchase')")
      If TxtVenderID.Text = .Fields(0).Value Then
         TxtPaidAmount.Text = TxtNetAmount.Text
      End If
   End With
  
   Call SubCalculateFooter
   
'   If ObjRegistry.SerialCompulsoryinInvoice And RsBodySerial.RecordCount <> Grid.Rows - 1 Then
'      MsgBox "Data can not be saved because Serial No. is not entered", vbInformation, Me.Caption
'      Exit Sub
'   End If
   
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
   If ObjUserSecurity.IsEditClosingInvoice = False Then
      vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where ToUserNo = " & vUser & " and Entrydate <='" & Date & "'"
      With CN.Execute(vStrSQL)
          If .Fields(0).Value >= DtpPurchaseDate.DateValue Then
              MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
              Exit Sub
          End If
      End With
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
       If DtpPurchaseDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
    
    '''''''''''''''''''''''Check Import / Export'''''''''''''''''''''''''''''''''
    If ObjRegistry.ShowMultiBranches = True Then
      vStrSQL = "select * from PurchaseHeader where Tag is not null And PurID=" & Val(TxtPurchaseID.Text) & " and Purchasedate='" & DtpPurchaseDate.DateValue & "'"
      With CN.Execute(vStrSQL)
          If Not .EOF Then
              MsgBox "Import/Export Record Cannot be Update", vbInformation, Me.Caption
              Exit Sub
          End If
      End With
   End If
   
  RsBody.Filter = 0
  If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Purchase", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
  End If
  
   If vIsNewRecord = False Then
'      Call ActivityLog("Purchase Invoice", eEdit, TxtPurchaseID.Text, DtpPurchaseDate.DateValue)
   End If
  
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
   
  
    ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'Purchase Invoice' and LocalComputerName = '" & LocalComputerName & "'"
   If CN.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview, IsPrint ) Values ('" & LocalComputerName & "', 'Purchase Invoice','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & "," & ChkIsPrint.Value & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & ", IsPrint = " & ChkIsPrint.Value & " Where FormType = 'Purchase Invoice' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   CN.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   
   CN.BeginTrans
   
   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmPurchaseInvoice, eEdit, TxtPurchaseID.Text, DtpPurchaseDate.DateValue, "Amount: " & Val(TxtNetAmount.Text))
   
'
   If ObjRegistry.ShowDispatchDate = True Then CN.Execute ("Update sysindexs Set value = '" & Val(TxtTerms.Text) & "'Where Registrykey = 'PurchasePromiseTerms'")
   
   If vIsNewRecord = True Then
      If CN.Execute("Select * from PurchaseHeader where PurID = " & Val(TxtPurchaseID.Text) & " and PurchaseDate='" & DtpPurchaseDate.DateValue & "'").RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtPurchaseID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
   
   ''''''''''''''''' Following Code used for Debug When Make a new Purid in PurcahseBody and New SID in PurchaseHeader
   If Val(TxtPurchaseID.Text) = 0 Then
      TxtPurchaseID.Text = FunGetMaxID
      If vIsNewRecord = False Then
         MsgBox "Please take a screen shot and send to SoftInn.PurchaseHeader  PurID = 0 and vIsNewRecord = False", vbExclamation, "Alert"
         Unload Me
      End If
   End If
   
   ''''''''''''''''' Following Code used for Debug When Make a new Purid in PurcahseBody and New SID in PurchaseHeader
   If vIsNewRecord = False Then
      ssql = "select PurID from PurchaseHeader where SID=" & Val(TxtSID.Text)
      If CN.Execute(ssql).Fields(0).Value <> Val(TxtPurchaseID.Text) Then
         MsgBox "Please take a screen shot and send to SoftInn. PurchaseHeader PurID = " & Val(TxtPurchaseID.Text) & " and SID = " & Val(TxtSID.Text) & " and vIsNewRecord = False", vbExclamation, "Alert"
         Unload Me
      End If
   End If
   '''''''''''''''''''''''''''''''''''''''''
   
   vPurchaseID = TxtPurchaseID.Text
   vPurchaseDate = DtpPurchaseDate.DateValue
   
'   Call UserActivities
   
   ssql = "select * from PurchaseHeader where PurID=" & Val(TxtPurchaseID.Text) & " and Purchasedate='" & DtpPurchaseDate.DateValue & "'"
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, CN, adOpenDynamic, adLockPessimistic
      If .BOF Then
         .AddNew
         !PurID = Val(TxtPurchaseID.Text)
         !PurchaseDate = DtpPurchaseDate.DateValue
         !OrderID = IIf(Val(TxtOrderID.Text) = 0, Null, TxtOrderID.Text)
         !OrderDate = DtpOrderDate.DateValue
         !UserNo = vUser
      End If
      !EntryDate = DtpEntryDate.DateValue
      !PromiseDate = IIf(DtpPromiseDate.DateValue = Empty, Null, DtpPromiseDate.DateValue)
      !vendorID = TxtVenderID.Text
      !StoreID = TxtStoreID.Text
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !BillNo = IIf(TxtBillNo.Text = "", Null, TxtBillNo.Text)
      !BiltyNo = IIf(TxtBiltyNo.Text = "", Null, TxtBiltyNo.Text)
      !VehicleNo = IIf(TxtVehicleNo.Text = "", Null, TxtVehicleNo.Text)
      !TotalAmount = Round(Val(TxtGrossAmount.Text))
      !SumDiscAmount = Round(Val(TxtSumDiscAmount.Text))
      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
      
      !AdvTaxVal = IIf(TxtAdvTaxVal.Text = "", Null, Val(TxtAdvTaxVal.Text))
      !AdvTaxPer = IIf(TxtAdvTaxPer.Text = "", Null, Val(TxtAdvTaxPer.Text))
      
      !ExtraTaxVal = IIf(TxtExtraTaxVal.Text = "", Null, Val(TxtExtraTaxVal.Text))
      !ExtraTaxPer = IIf(TxtExtraTaxPer.Text = "", Null, Val(TxtExtraTaxPer.Text))
      
      !OtherCharges = IIf(Val(TxtOtherCharges.Text) = 0, Null, Val(TxtOtherCharges.Text))
      !Freight = IIf(Val(TxtFreight.Text) = 0, Null, Val(TxtFreight.Text))
      !IsVenderFreight = IIf(OptVender.Value = True, 1, 0)
      !IsOurFreight = IIf(OptMe.Value = True, 1, 0)
      !IsExpense = IIf(OptExpense.Value = True, 1, 0)
      !TotalExpense = IIf(Val(TxtTotalExpense.Text) = 0, Null, Val(TxtTotalExpense.Text))
      !PAIDAMOUNT = IIf(TxtPaidAmount.Text = "", Null, Val(TxtPaidAmount.Text))
      !Description = IIf(TxtDescription.Text = "", Null, TxtDescription.Text)
      !Remarks = IIf(TxtRemarks.Text = "", Null, TxtRemarks.Text)
      !PreviousAmount = IIf(lblPayable.Caption = "Previous Receivable", Val(TxtPreviousPayable.Text), Val(TxtPreviousPayable.Text) * -1)
'      !UserNo = vUser
      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
      .Update
      .Close
      If vIsNewRecord = True Then TxtSID.Text = CN.Execute("select @@identity").Fields(0).Value
   End With
   
   Dim PurMaxDate As String
   
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !SID = Val(TxtSID.Text)
         !PurID = Val(TxtPurchaseID.Text)
         !PurchaseDate = DtpPurchaseDate.DateValue
         PurMaxDate = CN.Execute("SELECT dbo.FunMaxPurDate('" & RsBody!Productid & "')").Fields(0).Value
         If DtpPurchaseDate.DateValue >= CDate(PurMaxDate) Then
            ssql = "update Products set IsSync = 0, PurPrice = " & RsBody!Price & ", PurDiscPC = " & RsBody!DiscPC & IIf(ObjRegistry.ShowChangeRetailInPurchaseInvoice = True, ", RetailPrice = " & RsBody!RetailPrice, "") & IIf(ObjRegistry.isShowListPrice = True, IIf(RsBody!DiscPC <> 0, ", ListPrice = " & RsBody!Price, ""), "") & IIf(IsNull(RsBody!PackingID), "", " , PurchasePackingID = " & RsBody!PackingID) & " Where ProductID = " & Val(RsBody!Productid)
            CN.Execute ssql
            If (Not IsNull(RsBody!PackingID)) And (Not IsNull(RsBody!Multiplier)) And (RsBody!Multiplier <> 0) Then
               If CN.Execute("select * from ProductPacking Where ProductID = " & Val(RsBody!Productid) & " and PackingID = " & RsBody!PackingID).RecordCount = 0 Then
                  ssql = "INSERT INTO ProductPacking(PackingID,Multiplier,ProductID) VALUES ('" & RsBody!PackingID & "','" & RsBody!Multiplier & "'," & Val(RsBody!Productid) & ")"
                  CN.Execute ssql
               Else
                  ssql = "update ProductPacking set IsSync = 0, Multiplier = " & IIf(IsNull(RsBody!Multiplier), 0, RsBody!Multiplier) & " Where ProductID = " & Val(RsBody!Productid) & " and PackingID = " & RsBody!PackingID
                  CN.Execute ssql
               End If
            End If
         End If
        .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   
   If RsBodySerial.RecordCount > 0 Then
     With RsBodySerial
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !PurID = Val(TxtPurchaseID.Text)
         !PurchaseDate = DtpPurchaseDate.DateValue
         .MoveNext
      Next vCounter
      .UpdateBatch
     End With
   End If
   
   With RsProductOffer
      .Filter = 0
      If Not .EOF Then
        .MoveFirst
        For vCounter = 1 To .RecordCount
         !PurID = Val(TxtPurchaseID.Text)
         !PurchaseDate = DtpPurchaseDate.DateValue
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
         !PurID = Val(TxtPurchaseID.Text)
         !PurchaseDate = DtpPurchaseDate.DateValue
         .MoveNext
        Next vCounter
      End If
      .UpdateBatch
   End With
   
   ssql = " select sob.ProductID, ProductName, QtyPack - isnull(UPack,0) as RQtyPack, QtyLoose - isnull(UQty,0) as RQty, Bonus - isnull(UBonus,0) as RBonus, sob.*" & vbCrLf _
      + " from (select OrderID, OrderDate, ProductID, Sum(QtyLoose) as UQty, Sum(QtyPack) as UPack, Sum(Bonus) as UBonus from PurchaseBody b inner join PurchaseHeader h on h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate Group By OrderID, OrderDate, ProductID) b " & vbCrLf _
      + " right outer join PurchaseOrderbody sob on sob.OrderID = b.orderid and sob.OrderDate = b.orderdate and b.ProductID = sob.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = sob.productid" & vbCrLf _
      + " where sob.OrderID = " & Val(TxtOrderID.Text) & " and sob.OrderDate = '" & DtpOrderDate.DateValue & "' and (QtyPack - isnull(UPack,0) <> 0 or Qtyloose - isnull(UQty,0) <> 0 or Bonus - isnull(UBonus,0) <> 0)"

   With CN.Execute(ssql)
      If .RecordCount = 0 Then
         CN.Execute ("Update PurchaseOrderHeader set IsPurchase = 1 Where OrderID = " & Val(TxtOrderID.Text) & " And Orderdate ='" & DtpOrderDate.DateValue & "'")
      End If
   End With
   
   '''' update product purprice if any product is deleted from purchase price
   vDel = Split(vDelProductID, " ")
   For i = 0 To UBound(vDel)
      vDelProductID = vDel(i)
      If vDelProductID <> "" Then
         ssql = "if exists (Select top 1 Price from PurchaseBody where productid = " & vDelProductID & ") Begin Update products Set IsSync = 0, PurPrice = (Select top 1 Price from PurchaseBody where productid = " & vDelProductID & " Order by purchasedate desc) , PurDiscPC = (Select top 1 DiscPC from PurchaseBody where productid = " & vDelProductID & " Order by purchasedate desc) Where productid = " & vDelProductID & " End"
         CN.Execute (ssql)
      End If
   Next
            
'   If vIsNewRecord = True Then Call ActivityLog("Purchase Invoice", eAdd, TxtPurchaseID.Text, DtpPurchaseDate.DateValue)
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmPurchaseInvoice, eAdd, TxtPurchaseID.Text, DtpPurchaseDate.DateValue, Grid.Rows - 1 & " New Product/s Added Amount: " & Val(TxtNetAmount.Text))
   CN.CommitTrans
   
   If ObjRegistry.ShowChangePriceOnSavePI = True And (Me.ActiveControl.Name <> BtnChangePrice.Name) And (Me.ActiveControl.Name <> BtnBarCode.Name) Then BtnChangePrice_Click
   
'   If ObjRegistry.IsShowPrintYesOrNo = True Then
'      If MsgBox("Are you sure to print Current Purchase invoice", vbInformation + vbYesNo, "Alert") = vbYes Then
'         Call BtnPrint_Click
'      End If
'   End If
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
      " where b.productid = " & Val(TxtProductID.Text) & " and h.purchasedate <= '" & DtpPurchaseDate.DateValue & "' order by b.PurchaseDate Desc, b.purid Desc"
      
      With CN.Execute(ssql)
         GridHistory.Redraw = False
         GridHistory.MoveFirst
         GridHistory.RemoveAll
         GridHistory.AllowAddNew = True
         
         If ObjRegistry.isShowListPrice Then
            If Not .EOF Then
               LblCaptionRetailPrice = "Last Price"
               LblRetailPrice.Caption = !Price & ", DiscPack=" & Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * Val(IIf(IsNull(!Multiplier), "1", !Multiplier)) & vbCrLf & _
               "Disc%=" & !DiscPer & ", " & Format(!PurchaseDate, "dd/MM/yyyy")
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
   RsBody.Open "Select * from PurchaseBody where SID = " & Val(TxtSID.Text) & " and Purchasedate = '" & DtpPurchaseDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.ProductName, code,  b.* from Purchasebody b join products p on p.productid = b.productid where SID=" & Val(TxtSID.Text) & " and Purchasedate='" & DtpPurchaseDate.DateValue & "' order by SerialNo"
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         TxtSumDiscAmount.Text = 0
         TxtGrossAmount.Text = 0
         TxtTotalItems.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("SRNo").Text = Grid.Rows
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = !Code
            Grid.Columns("BatchNo").Text = IIf(IsNull(!BatchNo), "", !BatchNo)
            Grid.Columns("ExpiryDate").Text = IIf(IsNull(!ExpiryDate), "", !ExpiryDate)
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
            Grid.Columns("GrossQty").Value = IIf(IsNull(!GrossQty), "", !GrossQty)
            Grid.Columns("GrossUnit").Value = IIf(IsNull(!GrossUnit), "", !GrossUnit)
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            Grid.Columns("QtyLoose").Value = !QtyLoose
            Grid.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            Grid.Columns("Price").Value = !Price
            
            Grid.Columns("RetailPrice").Value = !RetailPrice
            Grid.Columns("IsWSDiscb4ST").Value = !IsWSDiscb4ST
            Grid.Columns("IsWSSaleTax").Value = !IsWSSaleTax
            Grid.Columns("IsRetailSaleTax").Value = !IsRetailSaleTax
            Grid.Columns("IsSerial").Value = !IsSerial
            vIsSerial = !IsSerial
            
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPack").Value = IIf(IsNull(!DiscPack), "", !DiscPack)
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            
            Grid.Columns("DiscPer2").Value = IIf(IsNull(!DiscPer2), "", !DiscPer2)
            Grid.Columns("DiscVal2").Value = IIf(IsNull(!DiscVal2), "", !DiscVal2)
            
            Grid.Columns("isDiscB4TradeOffer").Value = IIf(IsNull(!isDiscB4TradeOffer), "", !isDiscB4TradeOffer)
            Grid.Columns("isDiscB4ExtraScheme").Value = IIf(IsNull(!IsDiscB4ExtraScheme), "", !IsDiscB4ExtraScheme)
            Grid.Columns("isDiscB4SaleTax").Value = IIf(IsNull(!isDiscB4SaleTax), "", !isDiscB4SaleTax)
            Grid.Columns("TradeOffer1").Value = IIf(IsNull(!TradeOffer1), "", !TradeOffer1)
            Grid.Columns("TradeOffer2").Value = IIf(IsNull(!TradeOffer2), "", !TradeOffer2)
            Grid.Columns("ExtraSchemePer").Value = IIf(IsNull(!ExtraSchemePer), "", !ExtraSchemePer)
            Grid.Columns("TradeValue").Value = IIf(IsNull(!TradeValue), "", !TradeValue)
            Grid.Columns("ExtraSchemeValue").Value = IIf(IsNull(!ExtraSchemeValue), "", !ExtraSchemeValue)
            
            Grid.Columns("SC").Value = IIf(IsNull(!SC), "", !SC)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("DiscAmount").Value = !DiscAmount
            Grid.Columns("SaleDiscPer").Value = IIf(IsNull(!SaleDiscPer), "", !SaleDiscPer)
            Grid.Columns("RetailAmount").Value = IIf(IsNull(!RetailAmount), "", !RetailAmount)
            Grid.Columns("ProfitAmount").Value = IIf(IsNull(!ProfitAmount), "", !ProfitAmount)
            
            
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(!Amount) + Val(!DiscVal)
            TxtSumDiscAmount.Text = Val(TxtSumDiscAmount.Text) + Val(!DiscVal)
            TxtGrossAmount.Text = Val(TxtGrossAmount.Text) + Val(!Amount)
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
   RsBodySerial.Open "Select * from PurchaseBodySerial where PurID =" & Val(TxtPurchaseID.Text) & " and PurchaseDate ='" & DtpPurchaseDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   RsBodySerial.Filter = 0
   
   
   Call PopulateDataToGridOffer
   Call PopulateDataToGridExpense
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

Private Sub PopulatePOToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from PurchaseBody where PurID=" & Val(TxtPurchaseID.Text) & " and Purchasedate = '" & DtpPurchaseDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      ssql = " select p.ProductID, p.ItemCode, ProductName, QtyPack - isnull(UPack,0) as RQtyPack, QtyLoose - isnull(UQty,0) as RQty, Bonus - isnull(UBonus,0) as RBonus, P.DiscPer SaleDiscPer, sob.*" & vbCrLf _
      + " from (select OrderID, OrderDate, ProductID, Sum(QtyLoose) as UQty, Sum(QtyPack) as UPack, Sum(Bonus) as UBonus from PurchaseBody b inner join PurchaseHeader h on h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate Group By OrderID, OrderDate, ProductID) b " & vbCrLf _
      + " right outer join PurchaseOrderbody sob on sob.OrderID = b.orderid and sob.OrderDate = b.orderdate and b.ProductID = sob.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = sob.productid" & vbCrLf _
      + " where sob.OrderID = " & Val(TxtOrderID.Text) & " and sob.OrderDate = '" & DtpOrderDate.DateValue & "' and (QtyPack - isnull(UPack,0) <> 0 or QtyLoose - isnull(UQty,0) <> 0 or Bonus - isnull(UBonus,0) <> 0) order by sob.serialno"

      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtGrossAmount.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("SrNo").Text = Grid.Rows
            Grid.Columns("ProductID").Text = Val(!Productid)
            Grid.Columns("Code").Text = IIf(vColour = True, !ItemCode, !Productid)
            Grid.Columns("BatchNo").Text = IIf(IsNull(!BatchNo), "", !BatchNo)
            Grid.Columns("ExpiryDate").Text = IIf(IsNull(!ExpiryDate), "", !ExpiryDate)
            Grid.Columns("ProductName").Text = !ProductName
            
            RsBody.AddNew
            RsBody!Productid = Val(!Productid)
            RsBody!Code = IIf(vColour = True, !ItemCode, !Productid)
            
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
            Grid.Columns("QtyPack").Value = IIf(IsNull(!RQtyPack), "", !RQtyPack)
            Grid.Columns("QtyLoose").Value = !RQty
            Grid.Columns("Bonus").Value = IIf(IsNull(!RBonus), "", !RBonus)
            Grid.Columns("Price").Value = !Price
            GetOldPrice
            Grid.Columns("OldPrice").Value = IIf(vOldPrice = 0, "", vOldPrice)
            Grid.Columns("RetailPrice").Value = !RetailPrice
            Grid.Columns("IsWSDiscb4ST").Value = !IsWSDiscb4ST
            Grid.Columns("IsWSSaleTax").Value = !IsWSSaleTax
            Grid.Columns("IsRetailSaleTax").Value = !IsRetailSaleTax
            
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscPack").Value = IIf(IsNull(!DiscPack), "", !DiscPack)
            Grid.Columns("DiscVal").Value = Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty) 'IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = ((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty) '!Amount
            Grid.Columns("SaleDiscPer").Value = IIf(IsNull(!SaleDiscPer), "", !SaleDiscPer)
            Grid.Columns("RetailAmount").Value = Round((Val(!RetailPrice) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), 0, !Multiplier) + IIf(IsNull(!RQty), 0, !RQty))) - (!RetailPrice * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), 0, !Multiplier) + IIf(IsNull(!RQty), 0, !RQty)) * IIf(IsNull(!SaleDiscPer), 0, !SaleDiscPer) / 100), 2) '!RetailAmount
            Grid.Columns("ProfitAmount").Value = Grid.Columns("RetailAmount").Value - Grid.Columns("Amount").Value
            '''''
            RsBody!Multiplier = IIf(IsNull(!Multiplier), Null, !Multiplier)
            RsBody!QtyPack = IIf(IsNull(!RQtyPack), Null, !RQtyPack)
            RsBody!QtyLoose = !RQty
            RsBody!Bonus = IIf(IsNull(!RBonus), Null, !RBonus)
            RsBody!Price = !Price
            RsBody!OldPrice = IIf(vOldPrice = 0, Null, vOldPrice)
            
            RsBody!RetailPrice = !RetailPrice
            RsBody!IsWSDiscb4ST = !IsWSDiscb4ST
            RsBody!IsWSSaleTax = !IsWSSaleTax
            RsBody!IsRetailSaleTax = !IsRetailSaleTax
            
            RsBody!DiscPC = IIf(IsNull(!DiscPC), Null, !DiscPC)
            RsBody!DiscPack = IIf(IsNull(!DiscPack), Null, !DiscPack)

            RsBody!Offer = IIf(IsNull(!Offer), Null, !Offer)
            RsBody!SaleTaxPer = IIf(IsNull(!SaleTaxPer), Null, !SaleTaxPer)
            RsBody!SaleTaxval = IIf(IsNull(!SaleTaxval), Null, !SaleTaxval)
            RsBody!DiscPer = IIf(IsNull(!DiscPer), Null, !DiscPer)
            RsBody!DiscVal = Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty) 'IIf(IsNull(!DiscVal), "", !DiscVal)
            RsBody!Amount = ((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty)  '!Amount
            RsBody!SaleDiscPer = IIf(IsNull(!SaleDiscPer), Null, !SaleDiscPer)
            RsBody!RetailAmount = Round((Val(!RetailPrice) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), 0, !Multiplier) + IIf(IsNull(!RQty), 0, !RQty))) - (!RetailPrice * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), 0, !Multiplier) + IIf(IsNull(!RQty), 0, !RQty)) * IIf(IsNull(!SaleDiscPer), 0, !SaleDiscPer) / 100), 2) '!RetailAmount
            RsBody!ProfitAmount = RsBody!RetailAmount - RsBody!Amount
            RsBody.Update
            ''''
            
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + ((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty)  '!Amount
            TxtGrossAmount.Text = Val(TxtGrossAmount.Text) + ((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty)  '!Amount
            TxtTotalItems.Text = Val(TxtTotalItems.Text) + !RQty + IIf(IsNull(!RBonus), "0", !RBonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!RQtyPack), 0, !RQtyPack))
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
'   End If
   
   RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   RsBodySerial.Open "Select * from PurchaseBodySerial where PurID=" & Val(TxtPurchaseID.Text) & " and Purchasedate = '" & DtpPurchaseDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
   Call PopulatePOToGridOffer
   Call PopulatePOToGridSerial
   Call PopulatePOToGridExpense
End Sub

Private Sub PopulateDataToGridExpense()
    If RsExpense.State = adStateOpen Then RsExpense.Close
    RsExpense.Open "Select * from PurchaseExpense where PurID =" & Val(TxtPurchaseID.Text) & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
'    GridExpense.Visible = True
    ssql = "select EA.AccountNo, Accountname, PE.ExpAmount from ExpenseAccounts EA Left Outer join ChartofAccounts C on C.AccountNo = EA.AccountNo Left Outer Join (Select * from PurchaseExpense where PurID =" & Val(TxtPurchaseID.Text) & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "') PE On PE.ExpenseID = EA.AccountNo"
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

Private Sub PopulatePOToGridExpense()
    If RsExpense.State = adStateOpen Then RsExpense.Close
    RsExpense.Open "Select * from PurchaseExpense where PurID =" & Val(TxtPurchaseID.Text) & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
'    GridExpense.Visible = True
    ssql = "select EA.AccountNo, Accountname, PE.ExpAmount from ExpenseAccounts EA Left Outer join ChartofAccounts C on C.AccountNo = EA.AccountNo Left Outer Join (Select * from PurchaseOrderExpense where OrderID =" & Val(TxtOrderID.Text) & " And OrderDate = '" & DtpOrderDate.DateValue & "') PE On PE.ExpenseID = EA.AccountNo"
      With CN.Execute(ssql)
         GridExpense.Redraw = False
         GridExpense.MoveFirst
         GridExpense.RemoveAll
         While Not .EOF
            GridExpense.AddNew
            GridExpense.Columns("ID").Text = !AccountNo
            GridExpense.Columns("Name").Text = !AccountName
            GridExpense.Columns("Value").Value = IIf(IsNull(!ExpAmount), 0, !ExpAmount)
'            TxtTotalExpense.Text = Val(TxtTotalExpense.Text) + Val(!expAmount)
            GridExpense.Update
            
            RsExpense.AddNew
            RsExpense!ExpenseID = !AccountNo
            RsExpense!ExpAmount = IIf(IsNull(!ExpAmount), 0, !ExpAmount)
            RsExpense.Update
            .MoveNext
         Wend
      End With

     If GridExpense.Rows > 0 Then GridExpense.FirstRow = 0
     GridExpense.Redraw = True
'      GridExpense.Visible = False
End Sub


Private Sub PopulateDataToGridserial()
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      RsBodySerial.Filter = 0
   Else
      RsBodySerial.Filter = 0
      RsBodySerial.Filter = "ProductID = " & Grid.Columns("ProductID").Text
   End If

   If RsBodySerial.RecordCount > 0 Then
'       sSql = "select d.* from PurchaseBodySerial d  where PurID=" & Val(TxtPurchaseID.Text) & " and Purchasedate='" & DtpPurchaseDate.DateValue & "' and ProductID = '" & Grid.Columns("ProductID").Text & "'"
'      With CN.Execute(sSql)
       With RsBodySerial
         GridSerial.Redraw = False
         GridSerial.MoveFirst
         GridSerial.RemoveAll
         GridSerial.AllowAddNew = True
         .MoveFirst
         While Not .EOF
            GridSerial.AddNew
            GridSerial.Columns("ProductID").Text = !Productid
            GridSerial.Columns("Serial").Text = !serial
            .MoveNext
         Wend
'         .Close
      End With
      GridSerial.AddNew
      GridSerial.Columns("Serial").Text = " "
      GridSerial.AllowAddNew = False
      GridSerial.Redraw = True
   Else
    Call SubClearSerialFields
   End If
   RsBodySerial.Filter = 0
End Sub

Private Sub PopulatePOToGridSerial()
'   RsBodySerial.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
'   If RsBodySerial.RecordCount > 0 Then
       ssql = "select d.* from PurchaseOrderBodySerial d  where OrderID=" & Val(TxtOrderID.Text) & " and Orderdate='" & DtpOrderDate.DateValue & "'"
      With CN.Execute(ssql)
'       With RsBodySerial
'         GridSerial.Redraw = False
'         GridSerial.MoveFirst
'         GridSerial.RemoveAll
'         GridSerial.AllowAddNew = True
'         .MoveFirst
         While Not .EOF
'            GridSerial.AddNew
'            GridSerial.Columns("ProductID").Text = !Productid
'            GridSerial.Columns("Serial").Text = !Serial
             
            RsBodySerial.AddNew
            RsBodySerial!Productid = !Productid
            RsBodySerial!Productid = !serial
            RsBodySerial.Update
            .MoveNext
         Wend
'         .Close
      End With
'      GridSerial.AddNew
'      GridSerial.Columns("Serial").Text = " "
'      GridSerial.AllowAddNew = False
'      GridSerial.Redraw = True
'   Else
'    Call SubClearSerialFields
'   End If
'   RsBodySerial.Filter = 0
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
      FrmHistory.Visible = False
      FrmProductPrices.Visible = False
      vServerDate = CN.Execute("Select CONVERT(datetime, CONVERT(varchar, GETDATE(), 110)) ServerDate").Fields(0).Value
      vSystemDate = Abs(ObjRegistry.SystemDate)
      vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
      vDate = IIf(vSystemDate = True, CN.Execute("Select SystemDate From SystemDate").Fields(0).Value, vServerDate)
'      vDate = IIf(vSystemDate = True, IIf(IsNull(vDate), IIf(Format(Now, "hh") >= vHDiff, Date, DateAdd("d", -1, Date)), Date), IIf(Format(cn.Execute("Select getdate()").Fields(0).Value, "hh") >= vHDiff, vDate, DateAdd("d", -1, vDate)))
      
      If vSystemDate = True Then
         If IsNull(vDate) Then
            If Format(Now, "hh") >= vHDiff Then
               vDate = Date
            Else
               vDate = DateAdd("d", -1, Date)
            End If
         Else
            If Format(Now, "hh") >= vHDiff Then
               vDate = vDate
            Else
               vDate = DateAdd("d", -1, vDate)
            End If
         End If
      Else
         If Format(CN.Execute("Select getdate()").Fields(0).Value, "hh") >= vHDiff Then
            vDate = vDate
         Else
            vDate = DateAdd("d", -1, vDate)
         End If
      End If
      
      DtpPurchaseDate.DateValue = vDate
      vNow = vDate & " " & Format(IIf(vSystemDate = True, Now, CN.Execute("Select getdate()").Fields(0).Value), "hh:mm:ss")
      DtpEntryDate.DateValue = vNow
      TxtPurchaseID.Text = FunGetMaxID()
      Call PopulateDataToGrid
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      TxtCode.Enabled = True
      TxtStoreID.Enabled = True
      BtnStore.Enabled = True
      LblStock.Visible = False
      LblAllStock.Visible = False
      LblStockCaption.Visible = False
      LblCaptionRetailPrice.Visible = False
      LblRetailPrice.Visible = False
      BtnProduct.Enabled = True
      'TxtPurchaseID.Enabled = True
      DtpPurchaseDate.Enabled = True
      If DtpPurchaseDate.Enabled And DtpPurchaseDate.Visible Then DtpPurchaseDate.SetFocus
      GridOffer.Visible = False
      FramExpense.ZOrder 0
      vIsNewRecord = True
   Case Is = OpenMode
      'TxtPurchaseID.Enabled = False
      DtpPurchaseDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      'TxtStoreID.Enabled = False
      'BtnStore.Enabled = False
      LblStock.Visible = False
      LblAllStock.Visible = False
      LblStockCaption.Visible = False
      LblCaptionRetailPrice.Visible = False
      LblRetailPrice.Visible = False
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      'DtpPurchaseDate.SetFocus
      DtpEntryDate.SetFocus
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
      If TxtOrganizationID.Visible And TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
      TxtBillNo.SetFocus
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
         With CN.Execute("select * from ProductPacking where ProductID = " & Val(TxtProductID.Text) & " and packingid=" & CmbPackName.ItemData(CmbPackName.ListIndex))
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

Private Sub DtpPurchaseDate_Validate(Cancel As Boolean)
   If DtpPurchaseDate.Enabled = False Then Exit Sub
   TxtPurchaseID.Text = FunGetMaxID()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtCode.Enabled Then
         TxtCode.SetFocus
         Call SubClearDetailArea
         RsBodySerial.Filter = ""
         RsBodySerial.Filter = "ProductID = " & Val(TxtProductID.Text)
         If RsBodySerial.RecordCount > 0 Then
            RsBodySerial.Delete
            SubClearSerialFields
         End If
      End If
   ElseIf Shift = vbCtrlMask Then
      If KeyCode = vbKeyDelete Then
         If ActiveControl.Name = Grid.Name Then
            If Trim(Grid.Columns("ProductID").Text <> "") Then Call mniRemoveRow_Click
         ElseIf ActiveControl.Name = GridSerial.Name Then
            If Trim(GridSerial.Columns("ProductID").Text <> "") Then Call mniRemoveRow_Click
         ElseIf ActiveControl.Name = GridOffer.Name Then
            If Trim(GridOffer.Columns("ProductID").Text <> "") Then Call mniRemoveRow_Click
         Else
            KeyCode = 0: Exit Sub
         End If
      End If
      Select Case KeyCode
         Case vbKeyS
'            If ObjRegistry.ShowChangePriceOnSavePI = True Then
'               BtnChangePrice_Click
'            Else
'               If BtnSave.Enabled Then BtnSave_Click
'            End If
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
         Case vbKeyB
            If BtnBarCode.Enabled Then BtnBarCode_Click
            KeyCode = 0
         Case vbKeyC
            If BtnChangePrice.Enabled Then BtnChangePrice_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then If TxtBatchNo.Visible Then TxtBatchNo.SetFocus Else GetDataFromTexBoxesToGrid
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, False) = True Then TxtBillNo.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus Else TxtOrganizationID.SetFocus Else TxtStoreID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then If TxtVenderID.Enabled Then TxtVenderID.SetFocus Else TxtOrganizationID.SetFocus
      End Select
   ElseIf KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name
         Grid_DblClick
      Case GridSerial.Name
         GridSerial_DblClick
      Case TxtCode.Name
         If FunSelectProduct(ssValidate, False) = True Then If TxtBatchNo.Visible Then TxtBatchNo.SetFocus Else GetDataFromTexBoxesToGrid
      Case TxtDiscVal2.Name, TxtAmount.Name
         'Grid.SetFocus
         
         If vIsSerial = False Then
            GetDataFromTexBoxesToGrid
         Else
            Frame1.Visible = True
            Frame1.ZOrder 0
            TxtSerial.Enabled = True
            TxtSerial.SetFocus
         End If
      Case TxtProductName.Name
           ' Call SearchCode
      Case Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End Select
      
   ElseIf KeyCode = vbKeyF2 Then
      If Frame1.Visible = True Then
         Frame1.Visible = False
         If TxtCode.Enabled = True Then TxtCode.SetFocus Else Grid.SetFocus
      Else
            Frame1.Visible = True
            Frame1.ZOrder 0
            If TxtSerial.Enabled = True Then TxtSerial.SetFocus
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
   SetWindowText Me.hWnd, "Purchase Invoice"
   HelpLocation Me
   
   vServerDate = CN.Execute("Select CONVERT(datetime, CONVERT(varchar, GETDATE(), 110)) ServerDate").Fields(0).Value
   vSystemDate = Abs(ObjRegistry.SystemDate)
   vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
   
   
   
   
   vDiscPackFlag = False
'   With CN.Execute("Select * from Packings")
'      CmbPackName.AddItem ""
'      While Not .EOF
'         CmbPackName.AddItem !Packingname
'         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
'         .MoveNext
'      Wend
'      .Close
'   End With
   vNoofPrints = IIf(IsNull(ObjUserSecurity.NoofPurPrints) Or Val(ObjUserSecurity.NoofPurPrints) = 0, 1, ObjUserSecurity.NoofPurPrints)
   If ObjRegistry.ShowDispatchDate = True Then
      LblTerms.Visible = True
      TxtTerms.Visible = True
      TxtTerms.Text = CN.Execute("Select value from sysindexs Where Registrykey = 'PurchasePromiseTerms'").Fields(0).Value
   Else
      LblTerms.Visible = False
      TxtTerms.Visible = False
      TxtTerms.Text = 0
   End If
   
   If ObjUserSecurity.ShowStock = True Or ObjUserSecurity.IsAdministrator Then
      vShowStock = True
   Else
      vShowStock = False
   End If
   
   
   LblBarCode.Visible = ObjRegistry.ShowAddBarCode
   TxtBarcode.Visible = ObjRegistry.ShowAddBarCode
   
   LblRetailAmount.Visible = ObjRegistry.ShowPurchaseProfit
   LblProfitAmount.Visible = ObjRegistry.ShowPurchaseProfit
   LblSaleDiscPer.Visible = ObjRegistry.ShowPurchaseProfit
   TxtRetailAmount.Visible = ObjRegistry.ShowPurchaseProfit
   TxtProfitAmount.Visible = ObjRegistry.ShowPurchaseProfit
   TxtSaleDiscPer.Visible = ObjRegistry.ShowPurchaseProfit
   
   LblDiscPer2.Visible = ObjRegistry.ShowDisc2
   TxtDiscPer2.Visible = ObjRegistry.ShowDisc2
   LblDiscVal2.Visible = ObjRegistry.ShowDisc2
   TxtDiscVal2.Visible = ObjRegistry.ShowDisc2
   
   vNoofPrints = IIf(IsNull(ObjRegistry.NoofPrints) Or Val(ObjRegistry.NoofPrints) = 0, 1, ObjRegistry.NoofPrints)
   
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

   
   
   '''''''''''''''' Form Default Setting  ''''''''''''''''''''''
   ssql = "select * from FormDefaultSetting Where FormType = 'Purchase Invoice' and LocalComputerName = '" & LocalComputerName & "'"
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
   '''''''''''''''''''''''   '''''''''''''''''''''''''''''''''
   
   '''''''''''''''' isChangePrice Form Allow  ''''''''''''''''''''''
   If ObjUserSecurity.IsAdministrator = False Then
      ssql = "select * from usertasks Where TaskKey = 'MniChangePrice' And UserNo = " & vUser
      With CN.Execute(ssql)
        If .EOF = True Then
           BtnChangePrice.Visible = False
        End If
        .Close
      End With
   End If
   ''''''''''''''''''''''''''''''''''''''''''''''
   
   If ObjUserSecurity.IsAdministrator = False Then
      TxtDiscPack.Enabled = Not ObjRegistry.isShowListPrice
      TxtDiscPC.Enabled = Not ObjRegistry.isShowListPrice
      TxtDiscPer.Enabled = Not ObjRegistry.isShowListPrice
      TxtDiscVal.Enabled = Not ObjRegistry.isShowListPrice
   End If
   
   vColour = ObjRegistry.ShowColourSize
   
   LblColour.Visible = vColour
   CmbColourName.Visible = vColour
   LblSize.Visible = vColour
   cmbSizeName.Visible = vColour
   Grid.Columns("ColourName").Visible = vColour
   Grid.Columns("SizeName").Visible = vColour
   
'   If vColour = False Then
'      LblPackName.Left = LblPackName.Left - CmbColourName.Width - cmbSizeName.Width
'      CmbPackName.Left = CmbPackName.Left - CmbColourName.Width - cmbSizeName.Width
'      LblMultiplier.Left = LblMultiplier.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtMultiplier.Left = TxtMultiplier.Left - CmbColourName.Width - cmbSizeName.Width
'      LblQtyPack.Left = LblQtyPack.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtQtyPack.Left = TxtQtyPack.Left - CmbColourName.Width - cmbSizeName.Width
'      LblQtyLoose.Left = LblQtyLoose.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtQtyLoose.Left = TxtQtyLoose.Left - CmbColourName.Width - cmbSizeName.Width
'      LblBonus.Left = LblBonus.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtBonus.Left = TxtBonus.Left - CmbColourName.Width - cmbSizeName.Width
'      LblOffer.Left = LblOffer.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtOffer.Left = TxtOffer.Left - CmbColourName.Width - cmbSizeName.Width
'      LblPrice.Left = LblPrice.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtPrice.Left = TxtPrice.Left - CmbColourName.Width - cmbSizeName.Width
'      LblDiscPC.Left = LblDiscPC.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtDiscPC.Left = TxtDiscPC.Left - CmbColourName.Width - cmbSizeName.Width
'
'      LblDiscPack.Left = LblDiscPack.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtDiscPack.Left = TxtDiscPack.Left - CmbColourName.Width - cmbSizeName.Width
'
'      LblSaleTaxPer.Left = LblSaleTaxPer.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtSaleTaxPer.Left = TxtSaleTaxPer.Left - CmbColourName.Width - cmbSizeName.Width
'      LblDiscPer.Left = LblDiscPer.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtDiscPer.Left = TxtDiscPer.Left - CmbColourName.Width - cmbSizeName.Width
'      LblDiscVal.Left = LblDiscVal.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtDiscVal.Left = TxtDiscVal.Left - CmbColourName.Width - cmbSizeName.Width
'      LblSaleTaxVal.Left = LblSaleTaxVal.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtSaleTaxVal.Left = TxtSaleTaxVal.Left - CmbColourName.Width - cmbSizeName.Width
'      LblAmount.Left = LblAmount.Left - CmbColourName.Width - cmbSizeName.Width
'      TxtAmount.Left = TxtAmount.Left - CmbColourName.Width - cmbSizeName.Width
'      Grid.Width = Grid.Width - CmbColourName.Width - cmbSizeName.Width
'   End If
'
   TxtStoreID.Text = IIf((ObjRegistry.StoreID = ""), "", ObjRegistry.StoreID)
   FunSelectStore ssValidate, True
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   
   TxtBatchNo.Visible = ObjRegistry.BatchNoVisible
   DtpExpiryDate.Visible = ObjRegistry.BatchNoVisible
   If ObjRegistry.ShowPromiseDateInSalaPurchase = True Then
      LblPromiseDate.Visible = True
      DtpPromiseDate.Visible = True
      DtpPromiseDate.DateValue = Null
   Else
      LblPromiseDate.Visible = False
      DtpPromiseDate.Visible = False
      DtpPromiseDate.DateValue = Null
   End If
   
   vTradeOffer = ObjRegistry.ShowTradeOffer
   LblTradeOffer.Visible = vTradeOffer
   TxtTradeOffer1.Visible = vTradeOffer
   TxtTradeOffer2.Visible = vTradeOffer
   LblPlusSign.Visible = vTradeOffer
   LblTradeValue.Visible = vTradeOffer
   TxtTradeOfferValue.Visible = vTradeOffer
   ChkDiscB4TradeOffer.Visible = vTradeOffer
   
'   LblExtraSchemePer.Visible = vTradeOffer
'   TxtExtraSchemePer.Visible = vTradeOffer
'   LblExtraSchemeValue.Visible = vTradeOffer
'   TxtExtraSchemeValue.Visible = vTradeOffer
'   LblGSTValue.Visible = vTradeOffer
'   TxtSaleTaxVal.Visible = vTradeOffer
'   ChkDiscB4ExtraScheme.Visible = vTradeOffer
'   ChkDiscB4SaleTax.Visible = vTradeOffer
'
   If ObjRegistry.BatchNoVisible = False Then LblProductName.Left = TxtProductName.Left

   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   Frame2.Visible = ObjRegistry.FreightVisible
   LblFreight.Visible = ObjRegistry.FreightVisible
   TxtFreight.Visible = ObjRegistry.FreightVisible
   
   If ObjUserSecurity.IsAdministrator = False Then
      LblPrice.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtPrice.Visible = Not ObjRegistry.HidePurchaseAmount
      LblAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      LblTotalAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtGrossAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      LblNetAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtNetAmount.Visible = Not ObjRegistry.HidePurchaseAmount
      lblPayable.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtPreviousPayable.Visible = Not ObjRegistry.HidePurchaseAmount
      LblTtlPayable.Visible = Not ObjRegistry.HidePurchaseAmount
      TxtTotalPayable.Visible = Not ObjRegistry.HidePurchaseAmount
      Grid.Columns("Price").Visible = Not ObjRegistry.HidePurchaseAmount
      Grid.Columns("Amount").Visible = Not ObjRegistry.HidePurchaseAmount
   End If
   
   LblGrossQty.Visible = ObjRegistry.IsGrossQty
   TxtGrossQty.Visible = ObjRegistry.IsGrossQty
   LblGrossUnit.Visible = ObjRegistry.IsGrossQty
   TxtGrossUnit.Visible = ObjRegistry.IsGrossQty
   
   With CN.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
      End If
      .Close
   End With
   
   If ObjRegistry.ShowBonus = False Then
      LblBonus.Visible = False
      TxtBonus.Visible = False
      Grid.Columns("Bonus").Visible = False
      
      LblOffer.Left = LblOffer.Left - TxtBonus.Width
      TxtOffer.Left = TxtOffer.Left - TxtBonus.Width
      
      LblPrice.Left = LblPrice.Left - TxtBonus.Width
      TxtPrice.Left = TxtPrice.Left - TxtBonus.Width
      
      LblRetail.Left = LblRetail.Left - TxtBonus.Width
      TxtRetailPrice.Left = TxtRetailPrice.Left - TxtBonus.Width
      
      LblDiscPC.Left = LblDiscPC.Left - TxtBonus.Width
      TxtDiscPC.Left = TxtDiscPC.Left - TxtBonus.Width
      
      LblDiscPack.Left = LblDiscPack.Left - TxtBonus.Width
      TxtDiscPack.Left = TxtDiscPack.Left - TxtBonus.Width
      
     
      
      LblDiscPer.Left = LblDiscPer.Left - TxtBonus.Width
      TxtDiscPer.Left = TxtDiscPer.Left - TxtBonus.Width
      
      LblDiscVal.Left = LblDiscVal.Left - TxtBonus.Width
      TxtDiscVal.Left = TxtDiscVal.Left - TxtBonus.Width
      
      LblSaleTaxPer.Left = LblSaleTaxPer.Left - TxtBonus.Width
      TxtSaleTaxPer.Left = TxtSaleTaxPer.Left - TxtBonus.Width
      
      LblSaleTaxVal.Left = LblSaleTaxVal.Left - TxtBonus.Width
      TxtSaleTaxVal.Left = TxtSaleTaxVal.Left - TxtBonus.Width
      
      LblAmount.Left = LblAmount.Left - TxtBonus.Width
      TxtAmount.Left = TxtAmount.Left - TxtBonus.Width
      
      Grid.Width = Grid.Width - TxtBonus.Width
   End If
   
   If ObjRegistry.ShowOffer = False Then
      LblOffer.Visible = False
      TxtOffer.Visible = False
      Grid.Columns("Offer").Visible = False
      
      LblPrice.Left = LblPrice.Left - TxtOffer.Width
      TxtPrice.Left = TxtPrice.Left - TxtOffer.Width
      
      LblRetail.Left = LblRetail.Left - TxtOffer.Width
      TxtRetailPrice.Left = TxtRetailPrice.Left - TxtOffer.Width
      
      LblDiscPC.Left = LblDiscPC.Left - TxtOffer.Width
      TxtDiscPC.Left = TxtDiscPC.Left - TxtOffer.Width
      
      LblDiscPack.Left = LblDiscPack.Left - TxtOffer.Width
      TxtDiscPack.Left = TxtDiscPack.Left - TxtOffer.Width
      
      
      
      LblDiscPer.Left = LblDiscPer.Left - TxtOffer.Width
      TxtDiscPer.Left = TxtDiscPer.Left - TxtOffer.Width
      
      LblDiscVal.Left = LblDiscVal.Left - TxtOffer.Width
      TxtDiscVal.Left = TxtDiscVal.Left - TxtOffer.Width
      
      LblSaleTaxPer.Left = LblSaleTaxPer.Left - TxtOffer.Width
      TxtSaleTaxPer.Left = TxtSaleTaxPer.Left - TxtOffer.Width
      
      LblSaleTaxVal.Left = LblSaleTaxVal.Left - TxtOffer.Width
      TxtSaleTaxVal.Left = TxtSaleTaxVal.Left - TxtOffer.Width
      
      LblAmount.Left = LblAmount.Left - TxtOffer.Width
      TxtAmount.Left = TxtAmount.Left - TxtOffer.Width
      
      Grid.Width = Grid.Width - TxtOffer.Width
   End If
  
   If ObjRegistry.ShowChangeRetailInPurchaseInvoice = False Then
      LblRetail.Visible = False
      TxtRetailPrice.Visible = False
      Grid.Columns("RetailPrice").Visible = False
      
      LblDiscPC.Left = LblDiscPC.Left - TxtRetailPrice.Width
      TxtDiscPC.Left = TxtDiscPC.Left - TxtRetailPrice.Width
      
      LblDiscPack.Left = LblDiscPack.Left - TxtRetailPrice.Width
      TxtDiscPack.Left = TxtDiscPack.Left - TxtRetailPrice.Width
      
      
      
      LblDiscPer.Left = LblDiscPer.Left - TxtRetailPrice.Width
      TxtDiscPer.Left = TxtDiscPer.Left - TxtRetailPrice.Width
      
      LblDiscVal.Left = LblDiscVal.Left - TxtRetailPrice.Width
      TxtDiscVal.Left = TxtDiscVal.Left - TxtRetailPrice.Width
      
      LblSaleTaxPer.Left = LblSaleTaxPer.Left - TxtRetailPrice.Width
      TxtSaleTaxPer.Left = TxtSaleTaxPer.Left - TxtRetailPrice.Width
      
      LblSaleTaxVal.Left = LblSaleTaxVal.Left - TxtRetailPrice.Width
      TxtSaleTaxVal.Left = TxtSaleTaxVal.Left - TxtRetailPrice.Width
      
      LblAmount.Left = LblAmount.Left - TxtRetailPrice.Width
      TxtAmount.Left = TxtAmount.Left - TxtRetailPrice.Width
      
      Grid.Width = Grid.Width - TxtRetailPrice.Width
   
   End If
      
   If ObjRegistry.isShowListPrice = False Then
      
      LblDiscPack.Visible = False
      TxtDiscPack.Visible = False
      
      Grid.Columns("DiscPack").Visible = False
      
      LblSaleTaxPer.Left = LblSaleTaxPer.Left - TxtDiscPack.Width
      TxtSaleTaxPer.Left = TxtSaleTaxPer.Left - TxtDiscPack.Width
      
      LblDiscPer.Left = LblDiscPer.Left - TxtDiscPack.Width
      TxtDiscPer.Left = TxtDiscPer.Left - TxtDiscPack.Width
      
      LblDiscVal.Left = LblDiscVal.Left - TxtDiscPack.Width
      TxtDiscVal.Left = TxtDiscVal.Left - TxtDiscPack.Width
      
      LblSaleTaxVal.Left = LblSaleTaxVal.Left - TxtDiscPack.Width
      TxtSaleTaxVal.Left = TxtSaleTaxVal.Left - TxtDiscPack.Width
      
      LblAmount.Left = LblAmount.Left - TxtDiscPack.Width
      TxtAmount.Left = TxtAmount.Left - TxtDiscPack.Width
      
      Grid.Width = Grid.Width - TxtDiscPack.Width
   End If
  
  
   LblSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   TxtSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   LblSaleTaxVal.Visible = ObjRegistry.ShowSaleTax
   TxtSaleTaxVal.Visible = ObjRegistry.ShowSaleTax
   ChkDiscB4SaleTax.Visible = ObjRegistry.ShowSaleTax
   
   
   If ObjRegistry.ShowSaleTax = False Then
      LblSaleTaxPer.Visible = False
      TxtSaleTaxPer.Visible = False
      LblSaleTaxVal.Visible = False
      TxtSaleTaxVal.Visible = False
    
    Grid.Columns("SaleTaxPer").Visible = False
    Grid.Columns("SaleTaxVal").Visible = False
        
'    LblDiscPer.Left = LblDiscPer.Left - TxtSaleTaxPer.Width
'    TxtDiscPer.Left = TxtDiscPer.Left - TxtSaleTaxPer.Width
'
'    LblDiscVal.Left = LblDiscVal.Left - TxtSaleTaxPer.Width
'    TxtDiscVal.Left = TxtDiscVal.Left - TxtSaleTaxPer.Width
    
    LblAmount.Left = LblAmount.Left - TxtSaleTaxPer.Width - TxtSaleTaxVal.Width
    TxtAmount.Left = TxtAmount.Left - TxtSaleTaxPer.Width - TxtSaleTaxVal.Width
    
    Grid.Width = Grid.Width - TxtSaleTaxPer.Width - TxtSaleTaxVal.Width
   End If
 
   LblSC.Visible = False
   TxtSC.Visible = False
    
   LblExtraSchemePer.Visible = vTradeOffer
   TxtExtraSchemePer.Visible = vTradeOffer
   LblExtraSchemeValue.Visible = vTradeOffer
   TxtExtraSchemeValue.Visible = vTradeOffer
   ChkDiscB4ExtraScheme.Visible = vTradeOffer
   FormStatus = NewMode
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpPurchaseDate.IsDateValid = False Then Exit Function
   If ObjRegistry.AllowContinuousBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(PurID),0)+1 from PurchaseHeader").Fields(0)
   ElseIf ObjRegistry.AllowMonthlyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(PurID),0)+1 from PurchaseHeader where Month(Purchasedate) = '" & Month(DtpPurchaseDate.DateValue) & "' and  year(Purchasedate) ='" & Year(DtpPurchaseDate.DateValue) & "'").Fields(0)
   ElseIf ObjRegistry.AllowDailyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(PurID),0)+1 from PurchaseHeader where Purchasedate = '" & DtpPurchaseDate.DateValue & "'").Fields(0)
   Else
'      FunGetMaxID = cn.Execute("Select isnull(max(PurID),0)+1 from PurchaseHeader where Purchasedate = '" & DtpPurchaseDate.DateValue & "' and StoreID = " & TxtStoreID.Text).Fields(0)
      FunGetMaxID = CN.Execute("Select isnull(max(PurID),0)+1 from PurchaseHeader where Purchasedate = '" & DtpPurchaseDate.DateValue & "'").Fields(0)
   End If

   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpPurchaseDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = CN.Execute("Select isnull(max(BinID),0)+1 from Bin_PurchaseHeader ").Fields(0)
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
   vAdvDiscPerFlag = True
   vDelProductID = ""
   ChkPurchaseReplacement.Value = 0
   TxtNetAmount.Text = 0
   Grid.CancelUpdate
   Grid.DataMode = ssDataModeAddItem
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("Code").Text = " "
   Grid.Update
   GridOffer.CancelUpdate
   GridOffer.RemoveAll
   GridOffer.AddNew
   GridOffer.Columns("ProductID").Text = " "
   GridOffer.Update
   GridOffer.Visible = False
   DtpExpiryDate.DateValue = ""
   DtpPromiseDate.DateValue = Null
   If ObjRegistry.TermAllowZero = False Then TxtTerms.Text = 0
   If ObjRegistry.ShowDispatchDate = True And ObjRegistry.ShowPromiseDateInSalaPurchase Then
      DtpPromiseDate.DateValue = DateAdd("d", Val(TxtTerms.Text), DtpPurchaseDate.DateValue)
   End If
   Call SubClearSerialFields
   Frame1.Visible = False
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
    Set FrmPurchaseInvoice = Nothing
   End If
   '''''''''''''''''' ActivityLogBin For Close Action
'      Call DeleteTempActivityLogBin(vRandomID)
      If Grid.Rows > 1 And Cancel = 0 Then
         vGridRows = 0
         Grid.Redraw = False
         Grid.MoveFirst
         For vCounter = 2 To Grid.Rows
            vGridRows = vGridRows + 1
            If Trim(Grid.Columns("Code").Text) <> "" Then
               ssql = "Select Productid From purchasebody where PurID = " & Val(TxtPurchaseID.Text) & " and PurchaseDate='" & DtpPurchaseDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
               With CN.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmPurchaseInvoice, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtPurchaseID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPurchaseDate.Date), "Closed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmPurchaseInvoice, eCloseSavedRecord, TxtPurchaseID.Text, DtpPurchaseDate.DateValue, vGridRows & " Product/s Closed")
         Grid.Redraw = True
      End If
  ''''''''''''''''''
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("Amount").Value - Grid.Columns("DiscVal").Value
   TxtSumDiscAmount.Text = Val(TxtSumDiscAmount.Text) - Grid.Columns("DiscVal").Value
   TxtGrossAmount.Text = Val(TxtGrossAmount.Text) - Grid.Columns("Amount").Value
   TxtTotalItems.Text = Val(TxtTotalItems.Text) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Grid.Columns("QtyPack").Value)))
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   On Error GoTo ErrorHandler
   Call Grid_LostFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_GotFocus()
   On Error GoTo ErrorHandler
   Flag = True
   TxtCode.Enabled = False
   BtnProduct.Enabled = False
   'TxtCode.BackColor = TxtProductName.BackColor
   'TxtCode.TabStop = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_LostFocus()
   On Error GoTo ErrorHandler
   Flag = False
   If Trim(Grid.Columns("Code").Text) = "" Then
      TxtCode.Text = ""
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      TxtCode.SetFocus
   Else
      TxtCode.Enabled = False
      BtnProduct.Enabled = False
      'If Me.ActiveControl.Name = Grid.Name Then
      CmbPackName.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo ErrorHandler
    If Grid.Visible = False Then Exit Sub
    If Flag Then Call GetDataBackFromGridToTexBoxes
    Call PopulateDataToGridserial
    If Trim(Grid.Columns("Code").Text) = "" Then
'        TxtSerial.Enabled = False
    Else
        TxtSerial.Enabled = True
    End If
    GridOffer.MoveFirst
    For vCounter = 1 To GridOffer.Rows
    If GridOffer.Columns("ProductID").Text = Grid.Columns("ProductID").Text Then Exit Sub
        GridOffer.MoveNext
    Next vCounter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name = "Grid" Then
    If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
    
    If ObjRegistry.NegativeSale = False Then
      ssql = "Select QtyPack, Multiplier, Qtyloose From PurchaseBody Where SID = " & Val(TxtSID.Text) & " and Productid = " & Val(TxtProductID.Text)
         With CN.Execute(ssql)
            If .EOF = False Then
               If (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), 0, !Multiplier)) + IIf(IsNull(!QtyLoose), 0, !QtyLoose) > Val(vQtyLoose) Then
               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            Exit Sub
         End If
            End If
         .Close
         End With
    End If
    
   ssql = "Select Productid From Purchasebody where purid=" & Val(TxtPurchaseID.Text) & " and Purchasedate ='" & DtpPurchaseDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
   With CN.Execute(ssql)
      If .EOF Then
         Call ActivityLogBin("", eFrmPurchaseInvoice, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtPurchaseID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPurchaseDate.Date), "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      Else
         Call ActivityLogBin("", eFrmPurchaseInvoice, eRemoveRow, TxtPurchaseID.Text, DtpPurchaseDate.DateValue, "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
         Call ActivityLogBin(vRandomID, eFrmPurchaseInvoice, eAddTempRecord, TxtPurchaseID.Text, DtpPurchaseDate.DateValue, "Pending Remove Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      End If
   End With
    RsBody.Filter = "ProductID = " & Val(TxtProductID.Text) & " and BatchNo = " & IIf(Trim(TxtBatchNo.Text) = "", "null", "'" & Trim(TxtBatchNo.Text) & "'") & " and Price = " & Val(TxtPrice.Text)
    If RsBody.RecordCount > 0 Then RsBody.Delete
    RsProductOffer.Filter = "ProductID = " & Val(GridOffer.Columns("ProductID").Text)
    If RsProductOffer.RecordCount > 0 Then
        RsProductOffer.Delete
        GridOffer.SelBookmarks.RemoveAll
        GridOffer.SelBookmarks.Add GridOffer.Bookmark
        GridOffer.DeleteSelected
        GridOffer.Refresh
        RsProductOffer.Filter = 0
    End If
    RsBodySerial.Filter = ""
    RsBodySerial.Filter = "ProductID = " & Val(TxtProductID.Text)
    If RsBodySerial.RecordCount > 0 Then RsBodySerial.Delete
          
    Dim SrNo As Integer
    vDelProductID = vDelProductID & TxtProductID.Text & " "
    CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
    Grid.SelBookmarks.RemoveAll
    Grid.SelBookmarks.Add Grid.Bookmark
    SrNo = Val(Grid.Columns("SrNo").Text)
    Grid.DeleteSelected
    While Grid.Columns("ProductName").Text <> ""
      Grid.Columns("SrNo").Text = SrNo
      SrNo = SrNo + 1
      Grid.MoveNext
    Wend
    Grid.Refresh
    RsBody.Filter = 0
    Grid.MoveLast
    GetDataBackFromGridToTexBoxes
   ElseIf Me.ActiveControl.Name = "GridSerial" Then
    If TxtCode.Enabled = True Then
      MsgBox "Please Select the parent row to delete the child row", vbInformation + vbOKOnly, "Error"
      Exit Sub
    End If
    If Trim(GridSerial.Columns("Serial").Text) = "" Then Exit Sub
    RsBodySerial.Filter = "Serial = '" & TxtSerial.Text & "'"
    If RsBodySerial.RecordCount > 0 Then RsBodySerial.Delete
    CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Removed ProdcutID-" & GridSerial.Columns("ProductID").Text & " Serial-" & GridSerial.Columns("Serial").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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
   
   If CmbColourName.Text = "" And cmbSizeName.Text = "" And vColour = True Then
     MsgBox "Please Select Colour and Size", vbInformation + vbOKOnly, "Error"
     Exit Sub
   End If
   
   '' comment on 14/05/2022
'   If Val(TxtPrice.Text) <> 0 Then
'      If Abs(Round(Round(Val(TxtDiscPer.Text), 2) - (Round((Val(TxtDiscPC.Text) * 100) / (Val(TxtPrice.Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))), 2)))) > 0.01 Then
'         MsgBox "Please update the Discount for change Price.", vbExclamation, "Alert"
'         If TxtDiscPer.Enabled And TxtDiscPer.Visible Then TxtDiscPer.SetFocus
'         Exit Sub
'      End If
'   End If
   If vUnitRetailPrice < vUnitPrice Then
      MsgBox "Retail Price is Less Than Purchase Price", vbExclamation, "Alert"
   End If
   FrmHistory.Visible = False
   FrmProductPrices.Visible = False
   If ChkPurchaseReplacement.Value = 1 Then
      TxtQtyPack.Text = -1 * Val(TxtQtyPack.Text)
      TxtQtyLoose.Text = -1 * Val(TxtQtyLoose.Text)
      TxtAmount.Text = -1 * Val(TxtAmount.Text)
   End If
   Call BtnAddBarCode_Click
   If Trim(Grid.Columns("Productid").Text) = "" Then
'      vNewRow = True
      RsBody.Filter = "ProductID = " & Val(TxtProductID.Text) & " and BatchNo = " & IIf(Trim(TxtBatchNo.Text) = "", "null", "'" & Trim(TxtBatchNo.Text) & "'") & " and Price = " & Val(TxtPrice.Text)
   Else
'      vNewRow = False
      RsBody.Filter = "ProductID = " & Val(Grid.Columns("Productid").Text) & " and BatchNo = " & IIf(Grid.Columns("BatchNo").Text = "", "null", "'" & Grid.Columns("BatchNo").Text & "'") & " and Price = " & Val(Grid.Columns("Price").Text)
      'RsBody.Filter = "ProductID='" & Grid.Columns("Productid").Text & "'"
   End If
 
   
   
   If ObjRegistry.NegativeSale = False And TxtCode.Enabled = False And vIsNewRecord = False Then
     ssql = "Select QtyPack, Multiplier, Qtyloose From PurchaseBody Where SID = " & Val(TxtSID.Text) & " and Productid = " & Val(TxtProductID.Text)
     With CN.Execute(ssql)
         If ((IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), 0, !Multiplier)) + IIf(IsNull(!QtyLoose), 0, !QtyLoose)) - (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) > Val(vQtyLoose) Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            Exit Sub
         End If
         .Close
      End With
    End If
   
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("SRNo").Text = Grid.Rows
         Grid.Columns("ProductID").Text = TxtProductID.Text
         Grid.Columns("Code").Text = TxtCode.Text
         Grid.Columns("Price").Value = Val(TxtPrice.Text)
         GetOldPrice
         Grid.Columns("OldPrice").Value = IIf(vOldPrice = 0, "", vOldPrice)
         Grid.Columns("BatchNo").Text = Trim(TxtBatchNo.Text)
         RsBody!Productid = TxtProductID.Text
         RsBody!Code = TxtCode.Text
         RsBody!Price = Val(TxtPrice.Text)
         RsBody!BatchNo = Trim(TxtBatchNo.Text)
         RsBody!ExpiryDate = DtpExpiryDate.DateValue
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("Productid").Text = TxtProductID.Text And Grid.Columns("BatchNo").Text = Trim(TxtBatchNo.Text) And Val(Grid.Columns("Price").Text) = Val(TxtPrice.Text) Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  '''''' check expiry ''''''''
                  vExpiryTime = 0
'                  With cn.Execute("Select dbo.GetExpiryTime('" & TxtProductID.Text & "', " & IIf(TxtBatchNo.Text = "", "Null", "'" & TxtBatchNo.Text & "'") & " , getdate()) as Day ")
'                     If .RecordCount > 0 Then
'                        vExpiryTime = !Day
'                     End If
'                  End With
                  
                  ssql = "Select Productid From Purchasebody where purid=" & Val(TxtPurchaseID.Text) & " and Purchasedate ='" & DtpPurchaseDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmPurchaseInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtPurchaseID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPurchaseDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmPurchaseInvoice, eEdit, TxtPurchaseID.Text, DtpPurchaseDate.DateValue, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  If ObjRegistry.BatchNoVisible = True Then
                     ssql = "Select  isnull(min(DATEDIFF (day, getdate(), '" & DtpExpiryDate.DateValue & "')),0)"
                     vExpiryTime = CN.Execute(ssql).Fields(0).Value
                  End If
                  '''''''''''''''''''''''''This QtyOffer Is used for DetailGrid
                  QtyOffer = Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value) + Val(Grid.Columns("QtyLoose").Value)
                  GetDataFromTextBoxesToGridOffer
                  TxtOffer.Text = Val(TxtOffer.Text) + Val(Grid.Columns("Offer").Text)
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  TxtQtyPack.Text = Val(TxtQtyPack.Text) + Val(Grid.Columns("QtyPack").Value)
                  TxtQtyLoose.Text = Val(TxtQtyLoose.Text) + Val(Grid.Columns("QtyLoose").Value)
                  TxtBonus.Text = Val(TxtBonus.Text) + Val(Grid.Columns("Bonus").Value)
                  TxtTradeOfferValue.Text = Val(TxtTradeOfferValue.Text) + Val(TxtTradeOfferValue.Text) - Val(Grid.Columns("TradeValue").Text)
                  TxtExtraSchemeValue.Text = Val(TxtExtraSchemeValue.Text) + Val(TxtExtraSchemeValue.Text) - Val(Grid.Columns("ExtraSchemeValue").Text)
                  Call SubCalculateBody
                  
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text) + Val(TxtDiscAmount.Text) - Val(Grid.Columns("DiscAmount").Text)
                  TxtSumDiscAmount.Text = Val(TxtSumDiscAmount.Text) + Val(TxtDiscAmount.Text) - Val(Grid.Columns("DiscAmount").Text)
                  TxtGrossAmount.Text = Val(TxtGrossAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Val(Grid.Columns("QtyLoose").Value) + Val(Grid.Columns("Bonus").Value) + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
                  
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
                  Grid.Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
                  Grid.Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, "", Val(TxtMultiplier.Text))
                  Grid.Columns("GrossQty").Value = IIf(Val(TxtGrossQty.Text) = 0, 0, Val(TxtGrossQty.Text))
                  Grid.Columns("GrossUnit").Value = IIf(Val(TxtGrossUnit.Text) = 0, 0, Val(TxtGrossUnit.Text))
                  Grid.Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, "", Val(TxtQtyPack.Text))
                  Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
                  Grid.Columns("Bonus").Value = Val(TxtBonus.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  GetOldPrice
                  Grid.Columns("OldPrice").Value = IIf(vOldPrice = 0, "", vOldPrice)
                  Grid.Columns("RetailPrice").Value = Val(TxtRetailPrice.Text)
                  Grid.Columns("IsWSDiscb4ST").Value = vIsWSDiscb4ST
                  Grid.Columns("IsWSSaleTax").Value = vIsWSSaleTax
                  Grid.Columns("IsRetailSaleTax").Value = vIsRetailSaleTax
                  Grid.Columns("IsSerial").Value = vIsSerial
                  Grid.Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  Grid.Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  Grid.Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  Grid.Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  Grid.Columns("DiscPack").Value = IIf(Val(TxtDiscPack.Text) = 0, 0, Val(TxtDiscPack.Text))
                  Grid.Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  Grid.Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  Grid.Columns("DiscPer2").Value = IIf(Val(TxtDiscPer2.Text) = 0, 0, Val(TxtDiscPer2.Text))
                  Grid.Columns("DiscVal2").Value = IIf(Val(TxtDiscVal2.Text) = 0, 0, Val(TxtDiscVal2.Text))
                  Grid.Columns("isDiscB4TradeOffer").Value = Abs(ChkDiscB4TradeOffer.Value)
                  Grid.Columns("isDiscB4ExtraScheme").Value = Abs(ChkDiscB4ExtraScheme.Value)
                  Grid.Columns("isDiscB4SaleTax").Value = Abs(ChkDiscB4SaleTax.Value)
                  Grid.Columns("TradeOffer1").Value = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
                  Grid.Columns("TradeOffer2").Value = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
                  Grid.Columns("ExtraSchemePer").Value = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
                  Grid.Columns("TradeValue").Value = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
                  Grid.Columns("ExtraSchemeValue").Value = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
                  Grid.Columns("SC").Value = IIf(Val(TxtSC.Text) = 0, 0, Val(TxtSC.Text))
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("DiscAmount").Value = Val(TxtDiscAmount.Text)
                  Grid.Columns("ExpiryTime").Value = Val(vExpiryTime)
                  Grid.Columns("SaleDiscPer").Value = IIf(Val(TxtSaleDiscPer.Text) = 0, 0, Val(TxtSaleDiscPer.Text))
                  Grid.Columns("RetailAmount").Value = Val(TxtRetailAmount.Text)
                  Grid.Columns("ProfitAmount").Value = Val(TxtProfitAmount.Text)
                  RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
                  RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
                  RsBody!GrossQty = IIf(Val(TxtGrossQty.Text) = 0, Null, Val(TxtGrossQty.Text))
                  RsBody!GrossUnit = IIf(Val(TxtGrossUnit.Text) = 0, Null, Val(TxtGrossUnit.Text))
                  RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
                  RsBody!QtyLoose = Val(TxtQtyLoose.Text)
                  RsBody!Bonus = Val(TxtBonus.Text)
                  RsBody!Price = Val(TxtPrice.Text)
                  RsBody!OldPrice = IIf(vOldPrice = 0, Null, vOldPrice)
                  RsBody!RetailPrice = Val(TxtRetailPrice.Text)
                  RsBody!IsWSDiscb4ST = vIsWSDiscb4ST
                  RsBody!IsWSSaleTax = vIsWSSaleTax
                  RsBody!IsRetailSaleTax = vIsRetailSaleTax
                  RsBody!IsSerial = vIsSerial
                  RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  RsBody!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  RsBody!SaleTaxval = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  RsBody!DiscPack = IIf(Val(TxtDiscPack.Text) = 0, 0, Val(TxtDiscPack.Text))
                  RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  RsBody!DiscPer2 = IIf(Val(TxtDiscPer2.Text) = 0, 0, Val(TxtDiscPer2.Text))
                  RsBody!DiscVal2 = IIf(Val(TxtDiscVal2.Text) = 0, 0, Val(TxtDiscVal2.Text))
                  RsBody!isDiscB4TradeOffer = Abs(ChkDiscB4TradeOffer.Value)
                  RsBody!IsDiscB4ExtraScheme = Abs(ChkDiscB4ExtraScheme.Value)
                  RsBody!isDiscB4SaleTax = Abs(ChkDiscB4SaleTax.Value)
                  RsBody!TradeOffer1 = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
                  RsBody!TradeOffer2 = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
                  RsBody!ExtraSchemePer = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
                  RsBody!TradeValue = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
                  RsBody!ExtraSchemeValue = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
                  RsBody!SC = IIf(Val(TxtSC.Text) = 0, 0, Val(TxtSC.Text))
                  RsBody!Amount = Val(TxtAmount.Text)
                  RsBody!DiscAmount = Val(TxtDiscAmount.Text)
                  RsBody!SaleDiscPer = IIf(Val(TxtSaleDiscPer.Text) = 0, 0, Val(TxtSaleDiscPer.Text))
                  RsBody!RetailAmount = Val(TxtRetailAmount.Text)
                  RsBody!ProfitAmount = Val(TxtProfitAmount.Text)
                  ssql = "Select Productid From Purchasebody where purid=" & Val(TxtPurchaseID.Text) & " and Purchasedate ='" & DtpPurchaseDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmPurchaseInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtPurchaseID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPurchaseDate.Date), "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmPurchaseInvoice, eEdit, TxtPurchaseID.Text, DtpPurchaseDate.DateValue, "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmPurchaseInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtPurchaseID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPurchaseDate.Date), "Pending Update Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
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
'      If Trim(Grid.Columns("ProductID").Text) = "" Then
       If TxtCode.Enabled = True Then
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) + Val(TxtDiscAmount.Text)
         TxtSumDiscAmount.Text = Val(TxtSumDiscAmount.Text) + Val(TxtDiscAmount.Text)
         TxtGrossAmount.Text = Val(TxtGrossAmount.Text) + Val(TxtAmount.Text)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text)))
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmPurchaseInvoice, eAddNewRowByEdit, TxtPurchaseID.Text, DtpPurchaseDate.DateValue, "Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
         Call ActivityLogBin(vRandomID, eFrmPurchaseInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtPurchaseID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPurchaseDate.Date), "Pending Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text) + Val(TxtDiscAmount.Text) - Val(.Columns("DiscAmount").Text)
         TxtSumDiscAmount.Text = Val(TxtSumDiscAmount.Text) + Val(TxtDiscAmount.Text) - Val(.Columns("DiscAmount").Text)
         TxtGrossAmount.Text = Val(TxtGrossAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Val(Grid.Columns("Pack").Value)) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
         ssql = "Select Productid From Purchasebody where purid=" & Val(TxtPurchaseID.Text) & " and Purchasedate ='" & DtpPurchaseDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
         With CN.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmPurchaseInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtPurchaseID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPurchaseDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmPurchaseInvoice, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtPurchaseID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPurchaseDate.Date), "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            Else
               Call ActivityLogBin("", eFrmPurchaseInvoice, eEdit, TxtPurchaseID.Text, DtpPurchaseDate.Date, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmPurchaseInvoice, eEdit, TxtPurchaseID.Text, DtpPurchaseDate.Date, "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmPurchaseInvoice, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtPurchaseID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPurchaseDate.Date), "Pending Update Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
      End If
      .Columns("ProductID").Text = Val(TxtProductID.Text)
      .Columns("Code").Text = TxtCode.Text
      
      .Columns("ColourName").Text = CmbColourName.Text
      If CmbColourName.Text <> "" Then .Columns("ColourID").Value = CmbColourName.ItemData(CmbColourName.ListIndex)
      .Columns("SizeName").Text = cmbSizeName.Text
      If cmbSizeName.Text <> "" Then .Columns("SizeID").Value = cmbSizeName.ItemData(cmbSizeName.ListIndex)
      
      If vColour = True And .Columns("ColourID").Text <> "" Then
         RsBody!ColourID = .Columns("ColourID").Text
         RsBody!SizeID = .Columns("SizeID").Text
      End If
      .Columns("BatchNo").Text = Trim(TxtBatchNo.Text)
      .Columns("ExpiryDate").Text = DtpExpiryDate.DateValue
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
      .Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, "", Val(TxtMultiplier.Text))
      .Columns("GrossQty").Value = IIf(Val(TxtGrossQty.Text) = 0, 0, Val(TxtGrossQty.Text))
      .Columns("GrossUnit").Value = IIf(Val(TxtGrossUnit.Text) = 0, 0, Val(TxtGrossUnit.Text))
      .Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, "", Val(TxtQtyPack.Text))
      .Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
      .Columns("Bonus").Value = Val(TxtBonus.Text)
      .Columns("Price").Value = Val(TxtPrice.Text)
      GetOldPrice
      .Columns("OldPrice").Value = IIf(vOldPrice = 0, "", vOldPrice)
      .Columns("RetailPrice").Value = Val(TxtRetailPrice.Text)
      .Columns("IsWSDiscb4ST").Value = vIsWSDiscb4ST
      .Columns("IsWSSaleTax").Value = vIsWSSaleTax
      .Columns("IsRetailSaleTax").Value = vIsRetailSaleTax
      .Columns("IsSerial").Value = vIsSerial
      .Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      .Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
      .Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
      .Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      .Columns("DiscPack").Value = IIf(Val(TxtDiscPack.Text) = 0, 0, Val(TxtDiscPack.Text))
      .Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      .Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      .Columns("DiscPer2").Value = IIf(Val(TxtDiscPer2.Text) = 0, 0, Val(TxtDiscPer2.Text))
      .Columns("DiscVal2").Value = IIf(Val(TxtDiscVal2.Text) = 0, 0, Val(TxtDiscVal2.Text))
      .Columns("isDiscB4TradeOffer").Value = Abs(ChkDiscB4TradeOffer.Value)
      .Columns("isDiscB4ExtraScheme").Value = Abs(ChkDiscB4ExtraScheme.Value)
      .Columns("isDiscB4SaleTax").Value = Abs(ChkDiscB4SaleTax.Value)
      .Columns("TradeOffer1").Value = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
      .Columns("TradeOffer2").Value = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
      .Columns("ExtraSchemePer").Value = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
      .Columns("TradeValue").Value = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
      .Columns("ExtraSchemeValue").Value = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
      .Columns("SC").Value = IIf(Val(TxtSC.Text) = 0, 0, Val(TxtSC.Text))
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("DiscAmount").Value = Val(TxtDiscAmount.Text)
      .Columns("SaleDiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      .Columns("RetailAmount").Value = Val(TxtRetailAmount.Text)
      .Columns("ProfitAmount").Value = Val(TxtProfitAmount.Text)
      
      ''''' Check Expiry
      vExpiryTime = 0
'         With cn.Execute("Select dbo.GetExpiryTime('" & TxtProductID.Text & "', " & IIf(TxtBatchNo.Text = "", "Null", "'" & TxtBatchNo.Text & "'") & " , GetDate()) as Day ")
'            If .RecordCount > 0 Then
'               vExpiryTime = !Day
'            End If
'         End With
      '''''' GetExpiryTime
         If ObjRegistry.BatchNoVisible = True Then
             ssql = "Select  isnull(min(DATEDIFF (day, getdate(), '" & DtpExpiryDate.DateValue & "')),0)"
            vExpiryTime = CN.Execute(ssql).Fields(0).Value
         End If
         
      .Columns("ExpiryTime").Value = Val(vExpiryTime)
      RsBody!BatchNo = IIf(Trim(TxtBatchNo.Text) = "", Null, Trim(TxtBatchNo.Text))
      RsBody!ExpiryDate = IIf(DtpExpiryDate.DateValue = "", Null, DtpExpiryDate.DateValue)
      RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
      RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
      RsBody!GrossQty = IIf(Val(TxtGrossQty.Text) = 0, Null, Val(TxtGrossQty.Text))
      RsBody!GrossUnit = IIf(Val(TxtGrossUnit.Text) = 0, Null, Val(TxtGrossUnit.Text))
      RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
      RsBody!QtyLoose = Val(TxtQtyLoose.Text)
      RsBody!Bonus = Val(TxtBonus.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!OldPrice = IIf(vOldPrice = 0, Null, vOldPrice)
      RsBody!RetailPrice = Val(TxtRetailPrice.Text)
      RsBody!IsWSDiscb4ST = vIsWSDiscb4ST
      RsBody!IsWSSaleTax = vIsWSSaleTax
      RsBody!IsRetailSaleTax = vIsRetailSaleTax
      RsBody!IsSerial = vIsSerial
      RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      RsBody!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
      RsBody!SaleTaxval = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
      RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      RsBody!DiscPack = IIf(Val(TxtDiscPack.Text) = 0, 0, Val(TxtDiscPack.Text))
      RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      RsBody!DiscPer2 = IIf(Val(TxtDiscPer2.Text) = 0, 0, Val(TxtDiscPer2.Text))
      RsBody!DiscVal2 = IIf(Val(TxtDiscVal2.Text) = 0, 0, Val(TxtDiscVal2.Text))
      RsBody!isDiscB4TradeOffer = Abs(ChkDiscB4TradeOffer.Value)
      RsBody!IsDiscB4ExtraScheme = Abs(ChkDiscB4ExtraScheme.Value)
      RsBody!isDiscB4SaleTax = Abs(ChkDiscB4SaleTax.Value)
      RsBody!TradeOffer1 = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
      RsBody!TradeOffer2 = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
      RsBody!ExtraSchemePer = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
      RsBody!TradeValue = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
      RsBody!ExtraSchemeValue = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
      RsBody!SC = IIf(Val(TxtSC.Text) = 0, 0, Val(TxtSC.Text))
      RsBody!Amount = Val(TxtAmount.Text)
      RsBody!DiscAmount = Val(TxtDiscAmount.Text)
      RsBody!SaleDiscPer = IIf(Val(TxtSaleDiscPer.Text) = 0, 0, Val(TxtSaleDiscPer.Text))
      RsBody!RetailAmount = Val(TxtRetailAmount.Text)
      RsBody!ProfitAmount = Val(TxtProfitAmount.Text)
     If TxtCode.Enabled = False And ObjRegistry.AfterRowEditFocusNextGridLine = True Then
         .MoveNext
         Call Grid_GotFocus
'         CmbPackName.SetFocus
      Else
         .MoveLast
         If Trim(.Columns("ProductID").Text) <> "" Then
            .AllowAddNew = True
            .AddNew
            .Columns("Code").Text = " "
            .AllowAddNew = False
         End If
      End If
   End With
   
   QtyOffer = 0
   GetDataFromTextBoxesToGridOffer
   
'   If TxtCode.Enabled = True And TxtCode.Visible = True Then TxtCode.SetFocus
'   If Trim(Grid.Columns("ProductID").Text) = "" Then
'      Call SubClearDetailArea
'      If TxtCode.Enabled = True And TxtCode.Visible = True Then TxtCode.SetFocus
'   End If
   
   GetDataBackFromGridToTexBoxes
   Grid_LostFocus
   
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
   TxtBarcode.Text = ""
   TxtProductName.Text = ""
   CmbPackName.ListIndex = 0
   TxtBatchNo.Text = ""
   DtpExpiryDate.DateValue = ""
   TxtMultiplier.Text = ""
   TxtGrossQty.Text = ""
   TxtGrossUnit.Text = ""
   TxtQtyPack.Text = ""
   TxtQtyLoose.Text = ""
   TxtBonus.Text = ""
   TxtPrice.Text = ""
   TxtRetailPrice.Text = ""
   TxtOffer.Text = ""
   TxtSaleTaxPer.Text = ""
   TxtSaleTaxVal.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtTradeOffer1.Text = ""
   TxtTradeOffer2.Text = ""
   TxtExtraSchemePer.Text = ""
   TxtTradeOfferValue.Text = ""
   TxtExtraSchemeValue.Text = ""
   ChkDiscB4TradeOffer.Value = 0
   ChkDiscB4ExtraScheme.Value = 0
   ChkDiscB4SaleTax.Value = 0
   TxtDiscVal.Text = ""
   TxtSC.Text = ""
   TxtAmount.Text = ""
   TxtDiscAmount.Text = ""
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
                        GridOffer.Columns("ProductID").Text = Val(TxtProductID.Text)
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
    RsProductOffer.Open "Select * from PurchaseBodyOffer where PurId =" & Val(TxtPurchaseID.Text) & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
    If RsProductOffer.RecordCount > 0 Then
    GridOffer.Visible = True
    ssql = "select p.productname, D.* from PurchaseBodyOffer D Inner join products p on p.productid = D.productOfferid where PurId =" & Val(TxtPurchaseID.Text) & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
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

Private Sub PopulatePOToGridOffer()
    If RsProductOffer.State = adStateOpen Then RsProductOffer.Close
    RsProductOffer.Open "Select * from PurchaseBodyOffer where PurId =" & Val(TxtPurchaseID.Text) & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
'    If RsProductOffer.RecordCount > 0 Then
'    GridOffer.Visible = True
    ssql = "select p.productname, D.* from PurchaseOrderBodyOffer D Inner join products p on p.productid = D.productOfferid where OrderID =" & Val(TxtOrderID.Text) & " And OrderDate = '" & DtpOrderDate.DateValue & "'"
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
            
            RsProductOffer.AddNew
            RsProductOffer!Productid = !Productid
            RsProductOffer!ProductOfferID = !ProductOfferID
            RsProductOffer!QtyOffer = !QtyOffer
            RsProductOffer.Update
            .MoveNext
         Wend
      
     If .RecordCount > 0 Then
        GridOffer.AddNew
        GridOffer.Columns("ProductID").Text = " "
        GridOffer.AllowAddNew = False
        GridOffer.Redraw = True
        GridOffer.Visible = True
     Else
        GridOffer.CancelUpdate
        GridOffer.RemoveAll
        GridOffer.AddNew
        GridOffer.Columns("ProductID").Text = " "
        GridOffer.Update
        GridOffer.Visible = False
     End If
     End With
'    End If
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler

   With Grid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("Code").Text
      TxtBatchNo.Text = .Columns("BatchNo").Text
      DtpExpiryDate.DateValue = .Columns("ExpiryDate").Text
      TxtProductName.Text = .Columns("ProductName").Text
      CmbPackName.Clear
      vStrSQL = "select distinct pp.PackingID, Packingname from ProductPacking pp inner join packings p on p.packingid = pp.packingid" & vbCrLf _
           + "left outer join ProductBarcodes b on b.productid = pp.productid" & vbCrLf _
           + " where pp.productid = " & Val(TxtCode.Text) & " or code = '" & TxtCode.Text & "'"
      With CN.Execute(vStrSQL)
         CmbPackName.AddItem ""
         While Not .EOF
            CmbPackName.AddItem !PackingName
            CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
            .MoveNext
         Wend
            .Close
      End With
      If Trim(.Columns("PackName").Text) = "" Then
         CmbPackName.ListIndex = 0
      Else
         CmbPackName.Text = .Columns("PackName").Text
      End If
      
      TxtMultiplier.Text = .Columns("Pack").Text
      TxtGrossQty.Text = .Columns("GrossQty").Text
      TxtGrossUnit.Text = .Columns("GrossUnit").Text
      
      
      
      If .Columns("QtyLoose").Text = "" Then
         TxtQtyLoose.Text = .Columns("QtyLoose").Text
      Else
         TxtQtyLoose.Text = Abs(.Columns("QtyLoose").Text)
      End If
      
      If .Columns("QtyPack").Text = "" Then
         TxtQtyPack.Text = .Columns("QtyPack").Text
      Else
         TxtQtyPack.Text = Abs(.Columns("QtyPack").Text)
      End If
      
      vUnitPrice = Val(.Columns("Price").Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))
'      vUnitRetailPrice = Val(.Columns("RetailPrice").Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))
      vUnitRetailPrice = Val(.Columns("RetailPrice").Text)
      
      TxtPrice.Text = .Columns("Price").Text
      TxtRetailPrice.Text = .Columns("RetailPrice").Text
      vIsWSDiscb4ST = .Columns("IsWSDiscb4ST").Value
      vIsRetailSaleTax = .Columns("IsRetailSaleTax").Value
      vIsRetailSaleTax = .Columns("IsRetailSaleTax").Value
      vIsSerial = .Columns("IsSerial").Value
      TxtBonus.Text = .Columns("Bonus").Text
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtDiscPack.Text = .Columns("DiscPack").Value
      vDiscPackFlag = IIf(TxtDiscPack.Visible, Val(TxtDiscPack.Text) <> 0, False)
      TxtOffer.Text = .Columns("Offer").Value
      TxtSaleTaxPer.Text = .Columns("SaleTaxPer").Value
      TxtSaleTaxVal.Text = .Columns("SaleTaxVal").Value
      ChkDiscB4TradeOffer = Abs(Val(.Columns("isDiscB4TradeOffer").Value))
      ChkDiscB4ExtraScheme = Abs(Val(.Columns("isDiscB4ExtraScheme").Value))
      ChkDiscB4SaleTax = Abs(Val(.Columns("isDiscB4SaleTax").Value))
      TxtTradeOffer1.Text = .Columns("TradeOffer1").Value
      TxtTradeOffer2.Text = .Columns("TradeOffer2").Value
      TxtExtraSchemePer.Text = .Columns("ExtraSchemePer").Value
      TxtTradeOfferValue.Text = .Columns("TradeValue").Value
      TxtExtraSchemeValue.Text = .Columns("ExtraSchemeValue").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscPack.Text = .Columns("DiscPack").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      TxtDiscPer2.Text = .Columns("DiscPer2").Value
      TxtDiscVal2.Text = .Columns("DiscVal2").Value
      TxtSC.Text = .Columns("SC").Value
      TxtAmount.Text = Abs(.Columns("Amount").Value)
      TxtDiscAmount.Text = Abs(Val(.Columns("DiscAmount").Value))
      TxtSaleDiscPer.Text = .Columns("SaleDiscPer").Value
      TxtRetailAmount.Text = Abs(Val(.Columns("RetailAmount").Value))
      TxtProfitAmount.Text = Abs(Val(.Columns("ProfitAmount").Value))
      If ObjRegistry.ShowAllPrices Then
         PopulateDataToPriceGrid
         FrmProductPrices.Visible = True
      End If
            
      If TxtCode.Enabled = False Then
         If LblStock.Visible = False Then
            LblStock.Visible = vShowStock
            LblAllStock.Visible = vShowStock
            LblStockCaption.Visible = vShowStock
         End If
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
        vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & Date + 1 & "',0),0)"
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
         End If
        
         If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & ",Null,0,0,0,0,0,0,'" & DtpPurchaseDate.DateValue + 1 & "',0),0)"
            With CN.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
            LblAllStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
            With CN.Execute("Select isnull(abbreviation,'') from packings where packingname = '" & CmbPackName.Text & "'")
               If .RecordCount > 0 Then
                  LblAllStock.Caption = LblAllStock.Caption & " " & .Fields(0).Value
               Else
                  LblAllStock.Caption = LblAllStock.Caption & " "
               End If
            End With
            LblAllStock.Caption = LblAllStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & "Loose"
            LblAllStock.Visible = True
            LblStock.Visible = Not LblAllStock.Visible
         Else
            LblAllStock.Visible = False
            LblStock.Visible = True
         End If
'      With CN.Execute("select QtyLoose from currentstockStore where productid ='" & TxtProductID.Text & "' and storeid = " & TxtStoreID.Text)
'         If .RecordCount > 0 Then
''            vQtyLoose = !QtyLoose
'            LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack('" & TxtProductID.Text & "',Floor(" & !QtyLoose & "))").Fields(0).Value
'            LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'            LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & !QtyLoose & "))").Fields(0).Value
'            LblStock.Caption = LblStock.Caption & " " & "Loose"
'
'            'LblStock.Caption = !QtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtProductID.Text & "')").Fields(0).Value
'         Else
''            vQtyLoose = 0
'            LblStock.Caption = 0
'         End If
'      End With
      If ObjRegistry.isShowListPrice Then
         If Trim(TxtProductID.Text) <> "" Then
             ssql = "select top 3 h.purID, pt.PartyName, VendorID, code, b.* " & vbCrLf & _
            " from PurchaseHeader h inner join Purchasebody b on h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate" & vbCrLf & _
            " inner join Parties pt on pt.PartyID = h.VendorID " & vbCrLf & _
            " where b.productid = " & Val(TxtProductID.Text) & " and h.purchasedate < '" & DtpPurchaseDate.DateValue & "' order by b.PurchaseDate Desc, b.purid Desc"
            With CN.Execute(ssql)
               If Not .EOF Then
                  LblCaptionRetailPrice = "Last Price"
                  LblRetailPrice.Caption = !Price & ", DiscPack=" & Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * Val(IIf(IsNull(!Multiplier), "1", !Multiplier)) & vbCrLf & _
                  "Disc%=" & !DiscPer & ", " & Format(!PurchaseDate, "dd/MM/yyyy")
               Else
                  LblRetailPrice.Caption = ""
               End If
            End With
         Else
            LblRetailPrice.Caption = ""
         End If
      Else
         If Trim(TxtProductID.Text) <> "" Then
            LblCaptionRetailPrice = "Retail Price"
            LblRetailPrice.Caption = CN.Execute("Select RetailPrice from Products where ProductID = " & Val(TxtProductID.Text)).Fields(0).Value
         End If
      End If
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetPurchase()
   On Error GoTo ErrorHandler
   CmbPackName.Clear
   With CN.Execute("Select * from Packings")
      CmbPackName.AddItem ""
      While Not .EOF
         CmbPackName.AddItem !PackingName
         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
         .MoveNext
      Wend
      .Close
   End With
   
   ssql = "Select h.*, OrganizationName, isnull(p.partyname,EmpName) partyname, isnull(p.Address,emp.Address) Address , isnull(p.City,emp.city) City, StoreName FROM PurchaseHeader h Left Outer join parties p on h.Vendorid = p.partyid Left outer join Employees Emp on Emp.EmpID = H.VendorID inner join stores s on s.storeid = h.storeid left outer join Organizations o on o.OrganizationID = h.OrganizationID where h.SID=" & Val(TxtSID.Text) & " and PurchaseDate='" & DtpPurchaseDate.DateValue & "'" & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With CN.Execute(ssql)
      If Not .BOF Then
          DtpPurchaseDate.DateValue = !PurchaseDate
          DtpPromiseDate.DateValue = !PromiseDate
          TxtOrderID.Text = IIf(IsNull(!OrderID), "", !OrderID)
          DtpOrderDate.DateValue = IIf(IsNull(!OrderDate), "01/01/1990", !OrderDate)
          DtpEntryDate.DateValue = IIf(IsNull(!EntryDate), !PurchaseDate, !EntryDate)
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
          TxtTotalAmount.Text = !TotalAmount + IIf(IsNull(!SumDiscAmount), 0, !SumDiscAmount)
          TxtGrossAmount.Text = !TotalAmount
          TxtSumDiscAmount.Text = !SumDiscAmount
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          TxtAdvTaxVal.Text = IIf(IsNull(!AdvTaxVal), "", !AdvTaxVal)
          TxtAdvTaxPer.Text = IIf(IsNull(!AdvTaxPer), "", !AdvTaxPer)
          TxtExtraTaxVal.Text = IIf(IsNull(!ExtraTaxVal), "", !ExtraTaxVal)
          TxtExtraTaxPer.Text = IIf(IsNull(!ExtraTaxPer), "", !ExtraTaxPer)
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtTotalExpense.Text = IIf(IsNull(!TotalExpense), "", !TotalExpense)
          TxtFreight.Text = IIf(IsNull(!Freight), "", !Freight)
          OptVender.Value = IIf((IsNull(!IsVenderFreight) = True) Or (!IsVenderFreight = False), 0, 1)
          OptExpense.Value = IIf((IsNull(!IsExpense) = True) Or (!IsExpense = False), 0, 1)
          OptMe.Value = IIf((IsNull(!IsOurFreight) = True) Or (!IsOurFreight = False), 0, 1)
          TxtPaidAmount.Text = IIf(IsNull(!PAIDAMOUNT), "", !PAIDAMOUNT)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
          TxtPaidAmount.Text = IIf(IsNull(!PAIDAMOUNT), "", !PAIDAMOUNT)
          
'          TxtPreviousPayable.Text = CN.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtVenderID.Text & "','" & DtpPurchaseDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
'          VStrSQL = " Select isnull(Sum(TotalAmount - isnull(BillDisc,0) + isnull(OtherCharges,0)),0) as Amount " & vbCrLf _
'                  + " FROM PurchaseHeader h INNER JOIN (Select PurId, PurchaseDate, Sum(amount) TTLValue FROM PurchaseBody Group By PurId, PurchaseDate)B " & vbCrLf _
'                  + " ON h.PurId = B.PurId and h.PurchaseDate = B.PurchaseDate " & vbCrLf _
'                  + " where VendorID = '" & (TxtVenderID.Text) & "' and h.PurchaseDate = '" & DtpPurchaseDate.DateValue & "' and h.PurID >= " & Val(TxtPurchaseID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
'          TxtPreviousPayable.Text = TxtPreviousPayable.Text + CN.Execute(VStrSQL).Fields(0).Value
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

Private Sub GetPurchaseOrder()
   On Error GoTo ErrorHandler
   TxtPurchaseID.Text = FunGetMaxID
   ssql = "Select h.*, OrganizationName, p.partyname, Address, City, StoreName FROM PurchaseOrderHeader h join parties p on h.Vendorid = p.partyid inner join stores s on s.storeid = h.storeid left outer join Organizations o on o.OrganizationID = h.OrganizationID where h.OrderID=" & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'"
   With CN.Execute(ssql)
      If Not .BOF Then
          DtpOrderDate.DateValue = !OrderDate
'          DtpEntryDate.DateValue = IIf(IsNull(!EntryDate), !OrderDate, !EntryDate)
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
          TxtGrossAmount.Text = !TotalAmount
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtTotalExpense.Text = IIf(IsNull(!TotalExpense), "", !TotalExpense)
          TxtPaidAmount.Text = IIf(IsNull(!PAIDAMOUNT), "", !PAIDAMOUNT)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtPaidAmount.Text = IIf(IsNull(!PAIDAMOUNT), "", !PAIDAMOUNT)
          TxtPreviousPayable.Text = IIf(IsNull(!PreviousAmount), "", !PreviousAmount)
          lblPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Previous Receivable", "Previous Payable")
          LblTtlPayable.Caption = IIf(Val(TxtPreviousPayable.Text) > 0, "Total Receivable", "Total Payable")
          TxtPreviousPayable.Text = Abs(Val(TxtPreviousPayable.Text))
      End If
      .Close
   End With
   Call PopulatePOToGrid
'   FormStatus = OpenMode
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

Private Sub GridSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub GridSerial_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(GridSerial.Columns("Serial").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub GridSerial_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    GetDataBackFromGridSerialToTexBoxes
End Sub

Private Sub TxtBillDisc_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtBillDisc.Name Then Exit Sub
   TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / IIf(Val(TxtGrossAmount.Text) = 0, 1, Val(TxtGrossAmount.Text)), 6)
   Call SubCalculateFooter
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtBillDiscPer_Change()
   If ActiveControl.Name <> TxtBillDiscPer.Name Then Exit Sub
   TxtBillDisc.Text = SelfRound((Val(TxtGrossAmount.Text) * Val(TxtBillDiscPer.Text) / 100))
   Call SubCalculateFooter
End Sub

Private Sub TxtDiscPack_Change()
   If ActiveControl.Name <> TxtDiscPack.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   TxtDiscPC.Text = Round(IIf(Val(TxtMultiplier.Text) <> 0, Val(TxtDiscPack.Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text)), ""), IIf(ObjRegistry.IsRoundFigure, 2, 5))
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, IIf(ObjRegistry.IsRoundFigure, 3, 5))
   If Val(TxtDiscPer.Text) = 0 Then TxtDiscPer.Text = ""
   'If Val(TxtDiscPC.Text) = 0 Then TxtDiscPC.Text = ""
   Call SubCalculateBody
End Sub

Private Sub TxtDiscPC_Change()
   If ActiveControl.Name <> TxtDiscPC.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   If TxtDiscPack.Visible Then TxtDiscPack.Text = IIf(Val(TxtMultiplier.Text) <> 0, Val(TxtDiscPC.Text) * Val(TxtMultiplier.Text), Val(TxtDiscPC.Text))
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, IIf(ObjRegistry.IsRoundFigure, 3, 5))
   If Val(TxtDiscPer.Text) = 0 Then TxtDiscPer.Text = ""
   Call SubCalculateBody
End Sub

Private Sub TxtDiscPer_Change()
   If ActiveControl.Name <> TxtDiscPer.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   TxtDiscPC.Text = Round((vUnitPrice * Val(TxtDiscPer.Text) / 100), IIf(ObjRegistry.IsRoundFigure, 2, 5))
   If TxtDiscPack.Visible Then TxtDiscPack.Text = IIf(Val(TxtMultiplier.Text) <> 0, Val(TxtDiscPC.Text) * Val(TxtMultiplier.Text), Val(TxtDiscPC.Text))
   If Val(TxtDiscPC.Text) = 0 Then TxtDiscPC.Text = ""
   Call SubCalculateBody
End Sub

Private Sub TxtDiscVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDiscVal.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   If (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) = 0 Then Exit Sub
   TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 4)
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
'   TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), 2)
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDiscVal_LostFocus()
'   Select Case ActiveControl.Name
'   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtBonus.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtOffer.Name, TxtSaleTaxPer.Name, TxtSC.Name
'      Exit Sub
'   End Select
'   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtMultiplier_Change()
   If ActiveControl.Name <> TxtMultiplier.Name Then Exit Sub
   If Val(TxtMultiplier.Text) <> 0 Then
      TxtPrice.Text = Round(vUnitPrice * Val(TxtMultiplier.Text), 3)
   Else
      TxtPrice.Text = Round(vUnitPrice, 3)
   End If
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub TxtMultiplier_Validate(Cancel As Boolean)
   If ActiveControl.Name <> TxtQtyPack.Name Then Exit Sub
End Sub

Private Sub TxtOffer_Change()
    If ActiveControl.Name <> TxtOffer.Name Then Exit Sub
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

Private Sub TxtProductName_Change()
   TxtProductName2.Text = TxtProductName.Text
End Sub

Private Sub TxtQtyLoose_Change()
   If ActiveControl.Name <> TxtQtyLoose.Name Then Exit Sub
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub TxtQtyLoose_Validate(Cancel As Boolean)
   If ActiveControl.Name <> TxtQtyLoose.Name Then Exit Sub
End Sub

Private Sub TxtQtyPack_Change()
   If ActiveControl.Name <> TxtQtyPack.Name Then Exit Sub
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub TxtRetailPrice_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtRetailPrice.Name Then Exit Sub
   vUnitRetailPrice = Val(TxtRetailPrice.Text)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSaleTaxPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSaleTaxPer.Name Then Exit Sub
      If vIsWSSaleTax = True And vIsWSDiscb4ST = True Then
          TxtSaleTaxVal.Text = Round(Val(TxtAmount.Text) * Val(TxtSaleTaxPer.Text) / 100, 3)
   '   ElseIf vIsWSSaleTax = True And vIsWSDiscb4ST = False Then
   '       TxtSaleTaxVal.Text = Round(Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
      ElseIf vIsRetailSaleTax = True Then
          TxtSaleTaxVal.Text = Round((vUnitRetailPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) - (Val(vUnitRetailPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) / (100 + Val(TxtSaleTaxPer.Text)) * 100), 3)
      Else
          TxtSaleTaxVal.Text = Round(Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
      End If
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSaleTaxVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSaleTaxVal.Name Then Exit Sub
   If vIsWSSaleTax = True And vIsWSDiscb4ST = True Then
       TxtSaleTaxPer.Text = Round(Val(TxtSaleTaxVal.Text) * 100 / Round(((Val(vUnitPrice) - Val(TxtDiscPC.Text)) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))), 2), 3)
   ElseIf vIsRetailSaleTax = True Then
      If Val(TxtSaleTaxVal.Text) <> 0 Then
        TxtSaleTaxPer.Text = Round(Val(TxtSaleTaxVal.Text) / ((vUnitRetailPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) - Val(TxtSaleTaxVal.Text)) * 100, 2)
       End If
   Else
       TxtSaleTaxPer.Text = Round(Val(TxtSaleTaxVal.Text) * 100 / (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))), 3)
   End If
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSC_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSC.Name Then Exit Sub
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSC_LostFocus()
   On Error GoTo ErrorHandler
   Select Case ActiveControl.Name
   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtBonus.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtOffer.Name, TxtSaleTaxPer.Name, TxtSaleTaxVal.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSerial_LostFocus()
   On Error GoTo ErrorHandler
   GetDataFromTexBoxesToGridSerial
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGridSerial()
   On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   If Trim(TxtCode.Text) = "" Then
      If TxtCode.Enabled = True Then TxtCode.SetFocus
      TxtSerial.Text = ""
      Exit Sub
   End If
     
   If Trim(TxtSerial.Text) = "" Then
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
'      TxtSerial.SetFocus
      Exit Sub
   End If
   
   
'   vStrSQL = "Select Distinct ProductID from vuPurchaseSerial where SerialAdd = 1 and Serial = '" & Trim(TxtSerial.Text) & "'"
   vStrSQL = "Select ProductID, SerialAdd from vuPurchaseSerial where Serial = '" & Trim(TxtSerial.Text) & "' Order by SerialAdd Desc"
   With CN.Execute(vStrSQL)
      If Not .EOF Then
         If !SerialAdd = True Then
            MsgBox "The Serial cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
            TxtSerial.SetFocus
            Exit Sub
          ElseIf !Productid = Val(TxtProductID.Text) Then
            MsgBox "Same Serial cannot be inserted on Same Product Again", vbInformation + vbOKOnly, "Error"
            TxtSerial.SetFocus
            Exit Sub
         End If
      End If
   End With
   RsBodySerial.Filter = ""
   RsBodySerial.Filter = "Serial='" & Trim(TxtSerial.Text) & "'"
   If TxtSerial.Enabled Then
      If RsBodySerial.RecordCount = 0 Then
         RsBodySerial.AddNew
         GridSerial.Columns("ProductID").Text = TxtCode.Text
         GridSerial.Columns("Serial").Text = Trim(TxtSerial.Text)
         RsBodySerial!Productid = TxtCode.Text
         RsBodySerial!serial = Trim(TxtSerial.Text)
         RsBodySerial.Update
         TxtSerial.Text = ""
      Else
'         GridSerial.Redraw = False
'         GridSerial.MoveFirst
'            For vrowcounter = 1 To GridSerial.Rows
'               If GridSerial.Columns("Serial").Text = TxtSerial.Text Then
'                  MsgBox "The Serial cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
'                  'SubClearDetailArea
'                  GridSerial.MoveLast
'                  TxtSerial.SetFocus
'                  GridSerial.Redraw = True
'                  Exit Sub
'               End If
'               GridSerial.MoveNext
'            Next vrowcounter
         MsgBox "The Serial Already Exist", vbInformation + vbOKOnly, "Alert"
         
         
         
'         GridSerial.MoveLast
         RsBodySerial.Filter = "ProductID = " & Val(TxtCode.Text)
         If TxtQtyLoose.Text = RsBodySerial.RecordCount And vIsSerial = True Then GetDataFromTexBoxesToGrid
         TxtSerial.Text = ""
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
   
   
   RsBodySerial.Filter = "ProductID = " & Val(TxtCode.Text)
   If Val(TxtQtyLoose.Text) = RsBodySerial.RecordCount And vIsSerial = True Then
      GetDataFromTexBoxesToGrid
      Exit Sub
   End If
   '' automove grid by enter serial No
   If ObjRegistry.AutoMoveGridWhenSerialEntered = True Then
      If Grid.Rows = Grid.Row + 1 Then Exit Sub
      Grid.MoveNext
      Call Grid_GotFocus
   End If
   GridSerial.Redraw = True
   Call GridSerial_DblClick
   If TxtSerial.Enabled = True Then TxtSerial.SetFocus
   
   
   Exit Sub
ErrorHandler:
   GridSerial.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtTerms_Change()
On Error GoTo ErrorHandler
   If TxtTerms.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtTerms.Name Then Exit Sub
   DtpPromiseDate.DateValue = DateAdd("d", Val(TxtTerms.Text), DtpPurchaseDate.DateValue)
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub TxtGrossAmount_Change()
   TxtBillDisc.Text = SelfRound((Val(TxtGrossAmount.Text) * Val(TxtBillDiscPer.Text) / 100))
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

Private Sub FindRebate()
   On Error GoTo ErrorHandler
'    With CN.Execute("Select * from ProductOffers where Rebate <> 0 and ProductID = '" & TxtProductID.Text & "'")
'        If .RecordCount > 0 Then
'            Rebate = Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)
'            Rebate = Rebate \ !Qty
'            Rebate = Rebate * !Rebate
'            TxtOffer.Text = Rebate
'        End If
'    End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub UserActivities()
     If vIsNewRecord = False Then
    With CN.Execute("Select  * from PurchaseHeader where PurID =" & TxtPurchaseID.Text & " And PurchaseDate = '" & DtpPurchaseDate.DateValue & "'")
'        If Val(TxtVenderID.Text) <> IIf(IsNull(!VENDORID), 0, !VENDORID) Then
'            CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Updated VenderID-" & !VENDORID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
'        End If
'        If TxtStoreID.Text <> !StoreID Then
'            CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
'        End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.Rows - 1
        With CN.Execute("Select * from PurchaseBody Where PurID = " & TxtPurchaseID.Text & " and PurchaseDate ='" & DtpPurchaseDate.DateValue & "' and Productid = " & Val(Grid.Columns("Productid").Text))
             If .EOF = True Then
                ssql = "Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")"
                CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("QtyLoose").Text <> !QtyLoose Or Grid.Columns("Price").Text <> !Price Or Grid.Columns("discper").Text <> !DiscPer Then
                   CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Updated ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
    
   Else
    CN.Execute ("Insert Into UserActivities values ('Purchase Invoice'" & "," & TxtPurchaseID.Text & ",'" & DtpPurchaseDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub CalculateValue()
On Error GoTo ErrorHandler

'''''' Calculate Trade value ''''''
If Val(TxtTradeOffer1.Text) <> 0 And Val(TxtTradeOffer2.Text) <> 0 Then
   If ChkDiscB4TradeOffer.Value = 1 Then
      TxtTradeOfferValue.Text = Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))
      TxtTradeOfferValue.Text = Round((Val(TxtTradeOffer2.Text) * Val(TxtTradeOfferValue.Text)) / (Val(TxtTradeOffer1.Text) + Val(TxtTradeOffer2.Text)), 4)
   Else
      TxtTradeOfferValue.Text = (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - Val(TxtDiscVal.Text)
      TxtTradeOfferValue.Text = Round((Val(TxtTradeOffer2.Text) * Val(TxtTradeOfferValue.Text)) / (Val(TxtTradeOffer1.Text) + Val(TxtTradeOffer2.Text)), 4)
   End If
Else
   TxtTradeOfferValue.Text = 0
End If
'''''''''''''''''''''''''''''''''


'''''' Calculate Extra Scheme value ''''''
If Val(TxtExtraSchemePer.Text) <> 0 Then
   If ChkDiscB4ExtraScheme.Value = 1 Then
      TxtExtraSchemeValue.Text = (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)))
'      TxtExtraSchemeValue.Text = Val(TxtExtraSchemeValue.Text) - Val(TxtTradeOfferValue.Text)
      TxtExtraSchemeValue.Text = Val(TxtExtraSchemeValue.Text) * Val(TxtExtraSchemePer.Text) / 100
   Else
      TxtExtraSchemeValue.Text = (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)))
      TxtExtraSchemeValue.Text = Val(TxtExtraSchemeValue.Text) - Val(TxtTradeOfferValue.Text) - Val(TxtDiscVal.Text)
      TxtExtraSchemeValue.Text = Val(TxtExtraSchemeValue.Text) * Val(TxtExtraSchemePer.Text) / 100
   End If
Else
   TxtExtraSchemeValue.Text = 0
End If
'''''''''''''''''''''''''''''''''


'''''' Calculate GST value ''''''
'If Val(TxtSaleTaxPer.Text) <> 0 Then
'   If ChkDiscB4SaleTax.Value = 1 Then
'      TxtSaleTaxVal.Text = (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)))
'      TxtSaleTaxVal.Text = Val(TxtSaleTaxVal.Text) * Val(TxtSaleTaxPer.Text) / 100
'   Else
'      TxtSaleTaxVal.Text = (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)))
'      TxtSaleTaxVal.Text = Val(TxtSaleTaxVal.Text) - Val(TxtDiscVal.Text) - Val(TxtTradeOfferValue.Text) - Val(TxtExtraSchemeValue.Text)
'      TxtSaleTaxVal.Text = Round(Val(TxtSaleTaxVal.Text) * Val(TxtSaleTaxPer.Text) / 100, 4)
'   End If
''Else
''   TxtSaleTaxPer.Text = 0
'End If
'''''''''''''''''''''''''''''''''

Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTradeOffer1_Change()
   If ActiveControl.Name <> TxtTradeOffer1.Name Then Exit Sub
   Call CalculateValue
End Sub

Private Sub TxtTradeOffer2_Change()
   If ActiveControl.Name <> TxtTradeOffer2.Name Then Exit Sub
   Call CalculateValue
End Sub

Private Sub TxtExtraSchemePer_Change()
   If ActiveControl.Name <> TxtExtraSchemePer.Name Then Exit Sub
   Call CalculateValue
End Sub

Private Sub ChkDiscB4ExtraScheme_Click()
   If ActiveControl.Name <> ChkDiscB4ExtraScheme.Name Then Exit Sub
   Call CalculateValue
End Sub

Private Sub ChkDiscB4SaleTax_Click()
   If ActiveControl.Name <> ChkDiscB4SaleTax.Name Then Exit Sub
   Call CalculateValue
End Sub

Private Sub ChkDiscB4TradeOffer_Click()
   If ActiveControl.Name <> ChkDiscB4TradeOffer.Name Then Exit Sub
   Call CalculateValue
End Sub

Private Sub TxtAdvTaxPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtAdvTaxPer.Name Then Exit Sub
'   vAdvDiscPerFlag = False
   TxtNetAmount.Text = SelfRound(Val(TxtGrossAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text)
   TxtAdvTaxVal.Text = SelfRound((Val(TxtNetAmount.Text) * Val(TxtAdvTaxPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtAdvTaxVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtAdvTaxVal.Name Then Exit Sub
'   vAdvDiscPerFlag = True
   TxtNetAmount.Text = SelfRound(Val(TxtGrossAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text)
   TxtAdvTaxPer.Text = Round((Val(TxtAdvTaxVal.Text) * 100) / IIf(Val(TxtNetAmount.Text) = 0, 1, Val(TxtNetAmount.Text)), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetOldPrice()
   On Error GoTo ErrorHandler
      vOldPrice = 0
      
      ssql = "select top 1 b.Price  " & vbCrLf & _
      " from PurchaseHeader h inner join Purchasebody b on h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate" & vbCrLf & _
      " inner join Parties pt on pt.PartyID = h.VendorID " & vbCrLf & _
      " where b.productid = " & Val((Grid.Columns("ProductID").Text)) & " order by b.PurchaseDate Desc, b.purid Desc"
      With CN.Execute(ssql)
      If Not .EOF Then
           If Grid.Columns("Price").Value <> .Fields("Price").Value Then vOldPrice = .Fields("Price").Value
      End If
      End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.PurchaseHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmPurchaseInvoice) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmPurchaseInvoice & ", " & vUser & "," & TableHeaderFields(eFrmPurchaseInvoice) & " from PurchaseHeader " & vbCrLf _
             & "Where PurID = " & TxtPurchaseID.Text & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
      CN.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.PurchaseBodyBin (" & TableBodyFields(eFrmPurchaseInvoice) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmPurchaseInvoice) & " from PurchaseBody " & vbCrLf _
             & "Where PurID = " & TxtPurchaseID.Text & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
      CN.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

