VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmSaleInvoiceDist 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15345
   Icon            =   "FrmSaleInvoiceDist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   684
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1023
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkIsPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Print"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11880
      TabIndex        =   257
      Top             =   9000
      Width           =   1290
   End
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10800
      TabIndex        =   255
      Top             =   9000
      Width           =   1245
   End
   Begin VB.Frame FrmHistory 
      Height          =   1635
      Left            =   2115
      TabIndex        =   171
      Top             =   5925
      Visible         =   0   'False
      Width           =   10125
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridHistory 
         Height          =   1455
         Left            =   90
         TabIndex        =   172
         Top             =   135
         Width           =   9885
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
         stylesets(0).Picture=   "FrmSaleInvoiceDist.frx":000C
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
         stylesets(1).Picture=   "FrmSaleInvoiceDist.frx":0028
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
         stylesets(2).Picture=   "FrmSaleInvoiceDist.frx":0044
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
         Columns(0).Width=   953
         Columns(0).Caption=   "CID"
         Columns(0).Name =   "ID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3281
         Columns(1).Caption=   "Customer Name"
         Columns(1).Name =   "Name"
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   1429
         Columns(2).Caption=   "BillID"
         Columns(2).Name =   "BillID"
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
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
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
         _ExtentX        =   17436
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
      Left            =   8490
      TabIndex        =   124
      Top             =   5400
      Visible         =   0   'False
      Width           =   4215
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridExpense 
         Height          =   1860
         Left            =   120
         TabIndex        =   125
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
         stylesets(0).Picture=   "FrmSaleInvoiceDist.frx":0060
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
         stylesets(1).Picture=   "FrmSaleInvoiceDist.frx":007C
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
         stylesets(2).Picture=   "FrmSaleInvoiceDist.frx":0098
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
         TabIndex        =   127
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
         TabIndex        =   128
         Top             =   2100
         Width           =   1020
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridOffer 
      Height          =   1365
      Left            =   1215
      TabIndex        =   100
      Top             =   6480
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
      stylesets(0).Picture=   "FrmSaleInvoiceDist.frx":00B4
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
   Begin VB.Frame FrmProductPrices 
      Height          =   1095
      Left            =   7515
      TabIndex        =   221
      Top             =   150
      Visible         =   0   'False
      Width           =   6270
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridProductPrices 
         Height          =   885
         Left            =   15
         TabIndex        =   222
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
         stylesets(0).Picture=   "FrmSaleInvoiceDist.frx":00D0
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
         stylesets(1).Picture=   "FrmSaleInvoiceDist.frx":00EC
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
         stylesets(2).Picture=   "FrmSaleInvoiceDist.frx":0108
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
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmSaleInvoiceDist.frx":0124
      Left            =   9705
      List            =   "FrmSaleInvoiceDist.frx":0126
      Style           =   2  'Dropdown List
      TabIndex        =   216
      Tag             =   "1"
      Top             =   9870
      Width           =   3276
   End
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   10860
      TabIndex        =   214
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   9510
      Width           =   2115
   End
   Begin VB.CheckBox ChkDiscB4TradeOffer 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4 "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13815
      TabIndex        =   189
      Top             =   4620
      Width           =   1290
   End
   Begin VB.CheckBox ChkDiscB4ExtraScheme 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13860
      TabIndex        =   188
      Top             =   5700
      Width           =   1290
   End
   Begin VB.CheckBox ChkDiscB4SaleTax 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13815
      TabIndex        =   187
      Top             =   6870
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   1350
      TabIndex        =   168
      Top             =   5610
      Visible         =   0   'False
      Width           =   2295
      Begin SITextBox.Txt TxtSerial 
         Height          =   315
         Left            =   120
         TabIndex        =   169
         Top             =   240
         Width           =   2040
         _ExtentX        =   3598
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridSerial 
         Height          =   1500
         Left            =   120
         TabIndex        =   170
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
         stylesets(0).Picture=   "FrmSaleInvoiceDist.frx":0128
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
   Begin VB.Frame FrmExpiry 
      Height          =   1590
      Left            =   4455
      TabIndex        =   157
      Top             =   5415
      Visible         =   0   'False
      Width           =   4095
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridExpiry 
         Height          =   1395
         Left            =   120
         TabIndex        =   158
         Top             =   -330
         Width           =   3885
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   2
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
         stylesets(0).Picture=   "FrmSaleInvoiceDist.frx":0144
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
         stylesets(1).Picture=   "FrmSaleInvoiceDist.frx":0160
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
         stylesets(2).Picture=   "FrmSaleInvoiceDist.frx":017C
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
         stylesets(3).Picture=   "FrmSaleInvoiceDist.frx":0198
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
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "BatchNo"
         Columns(0).Name =   "BatchNo"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "ExpiryDate"
         Columns(1).Name =   "ExpiryDate"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   6853
         _ExtentY        =   2461
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00EFC09E&
      Caption         =   "Freight"
      Height          =   720
      Left            =   9465
      TabIndex        =   141
      Top             =   8610
      Width           =   1245
      Begin VB.OptionButton OptCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Customer"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   143
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
         TabIndex        =   142
         Top             =   450
         Value           =   -1  'True
         Width           =   1530
      End
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   6150
      MaxLength       =   50
      TabIndex        =   104
      Top             =   9870
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CheckBox ChkIsProduct 
      Caption         =   "Is Product"
      Height          =   255
      Left            =   6330
      TabIndex        =   98
      Top             =   570
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   4995
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   4305
      Width           =   1425
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
      Left            =   15000
      TabIndex        =   73
      Top             =   270
      Visible         =   0   'False
      Width           =   3660
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
         Left            =   540
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   74
         Tag             =   "NC"
         Text            =   "FrmSaleInvoiceDist.frx":01B4
         Top             =   315
         Width           =   4095
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
         TabIndex        =   75
         Top             =   90
         Width           =   135
      End
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   705
      TabIndex        =   0
      Top             =   1230
      Width           =   570
      _ExtentX        =   1005
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
      Left            =   7650
      TabIndex        =   54
      Top             =   9375
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
      MICON           =   "FrmSaleInvoiceDist.frx":02CB
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6345
      TabIndex        =   50
      Top             =   9375
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
      MICON           =   "FrmSaleInvoiceDist.frx":02E7
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3735
      TabIndex        =   52
      Top             =   9375
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
      MICON           =   "FrmSaleInvoiceDist.frx":0303
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8955
      TabIndex        =   55
      Top             =   9375
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
      MICON           =   "FrmSaleInvoiceDist.frx":031F
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5040
      TabIndex        =   51
      Top             =   9375
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
      MICON           =   "FrmSaleInvoiceDist.frx":033B
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   2475
      TabIndex        =   56
      Top             =   8265
      Width           =   1185
      _ExtentX        =   2090
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
      Left            =   5100
      TabIndex        =   47
      Top             =   8265
      Width           =   975
      _ExtentX        =   1720
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
   Begin SITextBox.Txt TxtNetAmount 
      Height          =   315
      Left            =   8235
      TabIndex        =   58
      Top             =   8265
      Width           =   1395
      _ExtentX        =   2461
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
   Begin JeweledBut.JeweledButton BtnProduct 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2070
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4305
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
      MICON           =   "FrmSaleInvoiceDist.frx":0357
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpBillDate 
      Height          =   315
      Left            =   1245
      TabIndex        =   1
      Top             =   1230
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
      Left            =   7890
      TabIndex        =   6
      Tag             =   "NC"
      Top             =   1230
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
      Left            =   8925
      TabIndex        =   7
      Tag             =   "NC"
      Top             =   1230
      Width           =   1380
      _ExtentX        =   2434
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
      Left            =   8565
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   1230
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
      MICON           =   "FrmSaleInvoiceDist.frx":0373
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2400
      TabIndex        =   53
      Top             =   9375
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
      MICON           =   "FrmSaleInvoiceDist.frx":038F
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillDisc 
      Height          =   315
      Left            =   6075
      TabIndex        =   48
      Top             =   8265
      Width           =   975
      _ExtentX        =   1720
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   8340
      TabIndex        =   67
      Top             =   285
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
   Begin SITextBox.Txt TxtTotalQtys 
      Height          =   315
      Left            =   1590
      TabIndex        =   70
      Top             =   8265
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
      Left            =   7050
      TabIndex        =   49
      Top             =   8265
      Width           =   1185
      _ExtentX        =   2090
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
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2430
      TabIndex        =   26
      Top             =   4305
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
   Begin SITextBox.Txt TxtMultiplier 
      Height          =   315
      Left            =   6420
      TabIndex        =   28
      Top             =   4305
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
      Left            =   7440
      TabIndex        =   31
      Top             =   4305
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
      Left            =   6930
      TabIndex        =   30
      Top             =   4305
      Width           =   510
      _ExtentX        =   900
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
   Begin SITextBox.Txt TxtDiscVal 
      Height          =   315
      Left            =   11325
      TabIndex        =   42
      Top             =   4305
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
      DecimalPoint    =   3
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtPrice 
      Height          =   315
      Left            =   9000
      TabIndex        =   34
      Top             =   4305
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
      Left            =   10830
      TabIndex        =   38
      Top             =   4305
      Width           =   495
      _ExtentX        =   873
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
      Left            =   9645
      TabIndex        =   36
      Top             =   4305
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
      Masked          =   2
      DecimalPoint    =   3
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtBonus 
      Height          =   315
      Left            =   7980
      TabIndex        =   32
      Top             =   4305
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
      Left            =   8520
      TabIndex        =   33
      Top             =   4305
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
      Left            =   10320
      TabIndex        =   37
      Top             =   4305
      Width           =   510
      _ExtentX        =   900
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
      DecimalPoint    =   3
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   12690
      TabIndex        =   44
      Top             =   4305
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
   Begin SITextBox.Txt TxtMemberID 
      Height          =   315
      Left            =   12285
      TabIndex        =   16
      Top             =   2520
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
      Left            =   13320
      TabIndex        =   94
      Top             =   2520
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
      Left            =   12960
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   2520
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
      MICON           =   "FrmSaleInvoiceDist.frx":03AB
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCost 
      Height          =   315
      Left            =   10185
      TabIndex        =   99
      Top             =   285
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
   Begin SITextBox.Txt TxtCommission 
      Height          =   315
      Left            =   9300
      TabIndex        =   101
      Tag             =   "NC"
      Top             =   285
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
   Begin SITextBox.Txt TxtRemarks 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7815
      TabIndex        =   21
      Top             =   3180
      Width           =   2130
      _ExtentX        =   3757
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
      IntegralPoint   =   4
   End
   Begin SITextBox.Txt TxtManualBillNo 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6495
      TabIndex        =   106
      Top             =   300
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
   Begin SITextBox.Txt TxtBillNo 
      Height          =   315
      Left            =   675
      TabIndex        =   17
      Top             =   3180
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
      Left            =   1395
      TabIndex        =   18
      Top             =   3180
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
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   315
      Left            =   675
      TabIndex        =   14
      Top             =   2520
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   15
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
   Begin SITextBox.Txt TxtAddress 
      Height          =   315
      Left            =   5265
      TabIndex        =   108
      Top             =   2520
      Width           =   3405
      _ExtentX        =   6006
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
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   5040
      TabIndex        =   20
      Top             =   3180
      Width           =   2775
      _ExtentX        =   4895
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
   Begin SITextBox.Txt TxtRetailPrice 
      Height          =   315
      Left            =   11730
      TabIndex        =   119
      Top             =   285
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
   Begin SITextBox.Txt TxtCity 
      Height          =   315
      Left            =   8670
      TabIndex        =   121
      Top             =   2520
      Width           =   1470
      _ExtentX        =   2593
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
   Begin SITextBox.Txt TxtTokenVal 
      Height          =   315
      Left            =   12450
      TabIndex        =   122
      Top             =   285
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
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   1965
      TabIndex        =   15
      Top             =   2520
      Width           =   2940
      _ExtentX        =   5186
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
   Begin JeweledBut.JeweledButton BtnCustomer 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   1605
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   2520
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
      MICON           =   "FrmSaleInvoiceDist.frx":03C7
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1215
      TabIndex        =   23
      Top             =   4305
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
      Masked          =   1
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3240
      Left            =   675
      TabIndex        =   129
      Top             =   4620
      Visible         =   0   'False
      Width           =   13155
      ScrollBars      =   3
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   47
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
      stylesets(0).Picture=   "FrmSaleInvoiceDist.frx":03E3
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
      stylesets(1).Picture=   "FrmSaleInvoiceDist.frx":03FF
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
      stylesets(2).Picture=   "FrmSaleInvoiceDist.frx":041B
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
      stylesets(3).Picture=   "FrmSaleInvoiceDist.frx":0437
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
      Columns.Count   =   47
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ProductID"
      Columns(0).Name =   "ProductID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   979
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
      Columns(4).Width=   2514
      Columns(4).Caption=   "Pack Name"
      Columns(4).Name =   "PackName"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   900
      Columns(5).Caption=   "Pack"
      Columns(5).Name =   "Pack"
      Columns(5).Alignment=   1
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   900
      Columns(6).Caption=   "Q(P)"
      Columns(6).Name =   "QtyPack"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   4
      Columns(6).FieldLen=   256
      Columns(7).Width=   953
      Columns(7).Caption=   "Q(L)"
      Columns(7).Name =   "QtyLoose"
      Columns(7).Alignment=   1
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   953
      Columns(8).Caption=   "Bns"
      Columns(8).Name =   "Bonus"
      Columns(8).Alignment=   1
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   4
      Columns(8).FieldLen=   256
      Columns(9).Width=   847
      Columns(9).Caption=   "Offer"
      Columns(9).Name =   "Offer"
      Columns(9).Alignment=   1
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1138
      Columns(10).Caption=   "Price"
      Columns(10).Name=   "Price"
      Columns(10).Alignment=   1
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   4
      Columns(10).FieldLen=   256
      Columns(11).Width=   1191
      Columns(11).Caption=   "DiscPC"
      Columns(11).Name=   "DiscPC"
      Columns(11).Alignment=   1
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   926
      Columns(12).Caption=   "Tax%"
      Columns(12).Name=   "SaleTaxPer"
      Columns(12).Alignment=   1
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   847
      Columns(13).Caption=   "Dis%"
      Columns(13).Name=   "DiscPer"
      Columns(13).Alignment=   1
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   1191
      Columns(14).Caption=   "Dis.Val"
      Columns(14).Name=   "DiscVal"
      Columns(14).Alignment=   1
      Columns(14).CaptionAlignment=   2
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   4
      Columns(14).FieldLen=   256
      Columns(15).Width=   1217
      Columns(15).Caption=   "SC"
      Columns(15).Name=   "SC"
      Columns(15).Alignment=   1
      Columns(15).CaptionAlignment=   2
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   1508
      Columns(16).Caption=   "Amount"
      Columns(16).Name=   "Amount"
      Columns(16).Alignment=   1
      Columns(16).CaptionAlignment=   2
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   5
      Columns(16).FieldLen=   256
      Columns(17).Width=   3200
      Columns(17).Visible=   0   'False
      Columns(17).Caption=   "PackingID"
      Columns(17).Name=   "PackingID"
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      Columns(18).Width=   3200
      Columns(18).Visible=   0   'False
      Columns(18).Caption=   "SaleTaxVal"
      Columns(18).Name=   "SaleTaxVal"
      Columns(18).Alignment=   1
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(19).Width=   3200
      Columns(19).Visible=   0   'False
      Columns(19).Caption=   "IsProduct"
      Columns(19).Name=   "IsProduct"
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   11
      Columns(19).FieldLen=   256
      Columns(20).Width=   3200
      Columns(20).Visible=   0   'False
      Columns(20).Caption=   "Cost"
      Columns(20).Name=   "Cost"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      Columns(21).Width=   1058
      Columns(21).Caption=   "Qty(G)"
      Columns(21).Name=   "GrossQty"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   8
      Columns(21).FieldLen=   256
      Columns(22).Width=   1058
      Columns(22).Caption=   "Qty(U)"
      Columns(22).Name=   "GrossUnit"
      Columns(22).DataField=   "Column 22"
      Columns(22).DataType=   8
      Columns(22).FieldLen=   256
      Columns(23).Width=   1958
      Columns(23).Caption=   "RetailPrice"
      Columns(23).Name=   "RetailPrice"
      Columns(23).Alignment=   1
      Columns(23).DataField=   "Column 23"
      Columns(23).DataType=   8
      Columns(23).FieldLen=   256
      Columns(24).Width=   2461
      Columns(24).Caption=   "IsWSDiscb4ST"
      Columns(24).Name=   "IsWSDiscb4ST"
      Columns(24).DataField=   "Column 24"
      Columns(24).DataType=   11
      Columns(24).FieldLen=   256
      Columns(25).Width=   2328
      Columns(25).Caption=   "IsWSSaleTax"
      Columns(25).Name=   "IsWSSaleTax"
      Columns(25).DataField=   "Column 25"
      Columns(25).DataType=   11
      Columns(25).FieldLen=   256
      Columns(26).Width=   2461
      Columns(26).Caption=   "IsRetailSaleTax"
      Columns(26).Name=   "IsRetailSaleTax"
      Columns(26).DataField=   "Column 26"
      Columns(26).DataType=   11
      Columns(26).FieldLen=   256
      Columns(27).Width=   2037
      Columns(27).Caption=   "TokenVal"
      Columns(27).Name=   "TokenVal"
      Columns(27).DataField=   "Column 27"
      Columns(27).DataType=   8
      Columns(27).FieldLen=   256
      Columns(28).Width=   1640
      Columns(28).Caption=   "BatchNo"
      Columns(28).Name=   "BatchNo"
      Columns(28).DataField=   "Column 28"
      Columns(28).DataType=   8
      Columns(28).FieldLen=   256
      Columns(29).Width=   3200
      Columns(29).Visible=   0   'False
      Columns(29).Caption=   "ExpiryTime"
      Columns(29).Name=   "ExpiryTime"
      Columns(29).DataField=   "Column 29"
      Columns(29).DataType=   8
      Columns(29).FieldLen=   256
      Columns(30).Width=   1852
      Columns(30).Caption=   "ExpiryDate"
      Columns(30).Name=   "ExpiryDate"
      Columns(30).DataField=   "Column 30"
      Columns(30).DataType=   8
      Columns(30).FieldLen=   256
      Columns(31).Width=   3200
      Columns(31).Caption=   "EmpID"
      Columns(31).Name=   "EmpID"
      Columns(31).DataField=   "Column 31"
      Columns(31).DataType=   8
      Columns(31).FieldLen=   256
      Columns(32).Width=   3200
      Columns(32).Caption=   "EmpName"
      Columns(32).Name=   "EmpName"
      Columns(32).DataField=   "Column 32"
      Columns(32).DataType=   8
      Columns(32).FieldLen=   256
      Columns(33).Width=   3200
      Columns(33).Caption=   "StoreID"
      Columns(33).Name=   "StoreID"
      Columns(33).DataField=   "Column 33"
      Columns(33).DataType=   8
      Columns(33).FieldLen=   256
      Columns(34).Width=   3200
      Columns(34).Caption=   "StoreName"
      Columns(34).Name=   "StoreName"
      Columns(34).DataField=   "Column 34"
      Columns(34).DataType=   8
      Columns(34).FieldLen=   256
      Columns(35).Width=   3200
      Columns(35).Caption=   "isDiscB4TradeOffer"
      Columns(35).Name=   "isDiscB4TradeOffer"
      Columns(35).DataField=   "Column 35"
      Columns(35).DataType=   11
      Columns(35).FieldLen=   256
      Columns(36).Width=   3200
      Columns(36).Caption=   "isDiscB4ExtraScheme"
      Columns(36).Name=   "isDiscB4ExtraScheme"
      Columns(36).DataField=   "Column 36"
      Columns(36).DataType=   11
      Columns(36).FieldLen=   256
      Columns(37).Width=   3200
      Columns(37).Caption=   "isDiscB4SaleTax"
      Columns(37).Name=   "isDiscB4SaleTax"
      Columns(37).DataField=   "Column 37"
      Columns(37).DataType=   11
      Columns(37).FieldLen=   256
      Columns(38).Width=   3200
      Columns(38).Caption=   "TradeOffer1"
      Columns(38).Name=   "TradeOffer1"
      Columns(38).DataField=   "Column 38"
      Columns(38).DataType=   8
      Columns(38).FieldLen=   256
      Columns(39).Width=   3200
      Columns(39).Caption=   "TradeOffer2"
      Columns(39).Name=   "TradeOffer2"
      Columns(39).DataField=   "Column 39"
      Columns(39).DataType=   8
      Columns(39).FieldLen=   256
      Columns(40).Width=   3200
      Columns(40).Caption=   "ExtraSchemePer"
      Columns(40).Name=   "ExtraSchemePer"
      Columns(40).DataField=   "Column 40"
      Columns(40).DataType=   8
      Columns(40).FieldLen=   256
      Columns(41).Width=   3200
      Columns(41).Caption=   "TradeValue"
      Columns(41).Name=   "TradeValue"
      Columns(41).DataField=   "Column 41"
      Columns(41).DataType=   8
      Columns(41).FieldLen=   256
      Columns(42).Width=   3200
      Columns(42).Caption=   "ExtraSchemeValue"
      Columns(42).Name=   "ExtraSchemeValue"
      Columns(42).DataField=   "Column 42"
      Columns(42).DataType=   8
      Columns(42).FieldLen=   256
      Columns(43).Width=   3200
      Columns(43).Caption=   "BasicAmount"
      Columns(43).Name=   "DiscAmount"
      Columns(43).DataField=   "Column 43"
      Columns(43).DataType=   8
      Columns(43).FieldLen=   256
      Columns(44).Width=   3200
      Columns(44).Visible=   0   'False
      Columns(44).Caption=   "isLastPrice"
      Columns(44).Name=   "isLastPrice"
      Columns(44).DataField=   "Column 44"
      Columns(44).DataType=   11
      Columns(44).FieldLen=   256
      Columns(45).Width=   1773
      Columns(45).Caption=   "ReSPrice"
      Columns(45).Name=   "ReSPrice"
      Columns(45).Alignment=   1
      Columns(45).DataField=   "Column 45"
      Columns(45).DataType=   8
      Columns(45).FieldLen=   256
      Columns(46).Width=   2143
      Columns(46).Caption=   "ReSAmount"
      Columns(46).Name=   "ReSAmount"
      Columns(46).Alignment=   1
      Columns(46).DataField=   "Column 46"
      Columns(46).DataType=   8
      Columns(46).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   23204
      _ExtentY        =   5715
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
   Begin SITextBox.Txt TxtOrderID 
      Height          =   315
      Left            =   705
      TabIndex        =   9
      Top             =   1905
      Width           =   735
      _ExtentX        =   1296
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
   Begin JeweledBut.JeweledButton BtnSaleOrder 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2745
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   1905
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
      MICON           =   "FrmSaleInvoiceDist.frx":0453
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   11415
      TabIndex        =   22
      Tag             =   "NC"
      Top             =   1905
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
      Left            =   12525
      TabIndex        =   133
      Tag             =   "NC"
      Top             =   1905
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
      Left            =   12165
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   1905
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
      MICON           =   "FrmSaleInvoiceDist.frx":046F
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpOrderDate 
      Height          =   315
      Left            =   1440
      TabIndex        =   10
      Top             =   1905
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
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   11475
      TabIndex        =   137
      Tag             =   "NC"
      Top             =   1230
      Width           =   1665
      _ExtentX        =   2937
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
      Left            =   11115
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   1230
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
      MICON           =   "FrmSaleInvoiceDist.frx":048B
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtVehicleNo 
      Height          =   315
      Left            =   2175
      TabIndex        =   19
      Top             =   3180
      Width           =   2865
      _ExtentX        =   5054
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
   Begin SITextBox.Txt TxtBatchNo 
      Height          =   315
      Left            =   2205
      TabIndex        =   24
      Top             =   3990
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   15
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
   Begin SITextBox.Txt TxtExpiryDate 
      Height          =   315
      Left            =   3060
      TabIndex        =   144
      Top             =   3990
      Width           =   855
      _ExtentX        =   1508
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
   End
   Begin SITextBox.Txt TxtReceivedAmount 
      Height          =   315
      Left            =   7140
      TabIndex        =   145
      Top             =   8985
      Width           =   1395
      _ExtentX        =   2461
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
   Begin SITextBox.Txt TxtPreviousReceivable 
      Height          =   315
      Left            =   4170
      TabIndex        =   146
      Top             =   8985
      Width           =   1575
      _ExtentX        =   2778
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
   Begin SITextBox.Txt TxtTotalReceivable 
      Height          =   315
      Left            =   5745
      TabIndex        =   147
      Top             =   8985
      Width           =   1395
      _ExtentX        =   2461
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
   Begin SITextBox.Txt TxtFreight 
      Height          =   315
      Left            =   8535
      TabIndex        =   148
      Top             =   8985
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Alignment       =   1
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
   Begin JeweledBut.JeweledButton BtnProductRange 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   1845
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   3975
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
      MICON           =   "FrmSaleInvoiceDist.frx":04A7
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrintWarranty 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   960
      TabIndex        =   151
      Top             =   9375
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      TX              =   "Print Warranty"
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
      MICON           =   "FrmSaleInvoiceDist.frx":04C3
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPurchase 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5220
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   1905
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "Pur"
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
      MICON           =   "FrmSaleInvoiceDist.frx":04DF
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtPurID 
      Height          =   315
      Left            =   3180
      TabIndex        =   11
      Top             =   1905
      Width           =   735
      _ExtentX        =   1296
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpPurchaseDate 
      Height          =   315
      Left            =   3930
      TabIndex        =   12
      Top             =   1905
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
   Begin JeweledBut.JeweledButton BtnSaveAS 
      Height          =   420
      Left            =   2715
      TabIndex        =   153
      Top             =   8835
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save As"
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
      MICON           =   "FrmSaleInvoiceDist.frx":04FB
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnBatchPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4680
      TabIndex        =   154
      Top             =   9720
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Batch Print"
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
      MICON           =   "FrmSaleInvoiceDist.frx":0517
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpPromiseDate 
      Height          =   315
      Left            =   6345
      TabIndex        =   5
      Top             =   1230
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
   Begin SITextBox.Txt TxtSyllabusID 
      Height          =   315
      Left            =   5625
      TabIndex        =   13
      Top             =   1920
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
   Begin SITextBox.Txt TxtSyllabusName 
      Height          =   315
      Left            =   6690
      TabIndex        =   159
      Top             =   1920
      Width           =   2430
      _ExtentX        =   4286
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
   Begin JeweledBut.JeweledButton BtnSyllabus 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6330
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   1920
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
      MICON           =   "FrmSaleInvoiceDist.frx":0533
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   10410
      TabIndex        =   8
      Tag             =   "NC"
      Top             =   1230
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
   Begin JeweledBut.JeweledButton BtnAddCustomer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4905
      TabIndex        =   165
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   2520
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "+"
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
      MICON           =   "FrmSaleInvoiceDist.frx":054F
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtContactNo 
      Height          =   315
      Left            =   10140
      TabIndex        =   166
      Top             =   2520
      Width           =   1530
      _ExtentX        =   2699
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
   Begin SITextBox.Txt TxtServiceCharges 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3660
      TabIndex        =   45
      Top             =   8265
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
   Begin SITextBox.Txt TxtServiceChargesPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4500
      TabIndex        =   46
      Top             =   8265
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
   Begin SITextBox.Txt TxtSC 
      Height          =   315
      Left            =   12000
      TabIndex        =   43
      Top             =   4305
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
      Left            =   5835
      TabIndex        =   29
      Top             =   3765
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
      Left            =   7485
      TabIndex        =   40
      Top             =   3765
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
      Left            =   6720
      TabIndex        =   39
      Top             =   3765
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
   Begin SITextBox.Txt TxtTradeOfferValue 
      Height          =   315
      Left            =   13860
      TabIndex        =   181
      Top             =   5040
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
      Left            =   13860
      TabIndex        =   183
      Top             =   6120
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
   Begin SITextBox.Txt TxtExtraSchemePer 
      Height          =   315
      Left            =   11310
      TabIndex        =   41
      Top             =   3765
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
   Begin SITextBox.Txt TxtSaleTaxVal 
      Height          =   315
      Left            =   13860
      TabIndex        =   190
      Top             =   7320
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
   Begin SITextBox.Txt TxtLicenceNO 
      Height          =   315
      Left            =   9135
      TabIndex        =   191
      Top             =   1920
      Width           =   1680
      _ExtentX        =   2963
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpExpiryInvoice 
      Height          =   315
      Left            =   2550
      TabIndex        =   2
      Top             =   1230
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
   Begin SITextBox.Txt TxtTotalItems 
      Height          =   315
      Left            =   705
      TabIndex        =   194
      Top             =   8265
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
   Begin SITextBox.Txt TxtExtraTaxVal 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   13230
      TabIndex        =   202
      Top             =   9750
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
      Left            =   14430
      TabIndex        =   204
      Top             =   9750
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
      Left            =   13230
      TabIndex        =   199
      Top             =   9195
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
      Left            =   14430
      TabIndex        =   200
      Top             =   9195
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
   Begin SITextBox.Txt TxtCashCustomer 
      Height          =   315
      Left            =   11220
      TabIndex        =   196
      Top             =   8250
      Width           =   4065
      _ExtentX        =   7170
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
      Masked          =   5
   End
   Begin SITextBox.Txt TxtCNIC 
      Height          =   315
      Left            =   12000
      TabIndex        =   197
      Top             =   8610
      Width           =   1425
      _ExtentX        =   2514
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
      Masked          =   5
   End
   Begin SITextBox.Txt TxtCellNo 
      Height          =   315
      Left            =   14040
      TabIndex        =   198
      Top             =   8610
      Width           =   1185
      _ExtentX        =   2090
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
      Masked          =   5
   End
   Begin SITextBox.Txt TxtSumDiscAmount 
      Height          =   315
      Left            =   120
      TabIndex        =   211
      Top             =   8970
      Width           =   1425
      _ExtentX        =   2514
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
   Begin SITextBox.Txt TxtDiscAmount 
      Height          =   315
      Left            =   30
      TabIndex        =   213
      Top             =   4290
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpDispatchDate 
      Height          =   315
      Left            =   3870
      TabIndex        =   3
      Top             =   1215
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
   Begin SITextBox.Txt TxtSID 
      Height          =   315
      Left            =   -30
      TabIndex        =   236
      Top             =   1215
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
   Begin SITextBox.Txt TxtRefID 
      Height          =   315
      Left            =   30
      TabIndex        =   244
      Top             =   615
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   1058
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
   Begin SITextBox.Txt TxtRefComm 
      Height          =   315
      Left            =   960
      TabIndex        =   246
      Top             =   630
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtGrossQty 
      Height          =   315
      Left            =   4995
      TabIndex        =   251
      Top             =   3765
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
   Begin SITextBox.Txt TxtReSPrice 
      Height          =   315
      Left            =   9000
      TabIndex        =   35
      Top             =   3735
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
   Begin SITextBox.Txt TxtReSAmount 
      Height          =   315
      Left            =   12690
      TabIndex        =   252
      Top             =   3765
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
   Begin VB.Frame FrmMultipleStore 
      Caption         =   "Multiple Store"
      Height          =   2955
      Left            =   5355
      TabIndex        =   223
      Top             =   1575
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ComboBox CmbMSPackName 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   228
         Top             =   615
         Width           =   1425
      End
      Begin VB.ComboBox CmbMSStore 
         Height          =   315
         Left            =   2745
         Style           =   2  'Dropdown List
         TabIndex        =   227
         Top             =   615
         Width           =   1455
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridMultipleStore 
         Height          =   1560
         Left            =   180
         TabIndex        =   224
         Top             =   930
         Width           =   7320
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   11
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
         stylesets(0).Picture=   "FrmSaleInvoiceDist.frx":056B
         AllowDelete     =   -1  'True
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
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
         ForeColorEven   =   0
         BackColorOdd    =   15724527
         RowHeight       =   423
         ExtraHeight     =   26
         ActiveRowStyleSet=   "SelectedRow"
         Columns.Count   =   11
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "SID"
         Columns(0).Name =   "SID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "ProductID"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   4498
         Columns(2).Caption=   "ProductName"
         Columns(2).Name =   "ProductName"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "StoreID"
         Columns(3).Name =   "StoreID"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   2593
         Columns(4).Caption=   "Store Name"
         Columns(4).Name =   "StoreName"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "PackingID"
         Columns(5).Name =   "PackingID"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   2514
         Columns(6).Caption=   "Pack Name"
         Columns(6).Name =   "PackName"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   900
         Columns(7).Caption=   "Pack"
         Columns(7).Name =   "Pack"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   900
         Columns(8).Caption=   "Q (P)"
         Columns(8).Name =   "QtyPack"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   953
         Columns(9).Caption=   "Q (L)"
         Columns(9).Name =   "QtyLoose"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   3200
         Columns(10).Visible=   0   'False
         Columns(10).Caption=   "Bonus"
         Columns(10).Name=   "Bonus"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   12912
         _ExtentY        =   2752
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
      Begin SITextBox.Txt TxtMSCode 
         Height          =   315
         Left            =   180
         TabIndex        =   225
         Top             =   270
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin SITextBox.Txt TxtMSQtyLoose 
         Height          =   315
         Left            =   6645
         TabIndex        =   233
         Top             =   615
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
      Begin SITextBox.Txt TxtMSQtyPack 
         Height          =   315
         Left            =   6135
         TabIndex        =   230
         Top             =   615
         Width           =   510
         _ExtentX        =   900
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
         Mandatory       =   1
      End
      Begin SITextBox.Txt TxtMSMultiplier 
         Height          =   315
         Left            =   5625
         TabIndex        =   229
         Top             =   615
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   556
         Alignment       =   1
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin SITextBox.Txt TxtMSTotalItems 
         Height          =   315
         Left            =   6330
         TabIndex        =   239
         Top             =   2535
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
      Begin SITextBox.Txt TxtMSProductName 
         Height          =   315
         Left            =   180
         TabIndex        =   241
         Top             =   600
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
      Begin SITextBox.Txt TxtMSBonus 
         Height          =   315
         Left            =   1530
         TabIndex        =   242
         Top             =   2505
         Visible         =   0   'False
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
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Bns(L)"
         Height          =   195
         Left            =   930
         TabIndex        =   243
         Top             =   2550
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Store Items"
         Height          =   195
         Left            =   5040
         TabIndex        =   240
         Top             =   2610
         Width           =   1200
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Pack"
         Height          =   195
         Left            =   5625
         TabIndex        =   238
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty (P)"
         Height          =   195
         Left            =   6135
         TabIndex        =   235
         Top             =   390
         Width           =   480
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty (L)"
         Height          =   195
         Left            =   6645
         TabIndex        =   234
         Top             =   390
         Width           =   465
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Pack Name"
         Height          =   195
         Left            =   4200
         TabIndex        =   232
         Top             =   390
         Width           =   840
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Store Name"
         Height          =   195
         Left            =   2745
         TabIndex        =   231
         Top             =   390
         Width           =   840
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         Height          =   195
         Left            =   1080
         TabIndex        =   226
         Top             =   390
         Width           =   555
      End
   End
   Begin SITextBox.Txt TxtTerms 
      Height          =   315
      Left            =   5850
      TabIndex        =   4
      Tag             =   "NC"
      Top             =   1215
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
   Begin VB.Label LblCustomerDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   660
      TabIndex        =   256
      Top             =   3510
      Width           =   1200
   End
   Begin VB.Label LblPurPrice 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Price"
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
      Left            =   12780
      TabIndex        =   254
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label LblReSAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Re Sale Amount"
      Height          =   195
      Left            =   12510
      TabIndex        =   253
      Top             =   3555
      Width           =   1155
   End
   Begin VB.Label LblReSPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Re Sale Price"
      Height          =   195
      Left            =   8820
      TabIndex        =   250
      Top             =   3555
      Width           =   975
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Index           =   3
      Left            =   1215
      TabIndex        =   249
      Top             =   4110
      Width           =   375
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Index           =   2
      Left            =   1245
      TabIndex        =   248
      Top             =   1035
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "SID"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   247
      Top             =   1035
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label55 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Reference ID Comm %"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   30
      TabIndex        =   245
      Top             =   390
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill ID"
      Height          =   195
      Index           =   1
      Left            =   705
      TabIndex        =   237
      Top             =   1035
      Width           =   405
   End
   Begin VB.Label LblTerms 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Terms"
      Height          =   195
      Left            =   5850
      TabIndex        =   220
      Top             =   1020
      Width           =   435
   End
   Begin VB.Label LblDispatchDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dispatch Date"
      Height          =   195
      Left            =   3870
      TabIndex        =   219
      Top             =   1020
      Width           =   1020
   End
   Begin VB.Label LblLastPrice 
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
      Left            =   1440
      TabIndex        =   218
      Top             =   30
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   9030
      TabIndex        =   217
      Top             =   9915
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
      Left            =   10860
      TabIndex        =   215
      Top             =   9210
      Width           =   840
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sum of Disc Amount"
      Height          =   195
      Left            =   120
      TabIndex        =   212
      Top             =   8730
      Width           =   1440
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Amount"
      Height          =   195
      Index           =   7
      Left            =   30
      TabIndex        =   210
      Top             =   4050
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11220
      TabIndex        =   209
      Top             =   8010
      Width           =   1665
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "CNIC No"
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
      Left            =   11280
      TabIndex        =   208
      Top             =   8610
      Width           =   645
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cell:"
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
      Left            =   13560
      TabIndex        =   207
      Top             =   8610
      Width           =   360
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   14430
      TabIndex        =   206
      Top             =   9525
      Width           =   210
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Tax"
      Height          =   195
      Left            =   13230
      TabIndex        =   205
      Top             =   9525
      Width           =   675
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Tax"
      Height          =   195
      Left            =   13230
      TabIndex        =   203
      Top             =   8970
      Width           =   960
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Index           =   0
      Left            =   14430
      TabIndex        =   201
      Top             =   8970
      Width           =   210
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items"
      Height          =   195
      Index           =   11
      Left            =   705
      TabIndex        =   195
      Top             =   8040
      Width           =   780
   End
   Begin VB.Label LblExpiryInvoice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expirty Invoice"
      Height          =   195
      Left            =   2550
      TabIndex        =   193
      Top             =   1035
      Width           =   1035
   End
   Begin VB.Label LblLicenceNO 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Licence #"
      Height          =   195
      Left            =   9135
      TabIndex        =   192
      Top             =   1740
      Width           =   720
   End
   Begin VB.Label LblExtraSchemePer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ex. Scheme %"
      Height          =   195
      Left            =   10215
      TabIndex        =   186
      Top             =   3810
      Width           =   1020
   End
   Begin VB.Label LblGSTValue 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "GST Value"
      Height          =   195
      Left            =   13860
      TabIndex        =   185
      Top             =   7095
      Width           =   780
   End
   Begin VB.Label LblExtraSchemeValue 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ex Scheme Value"
      Height          =   195
      Left            =   13860
      TabIndex        =   184
      Top             =   5925
      Width           =   1260
   End
   Begin VB.Label LblTradeValue 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Value"
      Height          =   195
      Left            =   13860
      TabIndex        =   182
      Top             =   4845
      Width           =   870
   End
   Begin VB.Label LblTradeOffer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Offer"
      Height          =   195
      Left            =   6885
      TabIndex        =   180
      Top             =   3555
      Width           =   810
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
      Left            =   7350
      TabIndex        =   179
      Top             =   3810
      Width           =   120
   End
   Begin VB.Label LblGrossUnit 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Unit"
      Height          =   195
      Left            =   5805
      TabIndex        =   178
      Top             =   3555
      Width           =   735
   End
   Begin VB.Label LblGrossQty 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Qty"
      Height          =   195
      Left            =   4995
      TabIndex        =   177
      Top             =   3555
      Width           =   690
   End
   Begin VB.Label LblSC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "S.C."
      Height          =   195
      Left            =   12015
      TabIndex        =   176
      Top             =   4110
      Width           =   300
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Ch."
      Height          =   195
      Left            =   3660
      TabIndex        =   175
      Top             =   8040
      Width           =   825
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   4500
      TabIndex        =   174
      Top             =   8040
      Width           =   210
   End
   Begin VB.Label LblTotalAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amount"
      Height          =   195
      Left            =   2505
      TabIndex        =   173
      Top             =   8040
      Width           =   990
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      Height          =   195
      Left            =   10140
      TabIndex        =   167
      Top             =   2295
      Width           =   765
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Date"
      Height          =   195
      Left            =   3930
      TabIndex        =   164
      Top             =   1710
      Width           =   630
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur ID"
      Height          =   195
      Left            =   3180
      TabIndex        =   163
      Top             =   1710
      Width           =   450
   End
   Begin VB.Label LblSyllabusID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Syllabus ID"
      Height          =   195
      Left            =   5625
      TabIndex        =   162
      Top             =   1725
      Width           =   795
   End
   Begin VB.Label LblSyllabusName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Syllabus Name"
      Height          =   195
      Left            =   6690
      TabIndex        =   161
      Top             =   1725
      Width           =   1050
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
      Left            =   4080
      TabIndex        =   156
      Top             =   3885
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblPromiseDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Promise Date"
      Height          =   195
      Left            =   6345
      TabIndex        =   155
      Top             =   1035
      Width           =   945
   End
   Begin VB.Label LblLastPurPrice 
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
      Left            =   1260
      TabIndex        =   149
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblFreight 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      Height          =   195
      Left            =   8550
      TabIndex        =   140
      Top             =   8760
      Width           =   480
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Transport Name"
      Height          =   195
      Index           =   6
      Left            =   2160
      TabIndex        =   139
      Top             =   2970
      Width           =   1140
   End
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   12510
      TabIndex        =   136
      Top             =   1695
      Width           =   780
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   11400
      TabIndex        =   135
      Top             =   1695
      Width           =   525
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID"
      Height          =   195
      Index           =   5
      Left            =   720
      TabIndex        =   131
      Top             =   1710
      Width           =   600
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
      Height          =   195
      Index           =   6
      Left            =   1410
      TabIndex        =   130
      Top             =   1710
      Width           =   780
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Token Val"
      Height          =   195
      Left            =   12450
      TabIndex        =   123
      Top             =   105
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ret Price"
      Height          =   195
      Left            =   11730
      TabIndex        =   120
      Top             =   105
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Received Amount"
      Height          =   195
      Left            =   7140
      TabIndex        =   118
      Top             =   8760
      Width           =   1275
   End
   Begin VB.Label lblPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Receivable"
      Height          =   195
      Left            =   4185
      TabIndex        =   117
      Top             =   8760
      Width           =   1470
   End
   Begin VB.Label LblTtlPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Receivable"
      Height          =   195
      Left            =   5745
      TabIndex        =   116
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      Height          =   195
      Index           =   6
      Left            =   675
      TabIndex        =   115
      Top             =   2970
      Width           =   495
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bilty No."
      Height          =   195
      Index           =   6
      Left            =   1410
      TabIndex        =   114
      Top             =   2970
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      Height          =   195
      Index           =   4
      Left            =   690
      TabIndex        =   113
      Top             =   2310
      Width           =   870
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   195
      Index           =   4
      Left            =   1995
      TabIndex        =   112
      Top             =   2310
      Width           =   1125
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   5265
      TabIndex        =   111
      Top             =   2310
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   8670
      TabIndex        =   110
      Top             =   2310
      Width           =   255
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   6
      Left            =   5040
      TabIndex        =   109
      Top             =   2970
      Width           =   795
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill No"
      Height          =   195
      Left            =   6495
      TabIndex        =   107
      Top             =   75
      Width           =   1020
   End
   Begin VB.Label LblRemarks 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   7815
      TabIndex        =   105
      Top             =   2955
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   10230
      TabIndex        =   103
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      Height          =   195
      Left            =   9225
      TabIndex        =   102
      Top             =   60
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblMemberID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      Height          =   195
      Left            =   12270
      TabIndex        =   97
      Top             =   2295
      Width           =   780
   End
   Begin VB.Label LblMemberName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      Height          =   195
      Left            =   13305
      TabIndex        =   96
      Top             =   2295
      Width           =   1035
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   11580
      TabIndex        =   93
      Top             =   1020
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   10410
      TabIndex        =   92
      Top             =   1035
      Width           =   1095
   End
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   12660
      TabIndex        =   91
      Top             =   4110
      Width           =   540
   End
   Begin VB.Label LblSaleTaxPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax%"
      Height          =   195
      Left            =   10320
      TabIndex        =   90
      Top             =   4110
      Width           =   390
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Val"
      Height          =   195
      Left            =   11100
      TabIndex        =   89
      Top             =   105
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label LblOffer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Offer"
      Height          =   195
      Left            =   8490
      TabIndex        =   88
      Top             =   4110
      Width           =   345
   End
   Begin VB.Label LblPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   9000
      TabIndex        =   87
      Top             =   4110
      Width           =   405
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   6450
      TabIndex        =   86
      Top             =   4110
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   5025
      TabIndex        =   85
      Top             =   4110
      Width           =   840
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   7440
      TabIndex        =   84
      Top             =   4110
      Width           =   465
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   6930
      TabIndex        =   83
      Top             =   4110
      Width           =   480
   End
   Begin VB.Label LblDiscval 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.Val"
      Height          =   195
      Left            =   11310
      TabIndex        =   82
      Top             =   4110
      Width           =   585
   End
   Begin VB.Label LblDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Dis%"
      Height          =   195
      Left            =   10830
      TabIndex        =   81
      Top             =   4110
      Width           =   345
   End
   Begin VB.Label LblDiscPC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc/PC"
      Height          =   195
      Left            =   9645
      TabIndex        =   80
      Top             =   4110
      Width           =   600
   End
   Begin VB.Label LblBonus 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bns(L)"
      Height          =   195
      Left            =   7980
      TabIndex        =   79
      Top             =   4110
      Width           =   450
   End
   Begin VB.Label LblRetailPrice 
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
      Left            =   13680
      TabIndex        =   78
      Top             =   1095
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   13575
      TabIndex        =   77
      Top             =   810
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
      Left            =   14055
      TabIndex        =   76
      Top             =   1890
      Width           =   435
   End
   Begin VB.Label LblOtherChargesCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      Height          =   195
      Left            =   7050
      TabIndex        =   72
      Top             =   8040
      Width           =   1020
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty(s)"
      Height          =   195
      Left            =   1590
      TabIndex        =   71
      Top             =   8040
      Width           =   810
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Invoice Distribution"
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
      TabIndex        =   69
      Top             =   270
      Width           =   4260
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   8340
      TabIndex        =   68
      Top             =   60
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   6075
      TabIndex        =   66
      Top             =   8040
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
      Left            =   11700
      TabIndex        =   65
      Top             =   2940
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
      Left            =   11610
      TabIndex        =   64
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   8925
      TabIndex        =   63
      Top             =   1035
      Width           =   840
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   7890
      TabIndex        =   62
      Top             =   1050
      Width           =   585
   End
   Begin VB.Label LblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3915
      TabIndex        =   60
      Top             =   4110
      Width           =   1020
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   14385
      Top             =   1785
      Width           =   330
   End
   Begin VB.Label LblNetAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      Height          =   195
      Left            =   8250
      TabIndex        =   59
      Top             =   8040
      Width           =   840
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount (%)"
      Height          =   195
      Left            =   5085
      TabIndex        =   57
      Top             =   8040
      Width           =   885
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
Attribute VB_Name = "FrmSaleInvoiceDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public cnSale As New ADODB.Connection
Dim Application1 As New CRAXDRT.Application
Dim vMode As FormMode, vBillTime As Date
Dim vUnitPrice As Double, vQty As Integer
Dim vDate, vServerDate As Date, vHDiff As Integer, vSystemDate As Boolean
Dim vUnitRetailPrice As Double
Dim vIsWSDiscb4ST As Boolean
Dim vIsRetailSaleTax As Boolean
Dim vIsWSSaleTax As Boolean
Dim vIsNewRecord As Boolean
Dim vCounter As Integer, isUrdu As Boolean
Dim vBm As Variant, vExpiryTime As String
Dim vMaxBinID, vGridRows As Integer
Dim RsBody As New ADODB.Recordset
Dim RsBodyStore As New ADODB.Recordset
Dim RsExpense As New ADODB.Recordset
Dim RsBodySerial As New ADODB.Recordset
Dim RsProductOffer As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim QtyOffer As Integer
Dim Rebate As Integer
Dim DateFlag As Boolean
Dim Flag As Boolean, vAlreadySerial As Boolean
Dim ssql As String
Dim vStrSQL, vSamePid, vRandomID  As String
Dim vBillID  As Double, vZoneID As Byte
Dim vBillDate  As Date
Dim ExpenseFlag As Boolean, vAutoEnterQtyintoGridSaleInvoice As Boolean
Dim vExpAmount As Double
Dim vQtyLoose As Double, vTotalAmount As Double
Dim vProductDetailQty, vProductQtyPack, vProductQtyLoose, vProductBonus As Double, vAmount, vBottomPrice As Double
Dim vStrPara As String
Dim i As Integer, vNoofPrints As Byte, isWholeSale As Boolean, vShowStock As Boolean
Dim vCash, vCredit As Integer
Dim vMasterID As Long
Dim vStrDetail As String
Dim vMobileNo() As String, vMobile, vWhere As String
Dim vUpdateStock, vTradeOffer, vUseMultipleStore, vWholeSale, vAdvDiscPerFlag As Boolean
Dim vPrinter() As String
'----------------------------------

Private Sub SubCalculateBody()
   TxtDiscVal.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) * Val(TxtDiscPer.Text) / 100, 2)
   If Val(TxtDiscVal.Text) = 0 Then TxtDiscVal.Text = ""
   TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), 2)
   If vIsWSSaleTax = True And vIsWSDiscb4ST = True Then
       TxtSaleTaxVal.Text = Round(Val(TxtAmount.Text) * Val(TxtSaleTaxPer.Text) / 100, 3)
   ElseIf vIsWSSaleTax = True And vIsWSDiscb4ST = False Then
       TxtSaleTaxVal.Text = Round(Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
   ElseIf vIsRetailSaleTax = True Then
       TxtSaleTaxVal.Text = Round(Val(vUnitRetailPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
   Else
       TxtSaleTaxVal.Text = Round(Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
   End If
   TxtAmount.Text = Round(Val(TxtAmount.Text) + Val(TxtSC.Text) + Val(TxtSaleTaxVal.Text) - Val(TxtOffer.Text), 2)
   TxtDiscAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - Val(TxtDiscVal.Text) - Val(TxtExtraSchemeValue.Text), 2)
   If ObjRegistry.IsRoundFigure = True Then TxtAmount.Text = SelfRound(TxtAmount.Text)
'   If vTradeOffer = True Then CalculateValue
    Call CalculateValue
End Sub

Private Sub SubCalculateFooter()
   If TxtTotalAmount.Text = "" Then Exit Sub
'   If ObjRegistry.IsRoundFigure = True Then
'   TxtNetAmount.Text = SelfRound(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text) + Val(TxtServiceCharges.Text)
'   Else
'   TxtNetAmount.Text = Round(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text), 2) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text) + Val(TxtServiceCharges.Text)
'   End If
   
   
   If Val(TxtAdvTaxVal.Text) <> 0 Then
      If vAdvDiscPerFlag = False Then
         TxtAdvTaxPer.Text = Round((Val(TxtAdvTaxVal.Text) * 100) / IIf(Val(TxtNetAmount.Text) = 0, 1, Val(TxtNetAmount.Text)), 2)
      Else
         TxtAdvTaxVal.Text = SelfRound((Val(TxtNetAmount.Text) * Val(TxtAdvTaxPer.Text) / 100))
      End If
   End If
   TxtNetAmount.Text = Round(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text), 2) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text) + Val(TxtServiceCharges.Text) + Val(TxtAdvTaxVal.Text) + Val(TxtExtraTaxVal.Text)
   TxtNetAmount.Text = SelfRound(TxtNetAmount.Text)
'   TxtReceivedAmount.Text = Val(TxtNetAmount.Text)
   TxtTotalReceivable.Text = Abs(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", Val(TxtPreviousReceivable.Text) * -1, TxtPreviousReceivable.Text)))
   LblTtlPayable.Caption = IIf(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", Val(TxtPreviousReceivable.Text) * -1, TxtPreviousReceivable.Text)) < 0, "Total Payable", "Total Receivable")
End Sub

Private Sub SubClearSerialFields()
   TxtSerial.Text = ""
   TxtSerial.Enabled = False
   GridSerial.CancelUpdate
   GridSerial.RemoveAll
   GridSerial.AddNew
   GridSerial.Columns("Serial").Text = " "
   GridSerial.Update
End Sub

Private Sub SubClearMultipleStoreFields()
   FrmMultipleStore.Visible = False
   GridMultipleStore.CancelUpdate
   GridMultipleStore.RemoveAll
   GridMultipleStore.AddNew
   GridMultipleStore.Columns("ProductID").Text = " "
   GridMultipleStore.Update
   TxtMSTotalItems.Text = 0
   Grid.Enabled = True
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
   Dim vStrSQL As String, vStr As String
   'vStr = " and p.ProductID not In (Select ProductID from ZoneAllotment where ZoneID = " & vZoneID & " and PartyID is Not Null and PartyID <> '" & TxtCustomerID.Text & "')"
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.ParaInWholeSale = True
      SchProduct.ParaInWhere = " and isLocked = 0" & vStr
      SchProduct.ParainShowStock = vShowStock
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
   If Trim(TxtCode.Text) = "" Then Exit Function
   With CN.Execute("Select productid, serial from purchasebodyserial where serial = '" & TxtCode.Text & "'")
      If .RecordCount > 0 Then
         TxtCode.Text = !Productid
         TxtSerial.Text = !Serial
      End If
      .Close
   End With
   
   CmbPackName.Clear
   vStrSQL = "select * from ProductPacking pp inner join packings p on p.packingid = pp.packingid" & vbCrLf _
           + "left outer join ProductBarcodes b on b.productid = pp.productid" & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "pp.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')" & ""
              

   With CN.Execute(vStrSQL)
      CmbPackName.AddItem ""
      While Not .EOF
         CmbPackName.AddItem !Packingname
         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
         .MoveNext
      Wend
      .Close
   End With
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
        vStrSQL = "SELECT p.productid, Code, Qty, ProductName, PurPrice, WSPrice, RetailPrice, BottomPrice, " & vbCrLf _
           + " IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST, IsDiscB4TradeOffer, IsDiscB4ExtraScheme, isDiscB4SaleTax, TradeOffer1, TradeOffer2, ExtraSchemePer,  " & vbCrLf _
           + " TokenVal, SaleTaxPer, DiscPC, ServiceCharges, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.Salepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')" & " and isLocked = 0"
           
 
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = IIf(isWholeSale = True, !WSPrice, !RetailPrice)
         TxtRetailPrice.Text = IIf(isWholeSale = True, !WSPrice, !RetailPrice)
         vBottomPrice = IIf(IsNull(!BottomPrice), 0, !BottomPrice)
         ssql = "select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text) & ")"
         LblLastPurPrice.Caption = CN.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text) & ")").Fields(0).Value
         vIsWSDiscb4ST = !IsWSDiscb4ST
         vIsWSSaleTax = !IsWSSaleTax
         vIsRetailSaleTax = !IsRetailSaleTax
         ChkDiscB4TradeOffer.Value = Abs(!isDiscB4TradeOffer)
         ChkDiscB4ExtraScheme.Value = Abs(!isDiscB4ExtraScheme)
         ChkDiscB4SaleTax.Value = Abs(!isDiscB4SaleTax)
         TxtTradeOffer1.Text = IIf(IsNull(!TradeOffer1), 0, !TradeOffer1)
         TxtTradeOffer2.Text = IIf(IsNull(!TradeOffer2), 0, !TradeOffer2)
         TxtExtraSchemePer.Text = IIf(IsNull(!ExtraSchemePer), 0, !ExtraSchemePer)
         TxtSaleTaxPer.Text = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
         TxtTokenVal.Text = IIf(IsNull(!TokenVal), "", !TokenVal)
         LblRetailPrice.Caption = IIf(isWholeSale = False, !WSPrice, !RetailPrice)
         LblCaptionRetailPrice.Caption = IIf(isWholeSale = False, "WS Price", "Retail Price")
         LblPurPrice.Caption = "Purchase Price: " & !PurPrice
         TxtSC.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         LblLastPrice.Caption = CN.Execute("Select dbo.FunLastPrice('S','" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text) & ",'" & TxtCustomerID.Text & "')").Fields(0).Value

        With CN.Execute("select cost from currentstock where productid = " & Val(TxtProductID.Text))
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
         ChkIsProduct.Value = 1
         'vQty = 1
         vQty = IIf(Len(TxtCode.Text) <= 5 And IsNumeric(TxtCode.Text), 1, IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty))   'IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         'If vAutoEnterQtyintoGridSaleInvoice = True Then TxtQtyLoose.Text = IIf(Val(TxtQtyLoose.Text) = 0, 1, TxtQtyLoose.Text)
         If IsNull(!Packingname) Then
            vUnitPrice = IIf(isWholeSale = True, !WSPrice, !RetailPrice)
            vUnitRetailPrice = !RetailPrice
            TxtMultiplier.Text = ""
            CmbPackName.ListIndex = 0
         Else
            TxtMultiplier.Text = !Multiplier
            If !Multiplier <> 0 Then
               vUnitPrice = IIf(isWholeSale = True, !WSPrice, !RetailPrice) / !Multiplier
               vUnitRetailPrice = !RetailPrice / !Multiplier
            Else
               vUnitPrice = IIf(isWholeSale = True, !WSPrice, !RetailPrice)
               vUnitRetailPrice = !RetailPrice
            End If
            CmbPackName.Text = !Packingname
         End If
         If ObjRegistry.AllowDiscountOnSaleDistribution = True Then
            TxtDiscPC.Text = IIf(IsNull(!DiscPC), "", !DiscPC)
            If vUnitPrice = 0 Then
               TxtDiscPer.Text = "0"
            Else
               TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
            End If
         End If
'         If ObjRegistry.AlertAllocateProduct = True Then
'            If Trim(TxtCustomerID.Text) <> "" Then
'               vStrSQL = "Select * " & vbCrLf _
'                     + " from CustomerProductPrice" & vbCrLf _
'                     + " where CustomerID = '" & TxtCustomerID.Text & "' and ProductID = '" & TxtProductID.Text & "'"
'
'               With cn.Execute(vStrSQL)
'                  If .RecordCount > 0 Then
'                     TxtPrice.Text = !Price
'                     vUnitPrice = !Price
'                     TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
'                     TxtDiscPC.Text = Round((vUnitPrice * Val(TxtDiscPer.Text) / 100), 4)
'                  Else
'                     MsgBox "Allocate this Product to the Customer first.", vbInformation + vbOKOnly, "Information"
'                     FunSelectProduct = False
'                     Exit Function
'                  End If
'               End With
'            End If
'         End If
         If ObjRegistry.AutoApplyPartyLastPrice Then
            If Trim(TxtCustomerID.Text) <> "" Then
               vStrSQL = "Select Top 1 * " & vbCrLf _
                     + " from SaleHeader h inner join SaleBody b on H.SID = B.SID" & vbCrLf _
                     + " left outer join ProductPacking pp on pp.packingid = b.packingid and pp.productid = b.productid" & vbCrLf _
                     + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
                     + " where h.CustomerID = " & Val(TxtCustomerID.Text) & " and b.ProductID = " & Val(TxtProductID.Text) & vbCrLf _
                     + " Order by h.BillDate Desc, h.BillID desc"
               With CN.Execute(vStrSQL)
                  If .RecordCount > 0 Then
                     If IsNull(!Packingname) Then
                      vUnitPrice = !Price
                      TxtMultiplier.Text = ""
                        CmbPackName.ListIndex = 0
                     Else
                        If !Multiplier <> 0 Then
                          vUnitPrice = !Price / !Multiplier
                        End If
                        CmbPackName.Text = !Packingname
                     End If
                     TxtMultiplier.Text = IIf(IsNull(!Multiplier), "", !Multiplier)
                     TxtPrice.Text = !Price
                     TxtDiscPC.Text = !DiscPC
                     TxtDiscPer.Text = !DiscPer
                     TxtTradeOffer1.Text = IIf(IsNull(!TradeOffer1), "", !TradeOffer1)
                     TxtTradeOffer2.Text = IIf(IsNull(!TradeOffer2), "", !TradeOffer2)
                     TxtExtraSchemePer.Text = IIf(IsNull(!ExtraSchemePer), "", !ExtraSchemePer)
                     ChkDiscB4TradeOffer.Value = Abs(!isDiscB4TradeOffer)
                     ChkDiscB4ExtraScheme.Value = Abs(!isDiscB4ExtraScheme)
                     ChkDiscB4SaleTax.Value = Abs(!isDiscB4SaleTax)
                  End If
               End With
            End If
         End If
         If ObjRegistry.AutoApplyPartyLastDiscount Then
            If Trim(TxtCustomerID.Text) <> "" And Val(TxtCustomerID.Text) <> 621 Then
               vStrSQL = "Select Top 1 * " & vbCrLf _
                     + " from SaleHeader h inner join SaleBody b on H.SID = B.SID" & vbCrLf _
                     + " left outer join ProductPacking pp on pp.packingid = b.packingid and pp.productid = b.productid" & vbCrLf _
                     + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
                     + " where h.CustomerID = " & Val(TxtCustomerID.Text) & " and b.ProductID = " & Val(TxtProductID.Text) & vbCrLf _
                     + " Order by h.BillDate Desc, h.BillID desc"
               With CN.Execute(vStrSQL)
                  If .RecordCount > 0 Then
                     TxtDiscPC.Text = !DiscPC
                     TxtDiscPer.Text = !DiscPer
                     TxtTradeOffer1.Text = IIf(IsNull(!TradeOffer1), "", !TradeOffer1)
                     TxtTradeOffer2.Text = IIf(IsNull(!TradeOffer2), "", !TradeOffer2)
                     TxtExtraSchemePer.Text = IIf(IsNull(!ExtraSchemePer), "", !ExtraSchemePer)
                     ChkDiscB4TradeOffer.Value = Abs(!isDiscB4TradeOffer)
                     ChkDiscB4ExtraScheme.Value = Abs(!isDiscB4ExtraScheme)
                     ChkDiscB4SaleTax.Value = Abs(!isDiscB4SaleTax)
                  End If
               End With
            End If
         End If
         
'         ''' latest Comment
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & Val(TxtProductID.Text)
            With CN.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Caption = LblStock.Caption & " " & " Total Qty: " & vQtyLoose
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
         LblPurPrice.Visible = ObjRegistry.ShowPurPrice
'         ''' latest Comment
         
'         With cn.Execute("select Productid, QtyLoose from CurrentStockStore where ProductID ='" & TxtProductID.Text & "' and StoreID = " & TxtStoreID.Text)
'            If .RecordCount > 0 Then
'               vQtyLoose = !QtyLoose
'               LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & !Productid & "',Floor(" & !QtyLoose & "))").Fields(0).Value
''               LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'               LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & !Productid & "',Floor(" & !QtyLoose & "))").Fields(0).Value
''               LblStock.Caption = LblStock.Caption & " " & "Loose"
'            Else
'               vQtyLoose = 0
'               LblStock.Caption = 0
'            End If
'         End With
         
         
         If Val(TxtCustomerID.Text) <> 621 And TxtCustomerID.Text <> "" Then
            PopulateDataToHistoryGrid
            FrmHistory.Visible = True
            FrmHistory.ZOrder 0
            GridHistory.Visible = True
            GridHistory.ZOrder 0
         Else
            FrmHistory.Visible = False
         End If
         
         If ObjRegistry.ShowAllPrices Then
            PopulateDataToPriceGrid
            FrmProductPrices.Visible = True
         Else
            FrmProductPrices.Visible = False
         End If
         
         If ObjRegistry.BatchNoVisible Then
            PopulateDataToGridExpiry
            FrmExpiry.Visible = True
            FrmExpiry.ZOrder 0
            GridExpiry.Visible = True
            GridExpiry.ZOrder 0
         End If
         
         If ObjRegistry.NegativeSale = False Then
            If vQtyLoose <= 0 Then
               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
               FunSelectProduct = False
               Exit Function
            End If
         End If
         vExpiryTime = 0
         With CN.Execute("Select dbo.GetExpiryTime(" & Val(TxtProductID.Text) & ", " & IIf(TxtBatchNo.Text = "", "Null", "'" & TxtBatchNo.Text & "'") & " , GetDate()) as Day ")
            If .RecordCount > 0 Then
               vExpiryTime = !Day
            End If
         End With
         
         SubCalculateBody
''         VStrSQL = "select isnull(dbo.FunStock('" & TxtProductID.Text & "'," & TxtStoreID.Text & "," & Val(TxtBillID.Text) & "," & Val(0) & "," & Val(TxtBillID.Text) & "," & Val(0) & "," & Val(0) & "," & Val(0) & ",'" & DateAdd("D", 1, DtpBillDate.DateValue) & "'," & Val(0) & "),0)"
''         vQtyLoose = cn.Execute(VStrSQL).Fields(0).Value
''         LblStock.Caption = vQtyLoose
         'Char.Speak TxtProductName.Text
         FunSelectProduct = True
'         If CmbPackName.ListCount <= 1 Then CmbPackName.SetFocus
         CmbPackName.SetFocus
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
         TxtRetailPrice.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtSC.Text = ""
         TxtAmount.Text = ""
         TxtTradeOffer1.Text = ""
         TxtTradeOffer2.Text = ""
         TxtExtraSchemePer.Text = ""
         TxtTradeOfferValue.Text = ""
         TxtExtraSchemeValue.Text = ""
         ChkDiscB4TradeOffer.Value = 0
         ChkDiscB4ExtraScheme.Value = 0
         ChkDiscB4SaleTax.Value = 0
         LblStock.Visible = False
         LblStockCaption.Visible = False
         LblCaptionRetailPrice.Visible = False
         LblRetailPrice.Visible = False
         LblPurPrice.Visible = False
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub PopulateDataToHistoryGrid()
   If Val(TxtCustomerID.Text) = 621 Then Exit Sub
     If ObjRegistry.ShowHistoryofAllCustomer = True Then
        vWhere = " where 1=1  and b.productid = " & Val(TxtProductID.Text) & " and h.StoreID = " & TxtStoreID.Text & " order by b.BillDate Desc"
     Else
        vWhere = " where h.CustomerID = " & Val(TxtCustomerID.Text) & " and b.productid = " & Val(TxtProductID.Text) & " and h.StoreID = " & TxtStoreID.Text & " order by b.BillDate Desc"
     End If
    
      ssql = "select top 3 pt.PartyName, CustomerID, code, b.* " & vbCrLf & _
      " from SaleHeader h inner join Salebody b on H.SID = B.SID" & vbCrLf & _
      " inner join Parties pt on pt.PartyID = h.CustomerID " & vWhere
      
      With CN.Execute(ssql)
         GridHistory.Redraw = False
         GridHistory.MoveFirst
         GridHistory.RemoveAll
         GridHistory.AllowAddNew = True
         While Not .EOF
            GridHistory.AddNew
            GridHistory.Columns("ID").Text = !CustomerID
            GridHistory.Columns("Name").Text = !partyname
            GridHistory.Columns("BillID").Text = !BillID
            GridHistory.Columns("Date").Text = !BillDate
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
'            GridHistory.Columns("ExpiryDate").Value = IIf(IsNull(!ExpiryDate), "", !ExpiryDate)
            GridHistory.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
            GridHistory.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            GridHistory.Columns("QtyLoose").Value = !Qty
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
            GridProductPrices.Columns("Description").Text = IIf(IsNull(!desc1), "", !desc1)
            GridProductPrices.Columns("Pur").Value = !PurPrice
            GridProductPrices.Columns("List").Value = IIf(IsNull(!Listprice), "0", !Listprice)
            GridProductPrices.Columns("WS").Value = !WSPrice
            GridProductPrices.Columns("Retail").Value = !RetailPrice
            .MoveNext
         Wend
         .Close
         GridProductPrices.MoveFirst
         GridProductPrices.Redraw = True
      End With
End Sub

Private Sub BtnAddCustomer_Click()
   DefCustomers.Show
End Sub

Private Sub BtnBatchPrint_Click()
RptBatchPrint.Show vbModal
End Sub

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
'   If Grid.Rows <= 1 And TxtProductID.Text = "" Then Exit Sub
        
   vStrDetail = ""
   With Grid
      .Redraw = False
      .MoveFirst
      For vCounter = 1 To .rows
         If Trim(.Columns("Productid").Text) <> "" Then
            vStrDetail = vStrDetail & " (P" & .Columns("ProductID").Text & IIf(Val(.Columns("Pack").Value) = 0, "", " M" & .Columns("Pack").Value) & IIf(Val(.Columns("QtyPack").Value) = 0, "", " QP" & .Columns("QtyPack").Value) & IIf(Val(.Columns("QtyLoose").Value) = 0, "", " QL" & .Columns("QtyLoose").Value) & IIf(Val(.Columns("Bonus").Value) = 0, "", " QB" & .Columns("Bonus").Value) & " A" & .Columns("Amount").Text & ")"
         End If
         .MoveNext
      Next vCounter
      .Redraw = True
   End With
    '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave And vIsNewRecord = True And Grid.rows > 1 Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = " Cleared ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt" & TxtNetAmount.Text
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               CN.Execute ssql
            End If
         Next
   End If
    
    '''''''''''''''''' ActivityLogBin For Clear Action
'      Call DeleteTempActivityLogBin(vRandomID)
      vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From salebody where SID=" & Val(TxtSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmSaleInvoiceDIS, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Cleared Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
      
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmSaleInvoiceDIS, eClearSavedRecord, TxtBillID.Text, DtpBillDate.DateValue, vGridRows & " Product/s Cleared")
      Grid.Redraw = True
  ''''''''''''''''''
      
      
   FormStatus = NewMode
   'cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   
    ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSaleInvoice", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
      vStrSQL = "select * from SaleHeader where Tag is not null And SID=" & Val(TxtSID.Text) & " and Billdate='" & DtpBillDate.DateValue & "'"
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
   
   
   vMaxBinID = FunGetMaxBinID
   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   cn.Execute ("Insert Into Bin_SaleHeader Select " & vMaxBinID & ",'" & Date & "',* from SaleHeader Where BillID = " & TxtBillID.Text & " And BillDate ='" & DtpBillDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into Bin_SaleBody Select " & vMaxBinID & ",'" & Date & "', * from SaleBody Where BillID = " & TxtBillID.Text & " And BillDate ='" & DtpBillDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Serial''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into Bin_SaleBodySerial Select " & vMaxBinID & ",'" & Date & "', * from SaleBodySerial Where BillID = " & TxtBillID.Text & " And BillDate ='" & DtpBillDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin ProductOffer''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into Bin_SaleBodyOffer Select " & vMaxBinID & ",'" & Date & "', * from SaleBodyOffer Where BillID = " & TxtBillID.Text & " And BillDate ='" & DtpBillDate.DateValue & "'")
'
'   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  Call BinData
  Call ActivityLogBin("", eFrmSaleInvoiceDIS, eDelete, TxtBillID.Text, DtpBillDate.DateValue, Grid.rows - 1 & " Product/s Deleted Amount: " & Val(TxtNetAmount.Text))
   ''''''''''''''''''''''''''Delete Product OFfer'''''''''''''''''''''
   GridOffer.Redraw = False
   GridOffer.MoveFirst
   For vCounter = 1 To GridOffer.rows
      If Trim(GridOffer.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from SaleBodyOffer where BillID = " & Val(TxtBillID.Text) & " And BillDate ='" & DtpBillDate.DateValue & "' and productid = " & Val(GridOffer.Columns("Productid").Text)
      End If
      GridOffer.MoveNext
   Next vCounter
   GridOffer.RemoveAll
   GridOffer.Redraw = True
   GridOffer.ZOrder 1
   ''''''''''''''''''''''''''Delete Serials'''''''''''''''''''''
    If RsBodySerial.RecordCount > 0 Then
        RsBodySerial.MoveFirst
        For vCounter = 1 To RsBodySerial.RecordCount
            CN.Execute "Delete from SaleBodySerial where BillID = " & Val(TxtSID.Text) & " And BillDate ='" & DtpBillDate.DateValue & "' and productid = " & RsBodySerial!Productid & " and Serial ='" & RsBodySerial!Serial & "'"
            RsBodySerial.MoveNext
        Next vCounter
    End If
   ''''''''''''''''''''''''''Delete Sale Body'''''''''''''''''''''
   vStrDetail = ""
   Grid.Redraw = False
   Grid.MoveFirst
'   Call ActivityLogSale("Sale Invoice", eDelete, TxtBillID.Text, DtpBillDate.DateValue)
   vStrSQL = "Delete from SaleBody where SID = " & Val(TxtSID.Text)
   CN.Execute vStrSQL
   For vCounter = 1 To Grid.rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         vQtyLoose = (Val(Grid.Columns("Pack").Text) * Val(Grid.Columns("QtyPack").Text)) + Val(Grid.Columns("QtyLoose").Text) + Val(Grid.Columns("Bonus").Text)
         CN.Execute "Exec UpdateStockPlus " & TxtStoreID.Text & "," & Val(Grid.Columns("ProductID").Text) & "," & vQtyLoose & "," & Val(TxtBillID.Text) & ",'" & DtpBillDate.DateValue & "'"
         
         vStrDetail = vStrDetail & " (P" & Grid.Columns("ProductID").Text & IIf(Val(Grid.Columns("Pack").Value) = 0, "", " M" & Grid.Columns("Pack").Value) & IIf(Val(Grid.Columns("QtyPack").Value) = 0, "", " QP" & Grid.Columns("QtyPack").Value) & IIf(Val(Grid.Columns("QtyLoose").Value) = 0, "", " QL" & Grid.Columns("QtyLoose").Value) & IIf(Val(Grid.Columns("Bonus").Value) = 0, "", " QB" & Grid.Columns("Bonus").Value) & " A" & Grid.Columns("Amount").Text & ")"
'          cn.Execute ("Insert Into Bin_SaleBody Select " & FunGetMaxBinID & ", * from SaleBody Where BillID = " & TxtBillID.Text & " And BillDate ='" & DtpBillDate.DateValue & "' and productid ='" & Grid.Columns("Productid").Text & "'")
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   
   ''''''''''''''''''''''''''Delete SaleBody Store'''''''''''''''''''''
   
   GridMultipleStore.Redraw = False
'   GridMultipleStore.MoveFirst
   vStrSQL = "Delete from SaleBodyStore where SID = " & Val(TxtSID.Text)
   CN.Execute vStrSQL
   GridMultipleStore.RemoveAll
   GridMultipleStore.Redraw = True
   
   '''''''''''''''''''''''''''''''''''''''Delete Expense'''''''''''''''''''''''''''''''''''''''
   CN.Execute "Delete from SaleExpense where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   
   '''''''''''''''''''''''''''''''''''''''Delete Header'''''''''''''''''''''''''''''''''''''''
   CN.Execute "Delete from SaleHeader where SID = " & Val(TxtSID.Text)
   
   CN.Execute ("Update SaleOrderHeader set IsSale = 0 Where OrderID = " & Val(TxtOrderID.Text) & "And Orderdate ='" & DtpOrderDate.DateValue & "'")
    If Grid.rows > 1 Then
'      vStrSQL = "INSERT INTO ActivityLog(userno,FormType,EntryDate,Description,isnew,isedit,isdelete,isClear) values(" & vUser & ",'Sale Invoice', GetDate()," & "'BillID = " & TxtBillID.Text & " BillDate = " & DtpBillDate.DateValue & " Clear' ,0,0,0,1" & ")"
    'cnPOS.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Clear'," & vUser & ")")
      CN.Execute (vStrSQL)
'      Call Sub_Bin_Save
   End If
   
    '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = " Deleted ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt" & TxtNetAmount.Text
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               CN.Execute ssql
            End If
         Next
   End If
    '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = " Cleared ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt" & TxtNetAmount.Text
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
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

Private Sub Sub_Bin_Save()
  On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnEmployee_Click()
   On Error GoTo ErrorHandler
   If FunSelectEmployee(ssButton, False) = True Then
      If TxtMemberID.Visible = True Then TxtMemberID.SetFocus Else If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      TxtEmployeeID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      TxtBillNo.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

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
         
'         RsBody.AddNew
'         RsBody!Productid = !Productid
'         RsBody!Code = !Productid
         
         Grid.Columns("QtyLoose").Value = !QtyLoose
         Grid.Columns("Price").Value = !Price
         
         Grid.Columns("RetailPrice").Value = 0
         Grid.Columns("IsWSDiscb4ST").Value = 0
         Grid.Columns("IsWSSaleTax").Value = 0
         Grid.Columns("IsRetailSaleTax").Value = 0
         
         Grid.Columns("Amount").Value = (!Price * !QtyLoose)
         Grid.Columns("IsProduct").Value = 1
         
         
                Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
'                  Grid.Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
                  Grid.Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, "", Val(TxtMultiplier.Text))
                  Grid.Columns("GrossQty").Value = IIf(Val(TxtGrossQty.Text) = 0, Null, Val(TxtGrossQty.Text))
                  Grid.Columns("GrossUnit").Value = IIf(Val(TxtGrossUnit.Text) = 0, Null, Val(TxtGrossUnit.Text))
                  Grid.Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
'                  Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
                  Grid.Columns("Bonus").Value = Val(TxtBonus.Text)
'                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("RetailPrice").Value = Val(TxtRetailPrice.Text)
                  Grid.Columns("IsWSDiscb4ST").Value = vIsWSDiscb4ST
                  Grid.Columns("IsWSSaleTax").Value = vIsWSSaleTax
                  Grid.Columns("IsRetailSaleTax").Value = vIsRetailSaleTax
                  Grid.Columns("TokenVal").Value = IIf(Val(TxtTokenVal.Text) = 0, 0, Val(TxtTokenVal.Text))
                  Grid.Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  Grid.Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  Grid.Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  Grid.Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  Grid.Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  Grid.Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  Grid.Columns("isDiscB4TradeOffer").Value = Abs(ChkDiscB4TradeOffer.Value)
                  Grid.Columns("isDiscB4ExtraScheme").Value = Abs(ChkDiscB4ExtraScheme.Value)
                  Grid.Columns("isDiscB4SaleTax").Value = Abs(ChkDiscB4SaleTax.Value)
                  Grid.Columns("TradeOffer1").Value = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
                  Grid.Columns("TradeOffer2").Value = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
                  Grid.Columns("ExtraSchemePer").Value = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
                  Grid.Columns("TradeValue").Value = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
                  Grid.Columns("ExtraSchemeValue").Value = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
                  Grid.Columns("SC").Value = Val(TxtSC.Text)
'                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("Cost").Value = Val(TxtCost.Text)
                  Grid.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
                  Grid.Columns("ExpiryTime").Value = Val(vExpiryTime)
                  Grid.Columns("ExpiryDate").Value = TxtExpiryDate.Text
                  
         '''''
''         RsBody!Multiplier = Null
''         RsBody!QtyPack = Null
''         RsBody!Qty = !QtyLoose
''         RsBody!Bonus = Null
''         RsBody!Price = !Price
''         RsBody!isProduct = 1 '!isProduct
''
''         RsBody!RetailPrice = 0
''         RsBody!IsWSDiscb4ST = 0
''         RsBody!IsWSSaleTax = 0
''         RsBody!IsRetailSaleTax = 0
''
''         RsBody!Cost = 0
''         RsBody!DiscPC = 0
''         RsBody!Offer = Null
''         RsBody!SaleTaxPer = Null
''         RsBody!SaleTaxval = Null
''         RsBody!DiscPer = 0
''         RsBody!DiscVal = 0
''         RsBody!Amount = (!Price * !QtyLoose)
''         RsBody.Update
         ''''
         
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + (!Price * !QtyLoose)
         TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + !QtyLoose
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

Private Sub BtnPrintWarranty_Click()
   On Error GoTo ErrorHandler
   With CN.Execute("Select distinct o.OrganizationID, ReportName from SaleBody b inner join Products p on p.productid = b.productid inner join Organizations o on p.OrganizationID = o.OrganizationID where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'")
      For i = 1 To .RecordCount
         vStrSQL = "Select h.BillID, h.BillDate, EntryDate, h.OrganizationID, OrganizationName, Customerid, isnull(Pr.PartyName,AccountName) + ' - ' + H.CustomerID as Customer_Name_ID, Pr.Address, StoreName, BiltyNo, VehicleNo, h.Description," & vbCrLf _
            + " Isnull(H.BillDiscPer, 0) BillDiscPer, Isnull(H.BillDisc,0) BillDisc, isnull(OtherCharges,0) as OtherCharges," & vbCrLf _
            + " TotalAmount,  isnull(TotalExpense,0) as TotalExpense, b.ProductID as Code, " & IIf(isUrdu = True, "p.ProductName1", "p.ProductName") & " as ProductName, dbo.FunSaleBodySerial(b.BillID,b.BillDate, b.ProductId) Serial," & vbCrLf _
            + " dbo.FunSaleBodyOffer(b.BillID,b.BillDate, b.ProductId) ProductOffer, isnull(QtyPack,0)QtyPack, isnull(Multiplier,0)Multiplier, Qty," & vbCrLf _
            + " Bonus,b.DiscPc, b.DiscPer, DiscVal, Offer, b.SaleTaxPer, SaleTaxval," & vbCrLf _
            + " h.Empid, empname, price, Amount, previousAmount, CashReceived, b.RetailPrice, b.BatchNo " & vbCrLf _
            + " from SaleBody b inner join SaleHeader h on H.SID = B.SID" & vbCrLf _
            + " inner join products p on b.productid = p.productid" & vbCrLf _
            + " left outer join Organizations o on o.OrganizationID = p.OrganizationID" & vbCrLf _
            + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
            + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
            + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
            + " where h.SID = " & Val(TxtSID.Text) & " and p.OrganizationID = " & !OrganizationID
       
          If RsReport.State = adStateOpen Then RsReport.Close
          RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
         
          RptReportViewer.Report.SelectPrinter "abc", "xyz", "ghi"
          
          
          
          Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\" & !ReportName & ".rpt")
          'Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\SaleInvoiceWarranty1.rpt")
          RptReportViewer.Report.PaperSize = crPaperA4
          RptReportViewer.Report.PaperOrientation = crPortrait
          
          RptReportViewer.Report.DiscardSavedData
          RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
          RptReportViewer.Show vbModal, Me
          'cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
          .MoveNext
      Next i
      If .RecordCount = 0 Then MsgBox "Record not Found", vbInformation, "Sale Invoice"
      .Close
   End With
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
   If FunSelectPurchase(ssButton, False) = True Then
      TxtDescription.SetFocus
   Else
'      If TxtPurID.Enabled And TxtPurID.Visible Then TxtPurID.SetFocus
   End If
End Sub

Private Function FunSelectPurchase(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchPurchase.ParaInPurchasedate = DtpPurchaseDate.DateValue
        SchPurchase.Show vbModal, Me
        If SchPurchase.ParaOutPurchaseID = "" Then FunSelectPurchase = False: Exit Function
        TxtPurID.Text = SchPurchase.ParaOutPurchaseID
        DtpPurchaseDate.DateValue = SchPurchase.ParaOutPurchaseDate
    End If
    '---------------------------
    vStrSQL = "Select * from PurchaseHeader where PurID=" & Val(TxtPurID.Text) & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          If MsgBox("Do you want to Get Transfer Information from Purchase.", vbQuestion + vbYesNo, "Alert") = vbYes Then
            Call GetPurchase
          End If
          FunSelectPurchase = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          TxtPurID.Text = ""
          DtpPurchaseDate.DateValue = ""
          FunSelectPurchase = False
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub GetPurchase()
   On Error GoTo ErrorHandler
'   TxtBillID.Text = FunGetMaxID
   ssql = "select h.*, OrganizationName,StoreName FROM purchaseheader h left outer join Organizations o on o.OrganizationID = h.OrganizationID  inner join stores s on s.storeid = h.storeid Where h.purID=" & Val(TxtPurID.Text) & " and purchaseDate='" & DtpPurchaseDate.DateValue & "'"
   With CN.Execute(ssql)
      If Not .BOF Then
          DtpPurchaseDate.DateValue = !PurchaseDate
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          TxtBillNo.Text = IIf(IsNull(!BillNo), "", !BillNo)
          TxtBiltyNo.Text = IIf(IsNull(!BiltyNo), "", !BiltyNo)
          TxtVehicleNo.Text = IIf(IsNull(!VehicleNo), "", !VehicleNo)
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtTotalAmount.Text = !TotalAmount
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtTotalExpense.Text = IIf(IsNull(!TotalExpense), "", !TotalExpense)
'          TxtPaidAmount.Text = IIf(IsNull(!PAIDAMOUNT), "", !PAIDAMOUNT)
'          TxtReceivedAmount.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtPreviousReceivable.Text = IIf(IsNull(!PreviousAmount), "", !PreviousAmount)
          lblPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
          LblTtlPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Total Receivable", "Total Payable")
          TxtPreviousReceivable.Text = Abs(Val(TxtPreviousReceivable.Text))

      End If
      .Close
   End With
   Call PopulatePurchaseDataToGrid
'   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub
Private Sub PopulatePurchaseDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleBody where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      
'      ssql = " select pb.ProductID, ProductName, QtyPack - isnull(UPack,0) as RQUsePurPriceo9.pm;lmkmnn  bvtyPack, Qtyloose - isnull(UQty,0) as RQty, Bonus - isnull(UBonus,0) as RBonus, pb.*" & vbCrLf _
      + " from (select b.PurID, b.PurchaseDate, ProductID, Sum(Qtyloose) as UQty, Sum(QtyPack) as UPack, Sum(Bonus) as UBonus from PurchaseBody b inner join PurchaseHeader h on h.PurID = b.pURID and h.PurchaseDate = b.PurchaseDate Group By b.PurID, b.PurchaseDate, ProductID) b " & vbCrLf _
      + " right outer join PurchaseBody  pb on pb.PurID = b.PurID and pb.PurchaseDate = b.PurchaseDate and b.ProductID = pb.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = pb.productid" & vbCrLf _
      + " where pb.PurID = " & Val(TxtPurID.Text) & " and pb.PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
      vWholeSale = 0
      If Trim(TxtCustomerID.Text) <> "" Then
         vWholeSale = CN.Execute("Select isnull(isWholeSale,0) isWholeSale from parties where partyid = " & Val(TxtCustomerID.Text)).Fields("isWholeSale").Value
      End If
      ssql = "select p.productname, p.Retailprice product_RetailPrice, p.WSprice product_WSPrice, p.PurPrice product_PurPrice, b.* from PurchaseBody b join products p on p.productid = b.productid where PurID=" & Val(TxtPurID.Text) & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         While Not .EOF
            Grid.AddNew
'            RsBody.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = !Code
            Grid.Columns("ProductName").Text = !ProductName
            
'            RsBody!Productid = !Productid
'            RsBody!code = !code
            
            If !PackingID = 0 Or IsNull(!PackingID) Then
               Grid.Columns("PackingID").Value = ""
            Else
               Grid.Columns("PackingID").Value = !PackingID
'               RsBody!PackingID = !PackingID
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
            If ObjRegistry.UsePurPrice = True Then
               Grid.Columns("Price").Value = !product_PurPrice
               Grid.Columns("RetailPrice").Value = !product_PurPrice
            Else
               Grid.Columns("Price").Value = IIf(vWholeSale, !product_WSPrice, !product_RetailPrice)
               Grid.Columns("RetailPrice").Value = IIf(vWholeSale, !product_WSPrice, !product_RetailPrice)
            End If
            Grid.Columns("Cost").Value = 0
            Grid.Columns("isProduct").Value = 1 '!isProduct
            
            Grid.Columns("IsWSDiscb4ST").Value = !IsWSDiscb4ST
            Grid.Columns("IsWSSaleTax").Value = !IsWSSaleTax
            Grid.Columns("IsRetailSaleTax").Value = !IsRetailSaleTax
            Grid.Columns("TokenVal").Value = ""
            Grid.Columns("DiscPC").Value = 0
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = 0
            Grid.Columns("DiscVal").Value = 0
            Grid.Columns("Amount").Value = ((Grid.Columns("RetailPrice").Value / Val(IIf(IsNull(!Multiplier), "1", !Multiplier)))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !QtyLoose)  '!Amount
'            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(((!product_RetailPrice / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !QtyLoose))
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(((Grid.Columns("RetailPrice").Value / Val(IIf(IsNull(!Multiplier), "1", !Multiplier)))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !QtyLoose))
            TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + !QtyLoose + IIf(IsNull(!Bonus), "0", !Bonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack))
            
            'TxtAmount.Text = Round((Val(vUnitPrice) - Val(TxtDiscPC.Text)) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 3)

'            RsBody!Multiplier = !Multiplier
'            RsBody!QtyPack = IIf(IsNull(!QtyPack), 0, !QtyPack)
'            RsBody!Qty = !Qtyloose
'            RsBody!Bonus = !Bonus 'IIf(IsNull(), Null, !Bonus)
'            RsBody!Price = !product_RetailPrice
'            RsBody!Cost = 0
'            RsBody!isProduct = 1 '!isProduct
'            RsBody!RetailPrice = !product_RetailPrice
'            RsBody!IsWSDiscb4ST = !IsWSDiscb4ST
'            RsBody!IsWSSaleTax = !IsWSSaleTax
'            RsBody!IsRetailSaleTax = !IsRetailSaleTax
'            RsBody!TokenVal = 0
'            RsBody!DiscPC = IIf(IsNull(!DiscPC), 0, !DiscPC)
'            RsBody!Offer = !Offer
'            RsBody!SaleTaxPer = !SaleTaxPer
'            RsBody!SaleTaxval = !SaleTaxval
'            RsBody!DiscPer = !DiscPer
'            RsBody!DiscVal = Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !Qtyloose) 'IIf(IsNull(!DiscVal), "", !DiscVal)
'            RsBody!Amount = ((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !Qtyloose) '!Amount
'            RsBody.Update
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
'   End If

   
End Sub

Private Sub BtnSaleOrder_Click()
   On Error GoTo ErrorHandler
   SchSaleOrder.ParaInOrderDate = DtpOrderDate.DateValue
   SchSaleOrder.Show vbModal
   If SchSaleOrder.ParaOutOrderID <> -1 Then
      TxtOrderID.Text = SchSaleOrder.ParaOutOrderID
      'Dim a
      'a = Split(SchSale.ParaOutBillDate, "/")
      DtpOrderDate.DateValue = SchSaleOrder.ParaOutOrderDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
'      cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetSaleOrder
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSaveAS_Click()
vIsNewRecord = True
DtpBillDate.Date = Now
Call BtnSave_Click
End Sub

Private Sub BtnSyllabus_Click()
   If FunSelectSyllabus(ssButton, False) = True Then
      TxtCustomerID.SetFocus
   Else
      TxtSyllabusID.SetFocus
   End If
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


Private Sub CmbMSStore_Click()
On Error GoTo ErrorHandler
   If CmbMSStore.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbMSStore.Name Then Exit Sub
    ''' latest Comment
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & CmbMSStore.ItemData(CmbMSStore.ListIndex) & " and Productid = " & Val(TxtProductID.Text)
            With CN.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & CmbMSStore.ItemData(CmbMSStore.ListIndex) & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Caption = LblStock.Caption & " " & " Total Qty: " & vQtyLoose
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
'   If Trim(CmbMSPackName.Text) = "" Then TxtMSQtyLoose.SetFocus
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub



Private Sub DtpDispatchDate_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> DtpDispatchDate.Name Then Exit Sub
   DtpPromiseDate.DateValue = DateAdd("d", Val(TxtTerms.Text), DtpDispatchDate.DateValue)
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub DtpPromiseDate_Change()
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub DtpPromiseDate_DblClick()
   DtpPromiseDate.DateValue = Null
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub DtpExpiryInvoice_Change()
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub DtpExpiryInvoice_DblClick()
   DtpExpiryInvoice.DateValue = Null
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Form_Activate()
  If Trim(ParaCustID) <> "" Then
       TxtCustomerID.Text = ParaCustID
       TxtCustomerName.Text = ParaCustName
       If FunSelectCustomer(1, False) = True Then
       End If
         ParaCustID = ""
         ParaCustName = ""
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF7 Or KeyCode = vbKeyF6 Or KeyCode = vbKeyF5 Or KeyCode = vbKeyF4 Then
'      LblLastPurPrice.Visible = True
      LblCost.Visible = False
   End If
End Sub

Private Sub Grid_RowLoaded(ByVal Bookmark As Variant)
   With Grid
      If Val(.Columns("ExpiryTime").Value) = 0 Then
         .Columns("ProductName").CellStyleSet ""
      ElseIf Val(.Columns("ExpiryTime").Value) <= 90 And Val(.Columns("ExpiryTime").Value) > 30 Then
         .Columns("ProductName").CellStyleSet "Green"
      ElseIf Val(.Columns("ExpiryTime").Value) <= 30 And Val(.Columns("ExpiryTime").Value) > 0 Then
         .Columns("ProductName").CellStyleSet "Orange"
      ElseIf Val(.Columns("ExpiryTime").Value) < 0 Then
         .Columns("ProductName").CellStyleSet "Red"
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

Private Sub GridExpiry_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   TxtBatchNo.Text = GridExpiry.Columns("BatchNo").Text
   TxtExpiryDate.Text = GridExpiry.Columns("ExpiryDate").Text
End Sub

Private Sub TxtAmount_Change()
If ActiveControl.Name <> TxtAmount.Name Then Exit Sub
If ObjRegistry.ChangeQtyOnChangedPrice = True Then
   If Val(TxtMultiplier.Text) <> 0 Then
      vUnitPrice = Round(Val(TxtPrice.Text) / Val(TxtMultiplier.Text), 3)
   Else
      vUnitPrice = Val(TxtPrice.Text)
   End If
      vAmount = Val(TxtAmount.Text) + Val(TxtDiscVal.Text)
      vQtyLoose = Round(vAmount / IIf(Val(vUnitPrice) = 0, IIf(vAmount = 0, 1, vAmount), Val(vUnitPrice)), 3)
      If CmbPackName.Text <> "" Then
         TxtQtyPack.Text = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         TxtQtyLoose.Text = CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & "," & vQtyLoose & ")").Fields(0).Value
      Else
         TxtQtyLoose.Text = vQtyLoose
      End If
   End If
End Sub

Private Sub TxtAmount_LostFocus()
Select Case ActiveControl.Name
   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtBonus.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtOffer.Name, TxtSaleTaxPer.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtBatchNo_LostFocus()
   
   With CN.Execute("Select dbo.GetExpiryDate(" & Val(TxtProductID.Text) & ",'" & TxtBatchNo.Text & "') as ExpiryDate")
      If .RecordCount > 0 Then
          TxtExpiryDate.Text = IIf(IsNull(!ExpiryDate), "", !ExpiryDate)
      End If
   End With
End Sub

Private Sub TxtBillNo_Validate(Cancel As Boolean)
   
   On Error GoTo ErrorHandler
   Cancel = ObjRegistry.AllowManualBillNoValidation
   If ObjRegistry.AllowManualBillNoValidation = True Then
      If TxtBillNo.Text = "" Then Cancel = True: Exit Sub
      With CN.Execute("SELECT dbo.FunLastManualBillNo('%" & TxtBillNo.Text & "%')")
         If .Fields(0).Value <> 0 Then
            TxtBillNo.Text = .Fields(0).Value
         End If
      End With
      Cancel = False
   Else
      Cancel = False
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCustomerID_Change()
   If TxtCustomerID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   If TxtCustomerName.Text <> "" Then
      TxtCustomerName.Text = ""
   End If
End Sub

Private Sub TxtCustomerID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCustomer(ssValidate, True)
'   If vTemp = True Then
'      vTemp = Not FunSelectCustomer(ssButton, False)
'   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.CmbFilter = "Customers"
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = "Select c.AccountNo, c.AccountName as AccountName, RefID, RefComm, Address, City, p.phone1, p.phone2, p.mobile, p.mobile2, p.Description, isnull(p.isWholeSale,1) as isWholeSale, LicenceNO, TransportName, Remarks" & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " where (c.AccountNo = " & Val(TxtCustomerID.Text) & " or P.Phone1 = '" & (TxtCustomerID.Text) & "' or P.Phone2 = '" & (TxtCustomerID.Text) & "' or P.Mobile = '" & (TxtCustomerID.Text) & "' or P.Mobile2 = '" & (TxtCustomerID.Text) & "') and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and isDetailed = 1 and isLocked = 0"
    
    vStrSQL = vStrSQL + " union all Select EmpID, EmpName as AccountName, Null RefID, Null RefComm, Address, City, '' phone1,'' phone2, '' mobile, '' mobile2, '', 1 as isWholeSale, '' as LicenceNO, '' as TransportName, '' as Remarks" & vbCrLf _
         + " from Employees" & vbCrLf _
         + " where EmpID = '" & (TxtCustomerID.Text) & "' and isLockEmployee = 0"
    
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerID.Text = !AccountNo
          TxtCustomerName.Text = !AccountName
          TxtRefID.Text = IIf(IsNull(!RefID), "", !RefID)
          TxtRefComm.Text = IIf(IsNull(!RefComm), "", !RefComm)
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtLicenceNO.Text = IIf(IsNull(!LicenceNO), "", !LicenceNO)
          TxtContactNo.Text = IIf(IsNull(!Phone1), "", !Phone1 & " ") & IIf(IsNull(!Phone2), "", !Phone2 & " ") & IIf(IsNull(!Mobile), "", !Mobile & " ") & IIf(IsNull(!Mobile2), "", !Mobile2)
          LblCustomerDesc.Caption = IIf(IsNull(!Description), "", !Description)
          TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
          TxtVehicleNo.Text = IIf(IsNull(!TransportName), "", !TransportName)
'         VStrSQL = "SELECT isnull(dbo.FunCurrentDebit('" & TxtCustomerID.Text & "','" & DtpBillDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)"
          If Val(TxtCustomerID.Text) <> 621 Then
            TxtPreviousReceivable.Text = CN.Execute("SELECT isnull(dbo.FunCurrentDebit(" & Val(TxtCustomerID.Text) & ",'" & DtpBillDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value

            vStrSQL = " Select isnull(Sum(round(B.TTLValue,0) - isnull(BillDisc,0) + isnull(OtherCharges,0) + Isnull(TotalExpense,0) + isnull(servicecharges,0) + isnull(STax,0)),0) as Amount " & vbCrLf _
                  + " FROM SaleHeader h INNER JOIN (Select SID, Sum(Amount) TTLValue FROM SaleBody Group By SID)b " & vbCrLf _
                  + " ON H.SID = B.SID " & vbCrLf _
                  + " where CustomerID = " & Val(TxtCustomerID.Text) & " and h.BillDate = '" & DtpBillDate.DateValue & "' and h.BillID >= " & Val(TxtBillID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
            TxtPreviousReceivable.Text = TxtPreviousReceivable.Text - CN.Execute(vStrSQL).Fields(0).Value
            lblPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
            TxtPreviousReceivable.Text = Abs(TxtPreviousReceivable.Text)
          End If
          vZoneID = CN.Execute("SELECT isnull(dbo.FunGetZoneID(" & Val(TxtCustomerID.Text) & "),0)").Fields(0).Value
          isWholeSale = !isWholeSale
          FunSelectCustomer = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectCustomer = False
          .Close
          TxtCustomerID.Text = ""
          TxtCustomerName.Text = ""
          TxtRefID.Text = ""
          TxtRefComm.Text = ""
          TxtAddress.Text = ""
          TxtCity.Text = ""
          LblCustomerDesc.Caption = ""
          TxtPreviousReceivable.Text = ""
          lblPayable.Caption = "Previous Payable"
          LblTtlPayable.Caption = "Total Payable"
          vZoneID = 0
          isWholeSale = True
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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

Private Sub BtnMember_Click()
   On Error GoTo ErrorHandler
   If FunSelectMember(ssButton, False) = True Then
      If TxtCode.Enabled Then TxtCode.SetFocus
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


Private Sub TxtExtraSchemePer_Change()
   If ActiveControl.Name <> TxtExtraSchemePer.Name Then Exit Sub
   Call CalculateValue
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

Private Sub SubApplyMember()
   On Error GoTo ErrorHandler
   Dim vAmount, vDiscVal, vTotDisc As Double
   Grid.MoveFirst
   ssql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With CN.Execute(ssql)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = " & Val(Grid.Columns("ProductID").Text)
         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
'            RsBody.Filter = "ProductID='" & !Productid & "'"
            vDiscVal = Val(Grid.Columns("DiscVal").Value)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
            Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("Price").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = Val(Grid.Columns("DiscPC").Value) * (Val(Grid.Columns("Pack").Value) * Val(Grid.Columns("QtyPack").Value) + Val(Grid.Columns("QtyLoose").Value))
'            Grid.Columns("DiscVal").Value = Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qtyloose").Value)
            vAmount = Val(Grid.Columns("Amount").Value)
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * (Val(Grid.Columns("Pack").Value) * Val(Grid.Columns("QtyPack").Value) + Val(Grid.Columns("QtyLoose").Value))) - Val(Grid.Columns("DiscVal").Value)
'            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("QtyLoose").Value)) - Val(Grid.Columns("DiscVal").Value)
            
            TxtNetAmount.Text = Val(TxtNetAmount.Text) - vAmount + Val(Grid.Columns("Amount").Text)
            vTotDisc = Val(vTotDisc) - Val(vDiscVal) + Val(Grid.Columns("DiscVal").Text)
            vTotalAmount = vTotalAmount - vAmount + Val(Grid.Columns("Amount").Text)
            
            
'            RsBody!DiscPC = Val(Grid.Columns("DiscPC").Value)
'            RsBody!DiscPer = Val(Grid.Columns("DiscPer").Value)
'            RsBody!DiscVal = Val(Grid.Columns("DiscVal").Value)
'            RsBody!Amount = Val(Grid.Columns("Amount").Value)
     
         End If
         Grid.MoveNext
      Wend
      .Close
   End With
   Grid.MoveLast
    TxtBillDisc.Text = Val(TxtBillDisc.Text) + vTotDisc
    TxtBillDiscPer.Text = Val(TxtBillDisc.Text) / IIf(Val(TxtTotalAmount.Text) = 0, 1, Val(TxtNetAmount.Text)) * 100
   SubCalculateFooter
   
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubDestroyMember()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   ssql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With CN.Execute(ssql)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = " & Val(Grid.Columns("ProductID").Text)
         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
            
            RsBody.Filter = "ProductID = " & Val(!Productid)
            Grid.Columns("DiscPer").Value = 0 'IIf(IsNull(!DiscPer), 0, !DiscPer)
            Grid.Columns("DiscPC").Value = 0 'Round((Val(RsBody!Price) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = 0 'Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) - Val(Grid.Columns("DiscVal").Value)
            
'            TxtNetAmount.text = Val(TxtNetAmount.text) - RsBody!Amount + Val(Grid.Columns("Amount").Text)
'            vTotDisc = vTotDisc - RsBody!DiscVal + Val(Grid.Columns("DiscVal").Text)
'            vTotalAmount = vTotalAmount - RsBody!Amount + Val(Grid.Columns("Amount").Text)
            
            RsBody!DiscPC = Val(Grid.Columns("DiscPC").Value)
            RsBody!DiscPer = Val(Grid.Columns("DiscPer").Value)
            RsBody!DiscVal = Val(Grid.Columns("DiscVal").Value)
            RsBody!Amount = Val(Grid.Columns("Amount").Value)
         End If
         Grid.MoveNext
      Wend
      .Close
   End With
   SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchSale.ParaInBillDate = DtpBillDate.DateValue
   SchSale.Show vbModal
   If SchSale.ParaOutBillID <> -1 Then
      TxtSID.Text = SchSale.ParaOutSID
      TxtBillID.Text = SchSale.ParaOutBillID
      'Dim a
      'a = Split(SchSale.ParaOutBillDate, "/")
      DtpBillDate.DateValue = SchSale.ParaOutBillDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      vStrSQL = "Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")"
      CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetSale
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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

Private Sub BtnOrganization_Click()
   If FunSelectOrganization(ssButton, False) = True Then
      TxtCustomerID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   If ObjRegistry.AllowUrduProduct = True Then
      If MsgBox("Do you want to print this invoice in Urdu", vbQuestion + vbYesNo, "Alert") = vbYes Then
         isUrdu = True
      Else
         isUrdu = False
      End If
   End If
   
   With CN.Execute("Select distinct o.OrganizationID, ReportName, SaleReportName from SaleHeader h left  join Organizations o on h.OrganizationID = o.OrganizationID where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'")
      
      vStrSQL = "Select h.BillID, h.BillDate, h.StoreID, UserName, ExpiryInvoice, BillTIme as EntryTime, EntryDate, h.PromiseDate, h.OrganizationID, OrganizationName, Customerid, cast(H.CustomerID as varchar(10))  + ' - ' + isnull(Pr.PartyName,AccountName)  as Customer_Name_ID," & vbCrLf _
                + " pr.address, LicenceNo, SectorName, ZoneName, H.StoreID, StoreName, BiltyNo, VehicleNo, h.Description, h.Remarks," & vbCrLf _
                + " Case when CustomerID = 621 then isnull(CustomerName,AccountName) Else AccountName End as Customer, Isnull(H.BillDiscPer, 0) BillDiscPer, Isnull(H.BillDisc,0) BillDisc, isnull(OtherCharges,0) as OtherCharges,  isnull(h.ServiceCharges,0) as ServiceCharges," & vbCrLf _
                + " TotalAmount,  isnull(TotalExpense,0) as TotalExpense,  CompanyName, " & IIf(ObjRegistry.AllowUrduProduct = False, "GroupName", "GroupName1") & "  as GroupName, SubGroupName, BrandName, SeasonName, b.ProductID, b.Code as Code, " & IIf(isUrdu = True, "p.ProductName1", "p.ProductName") & " as ProductName, ProductName1, dbo.FunSaleBodySerial(b.SID,b.BillDate, b.ProductId) Serial," & vbCrLf _
                + " dbo.FunSaleBodyOffer(b.BillID,b.BillDate,b.ProductId) ProductOffer, isnull(QtyPack,0)QtyPack, isnull(b.Multiplier,0)Multiplier, isnull(b.GrossQty,0)GrossQty, isnull(b.GrossUnit,0)GrossUnit, Qty," & vbCrLf _
                + " P.RetailPrice, P.PurPrice, Bonus, b.DiscPc, b.DiscPer, DiscVal, Cash, Credit, BankCard, isPrinted, Stax, Offer, Cast(b.Tradeoffer1 as varchar(5)) + ' + ' + cast(b.tradeoffer2 as varchar(5)) TradeOffer_12, tradevalue, Extraschemevalue, b.ExtraSchemePer," & vbCrLf _
                + " b.SaleTaxPer, SaleTaxval, b.IsRetailSaleTax, AdvTaxVal, AdvTaxPer, ExtraTaxVal, ExtraTaxPer, Pr.CNIC, h.CNIC, h.MobileNo,  b.SC, h.Empid, empname, price, Amount, previousAmount, CashReceived, isnull(BankAmount,0) BankAmount, isnull(BatchNo,'') as BatchNo, BillNo," & vbCrLf _
                + " Cast(b.Tradeoffer1 as varchar(5)) + ' + ' + cast(b.tradeoffer2 as varchar(5)) TradeOffer_12, tradevalue, Extraschemevalue, b.ExtraSchemePer," & vbCrLf _
                + " Abbreviation + '/' + cast(cast(b.Multiplier as int)as varchar(10)) as packing, isnull(P.ListPrice,0) as ListPrice," & vbCrLf _
                + " isnull( pr.Phone1  + ', ','') + isnull( pr.Phone2 + ', ','')  + isnull( pr.mobile + ', ','') +  isnull( pr.mobile2 + ', ','') as Moblie, packingname, pr.city, " & vbCrLf _
                + " isnull(pr.Address,'') + isnull(' - ' + pr.City,'') + isnull(',' + pr.Phone1,'') + isnull(',' + pr.Phone2, '') + isnull(',' + pr.Mobile, '') as AddressFull, ServerEntry, isLastPrice, " & vbCrLf _
                + " AmountType = " & vbCrLf _
                + " CASE  " & vbCrLf _
                + " WHEN bankcard = 1 THEN ' Through Bank Card' " & vbCrLf _
                + " WHEN Cash = 1 THEN ' Through Cash' " & vbCrLf _
                + " WHEN Credit = 1 THEN ' Through Credit' " & vbCrLf _
                + " End "
                vStrSQL = vStrSQL + " from SaleBody b inner join products p on b.productid = p.productid" & vbCrLf _
                + " inner join SaleHeader h on H.SID = B.SID" & vbCrLf _
                + " inner join users ur on ur.UserNo = h.UserNo Left Outer jOin companies cmp on cmp.companyid = p.companyid " & vbCrLf _
                + " Left Outer jOin Groups g on g.Groupid = p.Groupid Left Outer jOin SubGroups sg on sg.subGroupid = p.subGroupid" & vbCrLf _
                + " Left Outer jOin Brands bd on bd.brandid = p.brandid" & vbCrLf _
                + " Left Outer jOin Seasons se on se.Seasonid = p.Seasonid" & vbCrLf _
                + " LEFT OUTER JOIN packings pak on pak.packingid = b.packingid" & vbCrLf _
                + " left outer join Organizations o on o.OrganizationID = h.OrganizationID" & vbCrLf _
                + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
                + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
                + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
                + " left outer join Sectors Sec on Sec.SectorID = Pr.SectorID" & vbCrLf _
                + " left outer join Zones Z on Z.ZoneID = Sec.ZoneID" & vbCrLf _
                + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
                + " where h.SID = " & Val(TxtSID.Text) & " and h.BillDate='" & DtpBillDate.DateValue & "'" & IIf(ObjRegistry.AllowOrderByCodeinInvoices, " Order By Code", " Order By SerialNo")
         
                

   If cmbPrintType.Text = "Thermal" Then
      If vIsNewRecord = False And ObjRegistry.ShowDuplicatePrint = True Then
         vStrSQL = "Exec ProdPrintSalePos " & Val(TxtSID.Text) & ",'Duplicate'"
      Else
         vStrSQL = "Exec ProdPrintSalePos " & Val(TxtSID.Text) & ",''"
      End If
   End If

    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   
   
   If cmbPrintType.Text = "Half Page" Then
      'Set RptReportViewer.Report = New CrpSaleInvoiceHalf1
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleInvoiceHalf1.rpt")
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.x
      RptReportViewer.Report.RightMargin = 225
   ElseIf cmbPrintType.Text = "Thermal" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleInvoiceAurora.rpt")
      RptReportViewer.Report.TopMargin = 0
      RptReportViewer.Report.LeftMargin = 0
      RptReportViewer.Report.RightMargin = 0
   Else
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\" & IIf(IsNull(!SaleReportName), "CrptSaleInvoice", !SaleReportName) & ".rpt")
   End If
   

   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ReportTitle = "Sale Invoice"
   
   If cmbPrintType.Text = "Thermal" Then
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
      RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.AddSpace = True, Left(".......................................", Val(ObjRegistry.BlankFooter)), "")
      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(ObjRegistry.CashReceived)
      RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(ObjRegistry.Statement)
      RptReportViewer.Report.ParameterFields(8).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(9).AddCurrentValue (IIf(ObjRegistry.PreviousBalanceVisible = True, 0, 0))
   Else
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
      RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CStr(ObjRegistry.Statement)
   End If
   If ObjRegistry.PrintHeadersSaleInvoice = True Then
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   Else
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ""
   End If
   
'   Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n """ & ObjRegistry.DeviceName & """"
   
   'RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   
   vPrinter = Split(CmbPrinters.Text, ",")
'   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   
   If ObjRegistry.isShowSeason = True Then
         RptReportViewer.Report.PaperOrientation = crLandscape
    ElseIf ObjRegistry.ShowTradeOffer = True Then
        RptReportViewer.Report.PaperSize = crPaperA4
        RptReportViewer.Report.PaperOrientation = crLandscape
    Else
'         RptReportViewer.Report.PaperSize = crPaperLegal
         RptReportViewer.Report.PaperOrientation = crPortrait
    End If
   If ObjRegistry.IsPortrait = False Then RptReportViewer.Report.PaperOrientation = crLandscape
   If cmbPrintType.Text = "Half Page" Then RptReportViewer.Report.PaperOrientation = crLandscape
   
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   
   If ObjRegistry.PreviewSaleInoice Or ObjRegistry.isShowSeason = True Or ChkIsPreview.Value = 1 Then
      If ChkIsPrint.Value = 1 Then
         RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
      End If
       RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
   End If
 
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmSaleInvoiceDIS, eRePrint, TxtBillID.Text, DtpBillDate.DateValue, "RePrinted Amount: " & Val(TxtNetAmount.Text))
'   cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
    .Close
   End With
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
'  If cn.State = adStateClosed Then cn.Open
  On Error GoTo ErrorHandler
   ''''''''''''' User Discount ''''''''''''''
   If Val(ObjUserSecurity.AllowMaximmDiscPer) <> 0 Then
      If Val(TxtBillDiscPer.Text) > Val(ObjUserSecurity.AllowMaximmDiscPer) Then
         MsgBox "Discount greater than Fixed Discount", vbCritical, "Error"
         Exit Sub
      End If
   End If
  ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSaleInvoice", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
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
  
   If Trim(TxtCustomerID.Text) = "" Then
      MsgBox "Enter Customer ID.", vbExclamation, Me.Caption
      TxtCustomerID.SetFocus
      Exit Sub
   Else
      With CN.Execute("Select * from Parties where CreditLimit <> 0 and CreditLimit is not null and PartyID = " & Val(TxtCustomerID.Text))
         If .RecordCount > 0 Then
            If !CreditLimit < (Val(TxtTotalReceivable.Text)) Then
               MsgBox "Credit Limit (" & !CreditLimit & ") is Exceed Balance (" & (Val(TxtTotalReceivable.Text)) & ") for this Customer.", vbExclamation, "Alert"
               Exit Sub
            End If
         End If
      End With
      With CN.Execute("Select * from Employees where CreditLimit <> 0 and CreditLimit is not null and EmpID = '" & TxtCustomerID.Text & "'")
         If .RecordCount > 0 Then
            If !CreditLimit < (Val(TxtTotalReceivable.Text)) Then
               MsgBox "Credit Limit (" & !CreditLimit & ") is Exceed Balance (" & (Val(TxtTotalReceivable.Text)) & ") for this Customer.", vbExclamation, "Alert"
               Exit Sub
            End If
         End If
      End With
   End If
    
   
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where ToUserNo = " & vUser & " and Entrydate <='" & Date & "'"
    With CN.Execute(vStrSQL)
        If .Fields(0).Value >= DtpBillDate.DateValue Then
            MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
   '''''''''''''''''''''''Check Entry Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.isEntryDate = True Then
       If ObjRegistry.FromDate > Date Or ObjRegistry.ToDate < Date Then
         MsgBox "Data can not be saved because date is not set according to the software's entry date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
    '''''''''''''''''''''''Check Current Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.CurrentDateDataEntry = True And ObjUserSecurity.IsAdministrator = False Then
       If DtpBillDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
    
     '''''''''''''''''''''''Check Import / Export'''''''''''''''''''''''''''''''''
    If ObjRegistry.ShowMultiBranches = True Then
      vStrSQL = "select * from SaleHeader where Tag is not null And SID=" & Val(TxtSID.Text) & " and Billdate='" & DtpBillDate.DateValue & "'"
      With CN.Execute(vStrSQL)
          If Not .EOF Then
              MsgBox "Import/Export Record Cannot be Updated", vbInformation, Me.Caption
              Exit Sub
          End If
      End With
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If ObjRegistry.DefaultCustomer = True Then
      Call DefaultCustomer
      vStrSQL = "Select AccountNo, VoucherDate, Debit, Credit, Datediff(day,VoucherDate,GETDATE()) Aging from AccountsLedger where Credit > 0  and Datediff(day,VoucherDate,GETDATE())  >=30 order by VoucherDate Desc "
      With CN.Execute(vStrSQL)
          If Not .EOF Then
              MsgBox "Default Customer becuase last payment was made " & !Aging & " days ago on " & !VoucherDate, vbInformation, Me.Caption
              Exit Sub
          End If
      End With
   End If
   
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      If TxtStoreID.Visible And TxtStoreID.Enabled Then TxtStoreID.SetFocus
      Exit Sub
   End If
   
   If ObjRegistry.AllowManualBillNoValidation = True And Trim(TxtBillNo.Text) = "" Then
      MsgBox "Enter Bill No.", vbExclamation, Me.Caption
      TxtBillNo.SetFocus
      Exit Sub
   End If
            
   If ObjRegistry.RemarksCompulsory = True And Trim(TxtRemarks.Text) = "" Then
      MsgBox "Enter Remarks.", vbExclamation, Me.Caption
      TxtRemarks.SetFocus
      Exit Sub
   End If

'   FrmPrint.TxtNetAmount.Text = TxtNetAmount.Text
'   If objRegistry.CashReceived = False Then
'      FrmPrint.TxtCashReceivedCash.Text = TxtNetAmount.Text
'   End If
'   FrmPrint.ParaInPrint = True
'   FrmPrint.ParaInChoice = "Cash"
'   FrmPrint.ParaInDate = DtpBillDate.DateValue
'   FrmPrint.Show vbModal, Me
'   If FrmPrint.ParaOutSelection = False Then Exit Sub
'   If DtpBillDate.Enabled And DtpBillDate.Date <> IIf(Format(Now, "hh") > 2, Date, DateAdd("d", -1, Date)) And DateFlag = True Then
'      If MsgBox("Are you sure to Change Bill Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
'         DtpBillDate.DateValue = IIf(Format(Now, "hh") > 2, Date, DateAdd("d", -1, Date))
'         TxtBillID.Text = FunGetMaxID()
'      End If
'      DateFlag = False
'   End If

  'Body Validation
  ' validation has been performed when a row is added to the grid
  'RsBody.Filter = 0
  If Grid.rows < 2 Then
      MsgBox "Please enter at least one product to Sale", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
  End If

  
  '''''' Check Multiple Store ''''''''''''''
  If vUseMultipleStore Then
   With Grid
   .Redraw = False
   .MoveFirst
   For vCounter = 1 To .rows
      PopulateDataToGridMultipleStore
        If Val(Grid.Columns("QtyLoose").Text) + Val(Grid.Columns("Bonus").Text) + (Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text)) <> Val(TxtMSTotalItems.Text) Then
         MsgBox .Columns("ProductName").Text & " has Quantity Must be equal to Store Quanity", vbCritical, "Alert"
         .Enabled = True
         .Redraw = True
         Exit Sub
      End If
      .MoveNext
   Next
   .Redraw = True
   End With
  Else
   RsBodyStore.Filter = 0
   If RsBodyStore.RecordCount > 0 Then
      MsgBox "This Bill ID include Muliple Store So Goto Software Default Setting and Click on UseMultipleStore Then Save ", vbCritical, "Alert"
      Exit Sub
   End If
  End If
  '''''''''''''''''''''''''
  
  'Saving record
  
   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'Sale Invoice DIS' and LocalComputerName = '" & LocalComputerName & "'"
   If CN.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview, IsPrint ) Values ('" & LocalComputerName & "', 'Sale Invoice DIS','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & "," & ChkIsPrint.Value & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & ", IsPrint = " & ChkIsPrint.Value & " Where FormType = 'Sale Invoice DIS' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   CN.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   
   CN.BeginTrans
   
   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmSaleInvoiceDIS, eEdit, TxtBillID.Text, DtpBillDate.DateValue, "Amount: " & Val(TxtNetAmount.Text))
   
   If ObjRegistry.ShowDispatchDate = True Then CN.Execute ("Update sysindexs Set value = '" & Val(TxtTerms.Text) & "' Where Registrykey = 'SalePromiseTerms'")
   If vIsNewRecord Then
      If CN.Execute("Select * from SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'").RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtBillID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
      
      
   ''''''''''''''''''''''''''Delete Product OFfer'''''''''''''''''''''
'   cn.Execute "Delete from SaleBodyOffer where BillID = " & Val(TxtBillID.Text) & " And BillDate ='" & DtpBillDate.DateValue & "'"
'
'   ''''''''''''''''''''''''''Delete Serials'''''''''''''''''''''
'   cn.Execute "Delete from SaleBodySerial where BillID = " & Val(TxtBillID.Text) & " And BillDate ='" & DtpBillDate.DateValue & "'"
'
'
'   ''''''''''''''''''''''''''Delete Sale Body'''''''''''''''''''''
'   cn.Execute "Delete from SaleBody where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
'
'
'   '''''''''''''''''''''''''''''''''''''''Delete Expense'''''''''''''''''''''''''''''''''''''''
'   cn.Execute "Delete from SaleExpense where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   
   '''''''''''''''''''''''''''''''''''''''Delete Header'''''''''''''''''''''''''''''''''''''''
'   cn.Execute "Delete from SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   
'   cn.Execute ("Update SaleOrderHeader set IsSale = 0 Where OrderID = " & Val(TxtOrderID.Text) & "And Orderdate ='" & DtpOrderDate.DateValue & "'")
   
   
   vBillID = TxtBillID.Text
   vBillDate = DtpBillDate.DateValue
   Dim Rs As New ADODB.Recordset
   
   If vIsNewRecord = False Then
'      Call ActivityLogSale("Sale Invoice", eEdit, TxtBillID.Text, DtpBillDate.DateValue)
   End If
   
''   Call UserActivities

'' Sale Header
   vStrPara = ""
   vStrPara = Abs(ObjRegistry.AllowContinuousBillNo) & "," 'AllowContinuousBillNo
   vStrPara = vStrPara & Abs(ObjRegistry.AllowMonthlyBillNo) & "," 'AllowMonthlyBillNo
   vStrPara = vStrPara & Abs(ObjRegistry.AllowDailyBillNo) & "," 'AllowDailyBillNo
   vStrPara = vStrPara & Val(TxtSID.Text) & "," 'SID
   vStrPara = vStrPara & TxtBillID.Text & ","               'BillID
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"    'BillDate
vStrPara = vStrPara & TxtCustomerID.Text & ","              'CustomerID
vStrPara = vStrPara & TxtTotalAmount.Text & ","             'TotalAmount
vStrPara = vStrPara & Val(TxtBillDisc.Text) & ","           'BillDisc
vStrPara = vStrPara & Val(TxtReceivedAmount.Text) & ","     'CashReceived
vStrPara = vStrPara & vUser & ","                           'UserNo
vStrPara = vStrPara & TxtStoreID.Text & ","                 'StoreID
vStrPara = vStrPara & 0 & ","                               'BankCard
With CN.Execute("select dbo.DefaultValue('Cash Counter')")
   If TxtCustomerID.Text = .Fields(0).Value Then
            vCash = 1
            vCredit = 0
         Else
            vCash = 0
            vCredit = 1
         End If
End With
vStrPara = vStrPara & vCredit & ","                         'Credit
vStrPara = vStrPara & vCash & ","                           'Cash
vStrPara = vStrPara & "'" & Null & "',"                     'BankMachineID
vStrPara = vStrPara & "'" & Null & "',"                     'InvoiceNo
vStrPara = vStrPara & "'" & IIf(Trim(TxtCashCustomer.Text) = "", Null, TxtCashCustomer.Text) & "',"   'CustomerName
vStrPara = vStrPara & Val(TxtBillDiscPer.Text) & ","        'BillDiscPer
vStrPara = vStrPara & "'" & Null & "',"                     'Commision
vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtCommission.Text)) & ","         'EmpComm
vStrPara = vStrPara & "'" & IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text)) & "',"  'EmpID
vStrPara = vStrPara & 0 & ","                               'isReplace
vStrPara = vStrPara & 0 & ","                               'isPosted
vStrPara = vStrPara & IIf(Trim(TxtMemberID.Text) = "", "''", TxtMemberID.Text) & ","                  'MemberID
vStrPara = vStrPara & "'" & vBillTime & "',"                       'BillTime
vStrPara = vStrPara & "'" & vIsNewRecord & "',"             'Tag
vStrPara = vStrPara & "'" & IIf(Trim(TxtManualBillNo.Text) = "", Null, TxtManualBillNo.Text) & "',"   'ManualBillNo
vStrPara = vStrPara & "'" & IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text) & "',"             'Remarks
vStrPara = vStrPara & IIf(Trim(TxtOrganizationID.Text) = "", "''", TxtOrganizationID.Text) & ","      'OrganizationID
vStrPara = vStrPara & "'" & IIf(Trim(TxtBillNo.Text) = "", Null, TxtBillNo.Text) & "',"               'BillNO
vStrPara = vStrPara & "'" & IIf(Trim(TxtBiltyNo.Text) = "", Null, TxtBiltyNo.Text) & "',"             'BILTYNO
vStrPara = vStrPara & "'" & IIf(Trim(TxtDescription.Text) = "", Null, TxtDescription.Text) & "',"     'DESCRIPTION
vStrPara = vStrPara & 0 & ","                               'PAIDAMOUNT
vStrPara = vStrPara & "'" & Null & "',"                     'EntryDate
vStrPara = vStrPara & IIf(lblPayable.Caption = "Previous Receivable", Val(TxtPreviousReceivable.Text), Val(TxtPreviousReceivable.Text) * -1) & "," 'PreviousAmount
vStrPara = vStrPara & Val(TxtOtherCharges.Text) & ","       'OtherCharges
vStrPara = vStrPara & "'" & Null & "',"                     'SaleManID
vStrPara = vStrPara & Val(TxtTotalExpense.Text) & ","       'TotalExpense
vStrPara = vStrPara & Val(TxtOrderID.Text) & ","            'OrderID
vStrPara = vStrPara & "'" & DtpOrderDate.DateValue & "',"   'OrderDate
vStrPara = vStrPara & Val(TxtFreight.Text) & ","            'Freight
vStrPara = vStrPara & IIf(OptCustomer.Value = True, 1, 0) & ","                                       'IsCustomerFreight
vStrPara = vStrPara & "'" & IIf(Trim(TxtVehicleNo.Text) = "", Null, TxtVehicleNo.Text) & "',"         'VehicleNo
vStrPara = vStrPara & IIf(TxtServiceCharges.Text = "", "''", Val(TxtServiceCharges.Text)) & "," 'ServiceCharges
vStrPara = vStrPara & IIf(TxtServiceChargesPer.Text = "", "''", Val(TxtServiceChargesPer.Text)) & "," 'ServiceChargesPer
vStrPara = vStrPara & 0 & ","                               'STax
vStrPara = vStrPara & 0 & ","                               'STaxPer
vStrPara = vStrPara & "'" & Null & "',"                     'TableID
vStrPara = vStrPara & "'" & Now & "',"                     'ServerEntry
vStrPara = vStrPara & "'" & Null & "',"                     'InvType
vStrPara = vStrPara & "'" & Null & "',"                     'DeliveryDate
vStrPara = vStrPara & "'" & Null & "',"                     'DeliveryTime
vStrPara = vStrPara & "'" & Null & "',"                     'isPrinted
vStrPara = vStrPara & "'" & Null & "',"                     'RemarksUrdu
'vStrPara = vStrPara & "'" & Null & "',"                     'StampID
vStrPara = vStrPara & 0 & ","                                'isTransfer
vStrPara = vStrPara & IIf(DtpPromiseDate.DateValue = Empty, "Null", "'" & DtpPromiseDate.DateValue & "'") & "," 'PromiseDate
vStrPara = vStrPara & IIf(DtpExpiryInvoice.DateValue = Empty, "Null", "'" & DtpExpiryInvoice.DateValue & "'") & "," 'ExpiryInvoice
vStrPara = vStrPara & "'" & IIf(Trim(TxtSyllabusID.Text) = "", Null, Val(TxtSyllabusID.Text)) & "',"  'SyllabusID
vStrPara = vStrPara & "'" & IIf(Trim(vSessionID) = 0, Null, Val(vSessionID)) & "',"  'vSessionID
vStrPara = vStrPara & IIf(TxtAdvTaxVal.Text = "", "''", Val(TxtAdvTaxVal.Text)) & "," 'AdvTaxVal
vStrPara = vStrPara & IIf(TxtAdvTaxPer.Text = "", "''", Val(TxtAdvTaxPer.Text)) & "," 'AdvTaxPer
vStrPara = vStrPara & IIf(TxtExtraTaxVal.Text = "", "''", Val(TxtExtraTaxVal.Text)) & "," 'ExtraTaxVal
vStrPara = vStrPara & IIf(TxtExtraTaxPer.Text = "", "''", Val(TxtExtraTaxPer.Text)) & "," 'ExtraTaxPer
vStrPara = vStrPara & "'" & IIf(Trim(TxtCNIC.Text) = "", Null, TxtCNIC.Text) & "',"  'CNIC
vStrPara = vStrPara & "'" & IIf(Trim(TxtCellNo.Text) = "", Null, TxtCellNo.Text) & "',"  'CellNo
vStrPara = vStrPara & Val(TxtSumDiscAmount.Text) & "," 'Sum Disc Amount
vStrPara = vStrPara & IIf(DtpDispatchDate.DateValue = Empty, "Null", "'" & DtpDispatchDate.DateValue & "'") & "," 'DispatchDate
vStrPara = vStrPara & IIf(Val(TxtTerms.Text) > 0, TxtTerms.Text, "''") & ","  'Terms
vStrPara = vStrPara & "'" & IIf(Trim(TxtRefID.Text) = "", Null, TxtRefID.Text) & "',"  'RefID
vStrPara = vStrPara & "'" & IIf(Trim(TxtRefComm.Text) = "", Null, TxtRefComm.Text) & "',"  'RefComm
vStrPara = vStrPara & "'" & Null & "'" 'Bank Amount in Credit Option
vStrPara = Replace(vStrPara, "''", "Null")

vStrPara = "DECLARE @returnvalue INT EXEC @returnvalue = saleheaderinsert " & vStrPara & " Select @returnvalue"
   vMasterID = CN.Execute(vStrPara).Fields(0).Value
   TxtSID.Text = vMasterID
'   MsgBox vMasterID
   

''' insert Sale Body
vStrDetail = ""
With Grid
 .Redraw = False
 .MoveFirst
   For vCounter = 1 To .rows
      If Trim(.Columns("Productid").Text) <> "" Then
      
      '''''' ActivityLogBin Follwoin lines check the same product id which was enter seperate row or new new row
      If (InStr(1, vSamePid, .Columns("Productid").Text)) = 0 Then vGridRows = vGridRows + 1
      vSamePid = vSamePid & " , " & .Columns("Productid").Text
      '''''''''''''''''''''''''''''''''''''
      
 vStrPara = ""
vStrPara = vStrPara & "'" & True & "',"
vStrPara = vStrPara & vMasterID & ","
vStrPara = vStrPara & CN.Execute("Select billID from Saleheader where SID = " & vMasterID).Fields(0).Value & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
'vStrPara = vStrPara & .Columns("SerialNo").Text & ","
'vStrPara = vStrPara & .Columns("BillID").Text & ","
'vStrPara = vStrPara & .Columns("BillDate").Text & ","
vStrPara = vStrPara & "'" & .Columns("ProductID").Text & "',"
vStrPara = vStrPara & .Columns("QtyLoose").Text & ","
vStrPara = vStrPara & .Columns("Price").Text & ","
vStrPara = vStrPara & .Columns("DiscPC").Text & ","
vStrPara = vStrPara & .Columns("Amount").Text & ","
vStrPara = vStrPara & "'" & .Columns("Code").Text & "',"
vStrPara = vStrPara & Val(.Columns("DiscPer").Text) & ","
vStrPara = vStrPara & Val(.Columns("DiscVal").Text) & ","

vStrPara = vStrPara & Val(.Columns("isDiscB4TradeOffer").Text) & "," ' isDiscB4TradeOffer
vStrPara = vStrPara & Val(.Columns("isDiscB4ExtraScheme").Text) & ","   'isDiscB4ExtraScheme
vStrPara = vStrPara & Val(.Columns("isDiscB4SaleTax").Text) & "," 'isDiscB4SaleTax
vStrPara = vStrPara & Val(.Columns("TradeOffer1").Text) & ","  'TradeOffer1
vStrPara = vStrPara & Val(.Columns("TradeOffer2").Text) & ","   'TradeOffer2
vStrPara = vStrPara & Val(.Columns("ExtraSchemePer").Text) & ","   'ExtraSchemePer
vStrPara = vStrPara & Val(.Columns("TradeValue").Text) & ","   'TradeValue
vStrPara = vStrPara & Val(.Columns("ExtraSchemeValue").Text) & ","   'ExtraSchemeValue


vStrPara = vStrPara & Val(.Columns("Cost").Text) & ","   'cost
vStrPara = vStrPara & .Columns("isProduct").Text & ","   'isProduct
vStrPara = vStrPara & IIf(Val(.Columns("PackingID").Text) > 0, .Columns("PackingID").Text, "''") & "," 'PackingID
vStrPara = vStrPara & IIf(.Columns("QtyPack").Text = "", "Null", .Columns("QtyPack").Text) & ","   'QtyPack
vStrPara = vStrPara & IIf(.Columns("Pack").Text = "", "Null", .Columns("Pack").Text) & ","   'Pack
vStrPara = vStrPara & Val(.Columns("Bonus").Text) & ","  'Bonus
vStrPara = vStrPara & Val(.Columns("Offer").Text) & ","  'Offer
vStrPara = vStrPara & Val(.Columns("SaleTaxPer").Text) & ","   'SaleTaxPer
vStrPara = vStrPara & Val(.Columns("SaleTaxVal").Text) & ","   'SaleTaxVal
vStrPara = vStrPara & Val(.Columns("TokenVal").Text) & ","  'TokenVal
vStrPara = vStrPara & Val(.Columns("RetailPrice").Text) & ","    'RetailPrice
vStrPara = vStrPara & Val(.Columns("IsWSSaleTax").Text) & ","    'IsWSSaleTax
vStrPara = vStrPara & .Columns("IsRetailSaleTax").Text & ","   'IsRetailSaleTax
vStrPara = vStrPara & Val(.Columns("IsWSDiscb4ST").Text) & ","   'IsWSDiscb4ST
vStrPara = vStrPara & IIf(.Columns("SC").Text = "", "Null", .Columns("SC").Text) & ","   'SC
vStrPara = vStrPara & "''" & "," 'EmpComm
vStrPara = vStrPara & IIf(Trim(.Columns("BatchNo").Text) <> "", "'" & .Columns("BatchNo").Text & "'", "''") & ","  'Batcho
'vStrPara = vStrPara & "''" & "," 'StampID
vStrPara = vStrPara & TxtStoreID.Text & "," 'StoreID
vStrPara = vStrPara & "''" & ","  'EmpID
vStrPara = vStrPara & "null" & "," 'ColourID
vStrPara = vStrPara & "null" & ","  'SizeID
vStrPara = vStrPara & IIf(.Columns("GrossQty").Text = "", "Null", .Columns("GrossQty").Text) & "," 'Gross Qty
vStrPara = vStrPara & IIf(.Columns("GrossUnit").Text = "", "Null", .Columns("GrossUnit").Text) & "," 'Gross Unit
vStrPara = vStrPara & IIf(.Columns("Storeid").Text = "", "Null", .Columns("StoreID").Text) & ","                 'HeaderStoreID
vStrPara = vStrPara & Val(.Columns("DiscAmount").Value) & "," ' Disc Amount
vStrPara = vStrPara & Val(.Columns("isLastPrice").Text) & "," ' isLastPrice
vStrPara = vStrPara & IIf(Val(.Columns("ReSPrice").Value) = 0, "Null", .Columns("ReSPrice").Value) & ","   'Re SPrice
vStrPara = vStrPara & IIf(Val(.Columns("ReSAmount").Value) = 0, "Null", .Columns("ReSAmount").Value) & ""    'Re SAmount

vStrPara = Replace(vStrPara, "''", "Null")
CN.Execute ("Exec SaleBodyInsert " & vStrPara)
vStrDetail = vStrDetail & " (P" & .Columns("ProductID").Text & IIf(Val(.Columns("Pack").Value) = 0, "", " M" & .Columns("Pack").Value) & IIf(Val(.Columns("QtyPack").Value) = 0, "", " QP" & .Columns("QtyPack").Value) & IIf(Val(.Columns("QtyLoose").Value) = 0, "", " QL" & .Columns("QtyLoose").Value) & IIf(Val(.Columns("Bonus").Value) = 0, "", " QB" & .Columns("Bonus").Value) & " A" & .Columns("Amount").Text & ")"

      ''''''''''''''''''''''''''''
      
      End If
      .MoveNext
   Next vCounter
   .RemoveAll
   .Redraw = True
   End With
   
''' insert Sale Body Offer
With GridOffer
 .Redraw = False
 .MoveFirst
   For vCounter = 1 To .rows
      If Trim(.Columns("Productid").Text) <> "" Then
      
      ''''''''''''''''''''''''''''
vStrPara = ""
vStrPara = vStrPara & vMasterID & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
vStrPara = vStrPara & Val(.Columns("ProductID").Text) & ","
vStrPara = vStrPara & "'" & .Columns("ProductOfferID").Text & "',"
vStrPara = vStrPara & .Columns("Qty").Text & ""

vStrPara = Replace(vStrPara, "''", "Null")
CN.Execute ("Exec SaleBodyOfferInsert " & vStrPara)

      ''''''''''''''''''''''''''''
      
      End If
      .MoveNext
   Next vCounter
   .RemoveAll
   .Redraw = True
   End With
   
 '''''' Sale Body Serial
 
 With GridSerial
 .Redraw = False
 .MoveFirst
  For vCounter = 1 To .rows
      If Trim(.Columns("Productid").Text) <> "" Then
      
      ''''''''''''''''''''''''''''
 vStrPara = ""
vStrPara = vStrPara & vMasterID & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
vStrPara = vStrPara & Val(.Columns("ProductID").Text) & ","
vStrPara = vStrPara & "'" & .Columns("Serial").Text & "'"

vStrPara = Replace(vStrPara, "''", "Null")
CN.Execute ("Exec SaleBodySerialInsert " & vStrPara)

      ''''''''''''''''''''''''''''
      
      End If
      .MoveNext
   Next vCounter
   .RemoveAll
   .Redraw = True
   End With
   
 '''''' Sale Expense
 
 With GridExpense
 .Redraw = False
 .MoveFirst
  For vCounter = 1 To .rows
      If Trim(.Columns("ID").Text) <> "" Then
      
      ''''''''''''''''''''''''''''
 vStrPara = ""
vStrPara = vStrPara & vMasterID & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
vStrPara = vStrPara & "'" & .Columns("ID").Text & "',"
vStrPara = vStrPara & .Columns("Value").Text & ""

vStrPara = Replace(vStrPara, "''", "Null")
CN.Execute ("Exec SaleExpenseInsert " & vStrPara)

      ''''''''''''''''''''''''''''
      
      End If
      .MoveNext
   Next vCounter
   .RemoveAll
   .Redraw = True
   End With
   
   If vUseMultipleStore = True Then
   With RsBodyStore
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !SID = Val(TxtSID.Text)
         !BillID = Val(TxtBillID.Text)
         !BillDate = DtpBillDate.DateValue
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   End If
'   TxtBillID.Text , DtpBillDate.DateValue, TxtCustomerID.Text, TxtTotalAmount.Text, IIf(TxtBillDisc.Text = "", Null, Round(Val(TxtBillDisc.Text), 3)), vbcrlf_
'   TxtReceivedAmount.Text , vUser, TxtStoreID.Text, BankCard, Credit, Cash, BankMachineID, InvoiceNo, CustomerName, IIf(TxtBillDiscPer.Text = "", Null, Round(Val(TxtBillDiscPer.Text), 3)), vbcrlf_
'   Commision , IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text)), IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text), vbcrlf_
'   isReplace , isPosted, IIf(Trim(TxtMemberID.Text) = "", Null, TxtMemberID.Text), Now, Tag, IIf(Trim(TxtManualBillNo.Text) = "", "", TxtManualBillNo.Text), IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text), IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text), IIf(TxtBillNo.Text = "", Null, TxtBillNo.Text), IIf(TxtBiltyNo.Text = "", Null, TxtBiltyNo.Text), IIf(TxtDescription.Text = "", Null, TxtDescription.Text), PaidAmount, vbcrlf_
'   EntryDate , IIf(lblPayable.Caption = "Previous Receivable", Val(TxtPreviousReceivable.Text), Val(TxtPreviousReceivable.Text) * -1), IIf(Val(TxtOtherCharges.Text) = 0, Null, Val(TxtOtherCharges.Text)), SalemanID, IIf(Val(TxtTotalExpense.Text) = 0, Null, Val(TxtTotalExpense.Text)), IIf(Val(TxtOrderID.Text) = 0, Null, TxtOrderID.Text), DtpOrderDate.DateValue, IIf(Val(TxtFreight.Text) = 0, Null, Val(TxtFreight.Text)), IIf(OptCustomer.Value = True, 1, 0), !VehicleNo = IIf(TxtVehicleNo.Text = "", Null, TxtVehicleNo.Text), vbcrlf_
'   ServiceCharges , ServiceChargesPer, STax, STaxPer, vbcrlf_
'   TableId , ServerEntry, InvType, DeliveryDate, DeliveryTime, isPrinted, RemarksUrdu, StampID, isTransfer
   
    
   
'   ssql = "select * from SaleHeader where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
'   'Dim Rs As New ADODB.Recordset
'   With Rs
'      .Open ssql, cn, adOpenDynamic, adLockPessimistic
'      If .BOF Then
'         .AddNew
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         !OrderID = IIf(Val(TxtOrderID.Text) = 0, Null, TxtOrderID.Text)
'         !OrderDate = DtpOrderDate.DateValue
'         !BillTime = Now
'      End If
'      !isReplace = 0
'      !isPosted = 0
'      !isTransfer = 0
'      !StoreID = TxtStoreID.Text
'      !CustomerID = TxtCustomerID.Text
'      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
'      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
'      !EmpComm = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text))
'      !MemberID = IIf(Trim(TxtMemberID.Text) = "", Null, TxtMemberID.Text)
'      !BillNo = IIf(TxtBillNo.Text = "", Null, TxtBillNo.Text)
'      !BiltyNo = IIf(TxtBiltyNo.Text = "", Null, TxtBiltyNo.Text)
'      !VehicleNo = IIf(TxtVehicleNo.Text = "", Null, TxtVehicleNo.Text)
'      !TotalAmount = Round(Val(TxtTotalAmount.Text))
'      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Round(Val(TxtBillDiscPer.Text), 3))
'      !Description = IIf(TxtDescription.Text = "", Null, TxtDescription.Text)
'      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Round(Val(TxtBillDisc.Text), 3))
'      !Freight = IIf(Val(TxtFreight.Text) = 0, Null, Val(TxtFreight.Text))
'      !IsCustomerFreight = IIf(OptCustomer.Value = True, 1, 0)

'      If FrmPrint.OptBankCard.Value = True Then
'         !InvoiceNo = FrmPrint.TxtInvoiceNo.Text
'         !Commision = FrmPrint.TxtCommision.Text
'         !BankMachineID = FrmPrint.TxtBankMachineID.Text
'         !CashReceived = 0
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(FrmPrint.TxtBankCustomer.Text) = "", Null, FrmPrint.TxtBankCustomer.Text)
'      End If
'      If FrmPrint.OptCash.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !BankMachineID = Null
'         !CashReceived = Val(FrmPrint.TxtCashReceivedCash.Text)
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(FrmPrint.TxtCashCustomer.Text) = "", Null, FrmPrint.TxtCashCustomer.Text)
'      End If
'      If FrmPrint.OptCredit.Value = True Then
'         !Commision = Null
'         !InvoiceNo = Null
'         !BankMachineID = Null
'         !CashReceived = Val(FrmPrint.TxtCashReceivedCredit.Text)
'         !CustomerID = FrmPrint.TxtCustomerID.Text
'         !CustomerName = Null
'      End If
      
'      !BankCard = FrmPrint.OptBankCard.Value
'      !Cash = FrmPrint.OptCash.Value
'      !Credit = FrmPrint.OptCredit.Value
'      !BankCard = 0
'      With cn.Execute("select dbo.DefaultValue('Cash Counter')")
'         If TxtCustomerID.Text = .Fields(0).Value Then
'            Rs!Cash = 1
'            Rs!Credit = 0
'         Else
'            Rs!Cash = 0
'            Rs!Credit = 1
'         End If
'      End With
'      !CashReceived = Val(TxtReceivedAmount.Text)
'      !PreviousAmount = IIf(lblPayable.Caption = "Previous Receivable", Val(TxtPreviousReceivable.Text), Val(TxtPreviousReceivable.Text) * -1)
'      '!Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
'      !Remarks = IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text)
'      !ManualBillNo = IIf(Trim(TxtManualBillNo.Text) = "", "", TxtManualBillNo.Text)
'      !OtherCharges = IIf(Val(TxtOtherCharges.Text) = 0, Null, Val(TxtOtherCharges.Text))
'      !TotalExpense = IIf(Val(TxtTotalExpense.Text) = 0, Null, Val(TxtTotalExpense.Text))
'      !UserNo = vUser
'      .Update
'      .Close
'   End With
   
'   With RsBody
'      .Filter = 0
'      .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
''         sSql = "update Products set PurPrice = " & RsBody!Price & ", PurDiscPC = " & RsBody!DiscPC & ", PurchasePackingID = " & IIf(IsNull(RsBody!PackingID), "Null", RsBody!PackingID) & " Where ProductID='" & RsBody!ProductID & "'"
''         cn.Execute ssql
'         If (Not IsNull(RsBody!PackingID)) And (Not IsNull(RsBody!Multiplier)) And (RsBody!Multiplier <> 0) Then
'            If cn.Execute("select * from ProductPacking Where ProductID='" & RsBody!Productid & "' and PackingID = " & RsBody!PackingID).RecordCount = 0 Then
'               ssql = "INSERT INTO ProductPacking(PackingID,Multiplier,ProductID) VALUES ('" & RsBody!PackingID & "','" & RsBody!Multiplier & "','" & RsBody!Productid & "')"
'               cn.Execute ssql
'            Else
'               ssql = "update ProductPacking set Multiplier = " & IIf(IsNull(RsBody!Multiplier), 0, RsBody!Multiplier) & " Where ProductID='" & RsBody!Productid & "' and PackingID = " & RsBody!PackingID
'               cn.Execute ssql
'            End If
'         End If
'         .MoveNext
'      Next vCounter
'      .UpdateBatch
'   End With
'
'   If RsBodySerial.RecordCount > 0 Then
'     With RsBodySerial
'      .Filter = 0
'      .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         .MoveNext
'      Next vCounter
'      .UpdateBatch
'     End With
'   End If
'
'   With RsProductOffer
'      .Filter = 0
'      If Not .EOF Then
'        .MoveFirst
'        For vCounter = 1 To .RecordCount
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         .MoveNext
'        Next vCounter
'      End If
'      .UpdateBatch
'   End With
'
'    With RsExpense
'      .Filter = 0
'      If Not .EOF Then
'        .MoveFirst
'        For vCounter = 1 To .RecordCount
'         !BillID = Val(TxtBillID.Text)
'         !BillDate = DtpBillDate.DateValue
'         .MoveNext
'        Next vCounter
'      End If
'      .UpdateBatch
'   End With
      
      ssql = " select sob.ProductID, ProductName, (isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty - isnull(uqty,0) as Qtyloose, sob.*" & vbCrLf _
      + " from (select OrderID, OrderDate, ProductID, Sum((isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty) as UQty from SaleBody b inner join SaleHeader h on H.SID = B.SID Group By OrderID, OrderDate, ProductID) b " & vbCrLf _
      + " right outer join SaleOrderBody sob on sob.OrderID = b.orderid and sob.OrderDate = b.orderdate and b.ProductID = sob.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = sob.productid" & vbCrLf _
      + " where sob.OrderID = " & Val(TxtOrderID.Text) & " and sob.OrderDate = '" & DtpOrderDate.DateValue & "' and (isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty - isnull(uqty,0)  <> 0"
   
   With CN.Execute(ssql)
      If .RecordCount = 0 Then
         CN.Execute ("Update SaleOrderHeader set IsSale = 1 Where OrderID = " & Val(TxtOrderID.Text) & " And Orderdate ='" & DtpOrderDate.DateValue & "'")
      End If
   End With

'   If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
'      Call BtnPrint_Click
'   End If

'   If vIsNewRecord = True Then Call ActivityLogSale("Sale Invoice", eAdd, TxtBillID.Text, DtpBillDate.DateValue)
   
   '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = " Saved ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt" & TxtNetAmount.Text
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               CN.Execute ssql
            End If
         Next
   End If
   
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmSaleInvoiceDIS, eAdd, TxtBillID.Text, DtpBillDate.DateValue, vGridRows & " New Product/s Added Amount: " & Val(TxtNetAmount.Text))
   
   CN.CommitTrans
   If ChkIsPreview.Value = 1 Or ChkIsPrint.Value = 1 Then
      Call BtnPrint_Click
   End If
   
   
'   cn.Close
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
'   RsBody.Open "Select * from SaleBody where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      ssql = "select p.ProductName, EmpName, StoreName, code, b.*, dbo.GetExpiryDate(b.ProductID,BatchNo) as ExpiryDate from SaleBody b join products p on p.productid = b.productid left outer join Employees e on e.empid = b.empid left outer join Stores s on s.StoreID = b.StoreID where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "' and sid = " & Val(TxtSID.Text) & " Order by SerialNo asc "
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("Serial").Text = Grid.rows
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = !Code
            If ObjRegistry.AllowEmployeProductWise Then
               Grid.Columns("EmpID").Text = IIf(IsNull(!EmpID), "", !EmpID)
               Grid.Columns("EmpName").Text = IIf(IsNull(!empname), "", !empname)
            End If
            If ObjRegistry.AllowStoreProductWise Then
               Grid.Columns("StoreID").Text = IIf(IsNull(!StoreID), "", !StoreID)
               Grid.Columns("StoreName").Text = IIf(IsNull(!StoreName), "", !StoreName)
            End If
            Grid.Columns("BatchNo").Text = IIf(IsNull(!BatchNo), "", !BatchNo)
            Grid.Columns("ExpiryDate").Value = IIf(IsNull(!ExpiryDate), "", !ExpiryDate)
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
            Grid.Columns("GrossQty").Value = IIf(IsNull(!GrossQty), "", !GrossQty)
            Grid.Columns("GrossUnit").Value = IIf(IsNull(!GrossUnit), "", !GrossUnit)
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            Grid.Columns("QtyLoose").Value = !Qty
            Grid.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            Grid.Columns("Price").Value = !Price
            Grid.Columns("ReSPrice").Value = !ReSPrice
            
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            Grid.Columns("isProduct").Value = !isProduct
            
            Grid.Columns("RetailPrice").Value = !RetailPrice
            Grid.Columns("IsWSDiscb4ST").Value = !IsWSDiscb4ST
            Grid.Columns("IsWSSaleTax").Value = !IsWSSaleTax
            Grid.Columns("IsRetailSaleTax").Value = !IsRetailSaleTax
            Grid.Columns("TokenVal").Value = IIf(IsNull(!TokenVal), "", !TokenVal)
            
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            
            Grid.Columns("isLastPrice").Value = IIf(IsNull(!isLastPrice), "", !isLastPrice)
            
            Grid.Columns("isDiscB4TradeOffer").Value = IIf(IsNull(!isDiscB4TradeOffer), "", !isDiscB4TradeOffer)
            Grid.Columns("isDiscB4ExtraScheme").Value = IIf(IsNull(!isDiscB4ExtraScheme), "", !isDiscB4ExtraScheme)
            Grid.Columns("isDiscB4SaleTax").Value = IIf(IsNull(!isDiscB4SaleTax), "", !isDiscB4SaleTax)
            Grid.Columns("TradeOffer1").Value = IIf(IsNull(!TradeOffer1), "", !TradeOffer1)
            Grid.Columns("TradeOffer2").Value = IIf(IsNull(!TradeOffer2), "", !TradeOffer2)
            Grid.Columns("ExtraSchemePer").Value = IIf(IsNull(!ExtraSchemePer), "", !ExtraSchemePer)
            Grid.Columns("TradeValue").Value = IIf(IsNull(!TradeValue), "", !TradeValue)
            Grid.Columns("ExtraSchemeValue").Value = IIf(IsNull(!ExtraSchemeValue), "", !ExtraSchemeValue)
            
            
            Grid.Columns("SC").Value = IIf(IsNull(!SC), "", !SC)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("ReSAmount").Value = !ReSAmount
            Grid.Columns("DiscAmount").Value = !DiscAmount
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(!Amount)
            TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + !Qty + IIf(IsNull(!Bonus), "0", !Bonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack))
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
'   End If
    TxtTotalItems.Text = Val(Grid.rows) - 1
   RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   RsBodySerial.Open "Select * from SaleBodySerial where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
   
   If RsBodyStore.State = adStateOpen Then RsBodyStore.Close
   RsBodyStore.Filter = 0
   RsBodyStore.Open "Select * from SaleBodyStore where BillID = " & Val(TxtBillID.Text) & " And SID = " & Val(TxtSID.Text) & " and Billdate = '" & DtpBillDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   Call PopulateDataToGridMultipleStore
   Call PopulateDataToGridOffer
   Call PopulateDataToGridExpense
   Call PopulateDataToGridserial
End Sub

Private Sub PopulateSaleOrderToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleBody where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      ssql = " select sob.ProductID, ProductName, QtyPack - isnull(UPack,0) as RQtyPack, Qty - isnull(UQty,0) as RQty, Bonus - isnull(UBonus,0) as RBonus, sob.*" & vbCrLf _
      + " from (select OrderID, OrderDate, ProductID, Sum(Qty) as UQty, Sum(QtyPack) as UPack, Sum(Bonus) as UBonus from SaleBody b inner join SaleHeader h on H.SID = B.SID Group By OrderID, OrderDate, ProductID) b " & vbCrLf _
      + " right outer join SaleOrderBody sob on sob.OrderID = b.orderid and sob.OrderDate = b.orderdate and b.ProductID = sob.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = sob.productid" & vbCrLf _
      + " where sob.OrderID = " & Val(TxtOrderID.Text) & " and sob.OrderDate = '" & DtpOrderDate.DateValue & "' and (QtyPack - isnull(UPack,0) <> 0 or Qty - isnull(UQty,0) <> 0 or Bonus - isnull(UBonus,0) <> 0) order by serialno"
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         While Not .EOF
            
'         ''' latest STock Comment
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & Val(!Productid)
            With CN.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(!Productid) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(!Productid) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(!Productid) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Caption = LblStock.Caption & " " & " Total Qty: " & vQtyLoose
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
         
         If (ObjRegistry.NegativeSale = False Or vQtyLoose >= 0) Then
            If (Val(vQtyLoose) - ((IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!RQtyPack), 0, !RQtyPack)) + Val(!RQty))) >= 0 Then
               Grid.AddNew
               RsBody.AddNew
               Grid.Columns("ProductID").Text = !Productid
               Grid.Columns("Code").Text = !Code
               Grid.Columns("ProductName").Text = !ProductName
            
               RsBody!Productid = !Productid
               RsBody!Code = !Code
            
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
               Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
               Grid.Columns("isProduct").Value = 1 '!isProduct
               Grid.Columns("RetailPrice").Value = !RetailPrice
               Grid.Columns("IsWSDiscb4ST").Value = !IsWSDiscb4ST
               Grid.Columns("IsWSSaleTax").Value = !IsWSSaleTax
               Grid.Columns("IsRetailSaleTax").Value = !IsRetailSaleTax
               Grid.Columns("TokenVal").Value = IIf(IsNull(!TokenVal), "", !TokenVal)
               Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
               Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
               Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
               Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
               Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
               Grid.Columns("DiscVal").Value = Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty) 'IIf(IsNull(!DiscVal), "", !DiscVal)
               Grid.Columns("Amount").Value = ((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty) '!Amount
               TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty))
               TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + !RQty + IIf(IsNull(!RBonus), "0", !RBonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!RQtyPack), 0, !RQtyPack))
            
               'TxtAmount.Text = Round((Val(vUnitPrice) - Val(TxtDiscPC.Text)) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 3)

               RsBody!Multiplier = !Multiplier
               RsBody!QtyPack = IIf(IsNull(!RQtyPack), 0, !RQtyPack)
               RsBody!Qty = !RQty
               RsBody!Bonus = !RBonus 'IIf(IsNull(), Null, !RBonus)
               RsBody!Price = !Price
               RsBody!Cost = !Cost
               RsBody!isProduct = 1 '!isProduct
               RsBody!RetailPrice = !RetailPrice
               RsBody!IsWSDiscb4ST = !IsWSDiscb4ST
               RsBody!IsWSSaleTax = !IsWSSaleTax
               RsBody!IsRetailSaleTax = !IsRetailSaleTax
               RsBody!TokenVal = !TokenVal
               RsBody!DiscPC = IIf(IsNull(!DiscPC), 0, !DiscPC)
               RsBody!Offer = !Offer
               RsBody!SaleTaxPer = !SaleTaxPer
               RsBody!SaleTaxval = !SaleTaxval
               RsBody!DiscPer = !DiscPer
               RsBody!DiscVal = Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty) 'IIf(IsNull(!DiscVal), "", !DiscVal)
               RsBody!Amount = ((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty) '!Amount
               RsBody.Update
               Else
                  MsgBox "Insufficient Stock Of " & !Productid & " ", vbInformation + vbOKOnly, "Error"
               End If
               
            End If
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
   RsBodySerial.Open "Select * from SaleBodySerial where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
   Call PopulateSaleOrderToGridOffer
   Call PopulateSaleOrderToGridSerial
   Call PopulateSaleOrderToGridExpense
End Sub

Private Sub PopulateDataToGridserial()
'   RsBodySerial.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
'   If RsBodySerial.RecordCount > 0 Then
       ssql = "select d.* from SaleBodySerial d  where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
      With CN.Execute(ssql)
'       With RsBodySerial
         GridSerial.Redraw = False
         GridSerial.MoveFirst
         GridSerial.RemoveAll
         GridSerial.AllowAddNew = True
'         .MoveFirst
         While Not .EOF
            GridSerial.AddNew
            GridSerial.Columns("ProductID").Text = !Productid
            GridSerial.Columns("Serial").Text = !Serial
            .MoveNext
         Wend
      .Close
      End With
      GridSerial.AddNew
      GridSerial.Columns("Serial").Text = " "
      GridSerial.AllowAddNew = False
      GridSerial.Redraw = True
'   Else
'    Call SubClearSerialFields
'   End If
'   RsBodySerial.Filter = 0
End Sub

Private Sub PopulateSaleOrderToGridSerial()
'   RsBodySerial.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
'   If RsBodySerial.RecordCount > 0 Then
       ssql = "select d.* from SaleOrderBodySerial d  where OrderID=" & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'"
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
            RsBodySerial!Serial = !Serial
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
'   If cn.State = adStateClosed Then cn.Open
   vMode = vNewValue
   Select Case vNewValue
   Case Is = NewMode
      If ObjRegistry.SaveAsNewBill = True Then BtnSaveAS.Visible = True Else BtnSaveAS.Visible = False
      Call SubClearFields
      vRandomID = Rnd() * 11111 & " " & Format(Now, "dd/mm hh:mm:ss")
      FrmHistory.Visible = False
      FrmProductPrices.Visible = False
      vZoneID = 0
      
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
      
    
      If Not ObjUserSecurity.IsAdministrator Then BtnOpen.Visible = ObjUserSecurity.OpenForm
      
      DtpBillDate.DateValue = vDate
      
      
      DtpOrderDate.DateValue = DtpBillDate.DateValue
      TxtBillID.Text = FunGetMaxID()
'      TxtStampID.Text = StampID()
      Call PopulateDataToGrid
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      TxtStoreID.Enabled = True
      BtnStore.Enabled = True
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      BtnSaveAS.Enabled = False
      TxtCode.Enabled = True
      TxtStoreID.Enabled = True
      BtnStore.Enabled = True
      LblStock.Visible = False
      LblStockCaption.Visible = False
      BtnProduct.Enabled = True
      'TxtBillID.Enabled = True
      DtpBillDate.Enabled = True
      If ObjUserSecurity.IsAdministrator = False Then DtpBillDate.Enabled = ObjUserSecurity.ChangeDate
      
      If DtpBillDate.Enabled And DtpBillDate.Visible Then
         DtpBillDate.SetFocus
      ElseIf TxtCustomerID.Enabled And TxtCustomerID.Visible Then
         TxtCustomerID.SetFocus
      End If
      GridOffer.Visible = False
      FramExpense.ZOrder 0
      vIsNewRecord = True
      isWholeSale = True
   Case Is = OpenMode
      'TxtBillID.Enabled = False
      FrmProductPrices.Visible = False
'      DtpBillDate.Enabled = False
      If ObjUserSecurity.IsAdministrator = False Then DtpBillDate.Enabled = ObjUserSecurity.ChangeDate
      TxtStoreID.Enabled = ObjUserSecurity.IsAdministrator
      BtnStore.Enabled = ObjUserSecurity.IsAdministrator
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      BtnSaveAS.Enabled = True
      'TxtStoreID.Enabled = False
      'BtnStore.Enabled = False
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      'DtpBillDate.SetFocus
      vIsNewRecord = False
   Case Is = ChangeMode
      If BtnSave.Enabled = False Then vBillTime = Now
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
      If TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus Else TxtCustomerID.SetFocus
   Else
      TxtStoreID.SetFocus
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

Private Sub DtpBillDate_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> DtpBillDate.Name Then Exit Sub
   If FormStatus = OpenMode Then Exit Sub
   TxtBillID.Text = FunGetMaxID()
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
'   If KeyCode = vbKeyReturn Then
'      If ActiveControl.Name = "Grid" Then
'         Grid_DblClick
'      ElseIf ActiveControl.Name = "GridSerial" Then
'         GridSerial_DblClick
'      ElseIf ActiveControl.Name = TxtCode.Name Then
'         If FunSelectProduct(ssValidate, False) = True Then If TxtBatchNo.Visible Then TxtBatchNo.SetFocus Else GetDataFromTexBoxesToGrid
'      Else
'         keybd_event 9, 1, 1, 1
'         KeyCode = 0
'      End If
     
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name
         Grid_DblClick
      Case GridSerial.Name
         GridSerial_DblClick
      Case GridMultipleStore.Name
         GridMultipleStore_DblClick
      Case TxtCode.Name
         If FunSelectProduct(ssValidate, False) = True Then
            If vAutoEnterQtyintoGridSaleInvoice = True And Val(TxtMultiplier.Text) = 0 And Len(TxtCode.Text) > 5 Then
               TxtQtyLoose.Text = IIf(Val(TxtQtyLoose.Text) = 0, vQty, TxtQtyLoose.Text)
               SubCalculateBody
               GetDataFromTexBoxesToGrid
            ElseIf vAutoEnterQtyintoGridSaleInvoice = True And Val(TxtMultiplier.Text) = 1 And Len(TxtCode.Text) > 5 Then
               TxtQtyPack.Text = IIf(Val(TxtQtyPack.Text) = 0, 1, TxtQtyPack.Text)
               SubCalculateBody
               GetDataFromTexBoxesToGrid
            ElseIf vAutoEnterQtyintoGridSaleInvoice = True And Val(TxtMultiplier.Text) > 1 And Len(TxtCode.Text) > 5 Then
               If CmbPackName.Enabled And CmbPackName.Visible Then CmbPackName.SetFocus
            Else
               keybd_event 9, 1, 1, 1: KeyCode = 0
               End If
         End If
      Case TxtQtyPack.Name, TxtQtyLoose.Name
         If vAutoEnterQtyintoGridSaleInvoice = True And Len(TxtCode.Text) > 5 Then
            GetDataFromTexBoxesToGrid
         Else
            keybd_event 9, 1, 1, 1: KeyCode = 0
         End If
      Case TxtDiscVal.Name, TxtSC.Name, TxtAmount.Name, TxtExtraSchemePer.Name
         Flag = True
         If vUseMultipleStore = True Then SubMultipleStore Else GetDataFromTexBoxesToGrid
'         GetDataFromTexBoxesToGrid
         'If Grid.Visible And Grid.Enabled Then Grid.SetFocus
      Case TxtMSQtyLoose.Name
         Flag = True
         GetDataFromTexBoxesToGridMultipleStore
      Case Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtCode.Enabled Then TxtCode.SetFocus: Call SubClearDetailArea
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
            If BtnSave.Enabled And BtnSave.Visible Then BtnSave_Click
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
            If BtnDelete.Enabled And BtnDelete.Visible Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus Else TxtCustomerID.SetFocus Else TxtStoreID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtCustomerID.SetFocus Else TxtOrganizationID.SetFocus
         Case TxtCustomerID.Name: If FunSelectCustomer(ssFunctionKey, False) = True Then TxtBillNo.SetFocus Else TxtCustomerID.SetFocus
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then If TxtMemberID.Visible = True Then TxtMemberID.SetFocus Else If TxtCode.Enabled Then TxtCode.SetFocus
         Case TxtMemberID.Name: If FunSelectMember(ssFunctionKey, True) = True Then If TxtCode.Enabled Then TxtCode.SetFocus Else TxtMemberID.SetFocus
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, False) = True Then If TxtBatchNo.Visible Then TxtBatchNo.SetFocus Else GetDataFromTexBoxesToGrid
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
   ElseIf KeyCode = vbKeyF4 Then
      If TxtProductID.Text <> "" And (ObjUserSecurity.ShowPrice = True Or ObjUserSecurity.IsAdministrator = True) Then
      Select Case ActiveControl.Name
      Case TxtProductID.Name, CmbPackName.Name, TxtMultiplier.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtSC.Name, Grid.Name
            If Val(TxtMultiplier.Text) <> 0 Then
               LblLastPurPrice.Caption = CN.Execute("select PurPrice from products where productid = " & Val(TxtProductID.Text)).Fields(0).Value * Val(TxtMultiplier.Text)
            Else
               LblLastPurPrice.Caption = CN.Execute("select PurPrice from products where productid = " & Val(TxtProductID.Text)).Fields(0).Value
            End If
'            LblLastPurPrice.Visible = True
            LblCost.Caption = LblLastPurPrice.Caption
            Call MniCostPrice_Click
      End Select
      End If
   ElseIf KeyCode = vbKeyF5 Then
      If TxtProductID.Text <> "" And (ObjUserSecurity.ShowPrice = True Or ObjUserSecurity.IsAdministrator = True) Then
      Select Case ActiveControl.Name
      Case TxtProductID.Name, CmbPackName.Name, TxtMultiplier.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtSC.Name, Grid.Name
            If Val(TxtMultiplier.Text) <> 0 Then
               LblLastPurPrice.Caption = CN.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text) & ")").Fields(0).Value * Val(TxtMultiplier.Text)
            Else
               LblLastPurPrice.Caption = CN.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text) & ")").Fields(0).Value
            End If
'            LblLastPurPrice.Visible = True
            LblCost.Caption = LblLastPurPrice.Caption
            Call MniCostPrice_Click
      End Select
      End If
   ElseIf KeyCode = vbKeyF6 Then
      If TxtProductID.Text <> "" And (ObjUserSecurity.ShowPrice = True Or ObjUserSecurity.IsAdministrator = True) Then
      Select Case ActiveControl.Name
      Case TxtProductID.Name, CmbPackName.Name, TxtMultiplier.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtSC.Name, Grid.Name
            CN.Execute "exec SPProductAverageCost '" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text)
            LblCost.Caption = CN.Execute("Select Price from TempPurchase Where Productid = " & Val(TxtProductID.Text)).Fields(0).Value
            Call MniCostPrice_Click
      End Select
      End If
   ElseIf KeyCode = vbKeyF7 Then
      If TxtProductID.Text <> "" And (ObjUserSecurity.ShowPrice = True Or ObjUserSecurity.IsAdministrator = True) Then
      Select Case ActiveControl.Name
      Case TxtProductID.Name, CmbPackName.Name, TxtMultiplier.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtSC.Name, Grid.Name
            LblCost.Caption = CN.Execute("Select WSPrice from Products Where Productid = " & Val(TxtProductID.Text)).Fields(0).Value
            Call MniCostPrice_Click
      End Select
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
   
'   Set cnSale = cn
'   If cnSale.State = adStateOpen Then cnSale.Close
'   cnSale.Open
'   cnSale.CursorLocation = adUseClient


   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Sale Invoice"
   HelpLocation Me
   
   
   vSystemDate = Abs(ObjRegistry.SystemDate)
   vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
   
   
'   With cn.Execute("Select * from Packings")
'      CmbPackName.AddItem ""
'      While Not .EOF
'         CmbPackName.AddItem !Packingname
'         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
'         .MoveNext
'      Wend
'      .Close
'   End With
   
   If ObjUserSecurity.ShowStock = True Or ObjUserSecurity.IsAdministrator Then
      vShowStock = True
   Else
      vShowStock = False
   End If
   
   With CN.Execute("Select * from Stores")
      While Not .EOF
         CmbMSStore.AddItem !StoreName
         CmbMSStore.ItemData(CmbMSStore.NewIndex) = !StoreID
         .MoveNext
      Wend
      .Close
   End With
   CmbMSStore.ListIndex = 0
   
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
   ssql = "select * from FormDefaultSetting Where FormType = 'Sale Invoice DIS' and LocalComputerName = '" & LocalComputerName & "'"
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
'
   BtnBatchPrint.Visible = ObjRegistry.ShowBatchPrint
   Frame2.Visible = ObjRegistry.FreightVisible
   LblFreight.Visible = ObjRegistry.FreightVisible
   TxtFreight.Visible = ObjRegistry.FreightVisible
  
   lblPayable.Visible = ObjRegistry.PreviousBalanceVisible
   LblTtlPayable.Visible = ObjRegistry.PreviousBalanceVisible
   TxtPreviousReceivable.Visible = ObjRegistry.PreviousBalanceVisible
   TxtTotalReceivable.Visible = ObjRegistry.PreviousBalanceVisible

   TxtStoreID.Text = IIf((ObjRegistry.StoreID = ""), "", ObjRegistry.StoreID)
   FunSelectStore ssValidate, True
                  
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
         
   TxtBatchNo.Visible = ObjRegistry.BatchNoVisible
   TxtExpiryDate.Visible = ObjRegistry.BatchNoVisible
   
   BtnPrintWarranty.Visible = ObjRegistry.ShowWarrantyinSaleInvoice
   LblLicenceNO.Visible = ObjRegistry.ShowWarrantyinSaleInvoice
   TxtLicenceNO.Visible = ObjRegistry.ShowWarrantyinSaleInvoice
   
   If ObjRegistry.BatchNoVisible = False Then LblProductName.Left = TxtProductName.Left

   
   LblEmpID.Visible = ObjRegistry.EmpVisible
   LblEmpName.Visible = ObjRegistry.EmpVisible
   TxtEmployeeID.Visible = ObjRegistry.EmpVisible
   TxtEmployeeName.Visible = ObjRegistry.EmpVisible
   BtnEmployee.Visible = ObjRegistry.EmpVisible
   
   LblReSPrice.Visible = ObjRegistry.ShowReSale
   TxtReSPrice.Visible = ObjRegistry.ShowReSale
   LblReSAmount.Visible = ObjRegistry.ShowReSale
   TxtReSAmount.Visible = ObjRegistry.ShowReSale
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible

   LblMemberID.Visible = ObjRegistry.MemberVisible
   LblMemberName.Visible = ObjRegistry.MemberVisible
   TxtMemberID.Visible = ObjRegistry.MemberVisible
   TxtMemberName.Visible = ObjRegistry.MemberVisible
   BtnMember.Visible = ObjRegistry.MemberVisible

   TxtManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   LblManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   
   TxtRemarks.Visible = ObjRegistry.RemarksVisible
   LblRemarks.Visible = ObjRegistry.RemarksVisible
   
   LblSyllabusID.Visible = ObjRegistry.ShowSyllabus
   LblSyllabusName.Visible = ObjRegistry.ShowSyllabus
   TxtSyllabusID.Visible = ObjRegistry.ShowSyllabus
   TxtSyllabusName.Visible = ObjRegistry.ShowSyllabus
   BtnSyllabus.Visible = ObjRegistry.ShowSyllabus
   
   If ObjRegistry.ShowBonus = False Then
    LblBonus.Visible = False
    TxtBonus.Visible = False
    Grid.Columns("Bonus").Visible = False
    
    LblOffer.Left = LblOffer.Left - TxtBonus.Width
    TxtOffer.Left = TxtOffer.Left - TxtBonus.Width
    
    LblPrice.Left = LblPrice.Left - TxtBonus.Width
    TxtPrice.Left = TxtPrice.Left - TxtBonus.Width
    
    LblDiscPC.Left = LblDiscPC.Left - TxtBonus.Width
    TxtDiscPC.Left = TxtDiscPC.Left - TxtBonus.Width
    
    LblSaleTaxPer.Left = LblSaleTaxPer.Left - TxtBonus.Width
    TxtSaleTaxPer.Left = TxtSaleTaxPer.Left - TxtBonus.Width
    
    LblDiscPer.Left = LblDiscPer.Left - TxtBonus.Width
    TxtDiscPer.Left = TxtDiscPer.Left - TxtBonus.Width
    
    LblDiscVal.Left = LblDiscVal.Left - TxtBonus.Width
    TxtDiscVal.Left = TxtDiscVal.Left - TxtBonus.Width
    
    LblSC.Left = LblSC.Left - TxtBonus.Width
    TxtSC.Left = TxtSC.Left - TxtBonus.Width
    
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
    
    LblDiscPC.Left = LblDiscPC.Left - TxtOffer.Width
    TxtDiscPC.Left = TxtDiscPC.Left - TxtOffer.Width
    
    LblSaleTaxPer.Left = LblSaleTaxPer.Left - TxtOffer.Width
    TxtSaleTaxPer.Left = TxtSaleTaxPer.Left - TxtOffer.Width
    
    LblDiscPer.Left = LblDiscPer.Left - TxtOffer.Width
    TxtDiscPer.Left = TxtDiscPer.Left - TxtOffer.Width
    
    LblDiscVal.Left = LblDiscVal.Left - TxtOffer.Width
    TxtDiscVal.Left = TxtDiscVal.Left - TxtOffer.Width
    
    LblSC.Left = LblSC.Left - TxtOffer.Width
    TxtSC.Left = TxtSC.Left - TxtOffer.Width
    
    LblAmount.Left = LblAmount.Left - TxtOffer.Width
    TxtAmount.Left = TxtAmount.Left - TxtOffer.Width
    
    Grid.Width = Grid.Width - TxtOffer.Width
   End If
  
   LblSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   TxtSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   LblGSTValue.Visible = ObjRegistry.ShowSaleTax
   TxtSaleTaxVal.Visible = ObjRegistry.ShowSaleTax
   ChkDiscB4SaleTax.Visible = ObjRegistry.ShowSaleTax
   
   If ObjRegistry.ShowSaleTax = False Then
    
    
    Grid.Columns("SaleTaxPer").Visible = False
        
    LblDiscPer.Left = LblDiscPer.Left - TxtSaleTaxPer.Width
    TxtDiscPer.Left = TxtDiscPer.Left - TxtSaleTaxPer.Width
    
    LblDiscVal.Left = LblDiscVal.Left - TxtSaleTaxPer.Width
    TxtDiscVal.Left = TxtDiscVal.Left - TxtSaleTaxPer.Width
    
    LblSC.Left = LblSC.Left - TxtSaleTaxPer.Width
    TxtSC.Left = TxtSC.Left - TxtSaleTaxPer.Width
    
    LblAmount.Left = LblAmount.Left - TxtSaleTaxPer.Width
    TxtAmount.Left = TxtAmount.Left - TxtSaleTaxPer.Width
    
    Grid.Width = Grid.Width - TxtSaleTaxPer.Width
   End If
 
   If ObjRegistry.ShowSC = False Then
    LblSC.Visible = False
    TxtSC.Visible = False
    Grid.Columns("SC").Visible = False
    
    LblAmount.Left = LblAmount.Left - TxtSC.Width
    TxtAmount.Left = TxtAmount.Left - TxtSC.Width
    
    Grid.Width = Grid.Width - TxtSC.Width
    
   End If
   
   ''''''' Show Trade '''''''
   vTradeOffer = ObjRegistry.ShowTradeOffer
   LblTradeOffer.Visible = vTradeOffer
   TxtTradeOffer1.Visible = vTradeOffer
   TxtTradeOffer2.Visible = vTradeOffer
   LblTradeValue.Visible = vTradeOffer
   TxtTradeOfferValue.Visible = vTradeOffer
   LblPlusSign.Visible = vTradeOffer
   ChkDiscB4TradeOffer.Visible = vTradeOffer
   
   LblExtraSchemePer.Visible = vTradeOffer
   TxtExtraSchemePer.Visible = vTradeOffer
   LblExtraSchemeValue.Visible = vTradeOffer
   TxtExtraSchemeValue.Visible = vTradeOffer
   ChkDiscB4ExtraScheme.Visible = vTradeOffer
   ''''''''''''''''''''''''''''''
   
   vAutoEnterQtyintoGridSaleInvoice = ObjRegistry.AutoEnterQtyintoGridSaleInvoice
   vUseMultipleStore = ObjRegistry.UseMultipleStore
   
     
   If ObjRegistry.ShowPromiseDateInSalaPurchase = True Then
      LblPromiseDate.Visible = True
      DtpPromiseDate.Visible = True
      DtpPromiseDate.DateValue = Null
   Else
      LblPromiseDate.Visible = False
      DtpPromiseDate.Visible = False
      DtpPromiseDate.DateValue = Null
   End If
   If ObjRegistry.ShowDispatchDate = True Then
      LblDispatchDate.Visible = True
      DtpDispatchDate.Visible = True
      DtpDispatchDate.DateValue = DtpBillDate.DateValue
      LblTerms.Visible = True
      TxtTerms.Visible = True
      TxtTerms.Text = CN.Execute("Select value from sysindexs Where Registrykey = 'SalePromiseTerms'").Fields(0).Value
   Else
      LblDispatchDate.Visible = False
      DtpDispatchDate.Visible = False
      DtpDispatchDate.DateValue = Null
      LblTerms.Visible = False
      TxtTerms.Visible = False
      TxtTerms.Text = 0
   End If
   
   If ObjRegistry.ShowExpiryInvoice = True Then
      LblExpiryInvoice.Visible = True
      DtpExpiryInvoice.Visible = True
      DtpExpiryInvoice.DateValue = Date
   Else
      LblExpiryInvoice.Visible = False
      DtpExpiryInvoice.Visible = False
      DtpExpiryInvoice.DateValue = Null
   End If
   
   
   
   LblGrossQty.Visible = ObjRegistry.IsGrossQty
   TxtGrossQty.Visible = ObjRegistry.IsGrossQty
   LblGrossUnit.Visible = ObjRegistry.IsGrossQty
   TxtGrossUnit.Visible = ObjRegistry.IsGrossQty
 
   If ObjUserSecurity.IsAdministrator = False Then
      TxtDiscPC.Enabled = ObjRegistry.DiscAllowed
      TxtDiscPer.Enabled = ObjRegistry.DiscAllowed
      TxtDiscVal.Enabled = ObjRegistry.DiscAllowed
      TxtBillDisc.Enabled = ObjRegistry.DiscAllowed
      TxtBillDiscPer.Enabled = ObjRegistry.DiscAllowed
      If ObjRegistry.DiscAllowed = False Then
         TxtDiscPC.Tag = "NC"
         TxtDiscPer.Tag = "NC"
         TxtDiscVal.Tag = "NC"
         TxtBillDisc.Tag = "NC"
         TxtBillDiscPer.Tag = "NC"
      End If
   End If
   
   If ObjUserSecurity.IsAdministrator = True Then
      TxtPrice.Enabled = True
      TxtPrice.Tag = ""
   Else
      TxtPrice.Enabled = ObjUserSecurity.ChangePriceSaleInvoice
      TxtPrice.Tag = IIf(TxtPrice.Enabled = True, "", "D")
   End If
   DateFlag = True
   isUrdu = False
   vNoofPrints = IIf(IsNull(ObjRegistry.NoofPrints) Or Val(ObjRegistry.NoofPrints) = 0, 1, ObjRegistry.NoofPrints)
   With CN.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         vNoofPrints = IIf(IsNull(!NoofPrints) Or Val(!NoofPrints) = 0, 1, !NoofPrints)
         FunSelectStore ssValidate, True
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
         If ObjRegistry.ChangePrice = True Then TxtPrice.Enabled = True
         vNoofPrints = IIf(IsNull(!NoofPrints) Or !NoofPrints = 0, 1, !NoofPrints)
      End If
      .Close
   End With
 
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   
   LblOtherChargesCaption.Caption = ObjRegistry.ChargesName
   
   'Grid.Left = (Me.Width / 2) - (Grid.Width / 2)
   vServerDate = CN.Execute("Select CONVERT(datetime, CONVERT(varchar, GETDATE(), 110)) ServerDate").Fields(0).Value
   DateFlag = True
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpBillDate.IsDateValid = False Then Exit Function
    
   If ObjRegistry.AllowContinuousBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader").Fields(0)
   ElseIf ObjRegistry.AllowMonthlyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where Month(BillDate) = '" & Month(DtpBillDate.DateValue) & "' and  year(BillDate) ='" & Year(DtpBillDate.DateValue) & "'").Fields(0)
   ElseIf ObjRegistry.AllowDailyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where BillDate = '" & DtpBillDate.DateValue & "'").Fields(0)
   Else
      FunGetMaxID = CN.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where BillDate = '" & DtpBillDate.DateValue & "' and StoreID = " & TxtStoreID.Text).Fields(0)
   End If
  
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function StampID() As Long
   On Error GoTo ErrorHandler
   StampID = CN.Execute("Select isnull(max(SID),0)+1 from Stamp").Fields(0)
   CN.Execute "update Stamp set SID = " & StampID
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpBillDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = CN.Execute("Select isnull(max(BillID),0)+1 from Bin_SaleHeader").Fields(0)
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
   TxtBillDisc.Text = ""
   TxtBillDiscPer.Text = ""
   TxtNetAmount.Text = 0
   TxtQtyLoose.Enabled = True
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("Code").Text = " "
   Grid.Update
   
   CmbMSPackName.Clear
   CmbMSPackName.AddItem ""
   CmbMSPackName.ListIndex = 0
   
   LblCustomerDesc.Caption = ""
   
   GridOffer.CancelUpdate
   GridOffer.RemoveAll
   GridOffer.AddNew
   GridOffer.Columns("ProductID").Text = " "
   GridOffer.Update
   DtpPromiseDate.DateValue = Null
   If ObjRegistry.ShowExpiryInvoice = True Then DtpExpiryInvoice.DateValue = Date Else DtpExpiryInvoice.DateValue = Null
   
   GridOffer.Visible = False
   FrmExpiry.Visible = False
   Call SubClearSerialFields
   Call SubClearMultipleStoreFields
   If ObjRegistry.ChangeQtyOnChangedPrice = True Then TxtAmount.Enabled = True
   If ObjRegistry.ShowDispatchDate = True And ObjRegistry.ShowPromiseDateInSalaPurchase Then
      DtpPromiseDate.DateValue = DateAdd("d", Val(TxtTerms.Text), DtpDispatchDate.DateValue)
   End If
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
      Else
          
         Dim frmObj As Object
         For Each frmObj In Forms
            Set frmObj = Nothing
         Next
         Set RsBodyStore = Nothing
         Set RsBody = Nothing
         Set RsBodySerial = Nothing
         Set RsProductOffer = Nothing
         Set FrmSaleInvoiceDist = Nothing
    End If
   
   End If
    '''''''''''''''''' ActivityLogBin For Close Action
'      Call DeleteTempActivityLogBin(vRandomID)
      If Grid.rows > 1 And Cancel = 0 Then
         vGridRows = 0
         Grid.Redraw = False
         Grid.MoveFirst
         For vCounter = 2 To Grid.rows
            vGridRows = vGridRows + 1
            If Trim(Grid.Columns("Code").Text) <> "" Then
               ssql = "Select Productid From salebody where SID=" & Val(TxtSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
               With CN.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmSaleInvoiceDIS, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Closed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmSaleInvoiceDIS, eCloseSavedRecord, TxtBillID.Text, DtpBillDate.DateValue, vGridRows & " Product/s Closed")
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
   TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("Amount").Value
   TxtTotalQtys.Text = Val(TxtTotalQtys.Text) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Grid.Columns("QtyPack").Value)))
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
'      If vUseMultipleStore = True Then TxtMSQtyLoose.SetFocus Else TxtCode.SetFocus
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
'    Call PopulateDataToGridserial
   If vUseMultipleStore Then
      If FrmMultipleStore.Visible = False Then FrmMultipleStore.Visible = True
      Call PopulateDataToGridMultipleStore
   End If
    
    If Trim(Grid.Columns("Code").Text) = "" Then
        TxtSerial.Enabled = False
    Else
        TxtSerial.Enabled = True
    End If
    
    GridOffer.MoveFirst
    For vCounter = 1 To GridOffer.rows
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
       '/******* Mobile SMS *************/
   vStrDetail = ""
   vStrDetail = vStrDetail & " (P" & Grid.Columns("ProductID").Text & IIf(Val(Grid.Columns("Pack").Value) = 0, "", " M" & Grid.Columns("Pack").Value) & IIf(Val(Grid.Columns("QtyPack").Value) = 0, "", " QP" & Grid.Columns("QtyPack").Value) & IIf(Val(Grid.Columns("QtyLoose").Value) = 0, "", " QL" & Grid.Columns("QtyLoose").Value) & IIf(Val(Grid.Columns("Bonus").Value) = 0, "", " QB" & Grid.Columns("Bonus").Value) & " A" & Grid.Columns("Amount").Text & ")"
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnClear Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = " Removed Item ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt" & TxtNetAmount.Text
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               CN.Execute ssql
            End If
         Next
   End If
   ssql = "Select Productid From salebody where sid = " & Val(TxtSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
   With CN.Execute(ssql)
      If .EOF Then
         Call ActivityLogBin("", eFrmSaleInvoiceDIS, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      Else
         Call ActivityLogBin("", eFrmSaleInvoiceDIS, eRemoveRow, TxtBillID.Text, DtpBillDate.DateValue, "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
         Call ActivityLogBin(vRandomID, eFrmSaleInvoiceDIS, eAddTempRecord, TxtBillID.Text, DtpBillDate.DateValue, "Pending Remove Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      End If
   End With
'      RsBody.Filter = "ProductID = '" & TxtProductID.Text & "'" & IIf(ObjRegistry.BatchNoVisible = True, IIf(Trim(TxtBatchNo.Text) = "", "", " and BatchNo = '" & Trim(TxtBatchNo.Text) & "'"), "") & IIf(ObjRegistry.SeperateProductWithPrice = True, " and Price = " & Val(TxtPrice.Text), "") & IIf(ObjRegistry.AllowEmployeProductWise = True, IIf(Trim(TxtEmployeeID.Text) = "", "", " and EmpID = '" & Trim(TxtEmployeeID.Text) & "'"), "") & IIf(ObjRegistry.AllowStoreProductWise = True, " and StoreID = " & Val(TxtStoreID.Text), "")
'      If RsBody.RecordCount > 0 Then RsBody.Delete
      RsProductOffer.Filter = "ProductID = " & Val(GridOffer.Columns("ProductID").Text)
      If RsProductOffer.RecordCount > 0 Then
         RsProductOffer.Delete
         GridOffer.SelBookmarks.RemoveAll
         GridOffer.SelBookmarks.Add GridOffer.Bookmark
         GridOffer.DeleteSelected
         GridOffer.Refresh
         RsProductOffer.Filter = 0
      End If
      RsBodySerial.Filter = "ProductID = " & Val(TxtProductID.Text) & " And Serial = '" & TxtSerial.Text & "'"
      If RsBodySerial.RecordCount > 0 Then
         With RsBodySerial
   '      .Filter = 0
            .MoveFirst
            For vCounter = 1 To .RecordCount
               RsBodySerial.Delete
               .MoveNext
            Next vCounter
         End With
      End If

      RsBodyStore.Filter = "ProductID = " & Val(TxtProductID.Text)
      If RsBodyStore.RecordCount > 0 Then
         With RsBodyStore
'        .Filter = 0
         .MoveFirst
         For vCounter = 1 To .RecordCount
            RsBodyStore.Delete
            .MoveNext
         Next vCounter
       End With
      End If

      CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed Code-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
      
      Grid.SelBookmarks.RemoveAll
      Grid.SelBookmarks.Add Grid.Bookmark
      Grid.DeleteSelected
      Grid.Refresh
      RsBody.Filter = 0
      Grid.MoveLast
      SubClearMultipleStoreFields
      GetDataBackFromGridToTexBoxes
   ElseIf Me.ActiveControl.Name = "GridSerial" Then
      If Trim(GridSerial.Columns("Serial").Text) = "" Then Exit Sub
      RsBodySerial.Filter = "Serial = '" & TxtSerial.Text & "'"
      If RsBodySerial.RecordCount > 0 Then RsBodySerial.Delete
       CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed Code-" & GridSerial.Columns("ProductID").Text & " Serial-" & GridSerial.Columns("Serial").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
       GridSerial.SelBookmarks.RemoveAll
       GridSerial.SelBookmarks.Add GridSerial.Bookmark
       GridSerial.DeleteSelected
       GridSerial.Refresh
       RsBodySerial.Filter = 0
   '    GridSerial.MoveLast
       GetDataBackFromGridSerialToTexBoxes
   ElseIf Me.ActiveControl.Name = "GridMultipleStore" Then
      If Trim(GridMultipleStore.Columns("ProductID").Text) = "" Then Exit Sub
      RsBodyStore.Filter = "ProductID = " & Val(TxtProductID.Text) & " and StoreID = " & GridMultipleStore.Columns("StoreID").Text
      If RsBodyStore.RecordCount > 0 Then RsBodyStore.Delete
       CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed Code-" & GridSerial.Columns("ProductID").Text & " Store-" & GridMultipleStore.Columns("StoreID").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
       GridMultipleStore.SelBookmarks.RemoveAll
       GridMultipleStore.SelBookmarks.Add GridMultipleStore.Bookmark
       GridMultipleStore.DeleteSelected
       GridMultipleStore.Refresh
       RsBodyStore.Filter = 0
   '    GridSerial.MoveLast
'       PopulateDataToGridMultipleStore
       GetDataBackFromGridMultipleStoreToTexBoxes
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
   TxtTotalItems.Text = Val(Grid.rows) - 1
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataBackFromGridSerialToTexBoxes()
   On Error GoTo ErrorHandler
   With GridSerial
      TxtSerial.Text = .Columns("Serial").Text
   End With
   If GridSerial.rows = 1 Then GridSerial.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
On Error GoTo ErrorHandler
 Dim vrowcounter As Integer
   
   If FunValidationDetailData = False Then Exit Sub
   
'   If Trim(Grid.Columns("Productid").Text) = "" Then
'      RsBody.Filter = "ProductID = '" & TxtProductID.Text & "'" & IIf(ObjRegistry.BatchNoVisible = True, IIf(Trim(TxtBatchNo.Text) = "", "", " and BatchNo = '" & Trim(TxtBatchNo.Text) & "'"), "") & IIf(ObjRegistry.SeperateProductWithPrice = True, " and Price = " & Val(TxtPrice.Text), "") & IIf(ObjRegistry.AllowEmployeProductWise = True, IIf(Trim(TxtEmployeeID.Text) = "", "", " and EmpID = '" & Trim(TxtEmployeeID.Text) & "'"), "") & IIf(ObjRegistry.AllowStoreProductWise = True, " and StoreID = " & Val(TxtStoreID.Text), "")
'   Else
'      RsBody.Filter = "ProductID = '" & Grid.Columns("Productid").Text & "'" & IIf(ObjRegistry.BatchNoVisible = True, IIf(Trim(Grid.Columns("BatchNo").Text) = "", "", " and BatchNo = '" & Trim(Grid.Columns("BatchNo").Text) & "'"), "") & IIf(ObjRegistry.SeperateProductWithPrice = True, " and Price = " & Val(Grid.Columns("Price").Text), "") & IIf(ObjRegistry.AllowEmployeProductWise = True, IIf(Trim(Grid.Columns("EmpID").Text) = "", "", " and EmpID = '" & Trim(Grid.Columns("EmpID").Text) & "'"), "") & IIf(ObjRegistry.AllowStoreProductWise = True, " and StoreID = " & Val(Trim(Grid.Columns("StoreID").Text)), "")
'   End If

   If TxtCode.Enabled Then
'      If RsBody.RecordCount = 0 Then
'         RsBody.AddNew
         
'         RsBody!Productid = TxtProductID.Text
'         RsBody!code = TxtCode.Text
'         RsBody!Price = Val(TxtPrice.Text)
'         RsBody!BatchNo = Trim(TxtBatchNo.Text)
       
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.rows
               If Grid.Columns("Productid").Text = TxtProductID.Text And IIf(ObjRegistry.BatchNoVisible = True, Grid.Columns("BatchNo").Text = Trim(TxtBatchNo.Text), True) And IIf(ObjRegistry.SeperateProductWithPrice = True, Val(Grid.Columns("Price").Text) = Val(TxtPrice.Text), True) Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  If ObjRegistry.NegativeSale = False Then
                     If vIsNewRecord = True Then
                        If (Val(vQtyLoose) - (Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)) - (Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value) + Val(Grid.Columns("QtyLoose").Value))) < 0 Then
                           MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
                           Grid.Redraw = True
                           Call SubClearDetailArea
                           Grid.MoveLast
                           If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
                           Exit Sub
                        End If
                     End If
                  End If
                  ssql = "Select Productid From salebody where sid=" & Val(TxtSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmSaleInvoiceDIS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmSaleInvoiceDIS, eEdit, TxtBillID.Text, DtpBillDate.DateValue, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  
                  '''''''''''''''''''''''''This QtyOffer Is used for DetailGrid
                  QtyOffer = Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value) + Val(Grid.Columns("QtyLoose").Value)
                  GetDataFromTextBoxesToGridOffer
                  TxtOffer.Text = Val(TxtOffer.Text) + Val(Grid.Columns("Offer").Text)
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  TxtQtyLoose.Text = Val(TxtQtyLoose.Text) + Val(Grid.Columns("QtyLoose").Value)
                  TxtQtyPack.Text = Val(TxtQtyPack.Text) + Val(Grid.Columns("QtyPack").Value)
                  Call FindRebate
                  TxtTradeOfferValue.Text = Val(TxtTradeOfferValue.Text) + Val(TxtTradeOfferValue.Text) - Val(Grid.Columns("TradeValue").Text)
                  TxtExtraSchemeValue.Text = Val(TxtExtraSchemeValue.Text) + Val(TxtExtraSchemeValue.Text) - Val(Grid.Columns("ExtraSchemeValue").Text)
                  Call SubCalculateBody
   
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  TxtSumDiscAmount.Text = Val(TxtSumDiscAmount.Text) + Val(TxtDiscAmount.Text) - Val(Grid.Columns("DiscAmount").Text)
                  
                  TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Val(Grid.Columns("QtyLoose").Value) + Val(Grid.Columns("Bonus").Value) + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
                  If ObjRegistry.AllowEmployeProductWise Then
                     Grid.Columns("EmpID").Text = TxtEmployeeID.Text
                     Grid.Columns("EmpName").Text = TxtEmployeeName.Text
                  End If
                  If ObjRegistry.AllowStoreProductWise Then
                     Grid.Columns("StoreID").Text = TxtStoreID.Text
                     Grid.Columns("StoreName").Text = TxtStoreName.Text
                  End If
                  'Grid.Columns("ProductID").Text = TxtProductID.Text
                  'Grid.Columns("Code").Text = TxtCode.Text
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
                  Grid.Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
                  Grid.Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, "", Val(TxtMultiplier.Text))
                  Grid.Columns("GrossQty").Value = IIf(Val(TxtGrossQty.Text) = 0, Null, Val(TxtGrossQty.Text))
                  Grid.Columns("GrossUnit").Value = IIf(Val(TxtGrossUnit.Text) = 0, Null, Val(TxtGrossUnit.Text))
                  Grid.Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
                  Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
                  Grid.Columns("DiscAmount").Value = Val(TxtDiscAmount.Text)
                  Grid.Columns("Bonus").Value = Val(TxtBonus.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("ReSPrice").Value = Val(TxtReSPrice.Text)
                  Grid.Columns("isLastPrice").Value = Abs(IIf(Val(LblLastPrice.Caption) = Val(TxtPrice.Text), 1, 0))
                  Grid.Columns("RetailPrice").Value = Val(TxtRetailPrice.Text)
                  Grid.Columns("IsWSDiscb4ST").Value = vIsWSDiscb4ST
                  Grid.Columns("IsWSSaleTax").Value = vIsWSSaleTax
                  Grid.Columns("IsRetailSaleTax").Value = vIsRetailSaleTax
                  Grid.Columns("TokenVal").Value = IIf(Val(TxtTokenVal.Text) = 0, 0, Val(TxtTokenVal.Text))
                  Grid.Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  Grid.Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  Grid.Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  Grid.Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  Grid.Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  Grid.Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  Grid.Columns("isDiscB4TradeOffer").Value = Abs(ChkDiscB4TradeOffer.Value)
                  Grid.Columns("isDiscB4ExtraScheme").Value = Abs(ChkDiscB4ExtraScheme.Value)
                  Grid.Columns("isDiscB4SaleTax").Value = Abs(ChkDiscB4SaleTax.Value)
                  Grid.Columns("TradeOffer1").Value = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
                  Grid.Columns("TradeOffer2").Value = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
                  Grid.Columns("ExtraSchemePer").Value = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
                  Grid.Columns("TradeValue").Value = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
                  Grid.Columns("ExtraSchemeValue").Value = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
                  Grid.Columns("SC").Value = Val(TxtSC.Text)
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("ReSAmount").Value = Val(TxtReSAmount.Text)
                  Grid.Columns("Cost").Value = Val(TxtCost.Text)
                  Grid.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
                  Grid.Columns("ExpiryTime").Value = Val(vExpiryTime)
                  Grid.Columns("ExpiryDate").Value = TxtExpiryDate.Text
'                  If ObjRegistry.AllowEmployeProductWise Then
'                     RsBody!EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text))
'                  End If
'                  If ObjRegistry.AllowStoreProductWise Then
'                     RsBody!StoreID = IIf(Trim(TxtStoreID.Text) = "", Null, Val(TxtStoreID.Text))
'                  End If
'                  RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
'                  RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
'                  RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
'                  RsBody!Qty = Val(TxtQtyLoose.Text)
'                  RsBody!Bonus = Val(TxtBonus.Text)
'                  RsBody!Price = Val(TxtPrice.Text)
'                  RsBody!RetailPrice = Val(TxtRetailPrice.Text)
'                  RsBody!IsWSDiscb4ST = vIsWSDiscb4ST
'                  RsBody!IsWSSaleTax = vIsWSSaleTax
'                  RsBody!IsRetailSaleTax = vIsRetailSaleTax
'                  RsBody!TokenVal = IIf(Val(TxtTokenVal.Text) = 0, 0, Val(TxtTokenVal.Text))
'                  RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
'                  RsBody!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
'                  RsBody!SaleTaxval = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
'                  RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
'                  RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
'                  RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
'                  RsBody!Amount = Val(TxtAmount.Text)
'                  RsBody!Cost = Val(TxtCost.Text)
'                  RsBody!isProduct = 1 'Abs(Grid.Columns("isProduct").Value)

                  ssql = "Select Productid From salebody where sid=" & Val(TxtSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmSaleInvoiceDIS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmSaleInvoiceDIS, eEdit, TxtBillID.Text, DtpBillDate.DateValue, "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmSaleInvoiceDIS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Pending Update Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                  Grid.MoveLast
                
                  Call SubClearDetailArea
                  TxtCode.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vrowcounter
         
            Grid.Columns("Serial").Text = Grid.rows
            Grid.Columns("ProductID").Text = TxtProductID.Text
            Grid.Columns("Code").Text = TxtCode.Text
            Grid.Columns("Price").Value = Val(TxtPrice.Text)
            Grid.Columns("ReSPrice").Value = Val(TxtReSPrice.Text)
            Grid.Columns("BatchNo").Text = Trim(TxtBatchNo.Text)
            
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
''         SubClearDetailArea
         Grid.MoveLast
         TxtCode.SetFocus
         
'         Exit Sub
'      End If
   End If
   
   If ActiveControl.Name <> "Grid" Then Grid.Redraw = False
   
   With Grid
      If TxtCode.Enabled = True Then
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text)
         TxtSumDiscAmount.Text = Val(TxtSumDiscAmount.Text) + Val(TxtDiscAmount.Text)
         TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text)))
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmSaleInvoiceDIS, eAddNewRowByEdit, TxtBillID.Text, DtpBillDate.DateValue, "Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
         Call ActivityLogBin(vRandomID, eFrmSaleInvoiceDIS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Pending Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         TxtSumDiscAmount.Text = Val(TxtSumDiscAmount.Text) + Val(TxtDiscAmount.Text) - Val(.Columns("DiscAmount").Text)
         TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Val(Grid.Columns("Pack").Value)) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
         ssql = "Select Productid From salebody where sid=" & Val(TxtSID.Text) & " and billdate ='" & DtpBillDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
         With CN.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmSaleInvoiceDIS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmSaleInvoiceDIS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            Else
               Call ActivityLogBin("", eFrmSaleInvoiceDIS, eEdit, TxtBillID.Text, DtpBillDate.Date, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmSaleInvoiceDIS, eEdit, TxtBillID.Text, DtpBillDate.Date, "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmSaleInvoiceDIS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtBillID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpBillDate.Date), "Pending Update Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
      End If
      
      .Columns("BatchNo").Text = Trim(TxtBatchNo.Text)
      If ObjRegistry.AllowEmployeProductWise Then
         .Columns("EmpID").Text = TxtEmployeeID.Text
         .Columns("EmpName").Text = TxtEmployeeName.Text
      End If
      If ObjRegistry.AllowStoreProductWise Then
         .Columns("StoreID").Text = TxtStoreID.Text
         .Columns("StoreName").Text = TxtStoreName.Text
      End If
      .Columns("ProductID").Text = TxtProductID.Text
      .Columns("Code").Text = TxtCode.Text
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
      .Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
      .Columns("GrossQty").Value = IIf(Val(TxtGrossQty.Text) = 0, Null, Val(TxtGrossQty.Text))
      .Columns("GrossUnit").Value = IIf(Val(TxtGrossUnit.Text) = 0, Null, Val(TxtGrossUnit.Text))
      .Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
      .Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
      .Columns("Bonus").Value = Val(TxtBonus.Text)
      .Columns("Price").Value = Val(TxtPrice.Text)
      .Columns("ReSPrice").Value = Val(TxtReSPrice.Text)
      .Columns("isLastPrice").Value = Abs(IIf(Val(LblLastPrice.Caption) = Val(TxtPrice.Text), 1, 0))
      .Columns("RetailPrice").Value = Val(TxtRetailPrice.Text)
      .Columns("IsWSDiscb4ST").Value = vIsWSDiscb4ST
      .Columns("IsWSSaleTax").Value = vIsWSSaleTax
      .Columns("IsRetailSaleTax").Value = vIsRetailSaleTax
      .Columns("TokenVal").Value = IIf(Val(TxtTokenVal.Text) = 0, 0, Val(TxtTokenVal.Text))
      .Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      .Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
      .Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
      .Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      .Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      .Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      .Columns("isDiscB4TradeOffer").Value = Abs(ChkDiscB4TradeOffer.Value)
      .Columns("isDiscB4ExtraScheme").Value = Abs(ChkDiscB4ExtraScheme.Value)
      .Columns("isDiscB4SaleTax").Value = Abs(ChkDiscB4SaleTax.Value)
      .Columns("TradeOffer1").Value = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
      .Columns("TradeOffer2").Value = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
      .Columns("ExtraSchemePer").Value = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
      .Columns("TradeValue").Value = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
      .Columns("ExtraSchemeValue").Value = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
      .Columns("SC").Value = Val(TxtSC.Text)
      .Columns("DiscAmount").Value = Val(TxtDiscAmount.Text)
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("ReSAmount").Value = Val(TxtReSAmount.Text)
      .Columns("Cost").Value = Val(TxtCost.Text)
      .Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
      .Columns("ExpiryTime").Value = Val(vExpiryTime)
      .Columns("ExpiryDate").Value = TxtExpiryDate.Text
'      If ObjRegistry.AllowEmployeProductWise Then
'         RsBody!EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text))
'      End If
'      If ObjRegistry.AllowStoreProductWise Then
'         RsBody!StoreID = IIf(Trim(TxtStoreID.Text) = "", Null, Val(TxtStoreID.Text))
'      End If
'      RsBody!BatchNo = IIf(Trim(TxtBatchNo.Text) = "", Null, Trim(TxtBatchNo.Text))
'      RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
'      RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
'      RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
'      RsBody!Qty = Val(TxtQtyLoose.Text)
'      RsBody!Bonus = Val(TxtBonus.Text)
'      RsBody!Price = Val(TxtPrice.Text)
'      RsBody!RetailPrice = Val(TxtRetailPrice.Text)
'      RsBody!IsWSDiscb4ST = vIsWSDiscb4ST
'      RsBody!IsWSSaleTax = vIsWSSaleTax
'      RsBody!IsRetailSaleTax = vIsRetailSaleTax
'      RsBody!TokenVal = IIf(Val(TxtTokenVal.Text) = 0, 0, Val(TxtTokenVal.Text))
'      RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
'      RsBody!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
'      RsBody!SaleTaxval = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
'      RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
'      RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
'      RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
'      RsBody!Amount = Val(TxtAmount.Text)
'      RsBody!Cost = Val(TxtCost.Text)
'      RsBody!isProduct = 1 'Abs(Grid.Columns("isProduct").Value)
        
      If TxtCode.Enabled = False And ObjRegistry.AfterRowEditFocusNextGridLine = True Then
         
         Grid.MoveNext
         Call Grid_GotFocus
         Call GetDataBackFromGridToTexBoxes
'         CmbPackName.SetFocus
      Else
        
         Grid.MoveLast
         If Trim(.Columns("Code").Text) <> "" Then
         Grid.AllowAddNew = True
         Grid.AddNew
         Grid.Columns("Code").Text = " "
         Grid.AllowAddNew = False
         
      End If
   End If
   
'      .MoveLast
   End With
   
   QtyOffer = 0
   
   GetDataFromTextBoxesToGridOffer
   Call SubClearDetailArea
   
   Grid.Redraw = True
   FrmExpiry.Visible = False
   TxtTotalItems.Text = Val(Grid.rows) - 1
   
   GetDataBackFromGridToTexBoxes
   Grid_LostFocus

   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtCode.Enabled = True
   BtnProduct.Enabled = True
   TxtCode.Text = ""
   TxtBatchNo.Text = ""
   TxtExpiryDate.Text = ""
   TxtProductName.Text = ""
   CmbPackName.ListIndex = 0
   TxtMultiplier.Text = ""
   TxtGrossQty.Text = ""
   TxtGrossUnit.Text = ""
   TxtQtyPack.Text = ""
   TxtQtyLoose.Text = ""
   TxtBonus.Text = ""
   TxtPrice.Text = ""
   TxtReSPrice.Text = ""
   TxtRetailPrice.Text = ""
   TxtTokenVal.Text = ""
   TxtOffer.Text = ""
   TxtSaleTaxPer.Text = ""
   TxtSaleTaxVal.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtDiscVal.Text = ""
   TxtSC.Text = ""
   TxtDiscAmount.Text = ""
   TxtAmount.Text = ""
   TxtReSAmount.Text = ""
   TxtTradeOffer1.Text = ""
   TxtTradeOffer2.Text = ""
   TxtExtraSchemePer.Text = ""
   TxtTradeOfferValue.Text = ""
   TxtExtraSchemeValue.Text = ""
   ChkDiscB4TradeOffer.Value = 0
   ChkDiscB4ExtraScheme.Value = 0
   ChkDiscB4SaleTax.Value = 0
   TxtQtyLoose.Enabled = True
End Sub

Private Sub GetDataFromTextBoxesToGridOffer()
On Error GoTo ErrorHandler

    With CN.Execute("Select * from ProductOffers where Rebate = 0 and ProductID = " & Val(TxtProductID.Text))
        
        If .RecordCount > 0 Then
            QtyOffer = QtyOffer + Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)
            QtyOffer = QtyOffer \ IIf(!Qty * !QtyOffer = 0, 1, !Qty * !QtyOffer)
            If QtyOffer > 0 Then
                
                RsProductOffer.Filter = "ProductID = " & Val(TxtProductID.Text)
                If TxtProductID.Enabled Then
                    If RsProductOffer.RecordCount = 0 Then
'                        RsProductOffer.AddNew
                        GridOffer.Columns("ProductID").Text = TxtProductID.Text
                        GridOffer.Columns("ProductOfferID").Text = IIf(IsNull(!ProductOfferID), "", !ProductOfferID)
'                        RsProductOffer!Productid = TxtProductID.Text
'                        RsProductOffer!ProductOfferID = !ProductOfferID
                    Else
                        GridOffer.Redraw = False
                        GridOffer.MoveFirst
                        For vCounter = 1 To GridOffer.rows
                        If GridOffer.Columns("ProductID").Text = TxtProductID.Text Then
                            GridOffer.Columns("ProductName").Text = CN.Execute("Select ProductName from products where productid = '" & GridOffer.Columns("ProductOfferID").Text & "'").Fields(0)
                            GridOffer.Columns("Qty").Value = QtyOffer
'                            RsProductOffer!QtyOffer = QtyOffer
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
'                        RsProductOffer.AddNew
                        GridOffer.Columns("ProductID").Text = TxtProductID.Text
                        GridOffer.Columns("ProductOfferID").Text = IIf(IsNull(!ProductOfferID), "", !ProductOfferID)
'                        RsProductOffer!Productid = TxtProductID.Text
'                        RsProductOffer!ProductOfferID = !ProductOfferID
                End If
                    GridOffer.Columns("ProductName").Text = CN.Execute("Select ProductName from products where productid = " & Val(GridOffer.Columns("ProductOfferID").Text)).Fields(0)
                    GridOffer.Columns("Qty").Value = QtyOffer
'                    RsProductOffer!QtyOffer = QtyOffer
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
                For vCounter = 1 To GridOffer.rows
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
    
'    If GridOffer.Rows <> 0 Then GridOffer.Visible = True
    
   Exit Sub
ErrorHandler:
   GridOffer.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGridOffer()
    If RsProductOffer.State = adStateOpen Then RsProductOffer.Close
    RsProductOffer.Open "Select * from SaleBodyOffer where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
    If RsProductOffer.RecordCount > 0 Then
    GridOffer.Visible = True
    ssql = "select p.productname, D.* from SaleBodyOffer D Inner join products p on p.productid = D.productOfferid where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'"
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

Private Sub PopulateSaleOrderToGridOffer()
    If RsProductOffer.State = adStateOpen Then RsProductOffer.Close
    RsProductOffer.Open "Select * from SaleBodyOffer where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
'    If RsProductOffer.RecordCount > 0 Then
    
    ssql = "select p.productname, D.* from SaleOrderBodyOffer D Inner join products p on p.productid = D.productOfferid where OrderID =" & Val(TxtOrderID.Text) & " And OrderDate = '" & DtpOrderDate.DateValue & "'"
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

Private Sub PopulateDataToGridExpense()
    If RsExpense.State = adStateOpen Then RsExpense.Close
    RsExpense.Open "Select * from SaleExpense where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
'    GridExpense.Visible = True
    ssql = "select EA.AccountNo, Accountname, SE.ExpAmount from ExpenseAccounts EA Left Outer join ChartofAccounts C on C.AccountNo = EA.AccountNo Left Outer Join (Select * from SaleExpense where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "') SE On SE.ExpenseID = EA.AccountNo"
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
            .MoveNext
         Wend
      End With

     If GridExpense.rows > 0 Then GridExpense.FirstRow = 0
     GridExpense.Redraw = True
'      GridExpense.Visible = False
End Sub

Private Sub PopulateSaleOrderToGridExpense()
    If RsExpense.State = adStateOpen Then RsExpense.Close
    RsExpense.Open "Select * from SaleExpense where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
'    GridExpense.Visible = True
    ssql = "select EA.AccountNo, Accountname, SE.ExpAmount from ExpenseAccounts EA Left Outer join ChartofAccounts C on C.AccountNo = EA.AccountNo Left Outer Join (Select * from SaleOrderExpense where OrderID =" & Val(TxtOrderID.Text) & " And OrderDate = '" & DtpOrderDate.DateValue & "') SE On SE.ExpenseID = EA.AccountNo"
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
            RsExpense.AddNew
            RsExpense!ExpenseID = !AccountNo
            RsExpense!ExpAmount = IIf(IsNull(!ExpAmount), 0, !ExpAmount)
            RsExpense.Update
            GridExpense.Update
            .MoveNext
         Wend
      End With

     If GridExpense.rows > 0 Then GridExpense.FirstRow = 0
     GridExpense.Redraw = True
'      GridExpense.Visible = False
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("Code").Text
      TxtBatchNo.Text = .Columns("BatchNo").Text
      TxtExpiryDate.Text = .Columns("ExpiryDate").Text
      If ObjRegistry.AllowEmployeProductWise Then
         TxtEmployeeID.Text = .Columns("EmpID").Text
         TxtEmployeeName.Text = .Columns("EmpName").Text
      End If
      If ObjRegistry.AllowStoreProductWise And (.Columns("StoreID").Text <> "") Then
         TxtStoreID.Text = .Columns("StoreID").Text
         TxtStoreName.Text = .Columns("StoreName").Text
      End If
      TxtProductName.Text = .Columns("ProductName").Text
      
      CmbPackName.Clear
      vStrSQL = "select * from ProductPacking pp inner join packings p on p.packingid = pp.packingid" & vbCrLf _
           + "left outer join ProductBarcodes b on b.productid = pp.productid" & vbCrLf _
           + " where pp.productid = " & Val(TxtCode.Text) & " or code='" & TxtCode.Text & "'"
      With CN.Execute(vStrSQL)
         CmbPackName.AddItem ""
         While Not .EOF
            CmbPackName.AddItem !Packingname
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
      TxtQtyLoose.Text = .Columns("QtyLoose").Text
      TxtQtyPack.Text = .Columns("QtyPack").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtReSPrice.Text = .Columns("ReSPrice").Text
      TxtRetailPrice.Text = .Columns("RetailPrice").Text
      vIsWSDiscb4ST = .Columns("IsWSDiscb4ST").Value
      vIsRetailSaleTax = .Columns("IsRetailSaleTax").Value
      vIsRetailSaleTax = .Columns("IsRetailSaleTax").Value
      TxtTokenVal.Text = .Columns("TokenVal").Value
      TxtBonus.Text = .Columns("Bonus").Text
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtOffer.Text = .Columns("Offer").Value
      TxtSaleTaxPer.Text = .Columns("SaleTaxPer").Value
      TxtSaleTaxVal.Text = .Columns("SaleTaxVal").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      ChkDiscB4TradeOffer.Value = Abs(Val(.Columns("isDiscB4TradeOffer").Value))
      ChkDiscB4ExtraScheme.Value = Abs(Val(.Columns("isDiscB4ExtraScheme").Value))
      ChkDiscB4SaleTax.Value = Abs(Val(.Columns("isDiscB4SaleTax").Value))
      TxtTradeOffer1.Text = .Columns("TradeOffer1").Value
      TxtTradeOffer2.Text = .Columns("TradeOffer2").Value
      TxtExtraSchemePer.Text = .Columns("ExtraSchemePer").Value
      TxtTradeOfferValue.Text = .Columns("TradeValue").Value
      TxtExtraSchemeValue.Text = .Columns("ExtraSchemeValue").Value
      TxtSC.Text = .Columns("SC").Value
      TxtDiscAmount.Text = .Columns("DiscAmount").Value
      TxtAmount.Text = .Columns("Amount").Value
      TxtReSAmount.Text = .Columns("ReSAmount").Value
      
      TxtCost.Text = .Columns("Cost").Value
      ChkIsProduct.Value = Abs(.Columns("isProduct").Value)
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
        If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & Val(TxtProductID.Text)
            With CN.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         End If
'         With cn.Execute("select QtyLoose from currentstockStore where productid ='" & TxtProductID.Text & "' and storeid = " & TxtStoreID.Text)
'         If .RecordCount > 0 Then
'            vQtyLoose = !QtyLoose
'            LblStock.Caption = !QtyLoose & " " & cn.Execute("SELECT dbo.FunGetUnit('" & TxtProductID.Text & "')").Fields(0).Value
'         Else
'            vQtyLoose = 0
'            LblStock.Caption = 0
'         End If
'      End With
         If TxtProductID.Enabled = False Then
            LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
            LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
   '         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
            LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
            LblStock.Caption = LblStock.Caption & " " & "Loose"
         End If
'      With cn.Execute("select QtyLoose from currentstockStore where productid ='" & TxtProductID.Text & "' and storeid = " & TxtStoreID.Text)
'         If .RecordCount > 0 Then
''            vQtyLoose = !QtyLoose
'            LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtProductID.Text & "',Floor(" & !QtyLoose & "))").Fields(0).Value
'            LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'            LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & !QtyLoose & "))").Fields(0).Value
'            LblStock.Caption = LblStock.Caption & " " & "Loose"
'            'LblStock.Caption = !QtyLoose & " " & cn.Execute("SELECT dbo.FunGetUnit('" & TxtProductID.Text & "')").Fields(0).Value
'         Else
''            vQtyLoose = 0
'            LblStock.Caption = 0
'         End If
'      End With
      
'      VStrSQL = "select isnull(dbo.FunStock('" & TxtProductID.Text & "'," & TxtStoreID.Text & "," & Val(TxtBillID.Text) & "," & Val(0) & "," & Val(TxtBillID.Text) & "," & Val(0) & "," & Val(0) & "," & Val(0) & ",'" & DateAdd("D", 1, DtpBillDate.DateValue) & "'," & Val(0) & "),0)"
'      vQtyLoose = cn.Execute(VStrSQL).Fields(0).Value
'      LblStock.Caption = vQtyLoose
         
      vUnitPrice = Val(.Columns("Price").Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))
      vUnitRetailPrice = Val(.Columns("RetailPrice").Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))
      If Trim(TxtProductID.Text) <> "" Then
         LblRetailPrice.Caption = CN.Execute("Select RetailPrice from Products where ProductID = " & Val(TxtProductID.Text)).Fields(0).Value
         LblPurPrice.Caption = CN.Execute("Select PurPrice from Products where ProductID = " & Val(TxtProductID.Text)).Fields(0).Value
         LblLastPurPrice.Caption = CN.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text) & ")").Fields(0).Value
         LblLastPrice.Caption = CN.Execute("Select dbo.FunLastPrice('S','" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text) & ",'" & TxtCustomerID.Text & "')").Fields(0).Value
      End If
   End With
   If Grid.rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetSale()
   On Error GoTo ErrorHandler
   ssql = "select h.*, EmpName, OrganizationName, isnull(p.partyname,c.accountname) as PartyName, SyllabusName, isnull(p.isWholeSale,1) as isWholeSale, " & vbCrLf _
         + " P.Description as CustomerDescription, P.Address, P.City, isnull(p.phone1 + ',','') + isnull(p.phone2 + ',','') + isnull(p.mobile + ',','')" & vbCrLf _
         + " + isnull(p.mobile2 + ',','') as Phone, c.AccountName, BankMachineName, StoreName, EmpName, MemberName, LicenceNO FROM SaleHeader h left outer join parties p on h.CustomerID = p.partyid left outer join Organizations o on o.OrganizationID = h.OrganizationID inner join ChartofAccounts c on h.customerid = c.AccountNo left outer join BankMachines b on b.BankMachineid = h.BankMachineid inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID left outer join Members M on M.MemberID = h.memberID left outer join SyllabusHeader syl on syl.syllabusID = h.syllabusID where isReplace=0 and h.sID=" & Val(TxtSID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'" & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With CN.Execute(ssql)
      If Not .BOF Then
          DtpBillDate.DateValue = !BillDate
          DtpPromiseDate.DateValue = !PromiseDate
          DtpDispatchDate.DateValue = !DispatchDate
          DtpExpiryInvoice.DateValue = !ExpiryInvoice
          If (IsNull(!terms)) Then
            TxtTerms.Text = 0
          Else
            TxtTerms.Text = !terms
          End If
          TxtOrderID.Text = IIf(IsNull(!OrderID), "", !OrderID)
          DtpOrderDate.DateValue = IIf(IsNull(!OrderDate), "01/01/1990", !OrderDate)
          TxtCustomerID.Text = Val(!CustomerID)
          TxtCustomerName.Text = !partyname
          TxtRefID.Text = IIf(IsNull(!RefID), "", !RefID)
          TxtRefComm.Text = IIf(IsNull(!RefComm), "", !RefComm)
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtContactNo.Text = IIf(IsNull(!Phone), "", !Phone)
          TxtLicenceNO.Text = IIf(IsNull(!LicenceNO), "", !LicenceNO)
          TxtMemberID.Text = IIf(IsNull(!MemberID), "", !MemberID)
          TxtMemberName.Text = IIf(IsNull(!MemberName), "", !MemberName)
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
          TxtEmployeeName.Text = IIf(IsNull(!empname), "", !empname)
          TxtBillNo.Text = IIf(IsNull(!BillNo), "", !BillNo)
          TxtBiltyNo.Text = IIf(IsNull(!BiltyNo), "", !BiltyNo)
          TxtVehicleNo.Text = IIf(IsNull(!VehicleNo), "", !VehicleNo)
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          LblCustomerDesc.Caption = IIf(IsNull(!CustomerDescription), "", !CustomerDescription)
          TxtTotalAmount.Text = !TotalAmount
          TxtSumDiscAmount.Text = !SumDiscAmount
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          TxtServiceChargesPer.Text = IIf(IsNull(!ServiceChargesPer), "", !ServiceChargesPer)
          TxtServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtAdvTaxVal.Text = IIf(IsNull(!AdvTaxVal), "", !AdvTaxVal)
          TxtAdvTaxPer.Text = IIf(IsNull(!AdvTaxPer), "", !AdvTaxPer)
          TxtExtraTaxVal.Text = IIf(IsNull(!ExtraTaxVal), "", !ExtraTaxVal)
          TxtExtraTaxPer.Text = IIf(IsNull(!ExtraTaxPer), "", !ExtraTaxPer)
          TxtCNIC.Text = IIf(IsNull(!CNIC), "", !CNIC)
          TxtCellNo.Text = IIf(IsNull(!MobileNo), "", !MobileNo)
          TxtCashCustomer.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
          TxtTotalExpense.Text = IIf(IsNull(!TotalExpense), "", !TotalExpense)
          TxtFreight.Text = IIf(IsNull(!Freight), "", !Freight)
'          TxtPaidAmount.Text = IIf(IsNull(!PAIDAMOUNT), "", !PAIDAMOUNT)
          TxtReceivedAmount.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
          TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
'          TxtPreviousReceivable.Text = IIf(IsNull(!PreviousAmount), "", !PreviousAmount)
          TxtPreviousReceivable.Text = CN.Execute("SELECT isnull(dbo.FunCurrentDebit(" & Val(TxtCustomerID.Text) & ",'" & DtpBillDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
           vStrSQL = " Select isnull(Sum(round(B.TTLValue,0) - isnull(BillDisc,0) + isnull(OtherCharges,0) + Isnull(TotalExpense,0) + isnull(servicecharges,0) + isnull(STax,0)),0) as Amount " & vbCrLf _
                  + " FROM SaleHeader h INNER JOIN (Select SID, Sum(Amount) TTLValue FROM SaleBody Group By SID)b " & vbCrLf _
                  + " ON H.SID = B.SID " & vbCrLf _
                  + " where CustomerID = " & Val(TxtCustomerID.Text) & " and h.BillDate = '" & DtpBillDate.DateValue & "' and h.BillID >= " & Val(TxtBillID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
          TxtPreviousReceivable.Text = TxtPreviousReceivable.Text - CN.Execute(vStrSQL).Fields(0).Value
          lblPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
          LblTtlPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Total Receivable", "Total Payable")
          TxtPreviousReceivable.Text = Abs(Val(TxtPreviousReceivable.Text))
          vZoneID = CN.Execute("SELECT isnull(dbo.FunGetZoneID(" & TxtCustomerID.Text & "),0)").Fields(0).Value
          TxtSyllabusID.Text = IIf(IsNull(!syllabusid), "", !syllabusid)
          TxtSyllabusName.Text = IIf(IsNull(!SyllabusName), "", !SyllabusName)
          isWholeSale = !isWholeSale
         If Val(TxtCustomerID.Text) = 621 Then
            TxtPreviousReceivable.Text = ""
'            TxtTotalReceivable.Text = ""
         End If
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

Private Sub GetSaleOrder()
   On Error GoTo ErrorHandler
   TxtBillID.Text = FunGetMaxID
   ssql = "select h.*, EmpName, OrganizationName, p.partyname, P.Address, P.City, c.AccountName, BankMachineName, StoreName, EmpName, MemberName FROM SaleOrderHeader h join parties p on h.CustomerID = p.partyid left outer join Organizations o on o.OrganizationID = h.OrganizationID left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join BankMachines b on b.BankMachineid = h.BankMachineid inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID left outer join Members M on M.MemberID = h.memberID where isReplace=0 and h.orderID=" & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'"
   With CN.Execute(ssql)
      If Not .BOF Then
          DtpOrderDate.DateValue = !OrderDate
          TxtCustomerID.Text = Val(!CustomerID)
          TxtCustomerName.Text = !partyname
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
          TxtEmployeeName.Text = IIf(IsNull(!empname), "", !empname)
          TxtBillNo.Text = IIf(IsNull(!BillNo), "", !BillNo)
          TxtBiltyNo.Text = IIf(IsNull(!BiltyNo), "", !BiltyNo)
          TxtVehicleNo.Text = IIf(IsNull(!VehicleNo), "", !VehicleNo)
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtTotalAmount.Text = !TotalAmount
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtTotalExpense.Text = IIf(IsNull(!TotalExpense), "", !TotalExpense)
'          TxtPaidAmount.Text = IIf(IsNull(!PAIDAMOUNT), "", !PAIDAMOUNT)
          TxtReceivedAmount.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtPreviousReceivable.Text = IIf(IsNull(!PreviousAmount), "", !PreviousAmount)
          lblPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
          LblTtlPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Total Receivable", "Total Payable")
          TxtPreviousReceivable.Text = Abs(Val(TxtPreviousReceivable.Text))

      End If
      .Close
   End With
   Call PopulateSaleOrderToGrid
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
        TxtSerial.Enabled = False
    End If
End Sub

Private Sub GridSerial_GotFocus()
   If Grid.Columns("Code").Text <> " " And GridSerial.Columns("Serial").Text = " " Then
        TxtSerial.Enabled = True
    Else
        TxtSerial.Enabled = False
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
   If ActiveControl.Name <> TxtBillDisc.Name Then Exit Sub
   TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / IIf(Val(TxtTotalAmount.Text) = 0, 1, Val(TxtTotalAmount.Text)), 2)
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

Private Sub TxtDiscVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDiscVal.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   If (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) = 0 Then Exit Sub
   TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 4)
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 2)
   Call CalculateValue
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDiscVal_LostFocus()
   If ObjRegistry.ChangeQtyOnChangedPrice = True Then Exit Sub
   Select Case ActiveControl.Name
   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtBonus.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtOffer.Name, TxtSaleTaxPer.Name, TxtSC.Name
      Exit Sub
   End Select
'    following function is already call from Form Key Down so dont need to call again
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
'    If ActiveControl.Name <> TxtOffer.Name Then Exit Sub
    Call SubCalculateBody
End Sub

Private Sub TxtOtherCharges_Change()
   On Error GoTo ErrorHandler
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtOtherCharges_GotFocus()
   On Error GoTo ErrorHandler
   If ObjRegistry.PackingChargesPer = "0" Then Exit Sub
   TxtOtherCharges.Text = Round(TxtNetAmount.Text * ObjRegistry.PackingChargesPer / 100, 0)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
   If ActiveControl.Name <> TxtQtyLoose.Name Then Exit Sub
   Call FindRebate
   Call SubCalculateBody
End Sub

Private Sub TxtQtyLoose_Validate(Cancel As Boolean)
   If ActiveControl.Name <> TxtQtyLoose.Name Then Exit Sub
End Sub

Private Sub TxtQtyPack_Change()
   Call FindRebate
   Call SubCalculateBody
End Sub

Private Sub TxtQtyPack_LostFocus()
 If ObjRegistry.EitherPackORLooseEnter = True And Val(TxtQtyPack.Text) > 0 Then
    TxtQtyLoose.Text = ""
    TxtQtyLoose.Enabled = False
 End If
End Sub

Private Sub TxtReSPrice_Change()
If ActiveControl.Name <> TxtReSPrice.Name Then Exit Sub
TxtReSAmount.Text = Round((Val(TxtReSPrice.Text) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), 2)
End Sub

Private Sub TxtSaleTaxPer_Change()
   If ActiveControl.Name <> TxtSaleTaxPer.Name Then Exit Sub
   Call SubCalculateBody
End Sub

Private Sub TxtSaleTaxPer_LostFocus()
   Select Case ActiveControl.Name
   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtBonus.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtOffer.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
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
Select Case ActiveControl.Name
   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtBonus.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtOffer.Name, TxtSaleTaxPer.Name, TxtSC.Name
      Exit Sub
   End Select
End Sub

Private Sub TxtSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridSerial.SetFocus
End Sub

Private Sub TxtSerial_LostFocus()
'   GetDataFromTexBoxesToGridSerial
End Sub

Private Sub GetDataFromTexBoxesToGridSerial()
   On Error GoTo ErrorHandler
'   Dim vrowcounter As Integer
'
'   If Trim(TxtSerial.Text) = "" Then
'      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
''      TxtSerial.SetFocus
'      Exit Sub
'   End If
''   RsBodySerial.Filter = "ProductID ='" & Grid.Columns("ProductID").Text & "' And Serial='" & TxtSerial.Text & "'"
'
'         GridSerial.Redraw = False
'         GridSerial.MoveFirst
'            For vrowcounter = 1 To GridSerial.rows
'               If GridSerial.Columns("Serial").Text = TxtSerial.Text Then
'                  MsgBox "The Product cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
'                  vAlreadySerial = True
'                  'SubClearDetailArea
'                  GridSerial.MoveLast
'                  TxtSerial.SetFocus
'                  GridSerial.Redraw = True
'                  Exit Sub
'               End If
'               GridSerial.MoveNext
'            Next vrowcounter
'         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
'
'  If TxtSerial.Enabled Then
''         RsBodySerial.AddNew
'         GridSerial.Columns("ProductID").Text = TxtCode.Text
'         GridSerial.Columns("Serial").Text = TxtSerial.Text
''         RsBodySerial!Productid = TxtCode.Text
''         RsBodySerial!Serial = TxtSerial.Text
'         TxtSerial.Text = ""
'  End If
'   'GridSerial.Redraw = False
'   With GridSerial
'      If Trim(.Columns("Serial").Text) <> "" Then
'         .AllowAddNew = True
'         .AddNew
'         .Columns("Serial").Text = " "
'         .AllowAddNew = False
'      End If
'   End With
'   If TxtSerial.Visible = True Then TxtSerial.SetFocus
'   GridSerial.Redraw = True
   Exit Sub
ErrorHandler:
   GridSerial.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtServiceCharges_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtServiceCharges.Name Then Exit Sub
   TxtServiceChargesPer.Text = Round((Val(TxtServiceCharges.Text) * 100) / IIf(Val(TxtTotalAmount.Text) = 0, 1, Val(TxtTotalAmount.Text)), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtServiceChargesPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtServiceChargesPer.Name Then Exit Sub
   TxtServiceCharges.Text = SelfRound((Val(TxtTotalAmount.Text) * Val(TxtServiceChargesPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSyllabusID_Change()
    If TxtSyllabusID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSyllabusID.Name Then Exit Sub
   If TxtSyllabusName.Text <> "" Then TxtSyllabusName.Text = ""
End Sub

Private Sub TxtSyllabusID_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler
    If TxtSyllabusName.Text <> "" Then Exit Sub
    If TxtSyllabusID.Text = "" Then Exit Sub
    Dim vTemp As Boolean
    vTemp = Not FunSelectSyllabus(ssValidate, True)
    If vTemp = True Then
        vTemp = Not FunSelectSyllabus(ssButton, False)
    End If
    Cancel = vTemp
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub TxtTerms_Change()
On Error GoTo ErrorHandler
   If TxtTerms.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtTerms.Name Then Exit Sub
   DtpPromiseDate.DateValue = DateAdd("d", Val(TxtTerms.Text), DtpDispatchDate.DateValue)
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub TxtTotalAmount_Change()
   'TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Text) * Val(TxtBillDiscPer.Text) / 100))
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
   If KeyCode = vbKeyDown Then
    Grid.Enabled = True
    Grid.Redraw = True
    Grid.Visible = True
    Grid.SetFocus
  End If
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
    With CN.Execute("Select * from ProductOffers where ProductID = " & Val(TxtProductID.Text))
        .Filter = "Rebate <> 0"
        If .RecordCount > 0 Then
            Rebate = Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)
            If !FixedRebate Then
               Rebate = IIf(Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text) >= !Qty, 1, 0)
            Else
               Rebate = Rebate \ !Qty
            End If
            
            Rebate = Rebate * !Rebate
            TxtOffer.Text = Rebate
'            TxtDiscVal.Text = Rebate
'            If Val(TxtPrice.Text) = 0 Then Exit Sub
'            If Val(TxtQty.Text) = 0 Then Exit Sub
'            TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (TxtQty.Text), 3)
'            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
'            TxtActualAmount.Text = Val(TxtQty.Text) * Val(TxtPrice.Text)
'            TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
'            TxtTotalDiscount.Caption = vTotDisc
'            SubCalculateBody
        End If
        .Filter = "QtyOffer <> 0"
        If .RecordCount > 0 Then
            Rebate = Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)
            Rebate = Rebate \ !Qty
            TxtBonus.Text = Rebate * !QtyOffer
        End If
    End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub UserActivities()
     If vIsNewRecord = False Then
    With CN.Execute("Select  * from SaleHeader where BillID =" & TxtBillID.Text & " And BillDate = '" & DtpBillDate.DateValue & "'")
        
        If TxtStoreID.Text <> !StoreID Then
            CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.rows - 1
        With CN.Execute("Select * from SaleBody Where BillID = " & TxtBillID.Text & " and BillDate ='" & DtpBillDate.DateValue & "' and Productid = " & Val(Grid.Columns("Productid").Text))
        
             If .EOF = True Then
                ssql = "Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Inserted New Code-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")"
                CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Inserted New Code-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("QtyLoose").Text <> !Qty Or Grid.Columns("Price").Text <> !Price Or Grid.Columns("discper").Text <> !DiscPer Then
                   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Updated Code-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
    
   Else
    CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub TxtTotalExpense_Change()
   Call SubCalculateFooter
End Sub
Private Sub ActivityLogSale(FormType As String, Mode As EntryMode, Optional Key1 As Long = 0, Optional Key2 As Date = "01-01-1900", Optional Key3 As String = "")
   Dim vSQL As String
   vSQL = "Exec ProdActivityLog '" & FormType & "'," & ObjUserSecurity.UserNo & "," & Mode & "," & Key1 & ",'" & Key2 & "','" & Key3 & "'"
   'vSQL = "INSERT into ActivityLogSaleSale(userno,FormType,EntryDate,Description,isnew,isedit,isdelete) values(" & ObjUserSecurity.UserNo & ",'" & FormType & "',getdate(),'" & Desc & "'," & IIf(Mode = eAdd, 1, 0) & "," & IIf(Mode = eEdit, 1, 0) & "," & IIf(Mode = eDelete, 1, 0) & ")"
   CN.Execute vSQL
End Sub
Private Sub MniCostPrice_Click()
   On Error GoTo ErrorHandler
'   If Trim(Grid.Columns("Cost").Text) = "" Then Exit Sub
'   If ObjUserSecurity.ShowPurchasePriceInInvoice = True Or ObjUserSecurity.IsAdministrator = True Then
'      LblCost.Caption = Grid.Columns("Cost").Value
      LblCost.Visible = True
'   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGridExpiry()
      ssql = "select BatchNo, ExpiryDate " & vbCrLf & _
      " from PurchaseHeader h inner join Purchasebody b on h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate" & vbCrLf & _
      " where BatchNo is not null and b.productid = " & Val(TxtProductID.Text) & " order by b.PurchaseDate Desc"
      
      With CN.Execute(ssql)
         GridExpiry.Redraw = False
         GridExpiry.MoveFirst
         GridExpiry.RemoveAll
         GridExpiry.AllowAddNew = True
         While Not .EOF
            GridExpiry.Columns("BatchNo").Text = !BatchNo
            GridExpiry.Columns("ExpiryDate").Text = IIf(IsNull(!ExpiryDate), Date, !ExpiryDate)
            GridExpiry.AddNew
            .MoveNext
         Wend
         .Close
         GridExpiry.MoveFirst
         TxtBatchNo.Text = GridExpiry.Columns("BatchNo").Text
         TxtExpiryDate.Text = GridExpiry.Columns("ExpiryDate").Text
         GridExpiry.Redraw = True
      End With
End Sub



Private Function FunSelectSyllabus(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        FrmSyllabusSelection.Show vbModal, Me
        If FrmSyllabusSelection.ParaOutID = "" Then FunSelectSyllabus = False: Exit Function
        TxtSyllabusID.Text = FrmSyllabusSelection.ParaOutID
    End If
    '---------------------------

    vStrSQL = " Select * FROM syllabusheader where SyllabusID=" & Val(TxtSyllabusID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSyllabusName.Text = !SyllabusName
          FunSelectSyllabus = True
          .Close
          GetSyllabus
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSyllabus = False
          .Close
          TxtSyllabusID.Text = ""
          TxtSyllabusName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub GetSyllabus()
   On Error GoTo ErrorHandler
   ssql = "select h.* from SyllabusHeader h Where h.SyllabusID=" & Val(TxtSyllabusID.Text)
   With CN.Execute(ssql)
      If Not .BOF Then
          ' Ok
      End If
      .Close
   End With
   Call PopulateSyllabusToGrid
'   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub
Private Sub PopulateSyllabusToGrid()
      ssql = " select b.ProductID, b.code, ProductName, 0 as packingID, Null as Multiplier, Null QtyPack, Null Bonus, Null Cost, " & vbCrLf _
             + " RetailPrice, IsWSDiscb4ST, IsWSSaleTax,  TokenVal, Null DiscPC, Null Offer, SaleTaxPer, Null SaleTaxval, " & vbCrLf _
             + " Price, QtyLoose Qty, Null DiscPer, 0 DiscPC, Null DiscVal, Amount From syllabusBody b left outer join products p on p.productid = b.productid where syllabusid =" & TxtSyllabusID.Text & " and isShow = 1"
       With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = !Code
            Grid.Columns("ProductName").Text = !ProductName
            
           
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
            Grid.Columns("QtyLoose").Value = !Qty
            Grid.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            Grid.Columns("Price").Value = !Price
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            Grid.Columns("isProduct").Value = 1 '!isProduct
            Grid.Columns("RetailPrice").Value = !RetailPrice
            Grid.Columns("IsWSDiscb4ST").Value = !IsWSDiscb4ST
            Grid.Columns("IsWSSaleTax").Value = !IsWSSaleTax
            Grid.Columns("IsRetailSaleTax").Value = False
            Grid.Columns("TokenVal").Value = IIf(IsNull(!TokenVal), "", !TokenVal)
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !Qty) 'IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = ((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !Qty) '!Amount
            Grid.Columns("Amount").Value = !Amount
'            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!RQtyPack), 0, !RQtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !RQty))
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + !Amount
'            TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + !RQty + IIf(IsNull(!RBonus), "0", !RBonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!RQtyPack), 0, !RQtyPack))
            TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + !Qty
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
'   End If
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
If Val(TxtSaleTaxPer.Text) <> 0 Then
   If ChkDiscB4SaleTax.Value = 1 Then
      TxtSaleTaxVal.Text = (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)))
      TxtSaleTaxVal.Text = Val(TxtSaleTaxVal.Text) * Val(TxtSaleTaxPer.Text) / 100
   Else
      TxtSaleTaxVal.Text = (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)))
      TxtSaleTaxVal.Text = Val(TxtSaleTaxVal.Text) - Val(TxtDiscVal.Text) - Val(TxtTradeOfferValue.Text) - Val(TxtExtraSchemeValue.Text)
      TxtSaleTaxVal.Text = Round(Val(TxtSaleTaxVal.Text) * Val(TxtSaleTaxPer.Text) / 100, 4)
   End If
Else
   TxtSaleTaxPer.Text = 0
End If
'''''''''''''''''''''''''''''''''

TxtAmount.Text = (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)))
TxtAmount.Text = Round(Val(TxtAmount.Text) - Val(TxtDiscVal.Text) - Val(TxtTradeOfferValue.Text) - Val(TxtExtraSchemeValue.Text) + Val(TxtSaleTaxVal.Text) + Val(TxtSC.Text) - Val(TxtOffer.Text), 2)
If ObjRegistry.IsRoundFigure = True Then TxtAmount.Text = SelfRound(TxtAmount.Text)
TxtDiscAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - Val(TxtDiscVal.Text) - Val(TxtExtraSchemeValue.Text), 2)
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

Private Sub TxtAdvTaxPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtAdvTaxPer.Name Then Exit Sub
   vAdvDiscPerFlag = True
   TxtNetAmount.Text = Round(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text), 2) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text) + Val(TxtServiceCharges.Text)
   TxtAdvTaxVal.Text = SelfRound((Val(TxtNetAmount.Text) * Val(TxtAdvTaxPer.Text) / 100))
'   TxtAdvTaxVal.Text = SelfRound((Val(TxtSumDiscAmount.Text) * Val(TxtAdvTaxPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtAdvTaxVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtAdvTaxVal.Name Then Exit Sub
   vAdvDiscPerFlag = False
   TxtNetAmount.Text = Round(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text), 2) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text) + Val(TxtServiceCharges.Text)
   TxtAdvTaxPer.Text = Round((Val(TxtAdvTaxVal.Text) * 100) / IIf(Val(TxtNetAmount.Text) = 0, 1, Val(TxtNetAmount.Text)), 2)
'   TxtAdvTaxPer.Text = Round((Val(TxtAdvTaxVal.Text) * 100) / IIf(Val(TxtSumDiscAmount.Text) = 0, 1, Val(TxtSumDiscAmount.Text)), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridMultipleStore_DblClick()
   
      If Trim(CmbMSPackName.Text) = "" Then
         TxtMSQtyLoose.SetFocus
      Else
         TxtMSQtyPack.SetFocus
      End If
  
End Sub

Private Sub GridMultipleStore_GotFocus()
   Flag = True
   If Trim(Grid.Columns("productID").Text) <> "" Then
      CmbMSStore.Enabled = False
    End If
End Sub

Private Sub GridMultipleStore_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub GridMultipleStore_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(GridMultipleStore.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub GridMultipleStore_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then GetDataBackFromGridMultipleStoreToTexBoxes
End Sub

Private Sub GridMultipleStore_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtMSTotalItems.Text = Val(TxtMSTotalItems.Text) - (GridMultipleStore.Columns("QtyLoose").Value + Val(GridMultipleStore.Columns("Bonus").Value) + (IIf(Val(GridMultipleStore.Columns("Pack").Value) = 0, 0, GridMultipleStore.Columns("Pack").Value) * IIf(Val(GridMultipleStore.Columns("QtyPack").Value) = 0, 0, GridMultipleStore.Columns("QtyPack").Value)))
   SubClearMSDetailArea
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub GridMultipleStore_LostFocus()
   Flag = False
   
'   If CmbMSStore.Enabled = True Then
'      CmbMSStore.SetFocus
'   Else
'      TxtQtyLoose.SetFocus
'   End If
End Sub

Private Sub SubMultipleStore()
   On Error GoTo ErrorHandler
      If FunValidationDetailData = False Then Exit Sub
      If vUseMultipleStore Then
         PopulateDataToGridMultipleStore
         SubCountTotalQty
         FrmMultipleStore.Visible = True
         FrmMultipleStore.Enabled = True
         FrmMultipleStore.ZOrder (0)
         Grid.Enabled = False
         CmbMSStore.Enabled = True
         TxtMSCode.Text = TxtCode.Text
         TxtMSProductName.Text = TxtProductName.Text
         TxtMSMultiplier.Text = TxtMultiplier.Text
         CmbMSStore.Text = TxtStoreName.Text
         TxtMSQtyLoose.SetFocus
      End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub GetDataFromTexBoxesToGridMultipleStore()
   On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   If Trim(TxtMSCode.Text) = "" Or Trim(TxtCode.Text) = "" Then
      TxtMSQtyLoose.Text = ""
      MsgBox "Please Select Product from Sale Invoice", vbInformation, Me.Caption
      Exit Sub
   End If
'   Val (TxtMSTotalItems.Text) + (Val(TxtMSQtyLoose.Text) + Val(TxtMSBonus.Text) + (Val(TxtMSQtyPack.Text) * Val(TxtMSMultiplier.Text)))
'   If TxtMSTotalItems.Text > Val(TxtQtyLoose.Text) Then
'      MsgBox "Please Total Store Qty should be equal to Sale Invoice Product Qty", vbInformation, Me.Caption
'      Exit Sub
'   End If
      RsBodyStore.Filter = 0
      RsBodyStore.Filter = "ProductID = " & Val(TxtMSCode.Text) & " and StoreID = " & CmbMSStore.ItemData(CmbMSStore.ListIndex)
   If CmbMSStore.Enabled = True Then
      If RsBodyStore.RecordCount = 0 Then
         RsBodyStore.AddNew
         GridMultipleStore.Columns("ProductID").Text = TxtMSCode.Text
         GridMultipleStore.Columns("ProductName").Text = TxtMSProductName.Text
         GridMultipleStore.Columns("StoreID").Text = CmbMSStore.ItemData(CmbMSStore.ListIndex)
         GridMultipleStore.Columns("StoreName").Text = CmbMSStore.Text
         RsBodyStore!Productid = TxtMSCode.Text
         RsBodyStore!StoreID = CmbMSStore.ItemData(CmbMSStore.ListIndex)
      Else
         GridMultipleStore.Redraw = False
         GridMultipleStore.MoveFirst
            For vrowcounter = 1 To GridMultipleStore.rows
               If GridMultipleStore.Columns("Productid").Text = TxtMSCode.Text And GridMultipleStore.Columns("StoreID").Text = CmbMSStore.ItemData(CmbMSStore.ListIndex) Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                                    
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  
                  TxtMSQtyPack.Text = Val(TxtMSQtyPack.Text) + Val(GridMultipleStore.Columns("QtyPack").Value)
                  TxtMSQtyLoose.Text = Val(TxtMSQtyLoose.Text) + Val(GridMultipleStore.Columns("QtyLoose").Value)
                  TxtMSBonus.Text = Val(TxtMSBonus.Text) + Val(GridMultipleStore.Columns("Bonus").Value)
                  
                  vProductQtyPack = vProductQtyPack - Val(TxtMSQtyPack.Text) + Val(GridMultipleStore.Columns("QtyPack").Value)
                  vProductQtyLoose = vProductQtyLoose - Val(TxtMSQtyLoose.Text) + Val(GridMultipleStore.Columns("QtyLoose").Value)
                  vProductBonus = vProductBonus - Val(TxtMSBonus.Text) + Val(GridMultipleStore.Columns("Bonus").Value)
                  
                  TxtMSTotalItems.Text = Val(TxtMSTotalItems.Text) + (Val(TxtMSQtyLoose.Text) + Val(TxtMSBonus.Text) + (Val(TxtMSQtyPack.Text) * Val(TxtMSMultiplier.Text))) - (Val(GridMultipleStore.Columns("QtyLoose").Value) + Val(GridMultipleStore.Columns("Bonus").Value) + (IIf(Val(GridMultipleStore.Columns("Pack").Value) = 0, 0, GridMultipleStore.Columns("Pack").Value) * IIf(Val(GridMultipleStore.Columns("QtyPack").Value) = 0, 0, Val(GridMultipleStore.Columns("QtyPack").Value))))
                  
                  
                  GridMultipleStore.Columns("PackName").Text = CmbMSPackName.Text
                  GridMultipleStore.Columns("PackingID").Value = IIf(CmbMSPackName.ListIndex > 0, CmbMSPackName.ItemData(CmbMSPackName.ListIndex), "")
                  GridMultipleStore.Columns("Pack").Value = IIf(Val(TxtMSMultiplier.Text) = 0, "", Val(TxtMSMultiplier.Text))
                  GridMultipleStore.Columns("QtyPack").Value = IIf(Val(TxtMSQtyPack.Text) = 0, "", Val(TxtMSQtyPack.Text))
                  GridMultipleStore.Columns("QtyLoose").Value = Val(TxtMSQtyLoose.Text)
                  RsBodyStore!PackingID = IIf(CmbMSPackName.ListIndex = 0, Null, CmbMSPackName.ItemData(CmbMSPackName.ListIndex))
                  RsBodyStore!QtyPack = IIf(Val(TxtMSQtyPack.Text) = 0, 0, Val(TxtMSQtyPack.Text))
                  RsBodyStore!Qty = Val(TxtMSQtyLoose.Text)
                  RsBodyStore!Bonus = Val(TxtMSBonus.Text)
'                  vProductQtyPack = Val(TxtMSQtyPack.Text) - vProductQtyPack
'                  vProductQtyLoose = Val(TxtMSQtyLoose.Text) - vProductQtyLoose
'                  vProductBonus = Val(TxtMSBonus.Text) - vProductBonus
                  GridMultipleStore.MoveLast
                  Call SubClearMSDetailArea
                  GridMultipleStore.Redraw = True
                  Exit Sub
               End If
               GridMultipleStore.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         GridMultipleStore.MoveLast
'         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   GridMultipleStore.Redraw = False
   With GridMultipleStore
      If CmbMSStore.Enabled = True Then
         TxtMSTotalItems.Text = Val(TxtMSTotalItems.Text) + (Val(TxtMSQtyLoose.Text) + Val(TxtMSBonus.Text) + (Val(TxtMSQtyPack.Text) * Val(TxtMSMultiplier.Text)))
         vProductQtyLoose = vProductQtyLoose - Val(TxtMSQtyLoose.Text)
         vProductQtyPack = vProductQtyPack - Val(TxtMSQtyPack.Text)
         vProductBonus = vProductBonus - Val(TxtMSBonus.Text)
      Else
         TxtMSTotalItems.Text = Val(TxtMSTotalItems.Text) + (Val(TxtMSQtyLoose.Text) + Val(TxtMSBonus.Text) + (Val(TxtMSQtyPack.Text) * Val(TxtMSMultiplier.Text))) - (GridMultipleStore.Columns("QtyLoose").Value + Val(GridMultipleStore.Columns("Bonus").Value) + (IIf(Val(GridMultipleStore.Columns("Pack").Value) = 0, 0, Val(GridMultipleStore.Columns("Pack").Value)) * IIf(Val(GridMultipleStore.Columns("QtyPack").Value) = 0, 0, Val(GridMultipleStore.Columns("QtyPack").Value))))
         vProductQtyLoose = vProductQtyLoose - Val(TxtMSQtyLoose.Text)
         vProductQtyPack = vProductQtyPack - Val(TxtMSQtyPack.Text)
         vProductBonus = vProductBonus - Val(TxtMSBonus.Text)
      End If
      .Columns("PackingID").Value = IIf(CmbMSPackName.ListIndex > 0, CmbMSPackName.ItemData(CmbMSPackName.ListIndex), "")
      .Columns("PackName").Text = CmbMSPackName.Text
      .Columns("Pack").Value = IIf(Val(TxtMSMultiplier.Text) = 0, "", Val(TxtMSMultiplier.Text))
      .Columns("QtyPack").Value = IIf(Val(TxtMSQtyPack.Text) = 0, "", Val(TxtMSQtyPack.Text))
      .Columns("QtyLoose").Value = Val(TxtMSQtyLoose.Text)
       RsBodyStore!PackingID = IIf(CmbMSPackName.ListIndex = 0, Null, CmbMSPackName.ItemData(CmbMSPackName.ListIndex))
       RsBodyStore!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMSMultiplier.Text))
       RsBodyStore!QtyPack = IIf(Val(TxtMSQtyPack.Text) = 0, 0, Val(TxtMSQtyPack.Text))
       RsBodyStore!Qty = Val(TxtMSQtyLoose.Text)
       RsBodyStore!Bonus = Val(TxtMSBonus.Text)
     If CmbMSStore.Enabled = True Then
         .MoveLast
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      Else
         .MoveLast
      End If
      
         
   End With
   
   SubClearMSDetailArea
   GridMultipleStore_LostFocus
   GridMultipleStore.Redraw = True
   
   Exit Sub
ErrorHandler:
   GridMultipleStore.Redraw = True
   Call ShowErrorMessage
End Sub
Private Sub GetDataBackFromGridMultipleStoreToTexBoxes()
   On Error GoTo ErrorHandler
   If Trim(TxtCode.Text) = "" Then
'      MsgBox "Please Select Product from Sale Invoice", vbInformation, Me.Caption
      Exit Sub
   End If
   TxtMSCode.Text = TxtCode.Text
   TxtMSProductName.Text = TxtProductName.Text
   With GridMultipleStore
      If Trim(.Columns("StoreName").Text) <> "" Then
         CmbMSStore.Text = .Columns("StoreName").Text
      End If
      If Trim(.Columns("PackName").Text) = "" Then
         CmbMSPackName.ListIndex = 0
      Else
         CmbMSPackName.Text = .Columns("PackName").Text
      End If
      TxtMSMultiplier.Text = .Columns("Pack").Text
      TxtMSQtyLoose.Text = .Columns("QtyLoose").Text
      TxtMSQtyPack.Text = .Columns("QtyPack").Text
      TxtBonus.Text = .Columns("Bonus").Text
      If Trim(.Columns("ProductID").Text) = "" Then CmbMSStore.Enabled = True Else CmbMSStore.Enabled = False
      End With
'      CmbMSStore.Enabled = False
'      If GridMultipleStore.Rows = 1 Then GridMultipleStore.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearMSDetailArea()
On Error GoTo ErrorHandler
   TxtMSQtyPack.Text = ""
   TxtMSQtyLoose.Text = ""
   CmbMSStore.Enabled = True
   CmbMSStore.SetFocus
   TxtMSQtyLoose.Text = vProductQtyLoose
   TxtMSQtyPack.Text = vProductQtyPack
   TxtMSBonus.Text = vProductBonus
   If Trim(CmbMSPackName.Text) = "" Then CmbMSPackName.Enabled = False Else CmbMSPackName.Enabled = True
   If Val(vProductDetailQty) = Val(TxtMSTotalItems.Text) Then
      Grid.Enabled = True
      GetDataFromTexBoxesToGrid
      If TxtCode.Enabled = True Then TxtCode.SetFocus
   End If
Exit Sub
ErrorHandler:
   GridMultipleStore.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGridMultipleStore()
   
   vProductQtyPack = 0
   vProductQtyLoose = 0
   vProductBonus = 0
   RsBodyStore.Filter = "ProductID = " & Val(IIf(Trim(TxtProductID.Text) = "", Grid.Columns("ProductID").Text, TxtProductID.Text))
   If RsBodyStore.RecordCount > 0 Then
'      sSql = "select d.* from PurchaseBodySerial d  where PurID=" & Val(TxtPurchaseID.Text) & " and Purchasedate='" & DtpPurchaseDate.DateValue & "' and ProductID = '" & Grid.Columns("ProductID").Text & "'"
'      With CN.Execute(sSql)
       With RsBodyStore
         GridMultipleStore.Redraw = False
         GridMultipleStore.MoveFirst
         GridMultipleStore.RemoveAll
         GridMultipleStore.AllowAddNew = True
         TxtMSTotalItems.Text = 0
         .MoveFirst
         While Not .EOF
            GridMultipleStore.AddNew
            GridMultipleStore.Columns("ProductID").Text = !Productid
            GridMultipleStore.Columns("StoreID").Text = !StoreID
            GridMultipleStore.Columns("StoreName").Text = CN.Execute("Select StoreName from Stores where StoreID = " & !StoreID).Fields(0).Value
            GridMultipleStore.Columns("ProductName").Text = CN.Execute("Select ProductName from Products where ProductID = " & !Productid).Fields(0).Value
            If !PackingID = 0 Or IsNull(!PackingID) Then
               GridMultipleStore.Columns("PackingID").Value = ""
            Else
               GridMultipleStore.Columns("PackingID").Value = !PackingID
            End If
            If !PackingID = 0 Or IsNull(!PackingID) Then
               GridMultipleStore.Columns("PackName").Text = ""
            Else
               GridMultipleStore.Columns("PackName").Text = CN.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
            End If
            GridMultipleStore.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
            GridMultipleStore.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            GridMultipleStore.Columns("QtyLoose").Value = !Qty
            GridMultipleStore.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            vProductQtyPack = vProductQtyPack + IIf(IsNull(!QtyPack), 0, !QtyPack)
            vProductQtyLoose = vProductQtyLoose + !Qty
            vProductBonus = Val(vProductBonus) + !Bonus
            TxtMSTotalItems.Text = Val(TxtMSTotalItems.Text) + !Qty + IIf(IsNull(!Bonus), "0", !Bonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack))
            .MoveNext
         Wend
'         .Close
      End With
      GridMultipleStore.AddNew
      GridMultipleStore.Columns("ProductID").Text = " "
      GridMultipleStore.AllowAddNew = False
      GridMultipleStore.Redraw = True
      FrmMultipleStore.Enabled = False
      FrmMultipleStore.ZOrder (0)
   Else
      Call SubClearMultipleStoreFields
'      vProductQtyPack = Val(TxtQtyPack.Text)
'      vProductQtyLoose = TxtQtyLoose.Text
'      vProductBonus = Val(TxtBonus.Text)
   End If
   RsBodyStore.Filter = 0
End Sub

Private Function FunValidationDetailData() As Boolean
   On Error GoTo ErrorHandler
   FunValidationDetailData = False
   
   If Trim(TxtCode.Text) = "" Then
      TxtCode.SetFocus
      Exit Function
   End If
   If CmbPackName.ListIndex > 0 Then
      If Trim(TxtMultiplier.Text) = 0 Then
         TxtMultiplier.SetFocus
         Exit Function
      End If
   End If
   
   If Val(TxtQtyPack.Text) = 0 And Val(TxtQtyLoose.Text) = 0 Then
      If TxtQtyPack.Enabled Then TxtQtyPack.SetFocus Else TxtQtyLoose.SetFocus
      Exit Function
   End If
   
   If vBottomPrice > 0 Then
    If Round(Val(TxtAmount.Text) / (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 2) < vBottomPrice Then
        MsgBox "Sale Price is less than Bottom Price.", vbExclamation, "Alert"
        Exit Function
    End If
   
   
   End If
   
   
   '************* Last Pur Price Check **************
   If Round(Val(LblLastPurPrice.Caption), 3) * IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text)) > Val(TxtPrice.Text) And ObjUserSecurity.SalePriceMustBeLessThanPurchase = True Then
      MsgBox "Sale Price is Less than Last (" & Round(Val(LblLastPurPrice.Caption), 3) & ").", vbInformation + vbOKOnly, "Alert"
      Exit Function
   End If
   
   
   
   
   If Round(Val(LblLastPurPrice.Caption), 3) * IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text)) > Val(TxtPrice.Text) Then
      If MsgBox("Sale Price is Less than Last (" & Round(Val(LblLastPurPrice.Caption), 3) & "). Do You want to continue?", vbQuestion + vbYesNo, "Alert") = vbNo Then Exit Function
   End If
   
'   If Round(Val(LblLastPurPrice.Caption), 3) * IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text)) > Val(TxtPrice.Text) And ObjUserSecurity.SalePriceMustBeLessThanPurchase = True Then
'      If ObjRegistry.isShowListPrice = True Then
'         MsgBox "Sale Price is Less than Last (" & Round(Val(LblLastPurPrice.Caption), 3) * IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text)) & ")", vbInformation + vbOKOnly, "Alert"
'         If TxtPrice.Enabled Then TxtPrice.SetFocus
'         Exit Function
'      Else
'         If MsgBox("Sale Price is Less than Last (" & Round(Val(LblLastPurPrice.Caption), 3) * IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text)) & "). Do You want to continue?", vbQuestion + vbYesNo, "Alert") = vbNo Then Exit Function
'      End If
'   End If
   
   FrmHistory.Visible = False
   FrmProductPrices.Visible = False
   
   If Val(TxtPrice.Text) <> 0 Then
      If Round(Val(TxtDiscPer.Text), 2) <> Round((Val(TxtDiscPC.Text) * 100) / (Val(TxtPrice.Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))), 2) Then
         MsgBox "Please update the Discount for change Price.", vbExclamation, "Alert"
         If TxtDiscPer.Enabled And TxtDiscPer.Visible Then TxtDiscPer.SetFocus
         Exit Function
      End If
   End If
   
   
   If ObjRegistry.NegativeSale = False Then
      If vIsNewRecord = True Then
         If (Val(vQtyLoose) - (Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text))) < 0 Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            Grid.Redraw = True
            Call SubClearDetailArea
            Grid.MoveLast
            If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
            Exit Function
         End If
      Else
         If (Val(vQtyLoose) - (Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)) + (Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value) + Val(Grid.Columns("QtyLoose").Value))) < 0 Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            Grid.Redraw = True
            Call SubClearDetailArea
            Grid.MoveLast
            If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
            Exit Function
         End If
      End If
   End If


 
 vAlreadySerial = False
 If TxtSerial.Text <> "" Then
      TxtSerial.Enabled = True
      GetDataFromTexBoxesToGridSerial
      If vAlreadySerial = True Then
         Call SubClearDetailArea
         If TxtCode.Enabled = True Then TxtCode.SetFocus
         Exit Function
      End If
 End If
   FunValidationDetailData = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Sub SubCountTotalQty2()
   On Error GoTo ErrorHandler
   If TxtCode.Enabled = True Then
      TxtMSQtyPack.Text = vProductQtyPack
      TxtMSQtyLoose.Text = vProductQtyLoose
      TxtMSBonus.Text = vProductBonus
   Else
      TxtMSQtyPack.Text = Val(TxtQtyPack.Text) - vProductQtyPack
      TxtMSQtyLoose.Text = Val(TxtQtyLoose.Text) - vProductQtyLoose
      TxtMSBonus.Text = Val(TxtBonus.Text) - vProductBonus
   End If
   vProductDetailQty = (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text)))
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub SubCountTotalQty()
   On Error GoTo ErrorHandler
   If TxtCode.Enabled = True Then
      If Trim(CmbPackName.Text) <> "" Then
         CmbMSPackName.Clear
         CmbMSPackName.AddItem ""
         CmbMSPackName.AddItem CmbPackName.Text
         CmbMSPackName.ItemData(CmbMSPackName.NewIndex) = CmbPackName.ItemData(CmbPackName.ListIndex)
         CmbMSPackName.ListIndex = 1
         
      End If
      
      TxtMSQtyPack.Text = Val(TxtQtyPack.Text)
      TxtMSQtyLoose.Text = Val(TxtQtyLoose.Text)
      TxtMSBonus.Text = Val(TxtBonus.Text)
      
      vProductQtyPack = vProductQtyPack + Val(TxtQtyPack.Text)
      vProductQtyLoose = vProductQtyLoose + Val(TxtQtyLoose.Text)
      vProductBonus = vProductBonus + Val(TxtBonus.Text)
      
   Else
      TxtMSQtyPack.Text = Val(TxtQtyPack.Text) - vProductQtyPack
      TxtMSQtyLoose.Text = Val(TxtQtyLoose.Text) - vProductQtyLoose
      TxtMSBonus.Text = Val(TxtBonus.Text) - vProductBonus
      
      vProductQtyPack = Val(TxtQtyPack.Text)
      vProductQtyLoose = Val(TxtQtyLoose.Text)
      vProductBonus = Val(TxtBonus.Text)
   End If

   vProductDetailQty = (Val(vProductQtyLoose) + Val(vProductBonus) + (Val(vProductQtyPack) * Val(TxtMultiplier.Text)))
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmSaleInvoiceDIS) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmSaleInvoiceDIS & ", " & vUser & "," & TableHeaderFields(eFrmSaleInvoiceDIS) & " from SaleHeader " & vbCrLf _
             & "Where SID = " & TxtSID.Text
      CN.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleBodyBin (" & TableBodyFields(eFrmSaleInvoiceDIS) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmSaleInvoiceDIS) & " from SaleBody " & vbCrLf _
             & "Where SID = " & TxtSID.Text
      CN.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DefaultCustomer()
      CN.Execute "Delete From AccountsLedger"
      
      vStrSQL = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[#AccountsLedger]') and OBJECTPROPERTY(id, N'IsTable') = 1)" & vbCrLf & _
     "drop Table [dbo].[#AccountsLedger]"
      CN.Execute vStrSQL
'     CN.Execute "drop Table [dbo].[#AccountsLedger]"
      
    
    vStrSQL = " CREATE TABLE [dbo].[#AccountsLedger] (" & vbCrLf & _
      " [organizationID] [tinyint] NULL ," & vbCrLf & _
      " [AccountNo] [varchar] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & vbCrLf & _
      " [VoucherType] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & vbCrLf & _
      " [VoucherNo] [int] NULL ," & vbCrLf & _
      " [StrVoucherNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & vbCrLf & _
      " [VoucherDate] [smalldatetime] NULL ," & vbCrLf & _
      " [Debit] [numeric](12, 2) NULL ," & vbCrLf & _
      " [Credit] [numeric](12, 2) NULL ," & vbCrLf & _
      " [Naration] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & vbCrLf & _
      " [EntryTime] [datetime] NULL ," & vbCrLf & _
      " [SessionID] [smallint] NULL" & vbCrLf & _
      ") ON [PRIMARY]"

   CN.Execute vStrSQL
   
      vStrSQL = "Select c.* from ChartofAccounts c  " & vbCrLf & _
      " left outer JOIN Parties p ON  p.PartyID = c.AccountNo " & vbCrLf & _
        " left outer Join Sectors S on S.sectorId = p.sectorID" & vbCrLf & _
        " left outer join Zones Z on Z.ZoneID = S.ZoneID " & vbCrLf & _
        " Where c.AccountNo Like '62%' and c.AccountNo <> '621' and C.AccountNo = " & TxtCustomerID.Text
      With CN.Execute(vStrSQL)
         While Not .EOF
            vStrSQL = "EXECUTE SPAccountsLedgerNew " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & ",'" & !AccountNo & "', '" & DtpBillDate.DateValue - 365 & "','" & DtpBillDate.DateValue & "'," & 1
            CN.Execute vStrSQL
            .MoveNext
         Wend
      End With
      CN.Execute "Insert into Accountsledger Select * from #AccountsLedger"
      CN.Execute "Delete Accountsledger where credit = 0"
      vStrSQL = "DELETE t1 " & vbCrLf & _
            "FROM Accountsledger t1, Accountsledger t2 " & vbCrLf & _
            "Where t1.AccountNo = t2.AccountNo  " & vbCrLf & _
            "AND t1.Entrytime < t2.Entrytime "
      CN.Execute (vStrSQL)
      vStrSQL = " Drop TABLE [dbo].[#AccountsLedger] "
      CN.Execute vStrSQL


End Sub
