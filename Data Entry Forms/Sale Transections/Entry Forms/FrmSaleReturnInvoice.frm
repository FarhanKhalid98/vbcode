VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmSaleReturnInvoice 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmSaleReturnInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkIsPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Print"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11310
      TabIndex        =   171
      Top             =   9930
      Width           =   1290
   End
   Begin VB.CheckBox ChkDiscB4SaleTax 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13725
      TabIndex        =   164
      Top             =   6975
      Width           =   1290
   End
   Begin VB.CheckBox ChkDiscB4ExtraScheme 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13770
      TabIndex        =   163
      Top             =   5805
      Width           =   1290
   End
   Begin VB.CheckBox ChkDiscB4TradeOffer 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4 "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13725
      TabIndex        =   162
      Top             =   4725
      Width           =   1290
   End
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10095
      TabIndex        =   158
      Top             =   9930
      Width           =   1290
   End
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   8835
      TabIndex        =   155
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   9885
      Width           =   1170
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmSaleReturnInvoice.frx":0ECA
      Left            =   4650
      List            =   "FrmSaleReturnInvoice.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   154
      Tag             =   "1"
      Top             =   9885
      Width           =   3276
   End
   Begin VB.Frame FrmHistory 
      Height          =   1635
      Left            =   2805
      TabIndex        =   145
      Top             =   6015
      Visible         =   0   'False
      Width           =   10260
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridHistory 
         Height          =   1455
         Left            =   90
         TabIndex        =   146
         Top             =   135
         Width           =   10020
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
         stylesets(0).Picture=   "FrmSaleReturnInvoice.frx":0ECE
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
         stylesets(1).Picture=   "FrmSaleReturnInvoice.frx":0EEA
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
         stylesets(2).Picture=   "FrmSaleReturnInvoice.frx":0F06
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
         Columns(2).Width=   1561
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
         _ExtentX        =   17674
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
   Begin VB.Frame FrmProductPrices 
      Height          =   1095
      Left            =   6615
      TabIndex        =   151
      Top             =   135
      Visible         =   0   'False
      Width           =   6270
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridProductPrices 
         Height          =   885
         Left            =   60
         TabIndex        =   152
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
         stylesets(0).Picture=   "FrmSaleReturnInvoice.frx":0F22
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
         stylesets(1).Picture=   "FrmSaleReturnInvoice.frx":0F3E
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
         stylesets(2).Picture=   "FrmSaleReturnInvoice.frx":0F5A
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
   Begin VB.Frame FramExpense 
      Height          =   2415
      Left            =   9315
      TabIndex        =   125
      Top             =   5355
      Visible         =   0   'False
      Width           =   4215
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridExpense 
         Height          =   1860
         Left            =   120
         TabIndex        =   126
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
         stylesets(0).Picture=   "FrmSaleReturnInvoice.frx":0F76
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
         stylesets(1).Picture=   "FrmSaleReturnInvoice.frx":0F92
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
         stylesets(2).Picture=   "FrmSaleReturnInvoice.frx":0FAE
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
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   10350
      MaxLength       =   50
      TabIndex        =   103
      Top             =   8865
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CheckBox ChkIsProduct 
      Caption         =   "Is Product"
      Height          =   255
      Left            =   690
      TabIndex        =   97
      Top             =   1035
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   5580
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   4380
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   1305
      TabIndex        =   74
      Top             =   7965
      Width           =   2295
      Begin SITextBox.Txt TxtSerial 
         Height          =   315
         Left            =   120
         TabIndex        =   75
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
         TabIndex        =   76
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
         stylesets(0).Picture=   "FrmSaleReturnInvoice.frx":0FCA
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
         Columns(1).Caption=   "Serial (s)"
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
      Left            =   13455
      TabIndex        =   68
      Top             =   45
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
         Left            =   270
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   69
         Tag             =   "NC"
         Text            =   "FrmSaleReturnInvoice.frx":0FE6
         Top             =   450
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
         TabIndex        =   70
         Top             =   90
         Width           =   135
      End
   End
   Begin SITextBox.Txt TxtReturnID 
      Height          =   315
      Left            =   1875
      TabIndex        =   0
      Top             =   1860
      Width           =   690
      _ExtentX        =   1217
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
      Left            =   9075
      TabIndex        =   45
      Top             =   9315
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
      MICON           =   "FrmSaleReturnInvoice.frx":10FD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7770
      TabIndex        =   41
      Top             =   9315
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
      MICON           =   "FrmSaleReturnInvoice.frx":1119
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5160
      TabIndex        =   43
      Top             =   9315
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
      MICON           =   "FrmSaleReturnInvoice.frx":1135
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10380
      TabIndex        =   46
      Top             =   9315
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
      MICON           =   "FrmSaleReturnInvoice.frx":1151
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6465
      TabIndex        =   42
      Top             =   9315
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
      MICON           =   "FrmSaleReturnInvoice.frx":116D
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   5340
      TabIndex        =   49
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
   Begin SITextBox.Txt TxtNetAmount 
      Height          =   315
      Left            =   11145
      TabIndex        =   37
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
      Height          =   330
      Left            =   2655
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4380
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
      MICON           =   "FrmSaleReturnInvoice.frx":1189
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpReturnDate 
      Height          =   315
      Left            =   2565
      TabIndex        =   1
      Top             =   1860
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
      Left            =   7230
      TabIndex        =   4
      Tag             =   "NC"
      Top             =   1860
      Width           =   615
      _ExtentX        =   1085
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
      Left            =   8205
      TabIndex        =   55
      Tag             =   "NC"
      Top             =   1860
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   7845
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   1860
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
      MICON           =   "FrmSaleReturnInvoice.frx":11A5
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3855
      TabIndex        =   44
      Top             =   9315
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
      MICON           =   "FrmSaleReturnInvoice.frx":11C1
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillDisc 
      Height          =   315
      Left            =   8940
      TabIndex        =   36
      Top             =   8265
      Width           =   975
      _ExtentX        =   1720
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   1890
      TabIndex        =   62
      Top             =   990
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
      Left            =   4455
      TabIndex        =   65
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
      Left            =   9915
      TabIndex        =   39
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
      Left            =   3015
      TabIndex        =   17
      Top             =   4380
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
      Left            =   7005
      TabIndex        =   19
      Top             =   4380
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
      Left            =   8025
      TabIndex        =   21
      Top             =   4380
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
      Left            =   7515
      TabIndex        =   20
      Top             =   4380
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
      Left            =   11895
      TabIndex        =   31
      Top             =   4380
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
      Left            =   9585
      TabIndex        =   24
      Top             =   4380
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
      Left            =   11415
      TabIndex        =   27
      Top             =   4380
      Width           =   480
      _ExtentX        =   847
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
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   10230
      TabIndex        =   25
      Top             =   4380
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
      Left            =   8565
      TabIndex        =   22
      Top             =   4380
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
      Left            =   9105
      TabIndex        =   23
      Top             =   4380
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
      Left            =   10905
      TabIndex        =   26
      Top             =   4380
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
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   12570
      TabIndex        =   32
      Top             =   4380
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
      DecimalPoint    =   3
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   9405
      TabIndex        =   5
      Tag             =   "NC"
      Top             =   1860
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
      Left            =   10470
      TabIndex        =   91
      Tag             =   "NC"
      Top             =   1860
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
      Left            =   10110
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   1860
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
      MICON           =   "FrmSaleReturnInvoice.frx":11DD
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtMemberID 
      Height          =   315
      Left            =   10545
      TabIndex        =   12
      Top             =   3225
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
      Left            =   11580
      TabIndex        =   93
      Top             =   3225
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
      Left            =   11220
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   3225
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
      MICON           =   "FrmSaleReturnInvoice.frx":11F9
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCost 
      Height          =   315
      Left            =   3765
      TabIndex        =   98
      Top             =   1020
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
      Left            =   2880
      TabIndex        =   100
      Top             =   1020
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Top             =   4380
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
   Begin SITextBox.Txt TxtRemarks 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   12540
      TabIndex        =   38
      Top             =   8265
      Width           =   1830
      _ExtentX        =   3228
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
      Left            =   12825
      TabIndex        =   106
      Top             =   8865
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridOffer 
      Height          =   1365
      Left            =   1800
      TabIndex        =   99
      Top             =   6555
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
      stylesets(0).Picture=   "FrmSaleReturnInvoice.frx":1215
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
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   2520
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
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   3210
      TabIndex        =   108
      Top             =   2520
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
      Left            =   6855
      TabIndex        =   109
      Top             =   2520
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
   Begin SITextBox.Txt TxtCity 
      Height          =   315
      Left            =   11385
      TabIndex        =   110
      Top             =   2520
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
   Begin JeweledBut.JeweledButton BtnCustomer 
      Height          =   330
      Left            =   2850
      TabIndex        =   111
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
      MICON           =   "FrmSaleReturnInvoice.frx":1231
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   3870
      TabIndex        =   2
      Top             =   1860
      Width           =   645
      _ExtentX        =   1138
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
      Left            =   4515
      TabIndex        =   3
      Top             =   1860
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
   Begin JeweledBut.JeweledButton BtnReturnAll 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6180
      TabIndex        =   116
      Top             =   1860
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
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
      MICON           =   "FrmSaleReturnInvoice.frx":124D
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSale 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5820
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   1860
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
      MICON           =   "FrmSaleReturnInvoice.frx":1269
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtPaidAmount 
      Height          =   315
      Left            =   8895
      TabIndex        =   40
      Top             =   8895
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
      Left            =   5925
      TabIndex        =   120
      Top             =   8895
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
      Left            =   7500
      TabIndex        =   121
      Top             =   8895
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
   Begin SITextBox.Txt TxtBillNo 
      Height          =   315
      Left            =   1890
      TabIndex        =   7
      Top             =   3225
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
      Left            =   2640
      TabIndex        =   8
      Top             =   3225
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
      Left            =   4905
      TabIndex        =   10
      Top             =   3225
      Width           =   3000
      _ExtentX        =   5292
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
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   7905
      TabIndex        =   11
      Top             =   3225
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
      Left            =   9015
      TabIndex        =   129
      Top             =   3225
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
      Left            =   8655
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   3225
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
      MICON           =   "FrmSaleReturnInvoice.frx":1285
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtVehicleNo 
      Height          =   315
      Left            =   3390
      TabIndex        =   9
      Top             =   3225
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
      Left            =   5355
      TabIndex        =   137
      Top             =   1035
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3240
      Left            =   1125
      TabIndex        =   139
      Top             =   4695
      Width           =   12615
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   36
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
      stylesets(0).Picture=   "FrmSaleReturnInvoice.frx":12A1
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
      Columns.Count   =   36
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ProductID"
      Columns(0).Name =   "ProductID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1191
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
      Columns(15).Width=   1508
      Columns(15).Caption=   "Amount"
      Columns(15).Name=   "Amount"
      Columns(15).Alignment=   1
      Columns(15).CaptionAlignment=   2
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   5
      Columns(15).FieldLen=   256
      Columns(16).Width=   3200
      Columns(16).Visible=   0   'False
      Columns(16).Caption=   "PackingID"
      Columns(16).Name=   "PackingID"
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      Columns(17).Width=   3200
      Columns(17).Visible=   0   'False
      Columns(17).Caption=   "SaleTaxVal"
      Columns(17).Name=   "SaleTaxVal"
      Columns(17).Alignment=   1
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      Columns(18).Width=   3200
      Columns(18).Visible=   0   'False
      Columns(18).Caption=   "IsProduct"
      Columns(18).Name=   "IsProduct"
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   11
      Columns(18).FieldLen=   256
      Columns(19).Width=   3200
      Columns(19).Visible=   0   'False
      Columns(19).Caption=   "Cost"
      Columns(19).Name=   "Cost"
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(20).Width=   1958
      Columns(20).Caption=   "RetailPrice"
      Columns(20).Name=   "RetailPrice"
      Columns(20).Alignment=   1
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      Columns(21).Width=   2461
      Columns(21).Caption=   "IsWSDiscb4ST"
      Columns(21).Name=   "IsWSDiscb4ST"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   11
      Columns(21).FieldLen=   256
      Columns(22).Width=   2328
      Columns(22).Caption=   "IsWSSaleTax"
      Columns(22).Name=   "IsWSSaleTax"
      Columns(22).DataField=   "Column 22"
      Columns(22).DataType=   11
      Columns(22).FieldLen=   256
      Columns(23).Width=   2461
      Columns(23).Caption=   "IsRetailSaleTax"
      Columns(23).Name=   "IsRetailSaleTax"
      Columns(23).DataField=   "Column 23"
      Columns(23).DataType=   11
      Columns(23).FieldLen=   256
      Columns(24).Width=   2037
      Columns(24).Caption=   "TokenVal"
      Columns(24).Name=   "TokenVal"
      Columns(24).DataField=   "Column 24"
      Columns(24).DataType=   8
      Columns(24).FieldLen=   256
      Columns(25).Width=   3200
      Columns(25).Caption=   "StampID"
      Columns(25).Name=   "StampID"
      Columns(25).DataField=   "Column 25"
      Columns(25).DataType=   8
      Columns(25).FieldLen=   256
      Columns(26).Width=   1640
      Columns(26).Caption=   "BatchNo"
      Columns(26).Name=   "BatchNo"
      Columns(26).DataField=   "Column 26"
      Columns(26).DataType=   8
      Columns(26).FieldLen=   256
      Columns(27).Width=   3200
      Columns(27).Visible=   0   'False
      Columns(27).Caption=   "isLastPrice"
      Columns(27).Name=   "isLastPrice"
      Columns(27).DataField=   "Column 27"
      Columns(27).DataType=   11
      Columns(27).FieldLen=   256
      Columns(28).Width=   3200
      Columns(28).Caption=   "isDiscB4TradeOffer"
      Columns(28).Name=   "isDiscB4TradeOffer"
      Columns(28).DataField=   "Column 28"
      Columns(28).DataType=   8
      Columns(28).FieldLen=   256
      Columns(29).Width=   3200
      Columns(29).Caption=   "isDiscB4ExtraScheme"
      Columns(29).Name=   "isDiscB4ExtraScheme"
      Columns(29).DataField=   "Column 29"
      Columns(29).DataType=   8
      Columns(29).FieldLen=   256
      Columns(30).Width=   3200
      Columns(30).Caption=   "isDiscB4SaleTax"
      Columns(30).Name=   "isDiscB4SaleTax"
      Columns(30).DataField=   "Column 30"
      Columns(30).DataType=   8
      Columns(30).FieldLen=   256
      Columns(31).Width=   3200
      Columns(31).Caption=   "TradeOffer1"
      Columns(31).Name=   "TradeOffer1"
      Columns(31).DataField=   "Column 31"
      Columns(31).DataType=   8
      Columns(31).FieldLen=   256
      Columns(32).Width=   3200
      Columns(32).Caption=   "TradeOffer2"
      Columns(32).Name=   "TradeOffer2"
      Columns(32).DataField=   "Column 32"
      Columns(32).DataType=   8
      Columns(32).FieldLen=   256
      Columns(33).Width=   3200
      Columns(33).Caption=   "ExtraSchemePer"
      Columns(33).Name=   "ExtraSchemePer"
      Columns(33).DataField=   "Column 33"
      Columns(33).DataType=   8
      Columns(33).FieldLen=   256
      Columns(34).Width=   3200
      Columns(34).Caption=   "TradeValue"
      Columns(34).Name=   "TradeValue"
      Columns(34).DataField=   "Column 34"
      Columns(34).DataType=   8
      Columns(34).FieldLen=   256
      Columns(35).Width=   3200
      Columns(35).Caption=   "ExtraSchemeValue"
      Columns(35).Name=   "ExtraSchemeValue"
      Columns(35).DataField=   "Column 35"
      Columns(35).DataType=   8
      Columns(35).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   22251
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
   Begin SITextBox.Txt TxtTokenVal 
      Height          =   315
      Left            =   4410
      TabIndex        =   140
      Top             =   1035
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
      Left            =   2655
      TabIndex        =   14
      Top             =   4065
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
   Begin SITextBox.Txt TxtExpiryDate 
      Height          =   315
      Left            =   3510
      TabIndex        =   15
      Top             =   4065
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
   Begin JeweledBut.JeweledButton BtnProductRange 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2295
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   4050
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
      MICON           =   "FrmSaleReturnInvoice.frx":12BD
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtSID 
      Height          =   315
      Left            =   1035
      TabIndex        =   147
      Top             =   1860
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
   Begin SITextBox.Txt TxtBillDiscPer 
      Height          =   315
      Left            =   7965
      TabIndex        =   35
      Top             =   8265
      Width           =   975
      _ExtentX        =   1720
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
   Begin SITextBox.Txt TxtServiceCharges 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6525
      TabIndex        =   33
      Top             =   8265
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
   Begin SITextBox.Txt TxtServiceChargesPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7365
      TabIndex        =   34
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
   Begin SITextBox.Txt TxtTradeOffer2 
      Height          =   315
      Left            =   8055
      TabIndex        =   29
      Top             =   3810
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
      Left            =   7290
      TabIndex        =   28
      Top             =   3810
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
      Left            =   10905
      TabIndex        =   30
      Top             =   3780
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
      Left            =   13770
      TabIndex        =   165
      Top             =   5145
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
      Left            =   13770
      TabIndex        =   166
      Top             =   6225
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
      Left            =   13770
      TabIndex        =   167
      Top             =   7425
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
   Begin VB.Label LblTradeValue 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Value"
      Height          =   195
      Left            =   13770
      TabIndex        =   170
      Top             =   4950
      Width           =   870
   End
   Begin VB.Label LblExtraSchemeValue 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ex Scheme Value"
      Height          =   195
      Left            =   13770
      TabIndex        =   169
      Top             =   6030
      Width           =   1260
   End
   Begin VB.Label LblGSTValue 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "GST Value"
      Height          =   195
      Left            =   13770
      TabIndex        =   168
      Top             =   7200
      Width           =   780
   End
   Begin VB.Label LblExtraSchemePer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ex. Scheme %"
      Height          =   195
      Left            =   9810
      TabIndex        =   161
      Top             =   3825
      Width           =   1020
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
      Left            =   7920
      TabIndex        =   160
      Top             =   3855
      Width           =   120
   End
   Begin VB.Label LblTradeOffer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Offer"
      Height          =   195
      Left            =   7455
      TabIndex        =   159
      Top             =   3600
      Width           =   810
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
      Left            =   7935
      TabIndex        =   157
      Top             =   9930
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
      Left            =   3975
      TabIndex        =   156
      Top             =   9930
      Width           =   570
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
      Left            =   720
      TabIndex        =   153
      Top             =   45
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   7365
      TabIndex        =   150
      Top             =   8040
      Width           =   210
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Ch."
      Height          =   195
      Left            =   6525
      TabIndex        =   149
      Top             =   8040
      Width           =   825
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "SID"
      Height          =   195
      Left            =   1035
      TabIndex        =   148
      Top             =   1665
      Visible         =   0   'False
      Width           =   270
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
      Left            =   675
      TabIndex        =   144
      Top             =   675
      Visible         =   0   'False
      Width           =   855
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
      Left            =   4410
      TabIndex        =   143
      Top             =   3915
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Token Val"
      Height          =   195
      Left            =   4410
      TabIndex        =   141
      Top             =   855
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ret Price"
      Height          =   195
      Left            =   5355
      TabIndex        =   138
      Top             =   855
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      Height          =   195
      Left            =   3375
      TabIndex        =   136
      Top             =   3015
      Width           =   780
   End
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   9000
      TabIndex        =   135
      Top             =   3015
      Width           =   780
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   7890
      TabIndex        =   134
      Top             =   3015
      Width           =   525
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      Height          =   195
      Left            =   1890
      TabIndex        =   133
      Top             =   3015
      Width           =   495
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bilty No."
      Height          =   195
      Left            =   2625
      TabIndex        =   132
      Top             =   3015
      Width           =   585
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   4905
      TabIndex        =   131
      Top             =   3015
      Width           =   795
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Amount"
      Height          =   195
      Left            =   8895
      TabIndex        =   124
      Top             =   8670
      Width           =   900
   End
   Begin VB.Label lblPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Receivable"
      Height          =   195
      Left            =   5925
      TabIndex        =   123
      Top             =   8670
      Width           =   1470
   End
   Begin VB.Label LblTtlPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Receivable"
      Height          =   195
      Left            =   7500
      TabIndex        =   122
      Top             =   8670
      Width           =   1215
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   4530
      TabIndex        =   119
      Top             =   1650
      Width           =   585
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " Bill ID"
      Height          =   195
      Left            =   3885
      TabIndex        =   118
      Top             =   1650
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      Height          =   195
      Left            =   1905
      TabIndex        =   115
      Top             =   2310
      Width           =   870
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   195
      Left            =   3210
      TabIndex        =   114
      Top             =   2310
      Width           =   1125
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   6855
      TabIndex        =   113
      Top             =   2310
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   11385
      TabIndex        =   112
      Top             =   2310
      Width           =   255
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill No"
      Height          =   195
      Left            =   12825
      TabIndex        =   107
      Top             =   8640
      Width           =   1020
   End
   Begin VB.Label LblRemarks 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   12540
      TabIndex        =   105
      Top             =   8040
      Width           =   630
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
      Left            =   3825
      TabIndex        =   104
      Top             =   8520
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   3810
      TabIndex        =   102
      Top             =   795
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      Height          =   195
      Left            =   2805
      TabIndex        =   101
      Top             =   795
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblMemberID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      Height          =   195
      Left            =   10530
      TabIndex        =   96
      Top             =   2985
      Width           =   780
   End
   Begin VB.Label LblMemberName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      Height          =   195
      Left            =   11565
      TabIndex        =   95
      Top             =   2985
      Width           =   1035
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   10545
      TabIndex        =   90
      Top             =   1635
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   9360
      TabIndex        =   89
      Top             =   1635
      Width           =   1095
   End
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   12570
      TabIndex        =   88
      Top             =   4185
      Width           =   540
   End
   Begin VB.Label LblSaleTaxPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax%"
      Height          =   195
      Left            =   10905
      TabIndex        =   87
      Top             =   4185
      Width           =   390
   End
   Begin VB.Label LblOffer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Offer"
      Height          =   195
      Left            =   9090
      TabIndex        =   86
      Top             =   4185
      Width           =   345
   End
   Begin VB.Label LblPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   9585
      TabIndex        =   85
      Top             =   4185
      Width           =   405
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   7035
      TabIndex        =   84
      Top             =   4185
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   5610
      TabIndex        =   83
      Top             =   4185
      Width           =   840
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   8025
      TabIndex        =   82
      Top             =   4185
      Width           =   465
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   7515
      TabIndex        =   81
      Top             =   4185
      Width           =   480
   End
   Begin VB.Label LblDiscVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.Val"
      Height          =   195
      Left            =   11895
      TabIndex        =   80
      Top             =   4185
      Width           =   585
   End
   Begin VB.Label LblDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Dis%"
      Height          =   195
      Left            =   11415
      TabIndex        =   79
      Top             =   4185
      Width           =   345
   End
   Begin VB.Label LblDiscPC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc/PC"
      Height          =   195
      Left            =   10230
      TabIndex        =   78
      Top             =   4185
      Width           =   600
   End
   Begin VB.Label LblBonus 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bns(L)"
      Height          =   195
      Left            =   8565
      TabIndex        =   77
      Top             =   4185
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
      Left            =   9705
      TabIndex        =   73
      Top             =   2190
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
      Left            =   8145
      TabIndex        =   72
      Top             =   2190
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
      Left            =   12908
      TabIndex        =   71
      Top             =   1793
      Width           =   435
   End
   Begin VB.Label LblOtherChargesCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      Height          =   195
      Left            =   9915
      TabIndex        =   67
      Top             =   8040
      Width           =   1020
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items"
      Height          =   195
      Left            =   4470
      TabIndex        =   66
      Top             =   8040
      Width           =   780
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Return Invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   64
      Top             =   270
      Width           =   3420
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   1920
      TabIndex        =   63
      Top             =   795
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   8940
      TabIndex        =   61
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
      Left            =   8640
      TabIndex        =   60
      Top             =   1215
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
      Left            =   8640
      TabIndex        =   59
      Top             =   1425
      Width           =   1035
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   8205
      TabIndex        =   58
      Top             =   1665
      Width           =   840
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   7230
      TabIndex        =   57
      Top             =   1665
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   1800
      TabIndex        =   54
      Top             =   4185
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   4410
      TabIndex        =   53
      Top             =   4185
      Width           =   1020
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   13328
      Top             =   1283
      Width           =   330
   End
   Begin VB.Label LblNetAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      Height          =   195
      Left            =   11160
      TabIndex        =   52
      Top             =   8040
      Width           =   840
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount (%)"
      Height          =   195
      Left            =   7950
      TabIndex        =   51
      Top             =   8040
      Width           =   885
   End
   Begin VB.Label LblTotalAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amount"
      Height          =   195
      Left            =   5370
      TabIndex        =   50
      Top             =   8040
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      Height          =   195
      Left            =   2580
      TabIndex        =   48
      Top             =   1665
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ReturnID"
      Height          =   195
      Left            =   1890
      TabIndex        =   47
      Top             =   1665
      Width           =   645
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
Attribute VB_Name = "FrmSaleReturnInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim vMode As FormMode
Dim vDate, vServerDate As Date, vHDiff As Integer, vSystemDate As Boolean
Dim vUnitPrice As Double
Dim vUnitRetailPrice As Double
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
Dim DateFlag As Boolean
Dim Flag As Boolean
Dim ssql As String
Dim vStrSQL, vRandomID As String
Dim vBillID  As Integer, vZoneID As Byte
Dim vBillDate  As Date
Dim ExpenseFlag As Boolean
Dim vSSID, vExpAmount As Double
Dim vQtyLoose As Double
Dim i As Integer, vNoofPrints As Byte, isWholeSale, vTradeOffer, vShowStock As Boolean
Dim vStrDetail As String
Dim vMobileNo() As String, vMobile As String
Dim vStampID As Variant
Dim vPrinter() As String
'----------------------------------

Private Sub SubCalculateBody()
   TxtDiscVal.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) * Val(TxtDiscPer.Text) / 100, 2)
   If Val(TxtDiscVal.Text) = 0 Then TxtDiscVal.Text = ""
   TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), 2)
   'TxtAmount.Text = Round((Val(vUnitPrice) - Val(TxtDiscPC.Text)) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 3)
   If vIsWSSaleTax = True And vIsWSDiscb4ST = True Then
        TxtSaleTaxVal.Text = Round(Val(TxtAmount.Text) * Val(TxtSaleTaxPer.Text) / 100, 3)
    ElseIf vIsWSSaleTax = True And vIsWSDiscb4ST = False Then
        TxtSaleTaxVal.Text = Round(Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
    ElseIf vIsRetailSaleTax = True Then
        TxtSaleTaxVal.Text = Round(Val(vUnitRetailPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
    Else
        TxtSaleTaxVal.Text = Round(Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtSaleTaxPer.Text) / 100, 3)
    End If
    TxtAmount.Text = Val(TxtAmount.Text) + Val(TxtSaleTaxVal.Text) - Val(TxtOffer.Text)
    Call CalculateValue
End Sub

Private Sub SubCalculateFooter()
   If TxtTotalAmount.Text = "" Then Exit Sub
   TxtNetAmount.Text = SelfRound(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text) + Val(TxtServiceCharges.Text)
   TxtTotalReceivable.Text = Abs(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", Val(TxtPreviousReceivable.Text), Val(TxtPreviousReceivable.Text) * -1)))
   LblTtlPayable.Caption = IIf(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", Val(TxtPreviousReceivable.Text), Val(TxtPreviousReceivable.Text) * -1)) > 0, "Total Payable", "Total Receivable")

'   TxtNetAmount.Text = SelfRound(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text)
'   TxtTotalReceivable.Text = Abs(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousReceivable.Text, Val(TxtPreviousReceivable.Text) * -1)))
'   LblTtlPayable.Caption = IIf(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", Val(TxtPreviousReceivable.Text) * 1, Val(TxtPreviousReceivable.Text) * -1)) < 0, "Total Receivable", "Total Payable")
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
      SchProduct.ParaInWhere = "and isLocked = 0"
      SchProduct.ParainShowStock = vShowStock
      SchProduct.Show vbModal, Me
      
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
   
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
  
  
  If Trim(TxtCode.Text) = "" Then Exit Function
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
        vStrSQL = " SELECT p.productid, Code, ProductName, WSPrice, RetailPrice, DiscPC, IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST, TokenVal, SaleTaxPer, PackingName, isnull(Multiplier,0) as Multiplier, " & vbCrLf _
           + " IsDiscB4TradeOffer, IsDiscB4ExtraScheme, isDiscB4SaleTax, TradeOffer1, TradeOffer2, ExtraSchemePer " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.salepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')" & " and isLocked = 0 "
           
 
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = IIf(isWholeSale = True, !WSPrice, !RetailPrice)
         TxtRetailPrice.Text = !RetailPrice
         vIsWSDiscb4ST = !IsWSDiscb4ST
         vIsWSSaleTax = !IsWSSaleTax
         ChkDiscB4TradeOffer.Value = Abs(!isDiscB4TradeOffer)
         ChkDiscB4ExtraScheme.Value = Abs(!isDiscB4ExtraScheme)
         ChkDiscB4SaleTax.Value = Abs(!isDiscB4SaleTax)
         TxtTradeOffer1.Text = IIf(IsNull(!TradeOffer1), 0, !TradeOffer1)
         TxtTradeOffer2.Text = IIf(IsNull(!TradeOffer2), 0, !TradeOffer2)
         TxtExtraSchemePer.Text = IIf(IsNull(!ExtraSchemePer), 0, !ExtraSchemePer)
         vIsRetailSaleTax = !IsRetailSaleTax
         TxtSaleTaxPer.Text = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
         TxtTokenVal.Text = IIf(IsNull(!TokenVal), "", !TokenVal)
         LblRetailPrice.Caption = !RetailPrice
         LblLastPrice.Caption = CN.Execute("Select dbo.FunLastPrice('S','" & DtpReturnDate.DateValue & "'," & Val(TxtProductID.Text) & ",'" & TxtCustomerID.Text & "')").Fields(0).Value
'         LblLastPrice.Caption = cn.Execute("Select dbo.FunLastPrice('S','" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text) & "," & Val(TxtCustomerID.Text) & ")").Fields(0).Value
         
         With CN.Execute("select cost from currentstock where productid = " & Val(TxtProductID.Text))
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
         ChkIsProduct.Value = 1
         If IsNull(!Packingname) Then
            vUnitPrice = IIf(isWholeSale = True, !WSPrice, !RetailPrice)
            vUnitRetailPrice = !RetailPrice
            TxtMultiplier.Text = ""
            CmbPackName.ListIndex = 0
'            Call CmbPackName_Click
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
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), "", !DiscPC)
         If vUnitPrice = 0 Then
            TxtDiscPer.Text = "0"
         Else
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
         End If
'         If ObjRegistry.AlertAllocateProduct = True Then
'            If Trim(TxtCustomerID.Text) <> "" Then
'               vStrSQL = "Select * " & vbCrLf _
'                     + " from CustomerProductPrice" & vbCrLf _
'                     + " where CustomerID = '" & TxtCustomerID.Text & "' and ProductID = '" & TxtProductID.Text & "'"
'
'               With cn.Execute(vStrSQL)
'                  If .RecordCount > 0 Then
'                     TxtPrice.Text = !price
'                     vUnitPrice = !price
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
                     + " from SaleHeader h inner join SaleBody b on h.BillID = b.BillID and h.BillDate = b.BillDate" & vbCrLf _
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
                           vUnitRetailPrice = 0 '!RetailPrice / !Multiplier
                        Else
                           vUnitPrice = !Price
                        End If
                        CmbPackName.Text = !Packingname
                     End If
                     TxtMultiplier.Text = IIf(IsNull(!Multiplier), "", !Multiplier)
                     TxtPrice.Text = !Price
                     TxtDiscPC.Text = !DiscPC
                     TxtDiscPer.Text = !DiscPer
                  End If
               End With
            End If
         End If
         If ObjRegistry.AutoApplyPartyLastDiscount Then
            If Val(TxtCustomerID.Text) <> 0 And Val(TxtCustomerID.Text) <> 621 Then
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
                  End If
               End With
            End If
         End If
         
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpReturnDate.DateValue + 1 & "',0),0)"
         vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         
'         With CN.Execute("select Productid, QtyLoose from CurrentStockStore where ProductID ='" & TxtProductID.Text & "' and StoreID = " & TxtStoreID.Text)
'            If .RecordCount > 0 Then
'               'vQtyLoose = !QtyLoose
'               LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack('" & !Productid & "',Floor(" & !QtyLoose & "))").Fields(0).Value
'               'LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'               LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose('" & !Productid & "',Floor(" & !QtyLoose & "))").Fields(0).Value
'               'LblStock.Caption = LblStock.Caption & " " & "Loose"
'            Else
'               'vQtyLoose = 0
'               LblStock.Caption = 0
'            End If
'         End With
'         If ObjRegistry.NegativeSale = False Then
'            If Val(LblStock.Caption) = 0 Then
'               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'               FunSelectProduct = False
'               Exit Function
'            End If
'         End If
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
         If TxtCustomerID.Text <> "" Then
            PopulateDataToHistoryGrid
            FrmHistory.Visible = True
            FrmHistory.ZOrder 0
            GridHistory.Visible = True
            GridHistory.ZOrder 0
         End If
         
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
      For vCounter = 2 To Grid.rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From saleReturnbody where SID=" & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Cleared Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
      
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eClearSavedRecord, TxtReturnID.Text, DtpReturnDate.DateValue, vGridRows & " Product/s Cleared ")
      Grid.Redraw = True
   FormStatus = NewMode
'   cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
     
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & Val(TxtReturnID.Text) & ",'" & DtpReturnDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      TxtBillNo.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
    ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSaleReturnInvoice", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
      vStrSQL = "select * from SaleReturnHeader where Tag is not null And SID=" & Val(TxtSID.Text) & " and Returndate='" & DtpReturnDate.DateValue & "'"
      With CN.Execute(vStrSQL)
          If Not .EOF Then
              MsgBox "Import/Export Record Cannot be Updated", vbInformation, Me.Caption
              Exit Sub
          End If
      End With
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   '''''''''''''''''''''''Check Import / Export'''''''''''''''''''''''''''''''''
    If ObjRegistry.ShowMultiBranches = True Then
      vStrSQL = "select * from SaleReturnHeader where Tag is not null And SID=" & Val(TxtSID.Text) & " and Returndate='" & DtpReturnDate.DateValue & "'"
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
   
   
'   vMaxBinID = FunGetMaxBinID
'   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_SaleReturnHeader Select " & vMaxBinID & ",'" & Date & "',* from SaleReturnHeader Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_SaleReturnBody Select " & vMaxBinID & ",'" & Date & "', * from SaleReturnBody Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Serial''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_SaleReturnSerial Select " & vMaxBinID & ",'" & Date & "', * from SaleReturnSerial Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin ProductOffer''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_SaleReturnOffer Select " & vMaxBinID & ",'" & Date & "', * from SaleReturnOffer Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")

   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  Call BinData
   Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eDelete, TxtReturnID.Text, DtpReturnDate.DateValue, Grid.rows - 1 & " Product/s Deleted Amount: " & Val(TxtNetAmount.Text))
   ''''''''''''''''''''''''''Delete Product OFfer'''''''''''''''''''''
   GridOffer.Redraw = False
   GridOffer.MoveFirst
   For vCounter = 1 To GridOffer.rows
      If Trim(GridOffer.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from SaleReturnOffer where ReturnID = " & Val(TxtReturnID.Text) & " And ReturnDate ='" & DtpReturnDate.DateValue & "' and productid ='" & GridOffer.Columns("Productid").Text & "'"
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
            CN.Execute "Delete from SaleReturnSerial where ReturnID = " & Val(TxtSID.Text) & " And ReturnDate ='" & DtpReturnDate.DateValue & "' and productid ='" & RsBodySerial!Productid & "' and Serial ='" & RsBodySerial!Serial & "'"
            RsBodySerial.MoveNext
        Next vCounter
   End If
   ''''''''''''''''''''''''''Delete Sale Return Body'''''''''''''''''''''
   Grid.Redraw = False
   Grid.MoveFirst
   Call ActivityLog("Sale Return Invoice", eDelete, TxtReturnID.Text, DtpReturnDate.DateValue)
   For vCounter = 1 To Grid.rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from SaleReturnBody where SID = " & Val(TxtSID.Text) & " And ReturnID = " & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "' and productid ='" & Grid.Columns("ProductID").Text & "' and BatchNo " & IIf(Trim(Grid.Columns("BatchNo").Text) = "", " is null", " = '" & Trim(Grid.Columns("BatchNo").Text) & "'") & " and Price = " & Val(Grid.Columns("Price").Text) & " and StoreID = " & Val(TxtStoreID.Text)
'         CN.Execute "Exec UpdateStockMinus " & TxtStoreID.Text & ",'" & Grid.Columns("ProductID").Text & "'," & Grid.Columns("Qty").Value
'          CN.Execute ("Insert Into Bin_SaleReturnBody Select " & FunGetMaxBinID & ", * from SaleReturnBody Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "' and productid ='" & Grid.Columns("Productid").Text & "'")
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   
    '''''''''''''''''''''''''''''''''''''''Delete Expense'''''''''''''''''''''''''''''''''''''''
   CN.Execute "Delete from SaleReturnExpense where ReturnID = " & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"

   CN.Execute "Delete from SaleReturnHeader where SID = " & Val(TxtSID.Text)
   
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnDelete Then
   vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = ObjRegistry.PrefixPhoneNo + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = ObjUserSecurity.UserName & " " & FrmSaleReturnInvoice.Caption & " Deleted ID:" & TxtReturnID.Text & vbCrLf & " Date:" & Format(DtpReturnDate.DateValue, "dd-MMM-yyyy") & " Time: " & Time & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt" & TxtNetAmount.Text
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
         RsBody!Productid = !Productid
         RsBody!Code = !Productid
         
         Grid.Columns("QtyLoose").Value = !QtyLoose
         Grid.Columns("Price").Value = !Price
         
         Grid.Columns("RetailPrice").Value = 0
         Grid.Columns("IsWSDiscb4ST").Value = 0
         Grid.Columns("IsWSSaleTax").Value = 0
         Grid.Columns("IsRetailSaleTax").Value = 0
         
         Grid.Columns("Amount").Value = (!Price * !QtyLoose)
         Grid.Columns("IsProduct").Value = 1
                  
         '''''
         RsBody!Multiplier = Null
         RsBody!QtyPack = Null
         RsBody!Qty = !QtyLoose
         RsBody!Bonus = Null
         RsBody!Price = !Price
         RsBody!isProduct = 1 '!isProduct
         
'         RsBody!RetailPrice = 0
'         RsBody!IsWSDiscb4ST = 0
'         RsBody!IsWSSaleTax = 0
'         RsBody!IsRetailSaleTax = 0
         
         RsBody!Cost = 0
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

Private Sub BtnReturnAll_Click()
   PopulateSaleDataToGrid
End Sub

Private Sub BtnSale_Click()
 On Error GoTo ErrorHandler
   If FunSelectSale(ssButton, False) = True Then
      BtnReturnAll.SetFocus
   Else
      TxtBillID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSale(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSale.ParaInBillDate = DtpBillDate.DateValue
        SchSale.Show vbModal, Me
        If SchSale.ParaOutBillID = "" Then FunSelectSale = False: Exit Function
        TxtSID.Text = SchSale.ParaOutSID
        TxtBillID.Text = SchSale.ParaOutBillID
        DtpBillDate.DateValue = SchSale.ParaOutBillDate
    End If
    '---------------------------
   vStrSQL = "select h.*, EmpName, OrganizationName, p.partyname, P.Address, P.City, c.AccountName, BankMachineName, StoreName, EmpName, MemberName FROM SaleHeader h left outer join parties p on h.CustomerID = p.partyid left outer join Organizations o on o.OrganizationID = h.OrganizationID left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join BankMachines b on b.BankMachineid = h.BankMachineid inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID left outer join Members M on M.MemberID = h.memberID where isReplace=0 and h.SID=" & Val(TxtSID.Text)
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerID.Text = !CustomerID
          TxtCustomerName.Text = !AccountName
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
          TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
          TxtTotalAmount.Text = !TotalAmount
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtTotalExpense.Text = IIf(IsNull(!TotalExpense), "", !TotalExpense)
'          TxtPaidAmount.Text = IIf(IsNull(!PAIDAMOUNT), "", !PAIDAMOUNT)
          TxtPreviousReceivable.Text = IIf(IsNull(!PreviousAmount), "", !PreviousAmount)
          lblPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
          LblTtlPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Total Receivable", "Total Payable")
          TxtPreviousReceivable.Text = Abs(Val(TxtPreviousReceivable.Text))
          FunSelectSale = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSale = False
          .Close
          TxtBillID.Text = ""
          TxtCustomerID.Text = ""
          TxtCustomerName.Text = ""
          TxtAddress.Text = ""
          TxtCity.Text = ""
          TxtOrganizationID.Text = ""
          TxtOrganizationName.Text = ""
          TxtEmployeeID.Text = ""
          TxtEmployeeName.Text = ""
          TxtBillNo.Text = ""
          TxtBiltyNo.Text = ""
          TxtVehicleNo.Text = ""
          TxtStoreID.Text = ""
          TxtStoreName.Text = ""
          TxtDescription.Text = ""
          TxtRemarks.Text = ""
          TxtTotalAmount.Text = ""
          TxtBillDiscPer.Text = ""
          TxtBillDisc.Text = ""
          TxtOtherCharges.Text = ""
          TxtTotalExpense.Text = ""
          TxtPreviousReceivable.Text = ""
          lblPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
          LblTtlPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Total Receivable", "Total Payable")
          TxtPreviousReceivable.Text = "0"
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
      .Close
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub PopulateSaleDataToGrid()
   
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleReturnBody where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
   ssql = "select p.productname, b.code, b.* from SaleBody b join products p on p.productid = b.productid where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "' order by serialno"
   With CN.Execute(ssql)
      If .RecordCount > 0 Then
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = "0"
         TxtTotalItems.Text = "0"
         vSSID = !SID
         While Not .EOF
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
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            Grid.Columns("QtyLoose").Value = !Qty
            Grid.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            Grid.Columns("Price").Value = !Price
            Grid.Columns("Cost").Value = !Cost
            Grid.Columns("isProduct").Value = !isProduct
            Grid.Columns("RetailPrice").Value = !RetailPrice
            
            Grid.Columns("isDiscB4TradeOffer").Value = IIf(IsNull(!isDiscB4TradeOffer), "", !isDiscB4TradeOffer)
            Grid.Columns("isDiscB4ExtraScheme").Value = IIf(IsNull(!isDiscB4ExtraScheme), "", !isDiscB4ExtraScheme)
            Grid.Columns("isDiscB4SaleTax").Value = IIf(IsNull(!isDiscB4SaleTax), "", !isDiscB4SaleTax)
            Grid.Columns("TradeOffer1").Value = IIf(IsNull(!TradeOffer1), "", !TradeOffer1)
            Grid.Columns("TradeOffer2").Value = IIf(IsNull(!TradeOffer2), "", !TradeOffer2)
            Grid.Columns("ExtraSchemePer").Value = IIf(IsNull(!ExtraSchemePer), "", !ExtraSchemePer)
            Grid.Columns("TradeValue").Value = IIf(IsNull(!TradeValue), "", !TradeValue)
            Grid.Columns("ExtraSchemeValue").Value = IIf(IsNull(!ExtraSchemeValue), "", !ExtraSchemeValue)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            
            Grid.Columns("IsWSDiscb4ST").Value = !IsWSDiscb4ST
            Grid.Columns("IsWSSaleTax").Value = !IsWSSaleTax
            Grid.Columns("IsRetailSaleTax").Value = !IsRetailSaleTax
            Grid.Columns("TokenVal").Value = IIf(IsNull(!TokenVal), "", !TokenVal)
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !Qty) 'IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = ((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !Qty) '!Amount
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !Qty))
            TxtTotalItems.Text = Val(TxtTotalItems.Text) + !Qty + IIf(IsNull(!Bonus), "0", !Bonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack))
            
            'TxtAmount.Text = Round((Val(vUnitPrice) - Val(TxtDiscPC.Text)) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 3)

            RsBody!Multiplier = !Multiplier
            RsBody!QtyPack = IIf(IsNull(!QtyPack), 0, !QtyPack)
            RsBody!Qty = !Qty
            RsBody!Bonus = IIf(IsNull(!Bonus), Null, !Bonus)
            RsBody!Price = !Price
            RsBody!Cost = !Cost
            RsBody!isProduct = 1 '!isProduct
'            RsBody!RetailPrice = !RetailPrice
'            RsBody!IsWSDiscb4ST = !IsWSDiscb4ST
'            RsBody!IsWSSaleTax = !IsWSSaleTax
'            RsBody!IsRetailSaleTax = !IsRetailSaleTax
'            RsBody!TokenVal = IIf(IsNull(!TokenVal), "", !TokenVal)
            RsBody!DiscPC = IIf(IsNull(!DiscPC), "", !DiscPC)
            RsBody!Offer = IIf(IsNull(!Offer), Null, !Offer)
            
            RsBody!isDiscB4TradeOffer = IIf(IsNull(!isDiscB4TradeOffer), "", !isDiscB4TradeOffer)
            RsBody!isDiscB4ExtraScheme = IIf(IsNull(!isDiscB4ExtraScheme), "", !isDiscB4ExtraScheme)
            RsBody!isDiscB4SaleTax = IIf(IsNull(!isDiscB4SaleTax), "", !isDiscB4SaleTax)
            RsBody!TradeOffer1 = IIf(IsNull(!TradeOffer1), "", !TradeOffer1)
            RsBody!TradeOffer2 = IIf(IsNull(!TradeOffer2), "", !TradeOffer2)
            RsBody!ExtraSchemePer = IIf(IsNull(!ExtraSchemePer), "", !ExtraSchemePer)
            RsBody!TradeValue = IIf(IsNull(!TradeValue), "", !TradeValue)
            RsBody!ExtraSchemeValue = IIf(IsNull(!ExtraSchemeValue), "", !ExtraSchemeValue)
            RsBody!SaleTaxPer = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            RsBody!SaleTaxval = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            
            RsBody!SaleTaxPer = IIf(IsNull(!SaleTaxPer), Null, !SaleTaxPer)
            RsBody!SaleTaxval = IIf(IsNull(!SaleTaxval), Null, !SaleTaxval)
            RsBody!DiscPer = IIf(IsNull(!DiscPer), "", !DiscPer)
            RsBody!DiscVal = Val(IIf(IsNull(!DiscPC), "0", !DiscPC)) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !Qty) 'IIf(IsNull(!DiscVal), "", !DiscVal)
            RsBody!Amount = ((!Price / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !Qty) '!Amount
            RsBody.Update
            .MoveNext
         Wend
         .Close
         Call SubCalculateBody
         Grid.AddNew
         Grid.Columns("productid").Text = " "
         Grid.AllowAddNew = False
         Grid.Redraw = True
         
         PopulateSaleDataToGridSerial
      End If
   End With
End Sub

Private Sub DtpReturnDate_LostFocus()
On Error GoTo ErrorHandler
 If Me.ActiveControl.Name <> DtpReturnDate.Name Then Exit Sub
    TxtReturnID.Text = FunGetMaxID()
    Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      LblCost.Visible = False
   End If
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
    vStrSQL = " Select c.AccountNo, c.AccountName as AccountName, Address, City, p.Description, p.Remarks, isnull(p.isWholeSale,1) as isWholeSale" & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " where c.AccountNo = " & Val(TxtCustomerID.Text) & " and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and isDetailed = 1 and isLocked = 0"
    
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerName.Text = !AccountName
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
          TxtPreviousReceivable.Text = CN.Execute("SELECT isnull(dbo.FunCurrentDebit(" & Val(TxtCustomerID.Text) & ",'" & DtpReturnDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
          vStrSQL = " Select isnull(Sum(round(B.TTLValue,0) - isnull(BillDisc,0) + isnull(OtherCharges,0) + Isnull(TotalExpense,0) + isnull(servicecharges,0) + isnull(STax,0)),0) as Amount " & vbCrLf _
                  + " FROM SaleReturnHeader h INNER JOIN (Select SID, Sum(Amount) TTLValue FROM SaleReturnBody Group By SID)b " & vbCrLf _
                  + " ON H.SID = B.SID " & vbCrLf _
                  + " where CustomerID = " & Val(TxtCustomerID.Text) & " and h.ReturnDate = '" & DtpReturnDate.DateValue & "' and h.ReturnID >= " & Val(TxtReturnID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
          TxtPreviousReceivable.Text = TxtPreviousReceivable.Text + CN.Execute(vStrSQL).Fields(0).Value
          lblPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
          TxtPreviousReceivable.Text = Abs(TxtPreviousReceivable.Text)
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
          TxtAddress.Text = ""
          TxtCity.Text = ""
          TxtDescription.Text = ""
          TxtRemarks.Text = ""
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
   Grid.MoveFirst
   ssql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With CN.Execute(ssql)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
            RsBody.Filter = "ProductID='" & !Productid & "'"
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
            Grid.Columns("DiscPC").Value = Round((Val(RsBody!Price) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) - Val(Grid.Columns("DiscVal").Value)
            
'            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - RsBody!Amount + Val(Grid.Columns("Amount").Text)
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
   Grid.MoveLast
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
         .Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
            
            RsBody.Filter = "ProductID='" & !Productid & "'"
            Grid.Columns("DiscPer").Value = 0 'IIf(IsNull(!DiscPer), 0, !DiscPer)
            Grid.Columns("DiscPC").Value = 0 'Round((Val(RsBody!Price) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = 0 'Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) - Val(Grid.Columns("DiscVal").Value)
            
'            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - RsBody!Amount + Val(Grid.Columns("Amount").Text)
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
   SchSaleReturn.ParaInReturnDate = DtpReturnDate.DateValue
   SchSaleReturn.Show vbModal
   If SchSaleReturn.ParaOutReturnID <> -1 Then
      TxtSID.Text = SchSaleReturn.ParaOutSID
      TxtReturnID.Text = SchSaleReturn.ParaOutReturnID
      'Dim a
      'a = Split(SchSaleReturn.ParaOutReturnDate, "/")
      DtpReturnDate.DateValue = SchSaleReturn.ParaOutReturnDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      CN.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetSaleReturn
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
      vStrSQL = "select h.ReturnID, h.ReturnDate, EntryDate, h.OrganizationID, OrganizationName, Customerid, isnull(Pr.PartyName,AccountName) + ' - ' + cast(H.CustomerID as varchar(10)) as Customer_Name_ID," & vbCrLf _
             + " pr.address, LicenceNo, SectorName, ZoneName, H.StoreID, StoreName, BiltyNo, VehicleNo, h.Description, h.Remarks," & vbCrLf _
             + " Isnull(H.BillDiscPer, 0) BillDiscPer, Isnull(H.BillDisc,0) BillDisc, isnull(OtherCharges,0) as OtherCharges,  isnull(h.ServiceCharges,0) as ServiceCharges," & vbCrLf _
             + " TotalAmount,  isnull(TotalExpense,0) as TotalExpense,  CompanyName, GroupName, SubGroupName, BrandName, SeasonName, b.ProductID as Code, p.ProductName as ProductName, '' Serial," & vbCrLf _
             + " '' ProductOffer, isnull(QtyPack,0)QtyPack, isnull(b.Multiplier,0)Multiplier, 0 as GrossQty, 0 as GrossUnit, Qty," & vbCrLf _
             + " P.RetailPrice, P.PurPrice, Bonus,b.DiscPc, b.DiscPer, DiscVal, Offer, 0 as TradeOffer_12, 0 as tradevalue, 0 as Extraschemevalue, 0 as ExtraSchemePer," & vbCrLf _
             + " b.SaleTaxPer, SaleTaxval, b.SC, h.Empid, empname, price, Amount, previousAmount, CashPaid, PaidAmount, isnull(BatchNo,'') as BatchNo, BillNo," & vbCrLf _
             + " Abbreviation + '/' + cast(b.Multiplier as varchar(10)) as packing, isnull(P.ListPrice,0) as ListPrice," & vbCrLf _
             + " isnull( pr.Phone1  + ', ','') + isnull( pr.Phone2 + ', ','')  + isnull( pr.mobile + ', ','') +  isnull( pr.mobile2 + ', ','') as Moblie, packingname, pr.city, " & vbCrLf _
             + " isnull(pr.Address,'') + isnull(' - ' + pr.City,'') + isnull(',' + pr.Phone1,'') + isnull(',' + pr.Phone2, '') + isnull(',' + pr.Mobile, '') as AddressFull, UserName, ReturnTime, isLastPrice"
             
      vStrSQL = vStrSQL + " from SaleReturnBody b inner join products p on b.productid = p.productid" & vbCrLf _
             + " inner join SaleReturnHeader h on H.SID = B.SID" & vbCrLf _
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
             + " where h.SID = " & Val(TxtSID.Text) & " and h.ReturnDate = '" & DtpReturnDate.DateValue & "'" & IIf(ObjRegistry.AllowOrderByCodeinInvoices, "Order By Code", "Order By SerialNo")
               
  
 
   If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
  
   If cmbPrintType.Text = "Half Page" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleReturnInvoiceHalf1.rpt")
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.x
      RptReportViewer.Report.RightMargin = 225
   ElseIf cmbPrintType.Text = "Thermal" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrptSaleReturnInvoiceAurora.rpt")
      RptReportViewer.Report.TopMargin = 0
      RptReportViewer.Report.LeftMargin = 0
      RptReportViewer.Report.RightMargin = 0
   Else
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrptSaleReturnInvoice.rpt")
      RptReportViewer.Report.LeftMargin = 225
      RptReportViewer.Report.RightMargin = 0
      RptReportViewer.Report.TopMargin = 255
   End If
   
   
   
   
 
   ''''''''
'   If ObjRegistry.LaserPrintofSaleInvoice = True Then
'      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleReturnInvoiceHalf1.rpt")
'      RptReportViewer.Report.TopMargin = ObjRegistry.Y
'      RptReportViewer.Report.LeftMargin = ObjRegistry.X
'      RptReportViewer.Report.RightMargin = 225
'   Else
'
'      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrptSaleReturnInvoice.rpt")
'      RptReportViewer.Report.PaperOrientation = crPortrait
'   End If
   
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ReportTitle = "Sale Return Invoice"
   
   If ObjRegistry.PrintHeadersSaleInvoice = True Then
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   Else
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ""
   End If
   
   If cmbPrintType.Text = "Half Page" Then
       RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy
       RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
       RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
       RptReportViewer.Report.ParameterFields(6).AddCurrentValue CStr(ObjRegistry.Statement)
   Else
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
   End If
   
'   If ObjRegistry.LaserPrintofSaleInvoice = True Then
'      RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
'      RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy
'       RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.CashReceived)
'       RptReportViewer.Report.ParameterFields(6).AddCurrentValue CStr(ObjRegistry.Statement)
'   Else
'      RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
'   End If

   
   
'      RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   
   
   vPrinter = Split(CmbPrinters.Text, ",")
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   
   
   If ObjRegistry.PreviewSaleInoice = True Or ChkIsPreview.Value = 1 Then
      If ChkIsPrint.Value = 1 Then
         RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
      End If
       RptReportViewer.Show vbModal, Me
   Else
      If ObjRegistry.IsPortrait = False Then RptReportViewer.Report.PaperOrientation = crLandscape
      If cmbPrintType.Text = "Half Page" Then RptReportViewer.Report.PaperOrientation = crLandscape
      RptReportViewer.Report.PrintOut False, CInt(IIf(IsNull(ObjRegistry.NoofPrints) Or Val(ObjRegistry.NoofPrints) = 0, 1, ObjRegistry.NoofPrints))
   End If
   
   'RptReportViewer.Report.ParameterFields(4).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
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
   
   ''''''''''''' User Discount ''''''''''''''
   If Val(ObjUserSecurity.AllowMaximmDiscPer) <> 0 Then
      If Val(TxtBillDiscPer.Text) > Val(ObjUserSecurity.AllowMaximmDiscPer) Then
         MsgBox "Discount greater than Fixed Discount", vbCritical, "Error"
         Exit Sub
      End If
   End If
  ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSaleReturnInvoice", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
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
   If Trim(TxtCustomerID.Text) = "" Then
      MsgBox "Enter Customer ID.", vbExclamation, Me.Caption
      TxtStoreID.SetFocus
      Exit Sub
   End If
   If CN.Execute("Select * From AdminClosing where ToUserNo = " & vUser & " and EntryDate = '" & DtpReturnDate.DateValue & "'").RecordCount > 0 Then
      MsgBox "You are not authorized to Add Record in Closing Dates.", vbCritical, "Alert"
      Exit Sub
   End If
   
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
   ssql = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where touserno = " & vUser & " and Entrydate <='" & Date & "'"
    With CN.Execute(ssql)
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
    '''''''''''''''''''''''Check Current Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.CurrentDateDataEntry = True And ObjUserSecurity.IsAdministrator = False Then
       If DtpReturnDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
    
    
   ''''''''''''''''''''''''Check Organization '''''''''''''''''''''''''''''''''
  If ObjRegistry.OrganizationMandatory = True And TxtOrganizationID.Text = "" Then
    MsgBox "Please Select Organization", vbInformation, Me.Caption
    If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
    Exit Sub
  End If
  
  
''''''''''''''''''''''''''''''''''
'   If FrmReturnPrint.OptCash.Visible Then FrmReturnPrint.OptCash.SetFocus
'   FrmReturnPrint.SubClearFields
'   FrmReturnPrint.TxtNetAmount.Text = TxtNetAmount.Text
'   FrmReturnPrint.Show vbModal, Me
'   If FrmReturnPrint.ParaOutSelection = False Then Exit Sub
   
   'Body Validation
   'validation has been performed when a row is added to the grid
   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Sale Return", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
   End If
   
   If Val(TxtBillID.Text) <> 0 Then
      RsBody.Filter = 0
      RsBody.MoveFirst
      For vCounter = 1 To RsBody.RecordCount
         With CN.Execute("select isnull(Qty,0) as Qty from salebody where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "' and ProductID = '" & RsBody!Productid & "'")
            If .RecordCount > 0 Then
               If !Qty < RsBody!Qty Then
                  MsgBox "Sale Quantity is less than Sale Return Quantity of Code = " & RsBody!Code, vbExclamation, "Alert"
                  If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
                  Exit Sub
               End If
            Else
               MsgBox "Sale of This Bill ID can not be found", vbExclamation, "Alert"
               If TxtBillID.Visible And TxtBillID.Enabled Then TxtBillID.SetFocus
               Exit Sub
            End If
         End With
         RsBody.MoveNext
      Next vCounter
   End If
   If DtpReturnDate.Date <> IIf(Format(Now, "hh") > 2, Date, DateAdd("d", -1, Date)) And DateFlag = True Then
      If MsgBox("Are you sure to Change Return Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
         DtpReturnDate.DateValue = IIf(Format(Now, "hh") > 2, Date, DateAdd("d", -1, Date))
      End If
      DateFlag = False
   End If
  
  'Saving record
  
   
    ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'Sale Return Invoice DIS' and LocalComputerName = '" & LocalComputerName & "'"
   If CN.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview, IsPrint ) Values ('" & LocalComputerName & "', 'Sale Return Invoice DIS','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & "," & ChkIsPrint.Value & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & ", IsPrint = " & ChkIsPrint.Value & " Where FormType = 'Sale Return Invoice DIS' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   CN.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   CN.BeginTrans
   
   If DtpReturnDate.Enabled Then
      If CN.Execute("Select * from SaleReturnHeader where ReturnID = " & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'").RecordCount > 0 Then
         'MsgBox "This Return ID already exists. A new Return ID. has been generated. Please try again", vbCritical, "Alert"
         TxtReturnID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
   
'   If vIsNewRecord = False Then Call ActivityLog("Sale Return Invoice", eEdit, TxtReturnID.Text, DtpReturnDate.DateValue)
'   Call UserActivities
   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Amount: " & Val(TxtNetAmount.Text))
   
   ssql = "select * from SaleReturnHeader where SID = " & Val(TxtSID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, CN, adOpenDynamic, adLockPessimistic
      If .BOF Then
         .AddNew
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
         !ReturnTime = Now
         !UserNo = vUser
      End If
      !isReplace = 0
      !isPosted = 0
      !isTransfer = 0
      !IsSync = 0
      !CustomerID = TxtCustomerID.Text
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
      !SalemanID = Null
      !EmpComm = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text))
      !BillID = IIf(Val(TxtBillID.Text) = 0, Null, Val(TxtBillID.Text))
      !BillDate = IIf(IsNull(!BillID), Null, DtpBillDate.DateValue)
      !StoreID = TxtStoreID.Text
      !TotalAmount = Round(Val(TxtTotalAmount.Text))
      !ServiceCharges = IIf(TxtServiceCharges.Text = "", Null, Val(TxtServiceCharges.Text))
      !ServiceChargesPer = IIf(TxtServiceChargesPer.Text = "", Null, Val(TxtServiceChargesPer.Text))
      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
      !Description = IIf(TxtDescription.Text = "", Null, TxtDescription.Text)
      !BiltyNo = IIf(TxtBiltyNo.Text = "", Null, TxtBiltyNo.Text)
      !BillNo = IIf(TxtBillNo.Text = "", Null, TxtBillNo.Text)
      !VehicleNo = IIf(TxtVehicleNo.Text = "", Null, TxtVehicleNo.Text)

'      If FrmReturnPrint.OptCash.Value = True Then
'         !CashPaid = IIf(FrmReturnPrint.TxtCashPaid.Text = "", Null, Val(FrmReturnPrint.TxtCashPaid.Text))
'         !CustomerID = "621"
'         !CustomerName = IIf(Trim(FrmReturnPrint.TxtCashCustomer.Text) = "", Null, FrmReturnPrint.TxtCashCustomer.Text)
'      End If
'      If FrmReturnPrint.OptCredit.Value = True Then
'         !CashPaid = 0
'         !CustomerID = FrmReturnPrint.TxtCustomerID.Text
'         !CustomerName = Null
'      End If
'      !BankCard = 0
      With CN.Execute("select dbo.DefaultValue('Cash Counter')")
         If TxtCustomerID.Text = .Fields(0).Value Then
            Rs!Cash = 1
            Rs!Credit = 0
            Rs!CashPaid = Round(Val(TxtTotalAmount.Text)) - Val(TxtBillDisc.Text)
         Else
            Rs!Cash = 0
            Rs!Credit = 1
            Rs!CashPaid = 0
         End If
      End With
      !BankCard = 0
      !PaidAmount = IIf(TxtPaidAmount.Text = "", Null, Val(TxtPaidAmount.Text))
      !TotalExpense = IIf(Val(TxtTotalExpense.Text) = 0, Null, Val(TxtTotalExpense.Text))
      !PreviousAmount = IIf(lblPayable.Caption = "Previous Receivable", Val(TxtPreviousReceivable.Text), Val(TxtPreviousReceivable.Text) * -1)
      !OtherCharges = IIf(Val(TxtOtherCharges.Text) = 0, Null, Val(TxtOtherCharges.Text))
      !ManualBillNo = IIf(Trim(TxtManualBillNo.Text) = "", "", TxtManualBillNo.Text)
      !Remarks = IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text)
      !Tag = IIf(Trim(TxtTag.Text) = "", Null, TxtTag.Text)
      
      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
      .Update
      .Close
      If vIsNewRecord = True Then TxtSID.Text = CN.Execute("select @@identity").Fields(0).Value
   End With
   vStrDetail = ""
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !SID = Val(TxtSID.Text)
         !ReturnID = Val(TxtReturnID.Text)
         !ReturnDate = DtpReturnDate.DateValue
         !StoreID = Val(TxtStoreID.Text)
          vStrDetail = vStrDetail & " (P" & !Productid & IIf(IsNull(!Multiplier), "", " M" & !Multiplier) & IIf(IsNull(!QtyPack), "", " QP" & !QtyPack) & IIf(IsNull(!Qty), "", " QL" & !Qty) & IIf(IsNull(!Bonus), "", " QB" & !Bonus) & " A" & !Amount & ")"
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   
   If RsBodySerial.RecordCount > 0 Then
     With RsBodySerial
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ReturnID = Val(TxtSID.Text)
         !ReturnDate = DtpReturnDate.DateValue
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
         !BillID = Val(TxtBillID.Text)
         !BillDate = DtpBillDate.DateValue
         .MoveNext
        Next vCounter
      End If
      .UpdateBatch
   End With
'   If vIsNewRecord = True Then Call ActivityLog("Sale Return Invoice", eAdd, TxtReturnID.Text, DtpReturnDate.DateValue)
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eAdd, TxtReturnID.Text, DtpReturnDate.DateValue, Grid.rows - 1 & " New Product/s Added Amount: " & Val(TxtNetAmount.Text))
   CN.CommitTrans
'   Char.Speak "Thank you for comming"
'   If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
'      Call BtnPrint_Click
'   End If
   
   If ChkIsPreview.Value = 1 Or ChkIsPrint.Value = 1 Then
      Call BtnPrint_Click
   End If
   
   '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = "Saved Return ID:" & TxtReturnID.Text & vbCrLf & " Date:" & Format(DtpReturnDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt:" & TxtNetAmount.Text
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               CN.Execute ssql
            End If
         Next
   End If
   
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
   RsBody.Open "Select * from SaleReturnBody where SID=" & Val(TxtSID.Text), CN, adOpenDynamic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.ProductName, code, b.* from SaleReturnBody b join products p on p.productid = b.productid where SID=" & Val(TxtSID.Text) & " Order by SerialNo asc "
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
            Grid.Columns("BatchNo").Text = IIf(IsNull(!BatchNo), "", !BatchNo)
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
            Grid.Columns("StampID").Text = !StampID
            Grid.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            Grid.Columns("QtyLoose").Value = !Qty
            Grid.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            Grid.Columns("Price").Value = !Price
            Grid.Columns("isLastPrice").Value = IIf(IsNull(!isLastPrice), "", !isLastPrice)
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("isDiscB4TradeOffer").Value = IIf(IsNull(!isDiscB4TradeOffer), "", !isDiscB4TradeOffer)
            Grid.Columns("isDiscB4ExtraScheme").Value = IIf(IsNull(!isDiscB4ExtraScheme), "", !isDiscB4ExtraScheme)
            Grid.Columns("isDiscB4SaleTax").Value = IIf(IsNull(!isDiscB4SaleTax), "", !isDiscB4SaleTax)
            Grid.Columns("TradeOffer1").Value = IIf(IsNull(!TradeOffer1), "", !TradeOffer1)
            Grid.Columns("TradeOffer2").Value = IIf(IsNull(!TradeOffer2), "", !TradeOffer2)
            Grid.Columns("ExtraSchemePer").Value = IIf(IsNull(!ExtraSchemePer), "", !ExtraSchemePer)
            Grid.Columns("TradeValue").Value = IIf(IsNull(!TradeValue), "", !TradeValue)
            Grid.Columns("ExtraSchemeValue").Value = IIf(IsNull(!ExtraSchemeValue), "", !ExtraSchemeValue)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(!Amount)
            TxtTotalItems.Text = Val(TxtTotalItems.Text) + !Qty + IIf(IsNull(!Bonus), "0", !Bonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack))
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
   RsBodySerial.Open "Select * from SaleReturnSerial where ReturnID=" & Val(TxtSID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
   Call PopulateDataToGridOffer
   Call PopulateDataToGridExpense
End Sub

Private Sub PopulateDataToGridExpense()
    If RsExpense.State = adStateOpen Then RsExpense.Close
    RsExpense.Open "Select * from SaleReturnExpense where Returnid =" & Val(TxtReturnID.Text) & " And ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
'    GridExpense.Visible = True
    ssql = "select EA.AccountNo, Accountname, SE.ExpAmount from ExpenseAccounts EA Left Outer join ChartofAccounts C on C.AccountNo = EA.AccountNo Left Outer Join (Select * from SaleReturnExpense where Returnid =" & Val(TxtReturnID.Text) & " And ReturnDate = '" & DtpReturnDate.DateValue & "') SE On SE.ExpenseID = EA.AccountNo"
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

Private Sub PopulateDataToGridserial()
   RsBodySerial.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
   If RsBodySerial.RecordCount > 0 Then
'       sSql = "select d.* from SaleReturnSerial d  where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "' and ProductID = '" & Grid.Columns("ProductID").Text & "'"
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
            GridSerial.Columns("Serial").Text = !Serial
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
           
      DtpReturnDate.DateValue = vDate
      
      DtpBillDate.DateValue = DtpReturnDate.DateValue
      TxtReturnID.Text = FunGetMaxID()
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
      LblStockCaption.Visible = False
      BtnProduct.Enabled = True
      'TxtReturnID.Enabled = True
      DtpReturnDate.Enabled = True
      If DtpReturnDate.Enabled And DtpReturnDate.Visible Then DtpReturnDate.SetFocus
      GridOffer.Visible = False
      FramExpense.ZOrder 0
      vIsNewRecord = True
      isWholeSale = True
   Case Is = OpenMode
      'TxtReturnID.Enabled = False
      DtpReturnDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      'TxtStoreID.Enabled = False
      'BtnStore.Enabled = False
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      'DtpReturnDate.SetFocus
      
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
      If TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus Else TxtCustomerID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub CmbPackName_Click()
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
         With CN.Execute("select * from ProductPacking where productid = " & Val(TxtProductID.Text) & " and packingid=" & CmbPackName.ItemData(CmbPackName.ListIndex))
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
   ElseIf ActiveControl.Name = Grid.Name And KeyCode = vbKeyF4 Then
      If Trim(Grid.Columns("ProductID").Text <> "") Then
         If MniCostPrice.Visible = True Then
'            Call MniCostPrice_Click
         End If
      End If
   ElseIf KeyCode = vbKeyF5 Then
      Select Case ActiveControl.Name
      Case TxtProductID.Name, CmbPackName.Name, TxtMultiplier.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, Grid.Name
            If Val(TxtMultiplier.Text) <> 0 Then
               LblLastPurPrice.Caption = CN.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text) & ")").Fields(0).Value * Val(TxtMultiplier.Text)
            Else
               LblLastPurPrice.Caption = CN.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "'," & Val(TxtProductID.Text) & ")").Fields(0).Value
            End If
'            LblLastPurPrice.Visible = True
            LblCost.Caption = LblLastPurPrice.Caption
            Call MniCostPrice_Click
      End Select
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
    RsProductOffer.Filter = "ProductID='" & GridOffer.Columns("ProductID").Text & "'"
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
   AddLabelEffect Me, 2, vbWhite, vbRed, lblEffectBorder
   SetWindowText Me.hWnd, "Sale Return Invoice"
   HelpLocation Me
   
   vSystemDate = Abs(ObjRegistry.SystemDate)
   vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
   
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
   ssql = "select * from FormDefaultSetting Where FormType = 'Sale Return Invoice DIS' and LocalComputerName = '" & LocalComputerName & "'"
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
   
   
   TxtStoreID.Text = IIf((ObjRegistry.StoreID = ""), "", ObjRegistry.StoreID)
   FunSelectStore ssValidate, True
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible

   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
         
   TxtBatchNo.Visible = ObjRegistry.BatchNoVisible
   
   LblEmpID.Visible = ObjRegistry.EmpVisible
   LblEmpName.Visible = ObjRegistry.EmpVisible
   TxtEmployeeID.Visible = ObjRegistry.EmpVisible
   TxtEmployeeName.Visible = ObjRegistry.EmpVisible
   BtnEmployee.Visible = ObjRegistry.EmpVisible
   
   LblMemberID.Visible = ObjRegistry.MemberVisible
   LblMemberName.Visible = ObjRegistry.MemberVisible
   TxtMemberID.Visible = ObjRegistry.MemberVisible
   TxtMemberName.Visible = ObjRegistry.MemberVisible
   BtnMember.Visible = ObjRegistry.MemberVisible
         
   TxtManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   LblManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   
   TxtRemarks.Visible = ObjRegistry.RemarksVisible
   LblRemarks.Visible = ObjRegistry.RemarksVisible
   
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
   
    
    LblAmount.Left = LblAmount.Left - TxtOffer.Width
    TxtAmount.Left = TxtAmount.Left - TxtOffer.Width
    
    Grid.Width = Grid.Width - TxtOffer.Width
   End If


   LblSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   TxtSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   
   If ObjRegistry.ShowSaleTax = False Then
    Grid.Columns("SaleTaxPer").Visible = False
        
    LblDiscPer.Left = LblDiscPer.Left - TxtSaleTaxPer.Width
    TxtDiscPer.Left = TxtDiscPer.Left - TxtSaleTaxPer.Width
    
    LblDiscVal.Left = LblDiscVal.Left - TxtSaleTaxPer.Width
    TxtDiscVal.Left = TxtDiscVal.Left - TxtSaleTaxPer.Width
    
    
    
    LblAmount.Left = LblAmount.Left - TxtSaleTaxPer.Width
    TxtAmount.Left = TxtAmount.Left - TxtSaleTaxPer.Width
    
    Grid.Width = Grid.Width - TxtSaleTaxPer.Width
   End If
   
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
      TxtPrice.Enabled = ObjUserSecurity.IsChangeRetail
      TxtPrice.Tag = IIf(TxtPrice.Enabled = True, "", "D")
   End If
'   DateFlag = True
   DateFlag = False
   
   With CN.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
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
   vServerDate = CN.Execute("Select CONVERT(datetime, CONVERT(varchar, GETDATE(), 110)) ServerDate").Fields(0).Value
   LblOtherChargesCaption.Caption = ObjRegistry.ChargesName
'   DateFlag = True
   DateFlag = False
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpReturnDate.IsDateValid = False Then Exit Function
   If ObjRegistry.AllowContinuousBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(ReturnID),0)+1 from SaleReturnHeader").Fields(0)
   ElseIf ObjRegistry.AllowMonthlyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(ReturnID),0)+1 from SaleReturnHeader where Month(ReturnDate) = '" & Month(DtpReturnDate.DateValue) & "' and  year(ReturnDate) ='" & Year(DtpReturnDate.DateValue) & "'").Fields(0)
   ElseIf ObjRegistry.AllowDailyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(ReturnID),0)+1 from SaleReturnHeader where ReturnDate = '" & DtpReturnDate.DateValue & "'").Fields(0)
   Else
      FunGetMaxID = CN.Execute("Select isnull(max(ReturnID),0)+1 from SaleReturnHeader where ReturnDate = '" & DtpReturnDate.DateValue & "' and StoreID = " & TxtStoreID.Text).Fields(0)
   End If
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpReturnDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = CN.Execute("Select isnull(max(BinID),0)+1 from Bin_SaleReturnHeader ").Fields(0)
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
   
   TxtNetAmount.Text = 0
   Grid.CancelUpdate
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
    Set FrmSaleReturnInvoice = Nothing
   End If
   ''''''''''''''''' ActivityLogBin For Close Action
'      Call DeleteTempActivityLogBin(vRandomID)
      If Grid.rows > 1 And Cancel = 0 Then
         vGridRows = 0
         Grid.Redraw = False
         Grid.MoveFirst
         For vCounter = 2 To Grid.rows
            vGridRows = vGridRows + 1
            If Trim(Grid.Columns("Code").Text) <> "" Then
               ssql = "Select Productid From saleReturnbody where SID=" & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("productid").Text)
               With CN.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Closed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eCloseSavedRecord, TxtReturnID.Text, DtpReturnDate.DateValue, vGridRows & " Product/s Closed")
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
      If Me.ActiveControl.Name = Grid.Name Then CmbPackName.SetFocus
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
    If Trim(Grid.Columns("Code").Text) = "" Then
'        TxtSerial.Enabled = False
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
   
   ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid='" & Grid.Columns("productid").Text & "'"
   With CN.Execute(ssql)
      If .EOF Then
         Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      Else
         Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eRemoveRow, TxtReturnID.Text, DtpReturnDate.DateValue, "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
         Call ActivityLogBin(vRandomID, eFrmSaleReturnInvoiceDIS, eAddTempRecord, TxtReturnID.Text, DtpReturnDate.DateValue, "Pending Remove Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      End If
   End With
   
   RsBody.Filter = "ProductID = " & Val(TxtProductID.Text) & IIf(ObjRegistry.BatchNoVisible = True, IIf(Trim(TxtBatchNo.Text) = "", "", " and BatchNo = '" & Trim(TxtBatchNo.Text) & "'"), "") & IIf(ObjRegistry.SeperateProductWithPrice = True, " and Price = " & Val(TxtPrice.Text), "")
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
'    cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed Code-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
    Grid.SelBookmarks.RemoveAll
    Grid.SelBookmarks.Add Grid.Bookmark
    Grid.DeleteSelected
    Grid.Refresh
    RsBody.Filter = 0
    Grid.MoveLast
    GetDataBackFromGridToTexBoxes
    SubClearSerialFields
   ElseIf Me.ActiveControl.Name = "GridSerial" Then
    If Trim(GridSerial.Columns("Serial").Text) = "" Then Exit Sub
    RsBodySerial.Filter = "Serial = '" & TxtSerial.Text & "'"
    If RsBodySerial.RecordCount > 0 Then RsBodySerial.Delete
    CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed Code-" & GridSerial.Columns("ProductID").Text & " Serial-" & GridSerial.Columns("Serial").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
    GridSerial.SelBookmarks.RemoveAll
    GridSerial.SelBookmarks.Add GridSerial.Bookmark
    GridSerial.DeleteSelected
    GridSerial.Refresh
    RsBodySerial.Filter = 0
'    GridSerial.MoveLast
    GetDataBackFromGridSerialToTexBoxes
   ElseIf Me.ActiveControl.Name = GridOffer.Name Then
    RsProductOffer.Filter = "ProductID='" & GridOffer.Columns("ProductID").Text & "'"
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
   If GridSerial.rows = 1 Then GridSerial.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
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
   FrmHistory.Visible = False
   FrmProductPrices.Visible = False
   DtpReturnDate.Refresh
On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Productid").Text) = "" Then
      RsBody.Filter = "ProductID = " & Val(TxtProductID.Text) & IIf(ObjRegistry.BatchNoVisible = True, IIf(Trim(TxtBatchNo.Text) = "", "", " and BatchNo = '" & Trim(TxtBatchNo.Text) & "'"), "") & IIf(ObjRegistry.SeperateProductWithPrice = True, " and Price = " & Val(TxtPrice.Text), "")
   Else
'      VStrSQL = "Select StampID from salereturnbody where sid = " & Val(TxtSID.Text) & " and ProductID = '" & Grid.Columns("Productid").Text & "'" & IIf(ObjRegistry.BatchNoVisible = True, IIf(Trim(Grid.Columns("BatchNo").Text) = "", "", " and BatchNo = '" & Trim(Grid.Columns("BatchNo").Text) & "'"), "") & IIf(ObjRegistry.SeperateProductWithPrice = True, " and Price = " & Val(Grid.Columns("Price").Text), "")
'      vStampID cn.Execute(VStrSQL).Fields(0).Value
      RsBody.Filter = "ProductID = " & Val(Grid.Columns("Productid").Text) & IIf(ObjRegistry.BatchNoVisible = True, IIf(Trim(Grid.Columns("BatchNo").Text) = "", "", " and BatchNo = '" & Trim(Grid.Columns("BatchNo").Text) & "'"), "") & IIf(ObjRegistry.SeperateProductWithPrice = True, " and Price = " & Val(Grid.Columns("Price").Text), "")
   End If
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("Serial").Text = Grid.rows
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
            For vrowcounter = 1 To Grid.rows
               If Grid.Columns("Productid").Text = TxtProductID.Text And IIf(ObjRegistry.BatchNoVisible = True, Grid.Columns("BatchNo").Text = Trim(TxtBatchNo.Text), True) And IIf(ObjRegistry.SeperateProductWithPrice = True, Val(Grid.Columns("Price").Text) = Val(TxtPrice.Text), True) Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid='" & Grid.Columns("Code").Text & "'"
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  
                  
                  '''''''''''''''''''''''''This QtyOffer Is used for DetailGrid
                  QtyOffer = Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value) + Val(Grid.Columns("QtyLoose").Value)
                  GetDataFromTextBoxesToGridOffer
                  TxtOffer.Text = Val(TxtOffer.Text) + Val(Grid.Columns("Offer").Text)
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  
                  TxtQtyLoose.Text = Val(TxtQtyLoose.Text) + Val(Grid.Columns("QtyLoose").Value)
                  TxtQtyPack.Text = Val(TxtQtyPack.Text) + Val(Grid.Columns("QtyPack").Value)
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Val(Grid.Columns("QtyLoose").Value) + Val(Grid.Columns("Bonus").Value) + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
                  TxtTradeOfferValue.Text = Val(TxtTradeOfferValue.Text) + Val(TxtTradeOfferValue.Text) - Val(Grid.Columns("TradeValue").Text)
                  TxtExtraSchemeValue.Text = Val(TxtExtraSchemeValue.Text) + Val(TxtExtraSchemeValue.Text) - Val(Grid.Columns("ExtraSchemeValue").Text)
                  Grid.Columns("isDiscB4TradeOffer").Value = Abs(ChkDiscB4TradeOffer.Value)
                  Grid.Columns("isDiscB4ExtraScheme").Value = Abs(ChkDiscB4ExtraScheme.Value)
                  Grid.Columns("isDiscB4SaleTax").Value = Abs(ChkDiscB4SaleTax.Value)
                  Grid.Columns("TradeOffer1").Value = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
                  Grid.Columns("TradeOffer2").Value = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
                  Grid.Columns("ExtraSchemePer").Value = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
                  Grid.Columns("TradeValue").Value = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
                  Grid.Columns("ExtraSchemeValue").Value = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
                  Grid.Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
                  Grid.Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, 0, Val(TxtMultiplier.Text))
                  Grid.Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
                  Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
                  Grid.Columns("Bonus").Value = Val(TxtBonus.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("isLastPrice").Value = Abs(IIf(Val(LblLastPrice.Caption) = Val(TxtPrice.Text), 1, 0))
                  Grid.Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  Grid.Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  Grid.Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  Grid.Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  Grid.Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  Grid.Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("Cost").Value = Val(TxtCost.Text)
                  Grid.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
                  
                  RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
                  RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
                  RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
                  RsBody!StoreID = Val(TxtStoreID.Text)
                  RsBody!HeaderStoreID = Val(TxtStoreID.Text)
                  RsBody!Qty = Val(TxtQtyLoose.Text)
                  RsBody!Bonus = Val(TxtBonus.Text)
                  RsBody!Price = Val(TxtPrice.Text)
                  RsBody!isLastPrice = Abs(IIf(Val(LblLastPrice.Caption) = Val(TxtPrice.Text), 1, 0))
                  RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  RsBody!isDiscB4TradeOffer = Abs(ChkDiscB4TradeOffer.Value)
                  RsBody!isDiscB4ExtraScheme = Abs(ChkDiscB4ExtraScheme.Value)
                  RsBody!isDiscB4SaleTax = Abs(ChkDiscB4SaleTax.Value)
                  RsBody!TradeOffer1 = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
                  RsBody!TradeOffer2 = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
                  RsBody!ExtraSchemePer = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
                  RsBody!TradeValue = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
                  RsBody!ExtraSchemeValue = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
                  RsBody!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  RsBody!SaleTaxval = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  RsBody!Amount = Val(TxtAmount.Text)
                  RsBody!Cost = Val(TxtCost.Text)
                  RsBody!isProduct = 1 'Abs(Grid.Columns("isProduct").Value)
                  
                  ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid='" & Grid.Columns("Code").Text & "'"
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmSaleReturnInvoiceDIS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Pending Update Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                 
                  Grid.MoveLast

                  Call SubClearDetailArea
                  TxtCode.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vrowcounter
'            Grid.Columns("Serial").Text = Grid.rows
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearDetailArea
         Grid.MoveLast
         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   
   Grid.Redraw = False
   With Grid
      If TxtCode.Enabled = True Then
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text)))
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eAddNewRowByEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
         Call ActivityLogBin(vRandomID, eFrmSaleReturnInvoiceDIS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Pending Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Val(Grid.Columns("Pack").Value)) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
          ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid='" & Grid.Columns("ProductID").Text & "'"
         With CN.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            Else
               Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eEdit, TxtReturnID.Text, DtpReturnDate.Date, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmSaleReturnInvoiceDIS, eEdit, TxtReturnID.Text, DtpReturnDate.Date, "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmSaleReturnInvoiceDIS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Pending Update Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
      End If
      
      .Columns("BatchNo").Text = Trim(TxtBatchNo.Text)
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
      .Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, 0, Val(TxtMultiplier.Text))
      .Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, 0, Val(TxtQtyPack.Text))
      .Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
      .Columns("Bonus").Value = Val(TxtBonus.Text)
      .Columns("Price").Value = Val(TxtPrice.Text)
      
      .Columns("isLastPrice").Value = Abs(IIf(Val(LblLastPrice.Caption) = Val(TxtPrice.Text), 1, 0))
      .Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      .Columns("isDiscB4TradeOffer").Value = Abs(ChkDiscB4TradeOffer.Value)
      .Columns("isDiscB4ExtraScheme").Value = Abs(ChkDiscB4ExtraScheme.Value)
      .Columns("isDiscB4SaleTax").Value = Abs(ChkDiscB4SaleTax.Value)
      .Columns("TradeOffer1").Value = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
      .Columns("TradeOffer2").Value = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
      .Columns("ExtraSchemePer").Value = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
      .Columns("TradeValue").Value = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
      .Columns("ExtraSchemeValue").Value = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
      .Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
      .Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
      .Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      .Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      .Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("Cost").Value = Val(TxtCost.Text)
      .Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
      RsBody!BatchNo = IIf(Trim(TxtBatchNo.Text) = "", Null, Trim(TxtBatchNo.Text))
      RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
      RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
      RsBody!StoreID = Val(TxtStoreID.Text)
      RsBody!HeaderStoreID = Val(TxtStoreID.Text)
      RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
      RsBody!Qty = Val(TxtQtyLoose.Text)
      RsBody!Bonus = Val(TxtBonus.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!isLastPrice = Abs(IIf(Val(LblLastPrice.Caption) = Val(TxtPrice.Text), 1, 0))
      RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      RsBody!isDiscB4TradeOffer = Abs(ChkDiscB4TradeOffer.Value)
      RsBody!isDiscB4ExtraScheme = Abs(ChkDiscB4ExtraScheme.Value)
      RsBody!isDiscB4SaleTax = Abs(ChkDiscB4SaleTax.Value)
      RsBody!TradeOffer1 = IIf(Val(TxtTradeOffer1.Text) = 0, 0, Val(TxtTradeOffer1.Text))
      RsBody!TradeOffer2 = IIf(Val(TxtTradeOffer2.Text) = 0, 0, Val(TxtTradeOffer2.Text))
      RsBody!ExtraSchemePer = IIf(Val(TxtExtraSchemePer.Text) = 0, 0, Val(TxtExtraSchemePer.Text))
      RsBody!TradeValue = IIf(Val(TxtTradeOfferValue.Text) = 0, 0, Val(TxtTradeOfferValue.Text))
      RsBody!ExtraSchemeValue = IIf(Val(TxtExtraSchemeValue.Text) = 0, 0, Val(TxtExtraSchemeValue.Text))
      RsBody!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
      RsBody!SaleTaxval = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
      RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      RsBody!Amount = Val(TxtAmount.Text)
      RsBody!Cost = Val(TxtCost.Text)
      RsBody!isProduct = 1 'Abs(Grid.Columns("isProduct").Value)
      .MoveLast
      
      If Trim(.Columns("Code").Text) <> "" Then
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
   TxtCode.Enabled = True
   BtnProduct.Enabled = True
   TxtCode.Text = ""
   TxtProductName.Text = ""
   TxtBatchNo.Text = ""
   CmbPackName.ListIndex = 0
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
                        For vCounter = 1 To GridOffer.rows
                        If GridOffer.Columns("ProductID").Text = TxtProductID.Text Then
                            GridOffer.Columns("ProductName").Text = CN.Execute("Select ProductName from products where productid = '" & GridOffer.Columns("ProductOfferID").Text & "'").Fields(0)
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
                    GridOffer.Columns("ProductName").Text = CN.Execute("Select ProductName from products where productid = '" & GridOffer.Columns("ProductOfferID").Text & "'").Fields(0)
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
   Exit Sub
ErrorHandler:
   GridOffer.Redraw = True
   Call ShowErrorMessage
End Sub


Private Sub PopulateDataToGridOffer()
    If RsProductOffer.State = adStateOpen Then RsProductOffer.Close
    RsProductOffer.Open "Select * from SaleReturnOffer where ReturnID =" & Val(TxtReturnID.Text) & " And ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
    If RsProductOffer.RecordCount > 0 Then
    GridOffer.Visible = True
    ssql = "select p.productname, D.* from SaleReturnOffer D Inner join products p on p.productid = D.productOfferid where ReturnID =" & Val(TxtReturnID.Text) & " And ReturnDate = '" & DtpReturnDate.DateValue & "'"
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

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("Code").Text
      TxtBatchNo.Text = .Columns("BatchNo").Text
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
      TxtQtyLoose.Text = .Columns("QtyLoose").Text
      TxtQtyPack.Text = .Columns("QtyPack").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtBonus.Text = .Columns("Bonus").Text
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtOffer.Text = .Columns("Offer").Value
      ChkDiscB4TradeOffer.Value = Abs(Val(.Columns("isDiscB4TradeOffer").Value))
      ChkDiscB4ExtraScheme.Value = Abs(Val(.Columns("isDiscB4ExtraScheme").Value))
      ChkDiscB4SaleTax.Value = Abs(Val(.Columns("isDiscB4SaleTax").Value))
      TxtTradeOffer1.Text = .Columns("TradeOffer1").Value
      TxtTradeOffer2.Text = .Columns("TradeOffer2").Value
      TxtExtraSchemePer.Text = .Columns("ExtraSchemePer").Value
      TxtTradeOfferValue.Text = .Columns("TradeValue").Value
      TxtExtraSchemeValue.Text = .Columns("ExtraSchemeValue").Value
      TxtSaleTaxPer.Text = .Columns("SaleTaxPer").Value
      TxtSaleTaxVal.Text = .Columns("SaleTaxVal").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      TxtAmount.Text = .Columns("Amount").Value
      TxtCost.Text = .Columns("Cost").Value
      ChkIsProduct.Value = Abs(.Columns("isProduct").Value)
      If ObjRegistry.ShowAllPrices Then
         PopulateDataToPriceGrid
         FrmProductPrices.Visible = True
      End If

      If Val(TxtMultiplier.Text) = 0 Then
         vUnitPrice = IIf(.Columns("Price").Text = "", 0, .Columns("Price").Text)
         vUnitRetailPrice = IIf(.Columns("RetailPrice").Text = "", 0, .Columns("RetailPrice").Text)
      Else
         vUnitPrice = Val(.Columns("Price").Text) / Val(TxtMultiplier.Text)
         vUnitRetailPrice = Val(.Columns("RetailPrice").Text) / Val(TxtMultiplier.Text)
      End If
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpReturnDate.DateValue + 1 & "',0),0)"
         vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         
         LblLastPrice.Caption = CN.Execute("Select dbo.FunLastPrice('S','" & DtpReturnDate.DateValue & "'," & Val(TxtProductID.Text) & ",'" & TxtCustomerID.Text & "')").Fields(0).Value

'         LblCaptionRetailPrice.Visible = True
'         LblRetailPrice.Visible = True
   End With
   If Grid.rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetSaleReturn()
   On Error GoTo ErrorHandler
   ssql = "select h.*, OrganizationName, c.AccountName, P.Address, P.City, StoreName, EmpName, MemberName FROM SaleReturnHeader h left outer join parties p on h.CustomerID = p.partyid left outer join Organizations o on o.OrganizationID = h.OrganizationID left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join Members m on m.MemberID = h.MemberID inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID where isReplace=0 and h.SID=" & Val(TxtSID.Text) & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With CN.Execute(ssql)
      If Not .BOF Then
          DtpReturnDate.DateValue = !ReturnDate
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          TxtCustomerID.Text = !CustomerID
          TxtCustomerName.Text = !AccountName
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
          TxtEmployeeName.Text = IIf(IsNull(!empname), "", !empname)
          TxtBillNo.Text = IIf(IsNull(!BillNo), "", !BillNo)
          TxtBiltyNo.Text = IIf(IsNull(!BiltyNo), "", !BiltyNo)
          TxtVehicleNo.Text = IIf(IsNull(!VehicleNo), "", !VehicleNo)
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
          TxtTotalAmount.Text = !TotalAmount
          TxtServiceChargesPer.Text = IIf(IsNull(!ServiceChargesPer), "", !ServiceChargesPer)
          TxtServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
          TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
          TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtTotalExpense.Text = IIf(IsNull(!TotalExpense), "", !TotalExpense)
          TxtPaidAmount.Text = IIf(IsNull(!PaidAmount), "", !PaidAmount)
          TxtPreviousReceivable.Text = IIf(IsNull(!PreviousAmount), "", !PreviousAmount)
          lblPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
          LblTtlPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Total Receivable", "Total Payable")
          TxtPreviousReceivable.Text = Abs(Val(TxtPreviousReceivable.Text))
          vZoneID = CN.Execute("SELECT isnull(dbo.FunGetZoneID(" & Val(TxtCustomerID.Text) & "),0)").Fields(0).Value
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
   TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / Val(TxtTotalAmount.Text), 2)
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
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 2)
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
   TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 3)
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
   TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), 2)
   'TxtAmount.Text = Round((Val(vUnitPrice) - Val(TxtDiscPC.Text)) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 2)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDiscVal_LostFocus()
   Select Case ActiveControl.Name
   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtBonus.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtOffer.Name, TxtSaleTaxPer.Name
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
   Call SubCalculateFooter
End Sub

Private Sub TxtPreviousReceivable_Change()
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

Private Sub TxtQtyLoose_Validate(Cancel As Boolean)
If ActiveControl.Name <> TxtQtyLoose.Name Then Exit Sub
End Sub

Private Sub TxtQtyPack_Change()
   Call SubCalculateBody
   Call FindRebate
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

Private Sub TxtSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridSerial.SetFocus
End Sub

Private Sub TxtSerial_LostFocus()
'    GetDataFromTexBoxesToGridSerial
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
'   RsBodySerial.Filter = "ProductID ='" & Grid.Columns("ProductID").Text & "' And Serial='" & TxtSerial.Text & "'"
'
'         GridSerial.Redraw = False
'         GridSerial.MoveFirst
'            For vrowcounter = 1 To GridSerial.rows
'               If GridSerial.Columns("Serial").Text = TxtSerial.Text Then
'                  MsgBox "The Product cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
''                  vAlreadySerial = True
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
'         RsBodySerial.AddNew
'          With cn.Execute("Select Distinct ProductID from vuPurchaseSerial where Serial = '" & Trim(TxtSerial.Text) & "'")
'            If .EOF Then Exit Sub
'            TxtCode.Text = .Fields("ProductId").Value
'            GridSerial.Columns("ProductID").Text = TxtCode.Text
'            GridSerial.Columns("Serial").Text = TxtSerial.Text
'            RsBodySerial!ProductID = TxtCode.Text
'            RsBodySerial!Serial = TxtSerial.Text
'          If FunSelectProduct(ssValidate, False) = True Then GetDataFromTexBoxesToGrid
'          End With
'
'
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
'   Frame1.Visible = True
'   GridSerial.Redraw = True
   Exit Sub
ErrorHandler:
   GridSerial.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtServiceCharges_Change()
On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtServiceCharges.Name Then Exit Sub
   TxtServiceChargesPer.Text = Round((Val(TxtServiceCharges.Text) * 100) / Val(TxtTotalAmount.Text), 2)
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
Private Sub UserActivities()
     If vIsNewRecord = False Then
    With CN.Execute("Select  * from SaleReturnHeader where ReturnID =" & TxtReturnID.Text & " And ReturnDate = '" & DtpReturnDate.DateValue & "'")
       If Not .EOF Then
         If TxtStoreID.Text <> !StoreID Then
            CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
         End If
       End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.rows - 1
        With CN.Execute("Select * from SaleReturnBody Where ReturnID = " & TxtReturnID.Text & " and ReturnDate ='" & DtpReturnDate.DateValue & "' and Productid ='" & Grid.Columns("Productid").Text & "'")
        
             If .EOF = True Then
                ssql = "Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Inserted New Code-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")"
                CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Inserted New Code-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("QtyLoose").Text <> !Qty Or Grid.Columns("Price").Text <> !Price Or Grid.Columns("discper").Text <> !DiscPer Then
                   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated Code-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
    
   Else
    CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub TxtTotalExpense_Change()
  Call SubCalculateFooter
End Sub

Private Sub MniCostPrice_Click()
   On Error GoTo ErrorHandler
'   If Trim(Grid.Columns("Cost").Text) = "" Then Exit Sub
   If ObjUserSecurity.ShowPurchasePriceInInvoice = True Or ObjUserSecurity.IsAdministrator = True Then
'      LblCost.Caption = Grid.Columns("Cost").Value
      LblCost.Visible = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToHistoryGrid()
'      ssql = "select top 3 pt.PartyName, CustomerID, code, b.* " & vbCrLf & _
'      " from SaleReturnHeader h inner join SaleReturnbody b on H.SID = B.SID" & vbCrLf & _
'      " inner join Parties pt on pt.PartyID = h.CustomerID " & vbCrLf & _
'      " where h.CustomerID = '" & TxtCustomerID.Text & "' and b.productid = '" & (TxtProductID.Text) & "' order by b.ReturnDate Desc"

       ssql = "select top 3 pt.PartyName, CustomerID, code, b.* " & vbCrLf & _
      " from SaleHeader h inner join Salebody b on H.SID = B.SID and h.billdate = b.billdate" & vbCrLf & _
      " inner join Parties pt on pt.PartyID = h.CustomerID " & vbCrLf & _
      " where h.CustomerID = " & Val(TxtCustomerID.Text) & " and b.productid = " & Val(TxtProductID.Text) & " order by b.BillDate Desc"
      
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


Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleReturnHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmSaleReturnInvoiceDIS) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmSaleReturnInvoiceDIS & ", " & vUser & "," & TableHeaderFields(eFrmSaleReturnInvoiceDIS) & " from SaleReturnHeader " & vbCrLf _
             & "Where SID = " & TxtSID.Text
      CN.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleReturnBodyBin (" & TableBodyFields(eFrmSaleReturnInvoiceDIS) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmSaleReturnInvoiceDIS) & " from SaleReturnBody " & vbCrLf _
             & "Where SID = " & TxtSID.Text
      CN.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub PopulateSaleDataToGridSerial()
   On Error GoTo ErrorHandler
   
   RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   RsBodySerial.Open "Select * from SaleReturnSerial where ReturnID=" & Val(TxtSID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
   ssql = "Select * from SaleBodySerial where BillID=" & Val(vSSID) & " and BillDate='" & DtpBillDate.DateValue & "'"
   With CN.Execute(ssql)
      If .RecordCount > 0 Then
         GridSerial.Redraw = False
         GridSerial.MoveFirst
         GridSerial.RemoveAll
         GridSerial.AllowAddNew = True
         While Not .EOF
            RsBodySerial.AddNew
            RsBodySerial!Productid = !Productid
            RsBodySerial!Serial = !Serial
            RsBodySerial.Update
            GridSerial.AddNew
            GridSerial.Columns("ProductID").Text = !Productid
            GridSerial.Columns("Serial").Text = !Serial
            .MoveNext
         Wend
         .Close
         GridSerial.AddNew
         GridSerial.Columns("productid").Text = " "
         GridSerial.AllowAddNew = False
         GridSerial.Redraw = True
      End If
   End With
   Exit Sub
ErrorHandler:
   GridSerial.Redraw = True
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
TxtAmount.Text = Round(Val(TxtAmount.Text) - Val(TxtDiscVal.Text) - Val(TxtTradeOfferValue.Text) - Val(TxtExtraSchemeValue.Text) + Val(TxtSaleTaxVal.Text) - Val(TxtOffer.Text), 2)
If ObjRegistry.IsRoundFigure = True Then TxtAmount.Text = SelfRound(TxtAmount.Text)
'TxtDiscAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - Val(TxtDiscVal.Text) - Val(TxtExtraSchemeValue.Text), 2)
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
