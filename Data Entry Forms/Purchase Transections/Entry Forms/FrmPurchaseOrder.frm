VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmPurchaseOrder 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmPurchaseOrder.frx":0000
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
      Left            =   13635
      TabIndex        =   141
      Top             =   10395
      Width           =   1290
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Print"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   140
      Top             =   0
      Width           =   1290
   End
   Begin VB.Frame FrameMultiBranchStock 
      Height          =   1275
      Left            =   360
      TabIndex        =   136
      Top             =   585
      Visible         =   0   'False
      Width           =   7080
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridBranch 
         Height          =   1140
         Left            =   45
         TabIndex        =   137
         Top             =   45
         Width           =   6960
         ScrollBars      =   1
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   10
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
         stylesets(0).Picture=   "FrmPurchaseOrder.frx":000C
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
         stylesets(1).Picture=   "FrmPurchaseOrder.frx":0028
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
         stylesets(2).Picture=   "FrmPurchaseOrder.frx":0044
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
         Columns.Count   =   10
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "Branch1"
         Columns(0).Name =   "Branch1"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Branch2"
         Columns(1).Name =   "Branch2"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "Branch3"
         Columns(2).Name =   "Branch3"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "Branch4"
         Columns(3).Name =   "Branch4"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "Branch5"
         Columns(4).Name =   "Branch5"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "Branch6"
         Columns(5).Name =   "Branch6"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Visible=   0   'False
         Columns(6).Caption=   "Branch7"
         Columns(6).Name =   "Branch7"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3200
         Columns(7).Visible=   0   'False
         Columns(7).Caption=   "Branch8"
         Columns(7).Name =   "Branch8"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   3200
         Columns(8).Visible=   0   'False
         Columns(8).Caption=   "Branch9"
         Columns(8).Name =   "Branch9"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3200
         Columns(9).Caption=   "Stock"
         Columns(9).Name =   "Stock"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   12277
         _ExtentY        =   2011
         _StockProps     =   79
         Caption         =   "Branch Stock"
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
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13635
      TabIndex        =   135
      Top             =   10080
      Width           =   1245
   End
   Begin VB.Frame FrmProductPrices 
      Height          =   1095
      Left            =   7605
      TabIndex        =   132
      Top             =   450
      Visible         =   0   'False
      Width           =   6270
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridProductPrices 
         Height          =   885
         Left            =   60
         TabIndex        =   133
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
         stylesets(0).Picture=   "FrmPurchaseOrder.frx":0060
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
         stylesets(1).Picture=   "FrmPurchaseOrder.frx":007C
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
         stylesets(2).Picture=   "FrmPurchaseOrder.frx":0098
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
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   11730
      TabIndex        =   128
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   9705
      Width           =   1170
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmPurchaseOrder.frx":00B4
      Left            =   10350
      List            =   "FrmPurchaseOrder.frx":00B6
      Style           =   2  'Dropdown List
      TabIndex        =   127
      Tag             =   "1"
      Top             =   10155
      Width           =   3276
   End
   Begin VB.ComboBox cmbSizeName 
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   123
      Top             =   4545
      Width           =   840
   End
   Begin VB.ComboBox CmbColourName 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   121
      Top             =   4545
      Width           =   1200
   End
   Begin VB.Frame FrmHistory 
      Height          =   1635
      Left            =   3465
      TabIndex        =   117
      Top             =   6278
      Visible         =   0   'False
      Width           =   9855
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridHistory 
         Height          =   1455
         Left            =   90
         TabIndex        =   118
         Top             =   135
         Width           =   10020
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   13
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
         stylesets(0).Picture=   "FrmPurchaseOrder.frx":00B8
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
         stylesets(1).Picture=   "FrmPurchaseOrder.frx":00D4
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
         stylesets(2).Picture=   "FrmPurchaseOrder.frx":00F0
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
         Columns.Count   =   13
         Columns(0).Width=   953
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
         Columns(2).Width=   1799
         Columns(2).Caption=   "Date"
         Columns(2).Name =   "Date"
         Columns(2).CaptionAlignment=   2
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).NumberFormat=   "dd/MM/yyyy"
         Columns(2).FieldLen=   256
         Columns(3).Width=   1852
         Columns(3).Caption=   "PackName"
         Columns(3).Name =   "Packname"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   847
         Columns(4).Caption=   "Pack"
         Columns(4).Name =   "Pack"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   1005
         Columns(5).Caption=   "Qty(P)"
         Columns(5).Name =   "QtyPack"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   1032
         Columns(6).Caption=   "Qty(L)"
         Columns(6).Name =   "QtyLoose"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   688
         Columns(7).Caption=   "Bns"
         Columns(7).Name =   "Bonus"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   1005
         Columns(8).Caption=   "Price"
         Columns(8).Name =   "Price"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   1111
         Columns(9).Caption=   "D(PC)"
         Columns(9).Name =   "DiscPC"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   979
         Columns(10).Caption=   "D(Per)"
         Columns(10).Name=   "DiscPer"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(11).Width=   1085
         Columns(11).Caption=   "D(Val)"
         Columns(11).Name=   "DiscVal"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(12).Width=   1588
         Columns(12).Caption=   "Amount"
         Columns(12).Name=   "Value"
         Columns(12).Alignment=   1
         Columns(12).CaptionAlignment=   2
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
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
   Begin VB.Frame FramExpense 
      Height          =   2415
      Left            =   9150
      TabIndex        =   104
      Top             =   5513
      Visible         =   0   'False
      Width           =   4215
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridExpense 
         Height          =   1860
         Left            =   120
         TabIndex        =   105
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
         stylesets(0).Picture=   "FrmPurchaseOrder.frx":010C
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
         stylesets(1).Picture=   "FrmPurchaseOrder.frx":0128
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
         stylesets(2).Picture=   "FrmPurchaseOrder.frx":0144
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
         TabIndex        =   106
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
         TabIndex        =   107
         Top             =   2100
         Width           =   1020
      End
   End
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4545
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   1305
      TabIndex        =   80
      Top             =   5640
      Width           =   2295
      Begin SITextBox.Txt TxtSerial 
         Height          =   315
         Left            =   120
         TabIndex        =   81
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
         TabIndex        =   82
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
         stylesets(0).Picture=   "FrmPurchaseOrder.frx":0160
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
      Left            =   14865
      TabIndex        =   74
      Top             =   1080
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
         TabIndex        =   75
         Tag             =   "NC"
         Text            =   "FrmPurchaseOrder.frx":017C
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
   Begin SITextBox.Txt TxtPaidAmount 
      Height          =   315
      Left            =   11715
      TabIndex        =   29
      Top             =   9098
      Visible         =   0   'False
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
      Left            =   1935
      TabIndex        =   4
      Top             =   2760
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
      Left            =   3225
      TabIndex        =   41
      Top             =   2753
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
      Left            =   6870
      TabIndex        =   40
      Top             =   2753
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
   Begin SITextBox.Txt TxtOrderID 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   2018
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
   Begin SITextBox.Txt TxtCity 
      Height          =   315
      Left            =   11400
      TabIndex        =   39
      Top             =   2753
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
      Left            =   2865
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2753
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
      MICON           =   "FrmPurchaseOrder.frx":0293
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9060
      TabIndex        =   36
      Top             =   9488
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
      MICON           =   "FrmPurchaseOrder.frx":02AF
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7740
      TabIndex        =   30
      Top             =   9488
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
      MICON           =   "FrmPurchaseOrder.frx":02CB
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5100
      TabIndex        =   32
      Top             =   9488
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
      MICON           =   "FrmPurchaseOrder.frx":02E7
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10380
      TabIndex        =   37
      Top             =   9488
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
      MICON           =   "FrmPurchaseOrder.frx":0303
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6420
      TabIndex        =   31
      Top             =   9488
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
      MICON           =   "FrmPurchaseOrder.frx":031F
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   5708
      TabIndex        =   49
      Top             =   8505
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
      Left            =   6893
      TabIndex        =   26
      Top             =   8505
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
      DecimalPoint    =   6
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtNetAmount 
      Height          =   315
      Left            =   10028
      TabIndex        =   52
      Top             =   8505
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   900
      TabIndex        =   9
      Top             =   4545
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
      Left            =   1755
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4545
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
      MICON           =   "FrmPurchaseOrder.frx":033B
      BC              =   12632256
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3120
      Left            =   300
      TabIndex        =   54
      Top             =   4860
      Width           =   14595
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   28
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
      stylesets(0).Picture=   "FrmPurchaseOrder.frx":0357
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
      Columns.Count   =   28
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ProductID"
      Columns(0).Name =   "ProductID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1085
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
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1455
      Columns(5).Caption=   "Size"
      Columns(5).Name =   "SizeName"
      Columns(5).CaptionAlignment=   2
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
      Columns(20).Width=   2064
      Columns(20).Caption=   "RetailPrice"
      Columns(20).Name=   "RetailPrice"
      Columns(20).Alignment=   1
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      Columns(21).Width=   2249
      Columns(21).Caption=   "IsWSDiscb4ST"
      Columns(21).Name=   "IsWSDiscb4ST"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   11
      Columns(21).FieldLen=   256
      Columns(22).Width=   2249
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
      Columns(24).Width=   1667
      Columns(24).Caption=   "BatchNo"
      Columns(24).Name=   "BatchNo"
      Columns(24).DataField=   "Column 24"
      Columns(24).DataType=   8
      Columns(24).FieldLen=   256
      Columns(25).Width=   2117
      Columns(25).Caption=   "ExpiryDate"
      Columns(25).Name=   "ExpiryDate"
      Columns(25).DataField=   "Column 25"
      Columns(25).DataType=   8
      Columns(25).FieldLen=   256
      Columns(26).Width=   3200
      Columns(26).Visible=   0   'False
      Columns(26).Caption=   "ColourID"
      Columns(26).Name=   "ColourID"
      Columns(26).DataField=   "Column 26"
      Columns(26).DataType=   8
      Columns(26).FieldLen=   256
      Columns(27).Width=   3200
      Columns(27).Visible=   0   'False
      Columns(27).Caption=   "SizeID"
      Columns(27).Name=   "SizeID"
      Columns(27).DataField=   "Column 27"
      Columns(27).DataType=   8
      Columns(27).FieldLen=   256
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpOrderDate 
      Height          =   315
      Left            =   2565
      TabIndex        =   1
      Top             =   2018
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
      Left            =   6240
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   2018
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
      Left            =   7275
      TabIndex        =   57
      Tag             =   "NC"
      Top             =   2033
      Width           =   2340
      _ExtentX        =   4128
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
      Left            =   6915
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   2018
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
      MICON           =   "FrmPurchaseOrder.frx":0373
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3780
      TabIndex        =   33
      Top             =   9488
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
      MICON           =   "FrmPurchaseOrder.frx":038F
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPreviousPayable 
      Height          =   315
      Left            =   8925
      TabIndex        =   64
      Top             =   9098
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtTotalPayable 
      Height          =   315
      Left            =   10320
      TabIndex        =   66
      Top             =   9098
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtBillDisc 
      Height          =   315
      Left            =   7868
      TabIndex        =   27
      Top             =   8505
      Width           =   975
      _ExtentX        =   1720
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   8220
      TabIndex        =   68
      Top             =   1613
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
      Left            =   3938
      TabIndex        =   71
      Top             =   8505
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
      Left            =   8843
      TabIndex        =   28
      Top             =   8505
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
   Begin JeweledBut.JeweledButton BtnBarCode 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5490
      TabIndex        =   34
      Top             =   8888
      Visible         =   0   'False
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
      MICON           =   "FrmPurchaseOrder.frx":03AB
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnChangePrice 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2430
      TabIndex        =   35
      Top             =   9495
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
      MICON           =   "FrmPurchaseOrder.frx":03C7
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2115
      TabIndex        =   13
      Top             =   4545
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
      Left            =   8130
      TabIndex        =   15
      Top             =   4545
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
      Left            =   9150
      TabIndex        =   17
      Top             =   4545
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
      Left            =   8640
      TabIndex        =   16
      Top             =   4545
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
      Left            =   13020
      TabIndex        =   24
      Top             =   4545
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
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtPrice 
      Height          =   315
      Left            =   10710
      TabIndex        =   20
      Top             =   4545
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
      Left            =   12540
      TabIndex        =   23
      Top             =   4545
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
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   11355
      TabIndex        =   21
      Top             =   4545
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
   Begin SITextBox.Txt TxtBonus 
      Height          =   315
      Left            =   9690
      TabIndex        =   18
      Top             =   4545
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
      Left            =   10230
      TabIndex        =   19
      Top             =   4545
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
      Left            =   9990
      TabIndex        =   94
      Top             =   1418
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
   Begin SITextBox.Txt TxtSaleTaxPer 
      Height          =   315
      Left            =   12030
      TabIndex        =   22
      Top             =   4545
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
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   13695
      TabIndex        =   25
      Top             =   4545
      Width           =   855
      _ExtentX        =   1508
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridOffer 
      Height          =   1365
      Left            =   1755
      TabIndex        =   97
      Top             =   6713
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
      stylesets(0).Picture=   "FrmPurchaseOrder.frx":03E3
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
      Left            =   9630
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   2033
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
      Left            =   10935
      TabIndex        =   100
      Tag             =   "NC"
      Top             =   2033
      Width           =   2205
      _ExtentX        =   3889
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
      Left            =   10575
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   2033
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
      MICON           =   "FrmPurchaseOrder.frx":03FF
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtRetailPrice 
      Height          =   315
      Left            =   10065
      TabIndex        =   102
      Top             =   3473
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
   Begin SITextBox.Txt TxtBillNo 
      Height          =   315
      Left            =   1935
      TabIndex        =   5
      Top             =   3473
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
      Left            =   2685
      TabIndex        =   6
      Top             =   3473
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
      Left            =   4950
      TabIndex        =   8
      Top             =   3473
      Width           =   5115
      _ExtentX        =   9022
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
      Left            =   3435
      TabIndex        =   7
      Top             =   3473
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
   Begin JeweledBut.JeweledButton BtnSale 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5895
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   2003
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
      MICON           =   "FrmPurchaseOrder.frx":041B
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   3870
      TabIndex        =   113
      Top             =   2018
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpBillDate 
      Height          =   315
      Left            =   4605
      TabIndex        =   114
      Top             =   2018
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
      Left            =   1440
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   4215
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
      MICON           =   "FrmPurchaseOrder.frx":0437
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBatchNo 
      Height          =   315
      Left            =   1755
      TabIndex        =   10
      Top             =   4215
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
      Left            =   2430
      TabIndex        =   11
      Top             =   4215
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
   Begin SITextBox.Txt TxtTotalQtys 
      Height          =   315
      Left            =   4823
      TabIndex        =   125
      Top             =   8505
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
   Begin SITextBox.Txt TxtProductName2 
      Height          =   315
      Left            =   2430
      TabIndex        =   131
      Top             =   10170
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpPromiseDate 
      Height          =   315
      Left            =   13230
      TabIndex        =   138
      Top             =   2745
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
   Begin VB.Label LblPromiseDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Promise Date"
      Height          =   195
      Left            =   13230
      TabIndex        =   139
      Top             =   2520
      Width           =   945
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
      Left            =   10350
      TabIndex        =   134
      Top             =   3780
      Width           =   1905
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
      Left            =   11745
      TabIndex        =   130
      Top             =   9450
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
      Left            =   9675
      TabIndex        =   129
      Top             =   10200
      Width           =   570
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty(s)"
      Height          =   195
      Left            =   4823
      TabIndex        =   126
      Top             =   8280
      Width           =   810
   End
   Begin VB.Label LblSize 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   5880
      TabIndex        =   124
      Top             =   4350
      Width           =   300
   End
   Begin VB.Label LblColour 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
      Height          =   195
      Left            =   4800
      TabIndex        =   122
      Top             =   4350
      Width           =   450
   End
   Begin VB.Label LblPre 
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
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   1935
      TabIndex        =   119
      Top             =   3758
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill ID"
      Height          =   195
      Left            =   3870
      TabIndex        =   116
      Top             =   1823
      Width           =   405
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   4605
      TabIndex        =   115
      Top             =   1823
      Width           =   585
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   4950
      TabIndex        =   111
      Top             =   3263
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bilty No."
      Height          =   195
      Left            =   2670
      TabIndex        =   110
      Top             =   3263
      Width           =   585
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      Height          =   195
      Left            =   1935
      TabIndex        =   109
      Top             =   3263
      Width           =   495
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      Height          =   195
      Left            =   3420
      TabIndex        =   108
      Top             =   3263
      Width           =   780
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ret Price"
      Height          =   195
      Left            =   10080
      TabIndex        =   103
      Top             =   3233
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   10935
      TabIndex        =   99
      Top             =   1793
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   9630
      TabIndex        =   98
      Top             =   1793
      Width           =   1095
   End
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   13695
      TabIndex        =   96
      Top             =   4350
      Width           =   540
   End
   Begin VB.Label LblSaleTaxPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax%"
      Height          =   195
      Left            =   12030
      TabIndex        =   95
      Top             =   4350
      Width           =   390
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Val"
      Height          =   195
      Left            =   9990
      TabIndex        =   93
      Top             =   1223
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label LblOffer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Offer"
      Height          =   195
      Left            =   10230
      TabIndex        =   92
      Top             =   4350
      Width           =   345
   End
   Begin VB.Label LblPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   10710
      TabIndex        =   91
      Top             =   4350
      Width           =   405
   End
   Begin VB.Label LblMultiplier 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   8160
      TabIndex        =   90
      Top             =   4350
      Width           =   375
   End
   Begin VB.Label LblPackName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   6780
      TabIndex        =   89
      Top             =   4350
      Width           =   840
   End
   Begin VB.Label LblQtyLoose 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   9150
      TabIndex        =   88
      Top             =   4350
      Width           =   465
   End
   Begin VB.Label LblQtyPack 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   8640
      TabIndex        =   87
      Top             =   4350
      Width           =   480
   End
   Begin VB.Label LblDiscVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.Val"
      Height          =   195
      Left            =   13020
      TabIndex        =   86
      Top             =   4350
      Width           =   585
   End
   Begin VB.Label LblDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Dis%"
      Height          =   195
      Left            =   12540
      TabIndex        =   85
      Top             =   4350
      Width           =   345
   End
   Begin VB.Label LblDiscPC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc/PC"
      Height          =   195
      Left            =   11355
      TabIndex        =   84
      Top             =   4350
      Width           =   600
   End
   Begin VB.Label LblBonus 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bns(L)"
      Height          =   195
      Left            =   9690
      TabIndex        =   83
      Top             =   4350
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
      Left            =   12240
      TabIndex        =   79
      Top             =   3533
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
      Left            =   11835
      TabIndex        =   78
      Top             =   3173
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
      Left            =   12915
      TabIndex        =   77
      Top             =   1598
      Width           =   435
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      Height          =   195
      Left            =   8843
      TabIndex        =   73
      Top             =   8280
      Width           =   1020
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items"
      Height          =   195
      Left            =   3938
      TabIndex        =   72
      Top             =   8280
      Width           =   780
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
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
      TabIndex        =   70
      Top             =   180
      Width           =   2775
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   8220
      TabIndex        =   69
      Top             =   1298
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LblTtlPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payable"
      Height          =   195
      Left            =   10335
      TabIndex        =   67
      Top             =   8873
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Payable"
      Height          =   195
      Left            =   8820
      TabIndex        =   65
      Top             =   8873
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   7868
      TabIndex        =   63
      Top             =   8280
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
      Left            =   10920
      TabIndex        =   62
      Top             =   3173
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
      Left            =   10815
      TabIndex        =   61
      Top             =   3488
      Width           =   1035
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   7275
      TabIndex        =   60
      Top             =   1823
      Width           =   840
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   6240
      TabIndex        =   59
      Top             =   1823
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   900
      TabIndex        =   56
      Top             =   4350
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3660
      TabIndex        =   55
      Top             =   4350
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
      Left            =   10028
      TabIndex        =   53
      Top             =   8280
      Width           =   840
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount (%)"
      Height          =   195
      Left            =   6893
      TabIndex        =   51
      Top             =   8280
      Width           =   885
   End
   Begin VB.Label LblTotalAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amount"
      Height          =   195
      Left            =   5708
      TabIndex        =   50
      Top             =   8280
      Width           =   990
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Amount"
      Height          =   195
      Left            =   11715
      TabIndex        =   48
      Top             =   8873
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   11400
      TabIndex        =   47
      Top             =   2573
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   6870
      TabIndex        =   46
      Top             =   2543
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Name"
      Height          =   195
      Left            =   3225
      TabIndex        =   45
      Top             =   2543
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender ID"
      Height          =   195
      Left            =   1920
      TabIndex        =   44
      Top             =   2543
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
      Height          =   195
      Left            =   2595
      TabIndex        =   43
      Top             =   1823
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID"
      Height          =   195
      Left            =   1935
      TabIndex        =   42
      Top             =   1823
      Width           =   600
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application

Dim vDate, vNow, vServerDate As Date, vHDiff As Integer, vSystemDate As Boolean
Dim vMode As FormMode
Dim vUnitPrice As Double
Dim vUnitWSPrice As Double
Dim vUnitRetailPrice As Double
Dim vIsWSDiscb4ST As Boolean
Dim vIsRetailSaleTax As Boolean
Dim vIsWSSaleTax As Boolean, isPrice As Boolean

Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim vMaxBinID As Integer
Dim RsBody As New ADODB.Recordset
Dim RsBodySerial As New ADODB.Recordset
Dim RsProductOffer As New ADODB.Recordset
Dim RsExpense As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim QtyOffer As Integer
Dim Rebate As Integer
Dim Flag As Boolean
Dim ssql As String
Dim vStrSQL, vRandomID As String
Dim vOrderID  As Integer
Dim vOrderDate  As Date
Dim ExpenseFlag As Boolean
Dim vExpAmount As Double
Dim vQtyLoose, vTotalQtyLoose, vError As Double
Dim vColour, vShowStock As Boolean
Dim vMargin As String
Dim i, vGridRows As Integer
Dim vNoofPrints As Byte
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
    TxtAmount.Text = Val(TxtAmount.Text) + Val(TxtSaleTaxVal.Text) - Val(TxtOffer.Text)
End Sub

Private Sub SubCalculateFooter()
   If TxtTotalAmount.Text = "" Then Exit Sub
   TxtNetAmount.Text = SelfRound(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text)
   TxtTotalPayable.Text = Abs(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousPayable.Text, Val(TxtPreviousPayable.Text) * -1)))
   LblTtlPayable.Caption = IIf(Val(TxtNetAmount.Text) + Val(IIf(lblPayable.Caption = "Previous Payable", TxtPreviousPayable.Text, Val(TxtPreviousPayable.Text) * -1)) < 0, "Total Receivable", "Total Payable")
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
          TxtPreviousPayable.Text = CN.Execute("SELECT isnull(dbo.FunCurrentDebit(" & Val(TxtVenderID.Text) & ",'" & DtpOrderDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
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
      SchProduct.ParaInPurchase = True
      SchProduct.ParaInWhere = " and isLocked = 0 and isNoCostProduct = 0 and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
      SchProduct.ParainShowStock = vShowStock
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
   If TxtCode.Enabled = False Then FunSelectProduct = False: Exit Function
   If Trim(TxtCode.Text) = "" Then FunSelectProduct = False: Exit Function
   CmbPackName.Clear
   vStrSQL = "select distinct pp.PackingID, Packingname from ProductPacking pp inner join packings p on p.packingid = pp.packingid" & vbCrLf _
           + "left outer join ProductBarcodes b on b.productid = pp.productid" & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "b.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')"
           
   With CN.Execute(vStrSQL)
      CmbPackName.AddItem ""
      While Not .EOF
         CmbPackName.AddItem !PackingName
         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
         .MoveNext
      Wend
      .Close
   End With
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
        vStrSQL = " SELECT p.productid, Code, ProductName, PurPrice, WSPrice, RetailPrice, IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST, SaleTaxPer, PurDiscPC, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
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
         vIsWSDiscb4ST = !IsWSDiscb4ST
         vIsWSSaleTax = !IsWSSaleTax
         vIsRetailSaleTax = !IsRetailSaleTax
         TxtSaleTaxPer.Text = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
         LblRetailPrice.Caption = !RetailPrice
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
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
         End If
         
         vStrSQL = "select h.OrderID, H.OrderDate, (isnull(QtyPack,0)*isnull(multiplier,0))+QtyLoose as Qty, pt.PartyName" & vbCrLf & _
               " from PurchaseOrderHeader h inner join PurchaseOrderbody b on h.OrderID = b.OrderID and h.OrderDate = b.OrderDate" & vbCrLf & _
               " inner join Parties pt on pt.PartyID = h.VendorID " & vbCrLf & _
               " where IsPurchase = 0 and h.OrderDate < '" & DtpOrderDate.DateValue & "' and ProductID = " & Val(TxtProductID.Text) & " Order by h.OrderDate desc"

         With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               LblPre.Caption = "ID = " & !OrderID & ", Date = " & Format(!OrderDate, "DD/MM/yyyy") & ", QtyLoose = " & !Qty & " Party = " & !PartyName
            Else
               LblPre.Caption = ""
            End If
            .Close
         End With
         LblPre.Visible = True
         
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
         
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & GetServerDate + 1 & "',0),0)"
          With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               vQtyLoose = .Fields(0).Value
            Else
               vQtyLoose = 0
            End If
         End With
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         With CN.Execute("Select isnull(abbreviation,'') from packings where packingname = '" & CmbPackName.Text & "'")
            If .RecordCount > 0 Then
               LblStock.Caption = LblStock.Caption & " " & .Fields(0).Value
            Else
               LblStock.Caption = LblStock.Caption & " "
            End If
         End With
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
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
'         If ObjRegistry.NegativeSale = False Then
'            If Val(LblStock.Caption) <= 0 Then
'               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'               FunSelectProduct = False
'               Exit Function
'            End If
'         End If
                  
         If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & ",Null,0,0,0,0,0,0,'" & DtpOrderDate.DateValue + 1 & "',0),0)"
            With CN.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
            LblAllStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & TxtProductID.Text & ",(" & vQtyLoose & "))").Fields(0).Value
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
         Else
            LblAllStock.Visible = False
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
         
         If ObjRegistry.ShowMultiBranches Then
'            PopulateDataToGridBranch
             PopulateDataToGridBranchLive
            FrameMultiBranchStock.Visible = True
         Else
            FrameMultiBranchStock.Visible = False
         End If
                  
         SubCalculateBody
         'Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FrmHistory.Visible = False
         FrmProductPrices.Visible = False
         FrameMultiBranchStock.Visible = False
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtProductID.Text = ""
         TxtCode.Text = ""
         If CmbPackName.ListCount > 0 Then CmbPackName.ListIndex = 0
         TxtProductName.Text = ""
         TxtMultiplier.Text = ""
         TxtPrice.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtAmount.Text = ""
         LblPre.Visible = False
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

Private Sub BtnBarCode_Click()
   On Error GoTo ErrorHandler
   If BtnSave.Enabled Then BtnSave_Click
   If vIsNewRecord = False Then
      vOrderID = TxtOrderID.Text
      vOrderDate = DtpOrderDate.DateValue
   End If
      ssql = " Select b.ProductID, ProductName, Price, isnull(Multiplier,1) * isnull(Qtypack,0)+ Bonus + QtyLoose as QtyLoose" & vbCrLf _
         + " from PurchaseOrderBody b inner join Products p on b.ProductID = p.ProductID" & vbCrLf _
         + " where OrderID = " & vOrderID & " and OrderDate = '" & vOrderDate & "'"
'   sSql = "select b.ProductID, Code, ProductName from ProductBarcodes b inner join Products p on p.productid = b.ProductID where len(code) = 11 and code like '110%'"
   
   Dim i As Integer
   With CN.Execute(ssql)
      FrmMultiBarcodes.SubClearFields
      FrmMultiBarcodes.TxtTotQty.Text = "0"
      For i = 1 To .RecordCount
         FrmMultiBarcodes.Grid.Columns("ID").Text = !Productid
         FrmMultiBarcodes.Grid.Columns("Name").Text = !ProductName
         FrmMultiBarcodes.Grid.Columns("QtyLoose").Value = !QtyLoose
         FrmMultiBarcodes.Grid.Update
         FrmMultiBarcodes.Grid.AddNew
         FrmMultiBarcodes.TxtTotQty.Text = Val(FrmMultiBarcodes.TxtTotQty.Text) + !QtyLoose
         .MoveNext
      Next i
   End With
   FrmMultiBarcodes.Grid.FirstRow = 0
   FrmMultiBarcodes.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnChangePrice_Click()
   On Error GoTo ErrorHandler
   Dim vOrderID As Integer
   Dim vOrderDate As Date
   vOrderID = Val(TxtOrderID.Text)
   vOrderDate = DtpOrderDate.DateValue
   If BtnSave.Enabled Then BtnSave_Click
   Dim PurchaseFlag As Boolean
   'PurchaseFlag = False
   'If MsgBox("Do you want to Show Purchase Price of Current Invoice?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbYes Then PurchaseFlag = True
   'If vIsNewRecord = False Then
   '   vOrderID = TxtOrderID.Text
   'End If
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
  
  
       ssql = " SELECT distinct p.*, isnull(PackingName,'') as PackingName, isnull(PP.Multiplier,0) as Multiplier,  PurPrice as PurchasePrice, Pb.SerialNo, " & vMargin & " from Products p" & vbCrLf _
       + " inner join (Select * from PurchaseOrderBody where OrderID = " & vOrderID & " and OrderDate = '" & vOrderDate & "')pb on p.productid = pb.productid" & vbCrLf _
       + " left outer join ProductBarCodes b on p.productid = b.productid " & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text) & " or b.Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & " Order by Pb.SerialNo"
      
      
   vStrSQL = "Select * from Products where ProductID in (select ProductID from PurchaseOrderBody where OrderID = " & vOrderID & " and OrderDate = '" & vOrderDate & "')"
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
         FrmChangePrice.Rs.Update
  
         FrmChangePrice.Grid.AddNew
         FrmChangePrice.Grid.Columns("ID").Text = !Productid
         FrmChangePrice.Grid.Columns("Name").Text = !ProductName
         FrmChangePrice.Grid.Columns("Packing").Text = !PackingName
         FrmChangePrice.Grid.Columns("Multiplier").Value = !Multiplier
         FrmChangePrice.Grid.Columns("PurPrice").Value = !PurPrice
         FrmChangePrice.Grid.Columns("RetailPrice").Value = !RetailPrice
         FrmChangePrice.Grid.Columns("WSPrice").Value = !WSPrice
         FrmChangePrice.Grid.Columns("ListPrice").Value = !ListPrice
         FrmChangePrice.Grid.Columns("Margin").Value = !Margin
         FrmChangePrice.Grid.Columns("DiscPC").Value = !DiscPC
         FrmChangePrice.Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
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
      
   '''''''''''''''''' ActivityLogBin For Clear Action
'      Call DeleteTempActivityLogBin(vRandomID)
      vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.Rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From purchaseOrderbody where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmPurchaseOrder, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Cleared Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmPurchaseOrder, eClearSavedRecord, TxtOrderID.Text, DtpOrderDate.DateValue, vGridRows & " Product/s Cleared")
      Grid.Redraw = True
  ''''''''''''''''''
   FormStatus = NewMode
'   cn.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & Val(TxtOrderID.Text) & ",'" & DtpOrderDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
     
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniPurchaseOrder", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
   Call ActivityLogBin("", eFrmPurchaseOrder, eDelete, TxtOrderID.Text, DtpOrderDate.DateValue, Grid.Rows - 1 & " Product/s Deleted Amount: " & Val(TxtNetAmount.Text))
   
'   vMaxBinID = FunGetMaxBinID
'   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_PurchaseOrderHeader Select " & vMaxBinID & ",'" & Date & "',* from PurchaseOrderHeader Where OrderID = " & TxtOrderID.Text & " And OrderDate ='" & DtpOrderDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_PurchaseOrderBody Select " & vMaxBinID & ",'" & Date & "', * from PurchaseOrderBody Where OrderID = " & TxtOrderID.Text & " And OrderDate ='" & DtpOrderDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Serial''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_PurchaseOrderBodySerial Select " & vMaxBinID & ",'" & Date & "', * from PurchaseOrderBodySerial Where OrderID = " & TxtOrderID.Text & " And OrderDate ='" & DtpOrderDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin ProductOffer''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_PurchaseOrderBodyOffer Select " & vMaxBinID & ",'" & Date & "', * from PurchaseOrderBodyOffer Where OrderID = " & TxtOrderID.Text & " And OrderDate ='" & DtpOrderDate.DateValue & "'")
'
'  '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  ''''''''''''''''''''''''''Delete Product Offer'''''''''''''''''''''
 
   Dim vBm As Variant
   GridOffer.Redraw = False
   vBm = GridOffer.Bookmark
   GridOffer.MoveFirst
   
   For vCounter = 0 To GridOffer.Rows - 1
      If Trim(GridOffer.Columns("Productid").CellValue(GridOffer.GetBookmark(i))) <> "" Then
         CN.Execute "Delete from PurchaseOrderBodyOffer where OrderID = " & Val(TxtOrderID.Text) & " And OrderDate ='" & DtpOrderDate.DateValue & "' and productid = " & Val(GridOffer.Columns("Productid").CellValue(GridOffer.GetBookmark(i)))
      End If
   Next vCounter
   
   GridOffer.Bookmark = vBm
   
   GridOffer.RemoveAll
   GridOffer.Redraw = True

  ''''''''''''''''''''''''''Delete Serials'''''''''''''''''''''
    If RsBodySerial.RecordCount > 0 Then
        RsBodySerial.MoveFirst
        For vCounter = 1 To RsBodySerial.RecordCount
            CN.Execute "Delete from PurchaseOrderBodySerial where OrderID = " & Val(TxtOrderID.Text) & " And OrderDate ='" & DtpOrderDate.DateValue & "' and productid = " & RsBodySerial!Productid & " and Serial ='" & RsBodySerial!serial & "'"
            RsBodySerial.MoveNext
        Next vCounter
    End If
      
  ''''''''''''''''''''''''''Delete Purchase Body'''''''''''''''''''''
   Grid.Redraw = False
   Grid.MoveFirst
   Call ActivityLog("Purchase Order", eDelete, Val(TxtOrderID.Text), DtpOrderDate.DateValue)
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from PurchaseOrderBody where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "' and productid = " & Grid.Columns("ProductID").Text & " and Price = " & Val(Grid.Columns("Price").Text)
'          CN.Execute ("Insert Into Bin_PurchaseOrderBody Select " & FunGetMaxBinID & ", * from PurchaseOrderBody Where OrderID = " & TxtOrderID.Text & " And OrderDate ='" & DtpOrderDate.DateValue & "' and productid ='" & Grid.Columns("Productid").Text & "'")
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
    '''''''''''''''''''''''''''''''''''''''Delete Expense'''''''''''''''''''''''''''''''''''''''
   CN.Execute "Delete from PurchaseOrderExpense where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'"
   
   CN.Execute "Delete from PurchaseOrderHeader where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'"
   
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
   SchPurchaseOrder.ParaInOrderDate = DtpOrderDate.DateValue
   SchPurchaseOrder.Show vbModal
   If SchPurchaseOrder.ParaOutOrderID <> "" Then
      TxtOrderID.Text = SchPurchaseOrder.ParaOutOrderID
      DtpOrderDate.DateValue = SchPurchaseOrder.ParaOutOrderDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
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
   If vColour = True Then
      FrmPurchaseOrderRangeDetail.Show vbModal, Me
   Else
'      FrmProductOrderRangeGrid.Show vbModal, Me
      FrmProductRangeGrid.Show vbModal, Me
   End If
   
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

Private Sub BtnSale_Click()
   On Error GoTo ErrorHandler
   SchSale.ParaInBillDate = DtpOrderDate.DateValue
   SchSale.Show vbModal
   If SchSale.ParaOutBillDate <> "" Then
      TxtBillID.Text = SchSale.ParaOutBillID
      DtpBillDate.DateValue = SchSale.ParaOutBillDate
      GetSale
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateSaleToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from PurchaseOrderBody where OrderID=" & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   'If RsBody.RecordCount > 0 Then
      ssql = "select p.ProductName, code, b.* from SaleBody b join products p on p.productid = b.productid where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
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
            
            RsBody.AddNew
            RsBody!Productid = Val(!Productid)
            RsBody!Code = !Productid
            
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
            
            Grid.Columns("RetailPrice").Value = !RetailPrice
            Grid.Columns("IsWSDiscb4ST").Value = !IsWSDiscb4ST
            Grid.Columns("IsWSSaleTax").Value = !IsWSSaleTax
            Grid.Columns("IsRetailSaleTax").Value = !IsRetailSaleTax
            
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            
            '''''
            RsBody!Multiplier = IIf(IsNull(!Multiplier), Null, !Multiplier)
            RsBody!QtyPack = IIf(IsNull(!QtyPack), Null, !QtyPack)
            RsBody!QtyLoose = !Qty
            RsBody!Bonus = IIf(IsNull(!Bonus), Null, !Bonus)
            RsBody!Price = !Price
            
            RsBody!RetailPrice = !RetailPrice
            RsBody!IsWSDiscb4ST = !IsWSDiscb4ST
            RsBody!IsWSSaleTax = !IsWSSaleTax
            RsBody!IsRetailSaleTax = !IsRetailSaleTax
            
            RsBody!DiscPC = IIf(IsNull(!DiscPC), Null, !DiscPC)
            RsBody!Offer = IIf(IsNull(!Offer), Null, !Offer)
            RsBody!SaleTaxPer = IIf(IsNull(!SaleTaxPer), Null, !SaleTaxPer)
            RsBody!SaleTaxval = IIf(IsNull(!SaleTaxval), Null, !SaleTaxval)
            RsBody!DiscPer = IIf(IsNull(!DiscPer), Null, !DiscPer)
            RsBody!DiscVal = IIf(IsNull(!DiscVal), Null, !DiscVal)
            RsBody!Amount = !Amount
            RsBody.Update
            ''''
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
   'End If
   
   RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   RsBodySerial.Open "Select * from PurchaseOrderBodySerial where OrderID=" & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
   Call PopulateDataToGridOffer
   Call PopulateDataToGridExpense
End Sub

Private Sub GetSale()
   On Error GoTo ErrorHandler
   TxtBillID.Text = FunGetMaxID
   ssql = "select h.*, OrganizationName, StoreName FROM SaleHeader h left outer join Organizations o on o.OrganizationID = h.OrganizationID inner join stores s on s.storeid = h.storeid where h.BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'"
   With CN.Execute(ssql)
      If Not .BOF Then
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
      End If
      .Close
   End With
   Call PopulateSaleToGrid
'   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub FrmMultiBranchStock_DragDrop(Source As Control, x As Single, Y As Single)

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


Private Sub SSOleDBGrid1_InitColumnProps()

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
   
'   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
   
   If MsgBox("Do you want to print this invoice With Amount", vbQuestion + vbYesNo, "Alert") = vbYes Then
      isPrice = True
   Else
      isPrice = False
   End If


'   If InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
'   vStrSQL = " select h.OrderID, h.OrderDate, BillDate as EntryDate, Pr.PartyName + ' - ' + H.VendorID as Vend_Name_ID, Pr.PartyName, Pr.Address, StoreName, BillNo, h.Description," & vbCrLf _
'      + " Code, '' Serial" & vbCrLf _
'      + ", '' ProductOffer, b.ProductID, ProductName, QtyPack, Multiplier, QtyLoose, Bonus, h.BillDisc," & vbCrLf _
'      + IIf(isPrice = True, " b.DiscPc, b.DiscPer, DiscVal, Offer, b.SaleTaxPer, SaleTaxval, ", " 0 DiscPc, 0 DiscPer, 0 DiscVal, 0 Offer, 0 SaleTaxPer, 0 SaleTaxval, ") & vbCrLf _
'      + IIf(isPrice = True, " (Amount-(Amount* (isnull(BillDiscPer,0)/100)) + (Amount* isnull(OtherCharges,0) /TotalAmount)) /((isnull(QtyPack,0)*isnull(Multiplier,0))+QtyLoose+isnull(Bonus,0)+(Amount* isnull(TotalExpense,0) /TotalAmount)) as price, ", " 0 price,") & vbCrLf _
'      + IIf(isPrice = True, " Amount-(Amount* (isnull(BillDiscPer,0)/100)) + (Amount* isnull(OtherCharges,0) /TotalAmount)+(Amount* isnull(TotalExpense,0) /TotalAmount) as Amount,", " 0 amount,") & vbCrLf _
'      + " 0 as LastPrice, isnull(PK.PackingName,'PC') as PackingName, z.ZoneID, ZoneName, Sec.SectorName, UserName, " & vbCrLf _
'      + " b.RetailPrice" & vbCrLf _
'      + " from PurchaseOrderBody b inner join products p on b.productid = p.productid" & vbCrLf _
'      + " inner join PurchaseOrderHeader h on h.OrderID = b.OrderID and h.OrderDate = b.OrderDate" & vbCrLf _
'      + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
'      + " inner join parties pr on pr.partyid = h.VendorID" & vbCrLf _
'      + " left outer join Packings PK on Pk.PackingID = B.PackingID" & vbCrLf _
'      + " left outer join sectors sec on sec.sectorid = pr.sectorid" & vbCrLf _
'      + " left outer join zones z on z.zoneid = sec.zoneid" & vbCrLf _
'      + " left outer join users u on u.userno = h.userno" & vbCrLf _
'      + " where h.OrderID = " & Val(TxtOrderID.Text) & " and h.OrderDate='" & DtpOrderDate.DateValue & "'"
'   Else
'      vStrSQL = "  select UserName, h.OrderID as BillID, h.OrderDate as BillDate, isnull(h.Description,'') as Remarks, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, AccountName as Customer, b.ProductID," & vbCrLf _
'            + " 0 as CashReceived, p.ProductName, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + QtyLoose as Qty" & IIf(isPrice = True, " ,b.price/isnull(multiplier,1) as price, b.amount, b.DiscPC, b.DiscPer, b.DiscVal", " ,0 price, 0 amount, 0 DiscPC, 0 DiscPer, 0 DiscVal") & vbCrLf _
'            + " from PurchaseOrderHeader h inner join PurchaseOrderBody b on h.OrderID = b.OrderID and h.OrderDate = b.OrderDate" & vbCrLf _
'            + " inner join products p on p.productid = b.productid" & vbCrLf _
'            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
'            + " left outer join ChartofAccounts c on c.AccountNo = h.VendorID" & vbCrLf _
'            + " where h.OrderID = " & Val(TxtOrderID.Text) & " and h.OrderDate='" & DtpOrderDate.DateValue & "' Order By SerialNo"
'   End If
   
    vStrSQL = " select h.OrderID, h.OrderDate, h.PromiseDate, BillDate as EntryDate, Pr.PartyName + ' - ' + cast(H.VendorID as varchar(10)) as Vend_Name_ID, Pr.PartyName, Pr.Address, StoreName, BillNo, h.Description," & vbCrLf _
      + " b.code, b.productID, '' Serial" & vbCrLf _
      + ", '' ProductOffer, b.ProductID, ProductName, QtyPack, Multiplier, QtyLoose, Bonus, h.BillDisc," & vbCrLf _
      + IIf(isPrice = True, " b.DiscPc, b.DiscPer, DiscVal, Offer, b.SaleTaxPer, SaleTaxval, ", " 0 DiscPc, 0 DiscPer, 0 DiscVal, 0 Offer, 0 SaleTaxPer, 0 SaleTaxval, ") & vbCrLf _
      + IIf(isPrice = True, "  price, ", " 0 price,") & vbCrLf _
      + IIf(isPrice = True, " Amount,", " 0 amount,") & vbCrLf _
      + " 0 as LastPrice, isnull(PK.PackingName,'PC') as PackingName, z.ZoneID, ZoneName, Sec.SectorName, UserName, " & vbCrLf _
      + " b.RetailPrice" & vbCrLf _
      + " from PurchaseOrderBody b inner join products p on b.productid = p.productid" & vbCrLf _
      + " inner join PurchaseOrderHeader h on h.OrderID = b.OrderID and h.OrderDate = b.OrderDate" & vbCrLf _
      + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
      + " inner join parties pr on pr.partyid = h.VendorID" & vbCrLf _
      + " left outer join Packings PK on Pk.PackingID = B.PackingID" & vbCrLf _
      + " left outer join sectors sec on sec.sectorid = pr.sectorid" & vbCrLf _
      + " left outer join zones z on z.zoneid = sec.zoneid" & vbCrLf _
      + " left outer join users u on u.userno = h.userno" & vbCrLf _
      + " where h.OrderID = " & Val(TxtOrderID.Text) & " and h.OrderDate='" & DtpOrderDate.DateValue & "'"
      
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   
   If cmbPrintType.Text = "Half Page" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseOrderHalf1.rpt")
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.x
      RptReportViewer.Report.RightMargin = 225
   ElseIf cmbPrintType.Text = "Thermal" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseOrderAurora.rpt")
      RptReportViewer.Report.TopMargin = 0
      RptReportViewer.Report.LeftMargin = 0
      RptReportViewer.Report.RightMargin = 0
   Else
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseOrder.rpt")
      RptReportViewer.Report.LeftMargin = 225
      RptReportViewer.Report.RightMargin = 0
      RptReportViewer.Report.TopMargin = 255
   End If
      
'   If InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
'      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CryRptPurchaseOrder.rpt")
'   Else
'      Set RptReportViewer.Report = New CrpPurchaseOrderInvoiceAurora
'   End If

   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   
   RptReportViewer.Report.ReportTitle = "Purchase Order"
   
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "0", "Phone # " & ObjRegistry.CompanyPhoneNo)
   
'   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   
   vPrinter = Split(CmbPrinters.Text, ",")
'   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)

   'RptReportViewer.Report.ParameterFields(4).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
   'RptReportViewer.Report.PrintOut False
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   If ObjRegistry.PreviewSaleInoice = True Or ChkIsPreview.Value = 1 Then
      If ChkIsPrint.Value = 1 Then
         RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
      End If
       RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
   End If
'   CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      GetDataFromTexBoxesToGrid
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniPurchaseOrder", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
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
   If Trim(TxtVenderID.Text) = "" Then
      MsgBox "Enter Vender ID.", vbExclamation, Me.Caption
      TxtVenderID.SetFocus
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
       If DtpBillDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
'   If Trim(TxtStoreID.Text) = "" Then
'      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
'      If TxtStoreID.Visible And TxtStoreID.Enabled Then TxtStoreID.SetFocus
'      Exit Sub
'   End If

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

   Call SubCalculateFooter

  RsBody.Filter = 0
  If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Purchase", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
  End If
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
  
  
   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'Purchase Order' and LocalComputerName = '" & LocalComputerName & "'"
   If CN.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview, IsPrint ) Values ('" & LocalComputerName & "', 'Purchase Order','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & "," & ChkIsPrint.Value & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & ", IsPrint = " & ChkIsPrint.Value & " Where FormType = 'Purchase Order' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   CN.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   CN.BeginTrans
   
   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmPurchaseOrder, eEdit, TxtOrderID.Text, DtpOrderDate.DateValue, "Amount: " & Val(TxtNetAmount.Text))
   
   If vIsNewRecord = True Then
      If CN.Execute("Select * from PurchaseOrderHeader where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'").RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtOrderID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
   
   vOrderID = TxtOrderID.Text
   vOrderDate = DtpOrderDate.DateValue
   
   If vIsNewRecord = False Then
'    Call ActivityLog("Purchase Order", eEdit, TxtOrderID.Text, DtpOrderDate.DateValue)
   End If
   
'   Call UserActivities
   
   ssql = "select * from PurchaseOrderHeader where OrderID=" & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'"
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, CN, adOpenDynamic, adLockPessimistic
      If .BOF Then
         .AddNew
         !OrderID = Val(TxtOrderID.Text)
         !OrderDate = DtpOrderDate.DateValue
         !BillID = IIf(Val(TxtBillID.Text) = 0, Null, TxtBillID.Text)
         !BillDate = DtpBillDate.DateValue
         !isPurchase = 0
         !UserNo = vUser
      End If
      !PromiseDate = IIf(DtpPromiseDate.DateValue = Empty, Null, DtpPromiseDate.DateValue)
      !vendorID = TxtVenderID.Text
      !StoreID = TxtStoreID.Text
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !BillNo = IIf(TxtBillNo.Text = "", Null, TxtBillNo.Text)
      !BiltyNo = IIf(TxtBiltyNo.Text = "", Null, TxtBiltyNo.Text)
      !VehicleNo = IIf(TxtVehicleNo.Text = "", Null, TxtVehicleNo.Text)
      !TotalAmount = Round(Val(TxtTotalAmount.Text))
      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
      !OtherCharges = IIf(Val(TxtOtherCharges.Text) = 0, Null, Val(TxtOtherCharges.Text))
      !TotalExpense = IIf(Val(TxtTotalExpense.Text) = 0, Null, Val(TxtTotalExpense.Text))
      !PAIDAMOUNT = IIf(TxtPaidAmount.Text = "", Null, Val(TxtPaidAmount.Text))
      !Description = IIf(TxtDescription.Text = "", Null, TxtDescription.Text)
      !PreviousAmount = IIf(lblPayable.Caption = "Previous Receivable", Val(TxtPreviousPayable.Text), Val(TxtPreviousPayable.Text) * -1)
'      !UserNo = vUser
      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
      .Update
      .Close
   End With
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !OrderID = Val(TxtOrderID.Text)
         !OrderDate = DtpOrderDate.DateValue
         'ssql = "update Products set PurPrice = " & RsBody!Price & ", PurDiscPC = " & RsBody!DiscPC & ", PurchasePackingID = " & IIf(IsNull(RsBody!PackingID), "Null", RsBody!PackingID) & " Where ProductID='" & RsBody!Productid & "'"
         'CN.Execute ssql
         If (Not IsNull(RsBody!PackingID)) And (Not IsNull(RsBody!Multiplier)) And (RsBody!Multiplier <> 0) Then
            If CN.Execute("select * from ProductPacking Where ProductID = " & Val(RsBody!Productid) & " and PackingID = " & RsBody!PackingID).RecordCount = 0 Then
               ssql = "INSERT INTO ProductPacking(PackingID,Multiplier,ProductID) VALUES ('" & RsBody!PackingID & "','" & RsBody!Multiplier & "'," & Val(RsBody!Productid) & ")"
               CN.Execute ssql
            Else
               ssql = "update ProductPacking set Multiplier = " & IIf(IsNull(RsBody!Multiplier), 0, RsBody!Multiplier) & " Where ProductID = " & Val(RsBody!Productid) & " and PackingID = " & RsBody!PackingID
               CN.Execute ssql
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
         !OrderID = Val(TxtOrderID.Text)
         !OrderDate = DtpOrderDate.DateValue
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
         !OrderID = Val(TxtOrderID.Text)
         !OrderDate = DtpOrderDate.DateValue
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
         !OrderID = Val(TxtOrderID.Text)
         !OrderDate = DtpOrderDate.DateValue
         .MoveNext
        Next vCounter
      End If
      .UpdateBatch
   End With
   
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmPurchaseOrder, eAdd, TxtOrderID.Text, DtpOrderDate.DateValue, Grid.Rows - 1 & " New Product/s Added Amount: " & Val(TxtNetAmount.Text))
'   If vIsNewRecord = True Then Call ActivityLog("Purchase Order", eAdd, TxtOrderID.Text, DtpOrderDate.DateValue)
   CN.CommitTrans
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
      ssql = "select top 3 pt.PartyName, VendorID, code, b.* " & vbCrLf & _
      " from PurchaseOrderHeader h inner join PurchaseOrderbody b on h.OrderID = b.OrderID and h.OrderDate = b.OrderDate" & vbCrLf & _
      " inner join Parties pt on pt.PartyID = h.VendorID " & vbCrLf & _
      " where b.productid = " & Val(TxtProductID.Text) & " order by b.OrderDate Desc"
      
      With CN.Execute(ssql)
         GridHistory.Redraw = False
         GridHistory.MoveFirst
         GridHistory.RemoveAll
         GridHistory.AllowAddNew = True
         While Not .EOF
            GridHistory.AddNew
            GridHistory.Columns("ID").Text = !vendorID
            GridHistory.Columns("Name").Text = !PartyName
            GridHistory.Columns("Date").Text = !OrderDate
'            If !PackingID = 0 Or IsNull(!PackingID) Then
'               GridHistory.Columns("PackingID").Value = ""
'            Else
'               GridHistory.Columns("PackingID").Value = !PackingID
'            End If
            If !PackingID = 0 Or IsNull(!PackingID) Then
               GridHistory.Columns("PackName").Text = ""
            Else
               GridHistory.Columns("PackName").Text = CN.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
            End If
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
   RsBody.Open "Select * from PurchaseOrderBody where OrderID=" & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.ProductName, code, b.* from PurchaseOrderBody b join products p on p.productid = b.productid where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "' Order by SerialNo asc "
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
            Grid.Columns("ExpiryDate").Text = IIf(IsNull(!ExpiryDate), "", !ExpiryDate)
            Grid.Columns("ProductName").Text = !ProductName
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
               Grid.Columns("sizeID").Value = ""
            Else
               Grid.Columns("SizeID").Value = !SizeID
            End If
            If !SizeID = 0 Or IsNull(!SizeID) Then
               Grid.Columns("SizeName").Text = ""
            Else
               Grid.Columns("SizeName").Text = CN.Execute("Select SizeName from Sizes where SizeID=" & !SizeID).Fields(0).Value
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
            Grid.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            Grid.Columns("Price").Value = !Price
            
            Grid.Columns("RetailPrice").Value = !RetailPrice
            Grid.Columns("IsWSDiscb4ST").Value = !IsWSDiscb4ST
            Grid.Columns("IsWSSaleTax").Value = !IsWSSaleTax
            Grid.Columns("IsRetailSaleTax").Value = !IsRetailSaleTax
            
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("Offer").Value = IIf(IsNull(!Offer), "", !Offer)
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(!Amount)
            TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + !QtyLoose + IIf(IsNull(!Bonus), "0", !Bonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack))
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
   
   TxtTotalItems.Text = Val(Grid.Rows) - 1
   
   RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   RsBodySerial.Open "Select * from PurchaseOrderBodySerial where OrderID=" & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
   Call PopulateDataToGridOffer
   Call PopulateDataToGridExpense
End Sub

Private Sub PopulateDataToGridExpense()
    If RsExpense.State = adStateOpen Then RsExpense.Close
    RsExpense.Open "Select * from PurchaseOrderExpense where OrderID =" & Val(TxtOrderID.Text) & " And OrderDate = '" & DtpOrderDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
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
            .MoveNext
         Wend
      End With

     If GridExpense.Rows > 0 Then GridExpense.FirstRow = 0
     GridExpense.Redraw = True
'      GridExpense.Visible = False
End Sub

Private Sub PopulateDataToGridserial()
   RsBodySerial.Filter = "ProductID = " & Val(Grid.Columns("ProductID").Text)
   If RsBodySerial.RecordCount > 0 Then
'       sSql = "select d.* from PurchaseOrderBodySerial d  where OrderID=" & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "' and ProductID = '" & Grid.Columns("ProductID").Text & "'"
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
      FrameMultiBranchStock.Visible = False
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
      
    
      
      DtpBillDate.DateValue = vDate
      DtpOrderDate.DateValue = vDate

      TxtOrderID.Text = FunGetMaxID()
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
      'TxtOrderID.Enabled = True
      DtpOrderDate.Enabled = True
      If DtpOrderDate.Enabled And DtpOrderDate.Visible Then DtpOrderDate.SetFocus
      GridOffer.Visible = False
      FramExpense.ZOrder 0
      vIsNewRecord = True
   Case Is = OpenMode
      'TxtOrderID.Enabled = False
      DtpOrderDate.Enabled = False
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
      'DtpOrderDate.SetFocus
      If DtpOrderDate.Enabled And DtpOrderDate.Visible Then DtpOrderDate.SetFocus
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
      If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus
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

Private Sub DtpOrderDate_Validate(Cancel As Boolean)
   If DtpOrderDate.Enabled = False Then Exit Sub
   TxtOrderID.Text = FunGetMaxID()
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = "Grid" Then
         Grid_DblClick
      ElseIf ActiveControl.Name = "GridSerial" Then
         GridSerial_DblClick
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
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then GetDataFromTexBoxesToGrid
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, False) = True Then TxtBiltyNo.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus Else TxtStoreID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtVenderID.SetFocus Else TxtBillNo.SetFocus
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
   On Error GoTo ErrorHandler
   If KeyAscii = vbKeyReturn Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
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
   On Error GoTo ErrorHandler
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Purchase Order"
   HelpLocation Me
   
   If ObjUserSecurity.ShowStock = True Or ObjUserSecurity.IsAdministrator Then
      vShowStock = True
   Else
      vShowStock = False
   End If
   
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
   ssql = "select * from FormDefaultSetting Where FormType = 'Purchase Order' and LocalComputerName = '" & LocalComputerName & "'"
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
   
   GridBranch.AddNew
   If ObjRegistry.ShowMultiBranches = True Then
      vStrSQL = "SELECT 'Braches' AS Branches," & vbCrLf _
            + "[1] Branch1, [2] Branch2, [3] Branch3, [4] Branch4, [5] Branch5, [6] Branch6, [7] Branch7, [8] Branch8, [9] Branch9," & vbCrLf _
            + "S1.StoreName BranchName1, S2.StoreName BranchName2, S3.StoreName BranchName3, S4.StoreName BranchName4, S5.StoreName BranchName5, S6.StoreName BranchName6, S7.StoreName BranchName7, S8.StoreName BranchName8, S9.StoreName BranchName9" & vbCrLf _
            + "FROM (" & vbCrLf _
            + "SELECT StoreID FROM Stores" & vbCrLf _
            + ") AS SourceTable" & vbCrLf _
            + "PIVOT(" & vbCrLf _
            + "AVG(StoreID) FOR [StoreID] IN([1], [2], [3], [4], [5], [6], [7], [8], [9])" & vbCrLf _
            + ") AS PivotStore" & vbCrLf _
            + "Left Outer Join Stores S1 on S1.StoreID = PivotStore.[1]" & vbCrLf _
            + "Left Outer Join Stores S2 on S2.StoreID = PivotStore.[2]" & vbCrLf _
            + "Left Outer Join Stores S3 on S3.StoreID = PivotStore.[3]" & vbCrLf _
            + "Left Outer Join Stores S4 on S4.StoreID = PivotStore.[4]" & vbCrLf _
            + "Left Outer Join Stores S5 on S5.StoreID = PivotStore.[5]" & vbCrLf _
            + "Left Outer Join Stores S6 on S6.StoreID = PivotStore.[6]" & vbCrLf _
            + "Left Outer Join Stores S7 on S7.StoreID = PivotStore.[7]" & vbCrLf _
            + "Left Outer Join Stores S8 on S8.StoreID = PivotStore.[8]" & vbCrLf _
            + "Left Outer Join Stores S9 on S9.StoreID = PivotStore.[9]"
            
      With CN.Execute(vStrSQL)
         If Not .EOF Then
            If Not (IsNull(!Branch1)) Then
               GridBranch.Columns("Branch1").Visible = True
               GridBranch.Columns("Branch1").Width = 120 * Len(IIf(IsNull(!BranchName1), 1, !BranchName1))
               GridBranch.Columns("Branch1").Caption = IIf(IsNull(!BranchName1), "", !BranchName1)
            End If
            If Not (IsNull(!Branch2)) Then
               GridBranch.Columns("Branch2").Visible = True
               GridBranch.Columns("Branch2").Width = 120 * Len(IIf(IsNull(!BranchName2), 1, !BranchName2))
               GridBranch.Columns("Branch2").Caption = IIf(IsNull(!BranchName2), "", !BranchName2)
            End If
            If Not (IsNull(!Branch3)) Then
               GridBranch.Columns("Branch3").Visible = True
               GridBranch.Columns("Branch3").Width = 120 * Len(IIf(IsNull(!BranchName3), 1, !BranchName3))
               GridBranch.Columns("Branch3").Caption = IIf(IsNull(!BranchName3), "", !BranchName3)
            End If
            If Not (IsNull(!Branch4)) Then
               GridBranch.Columns("Branch4").Visible = True
               GridBranch.Columns("Branch4").Width = 120 * Len(IIf(IsNull(!BranchName4), 1, !BranchName4))
               GridBranch.Columns("Branch4").Caption = IIf(IsNull(!BranchName4), "", !BranchName4)
            End If
            If Not (IsNull(!Branch5)) Then
               GridBranch.Columns("Branch5").Visible = True
               GridBranch.Columns("Branch5").Width = 120 * Len(IIf(IsNull(!BranchName5), 1, !BranchName5))
               GridBranch.Columns("Branch5").Caption = IIf(IsNull(!BranchName5), "", !BranchName5)
            End If
            If Not (IsNull(!Branch6)) Then
               GridBranch.Columns("Branch6").Visible = True
               GridBranch.Columns("Branch6").Width = 120 * Len(IIf(IsNull(!BranchName6), 1, !BranchName6))
               GridBranch.Columns("Branch6").Caption = IIf(IsNull(!BranchName6), "", !BranchName6)
            End If
            If Not (IsNull(!Branch7)) Then
               GridBranch.Columns("Branch7").Visible = True
               GridBranch.Columns("Branch7").Width = 120 * Len(IIf(IsNull(!BranchName7), 1, !BranchName7))
               GridBranch.Columns("Branch7").Caption = IIf(IsNull(!BranchName8), "", !BranchName7)
            End If
            If Not (IsNull(!Branch8)) Then
               GridBranch.Columns("Branch8").Visible = True
               GridBranch.Columns("Branch8").Width = 120 * Len(IIf(IsNull(!BranchName8), 1, !BranchName8))
               GridBranch.Columns("Branch8").Caption = IIf(IsNull(!BranchName8), "", !BranchName8)
            End If
            If Not (IsNull(!Branch9)) Then
               GridBranch.Columns("Branch9").Visible = True
               GridBranch.Columns("Branch9").Width = 120 * Len(IIf(IsNull(!BranchName9), 1, !BranchName9))
               GridBranch.Columns("Branch9").Caption = IIf(IsNull(!BranchName9), "", !BranchName9)
            End If
            
         End If
      End With
   End If

   vNoofPrints = IIf(IsNull(ObjRegistry.NoofPrints) Or Val(ObjRegistry.NoofPrints) = 0, 1, ObjRegistry.NoofPrints)
   
   vColour = ObjRegistry.ShowColourSize
   
   LblColour.Visible = vColour
   CmbColourName.Visible = vColour
   LblSize.Visible = vColour
   cmbSizeName.Visible = vColour
   Grid.Columns("ColourName").Visible = vColour
   Grid.Columns("SizeName").Visible = vColour
   
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
   
   vServerDate = CN.Execute("Select CONVERT(datetime, CONVERT(varchar, GETDATE(), 110)) ServerDate").Fields(0).Value
   vSystemDate = Abs(ObjRegistry.SystemDate)
   vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
   
'   With cn.Execute("Select * from Packings")
'      CmbPackName.AddItem ""
'      While Not .EOF
'         CmbPackName.AddItem !PackingName
'         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
'         .MoveNext
'      Wend
'      .Close
'   End With
         
   TxtStoreID.Text = IIf((ObjRegistry.StoreID = ""), "", ObjRegistry.StoreID)
   FunSelectStore ssValidate, True
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   
   TxtBatchNo.Visible = ObjRegistry.BatchNoVisible
   DtpExpiryDate.Visible = ObjRegistry.BatchNoVisible

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
   TxtSaleTaxVal.Visible = ObjRegistry.ShowSaleTax
   
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


   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpOrderDate.IsDateValid = False Then Exit Function
   If ObjRegistry.AllowContinuousBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(OrderID),0)+1 from PurchaseOrderHeader").Fields(0)
   ElseIf ObjRegistry.AllowMonthlyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(OrderID),0)+1 from PurchaseOrderHeader where Month(OrderDate) = '" & Month(DtpOrderDate.DateValue) & "' and  year(OrderDate) ='" & Year(DtpOrderDate.DateValue) & "'").Fields(0)
   ElseIf ObjRegistry.AllowDailyBillNo = True Then
      FunGetMaxID = CN.Execute("Select isnull(max(OrderID),0)+1 from PurchaseOrderHeader where OrderDate = '" & DtpOrderDate.DateValue & "'").Fields(0)
   Else
      FunGetMaxID = CN.Execute("Select isnull(max(OrderID),0)+1 from PurchaseOrderHeader where OrderDate = '" & DtpOrderDate.DateValue & "' and StoreID = " & TxtStoreID.Text).Fields(0)
   End If

   Exit Function
ErrorHandler:
'   Call ShowErrorMessage
End Function

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpOrderDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = CN.Execute("Select isnull(max(BinID),0)+1 from Bin_PurchaseOrderHeader ").Fields(0)
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
   DtpPromiseDate.DateValue = Null
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
   DtpExpiryDate.DateValue = ""
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
    Set FrmPurchaseOrder = Nothing
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
               ssql = "Select Productid From purchaseOrderbody where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
               With CN.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmPurchaseOrder, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Closed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmPurchaseOrder, eCloseSavedRecord, TxtOrderID.Text, DtpOrderDate.DateValue, vGridRows & " Product/s Closed")
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
   If Trim(Grid.Columns("Code").Text) = "" Then
      TxtSerial.Enabled = False
   Else
      TxtSerial.Enabled = True
   End If
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
   
   ssql = "Select Productid From PurchaseOrderbody where Orderid=" & Val(TxtOrderID.Text) & " and Orderdate ='" & DtpOrderDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
   With CN.Execute(ssql)
      If .EOF Then
         Call ActivityLogBin("", eFrmPurchaseOrder, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      Else
         Call ActivityLogBin("", eFrmPurchaseOrder, eRemoveRow, TxtOrderID.Text, DtpOrderDate.DateValue, "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
         Call ActivityLogBin(vRandomID, eFrmPurchaseOrder, eAddTempRecord, TxtOrderID.Text, DtpOrderDate.DateValue, "Pending Remove Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      End If
   End With
   
   RsBody.Filter = "ProductID = " & Val(TxtProductID.Text) & " and Price = " & Val(TxtPrice.Text)
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
   RsBodySerial.Filter = "ProductID = " & Val(TxtCode.Text) & " And Serial = '" & TxtSerial.Text & "'"
   If RsBodySerial.RecordCount > 0 Then
      With RsBodySerial
'        .Filter = 0
         .MoveFirst
         For vCounter = 1 To .RecordCount
            RsBodySerial.Delete
            .MoveNext
         Next vCounter
      End With
   End If
      CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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
      CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Removed ProdcutID-" & GridSerial.Columns("ProductID").Text & " Serial-" & GridSerial.Columns("Serial").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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
   TxtTotalItems.Text = Val(Grid.Rows) - 1
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
   Dim vStrAdd As String
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
   'vStrAdd = IIf(vColour = True, " and ColorID = " & CmbColourName.ItemData(CmbColourName.ListIndex) & " and SizeID = " & cmbSizeName.ItemData(cmbSizeName.ListIndex), "")
   If vColour = True Then
    vStrAdd = " and ColorID = " & CmbColourName.ItemData(CmbColourName.ListIndex) & " and SizeID = " & cmbSizeName.ItemData(cmbSizeName.ListIndex)
   Else
    vStrAdd = ""
   End If
   FrmHistory.Visible = False
   FrmProductPrices.Visible = False
   FrameMultiBranchStock.Visible = False
   If Val(Grid.Columns("Productid").Text) = 0 Then
      RsBody.Filter = "ProductID = " & Val(TxtProductID.Text) & " and Price = " & Val(TxtPrice.Text) & vStrAdd
   Else
      RsBody.Filter = "ProductID = " & Val(Grid.Columns("Productid").Text) & " and Price = " & Val(Grid.Columns("Price").Text)
   End If
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("Serial").Text = Grid.Rows
         Grid.Columns("ProductID").Text = Val(TxtProductID.Text)
         Grid.Columns("Code").Text = TxtCode.Text
         Grid.Columns("Price").Value = Val(TxtPrice.Text)
         RsBody!Productid = Val(TxtProductID.Text)
         RsBody!Code = TxtCode.Text
         RsBody!Price = Val(TxtPrice.Text)
         RsBody!BatchNo = Trim(TxtBatchNo.Text)
         RsBody!ExpiryDate = DtpExpiryDate.DateValue
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("Productid").Text = TxtProductID.Text And Val(Grid.Columns("Price").Text) = Val(TxtPrice.Text) Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  
                  ssql = "Select Productid From PurchaseOrderbody where OrderID=" & Val(TxtOrderID.Text) & " and Orderdate ='" & DtpOrderDate.DateValue & "' and productid = " & Val(Grid.Columns("ProductID").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmPurchaseOrder, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmPurchaseOrder, eEdit, TxtOrderID.Text, DtpOrderDate.DateValue, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
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
                  TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Val(Grid.Columns("QtyLoose").Value) + Val(Grid.Columns("Bonus").Value) + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
                  Grid.Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
                  Grid.Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, "", Val(TxtMultiplier.Text))
                  Grid.Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, "", Val(TxtQtyPack.Text))
                  Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
                  Grid.Columns("Bonus").Value = Val(TxtBonus.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("RetailPrice").Value = Val(TxtRetailPrice.Text)
                  Grid.Columns("IsWSDiscb4ST").Value = vIsWSDiscb4ST
                  Grid.Columns("IsWSSaleTax").Value = vIsWSSaleTax
                  Grid.Columns("IsRetailSaleTax").Value = vIsRetailSaleTax
                  Grid.Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  Grid.Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  Grid.Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  Grid.Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  Grid.Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  Grid.Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
                  RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
                  RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
                  RsBody!QtyLoose = Val(TxtQtyLoose.Text)
                  RsBody!Bonus = Val(TxtBonus.Text)
                  RsBody!Price = Val(TxtPrice.Text)
                  RsBody!RetailPrice = Val(TxtRetailPrice.Text)
                  RsBody!IsWSDiscb4ST = vIsWSDiscb4ST
                  RsBody!IsWSSaleTax = vIsWSSaleTax
                  RsBody!IsRetailSaleTax = vIsRetailSaleTax
                  RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
                  RsBody!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
                  RsBody!SaleTaxval = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
                  RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
                  RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
                  RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
                  RsBody!Amount = Val(TxtAmount.Text)
                  
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmPurchaseOrder, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmPurchaseOrder, eEdit, TxtOrderID.Text, DtpOrderDate.DateValue, "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmPurchaseOrder, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Pending Update Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                  
                  Grid.MoveLast

                  Call SubClearDetailArea
                  If TxtCode.Enabled = True And TxtCode.Visible = True Then TxtCode.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearDetailArea
         Grid.MoveLast
         If TxtCode.Enabled = True And TxtCode.Visible = True Then TxtCode.SetFocus
         Exit Sub
      End If
   End If
   Grid.Redraw = False
   With Grid
      If Trim(Grid.Columns("Productid").Text) = "" Then
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text)
         TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text)))
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmPurchaseOrder, eAddNewRowByEdit, TxtOrderID.Text, DtpOrderDate.DateValue, "Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
         Call ActivityLogBin(vRandomID, eFrmPurchaseOrder, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Pending Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         TxtTotalQtys.Text = Val(TxtTotalQtys.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Val(Grid.Columns("Pack").Value)) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
         ssql = "Select Productid From PurchaseOrderbody where Orderid=" & Val(TxtOrderID.Text) & " and Orderdate ='" & DtpOrderDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
         With CN.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmPurchaseOrder, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmPurchaseOrder, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            Else
               Call ActivityLogBin("", eFrmPurchaseOrder, eEdit, TxtOrderID.Text, DtpOrderDate.Date, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("QtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("QtyLoose").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmPurchaseOrder, eEdit, TxtOrderID.Text, DtpOrderDate.Date, "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmPurchaseOrder, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Pending Update Code-" & TxtCode.Text & " Qty-" & Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
      End If
      .Columns("BatchNo").Text = Trim(TxtBatchNo.Text)
      .Columns("ExpiryDate").Text = DtpExpiryDate.DateValue
      .Columns("ProductName").Text = TxtProductName.Text
       .Columns("ColourName").Text = CmbColourName.Text
      If CmbColourName.Text <> "" Then .Columns("ColourID").Value = CmbColourName.ItemData(CmbColourName.ListIndex)
      .Columns("SizeName").Text = cmbSizeName.Text
      If cmbSizeName.Text <> "" Then .Columns("SizeID").Value = cmbSizeName.ItemData(cmbSizeName.ListIndex)
      
      If vColour = True And .Columns("ColourID").Text <> "" Then
         RsBody!ColourID = .Columns("ColourID").Text
         RsBody!SizeID = .Columns("SizeID").Text
      End If
'      .Columns("ColourName").Text = CmbColourName.Text
'      .Columns("ColourID").Value = IIf(CmbColourName.ListIndex > -1, CmbColourName.ItemData(CmbColourName.ListIndex), "")
'      .Columns("SizeName").Text = cmbSizeName.Text
'      .Columns("SizeID").Value = IIf(cmbSizeName.ListIndex > -1, cmbSizeName.ItemData(cmbSizeName.ListIndex), "")
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Value = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
      .Columns("Pack").Value = IIf(Val(TxtMultiplier.Text) = 0, "", Val(TxtMultiplier.Text))
      .Columns("QtyPack").Value = IIf(Val(TxtQtyPack.Text) = 0, "", Val(TxtQtyPack.Text))
      .Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
      .Columns("Bonus").Value = Val(TxtBonus.Text)
      .Columns("Price").Value = Val(TxtPrice.Text)
      .Columns("RetailPrice").Value = Val(TxtRetailPrice.Text)
      .Columns("IsWSDiscb4ST").Value = vIsWSDiscb4ST
      .Columns("IsWSSaleTax").Value = vIsWSSaleTax
      .Columns("IsRetailSaleTax").Value = vIsRetailSaleTax
      .Columns("Offer").Value = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      .Columns("SaleTaxPer").Value = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
      .Columns("SaleTaxVal").Value = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
      .Columns("DiscPC").Value = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      .Columns("DiscPer").Value = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      .Columns("DiscVal").Value = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      .Columns("Amount").Value = Val(TxtAmount.Text)
      RsBody!BatchNo = IIf(Trim(TxtBatchNo.Text) = "", Null, Trim(TxtBatchNo.Text))
      RsBody!ExpiryDate = IIf(DtpExpiryDate.DateValue = "", Null, DtpExpiryDate.DateValue)
'      RsBody!ColourID = IIf(CmbColourName.ListIndex = -1, Null, CmbColourName.ItemData(CmbColourName.ListIndex))
'      RsBody!SizeID = IIf(cmbSizeName.ListIndex = -1, Null, cmbSizeName.ItemData(cmbSizeName.ListIndex))
      RsBody!PackingID = IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
      RsBody!Multiplier = IIf(Val(TxtMultiplier.Text) = 0, Null, Val(TxtMultiplier.Text))
      RsBody!QtyPack = IIf(Val(TxtQtyPack.Text) = 0, Null, Val(TxtQtyPack.Text))
      RsBody!QtyLoose = Val(TxtQtyLoose.Text)
      RsBody!Bonus = Val(TxtBonus.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!RetailPrice = Val(TxtRetailPrice.Text)
      RsBody!IsWSDiscb4ST = vIsWSDiscb4ST
      RsBody!IsWSSaleTax = vIsWSSaleTax
      RsBody!IsRetailSaleTax = vIsRetailSaleTax
      RsBody!Offer = IIf(Val(TxtOffer.Text) = 0, 0, Val(TxtOffer.Text))
      RsBody!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, 0, Val(TxtSaleTaxPer.Text))
      RsBody!SaleTaxval = IIf(Val(TxtSaleTaxVal.Text) = 0, 0, Val(TxtSaleTaxVal.Text))
      RsBody!DiscPC = IIf(Val(TxtDiscPC.Text) = 0, 0, Val(TxtDiscPC.Text))
      RsBody!DiscPer = IIf(Val(TxtDiscPer.Text) = 0, 0, Val(TxtDiscPer.Text))
      RsBody!DiscVal = IIf(Val(TxtDiscVal.Text) = 0, 0, Val(TxtDiscVal.Text))
      RsBody!Amount = Val(TxtAmount.Text)
      .MoveLast
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   
   QtyOffer = 0
   GetDataFromTextBoxesToGridOffer
   Call SubClearDetailArea
   'If TxtCode.Enabled = True And TxtCode.Visible = True Then TxtCode.SetFocus
   
   TxtTotalItems.Text = Val(Grid.Rows) - 1

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
   TxtBatchNo.Text = ""
   DtpExpiryDate.DateValue = ""
   TxtProductName.Text = ""
   CmbPackName.ListIndex = 0
   TxtMultiplier.Text = ""
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
                    GridOffer.Columns("ProductName").Text = CN.Execute("Select ProductName from products where productid = " & GridOffer.Columns("ProductOfferID").Text).Fields(0)
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
    RsProductOffer.Open "Select * from PurchaseOrderBodyOffer where OrderID =" & Val(TxtOrderID.Text) & " And OrderDate = '" & DtpOrderDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
    If RsProductOffer.RecordCount > 0 Then
    GridOffer.Visible = True
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
      DtpExpiryDate.DateValue = .Columns("ExpiryDate").Text
      TxtProductName.Text = .Columns("ProductName").Text
'      If Trim(.Columns("ColourName").Text) = "" Then
'         CmbColourName.ListIndex = 0
'      Else
'         CmbColourName.Text = .Columns("ColourName").Text
'      End If
'      If Trim(.Columns("SizeName").Text) = "" Then
'         cmbSizeName.ListIndex = 0
'      Else
'         cmbSizeName.Text = .Columns("SizeName").Text
'      End If
      CmbPackName.Clear
      vStrSQL = "select distinct pp.PackingID, Packingname from ProductPacking pp inner join packings p on p.packingid = pp.packingid" & vbCrLf _
           + "left outer join ProductBarcodes b on b.productid = pp.productid" & vbCrLf _
           + " where pp.productid = " & Val(TxtCode.Text) & " or code='" & TxtCode.Text & "'"
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
      TxtQtyLoose.Text = .Columns("QtyLoose").Text
      TxtQtyPack.Text = .Columns("QtyPack").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtRetailPrice.Text = .Columns("RetailPrice").Text
      vIsWSDiscb4ST = .Columns("IsWSDiscb4ST").Value
      vIsRetailSaleTax = .Columns("IsRetailSaleTax").Value
      vIsRetailSaleTax = .Columns("IsRetailSaleTax").Value
      TxtBonus.Text = .Columns("Bonus").Text
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtOffer.Text = .Columns("Offer").Value
      TxtSaleTaxPer.Text = .Columns("SaleTaxPer").Value
      TxtSaleTaxVal.Text = .Columns("SaleTaxVal").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      TxtAmount.Text = .Columns("Amount").Value
      
      If ObjRegistry.ShowAllPrices Then
         PopulateDataToPriceGrid
         FrmProductPrices.Visible = True
      End If
      
      If ObjRegistry.ShowMultiBranches Then
'         PopulateDataToGridBranch
         PopulateDataToGridBranchLive
         FrameMultiBranchStock.Visible = True
      End If
      
      If LblStock.Visible = False Then
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
      End If
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
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & GetServerDate + 1 & "',0),0)"
          With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               vQtyLoose = .Fields(0).Value
            Else
               vQtyLoose = 0
            End If
         End With
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         With CN.Execute("Select isnull(abbreviation,'') from packings where packingname = '" & CmbPackName.Text & "'")
            If .RecordCount > 0 Then
               LblStock.Caption = LblStock.Caption & " " & .Fields(0).Value
            Else
               LblStock.Caption = LblStock.Caption & " "
            End If
         End With
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         
         If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & ",Null,0,0,0,0,0,0,'" & DtpOrderDate.DateValue + 1 & "',0),0)"
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
         Else
            LblAllStock.Visible = False
         End If
         
      If Val(TxtMultiplier.Text) = 0 Then
         vUnitPrice = IIf(.Columns("Price").Text = "", 0, .Columns("Price").Text)
         vUnitRetailPrice = IIf(.Columns("RetailPrice").Text = "", 0, .Columns("RetailPrice").Text)
      Else
         vUnitPrice = .Columns("Price").Text / Val(TxtMultiplier.Text)
'         vUnitRetailPrice = .Columns("RetailPrice").Text / Val(TxtMultiplier.Text)
         vUnitRetailPrice = .Columns("RetailPrice").Text
      End If
      LblRetailPrice.Caption = vUnitRetailPrice
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetPurchase()
   On Error GoTo ErrorHandler
   ssql = "Select h.*, OrganizationName, p.partyname, Address, City, StoreName FROM PurchaseOrderHeader h join parties p on h.Vendorid = p.partyid inner join stores s on s.storeid = h.storeid left outer join Organizations o on o.OrganizationID = h.OrganizationID where h.OrderID=" & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'" & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With CN.Execute(ssql)
      If Not .BOF Then
          DtpOrderDate.DateValue = !OrderDate
          DtpPromiseDate.DateValue = !PromiseDate
          TxtBillID.Text = IIf(IsNull(!BillID), "", !BillID)
          DtpBillDate.DateValue = IIf(IsNull(!BillDate), "01/01/1990", !BillDate)
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

Private Sub TxtDiscVal_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDiscVal.Name Then Exit Sub
   If vUnitPrice = 0 Then Exit Sub
   If (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) = 0 Then Exit Sub
   TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)), 4)
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
   Call SubCalculateBody
'   TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), 2)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDiscVal_LostFocus()
'   Select Case ActiveControl.Name
'   Case TxtCode.Name, CmbPackName.Name, TxtMultiplier.Name, TxtBonus.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtOffer.Name, TxtSaleTaxPer.Name
'      Exit Sub
'   End Select
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
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub TxtQtyLoose_Validate(Cancel As Boolean)
If ActiveControl.Name <> TxtQtyLoose.Name Then Exit Sub
End Sub

Private Sub TxtQtyPack_Change()
   Call SubCalculateBody
'   Call FindRebate
End Sub

Private Sub TxtSaleTaxPer_Change()
If ActiveControl.Name <> TxtSaleTaxPer.Name Then Exit Sub
   Call SubCalculateBody
End Sub

Private Sub TxtSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridSerial.SetFocus
End Sub

Private Sub TxtSerial_LostFocus()
    GetDataFromTexBoxesToGridSerial
End Sub

Private Sub GetDataFromTexBoxesToGridSerial()
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
    With CN.Execute("Select  * from PurchaseOrderHeader where OrderID =" & TxtOrderID.Text & " And OrderDate = '" & DtpOrderDate.DateValue & "'")
'        If Val(TxtVenderID.Text) <> IIf(IsNull(!VENDORID), 0, !VENDORID) Then
'            CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Updated VenderID-" & !VENDORID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
'        End If
'        If TxtStoreID.Text <> !StoreID Then
'            CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
'        End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.Rows - 1
        With CN.Execute("Select * from PurchaseOrderBody Where OrderID = " & TxtOrderID.Text & " and OrderDate ='" & DtpOrderDate.DateValue & "' and Productid = " & Val(Grid.Columns("Productid").Text))
        
             If .EOF = True Then
                ssql = "Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")"
                CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("QtyLoose").Text <> !QtyLoose Or Grid.Columns("Price").Text <> !Price Or Grid.Columns("discper").Text <> !DiscPer Then
                   CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Updated ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
    
   Else
    CN.Execute ("Insert Into UserActivities values ('Purchase Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
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
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.PurchaseOrderHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmPurchaseOrder) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmPurchaseOrder & ", " & vUser & "," & TableHeaderFields(eFrmPurchaseOrder) & " from PurchaseOrderHeader " & vbCrLf _
             & "Where OrderID = " & TxtOrderID.Text & " and OrderDate = '" & DtpOrderDate.DateValue & "'"
      CN.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.PurchaseOrderBodyBin (" & TableBodyFields(eFrmPurchaseOrder) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmPurchaseOrder) & " from PurchaseOrderBody " & vbCrLf _
             & "Where OrderID = " & TxtOrderID.Text & " and OrderDate = '" & DtpOrderDate.DateValue & "'"
      CN.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGridBranch()
    On Error GoTo ErrorHandler
      vStrSQL = " SELECT CS.QtyLoose StockLoose,  " & vbCrLf _
             + " [1] Branch1, [2] Branch2, [3] Branch3, [4] Branch4, [5] Branch5, [6] Branch6, [7] Branch7, [8] Branch8, [9] Branch9" & vbCrLf _
             + " From CurrentStock CS Left Outer Join (" & vbCrLf _
             + " SELECT * FROM (" & vbCrLf _
             + " SELECT ProductID, StoreID, QtyLoose FROM CurrentSTockStore" & vbCrLf _
             + " ) AS SourceTable" & vbCrLf _
             + " PIVOT(" & vbCrLf _
             + " Sum(QtyLoose) FOR [StoreID] IN([1], [2], [3], [4], [5], [6], [7], [8], [9])" & vbCrLf _
             + " ) AS PivotStore )CSS On CSS.ProductID = CS.ProductID " & vbCrLf _
             + " where CS.productID = " & Val(TxtProductID.Text)
      
      With CN.Execute(vStrSQL)
         If Not .EOF Then
            GridBranch.Columns("Branch1").Value = IIf(IsNull(!Branch1), "", !Branch1)
            GridBranch.Columns("Branch2").Value = IIf(IsNull(!Branch2), "", !Branch2)
            GridBranch.Columns("Branch3").Value = IIf(IsNull(!Branch3), "", !Branch3)
            GridBranch.Columns("Branch4").Value = IIf(IsNull(!Branch4), "", !Branch4)
            GridBranch.Columns("Branch5").Value = IIf(IsNull(!Branch5), "", !Branch5)
            GridBranch.Columns("Branch6").Value = IIf(IsNull(!Branch6), "", !Branch6)
            GridBranch.Columns("Branch7").Value = IIf(IsNull(!Branch7), "", !Branch7)
            GridBranch.Columns("Branch8").Value = IIf(IsNull(!Branch8), "", !Branch8)
            GridBranch.Columns("Branch9").Value = IIf(IsNull(!Branch9), "", !Branch9)
            GridBranch.Columns("Stock").Value = IIf(IsNull(!StockLoose), "", !StockLoose)
         End If
         .Close
      End With
      
      Exit Sub
ErrorHandler:
   Call ShowErrorMessage

End Sub

Private Sub PopulateDataToGridBranchLive()
    On Error GoTo ErrorHandler
      
      vStrSQL = " Select * from Stores "
      vTotalQtyLoose = 0
      i = 0
      With CN.Execute(vStrSQL)
         While Not .EOF
            vError = 0
            vStrSQL = "exec " & !config & "sp_executesql N'SELECT isnull(dbo.FunStock(" & TxtProductID.Text & ",1,0,0,0,0,0,0,''" & DtpOrderDate.DateValue & "'',0),0)'"
            With CN.Execute(vStrSQL)
               GridBranch.Columns(i).Text = ""
               If vError = 0 Then
                  If .RecordCount > 0 Then
                     vQtyLoose = .Fields(0).Value
                     vTotalQtyLoose = vTotalQtyLoose + vQtyLoose
                     GridBranch.Columns(i).Value = vQtyLoose
                  End If
                  .Close
               End If
               i = i + 1
            End With
            .MoveNext
         Wend
         .Close
      End With
      GridBranch.Columns("Stock").Value = vTotalQtyLoose
      Exit Sub
ErrorHandler:
   If err.Number = -2147217900 Then
      vError = err.Number
      Resume Next
   End If
   Call ShowErrorMessage
End Sub


