VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form frmSaleInvoice 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmSaleInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   758
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmHistory 
      Height          =   1635
      Left            =   2115
      TabIndex        =   169
      Top             =   7155
      Visible         =   0   'False
      Width           =   9270
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridHistory 
         Height          =   1455
         Left            =   90
         TabIndex        =   170
         Top             =   135
         Width           =   9075
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
         stylesets(0).Picture=   "FrmSaleInvoice.frx":000C
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
         stylesets(1).Picture=   "FrmSaleInvoice.frx":0028
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
         stylesets(2).Picture=   "FrmSaleInvoice.frx":0044
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
         Columns(2).Width=   1799
         Columns(2).Caption=   "Date"
         Columns(2).Name =   "Date"
         Columns(2).CaptionAlignment=   2
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).NumberFormat=   "dd/MM/yyyy"
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "Expiry Date"
         Columns(3).Name =   "ExpiryDate"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).NumberFormat=   "dd/MM/yyyy"
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
         _ExtentX        =   16007
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
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   1350
      TabIndex        =   166
      Top             =   6840
      Visible         =   0   'False
      Width           =   2295
      Begin SITextBox.Txt TxtSerial 
         Height          =   315
         Left            =   120
         TabIndex        =   167
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
         TabIndex        =   168
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
         stylesets(0).Picture=   "FrmSaleInvoice.frx":0060
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
      TabIndex        =   155
      Top             =   6645
      Visible         =   0   'False
      Width           =   4095
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridExpiry 
         Height          =   1395
         Left            =   90
         TabIndex        =   156
         Top             =   90
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
         stylesets(0).Picture=   "FrmSaleInvoice.frx":007C
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
         stylesets(1).Picture=   "FrmSaleInvoice.frx":0098
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
         stylesets(2).Picture=   "FrmSaleInvoice.frx":00B4
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
         stylesets(3).Picture=   "FrmSaleInvoice.frx":00D0
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
      Left            =   10305
      TabIndex        =   139
      Top             =   9840
      Width           =   1245
      Begin VB.OptionButton OptCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFC09E&
         Caption         =   "Customer"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   141
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
         TabIndex        =   140
         Top             =   450
         Value           =   -1  'True
         Width           =   1530
      End
   End
   Begin VB.Frame FramExpense 
      Height          =   2415
      Left            =   8490
      TabIndex        =   121
      Top             =   6630
      Visible         =   0   'False
      Width           =   4215
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridExpense 
         Height          =   1860
         Left            =   120
         TabIndex        =   122
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
         stylesets(0).Picture=   "FrmSaleInvoice.frx":00EC
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
         stylesets(1).Picture=   "FrmSaleInvoice.frx":0108
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
         stylesets(2).Picture=   "FrmSaleInvoice.frx":0124
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
         TabIndex        =   125
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
         TabIndex        =   126
         Top             =   2100
         Width           =   1020
      End
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   11430
      MaxLength       =   50
      TabIndex        =   100
      Top             =   10575
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CheckBox ChkIsProduct 
      Caption         =   "Is Product"
      Height          =   255
      Left            =   3210
      TabIndex        =   94
      Top             =   1200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   4995
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   5535
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
      Left            =   12870
      TabIndex        =   68
      Top             =   855
      Visible         =   0   'False
      Width           =   4260
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
         TabIndex        =   69
         Tag             =   "NC"
         Text            =   "FrmSaleInvoice.frx":0140
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
         TabIndex        =   70
         Top             =   90
         Width           =   135
      End
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   1305
      TabIndex        =   0
      Top             =   1800
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
      Left            =   8490
      TabIndex        =   46
      Top             =   10605
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
      MICON           =   "FrmSaleInvoice.frx":0257
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7185
      TabIndex        =   42
      Top             =   10605
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
      MICON           =   "FrmSaleInvoice.frx":0273
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4575
      TabIndex        =   44
      Top             =   10605
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
      MICON           =   "FrmSaleInvoice.frx":028F
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9795
      TabIndex        =   47
      Top             =   10605
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
      MICON           =   "FrmSaleInvoice.frx":02AB
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5880
      TabIndex        =   43
      Top             =   10605
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
      MICON           =   "FrmSaleInvoice.frx":02C7
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   3315
      TabIndex        =   50
      Top             =   9495
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
      Left            =   5940
      TabIndex        =   39
      Top             =   9495
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
   Begin SITextBox.Txt TxtNetAmount 
      Height          =   315
      Left            =   9075
      TabIndex        =   52
      Top             =   9495
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
      Left            =   2070
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5535
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
      MICON           =   "FrmSaleInvoice.frx":02E3
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpBillDate 
      Height          =   315
      Left            =   1875
      TabIndex        =   1
      Top             =   1800
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
      Left            =   3240
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   1800
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
      Left            =   4275
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   1800
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
      Left            =   3915
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   1800
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
      MICON           =   "FrmSaleInvoice.frx":02FF
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3240
      TabIndex        =   45
      Top             =   10605
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
      MICON           =   "FrmSaleInvoice.frx":031B
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillDisc 
      Height          =   315
      Left            =   6915
      TabIndex        =   40
      Top             =   9495
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
      Left            =   4440
      TabIndex        =   62
      Top             =   1185
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
      Left            =   2430
      TabIndex        =   65
      Top             =   9495
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
      Left            =   7890
      TabIndex        =   41
      Top             =   9495
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
      TabIndex        =   21
      Top             =   5535
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
      TabIndex        =   23
      Top             =   5535
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   4
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
      TabIndex        =   27
      Top             =   5535
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
      TabIndex        =   26
      Top             =   5535
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
      Left            =   11310
      TabIndex        =   34
      Top             =   5535
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
      TabIndex        =   30
      Top             =   5535
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
      TabIndex        =   33
      Top             =   5535
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
      Left            =   9645
      TabIndex        =   31
      Top             =   5535
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
      TabIndex        =   28
      Top             =   5535
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
      TabIndex        =   29
      Top             =   5535
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
      Left            =   7170
      TabIndex        =   85
      Top             =   1200
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
      Left            =   10320
      TabIndex        =   32
      Top             =   5535
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
      Left            =   12690
      TabIndex        =   36
      Top             =   5535
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
      Left            =   9915
      TabIndex        =   17
      Top             =   4665
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
      Left            =   10950
      TabIndex        =   90
      Top             =   4665
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
      Left            =   10590
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   4665
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
      MICON           =   "FrmSaleInvoice.frx":0337
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCost 
      Height          =   315
      Left            =   6285
      TabIndex        =   95
      Top             =   1185
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
      Left            =   5400
      TabIndex        =   97
      Tag             =   "NC"
      Top             =   1185
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
      Left            =   10470
      TabIndex        =   101
      Top             =   9495
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
      Left            =   1305
      TabIndex        =   103
      Top             =   1200
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
      Left            =   1215
      TabIndex        =   96
      Top             =   7710
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
      stylesets(0).Picture=   "FrmSaleInvoice.frx":0353
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
      Left            =   1260
      TabIndex        =   12
      Top             =   4665
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
      Left            =   2010
      TabIndex        =   13
      Top             =   4665
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
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   315
      Left            =   1260
      TabIndex        =   11
      Top             =   3270
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
   Begin SITextBox.Txt TxtAddress 
      Height          =   315
      Left            =   1230
      TabIndex        =   105
      Top             =   3900
      Width           =   5205
      _ExtentX        =   9181
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
      Left            =   4275
      TabIndex        =   15
      Top             =   4665
      Width           =   3000
      _ExtentX        =   5292
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
   Begin SITextBox.Txt TxtRetailPrice 
      Height          =   315
      Left            =   7830
      TabIndex        =   116
      Top             =   1185
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
      Left            =   6435
      TabIndex        =   118
      Top             =   3900
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
   Begin SITextBox.Txt TxtTokenVal 
      Height          =   315
      Left            =   8550
      TabIndex        =   119
      Top             =   1185
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
      Left            =   2550
      TabIndex        =   123
      Top             =   3270
      Width           =   4320
      _ExtentX        =   7620
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
      Height          =   330
      Left            =   2190
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   3270
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
      MICON           =   "FrmSaleInvoice.frx":036F
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1215
      TabIndex        =   18
      Top             =   5535
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3240
      Left            =   675
      TabIndex        =   127
      Top             =   5850
      Width           =   13155
      ScrollBars      =   3
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   35
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
      stylesets(0).Picture=   "FrmSaleInvoice.frx":038B
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
      stylesets(1).Picture=   "FrmSaleInvoice.frx":03A7
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
      stylesets(2).Picture=   "FrmSaleInvoice.frx":03C3
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
      stylesets(3).Picture=   "FrmSaleInvoice.frx":03DF
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
      Columns.Count   =   35
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
      Left            =   1290
      TabIndex        =   6
      Top             =   2475
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
      Left            =   3330
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   2475
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
      MICON           =   "FrmSaleInvoice.frx":03FB
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   7275
      TabIndex        =   16
      Tag             =   "NC"
      Top             =   4665
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
      Left            =   8385
      TabIndex        =   131
      Tag             =   "NC"
      Top             =   4665
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
      Left            =   8025
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   4665
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
      MICON           =   "FrmSaleInvoice.frx":0417
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpOrderDate 
      Height          =   315
      Left            =   2025
      TabIndex        =   7
      Top             =   2475
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
      Left            =   8085
      TabIndex        =   135
      Tag             =   "NC"
      Top             =   1800
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
      Left            =   7725
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   1800
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
      MICON           =   "FrmSaleInvoice.frx":0433
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtVehicleNo 
      Height          =   315
      Left            =   2760
      TabIndex        =   14
      Top             =   4665
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
   Begin SITextBox.Txt TxtBatchNo 
      Height          =   315
      Left            =   2205
      TabIndex        =   19
      Top             =   5220
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
      TabIndex        =   142
      Top             =   5220
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
      Left            =   7980
      TabIndex        =   143
      Top             =   10215
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
      Left            =   5010
      TabIndex        =   144
      Top             =   10215
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
      Left            =   6585
      TabIndex        =   145
      Top             =   10215
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
      Left            =   9375
      TabIndex        =   146
      Top             =   10215
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
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   5205
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
      MICON           =   "FrmSaleInvoice.frx":044F
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrintWarranty 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   1800
      TabIndex        =   149
      Top             =   10605
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
      MICON           =   "FrmSaleInvoice.frx":046B
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPurchase 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5805
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   2475
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
      MICON           =   "FrmSaleInvoice.frx":0487
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtPurID 
      Height          =   315
      Left            =   3765
      TabIndex        =   8
      Top             =   2475
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
      Left            =   4515
      TabIndex        =   9
      Top             =   2475
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
      Left            =   3555
      TabIndex        =   151
      Top             =   10065
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
      MICON           =   "FrmSaleInvoice.frx":04A3
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnBatchPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11655
      TabIndex        =   152
      Top             =   9930
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
      MICON           =   "FrmSaleInvoice.frx":04BF
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpPromiseDate 
      Height          =   315
      Left            =   5670
      TabIndex        =   4
      Top             =   1800
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
      Left            =   6210
      TabIndex        =   10
      Top             =   2490
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
      Left            =   7275
      TabIndex        =   157
      Top             =   2490
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
      Left            =   6915
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   2490
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
      MICON           =   "FrmSaleInvoice.frx":04DB
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   7020
      TabIndex        =   5
      Tag             =   "NC"
      Top             =   1800
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
      Left            =   6870
      TabIndex        =   163
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   3270
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
      MICON           =   "FrmSaleInvoice.frx":04F7
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtContactNo 
      Height          =   315
      Left            =   8205
      TabIndex        =   164
      Top             =   3900
      Width           =   4080
      _ExtentX        =   7197
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
      Left            =   4500
      TabIndex        =   37
      Top             =   9495
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
      Left            =   5340
      TabIndex        =   38
      Top             =   9495
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
      TabIndex        =   35
      Top             =   5535
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
   Begin SITextBox.Txt TxtGrossQty 
      Height          =   315
      Left            =   6900
      TabIndex        =   24
      Top             =   4995
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
   Begin SITextBox.Txt TxtGrossUnit 
      Height          =   315
      Left            =   8505
      TabIndex        =   25
      Top             =   4995
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
   Begin VB.Label LblGrossUnit 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Unit"
      Height          =   195
      Left            =   7695
      TabIndex        =   176
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label LblGrossQty 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Qty"
      Height          =   195
      Left            =   6120
      TabIndex        =   175
      Top             =   5040
      Width           =   690
   End
   Begin VB.Label LblSC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "S.C."
      Height          =   195
      Left            =   12015
      TabIndex        =   174
      Top             =   5340
      Width           =   300
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Ch."
      Height          =   195
      Left            =   4500
      TabIndex        =   173
      Top             =   9270
      Width           =   825
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   5340
      TabIndex        =   172
      Top             =   9270
      Width           =   210
   End
   Begin VB.Label LblTotalAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amount"
      Height          =   195
      Left            =   3345
      TabIndex        =   171
      Top             =   9270
      Width           =   990
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      Height          =   195
      Left            =   8235
      TabIndex        =   165
      Top             =   3690
      Width           =   855
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Date"
      Height          =   195
      Left            =   4515
      TabIndex        =   162
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur ID"
      Height          =   195
      Left            =   3765
      TabIndex        =   161
      Top             =   2280
      Width           =   450
   End
   Begin VB.Label LblSyllabusID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Syllabus ID"
      Height          =   195
      Left            =   6210
      TabIndex        =   160
      Top             =   2295
      Width           =   795
   End
   Begin VB.Label LblSyllabusName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Syllabus Name"
      Height          =   195
      Left            =   7275
      TabIndex        =   159
      Top             =   2295
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
      Left            =   3960
      TabIndex        =   154
      Top             =   5115
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblPromiseDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Promise Date"
      Height          =   195
      Left            =   5670
      TabIndex        =   153
      Top             =   1530
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
      TabIndex        =   147
      Top             =   5070
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblFreight 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      Height          =   195
      Left            =   9390
      TabIndex        =   138
      Top             =   9990
      Width           =   480
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      Height          =   195
      Left            =   2745
      TabIndex        =   137
      Top             =   4455
      Width           =   780
   End
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   8370
      TabIndex        =   134
      Top             =   4455
      Width           =   780
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   7260
      TabIndex        =   133
      Top             =   4455
      Width           =   525
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID"
      Height          =   195
      Left            =   1305
      TabIndex        =   129
      Top             =   2280
      Width           =   600
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
      Height          =   195
      Left            =   1995
      TabIndex        =   128
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Token Val"
      Height          =   195
      Left            =   8550
      TabIndex        =   120
      Top             =   1005
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ret Price"
      Height          =   195
      Left            =   7830
      TabIndex        =   117
      Top             =   1005
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Received Amount"
      Height          =   195
      Left            =   7980
      TabIndex        =   115
      Top             =   9990
      Width           =   1275
   End
   Begin VB.Label lblPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Receivable"
      Height          =   195
      Left            =   5025
      TabIndex        =   114
      Top             =   9990
      Width           =   1470
   End
   Begin VB.Label LblTtlPayable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Receivable"
      Height          =   195
      Left            =   6585
      TabIndex        =   113
      Top             =   9990
      Width           =   1215
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      Height          =   195
      Left            =   1260
      TabIndex        =   112
      Top             =   4455
      Width           =   495
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bilty No."
      Height          =   195
      Left            =   1995
      TabIndex        =   111
      Top             =   4455
      Width           =   585
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      Height          =   195
      Left            =   1275
      TabIndex        =   110
      Top             =   3060
      Width           =   870
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   195
      Left            =   2580
      TabIndex        =   109
      Top             =   3060
      Width           =   1125
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   1230
      TabIndex        =   108
      Top             =   3690
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   6435
      TabIndex        =   107
      Top             =   3690
      Width           =   255
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   4275
      TabIndex        =   106
      Top             =   4455
      Width           =   795
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill No"
      Height          =   195
      Left            =   1305
      TabIndex        =   104
      Top             =   975
      Width           =   1020
   End
   Begin VB.Label LblRemarks 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   10410
      TabIndex        =   102
      Top             =   9270
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   6330
      TabIndex        =   99
      Top             =   960
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      Height          =   195
      Left            =   5325
      TabIndex        =   98
      Top             =   960
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblMemberID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      Height          =   195
      Left            =   9900
      TabIndex        =   93
      Top             =   4455
      Width           =   780
   End
   Begin VB.Label LblMemberName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      Height          =   195
      Left            =   10935
      TabIndex        =   92
      Top             =   4455
      Width           =   1035
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   8205
      TabIndex        =   89
      Top             =   1605
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   7020
      TabIndex        =   88
      Top             =   1605
      Width           =   1095
   End
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   12660
      TabIndex        =   87
      Top             =   5340
      Width           =   540
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax%"
      Height          =   195
      Left            =   10320
      TabIndex        =   86
      Top             =   5340
      Width           =   390
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Val"
      Height          =   195
      Left            =   7170
      TabIndex        =   84
      Top             =   1005
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Offer"
      Height          =   195
      Left            =   8490
      TabIndex        =   83
      Top             =   5340
      Width           =   345
   End
   Begin VB.Label LblPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   9000
      TabIndex        =   82
      Top             =   5340
      Width           =   405
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   6450
      TabIndex        =   81
      Top             =   5340
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   5025
      TabIndex        =   80
      Top             =   5340
      Width           =   840
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   7440
      TabIndex        =   79
      Top             =   5340
      Width           =   465
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   6930
      TabIndex        =   78
      Top             =   5340
      Width           =   480
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.Val"
      Height          =   195
      Left            =   11310
      TabIndex        =   77
      Top             =   5340
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Dis%"
      Height          =   195
      Left            =   10830
      TabIndex        =   76
      Top             =   5340
      Width           =   345
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc/PC"
      Height          =   195
      Left            =   9645
      TabIndex        =   75
      Top             =   5340
      Width           =   600
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bns(L)"
      Height          =   195
      Left            =   7980
      TabIndex        =   74
      Top             =   5340
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
      Left            =   9405
      TabIndex        =   73
      Top             =   3375
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
      Left            =   7695
      TabIndex        =   72
      Top             =   3390
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
      Left            =   12240
      TabIndex        =   71
      Top             =   1470
      Width           =   435
   End
   Begin VB.Label LblOtherChargesCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      Height          =   195
      Left            =   7890
      TabIndex        =   67
      Top             =   9270
      Width           =   1020
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items"
      Height          =   195
      Left            =   2445
      TabIndex        =   66
      Top             =   9270
      Width           =   780
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Invoice"
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
      TabIndex        =   64
      Top             =   270
      Width           =   2160
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   4440
      TabIndex        =   63
      Top             =   960
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   6915
      TabIndex        =   61
      Top             =   9270
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
      Left            =   10125
      TabIndex        =   60
      Top             =   1035
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
      Left            =   9990
      TabIndex        =   59
      Top             =   1290
      Width           =   1035
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   4275
      TabIndex        =   58
      Top             =   1605
      Width           =   840
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   3240
      TabIndex        =   57
      Top             =   1605
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   1215
      TabIndex        =   55
      Top             =   5340
      Width           =   375
   End
   Begin VB.Label LblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3870
      TabIndex        =   54
      Top             =   5340
      Width           =   1020
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   12570
      Top             =   1358
      Width           =   330
   End
   Begin VB.Label LblNetAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      Height          =   195
      Left            =   9090
      TabIndex        =   53
      Top             =   9270
      Width           =   840
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount (%)"
      Height          =   195
      Left            =   5925
      TabIndex        =   51
      Top             =   9270
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   1875
      TabIndex        =   49
      Top             =   1605
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill ID"
      Height          =   195
      Left            =   1320
      TabIndex        =   48
      Top             =   1605
      Width           =   405
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
Attribute VB_Name = "frmSaleInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public cnSale As New ADODB.Connection
Dim Application1 As New CRAXDRT.Application
Dim vMode As FormMode
Dim vUnitPrice As Double
Dim vDate As Date, vHDiff As Integer, vSystemDate As Boolean
Dim vUnitRetailPrice As Double
Dim vIsWSDiscb4ST As Boolean
Dim vIsRetailSaleTax As Boolean
Dim vIsWSSaleTax As Boolean
Dim vIsNewRecord As Boolean
Dim vCounter As Integer, isUrdu As Boolean
Dim vBm As Variant, vExpiryTime As String
Dim vMaxBinID As Integer
Dim RsBody As New ADODB.Recordset
Dim RsExpense As New ADODB.Recordset
Dim RsBodySerial As New ADODB.Recordset
Dim RsProductOffer As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim QtyOffer As Integer
Dim Rebate As Integer
Dim DateFlag As Boolean
Dim Flag As Boolean, vAlreadySerial As Boolean
Dim ssql As String
Dim VStrSQL As String
Dim vBillID  As Integer, vZoneID As Byte
Dim vBillDate  As Date
Dim ExpenseFlag As Boolean, vAutoEnterQtyintoGridSaleInvoice As Boolean
Dim vExpAmount As Double
Dim vQtyLoose As Double, vTotalAmount As Double, vAmount As Double
Dim vStrPara As String
Dim i As Integer, vNoofPrints As Byte, isWholeSale As Boolean
Dim vCash, vCredit As Integer
Dim vMasterID As Integer
Dim vStrDetail As String
Dim vMobileNo() As String, vMobile As String
Dim vUpdateStock As Boolean
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
End Sub

Private Sub SubCalculateFooter()
   If TxtTotalAmount.Text = "" Then Exit Sub
   TxtNetAmount.Text = SelfRound(Val(TxtTotalAmount.Text) - Val(TxtBillDisc.Text)) + Val(TxtOtherCharges.Text) + Val(TxtTotalExpense.Text) + Val(TxtServiceCharges.Text)
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
    With cn.Execute(VStrSQL)
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
   Dim VStrSQL As String, vStr As String
   'vStr = " and p.ProductID not In (Select ProductID from ZoneAllotment where ZoneID = " & vZoneID & " and PartyID is Not Null and PartyID <> '" & TxtCustomerID.Text & "')"
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.ParaInWholeSale = True
      SchProduct.ParaInWhere = " and isLocked = 0" & vStr
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
   If Trim(TxtCode.Text) = "" Then Exit Function
   If IsNumeric(TxtCode.Text) = True Then
      If Len(TxtCode.Text) < 5 Then
         TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
      End If
   End If
   With cn.Execute("Select productid, serial from purchasebodyserial where serial = '" & TxtCode.Text & "'")
      If .RecordCount > 0 Then
         TxtCode.Text = !Productid
         TxtSerial.Text = !Serial
      End If
      .Close
   End With
   
   CmbPackName.Clear
   VStrSQL = "select * from ProductPacking pp inner join packings p on p.packingid = pp.packingid" & vbCrLf _
           + "left outer join ProductBarcodes b on b.productid = pp.productid" & vbCrLf _
           + " where pp.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
   With cn.Execute(VStrSQL)
      CmbPackName.AddItem ""
      While Not .EOF
         CmbPackName.AddItem !Packingname
         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
         .MoveNext
      Wend
      .Close
   End With
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
        VStrSQL = "SELECT p.productid, Code, ProductName, PurPrice, WSPrice, RetailPrice, " & vbCrLf _
           + " IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST, TokenVal, SaleTaxPer, DiscPC, ServiceCharges, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.Salepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where isLocked = 0 and (p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "')" & vStr
 
   With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = IIf(isWholeSale = True, !WSPrice, !RetailPrice)
         TxtRetailPrice.Text = !RetailPrice
         ssql = "select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "','" & TxtProductID.Text & "')"
         LblLastPurPrice.Caption = cn.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "','" & TxtProductID.Text & "')").Fields(0).Value
         vIsWSDiscb4ST = !IsWSDiscb4ST
         vIsWSSaleTax = !IsWSSaleTax
         vIsRetailSaleTax = !IsRetailSaleTax
         TxtSaleTaxPer.Text = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
         TxtTokenVal.Text = IIf(IsNull(!TokenVal), "", !TokenVal)
         LblRetailPrice.Caption = !RetailPrice
         TxtSC.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         
        With cn.Execute("select cost from currentstock where productid ='" & TxtProductID.Text & "'")
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
         ChkIsProduct.Value = 1
         If vAutoEnterQtyintoGridSaleInvoice = True Then TxtQtyLoose.Text = IIf(Val(TxtQtyLoose.Text) = 0, 1, TxtQtyLoose.Text)
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
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), "", !DiscPC)
         If vUnitPrice = 0 Then
            TxtDiscPer.Text = "0"
         Else
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / vUnitPrice, 3)
         End If
         If ObjRegistry.AlertAllocateProduct = True Then
            If Trim(TxtCustomerID.Text) <> "" Then
               VStrSQL = "Select * " & vbCrLf _
                     + " from CustomerProductPrice" & vbCrLf _
                     + " where CustomerID = '" & TxtCustomerID.Text & "' and ProductID = '" & TxtProductID.Text & "'"
   
               With cn.Execute(VStrSQL)
                  If .RecordCount > 0 Then
                     TxtPrice.Text = !Price
                     vUnitPrice = !Price
                     TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
                     TxtDiscPC.Text = Round((vUnitPrice * Val(TxtDiscPer.Text) / 100), 4)
                  Else
                     MsgBox "Allocate this Product to the Customer first.", vbInformation + vbOKOnly, "Information"
                     FunSelectProduct = False
                     Exit Function
                  End If
               End With
            End If
         End If
         If ObjRegistry.AutoApplyPartyLastPrice Then
            If Trim(TxtCustomerID.Text) <> "" Then
               VStrSQL = "Select Top 1 * " & vbCrLf _
                     + " from SaleHeader h inner join SaleBody b on h.BillID = b.BillID and h.BillDate = b.BillDate" & vbCrLf _
                     + " left outer join ProductPacking pp on pp.packingid = b.packingid and pp.productid = b.productid" & vbCrLf _
                     + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
                     + " where h.CustomerID = '" & TxtCustomerID.Text & "' and b.ProductID = '" & TxtProductID.Text & "'" & vbCrLf _
                     + " Order by h.BillDate Desc, h.BillID desc"
               With cn.Execute(VStrSQL)
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
                  End If
               End With
            End If
         End If
         If ObjRegistry.AutoApplyPartyLastDiscount Then
            If Trim(TxtCustomerID.Text) <> "" Then
               VStrSQL = "Select Top 1 * " & vbCrLf _
                     + " from SaleHeader h inner join SaleBody b on h.BillID = b.BillID and h.BillDate = b.BillDate" & vbCrLf _
                     + " left outer join ProductPacking pp on pp.packingid = b.packingid and pp.productid = b.productid" & vbCrLf _
                     + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
                     + " where h.CustomerID = '" & TxtCustomerID.Text & "' and b.ProductID = '" & TxtProductID.Text & "'" & vbCrLf _
                     + " Order by h.BillDate Desc, h.BillID desc"
               With cn.Execute(VStrSQL)
                  If .RecordCount > 0 Then
'                     TxtDiscPC.Text = !DiscPC
                     TxtDiscPer.Text = !DiscPer
                  End If
               End With
            End If
         End If
         
'         ''' latest Comment
         If ObjRegistry.ShowSavedStock = True Then
            VStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = '" & TxtProductID.Text & "'"
            With cn.Execute(VStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            VStrSQL = "select isnull(dbo.FunStock('" & TxtProductID.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(VStrSQL).Fields(0).Value
         End If
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Caption = LblStock.Caption & " " & " Total Qty: " & vQtyLoose
         LblStock.Visible = True
         LblStockCaption.Visible = True
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
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
         With cn.Execute("Select dbo.GetExpiryTime('" & TxtProductID.Text & "', " & IIf(TxtBatchNo.Text = "", "Null", "'" & TxtBatchNo.Text & "'") & " , GetDate()) as Day ")
            If .RecordCount > 0 Then
               vExpiryTime = !Day
            End If
         End With
         If TxtCustomerID.Text <> "" Then
            PopulateDataToHistoryGrid
            FrmHistory.Visible = True
            FrmHistory.ZOrder 0
            GridHistory.Visible = True
            GridHistory.ZOrder 0
         End If
         SubCalculateBody
''         VStrSQL = "select isnull(dbo.FunStock('" & TxtProductID.Text & "'," & TxtStoreID.Text & "," & Val(TxtBillID.Text) & "," & Val(0) & "," & Val(TxtBillID.Text) & "," & Val(0) & "," & Val(0) & "," & Val(0) & ",'" & DateAdd("D", 1, DtpBillDate.DateValue) & "'," & Val(0) & "),0)"
''         vQtyLoose = cn.Execute(VStrSQL).Fields(0).Value
''         LblStock.Caption = vQtyLoose
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

Private Sub PopulateDataToHistoryGrid()
      ssql = "select top 3 pt.PartyName, CustomerID, code, b.* " & vbCrLf & _
      " from SaleHeader h inner join Salebody b on h.billID = b.BillID and h.BillDate = b.BillDate" & vbCrLf & _
      " inner join Parties pt on pt.PartyID = h.CustomerID " & vbCrLf & _
      " where h.CustomerID = '" & TxtCustomerID.Text & "' and b.productid = '" & (TxtProductID.Text) & "' order by b.BillDate Desc"
      
      With cn.Execute(ssql)
         GridHistory.Redraw = False
         GridHistory.MoveFirst
         GridHistory.RemoveAll
         GridHistory.AllowAddNew = True
         While Not .EOF
            GridHistory.AddNew
            GridHistory.Columns("ID").Text = !CustomerID
            GridHistory.Columns("Name").Text = !PartyName
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
      For vCounter = 1 To .Rows
         If Trim(.Columns("Productid").Text) <> "" Then
            vStrDetail = vStrDetail & " (P" & .Columns("ProductID").Text & IIf(Val(.Columns("Pack").Value) = 0, "", " M" & .Columns("Pack").Value) & IIf(Val(.Columns("QtyPack").Value) = 0, "", " QP" & .Columns("QtyPack").Value) & IIf(Val(.Columns("QtyLoose").Value) = 0, "", " QL" & .Columns("QtyLoose").Value) & IIf(Val(.Columns("Bonus").Value) = 0, "", " QB" & .Columns("Bonus").Value) & " A" & .Columns("Amount").Text & ")"
         End If
         .MoveNext
      Next vCounter
      .Redraw = True
   End With
    '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave And vIsNewRecord = True And Grid.Rows > 1 Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = " Cleared ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt" & TxtNetAmount.Text
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               cn.Execute ssql
            End If
         Next
   End If
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
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   cn.BeginTrans
   
   
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
  
   ''''''''''''''''''''''''''Delete Product OFfer'''''''''''''''''''''
   GridOffer.Redraw = False
   GridOffer.MoveFirst
   For vCounter = 1 To GridOffer.Rows
      If Trim(GridOffer.Columns("Productid").Text) <> "" Then
         cn.Execute "Delete from SaleBodyOffer where BillID = " & Val(TxtBillID.Text) & " And BillDate ='" & DtpBillDate.DateValue & "' and productid ='" & GridOffer.Columns("Productid").Text & "'"
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
            cn.Execute "Delete from SaleBodySerial where BillID = " & Val(TxtBillID.Text) & " And BillDate ='" & DtpBillDate.DateValue & "' and productid ='" & RsBodySerial!Productid & "' and Serial ='" & RsBodySerial!Serial & "'"
            RsBodySerial.MoveNext
        Next vCounter
    End If
   ''''''''''''''''''''''''''Delete Sale Body'''''''''''''''''''''
   vStrDetail = ""
   Grid.Redraw = False
   Grid.MoveFirst
   Call ActivityLogSale("Sale Invoice", eDelete, TxtBillID.Text, DtpBillDate.DateValue)
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         VStrSQL = "Delete from SaleBody where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "' and productid ='" & Grid.Columns("ProductID").Text & "' and BatchNo " & IIf(Trim(Grid.Columns("BatchNo").Text) = "", " is null", " = '" & Trim(Grid.Columns("BatchNo").Text) & "'") & " and Price = " & Val(Grid.Columns("Price").Text) & " and EmpID " & IIf(Trim(Grid.Columns("EmpID").Text) = "", " is null", " = '" & Trim(Grid.Columns("EmpID").Text) & "'") & " and StoreID =" & Val(TxtStoreID.Text)
         cn.Execute VStrSQL
         vQtyLoose = (Val(Grid.Columns("Pack").Text) * Val(Grid.Columns("QtyPack").Text)) + Val(Grid.Columns("QtyLoose").Text) + Val(Grid.Columns("Bonus").Text)
         cn.Execute "Exec UpdateStockPlus " & TxtStoreID.Text & ",'" & Grid.Columns("ProductID").Text & "'," & vQtyLoose & "," & Val(TxtBillID.Text) & ",'" & DtpBillDate.DateValue & "'"
         
         vStrDetail = vStrDetail & " (P" & Grid.Columns("ProductID").Text & IIf(Val(Grid.Columns("Pack").Value) = 0, "", " M" & Grid.Columns("Pack").Value) & IIf(Val(Grid.Columns("QtyPack").Value) = 0, "", " QP" & Grid.Columns("QtyPack").Value) & IIf(Val(Grid.Columns("QtyLoose").Value) = 0, "", " QL" & Grid.Columns("QtyLoose").Value) & IIf(Val(Grid.Columns("Bonus").Value) = 0, "", " QB" & Grid.Columns("Bonus").Value) & " A" & Grid.Columns("Amount").Text & ")"
'          cn.Execute ("Insert Into Bin_SaleBody Select " & FunGetMaxBinID & ", * from SaleBody Where BillID = " & TxtBillID.Text & " And BillDate ='" & DtpBillDate.DateValue & "' and productid ='" & Grid.Columns("Productid").Text & "'")
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   '''''''''''''''''''''''''''''''''''''''Delete Expense'''''''''''''''''''''''''''''''''''''''
   cn.Execute "Delete from SaleExpense where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   
   '''''''''''''''''''''''''''''''''''''''Delete Header'''''''''''''''''''''''''''''''''''''''
   cn.Execute "Delete from SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   
   cn.Execute ("Update SaleOrderHeader set IsSale = 0 Where OrderID = " & Val(TxtOrderID.Text) & "And Orderdate ='" & DtpOrderDate.DateValue & "'")
    If Grid.Rows > 1 Then
      VStrSQL = "INSERT INTO ActivityLog(userno,FormType,EntryDate,Description,isnew,isedit,isdelete,isClear) values(" & vUser & ",'Sale Invoice', GetDate()," & "'BillID = " & TxtBillID.Text & " BillDate = " & DtpBillDate.DateValue & " Clear' ,0,0,0,1" & ")"
    'cnPOS.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Clear'," & vUser & ")")
      cn.Execute (VStrSQL)
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
               cn.Execute ssql
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
               cn.Execute ssql
            End If
         Next
   End If
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub
Private Sub Sub_Bin_Save()
  On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
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
         
         RsBody!RetailPrice = 0
         RsBody!IsWSDiscb4ST = 0
         RsBody!IsWSSaleTax = 0
         RsBody!IsRetailSaleTax = 0
         
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

Private Sub BtnPrintWarranty_Click()
   On Error GoTo ErrorHandler
   With cn.Execute("Select distinct o.OrganizationID, ReportName from SaleBody b inner join Products p on p.productid = b.productid inner join Organizations o on p.OrganizationID = o.OrganizationID where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'")
      For i = 1 To .RecordCount
         VStrSQL = "Select h.BillID, h.BillDate, EntryDate, h.OrganizationID, OrganizationName, Customerid, isnull(Pr.PartyName,AccountName) + ' - ' + H.CustomerID as Customer_Name_ID, Pr.Address, StoreName, BiltyNo, VehicleNo, h.Description," & vbCrLf _
            + " Isnull(H.BillDiscPer, 0) BillDiscPer, Isnull(H.BillDisc,0) BillDisc, isnull(OtherCharges,0) as OtherCharges," & vbCrLf _
            + " TotalAmount,  isnull(TotalExpense,0) as TotalExpense, b.ProductID as Code, " & IIf(isUrdu = True, "p.ProductName1", "p.ProductName") & " as ProductName, dbo.FunSaleBodySerial(b.BillID,b.BillDate, b.ProductId) Serial," & vbCrLf _
            + " dbo.FunSaleBodyOffer(b.BillID,b.BillDate, b.ProductId) ProductOffer, isnull(QtyPack,0)QtyPack, isnull(Multiplier,0)Multiplier, Qty," & vbCrLf _
            + " Bonus,b.DiscPc, b.DiscPer, DiscVal, Offer, b.SaleTaxPer, SaleTaxval," & vbCrLf _
            + " h.Empid, empname, price, Amount, previousAmount, CashReceived, b.RetailPrice, b.BatchNo " & vbCrLf _
            + " from SaleBody b inner join SaleHeader h on h.BillID = b.BillID and h.BillDate = b.BillDate" & vbCrLf _
            + " inner join products p on b.productid = p.productid" & vbCrLf _
            + " left outer join Organizations o on o.OrganizationID = p.OrganizationID" & vbCrLf _
            + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
            + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
            + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
            + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate='" & DtpBillDate.DateValue & "' and p.OrganizationID = " & !OrganizationID
       
          If RsReport.State = adStateOpen Then RsReport.Close
          RsReport.Open VStrSQL, cn, adOpenStatic, adLockReadOnly
         
          RptReportViewer.Report.SelectPrinter "abc", "xyz", "ghi"
          
          
          
          Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\" & !ReportName & ".rpt")
          'Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\SaleInvoiceWarranty1.rpt")
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
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchPurchase.ParaInPurchasedate = DtpPurchaseDate.DateValue
        SchPurchase.Show vbModal, Me
        If SchPurchase.ParaOutPurchaseID = "" Then FunSelectPurchase = False: Exit Function
        TxtPurID.Text = SchPurchase.ParaOutPurchaseID
        DtpPurchaseDate.DateValue = SchPurchase.ParaOutPurchaseDate
    End If
    '---------------------------
    VStrSQL = "Select * from PurchaseHeader where PurID=" & Val(TxtPurID.Text) & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
    With cn.Execute(VStrSQL)
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
   With cn.Execute(ssql)
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
   RsBody.Open "Select * from SaleBody where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      
'      ssql = " select pb.ProductID, ProductName, QtyPack - isnull(UPack,0) as RQtyPack, Qtyloose - isnull(UQty,0) as RQty, Bonus - isnull(UBonus,0) as RBonus, pb.*" & vbCrLf _
      + " from (select b.PurID, b.PurchaseDate, ProductID, Sum(Qtyloose) as UQty, Sum(QtyPack) as UPack, Sum(Bonus) as UBonus from PurchaseBody b inner join PurchaseHeader h on h.PurID = b.pURID and h.PurchaseDate = b.PurchaseDate Group By b.PurID, b.PurchaseDate, ProductID) b " & vbCrLf _
      + " right outer join PurchaseBody  pb on pb.PurID = b.PurID and pb.PurchaseDate = b.PurchaseDate and b.ProductID = pb.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = pb.productid" & vbCrLf _
      + " where pb.PurID = " & Val(TxtPurID.Text) & " and pb.PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
      ssql = "select p.productname, p.Retailprice product_RetailPrice, b.* from PurchaseBody b join products p on p.productid = b.productid where PurID=" & Val(TxtPurID.Text) & " and PurchaseDate = '" & DtpPurchaseDate.DateValue & "'"
      With cn.Execute(ssql)
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
               Grid.Columns("PackName").Text = cn.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
            End If
            
            Grid.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            Grid.Columns("QtyLoose").Value = !QtyLoose
            Grid.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            Grid.Columns("Price").Value = !product_RetailPrice
            Grid.Columns("Cost").Value = 0
            Grid.Columns("isProduct").Value = 1 '!isProduct
            Grid.Columns("RetailPrice").Value = !product_RetailPrice
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
            Grid.Columns("Amount").Value = ((!product_RetailPrice / Val(IIf(IsNull(!Multiplier), "1", !Multiplier)))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !QtyLoose)  '!Amount
'            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(((!product_RetailPrice / Val(IIf(IsNull(!Multiplier), "1", !Multiplier))) - Val(IIf(IsNull(!DiscPC), "0", !DiscPC))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !QtyLoose))
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(((!product_RetailPrice / Val(IIf(IsNull(!Multiplier), "1", !Multiplier)))) * (IIf(IsNull(!QtyPack), 0, !QtyPack) * IIf(IsNull(!Multiplier), "0", !Multiplier) + !QtyLoose))
            TxtTotalItems.Text = Val(TxtTotalItems.Text) + !QtyLoose + IIf(IsNull(!Bonus), "0", !Bonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!QtyPack), 0, !QtyPack))
            
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

Private Sub DtpPromiseDate_Change()
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub DtpPromiseDate_DblClick()
   DtpPromiseDate.DateValue = Null
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Or KeyCode = vbKeyF4 Then
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

Private Sub GridExpense_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
         TxtQtyPack.Text = cn.Execute("SELECT dbo.FunGetPack('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         TxtQtyLoose.Text = cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "'," & vQtyLoose & ")").Fields(0).Value
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
   
   With cn.Execute("Select dbo.GetExpiryDate('" & TxtProductID.Text & "','" & TxtBatchNo.Text & "') as ExpiryDate")
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
      With cn.Execute("SELECT dbo.FunLastManualBillNo('%" & TxtBillNo.Text & "%')")
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
   If vTemp = True Then
      vTemp = Not FunSelectCustomer(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
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
    VStrSQL = "Select c.AccountNo, c.AccountName as AccountName, Address, City, p.phone1, p.phone2, p.mobile, p.mobile2, p.Description, isnull(p.isWholeSale,1) as isWholeSale" & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " where c.AccountNo = '" & (TxtCustomerID.Text) & "' and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and isDetailed = 1 and isLocked = 0"
    
    VStrSQL = VStrSQL + " union all Select EmpID, EmpName as AccountName, Address, City, '' phone1,'' phone2, '' mobile, '' mobile2, '', 1 as isWholeSale" & vbCrLf _
         + " from Employees" & vbCrLf _
         + " where EmpID = '" & (TxtCustomerID.Text) & "' and isLockEmployee = 0"
    
    With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerName.Text = !AccountName
          TxtAddress.Text = IIf(IsNull(!Address), "", !Address)
          TxtCity.Text = IIf(IsNull(!City), "", !City)
          TxtContactNo.Text = IIf(IsNull(!Phone1), "", !Phone1 & " ") & IIf(IsNull(!Phone2), "", !Phone2 & " ") & IIf(IsNull(!Mobile), "", !Mobile & " ") & IIf(IsNull(!Mobile2), "", !Mobile2)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtPreviousReceivable.Text = cn.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtCustomerID.Text & "','" & DtpBillDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
          VStrSQL = " Select isnull(Sum(round(B.TTLValue,0) - isnull(BillDisc,0) + isnull(OtherCharges,0) + Isnull(TotalExpense,0) + isnull(servicecharges,0) + isnull(STax,0)),0) as Amount " & vbCrLf _
                  + " FROM SaleHeader h INNER JOIN (Select BillId, BillDate, Sum(Amount) TTLValue FROM SaleBody Group By BillId, BillDate)b " & vbCrLf _
                  + " ON h.BillId = B.BillId and h.BillDate = B.BillDate " & vbCrLf _
                  + " where CustomerID = '" & (TxtCustomerID.Text) & "' and h.BillDate = '" & DtpBillDate.DateValue & "' and h.BillID >= " & Val(TxtBillID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
          TxtPreviousReceivable.Text = TxtPreviousReceivable.Text - cn.Execute(VStrSQL).Fields(0).Value
          lblPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
          TxtPreviousReceivable.Text = Abs(TxtPreviousReceivable.Text)
          vZoneID = cn.Execute("SELECT isnull(dbo.FunGetZoneID('" & TxtCustomerID.Text & "'),0)").Fields(0).Value
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
    With cn.Execute(ssql)
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
    With cn.Execute(ssql)
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
   Dim vAmount, vDiscVal, vTotDisc As Double
   Grid.MoveFirst
   ssql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With cn.Execute(ssql)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
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
   With cn.Execute(ssql)
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
   On Error GoTo ErrorHandler
   SchSale.ParaInBillDate = DtpBillDate.DateValue
   SchSale.Show vbModal
   If SchSale.ParaOutBillID <> -1 Then
      TxtBillID.Text = SchSale.ParaOutBillID
      'Dim a
      'a = Split(SchSale.ParaOutBillDate, "/")
      DtpBillDate.DateValue = SchSale.ParaOutBillDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      VStrSQL = "Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")"
      cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetSale
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectOrganization(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganization = False: Exit Function
        TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Organizations where OrganizationID=" & Val(TxtOrganizationID.Text)
    With cn.Execute(VStrSQL)
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
   With cn.Execute("Select distinct o.OrganizationID, ReportName, SaleReportName from SaleHeader h left  join Organizations o on h.OrganizationID = o.OrganizationID where BillID = " & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'")
      VStrSQL = "  select h.BillID, h.BillDate, EntryDate, h.OrganizationID, OrganizationName, Customerid, isnull(Pr.PartyName,AccountName) + ' - ' + H.CustomerID as Customer_Name_ID," & vbCrLf _
               + " pr.address, StoreName, BiltyNo, VehicleNo, h.Description," & vbCrLf _
               + " Isnull(H.BillDiscPer, 0) BillDiscPer, Isnull(H.BillDisc,0) BillDisc, isnull(OtherCharges,0) as OtherCharges,  isnull(h.ServiceCharges,0) as ServiceCharges," & vbCrLf _
               + " TotalAmount,  isnull(TotalExpense,0) as TotalExpense,  CompanyName, GroupName, SubGroupName, BrandName, SeasonName, b.ProductID as Code, " & IIf(isUrdu = True, "p.ProductName1", "p.ProductName") & " as ProductName, dbo.FunSaleBodySerial(b.BillID,b.BillDate, b.ProductId) Serial," & vbCrLf _
               + " dbo.FunSaleBodyOffer(b.BillID,b.BillDate,b.ProductId) ProductOffer, isnull(QtyPack,0)QtyPack, isnull(b.Multiplier,0)Multiplier, isnull(b.GrossQty,0)GrossQty, isnull(b.GrossUnit,0)GrossUnit, Qty," & vbCrLf _
               + " P.RetailPrice, P.PurPrice, Bonus,b.DiscPc, b.DiscPer, DiscVal, Offer, b.SaleTaxPer, SaleTaxval, b.SC," & vbCrLf _
               + " h.Empid, empname, price, Amount, previousAmount, CashReceived, b.RetailPrice, isnull(BatchNo,'') as BatchNo, BillNo," & vbCrLf _
               + " Abbreviation + '/' + cast(b.Multiplier as varchar(10)) as packing, " & vbCrLf _
               + " isnull( pr.Phone1  + ', ','') + isnull( pr.Phone2 + ', ','')  + isnull( pr.mobile + ', ','') +  isnull( pr.mobile2 + ', ','') as Moblie, packingname, pr.city" & vbCrLf _
               + " from SaleBody b inner join products p on b.productid = p.productid" & vbCrLf _
               + " inner join SaleHeader h on h.BillID = b.BillID and h.BillDate = b.BillDate" & vbCrLf _
               + " Left Outer jOin companies cmp on cmp.companyid = p.companyid" & vbCrLf _
               + " Left Outer jOin Groups g on g.Groupid = p.Groupid" & vbCrLf _
               + " Left Outer jOin SubGroups sg on sg.subGroupid = p.subGroupid" & vbCrLf _
               + " Left Outer jOin Brands bd on bd.brandid = p.brandid" & vbCrLf _
               + " Left Outer jOin Seasons se on se.Seasonid = p.Seasonid" & vbCrLf _
               + " LEFT OUTER JOIN packings pak on pak.packingid = b.packingid" & vbCrLf _
               + " left outer join Organizations o on o.OrganizationID = h.OrganizationID" & vbCrLf _
               + " inner join stores s on s.storeid = h.storeid" & vbCrLf _
               + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
               + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
               + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
               + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate='" & DtpBillDate.DateValue & "'" & IIf(ObjRegistry.AllowOrderByCodeinInvoices, "Order By Code", "Order By SerialNo")
         
      If ObjRegistry.LaserPrintofSaleInvoice = True Then
         VStrSQL = "Select UserName, h.billid, h.BillDate, isnull(h.BillTime,0) as BillTime, h.Description, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, isnull(h.STax,0) as STax, isnull(h.cashReceived,0) as CashReceived, " & IIf(isUrdu = True, "p.ProductName1", "p.ProductName") & " as ProductName, unitname, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + Qty as Qty, round(cast(b.price as numeric(9,2))/isnull(multiplier,1),3) as price, b.amount, b.DiscVal, InvoiceNo" & vbCrLf _
               + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else h.CustomerID + ' - ' + AccountName End as Customer, isnull(pr.Address,'') + isnull(' (' + pr.City + ')','') as Address, pr.city, Cash, Credit, BankCard, b.ProductID, PreviousAmount, isnull(OtherCharges,0) as OtherCharges, h.Empid, e.empname, dbo.FunSaleBodySerial(b.BillID,b.BillDate, b.ProductId) Serial, h.TableID, isnull(TableName,'') as TableName, null as DeliveryDate, isnull(h.isPrinted,0) as isPrinted," & IIf(ObjRegistry.AllowUrduProduct = False, " isnull(Remarks,'')", " isnull(RemarksUrdu,'')") & " as Remarks, pr.Phone1, packingname" & vbCrLf _
               + " from saleHeader h inner join salebody b on h.billid = b.billid and h.BillDate = b.BillDate" & vbCrLf _
               + " inner join products p on p.productid = b.productid" & vbCrLf _
               + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
               + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
               + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
               + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
               + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
               + " left outer join Tables t on t.TableID = h.TableID " & vbCrLf _
               + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
               + " left outer join packings pk on pk.packingid = b.packingid" & vbCrLf _
               + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate ='" & DtpBillDate.DateValue & "'" & IIf(ObjRegistry.AllowOrderByCodeinInvoices, "Order By Code", "Order By SerialNo")
      
'      vStrSQL = " select UserName, h.billid, h.BillDate, isnull(h.BillTime,0) as BillTime, h.Description, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.cashReceived,0) as CashReceived, p.ProductName /*case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.billid, h.BillDate) end */ ProductName, unitname, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + Qty as Qty, b.price/isnull(multiplier,1) as price, b.amount, b.DiscVal, InvoiceNo" & vbCrLf _
            + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else h.CustomerID + ' - ' + AccountName End as Customer, isnull(pr.Address,'') as Address, Cash, Credit, BankCard, b.ProductID, PreviousAmount, isnull(OtherCharges,0) as OtherCharges,  h.Empid, e.empname, dbo.FunSaleBodySerial(b.BillID,b.BillDate, b.ProductId) Serial " & vbCrLf _
            + " from saleHeader h inner join salebody b on h.billid = b.billid and h.BillDate = b.BillDate" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
            + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
            + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
            + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
            + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate ='" & DtpBillDate.DateValue & "' Order By SerialNo"
      End If


    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open VStrSQL, cn, adOpenStatic, adLockReadOnly
   
    RptReportViewer.Report.SelectPrinter "abc", "xyz", "ghi"
   
'   Set RptReportViewer.Report = New CrpSaleInvoiceHalf1
   
   If ObjRegistry.LaserPrintofSaleInvoice = True Then
'      Set RptReportViewer.Report = New CrpSaleInvoiceHalf1
      Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\CrpSaleInvoiceHalf1.rpt")
      RptReportViewer.Report.PaperSize = crPaperA4
      RptReportViewer.Report.PaperOrientation = crLandscape
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.X
      RptReportViewer.Report.RightMargin = 225
   Else
      Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\" & IIf(IsNull(!SaleReportName), "CrptSaleInvoice", !SaleReportName) & ".rpt")
      
'      Set RptReportViewer.Report = New CrptSaleInvoice
'      RptReportViewer.Report.PaperOrientation = crPortrait
      
   End If
   
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ReportTitle = "Sale Invoice"
   If ObjRegistry.LaserPrintofSaleInvoice = True Then
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy  'cn.Execute("Select Name from Manufacturer").Fields(0).Value
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo) & IIf(ObjRegistry.CompanyEMail = "", "", ", E.Mail - " & ObjRegistry.CompanyEMail)
      RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.AddSpace = True, ".", "")
      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(ObjRegistry.CashReceived)
      RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(ObjRegistry.Statement)
      RptReportViewer.Report.ParameterFields(8).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(9).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
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
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   
   If ObjRegistry.isShowSeason = False Then
         RptReportViewer.Report.PaperOrientation = crPortrait
      Else
         RptReportViewer.Report.PaperSize = crPaperLegal
         RptReportViewer.Report.PaperOrientation = crLandscape
   End If
   
   If ObjRegistry.PreviewSaleInoice Or ObjRegistry.isShowSeason = True Then
      RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
   End If
   cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
      
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
      With cn.Execute("Select * from Parties where CreditLimit <> 0 and CreditLimit is not null and PartyID = '" & TxtCustomerID.Text & "'")
         If .RecordCount > 0 Then
            If !CreditLimit < (Val(TxtTotalReceivable.Text)) Then
               MsgBox "Credit Limit (" & !CreditLimit & ") is Exceed Balance (" & (Val(TxtTotalReceivable.Text)) & ") for this Customer.", vbExclamation, "Alert"
               Exit Sub
            End If
         End If
      End With
      With cn.Execute("Select * from Employees where CreditLimit <> 0 and CreditLimit is not null and EmpID = '" & TxtCustomerID.Text & "'")
         If .RecordCount > 0 Then
            If !CreditLimit < (Val(TxtTotalReceivable.Text)) Then
               MsgBox "Credit Limit (" & !CreditLimit & ") is Exceed Balance (" & (Val(TxtTotalReceivable.Text)) & ") for this Customer.", vbExclamation, "Alert"
               Exit Sub
            End If
         End If
      End With
   End If
   
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
    VStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where ToUserNo = " & vUser & " and Entrydate <='" & Date & "'"
    With cn.Execute(VStrSQL)
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
    If ObjRegistry.CurrentDateDataEntry = True Then
       If DtpBillDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
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

'  RsBody.Filter = 0
'  If RsBody.RecordCount = 0 Then
'      MsgBox "Please enter at least one product to Purchase", vbExclamation, "Alert"
'      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
'      Exit Sub
'  End If

  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
   cn.BeginTrans
   
   If vIsNewRecord Then
      If cn.Execute("Select * from SaleHeader where BillID = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'").RecordCount > 0 Then
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
      Call ActivityLogSale("Sale Invoice", eEdit, TxtBillID.Text, DtpBillDate.DateValue)
   End If
   
''   Call UserActivities

'' Sale Header
   vStrPara = ""
   vStrPara = Abs(ObjRegistry.AllowContinuousBillNo) & ","
   vStrPara = vStrPara & Abs(ObjRegistry.AllowMonthlyBillNo) & ","
   vStrPara = vStrPara & TxtBillID.Text & ","               'BillID
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"    'BillDate
vStrPara = vStrPara + "'" + TxtCustomerID.Text + "',"       'CustomerID
vStrPara = vStrPara & TxtTotalAmount.Text & ","             'TotalAmount
vStrPara = vStrPara & Val(TxtBillDisc.Text) & ","           'BillDisc
vStrPara = vStrPara & Val(TxtReceivedAmount.Text) & ","     'CashReceived
vStrPara = vStrPara & vUser & ","                           'UserNo
vStrPara = vStrPara & TxtStoreID.Text & ","                 'StoreID
vStrPara = vStrPara & 0 & ","                               'BankCard
With cn.Execute("select dbo.DefaultValue('Cash Counter')")
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
vStrPara = vStrPara & "'" & Null & "',"   'CustomerName
vStrPara = vStrPara & Val(TxtBillDiscPer.Text) & ","        'BillDiscPer
vStrPara = vStrPara & "'" & Null & "',"                     'Commision
vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtCommission.Text)) & ","         'EmpComm
vStrPara = vStrPara & "'" & IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text)) & "',"  'EmpID
vStrPara = vStrPara & 0 & ","                               'isReplace
vStrPara = vStrPara & 0 & ","                               'isPosted
vStrPara = vStrPara & IIf(Trim(TxtMemberID.Text) = "", "''", TxtMemberID.Text) & ","                  'MemberID
vStrPara = vStrPara & "'" & Now & "',"                      'BillTime
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
vStrPara = vStrPara & "'" & Null & "',"                     'ServerEntry
vStrPara = vStrPara & "'" & Null & "',"                     'InvType
vStrPara = vStrPara & "'" & Null & "',"                     'DeliveryDate
vStrPara = vStrPara & "'" & Null & "',"                     'DeliveryTime
vStrPara = vStrPara & "'" & Null & "',"                     'isPrinted
vStrPara = vStrPara & "'" & Null & "',"                     'RemarksUrdu
vStrPara = vStrPara & "'" & Null & "',"                     'StampID
vStrPara = vStrPara & 0 & ","                                'isTransfer
vStrPara = vStrPara & IIf(DtpPromiseDate.DateValue = Empty, "Null", "'" & DtpPromiseDate.DateValue & "'") & "," 'PromiseDate
vStrPara = vStrPara & "'" & IIf(Trim(TxtSyllabusID.Text) = "", Null, Val(TxtSyllabusID.Text)) & "',"  'SyllabusID
vStrPara = vStrPara & "'" & IIf(Trim(vSessionID) = 0, Null, Val(vSessionID)) & "'"  'vSessionID
vStrPara = Replace(vStrPara, "''", "Null")

vStrPara = "DECLARE @returnvalue INT EXEC @returnvalue = saleheaderinsert " & vStrPara & " Select @returnvalue"
   vMasterID = cn.Execute(vStrPara).Fields(0).Value
'   MsgBox vMasterID
   

''' insert Sale Body
vStrDetail = ""
With Grid
 .Redraw = False
 .MoveFirst
   For vCounter = 1 To .Rows
      If Trim(.Columns("Productid").Text) <> "" Then
      
      ''''''''''''''''''''''''''''
 vStrPara = ""
vStrPara = vStrPara & "'" & True & "',"
vStrPara = vStrPara & vMasterID & ","
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
vStrPara = vStrPara & Val(.Columns("Cost").Text) & ","
vStrPara = vStrPara & .Columns("isProduct").Text & ","
vStrPara = vStrPara & IIf(Val(.Columns("PackingID").Text) > 0, .Columns("PackingID").Text, "''") & ","
vStrPara = vStrPara & IIf(.Columns("QtyPack").Text = "", "Null", .Columns("QtyPack").Text) & ","
vStrPara = vStrPara & IIf(.Columns("Pack").Text = "", "Null", .Columns("Pack").Text) & ","
vStrPara = vStrPara & Val(.Columns("Bonus").Text) & ","
vStrPara = vStrPara & Val(.Columns("Offer").Text) & ","
vStrPara = vStrPara & Val(.Columns("SaleTaxPer").Text) & ","
vStrPara = vStrPara & Val(.Columns("SaleTaxVal").Text) & ","
vStrPara = vStrPara & Val(.Columns("TokenVal").Text) & ","
vStrPara = vStrPara & .Columns("RetailPrice").Text & ","
vStrPara = vStrPara & .Columns("IsWSSaleTax").Text & ","
vStrPara = vStrPara & .Columns("IsRetailSaleTax").Text & ","
vStrPara = vStrPara & .Columns("IsWSDiscb4ST").Text & ","
vStrPara = vStrPara & .Columns("SC").Text & "," 'SC
vStrPara = vStrPara & "''" & "," 'EmpComm
vStrPara = vStrPara & IIf(Trim(.Columns("BatchNo").Text) <> "", .Columns("BatchNo").Text, "''") & "," 'Batcho
vStrPara = vStrPara & "''" & "," 'StampID
vStrPara = vStrPara & TxtStoreID.Text & "," 'StoreID
vStrPara = vStrPara & "''" & ","  'EmpID
vStrPara = vStrPara & "null" & "," 'ColourID
vStrPara = vStrPara & "null" & ","  'SizeID
vStrPara = vStrPara & IIf(.Columns("GrossQty").Text = "", "Null", .Columns("GrossQty").Text) & "," 'Gross Qty
vStrPara = vStrPara & IIf(.Columns("GrossUnit").Text = "", "Null", .Columns("GrossUnit").Text) & "," 'Gross Unit
vStrPara = vStrPara & IIf(.Columns("Storeid").Text = "", "Null", .Columns("StoreID").Text) & ""                 'HeaderStoreID

vStrPara = Replace(vStrPara, "''", "Null")
cn.Execute ("Exec SaleBodyInsert " & vStrPara)
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
   For vCounter = 1 To .Rows
      If Trim(.Columns("Productid").Text) <> "" Then
      
      ''''''''''''''''''''''''''''
vStrPara = ""
vStrPara = vStrPara & vMasterID & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
vStrPara = vStrPara & "'" & .Columns("ProductID").Text & "',"
vStrPara = vStrPara & "'" & .Columns("ProductOfferID").Text & "',"
vStrPara = vStrPara & .Columns("Qty").Text & ""

vStrPara = Replace(vStrPara, "''", "Null")
cn.Execute ("Exec SaleBodyOfferInsert " & vStrPara)

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
  For vCounter = 1 To .Rows
      If Trim(.Columns("Productid").Text) <> "" Then
      
      ''''''''''''''''''''''''''''
 vStrPara = ""
vStrPara = vStrPara & vMasterID & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
vStrPara = vStrPara & "'" & .Columns("ProductID").Text & "',"
vStrPara = vStrPara & "'" & .Columns("Serial").Text & "'"

vStrPara = Replace(vStrPara, "''", "Null")
cn.Execute ("Exec SaleBodySerialInsert " & vStrPara)

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
  For vCounter = 1 To .Rows
      If Trim(.Columns("ID").Text) <> "" Then
      
      ''''''''''''''''''''''''''''
 vStrPara = ""
vStrPara = vStrPara & vMasterID & ","
vStrPara = vStrPara & "'" & DtpBillDate.DateValue & "',"
vStrPara = vStrPara & "'" & .Columns("ID").Text & "',"
vStrPara = vStrPara & .Columns("Value").Text & ""

vStrPara = Replace(vStrPara, "''", "Null")
cn.Execute ("Exec SaleExpenseInsert " & vStrPara)

      ''''''''''''''''''''''''''''
      
      End If
      .MoveNext
   Next vCounter
   .RemoveAll
   .Redraw = True
   End With
   
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
      + " from (select OrderID, OrderDate, ProductID, Sum((isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty) as UQty from SaleBody b inner join SaleHeader h on h.BillID = b.BillID and h.BillDate = b.BillDate Group By OrderID, OrderDate, ProductID) b " & vbCrLf _
      + " right outer join SaleOrderBody sob on sob.OrderID = b.orderid and sob.OrderDate = b.orderdate and b.ProductID = sob.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = sob.productid" & vbCrLf _
      + " where sob.OrderID = " & Val(TxtOrderID.Text) & " and sob.OrderDate = '" & DtpOrderDate.DateValue & "' and (isnull(QtyPack,0) * isnull(Multiplier,0)) + isnull(Bonus,0) + Qty - isnull(uqty,0)  <> 0"
   
   With cn.Execute(ssql)
      If .RecordCount = 0 Then
         cn.Execute ("Update SaleOrderHeader set IsSale = 1 Where OrderID = " & Val(TxtOrderID.Text) & " And Orderdate ='" & DtpOrderDate.DateValue & "'")
      End If
   End With

   If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
      Call BtnPrint_Click
   End If
   If vIsNewRecord = True Then Call ActivityLogSale("Sale Invoice", eAdd, TxtBillID.Text, DtpBillDate.DateValue)
   
   '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = " Saved ID:" & TxtBillID.Text & vbCrLf & " Date:" & Format(DtpBillDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtBillDisc.Text) & vbCrLf & " NetAmt" & TxtNetAmount.Text
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               cn.Execute ssql
            End If
         Next
   End If
   
   cn.CommitTrans
'   cn.Close
   FormStatus = NewMode
   Exit Sub
ErrorHandler:

   
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
   
End Sub

Private Sub PopulateDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
'   RsBody.Open "Select * from SaleBody where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      ssql = "select p.ProductName, EmpName, StoreName, code, b.*, dbo.GetExpiryDate(b.ProductID,BatchNo) as ExpiryDate from SaleBody b join products p on p.productid = b.productid left outer join Employees e on e.empid = b.empid left outer join Stores s on s.StoreID = b.StoreID where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
      With cn.Execute(ssql)
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
               Grid.Columns("PackName").Text = cn.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
            End If
            Grid.Columns("Pack").Value = IIf(IsNull(!Multiplier), "", !Multiplier)
            Grid.Columns("GrossQty").Value = IIf(IsNull(!GrossQty), "", !GrossQty)
            Grid.Columns("GrossUnit").Value = IIf(IsNull(!GrossUnit), "", !GrossUnit)
            Grid.Columns("QtyPack").Value = IIf(IsNull(!QtyPack), "", !QtyPack)
            Grid.Columns("QtyLoose").Value = !Qty
            Grid.Columns("Bonus").Value = IIf(IsNull(!Bonus), "", !Bonus)
            Grid.Columns("Price").Value = !Price
            
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
            Grid.Columns("SC").Value = IIf(IsNull(!SC), "", !SC)
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
'   End If
   
   RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   RsBodySerial.Open "Select * from SaleBodySerial where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
   
   Call PopulateDataToGridOffer
   Call PopulateDataToGridExpense
   Call PopulateDataToGridserial
End Sub

Private Sub PopulateSaleOrderToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleBody where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      ssql = " select sob.ProductID, ProductName, QtyPack - isnull(UPack,0) as RQtyPack, Qty - isnull(UQty,0) as RQty, Bonus - isnull(UBonus,0) as RBonus, sob.*" & vbCrLf _
      + " from (select OrderID, OrderDate, ProductID, Sum(Qty) as UQty, Sum(QtyPack) as UPack, Sum(Bonus) as UBonus from SaleBody b inner join SaleHeader h on h.BillID = b.BillID and h.BillDate = b.BillDate Group By OrderID, OrderDate, ProductID) b " & vbCrLf _
      + " right outer join SaleOrderBody sob on sob.OrderID = b.orderid and sob.OrderDate = b.orderdate and b.ProductID = sob.productid" & vbCrLf _
      + " inner join Products p on p.ProductID = sob.productid" & vbCrLf _
      + " where sob.OrderID = " & Val(TxtOrderID.Text) & " and sob.OrderDate = '" & DtpOrderDate.DateValue & "' and (QtyPack - isnull(UPack,0) <> 0 or Qty - isnull(UQty,0) <> 0 or Bonus - isnull(UBonus,0) <> 0) order by serialno"
      
      With cn.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
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
               Grid.Columns("PackName").Text = cn.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
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
            TxtTotalItems.Text = Val(TxtTotalItems.Text) + !RQty + IIf(IsNull(!RBonus), "0", !RBonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!RQtyPack), 0, !RQtyPack))
            
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
   RsBodySerial.Open "Select * from SaleBodySerial where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenDynamic, adLockBatchOptimistic
   
   Call PopulateSaleOrderToGridOffer
   Call PopulateSaleOrderToGridSerial
   Call PopulateSaleOrderToGridExpense
End Sub

Private Sub PopulateDataToGridserial()
'   RsBodySerial.Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
'   If RsBodySerial.RecordCount > 0 Then
       ssql = "select d.* from SaleBodySerial d  where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
      With cn.Execute(ssql)
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
      With cn.Execute(ssql)
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
      FrmHistory.Visible = False
      vZoneID = 0
      vDate = IIf(vSystemDate = True, cn.Execute("Select SystemDate From SystemDate").Fields(0).Value, Date)
      DtpBillDate.DateValue = IIf(vSystemDate = True, vDate, IIf(Format(Now, "hh") >= vHDiff, vDate, DateAdd("d", -1, vDate)))
      DtpOrderDate.DateValue = DtpBillDate.DateValue
      TxtBillID.Text = FunGetMaxID()
'      TxtStampID.Text = StampID()
      Call PopulateDataToGrid
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
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
      If DtpBillDate.Enabled And DtpBillDate.Visible Then DtpBillDate.SetFocus
      GridOffer.Visible = False
      FramExpense.ZOrder 0
      vIsNewRecord = True
      isWholeSale = True
   Case Is = OpenMode
      'TxtBillID.Enabled = False
      DtpBillDate.Enabled = False
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
      TxtMultiplier.Enabled = False
      TxtQtyPack.Enabled = False
      TxtMultiplier.Text = ""
      TxtQtyPack.Text = ""
      TxtPrice.Text = Round(vUnitPrice, 3)
   Else
      TxtMultiplier.Enabled = True
      TxtQtyPack.Enabled = True
      If Trim(TxtCode.Text) <> "" Then
         With cn.Execute("select * from ProductPacking where productid='" & TxtProductID.Text & "' and packingid=" & CmbPackName.ItemData(CmbPackName.ListIndex))
            TxtMultiplier.Text = IIf(.RecordCount = 0, "", !Multiplier)
            If Val(TxtMultiplier.Text) <> 0 Then
               TxtPrice.Text = Round(vUnitPrice * !Multiplier, 3)
            Else
               TxtPrice.Text = Round(vUnitPrice, 3)
            End If
         .Close
         End With
      End If
   End If
End Sub

Private Sub DtpBillDate_Validate(Cancel As Boolean)
   If ActiveControl.Name <> DtpBillDate.Name Then Exit Sub
   TxtBillID.Text = FunGetMaxID()
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
      Case TxtCode.Name
         If FunSelectProduct(ssValidate, False) = True Then If vAutoEnterQtyintoGridSaleInvoice = True Then GetDataFromTexBoxesToGrid Else keybd_event 9, 1, 1, 1: KeyCode = 0
      Case TxtDiscVal.Name, TxtSC.Name, TxtAmount.Name
         Grid.SetFocus
         GetDataFromTexBoxesToGrid
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
      Select Case ActiveControl.Name
      Case TxtProductID.Name, CmbPackName.Name, TxtMultiplier.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtSC.Name, Grid.Name
            If Val(TxtMultiplier.Text) <> 0 Then
               LblLastPurPrice.Caption = cn.Execute("select PurPrice from products where productid = '" & TxtProductID.Text & "'").Fields(0).Value * Val(TxtMultiplier.Text)
            Else
               LblLastPurPrice.Caption = cn.Execute("select PurPrice from products where productid = '" & TxtProductID.Text & "'").Fields(0).Value
            End If
'            LblLastPurPrice.Visible = True
            LblCost.Caption = LblLastPurPrice.Caption
            Call MniCostPrice_Click
      End Select
   ElseIf KeyCode = vbKeyF5 Then
      Select Case ActiveControl.Name
      Case TxtProductID.Name, CmbPackName.Name, TxtMultiplier.Name, TxtQtyLoose.Name, TxtQtyPack.Name, TxtPrice.Name, TxtDiscPC.Name, TxtSC.Name, Grid.Name
            If Val(TxtMultiplier.Text) <> 0 Then
               LblLastPurPrice.Caption = cn.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "','" & TxtProductID.Text & "')").Fields(0).Value * Val(TxtMultiplier.Text)
            Else
               LblLastPurPrice.Caption = cn.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "','" & TxtProductID.Text & "')").Fields(0).Value
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

Private Sub GridOffer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(GridOffer.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
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
   
   If ObjRegistry.BatchNoVisible = False Then LblProductName.Left = TxtProductName.Left

   
   LblEmpID.Visible = ObjRegistry.EmpVisible
   LblEmpName.Visible = ObjRegistry.EmpVisible
   TxtEmployeeID.Visible = ObjRegistry.EmpVisible
   TxtEmployeeName.Visible = ObjRegistry.EmpVisible
   BtnEmployee.Visible = ObjRegistry.EmpVisible
   
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
   
   vAutoEnterQtyintoGridSaleInvoice = ObjRegistry.AutoEnterQtyintoGridSaleInvoice
   
   If ObjRegistry.ShowPromiseDateInSalaPurchase = True Then
      LblPromiseDate.Visible = True
      DtpPromiseDate.Visible = True
      DtpPromiseDate.DateValue = Null
   Else
      LblPromiseDate.Visible = False
      DtpPromiseDate.Visible = False
      DtpPromiseDate.DateValue = Null
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
   With cn.Execute("select * from UserRegistry where UserNo = " & vUser)
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
      FunGetMaxID = cn.Execute("Select isnull(max(BillID),0)+1 from SaleHeader").Fields(0)
   ElseIf ObjRegistry.AllowMonthlyBillNo = True Then
      FunGetMaxID = cn.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where Month(BillDate) = '" & Month(DtpBillDate.DateValue) & "' and  year(BillDate) ='" & Year(DtpBillDate.DateValue) & "'").Fields(0)
   Else
      FunGetMaxID = cn.Execute("Select isnull(max(BillID),0)+1 from SaleHeader where BillDate = '" & DtpBillDate.DateValue & "'").Fields(0)
   End If
  
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function StampID() As Long
   On Error GoTo ErrorHandler
   StampID = cn.Execute("Select isnull(max(SID),0)+1 from Stamp").Fields(0)
   cn.Execute "update Stamp set SID = " & StampID
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpBillDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = cn.Execute("Select isnull(max(BillID),0)+1 from Bin_SaleHeader").Fields(0)
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
   DtpPromiseDate.DateValue = Null
   GridOffer.Visible = False
   FrmExpiry.Visible = False
   Call SubClearSerialFields
   If ObjRegistry.ChangeQtyOnChangedPrice = True Then TxtAmount.Enabled = True
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
    Set frmSaleInvoice = Nothing
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

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(Grid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    If Flag Then Call GetDataBackFromGridToTexBoxes
'    Call PopulateDataToGridserial
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
               cn.Execute ssql
            End If
         Next
   End If
'      RsBody.Filter = "ProductID = '" & TxtProductID.Text & "'" & IIf(ObjRegistry.BatchNoVisible = True, IIf(Trim(TxtBatchNo.Text) = "", "", " and BatchNo = '" & Trim(TxtBatchNo.Text) & "'"), "") & IIf(ObjRegistry.SeperateProductWithPrice = True, " and Price = " & Val(TxtPrice.Text), "") & IIf(ObjRegistry.AllowEmployeProductWise = True, IIf(Trim(TxtEmployeeID.Text) = "", "", " and EmpID = '" & Trim(TxtEmployeeID.Text) & "'"), "") & IIf(ObjRegistry.AllowStoreProductWise = True, " and StoreID = " & Val(TxtStoreID.Text), "")
'      If RsBody.RecordCount > 0 Then RsBody.Delete
      RsProductOffer.Filter = "ProductID='" & GridOffer.Columns("ProductID").Text & "'"
      If RsProductOffer.RecordCount > 0 Then
         RsProductOffer.Delete
         GridOffer.SelBookmarks.RemoveAll
         GridOffer.SelBookmarks.Add GridOffer.Bookmark
         GridOffer.DeleteSelected
         GridOffer.Refresh
         RsProductOffer.Filter = 0
      End If
      RsBodySerial.Filter = "ProductID ='" & TxtProductID.Text & "' And Serial = '" & TxtSerial.Text & "'"
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
      cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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
       cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Removed ProdcutID-" & GridSerial.Columns("ProductID").Text & " Serial-" & GridSerial.Columns("Serial").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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
   
   '************* Last Pur Price Check **************
   If Round(Val(LblLastPurPrice.Caption), 3) * IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text)) > Val(TxtPrice.Text) Then
      If MsgBox("Sale Price is Less than Last (" & Round(Val(LblLastPurPrice.Caption), 3) * IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text)) & "). Do You want to continue?", vbQuestion + vbYesNo, "Alert") = vbNo Then Exit Sub
   End If
   
   FrmHistory.Visible = False
   
   If Val(TxtPrice.Text) <> 0 Then
      If Round(Val(TxtDiscPer.Text), 2) <> Round((Val(TxtDiscPC.Text) * 100) / (Val(TxtPrice.Text) / IIf(Val(TxtMultiplier.Text) = 0, 1, Val(TxtMultiplier.Text))), 2) Then
         MsgBox "Please update the Discount for change Price.", vbExclamation, "Alert"
         If TxtDiscPer.Enabled And TxtDiscPer.Visible Then TxtDiscPer.SetFocus
         Exit Sub
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
            Exit Sub
         End If
      Else
         If (Val(vQtyLoose) - (Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)) + (Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value) + Val(Grid.Columns("QtyLoose").Value))) < 0 Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            Grid.Redraw = True
            Call SubClearDetailArea
            Grid.MoveLast
            If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
            Exit Sub
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
         Exit Sub
      End If
 End If
  
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
            For vrowcounter = 1 To Grid.Rows
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
                  '''''''''''''''''''''''''This QtyOffer Is used for DetailGrid
                  QtyOffer = Val(Grid.Columns("QtyPack").Value) * Val(Grid.Columns("Pack").Value) + Val(Grid.Columns("QtyLoose").Value)
                  GetDataFromTextBoxesToGridOffer
                  TxtOffer.Text = Val(TxtOffer.Text) + Val(Grid.Columns("Offer").Text)
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  TxtQtyLoose.Text = Val(TxtQtyLoose.Text) + Val(Grid.Columns("QtyLoose").Value)
                  TxtQtyPack.Text = Val(TxtQtyPack.Text) + Val(Grid.Columns("QtyPack").Value)
                  Call FindRebate
                  Call SubCalculateBody
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Val(Grid.Columns("QtyLoose").Value) + Val(Grid.Columns("Bonus").Value) + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Grid.Columns("Pack").Value) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
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
                  Grid.Columns("Bonus").Value = Val(TxtBonus.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
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
                  Grid.Columns("SC").Value = Val(TxtSC.Text)
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
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
                  Grid.MoveLast
                   
                  Call SubClearDetailArea
                  TxtCode.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vrowcounter
            Grid.Columns("Serial").Text = Grid.Rows
            Grid.Columns("ProductID").Text = TxtProductID.Text
            Grid.Columns("Code").Text = TxtCode.Text
            Grid.Columns("Price").Value = Val(TxtPrice.Text)
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
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text)))
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         TxtTotalItems.Text = Val(TxtTotalItems.Text) + (Val(TxtQtyLoose.Text) + Val(TxtBonus.Text) + (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text))) - (Grid.Columns("QtyLoose").Value + Grid.Columns("Bonus").Value + (IIf(Val(Grid.Columns("Pack").Value) = 0, 0, Val(Grid.Columns("Pack").Value)) * IIf(Val(Grid.Columns("QtyPack").Value) = 0, 0, Val(Grid.Columns("QtyPack").Value))))
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
      .Columns("SC").Value = Val(TxtSC.Text)
      .Columns("Amount").Value = Val(TxtAmount.Text)
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
'         CmbPackName.SetFocus
      Else
         Grid.MoveLast
         If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End If
   
'      .MoveLast
   End With
   
   QtyOffer = 0
   
   GetDataFromTextBoxesToGridOffer
   If Trim(Grid.Columns("Code").Text) = "" Then
      Call SubClearDetailArea
      TxtCode.SetFocus
   End If
   Grid.Redraw = True
   FrmExpiry.Visible = False
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
   TxtRetailPrice.Text = ""
   TxtTokenVal.Text = ""
   TxtOffer.Text = ""
   TxtSaleTaxPer.Text = ""
   TxtSaleTaxVal.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtDiscVal.Text = ""
   TxtSC.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub GetDataFromTextBoxesToGridOffer()
On Error GoTo ErrorHandler

    With cn.Execute("Select * from ProductOffers where Rebate = 0 and ProductID = '" & TxtProductID.Text & "'")
        
        If .RecordCount > 0 Then
            QtyOffer = QtyOffer + Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)
            QtyOffer = QtyOffer \ !Qty * !QtyOffer
            If QtyOffer > 0 Then
                
                RsProductOffer.Filter = "ProductID='" & TxtProductID.Text & "'"
                If TxtProductID.Enabled Then
                    If RsProductOffer.RecordCount = 0 Then
'                        RsProductOffer.AddNew
                        GridOffer.Columns("ProductID").Text = TxtProductID.Text
                        GridOffer.Columns("ProductOfferID").Text = !ProductOfferID
'                        RsProductOffer!Productid = TxtProductID.Text
'                        RsProductOffer!ProductOfferID = !ProductOfferID
                    Else
                        GridOffer.Redraw = False
                        GridOffer.MoveFirst
                        For vCounter = 1 To GridOffer.Rows
                        If GridOffer.Columns("ProductID").Text = TxtProductID.Text Then
                            GridOffer.Columns("ProductName").Text = cn.Execute("Select ProductName from products where productid = '" & GridOffer.Columns("ProductOfferID").Text & "'").Fields(0)
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
                        GridOffer.Columns("ProductOfferID").Text = !ProductOfferID
'                        RsProductOffer!Productid = TxtProductID.Text
'                        RsProductOffer!ProductOfferID = !ProductOfferID
                End If
                    GridOffer.Columns("ProductName").Text = cn.Execute("Select ProductName from products where productid = '" & GridOffer.Columns("ProductOfferID").Text & "'").Fields(0)
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
            RsProductOffer.Filter = "ProductID='" & TxtProductID.Text & "'"
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
    
'    If GridOffer.Rows <> 0 Then GridOffer.Visible = True
    
   Exit Sub
ErrorHandler:
   GridOffer.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGridOffer()
    If RsProductOffer.State = adStateOpen Then RsProductOffer.Close
    RsProductOffer.Open "Select * from SaleBodyOffer where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenStatic, adLockBatchOptimistic
    If RsProductOffer.RecordCount > 0 Then
    GridOffer.Visible = True
    ssql = "select p.productname, D.* from SaleBodyOffer D Inner join products p on p.productid = D.productOfferid where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'"
      With cn.Execute(ssql)
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
    RsProductOffer.Open "Select * from SaleBodyOffer where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenStatic, adLockBatchOptimistic
'    If RsProductOffer.RecordCount > 0 Then
    
    ssql = "select p.productname, D.* from SaleOrderBodyOffer D Inner join products p on p.productid = D.productOfferid where OrderID =" & Val(TxtOrderID.Text) & " And OrderDate = '" & DtpOrderDate.DateValue & "'"
      With cn.Execute(ssql)
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
    RsExpense.Open "Select * from SaleExpense where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenStatic, adLockBatchOptimistic
'    GridExpense.Visible = True
    ssql = "select EA.AccountNo, Accountname, SE.ExpAmount from ExpenseAccounts EA Left Outer join ChartofAccounts C on C.AccountNo = EA.AccountNo Left Outer Join (Select * from SaleExpense where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "') SE On SE.ExpenseID = EA.AccountNo"
      With cn.Execute(ssql)
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

Private Sub PopulateSaleOrderToGridExpense()
    If RsExpense.State = adStateOpen Then RsExpense.Close
    RsExpense.Open "Select * from SaleExpense where BillID =" & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'", cn, adOpenStatic, adLockBatchOptimistic
'    GridExpense.Visible = True
    ssql = "select EA.AccountNo, Accountname, SE.ExpAmount from ExpenseAccounts EA Left Outer join ChartofAccounts C on C.AccountNo = EA.AccountNo Left Outer Join (Select * from SaleOrderExpense where OrderID =" & Val(TxtOrderID.Text) & " And OrderDate = '" & DtpOrderDate.DateValue & "') SE On SE.ExpenseID = EA.AccountNo"
      With cn.Execute(ssql)
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

     If GridExpense.Rows > 0 Then GridExpense.FirstRow = 0
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
      VStrSQL = "select * from ProductPacking pp inner join packings p on p.packingid = pp.packingid" & vbCrLf _
           + "left outer join ProductBarcodes b on b.productid = pp.productid" & vbCrLf _
           + " where pp.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
      With cn.Execute(VStrSQL)
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
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      TxtSC.Text = .Columns("SC").Value
      TxtAmount.Text = .Columns("Amount").Value
      TxtCost.Text = .Columns("Cost").Value
      ChkIsProduct.Value = Abs(.Columns("isProduct").Value)
      If LblStock.Visible = False Then
         LblStock.Visible = True
         LblStockCaption.Visible = True
         LblCaptionRetailPrice.Visible = True
         LblRetailPrice.Visible = True
      End If
        If ObjRegistry.ShowSavedStock = True Then
            VStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = '" & TxtProductID.Text & "'"
            With cn.Execute(VStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            VStrSQL = "select isnull(dbo.FunStock('" & TxtProductID.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(VStrSQL).Fields(0).Value
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
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtProductID.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
     
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
         LblRetailPrice.Caption = cn.Execute("Select RetailPrice from Products where ProductID = '" & TxtProductID.Text & "'").Fields(0).Value
         LblLastPurPrice.Caption = cn.Execute("select dbo.FunLastPurPrice(1,'" & DtpBillDate.DateValue & "','" & TxtProductID.Text & "')").Fields(0).Value
      End If
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetSale()
   On Error GoTo ErrorHandler
   ssql = "select h.*, EmpName, OrganizationName, isnull(p.partyname,c.accountname) as PartyName, SyllabusName, P.Address, P.City, c.AccountName, BankMachineName, StoreName, EmpName, MemberName FROM SaleHeader h left outer join parties p on h.CustomerID = p.partyid left outer join Organizations o on o.OrganizationID = h.OrganizationID inner join ChartofAccounts c on h.customerid = c.AccountNo left outer join BankMachines b on b.BankMachineid = h.BankMachineid inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID left outer join Members M on M.MemberID = h.memberID left outer join SyllabusHeader syl on syl.syllabusID = h.syllabusID where isReplace=0 and h.BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'" & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With cn.Execute(ssql)
      If Not .BOF Then
          DtpBillDate.DateValue = !BillDate
          DtpPromiseDate.DateValue = !PromiseDate
          TxtOrderID.Text = IIf(IsNull(!OrderID), "", !OrderID)
          DtpOrderDate.DateValue = IIf(IsNull(!OrderDate), "01/01/1990", !OrderDate)
          TxtCustomerID.Text = !CustomerID
          TxtCustomerName.Text = !PartyName
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
          TxtServiceChargesPer.Text = IIf(IsNull(!ServiceChargesPer), "", !ServiceChargesPer)
          TxtServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
          TxtOtherCharges.Text = IIf(IsNull(!OtherCharges), "", !OtherCharges)
          TxtTotalExpense.Text = IIf(IsNull(!TotalExpense), "", !TotalExpense)
          TxtFreight.Text = IIf(IsNull(!Freight), "", !Freight)
'          TxtPaidAmount.Text = IIf(IsNull(!PAIDAMOUNT), "", !PAIDAMOUNT)
          TxtReceivedAmount.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
          TxtDescription.Text = IIf(IsNull(!Description), "", !Description)
          TxtPreviousReceivable.Text = IIf(IsNull(!PreviousAmount), "", !PreviousAmount)
'          TxtPreviousReceivable.Text = cn.Execute("SELECT isnull(dbo.FunCurrentDebit('" & TxtCustomerID.Text & "','" & DtpBillDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0)").Fields(0).Value
'          vStrSQL = " Select isnull(Sum(round(B.TTLValue,0) - isnull(BillDisc,0) + isnull(OtherCharges,0) + Isnull(TotalExpense,0) + isnull(servicecharges,0) + isnull(STax,0)),0) as Amount " & vbCrLf _
                  + " FROM SaleHeader h INNER JOIN (Select BillId, BillDate, Sum(Amount) TTLValue FROM SaleBody Group By BillId, BillDate)b " & vbCrLf _
                  + " ON h.BillId = B.BillId and h.BillDate = B.BillDate " & vbCrLf _
                  + " where CustomerID = '" & (TxtCustomerID.Text) & "' and h.BillDate = '" & DtpBillDate.DateValue & "' and h.BillID >= " & Val(TxtBillID.Text) & IIf(Val(TxtOrganizationID.Text) = 0, "", " and OrganizationID = " & Val(TxtOrganizationID.Text))
'          TxtPreviousReceivable.Text = TxtPreviousReceivable.Text - cn.Execute(vStrSQL).Fields(0).Value
          lblPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
          LblTtlPayable.Caption = IIf(Val(TxtPreviousReceivable.Text) > 0, "Total Receivable", "Total Payable")
          TxtPreviousReceivable.Text = Abs(Val(TxtPreviousReceivable.Text))
          vZoneID = cn.Execute("SELECT isnull(dbo.FunGetZoneID('" & TxtCustomerID.Text & "'),0)").Fields(0).Value
          TxtSyllabusID.Text = IIf(IsNull(!syllabusid), "", !syllabusid)
          TxtSyllabusName.Text = IIf(IsNull(!SyllabusName), "", !SyllabusName)
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
   With cn.Execute(ssql)
      If Not .BOF Then
          DtpOrderDate.DateValue = !OrderDate
          TxtCustomerID.Text = !CustomerID
          TxtCustomerName.Text = !PartyName
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

Private Sub GridSerial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
   'TxtAmount.Text = Round(Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) - Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text), 2)
   TxtAmount.Text = Round((Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text))) - (Val(vUnitPrice) * (Val(TxtQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtQtyLoose.Text)) * Val(TxtDiscPer.Text) / 100), 2)
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
'   RsBodySerial.Filter = "ProductID ='" & Grid.Columns("ProductID").Text & "' And Serial='" & TxtSerial.Text & "'"
        
         GridSerial.Redraw = False
         GridSerial.MoveFirst
            For vrowcounter = 1 To GridSerial.Rows
               If GridSerial.Columns("Serial").Text = TxtSerial.Text Then
                  MsgBox "The Product cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
                  vAlreadySerial = True
                  'SubClearDetailArea
                  GridSerial.MoveLast
                  TxtSerial.SetFocus
                  GridSerial.Redraw = True
                  Exit Sub
               End If
               GridSerial.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         
  If TxtSerial.Enabled Then
'         RsBodySerial.AddNew
         GridSerial.Columns("ProductID").Text = TxtCode.Text
         GridSerial.Columns("Serial").Text = TxtSerial.Text
'         RsBodySerial!Productid = TxtCode.Text
'         RsBodySerial!Serial = TxtSerial.Text
         TxtSerial.Text = ""
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
   If TxtSerial.Visible = True Then TxtSerial.SetFocus
   GridSerial.Redraw = True
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
    With cn.Execute("Select * from ProductOffers where ProductID = '" & TxtProductID.Text & "'")
        .Filter = "Rebate <> 0"
        If .RecordCount > 0 Then
            Rebate = Val(TxtMultiplier.Text) * Val(TxtQtyPack.Text) + Val(TxtQtyLoose.Text)
            Rebate = Rebate \ !Qty
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
    With cn.Execute("Select  * from SaleHeader where BillID =" & TxtBillID.Text & " And BillDate = '" & DtpBillDate.DateValue & "'")
        
        If TxtStoreID.Text <> !StoreID Then
            cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.Rows - 1
        With cn.Execute("Select * from SaleBody Where BillID = " & TxtBillID.Text & " and BillDate ='" & DtpBillDate.DateValue & "' and Productid ='" & Grid.Columns("Productid").Text & "'")
        
             If .EOF = True Then
                ssql = "Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")"
                cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("QtyLoose").Text <> !Qty Or Grid.Columns("Price").Text <> !Price Or Grid.Columns("discper").Text <> !DiscPer Then
                   cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Updated ProdcutID-" & Grid.Columns("Code").Text & " PackingID-" & Grid.Columns("PackName").Text & " Pack" & Grid.Columns("Pack").Text & " QtyPack-" & Grid.Columns("QtyPack").Text & " QtyLoose-" & Grid.Columns("QtyLoose").Text & " Bonus-" & Grid.Columns("Bonus").Text & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
    
   Else
    cn.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub TxtTotalExpense_Change()
   Call SubCalculateFooter
End Sub
Private Sub ActivityLogSale(FormType As String, Mode As EntryMode, Optional Key1 As Long = 0, Optional Key2 As Date = "01-01-1900", Optional Key3 As String = "")
   Dim vSQL As String
   vSQL = "Exec ProdActivityLog '" & FormType & "'," & ObjUserSecurity.UserNo & "," & Mode & "," & Key1 & ",'" & Key2 & "','" & Key3 & "'"
   'vSQL = "INSERT into ActivityLogSaleSale(userno,FormType,EntryDate,Description,isnew,isedit,isdelete) values(" & ObjUserSecurity.UserNo & ",'" & FormType & "',getdate(),'" & Desc & "'," & IIf(Mode = eAdd, 1, 0) & "," & IIf(Mode = eEdit, 1, 0) & "," & IIf(Mode = eDelete, 1, 0) & ")"
   cn.Execute vSQL
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

Private Sub PopulateDataToGridExpiry()
      ssql = "select BatchNo, ExpiryDate " & vbCrLf & _
      " from PurchaseHeader h inner join Purchasebody b on h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate" & vbCrLf & _
      " where BatchNo is not null and b.productid = '" & (TxtProductID.Text) & "' order by b.PurchaseDate Desc"
      
      With cn.Execute(ssql)
         GridExpiry.Redraw = False
         GridExpiry.MoveFirst
         GridExpiry.RemoveAll
         GridExpiry.AllowAddNew = True
         While Not .EOF
            GridExpiry.Columns("BatchNo").Text = !BatchNo
            GridExpiry.Columns("ExpiryDate").Text = !ExpiryDate
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
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        FrmSyllabusSelection.Show vbModal, Me
        If FrmSyllabusSelection.ParaOutID = "" Then FunSelectSyllabus = False: Exit Function
        TxtSyllabusID.Text = FrmSyllabusSelection.ParaOutID
    End If
    '---------------------------

    VStrSQL = " Select * FROM syllabusheader where SyllabusID=" & Val(TxtSyllabusID.Text)
    With cn.Execute(VStrSQL)
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
   With cn.Execute(ssql)
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
       With cn.Execute(ssql)
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
               Grid.Columns("PackName").Text = cn.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
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
'            TxtTotalItems.Text = Val(TxtTotalItems.Text) + !RQty + IIf(IsNull(!RBonus), "0", !RBonus) + (IIf(IsNull(!Multiplier), 0, !Multiplier) * IIf(IsNull(!RQtyPack), 0, !RQtyPack))
            TxtTotalItems.Text = Val(TxtTotalItems.Text) + !Qty
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


