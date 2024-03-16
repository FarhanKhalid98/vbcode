VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmSaleReturnInvoicePOS 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmSaleReturnInvoicePOS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   1395
      TabIndex        =   112
      Top             =   4860
      Visible         =   0   'False
      Width           =   2295
      Begin VB.TextBox TxtSerial 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   15
         TabIndex        =   113
         Top             =   180
         Width           =   2025
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridSerial 
         Height          =   1500
         Left            =   120
         TabIndex        =   114
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
         stylesets(0).Picture=   "FrmSaleReturnInvoicePOS.frx":0ECA
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
   Begin VB.ComboBox CmbColourName 
      Height          =   315
      Left            =   6825
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3210
      Width           =   1200
   End
   Begin VB.ComboBox cmbSizeName 
      Height          =   315
      Left            =   8025
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3240
      Width           =   840
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1748
      MaxLength       =   50
      TabIndex        =   60
      Top             =   9855
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
      Height          =   4560
      Left            =   14865
      TabIndex        =   56
      Top             =   840
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
         Height          =   4110
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   57
         Tag             =   "NC"
         Text            =   "FrmSaleReturnInvoicePOS.frx":0EE6
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
         TabIndex        =   58
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.CheckBox ChkIsProduct 
      Caption         =   "Is Product"
      Height          =   255
      Left            =   7470
      TabIndex        =   50
      Top             =   810
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1365
   End
   Begin SITextBox.Txt TxtReturnID 
      Height          =   315
      Left            =   1830
      TabIndex        =   27
      Top             =   2295
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
   Begin SITextBox.Txt TxtDiscVal 
      Height          =   315
      Left            =   12120
      TabIndex        =   30
      Top             =   3225
      Width           =   990
      _ExtentX        =   1746
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
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1110
      TabIndex        =   6
      Top             =   3225
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
      Left            =   8865
      TabIndex        =   9
      Top             =   3225
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
      Left            =   9645
      TabIndex        =   10
      Top             =   3225
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
      Left            =   13110
      TabIndex        =   28
      Top             =   3225
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   2970
      TabIndex        =   29
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":101F
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8993
      TabIndex        =   25
      Top             =   9420
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":103B
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7673
      TabIndex        =   21
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":1057
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5033
      TabIndex        =   23
      Top             =   9420
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":1073
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10313
      TabIndex        =   26
      Top             =   9420
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":108F
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6353
      TabIndex        =   22
      Top             =   9420
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":10AB
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   3330
      TabIndex        =   38
      Top             =   3225
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
      Height          =   3405
      Left            =   1125
      TabIndex        =   39
      Top             =   3540
      Width           =   13695
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   29
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
      stylesets(0).Picture=   "FrmSaleReturnInvoicePOS.frx":10C7
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
      ExtraHeight     =   1138
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   29
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
      Columns(4).Width=   2090
      Columns(4).Caption=   "Colour"
      Columns(4).Name =   "ColourName"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1508
      Columns(5).Caption=   "SizeName"
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
      Columns(8).Width=   1455
      Columns(8).Caption=   "Disc / Pc"
      Columns(8).Name =   "DiscPC"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1217
      Columns(9).Caption=   "Disc%"
      Columns(9).Name =   "DiscPer"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1746
      Columns(10).Caption=   "Disc. Val"
      Columns(10).Name=   "DiscVal"
      Columns(10).Alignment=   1
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   4
      Columns(10).FieldLen=   256
      Columns(11).Width=   2461
      Columns(11).Caption=   "Amount"
      Columns(11).Name=   "Amount"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   5
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "TotalAmount"
      Columns(12).Name=   "TotalAmount"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "Cost"
      Columns(13).Name=   "Cost"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   4
      Columns(13).FieldLen=   256
      Columns(14).Width=   3200
      Columns(14).Visible=   0   'False
      Columns(14).Caption=   "IsProduct"
      Columns(14).Name=   "IsProduct"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   11
      Columns(14).FieldLen=   256
      Columns(14).Style=   2
      Columns(15).Width=   3200
      Columns(15).Visible=   0   'False
      Columns(15).Caption=   "ColourID"
      Columns(15).Name=   "ColourID"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   3200
      Columns(16).Visible=   0   'False
      Columns(16).Caption=   "SizeID"
      Columns(16).Name=   "SizeID"
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      Columns(17).Width=   3200
      Columns(17).Visible=   0   'False
      Columns(17).Caption=   "EmpComm"
      Columns(17).Name=   "EmpComm"
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      Columns(18).Width=   3200
      Columns(18).Caption=   "SaletaxVal"
      Columns(18).Name=   "SaletaxVal"
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(19).Width=   3200
      Columns(19).Caption=   "SaletaxPer"
      Columns(19).Name=   "SaletaxPer"
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(20).Width=   3200
      Columns(20).Visible=   0   'False
      Columns(20).Caption=   "IsWSDiscb4ST"
      Columns(20).Name=   "IsWSDiscb4ST"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   11
      Columns(20).FieldLen=   256
      Columns(21).Width=   3200
      Columns(21).Visible=   0   'False
      Columns(21).Caption=   "IsWSSaleTax"
      Columns(21).Name=   "IsWSSaleTax"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   11
      Columns(21).FieldLen=   256
      Columns(22).Width=   3200
      Columns(22).Visible=   0   'False
      Columns(22).Caption=   "IsRetailSaleTax"
      Columns(22).Name=   "IsRetailSaleTax"
      Columns(22).DataField=   "Column 22"
      Columns(22).DataType=   11
      Columns(22).FieldLen=   256
      Columns(23).Width=   3200
      Columns(23).Caption=   "SC"
      Columns(23).Name=   "SC"
      Columns(23).DataField=   "Column 23"
      Columns(23).DataType=   8
      Columns(23).FieldLen=   256
      Columns(24).Width=   3200
      Columns(24).Caption=   "GrossQty"
      Columns(24).Name=   "GrossQty"
      Columns(24).DataField=   "Column 24"
      Columns(24).DataType=   8
      Columns(24).FieldLen=   256
      Columns(25).Width=   3200
      Columns(25).Caption=   "GrossUnit"
      Columns(25).Name=   "GrossUnit"
      Columns(25).DataField=   "Column 25"
      Columns(25).DataType=   8
      Columns(25).FieldLen=   256
      Columns(26).Width=   3200
      Columns(26).Caption=   "DiscAmount"
      Columns(26).Name=   "DiscAmount"
      Columns(26).DataField=   "Column 26"
      Columns(26).DataType=   8
      Columns(26).FieldLen=   256
      Columns(27).Width=   3200
      Columns(27).Caption=   "ReSPrice"
      Columns(27).Name=   "ReSPrice"
      Columns(27).DataField=   "Column 27"
      Columns(27).DataType=   8
      Columns(27).FieldLen=   256
      Columns(28).Width=   3200
      Columns(28).Caption=   "ReSAmount"
      Columns(28).Name=   "ReSAmount"
      Columns(28).DataField=   "Column 28"
      Columns(28).DataType=   8
      Columns(28).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   24156
      _ExtentY        =   6006
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
      Left            =   3735
      TabIndex        =   24
      Top             =   9405
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":10E3
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   10605
      TabIndex        =   11
      Top             =   3225
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
      Left            =   8835
      TabIndex        =   42
      Top             =   765
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpReturnDate 
      Height          =   315
      Left            =   2700
      TabIndex        =   0
      Top             =   2295
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
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   11430
      TabIndex        =   12
      Top             =   3225
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
      Left            =   9900
      TabIndex        =   45
      Top             =   765
      Visible         =   0   'False
      Width           =   780
      _ExtentX        =   1376
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
      Left            =   10650
      TabIndex        =   51
      Top             =   765
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
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1605
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
      Left            =   2445
      TabIndex        =   2
      Top             =   1605
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
      Left            =   4110
      TabIndex        =   92
      Top             =   1590
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":10FF
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSale 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   3750
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   1590
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":111B
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   2295
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
      Left            =   4995
      TabIndex        =   62
      Tag             =   "NC"
      Top             =   2295
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
      Left            =   4635
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   2295
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":1137
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   9330
      TabIndex        =   5
      Top             =   2295
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
      Left            =   10440
      TabIndex        =   64
      Top             =   2295
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
      Left            =   10080
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   2295
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":1153
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCommission 
      Height          =   315
      Left            =   11475
      TabIndex        =   70
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
   Begin SITextBox.Txt TxtManualBillNo 
      Height          =   315
      Left            =   11828
      TabIndex        =   20
      Top             =   9480
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
      Left            =   6390
      TabIndex        =   4
      Tag             =   "NC"
      Top             =   2295
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
      Left            =   7455
      TabIndex        =   73
      Tag             =   "NC"
      Top             =   2295
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
      Left            =   7095
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   2295
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":116F
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillDisc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1838
      TabIndex        =   13
      Top             =   7230
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
   Begin SITextBox.Txt TxtBillDiscPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2678
      TabIndex        =   14
      Top             =   7230
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
   Begin SITextBox.Txt TxtServiceCharges 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1838
      TabIndex        =   17
      Top             =   8355
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
      Left            =   2678
      TabIndex        =   18
      Top             =   8355
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
   Begin SITextBox.Txt TxtSTax 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1838
      TabIndex        =   15
      Top             =   7800
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
   Begin SITextBox.Txt TxtSTaxPer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2678
      TabIndex        =   16
      Top             =   7800
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
   Begin SITextBox.Txt TxtTableID 
      Height          =   315
      Left            =   1838
      TabIndex        =   19
      Top             =   8940
      Width           =   525
      _ExtentX        =   926
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
   Begin SITextBox.Txt TxtTableName 
      Height          =   315
      Left            =   2723
      TabIndex        =   91
      Top             =   8940
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
   Begin JeweledBut.JeweledButton BtnTable 
      Height          =   330
      Left            =   2363
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   8940
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":118B
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtRemarks 
      Height          =   315
      Left            =   10118
      TabIndex        =   96
      Top             =   8745
      Width           =   3180
      _ExtentX        =   5609
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
   Begin SITextBox.Txt TxtMemberID 
      Height          =   315
      Left            =   5670
      TabIndex        =   101
      Top             =   1575
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
   Begin SITextBox.Txt TxtMemberName 
      Height          =   315
      Left            =   7470
      TabIndex        =   102
      Top             =   1575
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
      Left            =   7110
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   1575
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
      MICON           =   "FrmSaleReturnInvoicePOS.frx":11A7
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtMemberBarCode 
      Height          =   315
      Left            =   8895
      TabIndex        =   104
      Top             =   1575
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
   Begin SITextBox.Txt TxtSID 
      Height          =   315
      Left            =   1080
      TabIndex        =   108
      Top             =   2295
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
   Begin SITextBox.Txt TxtAvgDisc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1125
      TabIndex        =   110
      Top             =   7215
      Visible         =   0   'False
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
   Begin SITextBox.Txt TxtEmpComm 
      Height          =   315
      Left            =   12330
      TabIndex        =   115
      Top             =   765
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
   Begin SITextBox.Txt TxtSaleTaxPer 
      Height          =   315
      Left            =   135
      TabIndex        =   117
      Top             =   870
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
   Begin SITextBox.Txt TxtSaleTaxValue 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   675
      TabIndex        =   118
      Top             =   855
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
   Begin SITextBox.Txt TxtTotalSaleTaxValue 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1575
      TabIndex        =   119
      Top             =   855
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
   Begin SITextBox.Txt TxtSSID 
      Height          =   315
      Left            =   630
      TabIndex        =   124
      Top             =   1605
      Width           =   1185
      _ExtentX        =   2090
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Barcode"
      Height          =   195
      Left            =   630
      TabIndex        =   123
      Top             =   1410
      Width           =   840
   End
   Begin VB.Label LblSaleTaxPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax%"
      Height          =   195
      Left            =   135
      TabIndex        =   122
      Top             =   675
      Width           =   390
   End
   Begin VB.Label LblSaleTaxValue 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Value"
      Height          =   195
      Left            =   630
      TabIndex        =   121
      Top             =   675
      Width           =   720
   End
   Begin VB.Label LblTotalSaleTaxValue 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Tax Value"
      Height          =   195
      Left            =   1575
      TabIndex        =   120
      Top             =   675
      Width           =   1125
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "EmpComm"
      Height          =   195
      Left            =   12375
      TabIndex        =   116
      Top             =   540
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label55 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Avg (%)"
      Height          =   195
      Left            =   1155
      TabIndex        =   111
      Top             =   6975
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "SID"
      Height          =   195
      Left            =   1080
      TabIndex        =   109
      Top             =   2085
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label LblMemberID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      Height          =   195
      Left            =   5670
      TabIndex        =   107
      Top             =   1350
      Width           =   780
   End
   Begin VB.Label LblMemberName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      Height          =   195
      Left            =   7470
      TabIndex        =   106
      Top             =   1350
      Width           =   1035
   End
   Begin VB.Label LblMemberBarCode 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member BarCode"
      Height          =   195
      Left            =   8850
      TabIndex        =   105
      Top             =   1350
      Width           =   1230
   End
   Begin VB.Label LblColour 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
      Height          =   195
      Left            =   6840
      TabIndex        =   100
      Top             =   3030
      Width           =   450
   End
   Begin VB.Label LblSize 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   8040
      TabIndex        =   99
      Top             =   3030
      Width           =   300
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
      Left            =   1935
      TabIndex        =   98
      Top             =   2955
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblRemarks 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   10118
      TabIndex        =   97
      Top             =   8535
      Width           =   630
   End
   Begin VB.Label LblTableName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name"
      Height          =   195
      Left            =   2678
      TabIndex        =   95
      Top             =   8760
      Width           =   870
   End
   Begin VB.Label LblTableID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Table ID"
      Height          =   195
      Left            =   1838
      TabIndex        =   94
      Top             =   8760
      Width           =   615
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
      Left            =   10973
      TabIndex        =   90
      Top             =   7125
      Width           =   1440
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
      Left            =   4358
      TabIndex        =   89
      Top             =   7095
      Width           =   1365
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
      Left            =   6368
      TabIndex        =   88
      Top             =   7125
      Width           =   1620
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
      Left            =   8558
      TabIndex        =   87
      Top             =   7125
      Width           =   1755
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
      Left            =   10343
      TabIndex        =   86
      Top             =   7425
      Width           =   2730
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
      Left            =   8558
      TabIndex        =   85
      Top             =   7425
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
      ForeColor       =   &H0000FFFF&
      Height          =   915
      Left            =   5783
      TabIndex        =   84
      Top             =   7425
      Width           =   2730
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
      ForeColor       =   &H0000FF00&
      Height          =   915
      Left            =   4358
      TabIndex        =   83
      Top             =   7425
      Width           =   1380
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   2678
      TabIndex        =   82
      Top             =   7575
      Width           =   210
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
      Height          =   195
      Left            =   1838
      TabIndex        =   81
      Top             =   7575
      Width           =   705
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Ch."
      Height          =   195
      Left            =   1838
      TabIndex        =   80
      Top             =   8130
      Width           =   825
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   2678
      TabIndex        =   79
      Top             =   8130
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   2678
      TabIndex        =   78
      Top             =   7005
      Width           =   210
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Disc."
      Height          =   195
      Left            =   1838
      TabIndex        =   77
      Top             =   7005
      Width           =   600
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   6390
      TabIndex        =   76
      Top             =   2055
      Width           =   1095
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   7575
      TabIndex        =   75
      Top             =   2055
      Width           =   1350
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill No"
      Height          =   195
      Left            =   11828
      TabIndex        =   72
      Top             =   9255
      Width           =   1020
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      Height          =   195
      Left            =   11340
      TabIndex        =   71
      Top             =   525
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   4995
      TabIndex        =   69
      Top             =   2055
      Width           =   840
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   3960
      TabIndex        =   68
      Top             =   2055
      Width           =   585
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   9330
      TabIndex        =   67
      Top             =   2055
      Width           =   525
   End
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   10440
      TabIndex        =   66
      Top             =   2055
      Width           =   780
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
      Height          =   225
      Left            =   1763
      TabIndex        =   61
      Top             =   9615
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
      Left            =   13133
      TabIndex        =   59
      Top             =   1905
      Width           =   435
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   " Bill ID"
      Height          =   195
      Left            =   1800
      TabIndex        =   54
      Top             =   1410
      Width           =   450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   2445
      TabIndex        =   53
      Top             =   1410
      Width           =   585
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   10650
      TabIndex        =   52
      Top             =   540
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
      Left            =   12600
      TabIndex        =   49
      Top             =   2490
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
      Left            =   12600
      TabIndex        =   48
      Top             =   2100
      Width           =   720
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Return Invoice"
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
      TabIndex        =   47
      Top             =   270
      Width           =   3420
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   9900
      TabIndex        =   46
      Top             =   570
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LblDiscPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. %"
      Height          =   195
      Left            =   11400
      TabIndex        =   44
      Top             =   3030
      Width           =   525
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Amount"
      Height          =   195
      Left            =   8820
      TabIndex        =   43
      Top             =   555
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label LblProdPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   9630
      TabIndex        =   41
      Top             =   3030
      Width           =   360
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   13268
      Top             =   1440
      Width           =   345
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3330
      TabIndex        =   40
      Top             =   3030
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      Height          =   195
      Left            =   2700
      TabIndex        =   37
      Top             =   2085
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Return ID"
      Height          =   195
      Left            =   1830
      TabIndex        =   36
      Top             =   2085
      Width           =   690
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   1110
      TabIndex        =   35
      Top             =   3030
      Width           =   375
   End
   Begin VB.Label LblDiscPC 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc / PC"
      Height          =   195
      Left            =   10590
      TabIndex        =   34
      Top             =   3030
      Width           =   690
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   8850
      TabIndex        =   33
      Top             =   3030
      Width           =   240
   End
   Begin VB.Label LblAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   13095
      TabIndex        =   32
      Top             =   3030
      Width           =   540
   End
   Begin VB.Label LblDiscVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. Val"
      Height          =   195
      Left            =   12105
      TabIndex        =   31
      Top             =   3030
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
Attribute VB_Name = "FrmSaleReturnInvoicePOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim Application1 As New CRAXDRT.Application
Dim vDate, vServerDate As Date, vHDiff As Integer, vSystemDate As Boolean
Dim vCounter As Integer
Dim RsBody As New ADODB.Recordset
Dim RsBodySerial As New ADODB.Recordset
Dim RsPurchaseSerial As New ADODB.Recordset
Dim RsReturnSerial As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vIsNewRecord As Boolean
Dim vMaxBinID As Integer
Dim Flag As Boolean
Dim DateFlag, vSerialAdd, vAlreadySerial As Boolean
Dim ssql As String
Dim i, vGridRows As Integer
Dim vStrSQL, vRandomID As String, vEmployeeCommision, vShowStock, vUpdateStock As Boolean
Dim vSSID, vQtyLoose As Double, vTotalAmount As Double
Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String, vTotDisc As Double
Dim vPrintHeader As Boolean, vLaserInvoice As Boolean, vX As Integer, vY As Integer, vNoofPrints As Byte
Dim vStrDetail As String
Dim vMobileNo() As String, vMobile As String
Dim vColour As Boolean
Public vReturnDate As String
Dim vIsWSDiscb4ST As Boolean
Dim vIsRetailSaleTax As Boolean
Dim vIsWSSaleTax As Boolean
Dim vMasterID As Long, vMasterID1 As Long
Public objFSO As New Scripting.FileSystemObject
Dim Cnn As New ADODB.Connection
Dim vPOSID As String, vFBRInvoiceNo As String, vUSIN As Long
Dim vProducts, vHeader As String, vConnStr As String
Dim vStrPara, vWhere, vSamePid As String



'----------------------------------
Private Sub SubCalculateBody()
   On Error GoTo ErrorHandler
   TxtActualAmount.Text = Val(TxtQty.Text) * Val(TxtPrice.Text)
   TxtDiscVal.Text = Round(Val(TxtQty.Text) * Val(TxtDiscPC.Text), 0)
   TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
   If vIsWSDiscb4ST = True Then
      TxtSaleTaxValue.Text = Val(TxtAmount.Text) - (Val(TxtAmount.Text) * (100 / (100 + Val(TxtSaleTaxPer.Text))))
   Else
      TxtSaleTaxValue.Text = Val(TxtActualAmount.Text) - (Val(TxtActualAmount.Text) * (100 / (100 + Val(TxtSaleTaxPer.Text))))
   End If
   If ObjRegistry.IsRoundFigure = True Then TxtAmount.Text = SelfRound(TxtAmount.Text)
   TxtTotalDiscount.Caption = vTotDisc
   SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateFooter()
   On Error GoTo ErrorHandler
   TxtTotalDiscount.Caption = Val(TxtBillDisc.Text) + vTotDisc
   TxtNetAmount.Caption = SelfRound(Val(TxtTotalAmount.Caption) - Val(TxtTotalDiscount.Caption) + Val(TxtServiceCharges.Text) + Val(TxtSTax.Text))
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSale(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSale.Show vbModal, Me
        If SchSale.ParaOutBillID = "" Then FunSelectSale = False: Exit Function
        TxtBillID.Text = SchSale.ParaOutBillID
        DtpBillDate.DateValue = SchSale.ParaOutBillDate
    End If
    '---------------------------
    vStrSQL = " Select * FROM SaleHeader where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          DtpBillDate.DateValue = !BillDate
          TxtStoreID.Text = !StoreID
          FunSelectSale = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSale = False
          .Close
          TxtBillID.Text = ""
          DtpBillDate.DateValue = Date
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
    If vColour = True Then
         SchItemCode.ParaInWhere = " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ") "
         SchItemCode.Show vbModal, Me
         TxtCode.Text = SchItemCode.ParaOutItemCode
      Else
         SchProduct.ParaInWhere = " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ") "
         SchProduct.ParainShowStock = vShowStock
         SchProduct.Show vbModal, Me
         TxtCode.Text = SchProduct.ParaOutID
      End If
   End If
    '---------------------------
   
   If TxtCode.Enabled = False Then FunSelectProduct = False: Exit Function
   If Trim(TxtCode.Text) = "" Then FunSelectProduct = False: Exit Function
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    
   
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
 
    
    ''''''''***********   Checking Union   ***********''''''''
   vStrSQL = " SELECT p.productid, Code, ProductName, RetailPrice, DiscPer, DiscPC" & vbCrLf _
         + " from PackageDealInfoHeader un inner join Products p on un.PackageDealID = p.productid" & vbCrLf _
         + " left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
         + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')" & " and isLocked = 0 "
         
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !RetailPrice
         TxtQty.Text = IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
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
         vStrSQL = " select Floor(min(css.QtyLoose/b.QtyLoose)) as QtyLoose " & vbCrLf _
                  + " from PackageDealInfoHeader h inner join PackageDealInfoBody b on h.id = b.id" & vbCrLf _
                  + " inner join Products p on p.productid = b.productid" & vbCrLf _
                  + " left outer join CurrentStockStore css on css.productid = p.productid " & vbCrLf _
                  + " where h.PackageDealID ='" & TxtPID.Text & "' and storeid = " & TxtStoreID.Text
         With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               vQtyLoose = !QtyLoose
               LblStock.Caption = IIf(IsNull(!QtyLoose), 0, !QtyLoose) & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
            Else
               vQtyLoose = 0
               LblStock.Caption = 0
            End If
         End With
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
               
         If ObjRegistry.NegativeSale = False Then
            If LblStock.Caption <= 0 Then
               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
               FunSelectProduct = False
               Exit Function
            End If
         End If
         
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtDiscPC.Text) <> 0 Then
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
         End If
'         ChkIsProduct.Value = 0
         SubCalculateBody
'         Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      End If
   End With
    
   ''''''''***********   Checking Product  ***********''''''''
    vStrSQL = " SELECT p.productid, Qty, code, ProductName, RetailPrice, DiscPer, DiscPC, EmpComm, SaletaxPer, TokenVal, isChangedPrice, IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')" & " and isLocked = 0 "

   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !RetailPrice
         TxtQty.Text = IIf(Len(TxtCode.Text) <= 5 And IsNumeric(TxtCode.Text), 1, IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty))  'IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         'TxtQty.Text = IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty) 'IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         With CN.Execute("select cost from currentstock where productid ='" & TxtPID.Text & "'")
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
'         VStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = '" & TxtPID.Text & "'"
'         With CN.Execute(VStrSQL)
'            If .RecordCount > 0 Then
'               vQtyLoose = .Fields(0).Value
'            Else
'               vQtyLoose = 0
'            End If
'         End With
'         LblStock.Caption = vQtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
         TxtSaleTaxPer.Text = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
         vIsWSDiscb4ST = !IsWSDiscb4ST
         vIsWSSaleTax = !IsWSSaleTax
         vIsRetailSaleTax = !IsRetailSaleTax
         
         vStrSQL = "select isnull(dbo.FunStock('" & TxtPID.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpReturnDate.DateValue + 1 & "',0),0)"
         vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = vQtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
         
         If ObjRegistry.NegativeSale = False Then
            If Val(vQtyLoose) <= 0 Then
               MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
               FunSelectProduct = False
               Exit Function
            End If
         End If
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtDiscPC.Text) <> 0 Then
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
         End If
         TxtEmpComm.Text = IIf(IsNull(!EmpComm), "", !EmpComm)
         ChkIsProduct.Value = 1
         If Val(TxtQty.Text) > 1 Then FindRebate
         SubCalculateBody
'         Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtPID.Text = ""
         TxtCode.Text = ""
         TxtProductName.Text = ""
         TxtPrice.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtAmount.Text = ""
         TxtCost.Text = ""
         TxtSaleTaxPer.Text = ""
         vIsWSDiscb4ST = 0
         vIsWSSaleTax = 0
         vIsRetailSaleTax = 0
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

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
'   cn.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   If MsgBox("Are you sure to Clear the Data?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Exit Sub
    '''''''''''''''''' ActivityLogBin For Clear Action
'      Call DeleteTempActivityLogBin(vRandomID)
      vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From saleReturnbody where SID = " & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Cleared Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
      
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eClearSavedRecord, TxtReturnID.Text, DtpReturnDate.DateValue, vGridRows & " Product/s Cleared ")
      Grid.Redraw = True
  ''''''''''''''''''
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
 '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & Val(TxtReturnID.Text) & ",'" & DtpReturnDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
    ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSaleReturnInvoicePOS", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
      For vCounter = 1 To Grid.rows
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
   
   vMaxBinID = FunGetMaxBinID
   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_SaleReturnHeader Select " & vMaxBinID & ",'" & Date & "', * from SaleReturnHeader Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
'    '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_SaleReturnBody Select " & vMaxBinID & ",'" & Date & "', * from SaleReturnBody Where ReturnID = " & TxtReturnID.Text & " And ReturnDate ='" & DtpReturnDate.DateValue & "'")
   
  '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Call BinData
   Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eDelete, TxtReturnID.Text, DtpReturnDate.DateValue, Grid.rows - 1 & " Product/s Deleted Amount: " & Val(TxtNetAmount.Caption))
   Grid.Redraw = False
   Grid.MoveFirst
'   Call ActivityLog("Sale Return Invoice", eDelete, TxtReturnID.Text, DtpReturnDate.DateValue)

     ''''''''''''''''''' Delete salebodyserial '''''''''''''''''''''''
   RsBodySerial.Filter = ""
   
   While Not RsBodySerial.EOF
      RsPurchaseSerial.Filter = ""
      RsPurchaseSerial.Filter = "Serial = " & RsBodySerial!Serial
      If RsPurchaseSerial.RecordCount > 0 Then RsPurchaseSerial!SerialAdd = 0
      
      CN.Execute "Delete from SaleReturnSerial where ReturnID=" & Val(TxtSID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Productid").Text)
      RsBodySerial.MoveNext
   Wend
      
   
'   RsBodySerial.Filter = ""
'   If RsBodySerial.RecordCount > 0 Then RsBodySerial.UpdateBatch
'
   RsPurchaseSerial.Filter = ""
   If RsPurchaseSerial.RecordCount > 0 Then RsPurchaseSerial.UpdateBatch
   
      
   For vCounter = 1 To Grid.rows
      If Trim(Grid.Columns("Productid").Text) <> "" Then
         CN.Execute "Delete from SaleReturnBody where SID = " & Val(TxtSID.Text) & " And ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Productid").Text) & " and StoreID = " & Val(TxtStoreID.Text)
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   CN.Execute "Delete from SaleReturnHeader where SID = " & Val(TxtSID.Text)
   
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnDelete Then
   vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = ObjRegistry.PrefixPhoneNo + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = ObjUserSecurity.UserName & " " & FrmSaleReturnInvoice.Caption & " Deleted ID:" & TxtReturnID.Text & vbCrLf & " Date:" & Format(DtpReturnDate.DateValue, "dd-MMM-yyyy") & " Time: " & Time & IIf(Val(TxtBillDisc.Text) = 0, "", " Disc:" & TxtTotalDiscount.Caption) & vbCrLf & " NetAmt" & TxtNetAmount.Caption
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
      If TxtCode.Visible = True Then TxtCode.SetFocus
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
      If TxtEmployeeID.Enabled And TxtEmployeeID.Visible Then TxtEmployeeID.SetFocus Else TxtCode.SetFocus
   Else
      TxtMemberID.SetFocus
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

Private Sub BtnOpen_Click()
   SchSaleReturn.ParaInReturnDate = DtpReturnDate.DateValue
   SchSaleReturn.Show vbModal
   If SchSaleReturn.ParaOutReturnID <> -1 Then
      TxtSID.Text = SchSaleReturn.ParaOutSID
      TxtReturnID.Text = SchSaleReturn.ParaOutReturnID
      'Dim a
      'a = Split(SchSaleReturn.ParaOutReturnDate, "/")
      DtpReturnDate.DateValue = SchSaleReturn.ParaOutReturnDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
'      cn.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
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
'      If TxtCustomerID.Enabled Then TxtCustomerID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   'VStrSQL = "select u.username, h.billid, h.BillDate, h.TotalAmount as tbill, isnull(h.discount,0) as discount, isnull(h.cashReceived,0) as cashReceived, p.productname, b.qty, b.price as price, b.amount, b.Disc" _
            + " from saleHeader h inner join salebody b on h.billid = b.billid and h.BillDate = b.BillDate" _
            + " inner join products p on p.productid = b.productid" _
            + " inner join users u on u.UserNo = h.UserNo" _
            + " where h.billid= " & Val(TxtBillID.Text) & " and h.BillDate='" & DtpBillDate.DateValue & "' order by SerialNo"
    
   'vStrSQL = "Select UserName, h.billid, h.BillDate, isnull(h.BillTime,0) as BillTime, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, isnull(h.STax,0) as STax, isnull(h.cashReceived,0) as CashReceived, p.ProductName /*case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.billid, h.BillDate) end */ ProductName, unitname, b.qty, b.price as price, b.amount, b.DiscPC, b.DiscPer, b.DiscVal, isnull(b.SC,0) as SC, InvoiceNo, isnull(Remarks,'') as Remarks" & vbCrLf _
            + " , '' as Desc1, Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName End as Customer, H.empid, isnull(EmpName,'') as EmpName, Cash, Credit, BankCard, b.ProductID, h.MemberID, isnull(cast(h.MemberID as varchar(6)) + '-' + MemberName,'') as MemberName, h.TableID, isnull(TableName,'') as TableName" & vbCrLf _
            + " from saleHeader h inner join salebody b on h.billid = b.billid and h.BillDate = b.BillDate" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
            + " left outer join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
            + " left outer join Members m on m.MemberID = h.MemberID" & vbCrLf _
            + " left outer join Tables t on t.TableID = h.TableID " & vbCrLf _
            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
            + " where h.BillID = " & Val(TxtBillID.Text) & " and h.BillDate ='" & DtpBillDate.DateValue & "' Order By SerialNo"
            
 vStrSQL = " Select UserName, h.Returnid as BillID, h.ReturnDate as BillDate, h.StoreID, isnull(h.ReturnTime,0) as BillTime, h.Description, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, isnull(h.ServiceChargesPer,0) as ServiceChargesPer,isnull(h.STax,0) as STax, isnull(h.cashpaid,0) as CashReceived, case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.billid, h.BillDate) end ProductName, b.qty , b.price as price, b.amount, b.DiscPC, b.DiscPer, b.DiscVal, isnull(b.SC,0) as SC, null as InvoiceNo, isnull(h.Remarks,'') as Remarks" & vbCrLf _
            + " , Case when CustomerID = 621 then isnull(CustomerName,AccountName) Else AccountName End as Customer, H.empid, isnull(EmpName,'') as EmpName, Cash, Credit, cast(0 as bit) as BankCard, b.ProductID, b.Code, p.ItemCode, right('00'+ cast(b.ColourID as varchar(2)),2) as ColourID, right('00'+ cast(b.SizeID as varchar(2)),2) as SizeID, ColourName, SizeName" & vbCrLf _
            + " from SaleReturnHeader h inner join SaleReturnbody b on H.SID = B.SID" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
            + " left outer join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
            + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
            + " Left outer join Colours Col on Col.Colourid = b.ColourID" & vbCrLf _
            + " Left Outer join Sizes Sz on Sz.SizeID = b.SizeID " & vbCrLf _
            + " where h.SID = " & Val(TxtSID.Text) & " Order By SerialNo"
       
'    VStrSQL = "Select username, h.Returnid as BillID, h.ReturnDate as BillDate, isnull(h.ReturnTime,0) as BillTime, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, isnull(h.STax,0) as STax, isnull(h.cashpaid,0) as CashReceived, case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.billid, h.BillDate) end ProductName, unitname, b.qty, b.price as price, b.amount, b.DiscPC, b.DiscPer, b.DiscVal, isnull(b.SC,0) as SC, null as InvoiceNo, isnull(Remarks,'') as Remarks" & vbCrLf _
            + " , Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName End as Customer, H.empid, isnull(EmpName,'') as EmpName, Cash, Credit, cast(0 as bit) as BankCard, b.ProductID" & vbCrLf _
            + " from SaleReturnHeader h inner join SaleReturnbody b on H.SID = B.SID" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
            + " left outer join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
            + " Left outer join Colours Col on Col.Colourid = b.ColourID" & vbCrLf _
            + " Left Outer join Sizes Sz on Sz.SizeID = b.SizeID " & vbCrLf _
            + " where h.ReturnID = " & Val(TxtReturnID.Text) & " and h.ReturnDate ='" & DtpReturnDate.DateValue & "' Order By SerialNo"

   If ObjRegistry.LaserPrintofSaleInvoice = True Then
      vStrSQL = " Select UserName, h.ReturnID, h.ReturnDate, h.StoreID, h.Description, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.PaidAmount,0) as PaidAmount, p.ProductName /*case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.billid, h.BillDate) end */ ProductName, unitname, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + Qty as Qty, b.price/isnull(multiplier,1) as price, b.amount, b.DiscVal, " & vbCrLf _
            + " Case when CustomerID = 621 then isnull(CustomerName, CustomerID + ' - ' + AccountName) Else AccountName End as Customer, Cash, Credit,  b.ProductID, PreviousAmount, isnull(OtherCharges,0) as OtherCharges" & vbCrLf _
            + " from saleReturnHeader h infner join saleReturnbody b on H.SID = B.SID" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
            + " left outer join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
            + " where h.SID = " & Val(TxtSID.Text) & " Order By SerialNo"
   End If

    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
  
'   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port

   If vLaserInvoice = True Then
'      Set RptReportViewer.Report = New CrpSaleReturnInvoiceHalf1
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleReturnInvoiceHalf1.rpt")
      RptReportViewer.Report.PaperSize = crPaperA4
      RptReportViewer.Report.PaperOrientation = crLandscape
      RptReportViewer.Report.TopMargin = vY
      RptReportViewer.Report.LeftMargin = vX
      RptReportViewer.Report.RightMargin = 225
   Else
      If InStr(1, Printer.DeviceName, "CBM1000") > 0 Then
         Set RptReportViewer.Report = New CrpSaleReturnInvoiceCBM
      ElseIf InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
         Set RptReportViewer.Report = New CrpSaleReturnInvoiceAurora
         RptReportViewer.Report.LeftMargin = 225
         RptReportViewer.Report.RightMargin = 0
         RptReportViewer.Report.TopMargin = 255
      Else 'InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
         Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleReturnInvoiceAurora.rpt")
         RptReportViewer.Report.TopMargin = 0
         RptReportViewer.Report.LeftMargin = 0
         RptReportViewer.Report.RightMargin = 0
'         Set RptReportViewer.Report = New CrpSaleReturnInvoiceAurora
         'RptReportViewer.Report.LeftMargin = 0
         'RptReportViewer.Report.RightMargin = 0
      End If
   End If
   
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ReportTitle = "Return Invoice"
'   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port

   If vPrintHeader = True Then
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   Else
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue ""
   End If
   
   If ObjRegistry.LaserPrintofSaleInvoice = True Then
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy
      RptReportViewer.Report.ParameterFields(5).AddCurrentValue CStr(ObjRegistry.Statement)
      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
   Else
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy  'CN.Execute("Select Name from Manufacturer").Fields(0).Value
      RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.AddSpace = True, ".", "")
      RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(ObjRegistry.CashReceived)
      RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(ObjRegistry.Statement)
'      RptReportViewer.Report.ParameterFields(8).AddCurrentValue IIf(ObjRegistry.AddSpace = True, ".", "")
      RptReportViewer.Report.ParameterFields(8).AddCurrentValue ""
   End If

   'RptReportViewer.Report.SelectPrinter "RASDD.DLL", "CBM1000 Partial Cut", "Com1" 'RptReportViewer.Report.SelectPrinter  "RASDD.DLL", "CBM1000 Partial Cut", "Com1"
'   cn.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
'   RptReportViewer.Report.PrintOut False, CInt(IIf(IsNull(ObjRegistry.NoofPrints) Or Val(ObjRegistry.NoofPrints) = 0, 1, ObjRegistry.NoofPrints))
   RptReportViewer.Report.PrintOut False
   'RptReportViewer.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtQty.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnReturnAll_Click()
'   PopulateSaleDataToGrid
   GetSale
End Sub

Private Sub BtnSale_Click()
   If FunSelectSale(ssButton, False) = True Then
      BtnReturnAll.SetFocus
   Else
      TxtBillID.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   Dim p As Object, a As String, B As String, vSQL As String
      
   ''''''''''''' User Discount ''''''''''''''
   If Val(ObjUserSecurity.AllowMaximmDiscPer) <> 0 Then
      If Val(TxtBillDiscPer.Text) > Val(ObjUserSecurity.AllowMaximmDiscPer) Then
         MsgBox "Discount greater than Fixed Discount", vbCritical, "Error"
         Exit Sub
      End If
   End If
  ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSaleReturnInvoicePOS", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
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
   If CN.Execute("Select * From AdminClosing where ToUserNo = " & vUser & " and EntryDate = '" & DtpReturnDate.DateValue & "'").RecordCount > 0 Then
      MsgBox "You are not authorized to Add Record in Closing Dates.", vbCritical, "Alert"
      Exit Sub
   End If
   
   Dim vBm As Variant
   Dim i As Integer
   Grid.Redraw = False
   vBm = Grid.Bookmark
   TxtTotalAmount.Caption = "0"
   vTotalAmount = 0
   Grid.MoveFirst
   For i = 0 To Grid.rows - 1
      TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(Grid.Columns("TotalAmount").CellValue(Grid.GetBookmark(i)))
      vTotalAmount = vTotalAmount + Val(Grid.Columns("Amount").CellValue(Grid.GetBookmark(i)))
   Next i
   Grid.Bookmark = vBm
   Grid.Redraw = True

   Call SubCalculateFooter
   
   '''''''''''''''''''''''Check Employee '''''''''''''''''''''''''''''''''
   If ObjRegistry.EmployeeMandatory = True And TxtEmployeeID.Text = "" Then
      MsgBox "Please Select Employee", vbInformation, Me.Caption
      If TxtEmployeeID.Visible = True Then TxtEmployeeID.SetFocus
      Exit Sub
   End If
   
   
   '''''''''''''''''''''''Check Employee '''''''''''''''''''''''''''''''''
'   If vEmployeeCommision = True Then
'      If Trim(TxtEmployeeID.Text) = "" Then
'         SubDestroyEmployeeCommision
'      Else
'         SubApplyEmployeeCommision
'      End If
'   End If
   
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
'    ssql = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where touserno = " & vUser & " and Entrydate <='" & Date & "'"
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where ToUserNo = " & vUser
    With CN.Execute(vStrSQL)
        If .Fields(0).Value >= DtpReturnDate.DateValue Then
            MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
    '''''''''''''''''''''''Check Current Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.CurrentDateDataEntry = True And ObjUserSecurity.IsAdministrator = False Then
       If DtpReturnDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
   ''''''''''''''''''''''''''''''''''
   
   
   ''''''''''''''''''''''''Check Organization '''''''''''''''''''''''''''''''''
  If ObjRegistry.OrganizationMandatory = True And TxtOrganizationID.Text = "" Then
    MsgBox "Please Select Organization", vbInformation, Me.Caption
    If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
    Exit Sub
  End If
  
   
    '''''''''''''''''''''''Check Entry Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.isEntryDate = True Then
       If ObjRegistry.FromDate > Date Or ObjRegistry.ToDate < Date Then
         MsgBox "Data can not be saved Because Date is not set according to the Software's Entry date", vbInformation, Me.Caption
         Exit Sub
       End If
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
   FrmReturnPrint.DtpReturnDate.DateValue = DtpReturnDate.DateValue
   FrmReturnPrint.TxtOrganizationID.Text = TxtOrganizationID.Text
   FrmReturnPrint.TxtReturnID.Text = TxtReturnID.Text
   FrmReturnPrint.TxtNetAmount.Text = TxtNetAmount.Caption
   FrmReturnPrint.Show vbModal, Me
   If FrmReturnPrint.ParaOutSelection = False Then Exit Sub
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
         With CN.Execute("select isnull(Qty,0) as Qty from salebody where BillID = " & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "' and ProductID = " & Val(RsBody!Productid))
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
   If DtpReturnDate.Date <> IIf(vSystemDate = True, IIf(IsNull(vDate), Date, vDate), IIf(Format(Now, "hh") >= vHDiff, vDate, DateAdd("d", -1, vDate))) And DateFlag = True Then
      If MsgBox("Are you sure to Change Return Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
         DtpReturnDate.DateValue = IIf(vSystemDate = True, IIf(IsNull(vDate), Date, vDate), IIf(Format(Now, "hh") >= vHDiff, vDate, DateAdd("d", -1, vDate)))
      End If
      DateFlag = False
   End If
  
    ''''''''''''''''' Get Commision from commisionDisc if not exists commision in employee
If Trim(TxtEmployeeID.Text) <> "" Then
If CN.Execute("Select commission from employees where EmpID = " & TxtEmployeeID.Text).Fields(0) = 0 Then
  TxtAvgDisc.Text = Round(TxtTotalDiscount.Caption / TxtTotalAmount.Caption * 100, 3)
  ssql = "Select * from commisionDisc Where " & Val(TxtAvgDisc.Text) & " >= DiscPerFrom and " & Val(TxtAvgDisc.Text) & " <= DiscPerTo"
   With CN.Execute(ssql)
'      TxtAvgDisc.Text = Round(TxtTotalDiscount.Caption / TxtNetAmount.Caption * 100, 3)
      If .RecordCount <> 0 Then
         TxtCommission.Text = !Commision
         TxtRemarks.Text = !CommisionName
      End If
   End With
End If
End If
 '''''''''''''''''''''''''''''''''''''''''''
  'Saving record
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
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Amount: " & Val(TxtNetAmount.Caption))
   
'   ssql = "select * from SaleReturnHeader where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
'   Dim Rs As New ADODB.Recordset
'   With Rs
'      .Open ssql, cn, adOpenDynamic, adLockPessimistic
'      If .BOF Then
'         .AddNew
'         !ReturnID = Val(TxtReturnID.Text)
'         !ReturnDate = DtpReturnDate.DateValue
'         !ReturnTime = Now
'         !UserNo = vUser
'      End If
'      !isReplace = 0
'      !isPosted = 0
'      !isTransfer = 0
'      !IsSync = 0
'      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
'      !EmpComm = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text))
'      !BillID = IIf(Val(TxtBillID.Text) = 0, Null, Val(TxtBillID.Text))
'      !BillDate = IIf(IsNull(!BillID), Null, DtpBillDate.DateValue)
'      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
'      !StoreID = TxtStoreID.Text
'      !TableId = IIf(Val(TxtTableID.Text) = 0, Null, TxtTableID.Text)
'      !PreviousAmount = IIf(FrmReturnPrint.lblPayable.Caption = "Previous Receivable", Val(FrmReturnPrint.TxtPreviousReceivable.Text), Val(FrmReturnPrint.TxtPreviousReceivable.Text) * -1)
'      !TotalAmount = SelfRound(vTotalAmount) 'SelfRound(Val(TxtTotalAmount.Caption))
'      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
'      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
'      !ServiceCharges = IIf(TxtServiceCharges.Text = "", Null, Val(TxtServiceCharges.Text))
'      !ServiceChargesPer = IIf(TxtServiceChargesPer.Text = "", Null, Val(TxtServiceChargesPer.Text))
'      !STax = IIf(TxtSTax.Text = "", Null, Val(TxtSTax.Text))
'      !STaxPer = IIf(TxtSTaxPer.Text = "", Null, Val(TxtSTaxPer.Text))
'      If FrmReturnPrint.OptCash.Value = True Then
'         !CashPaid = IIf(FrmReturnPrint.TxtCashPaid.Text = "", Null, Val(FrmReturnPrint.TxtCashPaid.Text))
'         !CustomerID = 621
'         !CustomerName = IIf(Trim(FrmReturnPrint.TxtCashCustomer.Text) = "", Null, FrmReturnPrint.TxtCashCustomer.Text)
'      End If
'      If FrmReturnPrint.OptCredit.Value = True Then
'         !CashPaid = IIf(FrmReturnPrint.TxtCashPaidCredit.Text = "", 0, Val(FrmReturnPrint.TxtCashPaidCredit.Text))
'         !CustomerID = FrmReturnPrint.TxtCustomerID.Text
'         !CustomerName = Null
'      End If
'      !Cash = FrmReturnPrint.OptCash.Value
'      !Credit = FrmReturnPrint.OptCredit.Value
'      !ManualBillNo = IIf(Trim(TxtManualBillNo.Text) = "", "", TxtManualBillNo.Text)
'      !Remarks = IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text)
'      !Tag = IIf(Trim(TxtTag.Text) = "", Null, TxtTag.Text)
'
'      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
'      .Update
'      .Close
'      If vIsNewRecord = True Then TxtSID.Text = cn.Execute("select @@identity").Fields(0).Value
'   End With


'' Sale Header

   
'   vNow = vDate & " " & Format(IIf(vSystemDate = True, Now, cn.Execute("Select getdate()").Fields(0).Value), "hh:mm:ss")
   

Dim vInvoiceNo, vComission, vBankMachineID, vCashReceived, vCashPaid, vCustomerID, vCustomerName As String

If FrmReturnPrint.OptCash.Value = True Then
   vCashPaid = IIf(FrmReturnPrint.TxtCashPaid.Text = "", Null, Val(FrmReturnPrint.TxtCashPaid.Text))
   vCustomerID = 621
   vCustomerName = IIf(Trim(FrmReturnPrint.TxtCashCustomer.Text) = "", Null, FrmReturnPrint.TxtCashCustomer.Text)
End If
If FrmReturnPrint.OptCredit.Value = True Then
   vCashPaid = IIf(FrmReturnPrint.TxtCashPaidCredit.Text = "", 0, Val(FrmReturnPrint.TxtCashPaidCredit.Text))
   vCustomerID = FrmReturnPrint.TxtCustomerID.Text
   vCustomerName = FrmReturnPrint.TxtCustomerName.Text
End If

vStrPara = ""
vStrPara = Abs(ObjRegistry.AllowContinuousBillNo) & ","
vStrPara = vStrPara & Abs(ObjRegistry.AllowMonthlyBillNo) & ","
vStrPara = vStrPara & Abs(ObjRegistry.AllowDailyBillNo) & "," 'AllowDailyBillNo
vStrPara = vStrPara & Val(TxtSID.Text) & "," 'SID
vStrPara = vStrPara & TxtReturnID.Text & ","
vStrPara = vStrPara & "'" & DtpReturnDate.DateValue & "',"
vStrPara = vStrPara & "'" & vCustomerID & "'," 'CustomerID
'vStrPara = vStrPara & SelfRound(vTotalAmount) & "," ' Total Amount
vStrPara = vStrPara & SelfRound(TxtNetAmount.Caption + Val(TxtBillDisc.Text) - Val(TxtServiceCharges.Text) - Val(TxtSTax.Text)) & ","    ' Total Amount
vStrPara = vStrPara & Val(TxtBillDisc.Text) & "," 'BillDisc
vStrPara = vStrPara & vCashPaid & "," ' 'CashPaid
vStrPara = vStrPara & vUser & "," 'UserNo
vStrPara = vStrPara & TxtStoreID.Text & "," 'StoreID
vStrPara = vStrPara & "0,"  'BankCard
vStrPara = vStrPara & IIf(FrmReturnPrint.OptCredit.Value = True, 1, 0) & "," 'Credit
vStrPara = vStrPara & IIf(FrmReturnPrint.OptCash.Value = True, 1, 0) & "," 'Cash
vStrPara = vStrPara & "" & "Null," 'BankMachineID
vStrPara = vStrPara & "'" & vInvoiceNo & "',"  'InvoiceNo
vStrPara = vStrPara & "'" & vCustomerName & "'," 'CustomerName
vStrPara = vStrPara & Val(TxtBillDiscPer.Text) & "," 'BillDiscPer
vStrPara = vStrPara & "Null,"    'Commision
vStrPara = vStrPara & IIf(Trim(TxtEmployeeID.Text) = "", "''", Val(TxtCommission.Text)) & "," 'EmpComm
vStrPara = vStrPara & "'" & IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtEmployeeID.Text)) & "'," 'EmpID
vStrPara = vStrPara & 0 & "," 'isReplace
vStrPara = vStrPara & 0 & "," 'isPosted
vStrPara = vStrPara & IIf(Trim(TxtMemberID.Text) = "", "''", TxtMemberID.Text) & "," 'MemberID
'vStrPara = vStrPara & "'" & vNow & "'," 'BillTime
vStrPara = vStrPara & "'" & vIsNewRecord & "'," 'Tag
vStrPara = vStrPara & "'" & IIf(Trim(TxtManualBillNo.Text) = "", Null, TxtManualBillNo.Text) & "'," 'ManualBillNo
vStrPara = vStrPara & "'" & IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text) & "',"  'Remarks
vStrPara = vStrPara & IIf(Trim(TxtOrganizationID.Text) = "", "''", TxtOrganizationID.Text) & ","  'OrganizationID
vStrPara = vStrPara & "'" & Null & "'," ' BillNo
vStrPara = vStrPara & "'" & Null & "'," ' Bilty No
vStrPara = vStrPara & "'" & Null & "'," 'Description
vStrPara = vStrPara & "''" & "," 'PAIDAMOUNT
vStrPara = vStrPara & "'" & Null & "',"  'EntryDate
vStrPara = vStrPara & IIf(FrmReturnPrint.OptCredit.Value = True, Val(FrmReturnPrint.TxtPreviousReceivable.Text), 0) & "," 'PreviousAmount
vStrPara = vStrPara & 0 & "," 'OtherCharges
vStrPara = vStrPara & "'" & Null & "'," 'SaleManID
vStrPara = vStrPara & 0 & "," 'TotalExpense
'vStrPara = vStrPara & IIf(Val(TxtOrderID.Text) = 0, "''", TxtOrderID.Text) & "," 'OrderID
'vStrPara = vStrPara & "'" & DtpOrderDate.DateValue & "'," 'OrderDate
'vStrPara = vStrPara & 0 & "," 'Freight
'vStrPara = vStrPara & 0 & "," 'IsCustomerFreight
vStrPara = vStrPara & "'" & Null & "'," 'VechicleNo
vStrPara = vStrPara & IIf(TxtServiceCharges.Text = "", "''", Val(TxtServiceCharges.Text)) & "," 'ServiceCharges
vStrPara = vStrPara & IIf(TxtServiceChargesPer.Text = "", "''", Val(TxtServiceChargesPer.Text)) & "," 'ServiceChargesPer
vStrPara = vStrPara & IIf(TxtSTax.Text = "", "''", Val(TxtSTax.Text)) & "," 'STax
vStrPara = vStrPara & IIf(TxtSTaxPer.Text = "", "''", Val(TxtSTaxPer.Text)) & "," 'STaxPer
vStrPara = vStrPara & "'" & IIf(Trim(TxtTableID.Text) = "", Null, TxtTableID.Text) & "'," 'TableID
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
vStrPara = vStrPara & "''" 'Bank Amount in Credit Option
vStrPara = Replace(vStrPara, "''", "Null")

vStrPara = "DECLARE @returnvalue INT EXEC @returnvalue = SaleReturnHeaderInsert " & vStrPara & " Select @returnvalue"
   vMasterID = CN.Execute(vStrPara).Fields(0).Value
   TxtSID.Text = vMasterID
   '/******* FBR Integeration*************/
   If vPOSID <> "" Then
      If ObjRegistry.AllowFBRContinuousNo Then
         vUSIN = CN.Execute("select isnull(max(USIN),0) + 1 as USIN from SaleReturnHeader").Fields(0).Value
      Else
         vUSIN = TxtSID.Text
      End If
      vHeader = "{InvoiceNumber:'',POSID:'" & vPOSID & "',DateTime:'" & Replace(DtpReturnDate.Date, "/", "-") & "',BuyerName:'" & FrmReturnPrint.TxtCustomerName.Text & "',TotalQuantity:" & Val(TxtTotalQty.Caption) & ",TotalSaleValue:" & Val(TxtNetAmount.Caption) - Val(TxtTotalSaleTaxValue.Text) + Val(TxtTotalDiscount.Caption) & ",Totaltaxcharged:" & Val(TxtTotalSaleTaxValue.Text) & ",Discount:" & Val(TxtTotalDiscount.Caption) & ",TotalBillAmount:" & Val(TxtNetAmount.Caption) + Val(TxtTotalDiscount.Caption) & ",PaymentMode:1,InvoiceType:1,USIN:'" & vUSIN & "', items : ["
   End If
   ''''''''''''''''''''''''''''

''' insert Sale Return Body
vStrDetail = ""
vGridRows = 0
vProducts = ""
i = 0
vSamePid = ""
With Grid
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
        vStrPara = vStrPara & "'" & vUpdateStock & "'," 'check stock update or not
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

        vStrPara = vStrPara & .Columns("Cost").Text & ","
        vStrPara = vStrPara & .Columns("isProduct").Text & ","
        vStrPara = vStrPara & "''" & "," ' Pack Name
        vStrPara = vStrPara & "''" & "," ' Qty Pack
        vStrPara = vStrPara & "''" & "," ' Pack
        vStrPara = vStrPara & "''" & "," ' Bonus
        vStrPara = vStrPara & "''" & "," 'Offer
        vStrPara = vStrPara & Val(.Columns("SaleTaxPer").Text) & ","  'SaleTaxPer
        vStrPara = vStrPara & Val(.Columns("SaleTaxVal").Text) & ","  ' SaleTaxVal
'        vStrPara = vStrPara & Val(.Columns("TokenVal").Text) & ","
'        vStrPara = vStrPara & Val(TxtPrice.Text) & "," 'RetailPrice
        vStrPara = vStrPara & Val(.Columns("IsWSSaleTax").Value) & "," 'IsWSSaleTax
        vStrPara = vStrPara & Val(.Columns("IsRetailSaleTax").Value) & ","  'IsRetailSaleTax
        vStrPara = vStrPara & Val(.Columns("IsWSDiscb4ST").Value) & ","  'IsWSDiscb4ST
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
        vStrPara = vStrPara & Val(.Columns("DiscAmount").Value) & "," ' Disc Amount
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




'   vStrDetail = ""
'   With RsBody
'      .Filter = 0
'      .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !SID = Val(TxtSID.Text)
'         !ReturnID = Val(TxtReturnID.Text)
'         !ReturnDate = DtpReturnDate.DateValue
'         !StoreID = Val(TxtStoreID.Text)
'         vStrDetail = vStrDetail & " (P" & !Productid & " Q" & !Qty & " A" & !Amount & ")"
'         .MoveNext
'      Next vCounter
'      .UpdateBatch
'   End With
   
   RsBodySerial.Filter = 0
   If RsBodySerial.RecordCount > 0 Then
     With RsBodySerial
'      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ReturnID = Val(TxtSID.Text)
         !ReturnDate = DtpReturnDate.DateValue

         RsPurchaseSerial.Filter = "Serial = " & RsBodySerial!Serial
         If RsPurchaseSerial.RecordCount > 0 Then RsPurchaseSerial!SerialAdd = 1

         .Update
         .MoveNext
      Next vCounter
      .UpdateBatch
     End With
   End If
   RsPurchaseSerial.Filter = ""
   If RsPurchaseSerial.RecordCount > 0 Then RsPurchaseSerial.UpdateBatch


'   If vIsNewRecord = True Then Call ActivityLog("Sale Return Invoice", eAdd, TxtReturnID.Text, DtpReturnDate.DateValue)
   
   '/******* FBR Integeration*************/

   
   If vPOSID <> "" Then
      vProducts = Left(vProducts, Len(vProducts) - 1)
      MsgBox "1"
      a = vHeader & vProducts & "]};"
      MsgBox "2"
      Set p = JSON.parse(a)
      MsgBox "3"
      B = JSON.toString(p)
      MsgBox "4"
      vFBRInvoiceNo = Webreq(B)
      MsgBox "5"
      vSQL = "update SaleReturnHeader set  FBRInvoiceNo = '" & vFBRInvoiceNo & "', POSID = " & vPOSID & " where sid = " & TxtSID.Text
      MsgBox "6"
      If ObjRegistry.AllowFBRContinuousNo Then
         vSQL = "update SaleReturnHeader set  FBRInvoiceNo = '" & vFBRInvoiceNo & "', POSID = " & vPOSID & ", USIN = " & vUSIN & " where sid = " & TxtSID.Text
      End If
      MsgBox "7"
      CN.Execute vSQL
      '============================================
      '   Start backup entry Master
      '============================================

      MsgBox "8"
      If vConnStr <> "" Then
         vSQL = "update SaleReturnHeader set  FBRInvoiceNo = '" & vFBRInvoiceNo & "', POSID = " & vPOSID & " where sid = " & vMasterID1
         If ObjRegistry.AllowFBRContinuousNo Then
            vSQL = "update SaleReturnHeader set  FBRInvoiceNo = '" & vFBRInvoiceNo & "', POSID = " & vPOSID & ", USIN = " & vUSIN & " where sid = " & vMasterID1
         End If
         Cnn.Execute vSQL
      End If
   End If
   
   ''''''''''''''''''''''''''''

   
   '/******* Mobile SMS *************/
   If ObjRegistry.OwnerMobileNo <> "" And ObjRegistry.AllowSMSOnSave Then
      vMobileNo = Split(ObjRegistry.OwnerMobileNo, " ")
         For i = 0 To UBound(vMobileNo)
            vMobile = "+92" + Right(vMobileNo(i), 10)
            If Len(vMobile) = 13 Then
               ssql = "Saved Return ID:" & TxtReturnID.Text & vbCrLf & " Date:" & Format(DtpReturnDate.DateValue, "dd-MMM-yyyy") & IIf(Val(TxtTotalDiscount.Caption) = 0, "", " Disc:" & TxtTotalDiscount.Caption) & vbCrLf & " NetAmt:" & TxtNetAmount.Caption
               ssql = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) values ('" & vMobile & "','','" & ssql & IIf(ObjRegistry.AllowSMSWithDetail = True, vStrDetail, "") & "','')"
               CN.Execute ssql
            End If
         Next
      
      
   End If
   
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eAdd, TxtReturnID.Text, DtpReturnDate.DateValue, Grid.rows - 1 & " New Product/s Added Amount: " & Val(TxtNetAmount.Caption))
   
   CN.CommitTrans
'   Char.Speak "Thank you for comming"
   'If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
   If FrmReturnPrint.ChkPrint.Value = 1 Then Call BtnPrint_Click
   'End If
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateSaleDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleReturnBody where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "' And StoreID = " & TxtStoreID.Text, CN, adOpenStatic, adLockBatchOptimistic
   ssql = "select p.productname, b.code, b.* from SaleBody b join products p on p.productid = b.productid where BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "' And b.StoreID = " & TxtStoreID.Text & " order by serialno"
   With CN.Execute(ssql)
      If .RecordCount > 0 Then
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalQty.Caption = 0
         TxtTotalSaleTaxValue.Text = ""
         vTotDisc = 0
         vTotalAmount = 0
         TxtTotalAmount.Caption = 0
         While Not .EOF
            RsBody.AddNew
            RsBody!StoreID = !StoreID
            RsBody!Productid = !Productid
            RsBody!Code = !Code
            RsBody!Qty = !Qty
            RsBody!Price = !Price
            RsBody!DiscPC = !DiscPC
            RsBody!DiscPer = !DiscPer
            RsBody!DiscVal = !DiscVal
            RsBody!Cost = !Cost
            RsBody!isProduct = !isProduct
            RsBody!Amount = !Amount
            RsBody.Update
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !Qty
            Grid.Columns("Price").Value = !Price
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("IsProduct").Value = Abs(!isProduct)
            Grid.Columns("TotalAmount").Value = Val(!Qty) * (Val(!Price) + Val(IIf(IsNull(!SC), "", !SC)))
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
            'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(!DiscVal)
            vTotDisc = vTotDisc + Val(!DiscVal)
'            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
            vTotalAmount = vTotalAmount + !Amount
            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
            TxtTotalSaleTaxValue.Text = Val(TxtTotalSaleTaxValue.Text) + IIf(IsNull(!SaleTaxval), 0, !SaleTaxval)
            'TxtLastRate.Caption = !Price
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

Private Sub GetSale()
   On Error GoTo ErrorHandler
   ssql = "select h.*, OrganizationName, c.AccountName, StoreName, EmpName, TableName FROM SaleHeader h left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join Organizations o on o.OrganizationID = h.OrganizationID inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID left outer join Tables t on t.TableID = h.TableID where h.billiD= " & Val(TxtBillID.Text) & " And BillDate = '" & DtpBillDate.DateValue & "'" & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With CN.Execute(ssql)
      If Not .BOF Then
         vSSID = !SID
         TxtBillID.Text = IIf(IsNull(!BillID), "", !BillID)
         DtpBillDate.DateValue = IIf(IsNull(!BillDate), "", !BillDate)
         TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
         TxtEmployeeName.Text = IIf(IsNull(!empname), "", !empname)
         TxtCommission.Text = IIf(IsNull(!EmpComm), "", !EmpComm)
         TxtManualBillNo.Text = IIf(IsNull(!ManualBillNo), "", !ManualBillNo)
         TxtStoreID.Text = !StoreID
         TxtStoreName.Text = !StoreName
         TxtTableID.Text = IIf(IsNull(!TableId), "", !TableId)
         TxtTableName.Text = IIf(IsNull(!TableName), "", !TableName)
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         TxtTotalAmount.Caption = !TotalAmount
         TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
         TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
         TxtServiceChargesPer.Text = IIf(IsNull(!ServiceChargesPer), "", !ServiceChargesPer)
         TxtServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtSTaxPer.Text = IIf(IsNull(!STaxPer), "", !STaxPer)
         TxtSTax.Text = IIf(IsNull(!STax), "", !STax)
         TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
         TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
         TxtNetAmount.Caption = !TotalAmount
         FrmReturnPrint.OptCash.Value = !Cash
         FrmReturnPrint.OptCredit.Value = !Credit
         If FrmReturnPrint.OptCash.Value = True Then
               FrmReturnPrint.TxtCashPaid.Text = IIf(IsNull(!Cash), "", !Cash)
               FrmReturnPrint.TxtCustomerID.Text = ""
               FrmReturnPrint.TxtCustomerName.Text = ""
               FrmReturnPrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
          End If
          If FrmReturnPrint.OptCredit.Value = True Then
               FrmReturnPrint.TxtCashPaidCredit.Text = IIf(IsNull(!Cash), "", !Cash)
               FrmReturnPrint.TxtCustomerID.Text = IIf(IsNull(!CustomerID), "", Val(!CustomerID))
               FrmReturnPrint.TxtCustomerName.Text = !AccountName
               FrmReturnPrint.TxtCashCustomer.Text = ""
          End If
      End If
      .Close
   End With
   Call PopulateSaleDataToGrid
'   FormStatus = OpenMode
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
      vRandomID = Rnd() * 11111 & " " & Format(Now, "dd/mm hh:mm:ss")
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
      Call PopulateDataPurchaseSerial
      Call PopulateDataReturnSerial
      'If RsBody.State = adStateOpen Then RsBody.Close
      TxtBillID.Enabled = True
      DtpBillDate.Enabled = True
      BtnReturnAll.Enabled = True
      BtnSale.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      'TxtCustomerID.Text = "621"
      'TxtCustomerName.Text = "Counter Sale"
      'DtpReturnDate.DateValue = Date
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
'      DtpReturnDate.Enabled = True
      'If DtpReturnDate.Enabled And DtpReturnDate.Visible Then DtpReturnDate.SetFocus
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      BtnSale.Enabled = False
      TxtBillID.Enabled = False
      DtpBillDate.Enabled = False
      BtnReturnAll.Enabled = False
      DtpReturnDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      TxtCode.SetFocus
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
      If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub GetSaleReturn()
   On Error GoTo ErrorHandler
   ssql = "select h.*, OrganizationName, c.AccountName, StoreName, EmpName, TableName FROM SaleReturnHeader h left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join Organizations o on o.OrganizationID = h.OrganizationID inner join stores s on s.storeid = h.storeid left outer join Employees e on e.EmpID = h.EmpID left outer join Tables t on t.TableID = h.TableID where h.SID=" & Val(TxtSID.Text) & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With CN.Execute(ssql)
      If Not .BOF Then
         TxtBillID.Text = IIf(IsNull(!BillID), "", !BillID)
         DtpBillDate.DateValue = IIf(IsNull(!BillDate), "", !BillDate)
         TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
         TxtEmployeeName.Text = IIf(IsNull(!empname), "", !empname)
         TxtCommission.Text = IIf(IsNull(!EmpComm), "", !EmpComm)
         TxtManualBillNo.Text = IIf(IsNull(!ManualBillNo), "", !ManualBillNo)
         TxtStoreID.Text = !StoreID
         TxtStoreName.Text = !StoreName
         TxtTableID.Text = IIf(IsNull(!TableId), "", !TableId)
         TxtTableName.Text = IIf(IsNull(!TableName), "", !TableName)
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         TxtTotalAmount.Caption = !TotalAmount
         TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
         TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
         TxtServiceChargesPer.Text = IIf(IsNull(!ServiceChargesPer), "", !ServiceChargesPer)
         TxtServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtSTaxPer.Text = IIf(IsNull(!STaxPer), "", !STaxPer)
         TxtSTax.Text = IIf(IsNull(!STax), "", !STax)
         TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
         TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
         TxtNetAmount.Caption = !TotalAmount
         FrmReturnPrint.OptCash.Value = !Cash
         FrmReturnPrint.OptCredit.Value = !Credit
         If FrmReturnPrint.OptCash.Value = True Then
               FrmReturnPrint.TxtCashPaid.Text = IIf(IsNull(!CashPaid), "", !CashPaid)
               FrmReturnPrint.TxtCustomerID.Text = ""
               FrmReturnPrint.TxtCustomerName.Text = ""
               FrmReturnPrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
          End If
          If FrmReturnPrint.OptCredit.Value = True Then
               FrmReturnPrint.TxtPreviousReceivable.Text = IIf(IsNull(!PreviousAmount), "", !PreviousAmount)
               FrmReturnPrint.lblPayable.Caption = IIf(Val(FrmReturnPrint.TxtPreviousReceivable.Text) > 0, "Previous Receivable", "Previous Payable")
'               FrmReturnPrint.LblTtlPayable.Caption = IIf(Val(FrmReturnPrint.TxtPreviousReceivable.Text) > 0, "Total Receivable", "Total Payable")
               FrmReturnPrint.TxtPreviousReceivable.Text = Abs(Val(FrmReturnPrint.TxtPreviousReceivable.Text))
               FrmReturnPrint.TxtCashPaidCredit.Text = IIf(IsNull(!CashPaid), "", !CashPaid)
               FrmReturnPrint.TxtCustomerID.Text = IIf(IsNull(!CustomerID), "", !CustomerID)
               FrmReturnPrint.TxtCustomerName.Text = !AccountName
               FrmReturnPrint.TxtCashCustomer.Text = ""
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

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleReturnBody where SID=" & Val(TxtSID.Text), CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      'ssql = "select p.productname, b.code,b.* from SaleReturnBody b join products p on p.productid = b.productid where ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "' order by serialno"
      ssql = " Select p.ProductName, b.code, b.*, ColourName, SizeName," & vbCrLf & _
       " Qty as Qty, Price, SC, b.DiscPC, b.DiscPer, b.DiscVal, Amount, Cost, b.EmpComm " & vbCrLf & _
       " from SaleReturnBody b join products p on p.productid = b.productid " & vbCrLf & _
       " Left outer join Colours Col on Col.Colourid = b.ColourID Left Outer join Sizes Sz on Sz.SizeID = b.SizeID" & vbCrLf & _
       " where SID=" & Val(TxtSID.Text) & " Order by SerialNo asc "
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         'TxtGrossAmount.Text = 0
         TxtTotalQty.Caption = 0
         TxtTotalSaleTaxValue.Text = ""
         'TxtTotalDiscount.Caption = 0
         vTotDisc = 0
         vTotalAmount = 0
         TxtTotalAmount.Caption = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !Qty
            'Grid.Columns("QtyOrigional").Value = !Qty
            Grid.Columns("Price").Value = !Price
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            
            Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), "", !SaleTaxPer)
            Grid.Columns("SaleTaxVal").Value = IIf(IsNull(!SaleTaxval), "", !SaleTaxval)

            Grid.Columns("ColourID").Value = IIf(IsNull(!ColourID), "", !ColourID)
            Grid.Columns("ColourName").Value = IIf(IsNull(!ColourName), "", !ColourName)
            Grid.Columns("SizeID").Value = IIf(IsNull(!SizeID), "", !SizeID)
            Grid.Columns("SizeName").Value = IIf(IsNull(!SizeName), "", !SizeName)

            Grid.Columns("IsProduct").Value = Abs(!isProduct)
            Grid.Columns("TotalAmount").Value = Val(!Qty) * (Val(!Price) + Val(IIf(IsNull(!SC), "", !SC)))
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            Grid.Columns("EmpComm").Value = IIf(IsNull(!EmpComm), "", !EmpComm)
            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
            'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(!DiscVal)
            vTotDisc = vTotDisc + Val(!DiscVal)
            vTotalAmount = vTotalAmount + !Amount
            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
            TxtTotalSaleTaxValue.Text = Val(TxtTotalSaleTaxValue.Text) + IIf(IsNull(!SaleTaxval), 0, !SaleTaxval)
'            TxtLastRate.Caption = !Price
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
   
   RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   RsBodySerial.Open "Select * from SaleReturnSerial where ReturnID=" & Val(TxtSID.Text) & " and ReturnDate = '" & DtpReturnDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   
   PopulateDataToGridserial
'   RsDetail.Filter = 0
'   If RsDetail.State = adStateOpen Then RsDetail.Close
'   RsDetail.Open "Select * from SaleUnionUsed where BillId=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenStatic, adLockBatchOptimistic
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub DtpReturnDate_Validate(Cancel As Boolean)
   TxtReturnID.Text = FunGetMaxID()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      Select Case ActiveControl.Name
         Case TxtCode.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name
         If TxtCode.Enabled Then TxtCode.SetFocus: Call SubClearDetailArea
      End Select
   ElseIf KeyCode = vbKeyReturn And Shift = vbShiftMask Then
      Select Case ActiveControl.Name
      Case TxtCode.Name
         If FunSelectProduct(ssValidate, False) = True Then TxtQty.SetFocus
      Case TxtQty.Name
         TxtDiscPC.SetFocus
      Case TxtDiscPC.Name
         TxtDiscPer.SetFocus
      End Select
      KeyCode = 0
      Shift = 0
   ElseIf KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name
         Grid_DblClick
      Case TxtCode.Name
         FunSelectProduct ssValidate, False
         GetDataFromTexBoxesToGrid
      Case TxtSerial.Name
         If Trim(TxtSerial.Text) = "" Or TxtCode.Enabled = False Then Exit Sub
         TxtCode.Text = Trim(TxtSerial.Text)
         If FunSelectProduct(ssValidate, False) = True Then
               GetDataFromTexBoxesToGrid
               TxtSerial.Text = ""
               TxtSerial.SetFocus
         Else
               keybd_event 9, 1, 1, 1
               KeyCode = 0
         End If
      Case TxtQty.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtPrice.Name, TxtAmount.Name
         GetDataFromTexBoxesToGrid
      Case Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
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
         'If ActiveControl.Name = Grid.Name Then KeyCode = 0: Exit Sub
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
         Case TxtBillID.Name: If FunSelectSale(ssFunctionKey, False) = True Then BtnReturnAll.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtCode.Enabled Then TxtCode.SetFocus
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then If TxtEmployeeID.Visible = True Then If TxtEmployeeID.Enabled Then TxtEmployeeID.SetFocus
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtQty.SetFocus
         Case TxtTableID.Name: If FunSelectTable(ssFunctionKey, False) = True Then TxtTableID.SetFocus
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
   ElseIf ActiveControl.Name = Grid.Name And KeyCode = vbKeyF4 Then
      If Trim(Grid.Columns("ProductID").Text <> "") Then
         If MniCostPrice.Visible = True Then
'            Call MniCostPrice_Click
         End If
      End If
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
         KeyCode = 0
         TxtBillDisc.SetFocus
      End If
   ElseIf KeyCode = vbKeyF5 Then
      If TxtPID.Text <> "" And ObjUserSecurity.ShowPrice = True Then
         Select Case ActiveControl.Name
         Case TxtCode.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name, Grid.Name
            LblCost.Caption = CN.Execute("select dbo.FunPurPrice('" & TxtPID.Text & "')").Fields(0).Value
            Call MniCostPrice_Click
'            LblCost.Visible = True
         End Select
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
   
   If objFSO.FileExists(vTmp & "\Settings.ini") Then
      Open vTmp & "\Settings.ini" For Input As #1
      Line Input #1, vPOSID
      Close #1
   Else
      vPOSID = ""
   End If
'
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
   
   
   
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbRed, lblEffectBorder
   SetWindowText Me.hWnd, "Sale Return Invoice (POS)"
   HelpLocation Me
   
'   FrmLoginSpecificForm .Show vbModal, Me
   
   If ObjUserSecurity.ShowStock = True Or ObjUserSecurity.IsAdministrator Then
      vShowStock = True
   Else
      vShowStock = False
   End If
   
   vColour = ObjRegistry.ShowColourSize
   
   LblColour.Visible = vColour
   CmbColourName.Visible = vColour
   LblSize.Visible = vColour
   cmbSizeName.Visible = vColour
   Grid.Columns("ColourName").Visible = vColour
   Grid.Columns("SizeName").Visible = vColour
   
   If vColour = False Then
      LblQty.Left = LblQty.Left - CmbColourName.Width - cmbSizeName.Width
      TxtQty.Left = TxtQty.Left - CmbColourName.Width - cmbSizeName.Width
      LblProdPrice.Left = LblProdPrice.Left - CmbColourName.Width - cmbSizeName.Width
      TxtPrice.Left = TxtPrice.Left - CmbColourName.Width - cmbSizeName.Width
      LblDiscPC.Left = LblDiscPC.Left - CmbColourName.Width - cmbSizeName.Width
      TxtDiscPC.Left = TxtDiscPC.Left - CmbColourName.Width - cmbSizeName.Width
      LblDiscPer.Left = LblDiscPer.Left - CmbColourName.Width - cmbSizeName.Width
      TxtDiscPer.Left = TxtDiscPer.Left - CmbColourName.Width - cmbSizeName.Width
      LblDiscVal.Left = LblDiscVal.Left - CmbColourName.Width - cmbSizeName.Width
      TxtDiscVal.Left = TxtDiscVal.Left - CmbColourName.Width - cmbSizeName.Width
      LblAmount.Left = LblAmount.Left - CmbColourName.Width - cmbSizeName.Width
      TxtAmount.Left = TxtAmount.Left - CmbColourName.Width - cmbSizeName.Width
      Grid.Width = Grid.Width - CmbColourName.Width - cmbSizeName.Width
   End If
   
   vSystemDate = Abs(ObjRegistry.SystemDate)
   vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
           
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
   LblMemberBarCode.Visible = ObjRegistry.MemberVisible
   TxtMemberBarCode.Visible = ObjRegistry.MemberVisible
   
   LblTableID.Visible = ObjRegistry.TableVisible
   LblTableName.Visible = ObjRegistry.TableVisible
   TxtTableID.Visible = ObjRegistry.TableVisible
   TxtTableName.Visible = ObjRegistry.TableVisible
   BtnTable.Visible = ObjRegistry.TableVisible

   LblSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   TxtSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   LblSaleTaxValue.Visible = ObjRegistry.ShowSaleTax
   TxtSaleTaxValue.Visible = ObjRegistry.ShowSaleTax
   LblTotalSaleTaxValue.Visible = ObjRegistry.ShowSaleTax
   TxtTotalSaleTaxValue.Visible = ObjRegistry.ShowSaleTax
         
   TxtManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   LblManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
         
   TxtRemarks.Visible = ObjRegistry.RemarksVisible
   LblRemarks.Visible = ObjRegistry.RemarksVisible
            
   vEmployeeCommision = ObjRegistry.EmployeeCommision
   
   vLaserInvoice = ObjRegistry.LaserPrintofSaleInvoice
   vPrintHeader = ObjRegistry.PrintHeadersSaleInvoice
         
   vX = IIf(IsNull(ObjRegistry.x), 0, Val(ObjRegistry.x))
   vY = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))
       
   If ObjUserSecurity.IsAdministrator = True Then
      TxtPrice.Enabled = True
   Else
      TxtPrice.Enabled = ObjUserSecurity.IsChangeRetail
   End If
   If ObjUserSecurity.IsAdministrator = False Then
      TxtDiscPC.Enabled = ObjRegistry.DiscAllowed
      TxtDiscPer.Enabled = ObjRegistry.DiscAllowed
      TxtDiscVal.Enabled = ObjRegistry.DiscAllowed
      TxtBillDisc.Enabled = ObjRegistry.DiscAllowed
      TxtBillDiscPer.Enabled = ObjRegistry.DiscAllowed
      TxtSTax.Enabled = ObjRegistry.DiscAllowed
      TxtSTaxPer.Enabled = ObjRegistry.DiscAllowed
      TxtServiceCharges.Enabled = ObjRegistry.DiscAllowed
      TxtServiceChargesPer.Enabled = ObjRegistry.DiscAllowed
      If ObjRegistry.DiscAllowed = False Then
         TxtDiscPC.Tag = "NC"
         TxtDiscPer.Tag = "NC"
         TxtDiscVal.Tag = "NC"
         TxtBillDisc.Tag = "NC"
         TxtBillDiscPer.Tag = "NC"
      End If
   End If
'   If ObjRegistry.ChangePrice = True Then
'      If ObjUserSecurity.IsAdministrator = True Then TxtPrice.Enabled = True
'   End If
   
   With CN.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
         vNoofPrints = IIf(IsNull(!NoofPrints) Or !NoofPrints = 0, 1, !NoofPrints)
      End If
      .Close
   End With
   
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   vServerDate = CN.Execute("Select CONVERT(datetime, CONVERT(varchar, GETDATE(), 110)) ServerDate").Fields(0).Value
   FormStatus = NewMode
'    sSql = InputBox("Enter Password", "Login")
'      If sSql = "" Then
'         BtnSave.Enabled = False
'         Unload Me
'      Else
'         vStrComp = "Select password FROM Users Where (islock = 0 or islock is null)  and password in ('" & EncryptStr(sSql, True) & "')"
'         If cn.Execute(vStrComp).EOF Then
'            BtnSave.Enabled = False
'            Unload Me
'         End If
'      End If
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
   'TxtLastRate.Caption = 0
   CmbColourName.Clear
   cmbSizeName.Clear
   TxtTotalQty.Caption = 0
   TxtTotalDiscount.Caption = 0
   TxtTotalAmount.Caption = 0
   TxtNetAmount.Caption = 0
   TxtTotalSaleTaxValue.Text = ""
   vTotDisc = 0
   vTotalAmount = 0
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ProductID").Text = " "
   Grid.Update
   Unload FrmReturnPrint
   Call SubClearSerialFields
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
    'CN.Execute ("exec spcurrentstock")
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set RsBody = Nothing
    Set RsReport = Nothing
    Set FrmSaleReturnInvoicePOS = Nothing
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
               ssql = "Select Productid From saleReturnbody where SID=" & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
               With CN.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Closed Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eCloseSavedRecord, TxtReturnID.Text, DtpReturnDate.DateValue, vGridRows & " Product/s Closed")
         Grid.Redraw = True
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
   TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) - Grid.Columns("TotalAmount").Value
   TxtTotalSaleTaxValue.Text = Val(TxtTotalSaleTaxValue.Text) - Val(Grid.Columns("SaleTaxVal").Value)
   vTotalAmount = vTotalAmount - Grid.Columns("Amount").Value
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
   TxtCode.Enabled = False
   BtnProduct.Enabled = False
   'TxtCode.BackColor = TxtProductName.BackColor
   'TxtCode.TabStop = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   On Error GoTo ErrorHandler
   Flag = False
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      TxtCode.Text = ""
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      If TxtCode.Enabled = True And TxtCode.Visible Then TxtCode.SetFocus
   Else
      TxtCode.Enabled = False
      BtnProduct.Enabled = False
      If TxtQty.Enabled = True And TxtQty.Visible Then TxtQty.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(Grid.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then
      Call GetDataBackFromGridToTexBoxes
      Call PopulateDataToGridserial
   End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
   
   If ObjRegistry.NegativeSale = False Then
      ssql = "Select QtyPack, Multiplier, Qty as Qtyloose From SaleReturnBody Where SID = " & Val(TxtSID.Text) & " and Productid = " & Val(TxtCode.Text)
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
    
   ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtSID.Text) & " and ReturnDate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
   With CN.Execute(ssql)
      If .EOF Then
         Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      Else
         Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eRemoveRow, TxtReturnID.Text, DtpReturnDate.DateValue, "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
         Call ActivityLogBin(vRandomID, eFrmSaleReturnInvoicePOS, eAddTempRecord, TxtReturnID.Text, DtpReturnDate.DateValue, "Pending Remove Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      End If
   End With
   RsBody.Filter = "Code='" & TxtCode.Text & "' and StoreID = " & Val(TxtStoreID.Text)
   If RsBody.RecordCount > 0 Then RsBody.Delete
'   cn.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   Grid.Refresh
   RsBody.Filter = 0
   Grid.MoveLast
   
   RsBodySerial.Filter = ""
   RsBodySerial.Filter = "ProductID = " & Val(TxtCode.Text)
   
   While Not RsBodySerial.EOF
      
      RsPurchaseSerial.Filter = "Serial = " & RsBodySerial!Serial
      If RsPurchaseSerial.RecordCount > 0 Then
         RsPurchaseSerial!SerialAdd = 0
         RsPurchaseSerial.Update
       End If
            
      RsBodySerial.Delete
      RsBodySerial.MoveNext
   Wend
   GetDataBackFromGridToTexBoxes
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   Dim vrowcounter As Integer
   If Trim(TxtCode.Text) = "" Then
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
      TxtCode.SetFocus
      Exit Sub
   End If
   If Val(TxtQty.Text) = 0 Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      TxtQty.SetFocus
      Exit Sub
   End If
   If Val(TxtPrice.Text) <> 0 Then
      If Round(Val(TxtDiscPer.Text), 2) <> Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2) Then
         MsgBox "Please update the Discount for change Price.", vbExclamation, "Alert"
         If TxtDiscPer.Enabled And TxtDiscPer.Visible Then TxtDiscPer.SetFocus
         Exit Sub
      End If
   End If
   If CmbColourName.Text = "" And cmbSizeName.Text = "" And vColour = True Then
     MsgBox "Please Select Colour and Size", vbInformation + vbOKOnly, "Error"
     Exit Sub
   End If
   
   If ObjRegistry.NegativeSale = False And TxtCode.Enabled = False And vIsNewRecord = False Then
     ssql = "Select QtyPack, Multiplier, Qty as Qtyloose From SaleReturnBody Where SID = " & Val(TxtSID.Text) & " and Productid = " & Val(TxtCode.Text)
     With CN.Execute(ssql)
         If IIf(IsNull(!QtyLoose), 0, !QtyLoose) - Val(TxtQty.Text) > Val(vQtyLoose) Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            Exit Sub
         End If
         .Close
      End With
    End If
    
   '''''''''   check Serial
   RsBodySerial.Filter = "ProductID =" & Val(TxtCode.Text)
   If (TxtCode.Enabled = False And RsBodySerial.RecordCount <> 0) And RsBodySerial.RecordCount <> TxtQty.Text Then
      MsgBox "Qty Should be equal to Serial", vbInformation + vbOKOnly, "Error"
      Call SubClearDetailArea
      If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
      Exit Sub
   End If
   RsBodySerial.Filter = ""
''''''''
On Error GoTo ErrorHandler
   RsBody.Filter = "ProductID = " & Val(TxtPID.Text)
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ProductID").Text = TxtPID.Text
         Grid.Columns("Code").Text = TxtCode.Text
         RsBody!Productid = TxtPID.Text
         RsBody!Code = TxtCode.Text
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.rows
               If Grid.Columns("Productid").Text = TxtPID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                   ssql = "Select Productid From saleReturnbody where sid = " & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  TxtQty.Text = Val(TxtQty.Text) + Grid.Columns("Qty").Value
                  TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text) - Val(Grid.Columns("Qty").Text)
                  vTotDisc = vTotDisc + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
                  'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
                  TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text) - Val(Grid.Columns("TotalAmount").Text)
                  'TxtLastRate.Caption = Val(TxtPrice.Text) - Val(TxtDiscPC.Text)
                  TxtTotalSaleTaxValue.Text = Val(TxtTotalSaleTaxValue.Text) + Val(TxtSaleTaxValue.Text) - Val(Grid.Columns("SaleTaxVal").Text)

                  vTotalAmount = vTotalAmount + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("Qty").Value = Val(TxtQty.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("DiscPC").Value = Val(TxtDiscPC.Text)
                  Grid.Columns("DiscPer").Value = Val(TxtDiscPer.Text)
                  Grid.Columns("DiscVal").Value = Val(TxtDiscVal.Text)
                  
                  Grid.Columns("SaleTaxPer").Value = Val(TxtSaleTaxPer.Text)
                  Grid.Columns("SaleTaxVal").Value = Val(TxtSaleTaxValue.Text)
                  
                  Grid.Columns("IsWSDiscb4ST").Value = vIsWSDiscb4ST
                  Grid.Columns("IsWSSaleTax").Value = vIsWSSaleTax
                  Grid.Columns("IsRetailSaleTax").Value = vIsRetailSaleTax

                  Grid.Columns("Cost").Value = Val(TxtCost.Text)
                  Grid.Columns("EmpComm").Value = IIf(Val(TxtEmpComm.Text) = 0, 0, Val(TxtEmpComm.Text))
                  Grid.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
                  RsBody!StoreID = Val(TxtStoreID.Text)
                  RsBody!HeaderStoreID = Val(TxtStoreID.Text)
                  RsBody!Qty = Val(TxtQty.Text)
                  RsBody!Price = Val(TxtPrice.Text)
                  RsBody!DiscPC = Val(TxtDiscPC.Text)
                  RsBody!DiscPer = Val(TxtDiscPer.Text)
                  RsBody!DiscVal = Val(TxtDiscVal.Text)
                  RsBody!SaleTaxPer = Val(TxtSaleTaxPer.Text)
                  RsBody!SaleTaxval = Val(TxtSaleTaxValue.Text)
                  RsBody!IsWSDiscb4ST = Val(vIsWSDiscb4ST)
                  RsBody!IsWSSaleTax = Val(vIsWSSaleTax)
                  RsBody!IsRetailSaleTax = Val(vIsRetailSaleTax)
                  RsBody!Cost = Val(TxtCost.Text)
                  RsBody!isProduct = Abs(ChkIsProduct.Value)
                  RsBody!EmpComm = Val(TxtEmpComm.Text)
                  RsBody!Amount = Val(TxtAmount.Text)
                  ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmSaleReturnInvoicePOS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Pending Update Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
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
      If TxtCode.Enabled = True Then
         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text)
         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text)
         TxtTotalSaleTaxValue.Text = Val(TxtTotalSaleTaxValue.Text) + Val(TxtSaleTaxValue.Text)
         
         'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(TxtDiscVal.Text)
         vTotalAmount = vTotalAmount + Val(TxtAmount.Text)
         vTotDisc = vTotDisc + Val(TxtDiscVal.Text)
         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text)
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eAddNewRowByEdit, TxtReturnID.Text, DtpReturnDate.DateValue, "Add New Code-" & TxtCode.Text & " Qty-" & TxtQty.Text & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
         Call ActivityLogBin(vRandomID, eFrmSaleReturnInvoicePOS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Pending Add New Code-" & TxtCode.Text & " Qty-" & TxtQty.Text & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
      Else
         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text) - Val(.Columns("Qty").Text)
         TxtTotalSaleTaxValue.Text = Val(TxtTotalSaleTaxValue.Text) + Val(TxtSaleTaxValue.Text) - Val(Grid.Columns("SaleTaxVal").Text)
         vTotDisc = vTotDisc + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
         'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
         vTotalAmount = vTotalAmount + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text) - Val(Grid.Columns("TotalAmount").Text)
         ssql = "Select Productid From saleReturnbody where sid=" & Val(TxtSID.Text) & " and Returndate ='" & DtpReturnDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
         With CN.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Updated Code-" & TxtCode.Text & " Qty-" & TxtQty.Text & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            Else
               Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eEdit, TxtReturnID.Text, DtpReturnDate.Date, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmSaleReturnInvoicePOS, eEdit, TxtReturnID.Text, DtpReturnDate.Date, "Updated Code-" & TxtCode.Text & " Qty-" & TxtQty.Text & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmSaleReturnInvoicePOS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtReturnID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpReturnDate.Date), "Pending Update Code-" & TxtCode.Text & " Qty-" & TxtQty.Text & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
      End If
      
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("Qty").Value = Val(TxtQty.Text)
      .Columns("Price").Value = Val(TxtPrice.Text)
      .Columns("DiscPC").Value = Val(TxtDiscPC.Text)
      .Columns("DiscPer").Value = Val(TxtDiscPer.Text)
      .Columns("DiscVal").Value = Val(TxtDiscVal.Text)
      .Columns("SaleTaxPer").Value = Val(TxtSaleTaxPer.Text)
      .Columns("SaleTaxVal").Value = Val(TxtSaleTaxValue.Text)
      .Columns("IsWSDiscb4ST").Value = vIsWSDiscb4ST
      .Columns("IsWSSaleTax").Value = vIsWSSaleTax
      .Columns("IsRetailSaleTax").Value = vIsRetailSaleTax
      If Trim(TxtCost.Text) <> "" Then
         .Columns("Cost").Value = Val(TxtCost.Text)
      End If
      .Columns("EmpComm").Value = IIf(Val(TxtEmpComm.Text) = 0, 0, Val(TxtEmpComm.Text))
      .Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
            
      Grid.Columns("ColourName").Text = CmbColourName.Text
      If CmbColourName.Text <> "" Then Grid.Columns("ColourID").Value = CmbColourName.ItemData(CmbColourName.ListIndex)
      Grid.Columns("SizeName").Text = cmbSizeName.Text
      If cmbSizeName.Text <> "" Then Grid.Columns("SizeID").Value = cmbSizeName.ItemData(cmbSizeName.ListIndex)

      'IIf(CmbPackName.ListIndex = 0, Null, CmbPackName.ItemData(CmbPackName.ListIndex))
      'TxtLastRate.Caption = Val(TxtPrice.Text) - Val(TxtDiscPC.Text)
      If CmbColourName.Text <> "" Then
         RsBody!ColourID = CmbColourName.ItemData(CmbColourName.ListIndex)
      Else
         RsBody!ColourID = Null
      End If
      If cmbSizeName.Text <> "" Then
         RsBody!SizeID = cmbSizeName.ItemData(cmbSizeName.ListIndex)
      Else
         RsBody!SizeID = Null
      End If
      RsBody!StoreID = Val(TxtStoreID.Text)
      RsBody!HeaderStoreID = Val(TxtStoreID.Text)
      RsBody!Qty = Val(TxtQty.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!DiscPC = Val(TxtDiscPC.Text)
      RsBody!DiscPer = Val(TxtDiscPer.Text)
      RsBody!DiscVal = Val(TxtDiscVal.Text)
      RsBody!SaleTaxPer = Val(TxtSaleTaxPer.Text)
      RsBody!SaleTaxval = Val(TxtSaleTaxValue.Text)
      
'      RsBody!IsWSDiscb4ST = Val(vIsWSDiscb4ST)
'      RsBody!IsWSSaleTax = Val(vIsWSSaleTax)
'      RsBody!IsRetailSaleTax = Val(vIsRetailSaleTax)
      
      If Trim(TxtCost.Text) <> "" Then
         RsBody!Cost = Val(TxtCost.Text)
      End If
      If IsNull(RsBody!Cost) Then RsBody!Cost = 0
      RsBody!EmpComm = Val(TxtEmpComm.Text)
      RsBody!Amount = Val(TxtAmount.Text)
      RsBody!isProduct = Abs(ChkIsProduct.Value)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearDetailArea
   TxtCode.SetFocus
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
   TxtQty.Text = ""
   TxtPrice.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtDiscVal.Text = ""
   TxtSaleTaxPer.Text = ""
   TxtSaleTaxValue.Text = ""
   TxtAmount.Text = ""
   TxtActualAmount.Text = ""
   TxtEmpComm.Text = ""
   ChkIsProduct.Value = 1
   TxtCost.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtPID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("code").Text
      TxtProductName.Text = .Columns("ProductName").Text
      
      If Trim(.Columns("ColourName").Text) <> "" Then
         CmbColourName.AddItem .Columns("ColourName").Text
         CmbColourName.ItemData(CmbColourName.NewIndex) = .Columns("ColourID").Text
         CmbColourName.ListIndex = 0
      End If
      
      If Trim(.Columns("SizeName").Text) <> "" Then
         cmbSizeName.AddItem .Columns("ColourName").Text
         cmbSizeName.ItemData(cmbSizeName.NewIndex) = .Columns("SizeID").Text
         cmbSizeName.ListIndex = 0
      End If
      
      TxtQty.Text = .Columns("Qty").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      TxtCost.Text = .Columns("Cost").Value
      TxtEmpComm.Text = .Columns("EmpComm").Value
      TxtAmount.Text = .Columns("Amount").Text
      TxtActualAmount.Text = .Columns("TotalAmount").Text
      
      TxtSaleTaxPer.Text = .Columns("SaleTaxPer").Value
      TxtSaleTaxValue.Text = .Columns("SaleTaxVal").Value
      vIsWSDiscb4ST = .Columns("IsWSDiscb4ST").Value
      vIsWSSaleTax = .Columns("IsWSSaleTax").Value
      vIsRetailSaleTax = .Columns("IsRetailSaleTax").Value

      ChkIsProduct.Value = Abs(.Columns("IsProduct").Value)
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
   vStrSQL = "select isnull(dbo.FunStock('" & TxtPID.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & Date + 1 & "',0),0)"
         vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         
         LblStock.Caption = vQtyLoose & " " & CN.Execute("SELECT dbo.FunGetUnit('" & TxtPID.Text & "')").Fields(0).Value
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock

   If Grid.rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub GetSaleReturn()
'   On Error GoTo ErrorHandler
'   sSql = "select h.*, p.partyname, StoreName FROM SaleReturnHeader h left outer join parties p on h.customerid=p.partyid inner join stores s on s.storeid = h.storeid where h.ReturnID=" & Val(TxtReturnID.Text) & " and ReturnDate='" & DtpReturnDate.DateValue & "'"
'   With CN.Execute(sSql)
'      If Not .BOF Then
'          TxtBillID.Text = IIf(IsNull(!BillId), "", !BillId)
'          If Not IsNull(!BillDate) Then
'             DtpBillDate.DateValue = !BillDate
'          End If
'          DtpBillDate.DateValue = IIf(IsNull(!BillDate), Null, !BillDate)
'          TxtStoreID.Text = !StoreID
'          TxtStoreName.Text = !StoreName
'          TxtTotalAmount.Caption = !TotalAmount
'          TxtBillDisc.Text = IIf(IsNull(!BillDiscount), "", !BillDiscount)
'          FrmReturnPrint.OptCash.Value = !Cash
'          FrmReturnPrint.OptCredit.Value = !Credit
'          FrmReturnPrint.TxtCashPaid.Text = IIf(IsNull(!CashPaid), "", !CashPaid)
'          FrmReturnPrint.TxtCustomerID.Text = IIf(IsNull(!CustomerID), "", !CustomerID)
'          FrmReturnPrint.TxtCustomerName.Text = IIf(IsNull(!PartyName), "", !PartyName)
'          FrmReturnPrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
'          TxtNetAmount.Caption = !TotalAmount
'      End If
'      .Close
'   End With
'   Call PopulateDataToGrid
'   FormStatus = OpenMode
'   Exit Sub
'ErrorHandler:
'   Grid.Redraw = True
'   Call ShowErrorMessage
'End Sub

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

Private Sub TxtBillID_Change()
   If Trim(TxtBillID.Text) = "" Then DtpBillDate.Enabled = False Else DtpBillDate.Enabled = True
End Sub

Private Sub TxtDiscPC_Change()
   If ActiveControl.Name <> TxtDiscPC.Name Then Exit Sub
   If Val(TxtPrice.Text) = 0 Then Exit Sub
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
   Call SubCalculateBody
End Sub

'Private Sub TxtDiscPC_LostFocus()
'   Select Case Me.ActiveControl.Name
'   Case TxtCode.Name, TxtQty.Name, TxtDiscPC.Name
'      Exit Sub
'   End Select
'   Call GetDataFromTexBoxesToGrid
'End Sub

Private Sub TxtDiscPer_Change()
   If ActiveControl.Name <> TxtDiscPer.Name Then Exit Sub
   TxtDiscPC.Text = Round((Val(TxtPrice.Text) * Val(TxtDiscPer.Text) / 100), 2)
   Call SubCalculateBody
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtCode.Text = ""
      TxtPID.Text = ""
      TxtProductName.Text = ""
      TxtPrice.Text = ""
      TxtDiscPC.Text = ""
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
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = Not FunSelectProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtPrice_Change()
   TxtDiscPC.Text = Round((Val(TxtPrice.Text) * Val(TxtDiscPer.Text) / 100), 2)
   Call SubCalculateBody
End Sub

Private Sub TxtQty_Change()
   Call SubCalculateBody
   Call FindRebate
End Sub

Private Sub TxtServiceCharges_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtServiceCharges.Name Then Exit Sub
   TxtServiceChargesPer.Text = Round((Val(TxtServiceCharges.Text) * 100) / Val(TxtTotalAmount.Caption), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtServiceChargesPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtServiceChargesPer.Name Then Exit Sub
   TxtServiceCharges.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtServiceChargesPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSSID_LostFocus()
On Error GoTo ErrorHandler
   ssql = "Select * from saleHeader where sid = " & Val(TxtSSID.Text)
   With CN.Execute(ssql)
      If .EOF = False Then
         TxtBillID.Text = !BillID
         DtpBillDate.DateValue = !BillDate
         vSessionID = IIf(IsNull(!SessionID), 0, !SessionID)
         GetSale
      End If
   End With
 Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub TxtSTax_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSTax.Name Then Exit Sub
   TxtSTaxPer.Text = Round((Val(TxtSTax.Text) * 100) / Val(TxtTotalAmount.Caption), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSTaxPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSTaxPer.Name Then Exit Sub
   TxtSTax.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtSTaxPer.Text) / 100))
   Call SubCalculateFooter
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

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpBillDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = CN.Execute("Select isnull(max(BinID),0)+1 from Bin_SaleReturnHeader ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
     If vIsNewRecord = False Then
    With CN.Execute("Select  * from SaleReturnHeader where ReturnID =" & TxtReturnID.Text & " And ReturnDate = '" & DtpReturnDate.DateValue & "'")
        If Val(TxtEmployeeID.Text) <> IIf(IsNull(!EmpID), 0, !EmpID) Then
            CN.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated EmpID-" & !EmpID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        
        If TxtStoreID.Text <> !StoreID Then
            CN.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    Grid.MoveFirst
    For i = 1 To Grid.rows - 1
        With CN.Execute("Select * from SaleReturnBody Where ReturnID = " & TxtReturnID.Text & " and ReturnDate ='" & DtpReturnDate.DateValue & "' and Productid = " & Val(Grid.Columns("Productid").Text))
        
             If .EOF = True Then
                CN.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Inserted New Code-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("Qty").Text <> !Qty Or Grid.Columns("Price").Text <> !Price Or Grid.Columns("discper").Text <> !DiscPer Then
                   CN.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Updated Code-" & Grid.Columns("Code").Text & " Qty-" & !Qty & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
    
   Else
    CN.Execute ("Insert Into UserActivities values ('Sale Return Invoice'" & "," & TxtReturnID.Text & ",'" & DtpReturnDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub BtnTable_Click()
   On Error GoTo ErrorHandler
   If FunSelectTable(ssButton, False) = True Then
      TxtTableID.SetFocus
   Else
      TxtTableID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectTable(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchTable.ParaInQuery = ""
        SchTable.Show vbModal, Me
        If SchTable.ParaOutTableID = "" Then FunSelectTable = False: Exit Function
        TxtTableID.Text = SchTable.ParaOutTableID
    End If
    '---------------------------
    If Trim(TxtTableID.Text) = "" Then Exit Function
    ssql = "Select TableID, TableName FROM Tables" & vbCrLf _
            + "where TableID = " & Val(TxtTableID.Text)
    With CN.Execute(ssql)
      If .RecordCount > 0 Then
        TxtTableName.Text = !TableName
        FunSelectTable = True
        .Close
        Exit Function
      Else
        FunSelectTable = False
        .Close
        TxtTableID.Text = ""
        TxtTableName.Text = ""
        Exit Function
      End If
    End With
Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub TxtTableID_Change()
   If ActiveControl.Name <> TxtTableID.Name Then Exit Sub
   If TxtTableName.Text <> "" Then TxtTableName.Text = ""
End Sub

Private Sub TxtTableID_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If TxtTableName.Text <> "" Then Exit Sub
    If TxtTableID.Text = "" Then Exit Sub
    Dim vTemp As Boolean
    vTemp = Not FunSelectTable(ssValidate, True)
    If vTemp = True Then
        vTemp = Not FunSelectTable(ssButton, False)
    End If
    Cancel = vTemp
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub TxtTotalAmount_Click()
   On Error GoTo ErrorHandler
   If Len(TxtTotalAmount.Caption) >= 6 Then
      TxtTotalAmount.FontSize = 36
   Else
      TxtTotalAmount.FontSize = 48
   End If
   TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtBillDiscPer.Text) / 100))
   TxtSTax.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtSTaxPer.Text) / 100))
   TxtServiceCharges.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtServiceChargesPer.Text) / 100))
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub FindRebate()
   Dim Rebate
   On Error GoTo ErrorHandler
    With CN.Execute("Select * from ProductOffers where Rebate <> 0 and ProductID = " & Val(TxtPID.Text))
        If .RecordCount > 0 Then
            Rebate = Val(TxtQty.Text)
            Rebate = Rebate \ !Qty
            Rebate = Rebate * !Rebate
            TxtDiscVal.Text = Rebate
            If Val(TxtPrice.Text) = 0 Then Exit Sub
            If Val(TxtQty.Text) = 0 Then Exit Sub
            TxtDiscPC.Text = Round(Val(TxtDiscVal.Text) / (TxtQty.Text), 3)
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
            TxtActualAmount.Text = Val(TxtQty.Text) * Val(TxtPrice.Text)
            TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
            TxtTotalDiscount.Caption = vTotDisc
            SubCalculateFooter
        End If
    End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      LblCost.Visible = False
   End If
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
    
        vStrSQL = " Select c.* FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " Left Outer join Members m on c.AccountNo = cast(m.Prefix as varchar(2))  + cast(m.MemberID as varchar(10)) " & vbCrLf & _
              " where p.BarCode = '" & (FrmReturnPrint.TxtCashCustomer.Text) & "' or m.BarCode = '" & (FrmReturnPrint.TxtCashCustomer.Text) & "' or (c.AccountNo = " & (FrmReturnPrint.TxtCashCustomer.Text) & " and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo like '3%') and c.isDetailed = 1 and c.isLocked = 0)"

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

Private Sub SubApplyEmployeeCommision()
   On Error GoTo ErrorHandler
   Grid.Redraw = False
   Grid.MoveFirst
      ssql = " select * " & vbCrLf _
            + " from Products where ProductID = " & Grid.Columns("ProductID").Text
      With CN.Execute(ssql)
         While Trim(Grid.Columns("ProductID").Text) <> ""
            RsBody.Filter = "ProductID = " & Val(Grid.Columns("ProductID").Text)
            Grid.Columns("EmpComm").Value = IIf(IsNull(!EmpComm), "", !EmpComm)
            Grid.MoveNext
         Wend
         .Close
      End With
   Grid.MoveLast
   Grid.Redraw = True
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubDestroyEmployeeCommision()
   On Error GoTo ErrorHandler
   Grid.Redraw = False
   Grid.MoveFirst
'   ssql = " select * " & vbCrLf _
         + " from Products"
   
   For vCounter = 1 To Grid.rows
      If Trim(Grid.Columns("ProductID").Text) <> "" Then
'         .Filter = "ProductID = '" & Grid.Columns("ProductID").Text & "'"
'         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
'            RsBody.Filter = "ProductID='" & !Productid & "'"
            Grid.Columns("EmpComm").Value = 0
'            RsBody!EmpComm = Null
'         End If
    End If
    Grid.MoveNext
    Next vCounter
   
   Grid.Redraw = True
   Grid.MoveLast
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubApplyMember()
   On Error GoTo ErrorHandler
   Dim vAmount, vDiscVal As Double
   Grid.MoveFirst
   ssql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With CN.Execute(ssql)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = " & Val(Grid.Columns("ProductID").Text)
         If .RecordCount > 0 Then
'            vDiscVal = Val(Grid.Columns("DiscVal").Value)
            'GetDataBackFromGridToTexBoxes
'            RsBody.Filter = "ProductID='" & !Productid & "'"
            vDiscVal = Val(Grid.Columns("DiscVal").Value)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
            Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("Price").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
            Grid.Columns("DiscVal").Value = Val(Grid.Columns("DiscPC").Value) * Val(Grid.Columns("Qty").Value)
            'Grid.Columns("SC").Value = IIf(IsNull(!Sc), 0, !Sc)
            vAmount = Val(Grid.Columns("Amount").Value)
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) + Val(Grid.Columns("SC").Value) - Val(Grid.Columns("DiscVal").Value)
            
            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - vAmount + Val(Grid.Columns("Amount").Text)
            vTotDisc = vTotDisc - vDiscVal + Val(Grid.Columns("DiscVal").Text)
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
'   TxtBillDisc.Text = Val(TxtBillDisc.Text) + vTotDisc
'   TxtBillDiscPer.Text = Val(TxtBillDisc.Text) / IIf(Val(TxtTotalAmount.Caption) = 0, 1, Val(TxtNetAmount.Caption)) * 100
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
            Grid.Columns("Amount").Value = (Val(Grid.Columns("Price").Value) * Val(Grid.Columns("Qty").Value)) + Val(Grid.Columns("SC").Value) - Val(Grid.Columns("DiscVal").Value)
       
            TxtNetAmount.Caption = Val(TxtNetAmount.Caption) - RsBody!Amount + Val(Grid.Columns("Amount").Text)
            vTotDisc = vTotDisc - RsBody!DiscVal + Val(Grid.Columns("DiscVal").Text)
            vTotalAmount = vTotalAmount - RsBody!Amount + Val(Grid.Columns("Amount").Text)
            
'            RsBody!DiscPC = Val(Grid.Columns("DiscPC").Value)
'            RsBody!DiscPer = Val(Grid.Columns("DiscPer").Value)
'            RsBody!DiscVal = Val(Grid.Columns("DiscVal").Value)
'            RsBody!Amount = Val(Grid.Columns("Amount").Value)
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

Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleReturnHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmSaleReturnInvoicePOS) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmSaleReturnInvoicePOS & ", " & vUser & "," & TableHeaderFields(eFrmSaleReturnInvoicePOS) & " from SaleReturnHeader " & vbCrLf _
             & "Where SID = " & TxtSID.Text
      CN.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleReturnBodyBin (" & TableBodyFields(eFrmSaleReturnInvoicePOS) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmSaleReturnInvoicePOS) & " from SaleReturnBody " & vbCrLf _
             & "Where SID = " & TxtSID.Text
      CN.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub TxtSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then GridSerial.SetFocus
End Sub

Private Sub GetDataFromTexBoxesToGridSerial()
   On Error GoTo ErrorHandler
   
     
   vStrSQL = "Select ProductID, Serial, SerialAdd from vuPurchaseSerial where Serial = '" & Trim(TxtSerial.Text) & "'"
      
      With CN.Execute(vStrSQL)
         If .EOF = True Then
            MsgBox "The Serail cannot be inserted because it is not Exist", vbInformation + vbOKOnly, "Error"
            TxtSerial.Text = ""
            Exit Sub
         ElseIf !SerialAdd = True Then
            MsgBox "The Serail Not Sold", vbInformation + vbOKOnly, "Error"
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
         RsBodySerial!Serial = TxtSerial.Text
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
            GridSerial.Columns("Serial").Text = !Serial
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

Private Sub SubClearSerialFields()
   TxtSerial.Text = ""
'   TxtSerial.Enabled = False
   GridSerial.CancelUpdate
   GridSerial.RemoveAll
   GridSerial.AddNew
   GridSerial.Columns("Serial").Text = " "
   GridSerial.Update
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

Private Sub PopulateDataPurchaseSerial()
   If RsPurchaseSerial.State = adStateOpen Then RsPurchaseSerial.Close
   vStrSQL = "select * from PurchaseBodySerial  "
   RsPurchaseSerial.Open vStrSQL, CN, adOpenDynamic, adLockBatchOptimistic
   RsPurchaseSerial.Filter = 0
End Sub

Private Sub PopulateDataReturnSerial()
   If RsReturnSerial.State = adStateOpen Then RsReturnSerial.Close
   vStrSQL = "select * from SaleReturnSerial "
   RsReturnSerial.Open vStrSQL, CN, adOpenDynamic, adLockBatchOptimistic
   RsReturnSerial.Filter = 0
End Sub

Private Function Webreq(postData As String) As String '(vStrUrl As String, vMessage As String, vCustNo As String)
   Dim vStrUrl As String, ReceiveText As String
   Dim winHttpReq As Object
   Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
   vStrUrl = "http://localhost:8524/api/IMSFiscal/GetInvoiceNumberByModel"
   winHttpReq.Open "POST", vStrUrl, False
   winHttpReq.setRequestHeader "Content-Type", "application/json"
   winHttpReq.Send (postData)
   ReceiveText = winHttpReq.responseText
   Webreq = Mid(ReceiveText, 19, InStr(19, ReceiveText, """") - 19)
End Function
