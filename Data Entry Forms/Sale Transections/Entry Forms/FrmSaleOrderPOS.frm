VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmSaleOrderPOS 
   BorderStyle     =   0  'None
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   Icon            =   "FrmSaleOrderPOS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   12630
      TabIndex        =   117
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   9375
      Width           =   2115
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmSaleOrderPOS.frx":0ECA
      Left            =   11565
      List            =   "FrmSaleOrderPOS.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   116
      Tag             =   "1"
      Top             =   9825
      Width           =   3276
   End
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12615
      TabIndex        =   115
      Top             =   8865
      Width           =   1245
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      Left            =   4545
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   8475
      Width           =   1950
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1845
      MaxLength       =   50
      TabIndex        =   87
      Top             =   9480
      Visible         =   0   'False
      Width           =   4125
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5985
      Top             =   9285
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
      Left            =   13725
      TabIndex        =   78
      Top             =   990
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
         TabIndex        =   79
         Tag             =   "NC"
         Text            =   "FrmSaleOrderPOS.frx":0ECE
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
         TabIndex        =   80
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.CheckBox ChkIsProduct 
      Caption         =   "Is Product"
      Height          =   255
      Left            =   13590
      TabIndex        =   64
      Top             =   7650
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1050
   End
   Begin SITextBox.Txt TxtOrderID 
      Height          =   315
      Left            =   1920
      TabIndex        =   29
      Top             =   1680
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
   Begin SITextBox.Txt TxtDiscVal 
      Height          =   315
      Left            =   10935
      TabIndex        =   11
      Top             =   2535
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1980
      TabIndex        =   6
      Top             =   2535
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
      Masked          =   1
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   7680
      TabIndex        =   7
      Top             =   2535
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
      Left            =   8460
      TabIndex        =   8
      Top             =   2535
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
      Left            =   11925
      TabIndex        =   30
      Top             =   2535
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
      Left            =   3825
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2535
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
      MICON           =   "FrmSaleOrderPOS.frx":1045
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9135
      TabIndex        =   27
      Top             =   8865
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
      MICON           =   "FrmSaleOrderPOS.frx":1061
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7815
      TabIndex        =   23
      Top             =   8865
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
      MICON           =   "FrmSaleOrderPOS.frx":107D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5175
      TabIndex        =   25
      Top             =   8865
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
      MICON           =   "FrmSaleOrderPOS.frx":1099
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10455
      TabIndex        =   28
      Top             =   8865
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
      MICON           =   "FrmSaleOrderPOS.frx":10B5
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6495
      TabIndex        =   24
      Top             =   8865
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
      MICON           =   "FrmSaleOrderPOS.frx":10D1
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBillDisc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1935
      TabIndex        =   12
      Top             =   6765
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
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   4185
      TabIndex        =   42
      Top             =   2535
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
      Height          =   3540
      Left            =   1965
      TabIndex        =   43
      Top             =   2850
      Width           =   11625
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   15
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
      stylesets(0).Picture=   "FrmSaleOrderPOS.frx":10ED
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
      ExtraHeight     =   106
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   15
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
      Columns(4).Width=   1376
      Columns(4).Caption=   "Qty"
      Columns(4).Name =   "Qty"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   4
      Columns(4).FieldLen=   256
      Columns(5).Width=   1693
      Columns(5).Caption=   "Price"
      Columns(5).Name =   "Price"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   4
      Columns(5).FieldLen=   256
      Columns(6).Width=   1455
      Columns(6).Caption=   "Disc / Pc"
      Columns(6).Name =   "DiscPC"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1217
      Columns(7).Caption=   "Disc%"
      Columns(7).Name =   "DiscPer"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1746
      Columns(8).Caption=   "Disc. Val"
      Columns(8).Name =   "DiscVal"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   4
      Columns(8).FieldLen=   256
      Columns(9).Width=   2461
      Columns(9).Caption=   "Amount"
      Columns(9).Name =   "Amount"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   5
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "TotalAmount"
      Columns(10).Name=   "TotalAmount"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Visible=   0   'False
      Columns(11).Caption=   "Cost"
      Columns(11).Name=   "Cost"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   4
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "QtyOrigional"
      Columns(12).Name=   "QtyOrigional"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   4
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "IsProduct"
      Columns(13).Name=   "IsProduct"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   11
      Columns(13).FieldLen=   256
      Columns(13).Style=   2
      Columns(14).Width=   3200
      Columns(14).Visible=   0   'False
      Columns(14).Caption=   "EmpComm"
      Columns(14).Name=   "EmpComm"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20505
      _ExtentY        =   6244
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
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
      Left            =   3855
      TabIndex        =   26
      Top             =   8865
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
      MICON           =   "FrmSaleOrderPOS.frx":1109
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   9420
      TabIndex        =   9
      Top             =   2535
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
      Left            =   9045
      TabIndex        =   49
      Top             =   9525
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
      Left            =   6165
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   1680
      Width           =   585
      _ExtentX        =   1032
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
      Left            =   7110
      TabIndex        =   51
      Tag             =   "NC"
      Top             =   1680
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
      Left            =   6750
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1680
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
      MICON           =   "FrmSaleOrderPOS.frx":1125
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpOrderDate 
      Height          =   315
      Left            =   2565
      TabIndex        =   31
      Top             =   1680
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
      Left            =   10245
      TabIndex        =   10
      Top             =   2535
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
      Left            =   10395
      TabIndex        =   56
      Top             =   8985
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
      Left            =   8205
      TabIndex        =   58
      Top             =   9510
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
      Left            =   2775
      TabIndex        =   13
      Top             =   6765
      Width           =   525
      _ExtentX        =   926
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
      Left            =   10950
      TabIndex        =   4
      Top             =   1680
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
      Left            =   12060
      TabIndex        =   70
      Top             =   1680
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
      Left            =   11700
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   1680
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
      MICON           =   "FrmSaleOrderPOS.frx":1141
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCommission 
      Height          =   315
      Left            =   7380
      TabIndex        =   74
      Top             =   9510
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
      Left            =   8505
      TabIndex        =   3
      Top             =   1680
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
      Left            =   9540
      TabIndex        =   82
      Top             =   1680
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
      Left            =   9180
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1680
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
      MICON           =   "FrmSaleOrderPOS.frx":115D
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtManualBillNo 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   13545
      TabIndex        =   22
      Top             =   7260
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
   Begin SITextBox.Txt TxtRemarks 
      Height          =   315
      Left            =   6705
      TabIndex        =   20
      Top             =   8385
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
   Begin SITextBox.Txt TxtServiceCharges 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1935
      TabIndex        =   16
      Top             =   7890
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
      Left            =   2775
      TabIndex        =   17
      Top             =   7890
      Width           =   525
      _ExtentX        =   926
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
      Left            =   1935
      TabIndex        =   14
      Top             =   7335
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
      Left            =   2775
      TabIndex        =   15
      Top             =   7335
      Width           =   525
      _ExtentX        =   926
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
      Left            =   6435
      TabIndex        =   96
      Top             =   9555
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
   Begin SITextBox.Txt TxtTableID 
      Height          =   315
      Left            =   1935
      TabIndex        =   18
      Top             =   8475
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
      Left            =   2820
      TabIndex        =   97
      Top             =   8475
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
      Left            =   2460
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   8475
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
      MICON           =   "FrmSaleOrderPOS.frx":1179
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   8505
      TabIndex        =   5
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
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   9570
      TabIndex        =   101
      Tag             =   "NC"
      Top             =   1230
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
      Left            =   9210
      TabIndex        =   102
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
      MICON           =   "FrmSaleOrderPOS.frx":1195
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpDeliveryDate 
      Height          =   315
      Left            =   3870
      TabIndex        =   0
      Top             =   1680
      Width           =   1260
      _Version        =   65543
      _ExtentX        =   2222
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
   Begin MSComCtl2.DTPicker DTPDeliveryTime 
      Height          =   315
      Left            =   5130
      TabIndex        =   1
      Top             =   1680
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "hh:mm tt"
      Format          =   117374979
      UpDown          =   -1  'True
      CurrentDate     =   39224.0416666667
   End
   Begin SITextBox.Txt TxtStampID 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6750
      TabIndex        =   108
      Top             =   1230
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
      Alignment       =   1
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
      IntegralPoint   =   4
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   5310
      TabIndex        =   110
      Top             =   2190
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpBillDate 
      Height          =   315
      Left            =   5880
      TabIndex        =   111
      Top             =   2190
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
   Begin JeweledBut.JeweledButton BtnSaleInvoice 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7200
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   2175
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
      MICON           =   "FrmSaleOrderPOS.frx":11B1
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpPromiseDate 
      Height          =   315
      Left            =   3870
      TabIndex        =   120
      Top             =   1155
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
      Left            =   3915
      TabIndex        =   121
      Top             =   945
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
      Left            =   12630
      TabIndex        =   119
      Top             =   9075
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
      Left            =   10890
      TabIndex        =   118
      Top             =   9870
      Width           =   570
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   195
      Left            =   5880
      TabIndex        =   114
      Top             =   1995
      Width           =   585
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill ID"
      Height          =   195
      Left            =   5325
      TabIndex        =   113
      Top             =   1995
      Width           =   405
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Stamp ID"
      Height          =   195
      Left            =   6075
      TabIndex        =   109
      Top             =   1320
      Visible         =   0   'False
      Width           =   660
   End
   Begin MSForms.TextBox TxtRemarksUrdu 
      Height          =   435
      Left            =   6705
      TabIndex        =   21
      ToolTipText     =   "Textbox1"
      Top             =   8340
      Visible         =   0   'False
      Width           =   7020
      VariousPropertyBits=   618678299
      ForeColor       =   0
      MaxLength       =   100
      BorderStyle     =   1
      Size            =   "12382;767"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Time"
      Height          =   195
      Left            =   5085
      TabIndex        =   107
      Top             =   1500
      Width           =   960
   End
   Begin VB.Label LblType 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Left            =   4545
      TabIndex        =   106
      Top             =   8250
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
      Height          =   195
      Left            =   3870
      TabIndex        =   105
      Top             =   1500
      Width           =   960
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   9690
      TabIndex        =   104
      Top             =   1005
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   8505
      TabIndex        =   103
      Top             =   1005
      Width           =   1095
   End
   Begin VB.Label LblTableID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Table ID"
      Height          =   195
      Left            =   1935
      TabIndex        =   100
      Top             =   8295
      Width           =   615
   End
   Begin VB.Label LblTableName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name"
      Height          =   195
      Left            =   2775
      TabIndex        =   99
      Top             =   8295
      Width           =   870
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "EmpComm"
      Height          =   195
      Left            =   6435
      TabIndex        =   95
      Top             =   9330
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   2790
      TabIndex        =   94
      Top             =   7110
      Width           =   210
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
      Height          =   195
      Left            =   1935
      TabIndex        =   93
      Top             =   7110
      Width           =   705
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Ch."
      Height          =   195
      Left            =   1935
      TabIndex        =   92
      Top             =   7665
      Width           =   825
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   2790
      TabIndex        =   91
      Top             =   7665
      Width           =   210
   End
   Begin VB.Label LblRemarks 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   6705
      TabIndex        =   90
      Top             =   8160
      Width           =   630
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill No"
      Height          =   195
      Left            =   13545
      TabIndex        =   89
      Top             =   7035
      Width           =   1020
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
      Height          =   225
      Left            =   1860
      TabIndex        =   88
      Top             =   9240
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
      Left            =   1890
      TabIndex        =   86
      Top             =   8970
      Width           =   165
   End
   Begin VB.Label LblMemberName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      Height          =   195
      Left            =   9540
      TabIndex        =   85
      Top             =   1500
      Width           =   1035
   End
   Begin VB.Label LblMemberID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      Height          =   195
      Left            =   8505
      TabIndex        =   84
      Top             =   1500
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
      Left            =   13005
      TabIndex        =   81
      Top             =   1365
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
      Left            =   11700
      TabIndex        =   77
      Top             =   960
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
      Left            =   11925
      TabIndex        =   76
      Top             =   1275
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      Height          =   195
      Left            =   7245
      TabIndex        =   75
      Top             =   9285
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblEmpName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   12060
      TabIndex        =   73
      Top             =   1500
      Width           =   780
   End
   Begin VB.Label LblEmpID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   10950
      TabIndex        =   72
      Top             =   1500
      Width           =   525
   End
   Begin VB.Label TxtTotalQty 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "VCRSCapsSSK"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   915
      Left            =   4515
      TabIndex        =   69
      Top             =   7170
      Width           =   1380
   End
   Begin VB.Label TxtTotalAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "VCRSCapsSSK"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   915
      Left            =   5940
      TabIndex        =   68
      Top             =   7170
      Width           =   2730
   End
   Begin VB.Label TxtTotalDiscount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "VCRSCapsSSK"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   8715
      TabIndex        =   67
      Top             =   7170
      Width           =   1740
   End
   Begin VB.Label TxtNetAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "VCRSCapsSSK"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   10500
      TabIndex        =   66
      Top             =   7170
      Width           =   2730
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      Height          =   195
      Left            =   2790
      TabIndex        =   65
      Top             =   6540
      Width           =   210
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Order POS"
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
      TabIndex        =   63
      Top             =   270
      Width           =   2655
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
      Left            =   2940
      TabIndex        =   62
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
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
      Left            =   11070
      TabIndex        =   61
      Top             =   1995
      Width           =   720
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
      Left            =   11880
      TabIndex        =   60
      Top             =   1995
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   8235
      TabIndex        =   59
      Top             =   9285
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   10665
      TabIndex        =   57
      Top             =   9330
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. %"
      Height          =   195
      Left            =   10230
      TabIndex        =   55
      Top             =   2340
      Width           =   525
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   6165
      TabIndex        =   54
      Top             =   1500
      Width           =   585
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   7110
      TabIndex        =   53
      Top             =   1500
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Amount"
      Height          =   195
      Left            =   9045
      TabIndex        =   50
      Top             =   9315
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
      Left            =   8715
      TabIndex        =   48
      Top             =   6870
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
      Left            =   6525
      TabIndex        =   47
      Top             =   6870
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   8460
      TabIndex        =   46
      Top             =   2340
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
      Left            =   4515
      TabIndex        =   45
      Top             =   6840
      Width           =   1365
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   13320
      Top             =   930
      Width           =   345
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   4185
      TabIndex        =   44
      Top             =   2340
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
      Left            =   11130
      TabIndex        =   41
      Top             =   6870
      Width           =   1440
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Order Disc."
      Height          =   195
      Left            =   1935
      TabIndex        =   40
      Top             =   6540
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
      Height          =   195
      Left            =   2565
      TabIndex        =   39
      Top             =   1500
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID"
      Height          =   195
      Left            =   1920
      TabIndex        =   38
      Top             =   1500
      Width           =   600
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   1965
      TabIndex        =   37
      Top             =   2340
      Width           =   375
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc / PC"
      Height          =   195
      Left            =   9420
      TabIndex        =   36
      Top             =   2340
      Width           =   690
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   7680
      TabIndex        =   35
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   11925
      TabIndex        =   34
      Top             =   2340
      Width           =   540
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. Val"
      Height          =   195
      Left            =   10935
      TabIndex        =   33
      Top             =   2340
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
Attribute VB_Name = "FrmSaleOrderPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim vMode As FormMode
Dim vCounter, vGridRows As Integer
Dim vDate, vServerDate As Date, vHDiff As Integer, vSystemDate As Boolean
'Dim RsDetail As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vMaxBinID As Integer
Dim vIsNewRecord As Boolean
Dim Flag As Boolean, vIsEdit, vAllowNegativeOrder As Boolean
Dim vSave As Boolean
Dim vBm As Variant
Dim UniCode As Variant
Dim DateFlag As Boolean, vAutoEnterBeforeQty, vShowStock As Boolean
'Dim vSystemDate As Date
Dim ssql As String, vWhere As String
Dim vStrSQL, vRandomID  As String, vAutoPrintSaleOrder As Boolean, vPrintKitchenInoices, vIsDisableCreditSale, vIsCreditSale As Boolean
Dim vQtyLoose As Double, vTotalAmount As Double
Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String, vTotDisc As Double
Dim i As Integer, vCashDrawer As Boolean, vLaserInvoice As Boolean, vPrintHeader  As Boolean, vNoofPrints As Byte, vX As Integer, vY As Integer
Dim vPrinter() As String
'----------------------------------

Private Sub FindRow()
   Dim vBm As Variant
   Dim lTotal As Long
   Dim i As Integer, vFind As String
   
   'vFind = InputBox("Enter Code", "Find Code")
   
   vBm = Grid.Bookmark
   Grid.MoveFirst
   
   For i = 0 To Grid.rows - 1
      'If Val(vFind) = Val(Grid.Columns("ProductID").CellValue(Grid.GetBookmark(i))) Or Val(vFind) = Val(Grid.Columns("Code").CellValue(Grid.GetBookmark(i))) Then
      If (Grid.Columns("ProductName").CellValue(Grid.GetBookmark(i))) Like TxtProductName.Text & "*" Then
         'MsgBox "1"
         Grid.Bookmark = Grid.GetBookmark(i)
         Exit Sub
      End If
   Next i
   Grid.Bookmark = vBm
End Sub

Private Sub SubCalculateBody()
    TxtDiscVal.Text = Val(TxtQty.Text) * Val(TxtDiscPC.Text)
    TxtActualAmount.Text = Val(TxtQty.Text) * Val(TxtPrice.Text)
    TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
    TxtTotalDiscount.Caption = vTotDisc
    SubCalculateFooter
End Sub

'Private Sub SubMakePackageDeal()
'   Dim RsTemp As New ADODB.Recordset
'   'Grid.Redraw = False
'   vBm = Grid.Bookmark
'   Grid.MoveFirst
'   ssql = " select * " & vbCrLf _
'         + " from PackageDealInfoBody b inner join PackageDealInfoHeader h on h.id = b.id"
'   With CN.Execute(ssql)
'      Grid.MoveFirst
'      While Grid.Columns("ProductID").Text <> ""
'         .Filter = "ProductID = " & val(Grid.Columns("ProductID").Text )
'         If .RecordCount > 0 Then
'            RsDetail.AddNew
'            RsDetail!Productid = Grid.Columns("ProductID").Text
'            RsDetail!Rate = Grid.Columns("Price").Text
'            RsDetail!QtyLoose = Grid.Columns("Qty").Text
'            RsDetail!Amount = Grid.Columns("Amount").Text
'            RsDetail.Update
'            RsBody.Filter = "ProductID = " & val(Grid.Columns("ProductID").Text )
'            If RsBody.RecordCount > 0 Then RsBody.Delete
'            Grid.SelBookmarks.RemoveAll
'            Grid.SelBookmarks.Add Grid.Bookmark
'            Grid.DeleteSelected
'         Else
'            Grid.MoveNext
'         End If
'      Wend
'      .Filter = "ProductID = " & RsDetail!Productid
'      If .RecordCount > 0 Then
'         If RsTemp.State = adStateOpen Then RsTemp.Close
'         vStrSQL = " SELECT p.productid, ProductName, RetailPrice, DiscPer, DiscPC, EmpComm" & vbCrLf _
'               + " from PackageDealInfoHeader un inner join Products p on un.PackageDealid = p.productid" & vbCrLf _
'               + " where p.productid = '" & !PackageDealID & "'"
'
'         RsTemp.Open vStrSQL, CN, adOpenDynamic, adLockReadOnly
'         If RsTemp.RecordCount > 0 Then
'            TxtCode.Text = RsTemp!Productid
'            TxtPID.Text = RsTemp!Productid
'            TxtProductName.Text = RsTemp!ProductName
'            TxtPrice.Text = RsTemp!RetailPrice
'            TxtQty.Text = RsDetail!QtyLoose
'            TxtCost.Text = 0
'            TxtDiscPC.Text = IIf(IsNull(RsTemp!DiscPC), 0, RsTemp!DiscPC)
'            TxtDiscPer.Text = IIf(IsNull(RsTemp!DiscPer), 0, RsTemp!DiscPer)
'            TxtEmpComm.Text = IIf(IsNull(RsTemp!EmpComm), 0, RsTemp!EmpComm)
'            If Val(TxtDiscPC.Text) <> 0 Then
'               TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
'            End If
''            ChkIsProduct.Value = 0
'            SubCalculateBody
'            Grid.MoveLast
'            GetDataFromTexBoxesToGrid
'         End If
'      End If
'      .Close
'   End With
'   'RsDetail.Filter = 0
'   'Grid.Bookmark = vBm
'   'Grid.Redraw = True
'End Sub

Private Sub SubApplyMember()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   ssql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With cn.Execute(ssql)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = " & Val(Grid.Columns("ProductID").Text)
         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
            RsBody.Filter = "ProductID = " & Val(!Productid)
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
   ssql = " select * " & vbCrLf _
         + " from MembersDiscount "
   With cn.Execute(ssql)
      While Trim(Grid.Columns("ProductID").Text) <> ""
         .Filter = "ProductID = " & Val(Grid.Columns("ProductID").Text)
         If .RecordCount > 0 Then
            'GetDataBackFromGridToTexBoxes
            
            RsBody.Filter = "ProductID = " & Val(!Productid)
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
      .Close
   End With
   SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateFooter()
   TxtTotalDiscount.Caption = Val(TxtBillDisc.Text) + vTotDisc
   TxtNetAmount.Caption = SelfRound(Val(TxtTotalAmount.Caption) - Val(TxtTotalDiscount.Caption) + Val(TxtServiceCharges.Text) + Val(TxtSTax.Text))
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

Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectStore = False: Exit Function
        TxtStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Stores where islock = 0 and StoreID=" & Val(TxtStoreID.Text)
    With cn.Execute(vStrSQL)
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

Private Function FunSelectTable(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchTable.ParaInQuery = "Select distinct t.TableID, TableName FROM Tables t left outer join (select TableID from SaleOrderHeader where issale = 0)h on t.TableID = h.TableID Where h.TableId Is Null"
        SchTable.Show vbModal, Me
        If SchTable.ParaOutTableID = "" Then FunSelectTable = False: Exit Function
        TxtTableID.Text = SchTable.ParaOutTableID
    End If
    '---------------------------
    If Trim(TxtTableID.Text) = "" Then Exit Function
    ssql = "Select distinct t.TableID, TableName FROM Tables t left outer join (select TableID from SaleOrderHeader where issale = 0)h on t.TableID = h.TableID Where h.TableId Is Null" & vbCrLf _
            + " and t.TableID = " & Val(TxtTableID.Text)
    With cn.Execute(ssql)
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

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.ParaInWhere = " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & "  and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
      SchProduct.ParainShowStock = vShowStock
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
   If TxtCode.Enabled = False Then FunSelectProduct = False: Exit Function
   If Trim(TxtCode.Text) = "" Then FunSelectProduct = False: Exit Function
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    
    ''''''''***********   Checking PackageDeal   ***********''''''''
    vStrSQL = "SELECT p.productid, Code, ProductName, RetailPrice, DiscPer, DiscPC, EmpComm" & vbCrLf _
         + " from PackageDealInfoHeader un inner join Products p on un.PackageDealid = p.productid" & vbCrLf _
         + " left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
         + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')" & " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
         
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !RetailPrice
         TxtEmpComm.Text = IIf(IsNull(!EmpComm), "", !EmpComm)
         TxtQty.Text = IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         vStrSQL = " select sum(isnull(Cost,PurPrice)* b.QtyLoose) as Cost from PackageDealInfoHeader h inner join PackageDealInfoBody b on h.id = b.id" & vbCrLf _
               + " inner join Products p on p.productid = b.productid" & vbCrLf _
               + " left outer join CurrentStock cs on cs.productid = p.productid " & vbCrLf _
               + " where h.PackageDealid ='" & TxtPID.Text & "'"
         With cn.Execute(vStrSQL)
            If .RecordCount > 0 Then
               TxtCost.Text = !Cost
            Else
               TxtCost.Text = "0"
            End If
         End With
         
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & Val(TxtPID.Text)
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         End If
          LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & Val(TxtPID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & val(TxtPID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = IIf(Val(LblStock.Caption) = 0, "", LblStock.Caption) & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(TxtPID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
'         vStrSQL = " select isnull(Floor(min(css.QtyLoose/b.QtyLoose)),0) as QtyLoose " & vbCrLf _
'                  + " from PackageDealInfoHeader h inner join PackageDealInfoBody b on h.id = b.id" & vbCrLf _
'                  + " inner join Products p on p.productid = b.productid" & vbCrLf _
'                  + " left outer join CurrentStockStore css on css.productid = p.productid " & vbCrLf _
'                  + " where h.PackageDealid ='" & TxtPID.Text & "' and css.storeid = " & TxtStoreID.Text
'
'         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpOrderDate.DateValue + 1 & "',0),0)"
'         vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
'         LblStock.Caption = vQtyLoose & " " & cn.Execute("SELECT dbo.FunGetUnit(" & Val(TxtPID.Text) & ")").Fields(0).Value
'         LblStock.Visible = True
'         LblStockCaption.Visible = True
        
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
'               If !AllowNegativeOrder = False Then
'                  If vQtyLoose <= 0 Then
'                     MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                     FunSelectProduct = False
'                     Exit Function
'                  End If
'               End If
         If ObjRegistry.LastRateVisible = True Then
            If FrmOrderPrint.TxtCustomerID.Text <> "" Then
               LblPrice = cn.Execute("Select dbo.FunLastPrice('S','" & DtpOrderDate.DateValue & "'," & Val(TxtPID.Text) & "," & Val(FrmOrderPrint.TxtCustomerID.Text) & ")").Fields(0).Value
               LblCaptionPrice.Visible = True
               LblPrice.Visible = True
            End If
         End If
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtDiscPC.Text) <> 0 Then
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
         End If
         ChkIsProduct.Value = 0
         SubCalculateBody
'         Char.Speak TxtProductName.Text
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      End If
   End With
    
''''''''***********   Checking Product  ***********''''''''
    vStrSQL = " SELECT p.productid, code, Qty, ProductName, RetailPrice, DiscPer, DiscPC, EmpComm" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')" & " and isLocked = 0 " & IIf(ObjRegistry.ShowRawMaterialProductInSaleInvoices, "", " and isRawProduct = 0 ") & " and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
   
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = !RetailPrice
         TxtEmpComm.Text = IIf(IsNull(!EmpComm), "", !EmpComm)
         TxtQty.Text = IIf(Len(TxtCode.Text) <= 5 And IsNumeric(TxtCode.Text), 1, IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty))  'IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         'TxtQty.Text = IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = " & Val(TxtPID.Text)
            With cn.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpBillDate.DateValue + 1 & "',0),0)"
            vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         End If
          LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & Val(TxtPID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & val(TxtPID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = IIf(Val(LblStock.Caption) = 0, "", LblStock.Caption) & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(TxtPID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
'            If !AllowNegativeOrder = False Then
'               If vQtyLoose <= 0 Then
'                  MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                  FunSelectProduct = False
'                  Exit Function
'               End If
'            End If
         If ObjRegistry.LastRateVisible = True Then
            If FrmOrderPrint.TxtCustomerID.Text <> "" Then
               LblPrice = cn.Execute("Select dbo.FunLastPrice('S','" & DtpOrderDate.DateValue & "'," & Val(TxtPID.Text) & "," & Val(FrmOrderPrint.TxtCustomerID.Text) & ")").Fields(0).Value
               LblCaptionPrice.Visible = True
               LblPrice.Visible = True
            End If
         End If
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtDiscPC.Text) <> 0 Then
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
         End If
         ChkIsProduct.Value = 1
         SubCalculateBody
 '        Char.Speak TxtProductName.Text
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
   If MsgBox("Are you sure to Clear the Data?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Exit Sub
   '''''''''''''''''' ActivityLogBin For Clear Action
'      Call DeleteTempActivityLogBin(vRandomID)
      vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From SaleOrderbody where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
            With cn.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmSaleOrderPOS, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Cleared Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmSaleOrderPOS, eClearSavedRecord, TxtOrderID.Text, DtpOrderDate.DateValue, vGridRows & " Product/s Cleared")
      Grid.Redraw = True
  ''''''''''''''''''
   ''''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   ''''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
    ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSaleOrderPOS", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
   cn.BeginTrans
   
   Call BinData
  Call ActivityLogBin("", eFrmSaleOrderPOS, eDelete, TxtOrderID.Text, DtpOrderDate.DateValue, Grid.rows - 1 & " Product/s Deleted Amount: " & Val(TxtNetAmount.Caption))
   
'   vMaxBinID = FunGetMaxBinID
'   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_SaleOrderHeader Select " & vMaxBinID & ",'" & Date & "',*," & 1 & "," & 0 & " from SaleOrderHeader Where OrderID = " & TxtOrderID.Text & " And OrderDate ='" & DtpOrderDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_SaleOrderBody Select " & vMaxBinID & ",'" & Date & "', * from SaleOrderBody Where OrderID = " & TxtOrderID.Text & " And OrderDate ='" & DtpOrderDate.DateValue & "'")
'   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Grid.Redraw = False
   Grid.MoveFirst
   Call ActivityLog("Sale Order", eDelete, TxtOrderID.Text, DtpOrderDate.DateValue)
'   For vCounter = 1 To RsDetail.RecordCount
'      CN.Execute "Delete from SaleOrderUnionUsed where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "' and Productid ='" & RsDetail!Productid & "'"
'      RsDetail.MoveNext
'   Next vCounter
   For vCounter = 1 To Grid.rows
      If Trim(Grid.Columns("ProductID").Text) <> "" Then
         cn.Execute "Delete from SaleOrderBody where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "' and ProductID = " & Val(Grid.Columns("Productid").Text)
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   cn.Execute "Delete from SaleOrderHeader where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'"
   cn.CommitTrans
   FormStatus = NewMode
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

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchSaleOrder.ParaInOrderDate = DtpOrderDate.DateValue
   SchSaleOrder.Show vbModal
   If SchSaleOrder.ParaOutOrderID <> -1 Then
      TxtOrderID.Text = SchSaleOrder.ParaOutOrderID
      'Dim a
      'a = Split(SchSaleOrder.ParaOutOrderDate, "/")
      DtpOrderDate.DateValue = SchSaleOrder.ParaOutOrderDate 'Val(a(1)) & "/" & Val(a(0)) & "/" & Val(a(2))
      ''''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      CN.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      GetSale
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PrintDepartment()
   On Error GoTo ErrorHandler
   
   If vSave = True Then
   vStrSQL = " Select Distinct DepartmentID" & vbCrLf _
            + " from SaleOrderHeader h inner join TempSaleOrderBody b on h.OrderID = b.OrderID and h.OrderDate = b.OrderDate" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " where h.OrderID = " & Val(TxtOrderID.Text) & " and h.OrderDate = '" & DtpOrderDate.DateValue & "'" & vWhere & " Order By DepartmentID"
   Else
   vStrSQL = " Select Distinct DepartmentID" & vbCrLf _
            + " from SaleOrderHeader h inner join SaleOrderBody b on h.OrderID = b.OrderID and h.OrderDate = b.OrderDate" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " where h.OrderID = " & Val(TxtOrderID.Text) & " and h.OrderDate = '" & DtpOrderDate.DateValue & "'" & vWhere & " Order By DepartmentID"
   End If
 




   '/********* Watch *************/
'   cn.Execute "Insert into Watch(ErrorFrom,Narration) values ('PrintDepartment After Master Query','OrderID = " & TxtOrderID.Text & ", OrderDate = " & DtpOrderDate.DateValue & "')"
   
   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
   
   If ObjRegistry.LaserPrintofSaleInvoice = True Then
      Set RptReportViewer.Report = New CrpDepartmentInvoiceHalf
      RptReportViewer.Report.PaperSize = crPaperA4
      RptReportViewer.Report.PaperOrientation = crLandscape
      RptReportViewer.Report.TopMargin = vY
      RptReportViewer.Report.LeftMargin = vX
      RptReportViewer.Report.RightMargin = 225
   Else
      Set RptReportViewer.Report = New CrpDepartmentInvoice
   End If
   
   RptReportViewer.Report.ReportTitle = "Kitchen Order"
      
   With cn.Execute(vStrSQL)
      '/********* Watch *************/
'      cn.Execute "Insert into Watch(ErrorFrom,Narration) values ('PrintDepartment Before For Master Query Count = " & .RecordCount & "','OrderID = " & TxtOrderID.Text & ", OrderDate = " & DtpOrderDate.DateValue & "')"
      
      For i = 1 To .RecordCount
         If vSave = True Then
         vStrSQL = " select UserName, h.OrderID as billid, h.OrderDate as BillDate, isnull(h.OrderTime,0) as BillTime, isnull(p.ProductName1,p.ProductName) as ProductName, tp.qty, " & vbCrLf _
                  + " H.empid, isnull(EmpName,'') as EmpName, tp.ProductID, h.TableID, isnull(TableName,'') as TableName, isnull(Department,'') as Department, h.DeliveryDate, isnull(h.DeliveryTime,0) as DeliveryTime, isnull(b.isPrinted,0) as isPrinted, h.InvType, " & IIf(ObjRegistry.AllowUrduProduct = False, " isnull(Remarks,'')", " isnull(RemarksUrdu,'')") & "  as Remarks" & vbCrLf _
                  + " from SaleOrderHeader h inner join TempSaleOrderBody tp on h.OrderID = tp.OrderID and h.OrderDate = tp.OrderDate " & vbCrLf _
                  + " inner join SaleOrderBody b on tp.OrderID = b.OrderID and tp.OrderDate = b.OrderDate and tp.ProductID = b.ProductID " & vbCrLf _
                  + " inner join products p on p.productid = b.productid" & vbCrLf _
                  + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
                  + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
                  + " left outer join Tables t on t.TableID = h.TableID " & vbCrLf _
                  + " left outer join Departments d on d.DepartmentID = p.DepartmentID " & vbCrLf _
                  + " where h.OrderID = " & Val(TxtOrderID.Text) & " and h.OrderDate = '" & DtpOrderDate.DateValue & "'" & vWhere & IIf(IsNull(!DepartmentID) = True, " and d.DepartmentID is null ", " and d.DepartmentID = " & !DepartmentID) & " Order By SerialNo"
         Else
         vStrSQL = " select UserName, h.OrderID as billid, h.OrderDate as BillDate, isnull(h.OrderTime,0) as BillTime, isnull(p.ProductName1,p.ProductName) as ProductName, b.qty, " & vbCrLf _
                  + " H.empid, isnull(EmpName,'') as EmpName, b.ProductID, h.TableID, isnull(TableName,'') as TableName, isnull(Department,'') as Department, h.DeliveryDate, isnull(h.DeliveryTime,0) as DeliveryTime, isnull(b.isPrinted,0) as isPrinted, h.InvType, " & IIf(ObjRegistry.AllowUrduProduct = False, " isnull(Remarks,'')", " isnull(RemarksUrdu,'')") & "  as Remarks" & vbCrLf _
                  + " from SaleOrderHeader h " & vbCrLf _
                  + " inner join SaleOrderBody b on h.OrderID = b.OrderID and h.OrderDate = b.OrderDate " & vbCrLf _
                  + " inner join products p on p.productid = b.productid" & vbCrLf _
                  + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
                  + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
                  + " left outer join Tables t on t.TableID = h.TableID " & vbCrLf _
                  + " left outer join Departments d on d.DepartmentID = p.DepartmentID " & vbCrLf _
                  + " where h.OrderID = " & Val(TxtOrderID.Text) & " and h.OrderDate = '" & DtpOrderDate.DateValue & "'" & vWhere & IIf(IsNull(!DepartmentID) = True, " and d.DepartmentID is null ", " and d.DepartmentID = " & !DepartmentID) & " Order By SerialNo"
         End If
         If RsReport.State = adStateOpen Then RsReport.Close
         RsReport.Open vStrSQL, cn, adOpenDynamic, adLockReadOnly
         '/********* Watch *************/
'         cn.Execute "Insert into Watch(ErrorFrom,Narration) values ('PrintDepartment Before For and Record Count = " & RsReport.RecordCount & "','OrderID = " & TxtOrderID.Text & ", OrderDate = " & DtpOrderDate.DateValue & "')"
         If RsReport.RecordCount > 0 Then
            RptReportViewer.Report.DiscardSavedData
            RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
            RptReportViewer.Report.PrintOut False
         End If
         .MoveNext
      Next i
      '/********* Watch *************/
'      cn.Execute "Insert into Watch(ErrorFrom,Narration) values ('PrintDepartment After For Loop','OrderID = " & TxtOrderID.Text & ", OrderDate = " & DtpOrderDate.DateValue & "')"
   End With
   Exit Sub
ErrorHandler:
   '/********* Watch *************/
'   cn.Execute "Insert into Watch(ErrorFrom,Narration) values ('PrintDepartment Error','OrderID = " & TxtOrderID.Text & ", OrderDate = " & DtpOrderDate.DateValue & "')"
   Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   
      If ObjRegistry.PrintKitchenInoices = True Then
         If vSave = False Then
            PrintDepartment
         End If
      End If
   
   vStrSQL = " select UserName, h.OrderID as billid, h.OrderDate as BillDate, isnull(h.OrderTime,0) as BillTime, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, isnull(h.STax,0) as STax, isnull(h.cashReceived,0) as CashReceived, p.ProductName /*case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.OrderID, h.OrderDate) end */ ProductName, unitname, b.qty, b.price as price, b.amount, b.DiscPC, b.DiscPer, b.DiscVal, 0 as SC, InvoiceNo," & IIf(ObjRegistry.AllowUrduProduct = False, " isnull(h.Remarks,'')", " isnull(RemarksUrdu,'')") & " as Remarks" & vbCrLf _
            + " , '' as Desc1, Case when CustomerID = 621 then isnull(CustomerName,AccountName) Else AccountName End as Customer, pr.Description PartyDescription, H.empid, isnull(EmpName,'') as EmpName, Cash, Credit, BankCard, b.ProductID, h.MemberID, isnull(cast(h.MemberID as varchar(6)) + '-' + MemberName,'') as MemberName, h.TableID, isnull(TableName,'') as TableName, DeliveryDate, isnull(h.DeliveryTime,0) as DeliveryTime, h.InvType, isnull(h.isPrinted,0) as isPrinted, " & vbCrLf _
            + " isnull( pr.Phone1  + ', ','') + isnull( pr.Phone2 + ', ','')  + isnull( pr.mobile + ', ','') +  isnull( pr.mobile2 + ', ','') as mobile, pr.Address + pr.city Address " & vbCrLf _
            + " from SaleOrderHeader h inner join SaleOrderBody b on h.OrderID = b.OrderID and h.OrderDate = b.OrderDate" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
            + " left outer join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
            + " left outer join Members m on m.MemberID = h.MemberID" & vbCrLf _
            + " left outer join Tables t on t.TableID = h.TableID " & vbCrLf _
            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
            + " left outer join parties pr on pr.partyid = h.customerid" & vbCrLf _
            + " where h.OrderID = " & Val(TxtOrderID.Text) & " and h.OrderDate = '" & DtpOrderDate.DateValue & "'" & " Order By SerialNo"
   
   If ObjRegistry.LaserPrintofSaleInvoice = True Then
      vStrSQL = "Select UserName, h.OrderID as billid, h.OrderDate as BillDate, isnull(h.OrderTime,0) as BillTime, h.Description, h.TotalAmount as tbill, isnull(h.Billdisc,0) as discount, isnull(h.ServiceCharges,0) as ServiceCharges, isnull(h.STax,0) as STax, isnull(h.cashReceived,0) as CashReceived, p.ProductName /*case when isproduct = 1 then p.ProductName else dbo.FunGetProduct(h.OrderID, h.OrderDate) end */ ProductName, unitname, isnull(QtyPack,0) * isnull(Multiplier,0) + Isnull(Bonus,0) + Qty as Qty, b.price/isnull(multiplier,1) as price, b.amount, b.DiscVal, InvoiceNo" & vbCrLf _
            + " , Case when CustomerID = 621 then isnull(CustomerName,AccountName) Else cast(h.CustomerID as varchar(10)) + ' - ' + AccountName End as Customer,  Cash, Credit, BankCard, b.ProductID, PreviousAmount, isnull(OtherCharges,0) as OtherCharges,  h.Empid, e.empname, dbo.FunSaleBodySerial(b.OrderID,b.OrderDate, b.ProductId) Serial, h.TableID, isnull(TableName,'') as TableName, DeliveryDate, isnull(h.DeliveryTime,0) as DeliveryTime, h.InvType, isnull(h.isPrinted,0) as isPrinted," & IIf(ObjRegistry.AllowUrduProduct = False, " isnull(h.Remarks,'')", " isnull(RemarksUrdu,'')") & " as Remarks, " & vbCrLf _
            + " isnull( pr.Phone1  + ', ','') + isnull( pr.Phone2 + ', ','')  + isnull( pr.mobile + ', ','') +  isnull( pr.mobile2 + ', ','') as mobile, pr.Address + pr.city Address " & vbCrLf _
            + " from SaleOrderHeader h inner join SaleOrderBody b on h.OrderID = b.OrderID and h.OrderDate = b.OrderDate" & vbCrLf _
            + " inner join products p on p.productid = b.productid" & vbCrLf _
            + " inner join users ur on ur.UserNo = h.UserNo" & vbCrLf _
            + " inner join ChartofAccounts c on c.AccountNo = h.CustomerID" & vbCrLf _
            + " left outer join parties pr on pr.partyid = h.CustomerID" & vbCrLf _
            + " left outer join Employees e on e.EmpID = h.EmpID" & vbCrLf _
            + " left outer join Tables t on t.TableID = h.TableID " & vbCrLf _
            + " left outer join Units u on u.unitid = p.unitid" & vbCrLf _
            + " left outer join employees emp on emp.empid = h.empid" & vbCrLf _
            + " where h.OrderID = " & Val(TxtOrderID.Text) & " and h.OrderDate ='" & DtpOrderDate.DateValue & "' Order By SerialNo"
   End If
     
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, cn, adOpenDynamic, adLockReadOnly
 
   If RsReport.RecordCount = 0 Then Exit Sub
   
   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
    
   If ObjRegistry.LaserPrintofSaleInvoice = True Then
      Set RptReportViewer.Report = New CrpSaleOrderHalf
      RptReportViewer.Report.PaperSize = crPaperA4
      RptReportViewer.Report.PaperOrientation = crLandscape
      RptReportViewer.Report.TopMargin = vY
      RptReportViewer.Report.LeftMargin = vX
      RptReportViewer.Report.RightMargin = 225
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
      ElseIf InStr(1, Printer.DeviceName, "AB-80K") > 0 Or InStr(1, Printer.DeviceName, "ARP-808K") Then
         Set RptReportViewer.Report = New CrpSaleInvoiceAurora
         RptReportViewer.Report.LeftMargin = 225
         RptReportViewer.Report.RightMargin = 0
         RptReportViewer.Report.TopMargin = 255
      ElseIf InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
         Set RptReportViewer.Report = New CrpSaleInvoice
         RptReportViewer.Report.TopMargin = 225
         RptReportViewer.Report.LeftMargin = 225
         RptReportViewer.Report.RightMargin = 225

      Else 'InStr(1, Printer.DeviceName, "AB-80K") > 0 Then
         Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpSaleOrderAurora.rpt")
'         Set RptReportViewer.Report = New CrpSaleInvoiceAurora
'         RptReportViewer.Report.TopMargin = vY
'         RptReportViewer.Report.LeftMargin = vX
        RptReportViewer.Report.TopMargin = 225
         RptReportViewer.Report.LeftMargin = 225
         RptReportViewer.Report.RightMargin = 0
      End If
      'RptReportViewer.Report.PaperOrientation = crPortrait
    End If
    RptReportViewer.Report.DiscardSavedData
    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
    RptReportViewer.Report.ReportTitle = "Sale Order"
    
   If vPrintHeader = True Then
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   Else
      RptReportViewer.Report.ParameterFields(1).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(4).AddCurrentValue ""
   End If
   
   'RptReportViewer.Report.ParameterFields(3).AddCurrentValue EncryptStr(CN.Execute("Select dbo.Value('1')").Fields(0), False) 'CN.Execute("Select Name from Manufacturer").Fields(0).Value

   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.DevelopedBy  'CN.Execute("Select Name from Manufacturer").Fields(0).Value
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.AddSpace = True, ".", "")
   RptReportViewer.Report.ParameterFields(6).AddCurrentValue CBool(ObjRegistry.CashReceived)
   RptReportViewer.Report.ParameterFields(7).AddCurrentValue CStr(ObjRegistry.Statement)
   
   If ObjRegistry.LaserPrintofSaleInvoice = True Then
      RptReportViewer.Report.ParameterFields(8).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(9).AddCurrentValue ObjRegistry.PreviousBalanceVisible
   Else
      RptReportViewer.Report.ParameterFields(8).AddCurrentValue ""
      RptReportViewer.Report.ParameterFields(9).AddCurrentValue IIf(ObjRegistry.PreviousBalanceVisible = True, FrmOrderPrint.ParaOutPrevious, 0)
   End If
   'RptReportViewer.Report.SelectPrinter "RASDD.DLL", "CBM1000 Partial Cut", "Com1" 'RptReportViewer.Report.SelectPrinter  "RASDD.DLL", "CBM1000 Partial Cut", "Com1"
   'RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
   'RptReportViewer.Show
   vPrinter = Split(CmbPrinters.Text, ",")
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)

   If ChkIsPreview.Value = 1 Then
      RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False, CInt(vNoofPrints)
   End If
'   RptReportViewer.Report.PrintOut False ', CInt(vNoofPrints)
   
   cn.Execute "update SaleOrderHeader set isPrinted = 1 where isnull(isPrinted,0) = 0 and OrderID = " & Val(TxtOrderID.Text) & " and OrderDate ='" & DtpOrderDate.DateValue & "'"
   cn.Execute "update SaleOrderBody set isPrinted = 1 where isnull(isPrinted,0) = 0 and OrderID = " & Val(TxtOrderID.Text) & " and OrderDate ='" & DtpOrderDate.DateValue & "'"
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Sale Invoice'" & "," & TxtBillID.Text & ",'" & DtpBillDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProduct_Click()
   On Error GoTo ErrorHandler
   If FunSelectProduct(ssButton, True) = True Then
      TxtQty.SetFocus
   Else
      TxtCode.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubAddTempSaleOrder()
 Dim vBm As Variant
   Dim i As Integer
   Grid.Redraw = False
   vBm = Grid.Bookmark
   TxtTotalAmount.Caption = "0"
   Grid.MoveFirst
   cn.Execute "Delete from TempSaleOrderBody"
   '/********* Watch *************/
'   cn.Execute "Insert into Watch(ErrorFrom,Narration) values ('Save Before For Grid Loop','OrderID = " & TxtOrderID.Text & ", OrderDate = " & DtpOrderDate.DateValue & "')"
   For i = 0 To Grid.rows - 1
      TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(Grid.Columns("TotalAmount").CellValue(Grid.GetBookmark(i)))
      If vPrintKitchenInoices = True Then
         If Trim(Grid.Columns("ProductID").CellValue(Grid.GetBookmark(i))) <> "" Then
            With cn.Execute("Select * from SaleOrderBody where OrderID = " & TxtOrderID.Text & " and OrderDate = '" & DtpOrderDate.DateValue & "' and ProductID = " & Val(Grid.Columns("ProductID").CellValue(Grid.GetBookmark(i))))
               If .RecordCount = 0 Then
                  vStrSQL = "insert into TempSaleOrderBody (OrderID, OrderDate, ProductID, Qty) values (" & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "'," & Val(Grid.Columns("ProductID").CellValue(Grid.GetBookmark(i))) & "'," & Grid.Columns("Qty").CellValue(Grid.GetBookmark(i)) & ")"
                  cn.Execute vStrSQL
               Else
                  RsBody.Filter = " ProductID = " & Val(Grid.Columns("ProductID").CellValue(Grid.GetBookmark(i)))
                  If Grid.Columns("Qty").CellValue(Grid.GetBookmark(i)) > !Qty Then
                     vStrSQL = "Insert into TempSaleOrderBody (OrderID, OrderDate, ProductID, Qty) values (" & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "'," & Val(Grid.Columns("ProductID").CellValue(Grid.GetBookmark(i))) & "," & Grid.Columns("Qty").CellValue(Grid.GetBookmark(i)) - !Qty & ")"
                     cn.Execute vStrSQL
                  End If
               End If
            End With
         End If
      End If
   Next i
   Grid.Bookmark = vBm
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
   
Private Sub BtnSaleInvoice_Click()
On Error GoTo ErrorHandler
   SchSale.ParaInBillDate = DtpBillDate.DateValue
   SchSale.Show vbModal
   If SchSale.ParaOutBillID <> -1 Then
      TxtBillID.Text = SchSale.ParaOutBillID
      DtpBillDate.DateValue = SchSale.ParaOutBillDate
      GetSaleInvoice
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSaleOrderPOS", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If ObjRegistry.TableVisible = True And Val(TxtTableID.Text) = 0 And ObjRegistry.TableIDMandatory = True Then
      MsgBox "Please enter Table ID ", vbExclamation, "Alert"
      Exit Sub
   End If
   If ObjRegistry.InvType = True And CmbType.Text = "" Then
      MsgBox "Please enter Invoice Type", vbExclamation, "Alert"
      CmbType.SetFocus
      Exit Sub
   End If
   
'   If vIsNewRecord = False And CmbType.Text <> "" Then
'      If vIsEdit = False Then
'         MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
'         Exit Sub
'      End If
'   ElseIf vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False Then
'      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
   
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False Then
'      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
   
   
'  Header Validation
   '''''''''''''''''''''''Check Entry Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.isEntryDate = True Then
       If ObjRegistry.FromDate > Date Or ObjRegistry.ToDate < Date Then
         MsgBox "Data can not be saved Because Date is not set according to the Software's Entry date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
    
    '''''''''''''''''''''''Check Current Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.CurrentDateDataEntry = True Then
       If DtpOrderDate.DateValue <> Date Then
         MsgBox "Data can not be saved because date is not current date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
   
   Call SubAddTempSaleOrder

   Call SubCalculateFooter
   
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      TxtStoreID.SetFocus
      Exit Sub
   End If

'   If DtpOrderDate.Enabled = True Then
'      If FrmOrderPrint.OptCash.Visible Then FrmOrderPrint.OptCash.SetFocus
'      FrmOrderPrint.SubClearFields
'   End If
   FrmOrderPrint.ParaInPrint = vAutoPrintSaleOrder
   FrmOrderPrint.ParaInKitchenPrint = vPrintKitchenInoices
       
   
   
   
   FrmOrderPrint.ParaInChoice = IIf(vIsCreditSale = True, "Credit", "Cash")
   
   FrmOrderPrint.ParaInID = Val(TxtOrderID.Text)
   FrmOrderPrint.ParaInDate = DtpOrderDate.DateValue
   
   FrmOrderPrint.TxtNetAmount.Text = TxtNetAmount.Caption
   
'   If ObjRegistry.CashReceived = False Then
'      FrmOrderPrint.TxtCashReceivedCash.Text = TxtNetAmount.Caption
'   End If
   
'   If vPrintKitchenInoices = True Then Call PrintDepartment
'   If vAutoPrintSaleOrder Then Call BtnPrint_Click
   
   
   FrmOrderPrint.Show vbModal, Me
   If FrmOrderPrint.ParaOutSelection = False Then Exit Sub
   
   
'   If DtpOrderDate.Enabled And DtpOrderDate.Date <> IIf(Format(Now, "hh") > IIf(IsNull(!HourDifference), 0, !HourDifference), Date, DateAdd("d", -1, Date)) And DateFlag = True Then
'      If MsgBox("Are you sure to Change Bill Date into Current Date", vbInformation + vbYesNo, "Alert") = vbYes Then
'         DtpOrderDate.DateValue = IIf(Format(Now, "hh") > IIf(IsNull(!HourDifference), 0, !HourDifference), Date, DateAdd("d", -1, Date))
'         TxtOrderID.Text = FunGetMaxID()
'      End If
'      DateFlag = False
'   End If
   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to sale", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
   End If
  'Body Validation
  ' validation has been performed when a row is added to the grid

  'Saving record
   cn.BeginTrans
    
   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmSaleOrderPOS, eEdit, TxtOrderID.Text, DtpOrderDate.DateValue, "Amount: " & Val(TxtNetAmount.Caption))
    
   '/********* Watch *************/
'   cn.Execute "Insert into Watch(ErrorFrom,Narration) values ('Save Before NewRecord is " & vIsNewRecord & "','OrderID = " & TxtOrderID.Text & ", OrderDate = " & DtpOrderDate.DateValue & "')"
   If vIsNewRecord = True Then
      If cn.Execute("Select * from SaleOrderHeader where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "' and StampID <> " & TxtStampID.Text).RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtOrderID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
   '/********* Watch *************/
'   cn.Execute "Insert into Watch(ErrorFrom,Narration) values ('Save After NewRecord is " & vIsNewRecord & "','OrderID = " & TxtOrderID.Text & ", OrderDate = " & DtpOrderDate.DateValue & "')"
    
   If vIsNewRecord = False Then Call ActivityLog("Sale Order", eEdit, TxtOrderID.Text, DtpOrderDate.DateValue)
   ''''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Call UserActivities
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
    ssql = "select * from SaleOrderHeader where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'"
    Dim Rs As New ADODB.Recordset
    With Rs
      .Open ssql, cn, adOpenDynamic, adLockPessimistic
      If .BOF Then
         .AddNew
         !OrderID = Val(TxtOrderID.Text)
         !OrderDate = DtpOrderDate.DateValue
         !OrderTime = Now
         !StampID = TxtStampID.Text
         !UserNo = vUser
         !BillID = IIf(Val(TxtBillID.Text) = 0, Null, TxtBillID.Text)
         !BillDate = DtpBillDate.DateValue
      End If
      !isReplace = 0
      !isPosted = 0
      !StoreID = TxtStoreID.Text
      !PromiseDate = IIf(DtpPromiseDate.DateValue = Empty, Null, DtpPromiseDate.DateValue)
      !InvType = IIf(CmbType.Visible = False, Null, CmbType.Text)
      !TableId = IIf(Trim(TxtTableID.Text) = "", Null, TxtTableID.Text)
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
      !EmpComm = IIf(Trim(TxtEmployeeID.Text) = "", Null, Val(TxtCommission.Text))
      !MemberID = IIf(Trim(TxtMemberID.Text) = "", Null, TxtMemberID.Text)
      !TotalAmount = SelfRound(vTotalAmount)
      !BillDisc = IIf(TxtBillDisc.Text = "", Null, Val(TxtBillDisc.Text))
      !BillDiscPer = IIf(TxtBillDiscPer.Text = "", Null, Val(TxtBillDiscPer.Text))
      !ServiceCharges = IIf(TxtServiceCharges.Text = "", Null, Val(TxtServiceCharges.Text))
      !ServiceChargesPer = IIf(TxtServiceChargesPer.Text = "", Null, Val(TxtServiceChargesPer.Text))
      !STax = IIf(TxtSTax.Text = "", Null, Val(TxtSTax.Text))
      !STaxPer = IIf(TxtSTaxPer.Text = "", Null, Val(TxtSTaxPer.Text))
      !DeliveryDate = DtpDeliveryDate.DateValue
      !DeliveryTime = DTPDeliveryTime.Value
      If FrmOrderPrint.OptCash.Value = True Then
         !Commision = Null
         !InvoiceNo = Null
         !BankMachineID = Null
         !CashReceived = Val(FrmOrderPrint.TxtCashReceivedCash.Text)
         !CustomerID = "621"
         !CustomerName = IIf(Trim(FrmOrderPrint.TxtCashCustomer.Text) = "", Null, FrmOrderPrint.TxtCashCustomer.Text)
      End If
      If FrmOrderPrint.OptCredit.Value = True Then
         !Commision = Null
         !InvoiceNo = Null
         !BankMachineID = Null
         !CashReceived = Val(FrmOrderPrint.TxtCashReceivedCredit.Text)
         !CustomerID = FrmOrderPrint.TxtCustomerID.Text
         !CustomerName = Null
      End If
      !BankCard = FrmOrderPrint.OptBankCard.Value
      !Cash = FrmOrderPrint.OptCash.Value
      !Credit = FrmOrderPrint.OptCredit.Value
      !Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
      !Remarks = IIf(Trim(TxtRemarks.Text) = "", Null, TxtRemarks.Text)
      !RemarksUrdu = IIf(Trim(TxtRemarksUrdu.Text) = "", Null, TxtRemarksUrdu.Text)
      !ManualBillNo = IIf(Trim(TxtManualBillNo.Text) = "", "", TxtManualBillNo.Text)
      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
      .Update
      .Close
   End With
   With RsBody
      .Filter = 0 '"StampID = " & TxtStampID.Text
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !OrderID = Val(TxtOrderID.Text)
         !OrderDate = DtpOrderDate.DateValue
         !StampID = TxtStampID.Text
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
'   With RsDetail
'      .Filter = 0
'      If .RecordCount > 0 Then .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !OrderID = Val(TxtOrderID.Text)
'         !OrderDate = DtpOrderDate.DateValue
'         .MoveNext
'      Next vCounter
'      .UpdateBatch
'   End With

   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmSaleOrderPOS, eAdd, TxtOrderID.Text, DtpOrderDate.DateValue, Grid.rows - 1 & " New Product/s Added Amount: " & Val(TxtNetAmount.Caption))
'   If vIsNewRecord = True Then Call ActivityLog("Sale Order", eAdd, TxtOrderID.Text, DtpOrderDate.DateValue)
   cn.CommitTrans
 '  Char.Speak "Thank you for comming"
   'If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
   vWhere = "" '" and isnull(isPrinted,0) = 0"
   
   '/********* Watch *************/
'   cn.Execute "Insert into Watch(ErrorFrom,Narration) values ('Save Before PrintKitchenInoices is " & vPrintKitchenInoices & "','OrderID = " & TxtOrderID.Text & ", OrderDate = " & DtpOrderDate.DateValue & "')"
   vSave = True
   If FrmOrderPrint.ChkKitchenPrint.Value = 1 Then Call PrintDepartment
   If FrmOrderPrint.ChkPrint.Value = 1 Then Call BtnPrint_Click

'   If vCashDrawer = True Then
'      'Shell "mode com1 9600,n,8,1", vbNormalFocus
'      'Shell "echo ^G>com1", vbNormalFocus
'      MSComm1.Output = "O"
'   End If

   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'Sale Order POS' and LocalComputerName = '" & LocalComputerName & "'"
   If cn.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview ) Values ('" & LocalComputerName & "', 'Sale Order POS','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & " Where FormType = 'Sale Order POS' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   cn.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   
   FrmOrderPrint.SubClearFields
   FormStatus = NewMode
   'End If
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleOrderBody where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "' and StampID = " & TxtStampID.Text, cn, adOpenDynamic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.ProductName, b.code, b.* from SaleOrderBody b join products p on p.productid = b.productid where OrderID=" & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "' order by serialno"
      With cn.Execute(ssql)
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
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !Qty
            Grid.Columns("QtyOrigional").Value = !Qty
            Grid.Columns("Price").Value = !Price
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("IsProduct").Value = Abs(!isProduct)
            Grid.Columns("TotalAmount").Value = Val(!Price) * Val(!Qty)
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            Grid.Columns("EmpComm").Value = IIf(IsNull(!EmpComm), "", !EmpComm)
            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
            'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(!DiscVal)
            vTotDisc = vTotDisc + Val(!DiscVal)
            vTotalAmount = vTotalAmount + !Amount
            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
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
'   RsDetail.Filter = 0
'   If RsDetail.State = adStateOpen Then RsDetail.Close
'   RsDetail.Open "Select * from SaleOrderUnionUsed where OrderID=" & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
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
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      vServerDate = cn.Execute("Select CONVERT(datetime, CONVERT(varchar, GETDATE(), 110)) ServerDate").Fields(0).Value
      vSystemDate = Abs(ObjRegistry.SystemDate)
      vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
   
      vDate = IIf(vSystemDate = True, cn.Execute("Select SystemDate From SystemDate").Fields(0).Value, vServerDate)

      
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
         If Format(cn.Execute("Select getdate()").Fields(0).Value, "hh") >= vHDiff Then
            vDate = vDate
         Else
            vDate = DateAdd("d", -1, vDate)
         End If
      End If
      
    
      
      DtpBillDate.DateValue = vDate
      
      
      DtpOrderDate.DateValue = DtpBillDate.DateValue

      
      DtpOrderDate.DateValue = vDate
      DtpDeliveryDate.DateValue = vDate
'      DTPDeliveryTime.Value = "00:00"
      TxtOrderID.Text = FunGetMaxID()
      TxtStampID.Text = StampID()
      Call PopulateDataToGrid
      'TxtCustomerID.Text = "621"
      'TxtCustomerName.Text = "Counter Sale"
'      DtpOrderDate.DateValue = vDate
      LblStock.Visible = False
      LblStockCaption.Visible = False
      LblCaptionPrice.Visible = False
      LblPrice.Visible = False
      TxtCode.Enabled = True
      TxtProductName.Enabled = False
      BtnProduct.Enabled = True
      vSave = False
'      DtpOrderDate.Enabled = True
      'If DtpOrderDate.Enabled And DtpOrderDate.Visible Then DtpOrderDate.SetFocus
      TxtCode.Enabled = True
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      DtpOrderDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      TxtCode.Enabled = True
      TxtProductName.Enabled = False
      BtnProduct.Enabled = True
      TxtCode.SetFocus
      LblStock.Visible = False
      LblStockCaption.Visible = False
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

Private Sub BtnTable_Click()
   On Error GoTo ErrorHandler
   If FunSelectTable(ssButton, False) = True Then
      BtnSave.SetFocus
   Else
      TxtTableID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbType_Click()
   On Error GoTo ErrorHandler
   With cn.Execute("Select ServiceChargesPer, IsEdit from InvTypes where InvType = '" & CmbType.Text & "'")
      If .RecordCount > 0 Then
         vIsEdit = !IsEdit
         TxtServiceChargesPer.Enabled = False
         TxtServiceCharges.Enabled = False
         TxtServiceCharges.Tag = "NC"
         TxtServiceChargesPer.Tag = "NC"
         TxtServiceChargesPer.Text = !ServiceChargesPer
      Else
         TxtServiceChargesPer.Enabled = True
         TxtServiceCharges.Enabled = True
         TxtServiceCharges.Tag = ""
         TxtServiceChargesPer.Tag = ""
         TxtServiceChargesPer.Text = ""
      End If
      .Close
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpOrderDate_LostFocus()
   On Error GoTo ErrorHandler
'   TxtOrderID.Text = FunGetMaxID()
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Activate()
'   On Error GoTo ErrorHandler
'   If vFlag = True Then
'      If TxtCode.Enabled = True And TxtCode.Visible = True Then TxtCode.SetFocus
'   Else
'      vFlag = True
'   End If
'   Exit Sub
'ErrorHandler:
'   If Err.Number = 5 Then Resume Next
'   Call ShowErrorMessage

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name = TxtRemarksUrdu.Name Then
      Call Textbox1_KeyDown(KeyCode, Shift)
      Exit Sub
   End If
   
   If KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      Select Case ActiveControl.Name
         Case TxtCode.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name
         If TxtCode.Enabled Then TxtCode.SetFocus: Call SubClearDetailArea
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
            If BtnSave.Enabled And BtnSave.Visible Then BtnSave_Click
            KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled And BtnClear.Visible Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled And BtnClose.Visible Then BtnClose_Click
            KeyCode = 0
'         Case vbKeyU
'            Call SubMakePackageDeal
         Case vbKeyM
               If TxtMemberID.Visible = True And TxtMemberID.Enabled = True Then TxtMemberID.SetFocus
               KeyCode = 0
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
               KeyCode = 0
         Case vbKeyO
            If BtnOpen.Enabled And BtnOpen.Visible Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled And BtnDelete.Visible Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled And BtnPrint.Visible Then BtnPrint_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyC And Shift = vbAltMask Then
      FrmOrderPrint.ParaInChoice = "Credit"
      FrmOrderPrint.Show vbModal, Me
   ElseIf KeyCode = vbKeyReturn And Shift = vbShiftMask Then
      Select Case ActiveControl.Name
      Case TxtCode.Name
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
      Case TxtCode.Name
         If Trim(TxtCode.Text) = "" Then
            If TxtTableID.Visible = True Then
               TxtTableID.SetFocus
            Else
               If BtnSave.Enabled And BtnSave.Visible Then BtnSave.SetFocus
            End If
         End If
         If FunSelectProduct(ssValidate, False) = True Then If vAutoEnterBeforeQty = True Then GetDataFromTexBoxesToGrid Else keybd_event 9, 1, 1, 1: KeyCode = 0
      Case TxtQty.Name, TxtDiscPC.Name, TxtDiscPer.Name, TxtDiscVal.Name, TxtPrice.Name
         GetDataFromTexBoxesToGrid
      Case Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtEmployeeID.SetFocus
         Case TxtTableID.Name: If FunSelectTable(ssFunctionKey, False) = True Then TxtTableID.SetFocus
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then If TxtMemberID.Visible = True Then If TxtCode.Enabled Then TxtCode.SetFocus
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtQty.SetFocus
         Case TxtMemberID.Name: If FunSelectMember(ssFunctionKey, True) = True Then If TxtCode.Enabled Then TxtCode.SetFocus
      End Select
   ElseIf KeyCode = vbKeyF3 Then
      TxtProductName.Enabled = True
      If TxtProductName.Enabled = True And TxtProductName.Visible = True Then TxtProductName.SetFocus
         'Call FindRow
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         If Grid.Visible And Grid.Enabled Then Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
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
         Case TxtCode.Name, TxtQty.Name, TxtPrice.Name, TxtDiscPC.Name, Grid.Name
            LblCost.Caption = cn.Execute("select dbo.FunPurPrice(" & Val(TxtPID.Text) & ")").Fields(0).Value
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
   If Me.ActiveControl.Name = TxtRemarksUrdu.Name Then
      Call Textbox1_KeyPress(KeyAscii)
      Exit Sub
   End If
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

   Call InvoiceNo
   SetWindowText Me.hWnd, "Sale Order (" & LblNo & ")"
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
   ssql = "select * from FormDefaultSetting Where FormType = 'Sale Order POS' and LocalComputerName = '" & LocalComputerName & "'"
   With cn.Execute(ssql)
     If .RecordCount > 0 Then
        cmbPrintType.Text = !Size
        ChkIsPreview.Value = Abs(!IsPreview)
        If Not IsNull(!DeviceName) Then
            CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
        Else
            CmbPrinters.ListIndex = 0
        End If
     End If
     .Close
   End With
   ''''''''''''''''''''''''''''''''''''''''''''''
   vSystemDate = Abs(ObjRegistry.SystemDate)
   vHDiff = IIf(IsNull(ObjRegistry.HourDifference), 0, ObjRegistry.HourDifference)
   
   DtpOrderDate.Enabled = False
   
   vAllowNegativeOrder = ObjRegistry.AllowNegativeOrder
   
   If ObjRegistry.ShowPromiseDateInSalaPurchase = True Then
      LblPromiseDate.Visible = True
      DtpPromiseDate.Visible = True
      DtpPromiseDate.DateValue = Null
   Else
      LblPromiseDate.Visible = False
      DtpPromiseDate.Visible = False
      DtpPromiseDate.DateValue = Null
   End If
   
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
                  
   LblTableID.Visible = ObjRegistry.TableVisible
   LblTableName.Visible = ObjRegistry.TableVisible
   TxtTableID.Visible = ObjRegistry.TableVisible
   TxtTableName.Visible = ObjRegistry.TableVisible
   BtnTable.Visible = ObjRegistry.TableVisible
         
   TxtManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   LblManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   
   TxtRemarks.Visible = ObjRegistry.RemarksVisible
   LblRemarks.Visible = ObjRegistry.RemarksVisible
   
   If LblRemarks.Visible = True Then
      TxtRemarks.Visible = Not ObjRegistry.AllowUrduProduct
      TxtRemarksUrdu.Visible = ObjRegistry.AllowUrduProduct
   End If
   
   If ObjUserSecurity.IsAdministrator = False Then
      TxtDiscPC.Enabled = ObjRegistry.DiscAllowed
      TxtDiscPer.Enabled = ObjRegistry.DiscAllowed
      TxtDiscVal.Enabled = ObjRegistry.DiscAllowed
      TxtBillDisc.Enabled = ObjRegistry.DiscAllowed
      TxtBillDiscPer.Enabled = ObjRegistry.DiscAllowed
   End If
   
   vIsDisableCreditSale = ObjUserSecurity.IsDisableCreditSale
   vIsCreditSale = ObjUserSecurity.IsCreditSale
   
   LblType.Visible = ObjRegistry.InvType
   CmbType.Visible = ObjRegistry.InvType
   
   CmbType.Clear
   CmbType.AddItem ""
   With cn.Execute("select * from InvTypes")
      If .RecordCount > 0 Then
         While Not .EOF
            CmbType.AddItem ![InvType]
            .MoveNext
         Wend
      End If
   End With

   vAutoEnterBeforeQty = ObjRegistry.AutoEnterBeforeQty
   vAutoPrintSaleOrder = Abs(ObjRegistry.AutoPrintSaleOrder)
   vPrintKitchenInoices = Abs(ObjRegistry.PrintKitchenInoices)
   
   'vCashDrawer = !CashDrawer
   vX = IIf(IsNull(ObjRegistry.x), 0, Val(ObjRegistry.x))
   vY = IIf(IsNull(ObjRegistry.Y), 0, Val(ObjRegistry.Y))
   vLaserInvoice = ObjRegistry.LaserPrintofSaleInvoice
   vPrintHeader = ObjRegistry.PrintHeadersSaleInvoice
   vNoofPrints = IIf(IsNull(ObjRegistry.NoofPrints) Or ObjRegistry.NoofPrints = 0, 1, ObjRegistry.NoofPrints)
   MniCostPrice.Visible = ObjRegistry.CostVisible
   
   If ObjUserSecurity.IsAdministrator = True Then
      TxtPrice.Enabled = True
   Else
      TxtPrice.Enabled = ObjUserSecurity.IsChangeRetail
   End If
   
   With cn.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         If !ChangePrice = True Then TxtPrice.Enabled = True
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
         vNoofPrints = IIf(IsNull(!NoofPrints) Or !NoofPrints = 0, 1, !NoofPrints)
      End If
      .Close
   End With
   
   DateFlag = True
   FormStatus = NewMode
   
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   vIsEdit = True
   'DtpOrderDate.DateValue = IIf(Format(Now, "hh") > IIf(IsNull(!HourDifference), 0, !HourDifference), Date, DateAdd("d", -1, Date))
   'If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
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
   With cn.Execute("Select * from TempNo where UserNo = " & vUser & " order by tempno")
      While (Not .EOF) And LoopFlag = True
         If vC <> !TempNo And Not .EOF Then
            LoopFlag = False
         Else
            vC = vC + 1
         End If
         .MoveNext
      Wend
'      LblNo.Caption = " Order. Open # " & CStr(vC)
      cn.Execute "INSERT INTO TempNo(TempNo,UserNo) VALUES (" & vC & "," & vUser & ")"
      .Close
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function StampID() As Long
   On Error GoTo ErrorHandler
   StampID = cn.Execute("Select isnull(max(SOID),0)+1 from Stamp").Fields(0)
   cn.Execute "update Stamp set SOID = " & StampID
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpOrderDate.IsDateValid = False Then FunGetMaxID = 1: Exit Function
   FunGetMaxID = cn.Execute("Select isnull(max(OrderID),0)+1 from SaleOrderHeader where OrderDate = '" & DtpOrderDate.DateValue & "'").Fields(0)
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
   DtpPromiseDate.DateValue = Null
   TxtRemarksUrdu.Text = ""
   TxtTotalQty.Caption = 0
   TxtTotalDiscount.Caption = 0
   TxtTotalAmount.Caption = 0
   TxtNetAmount.Caption = 0
   vTotDisc = 0
   vTotalAmount = 0
   vWhere = ""
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ProductID").Text = " "
   Grid.Update
   Unload FrmOrderPrint
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
   cn.Execute "delete from tempno where tempno = " & Val(Right(LblNo.Caption, 1))
    'CN.Execute ("exec spcurrentstock")
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set RsBody = Nothing
    Set RsReport = Nothing
    Set FrmSaleOrder = Nothing
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
               ssql = "Select Productid From SaleOrderbody where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
               With cn.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmSaleOrderPOS, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Closed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmSaleOrderPOS, eCloseSavedRecord, TxtOrderID.Text, DtpOrderDate.DateValue, vGridRows & " Product/s Closed")
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
   If Flag Then Call GetDataBackFromGridToTexBoxes
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
   LblCost.Visible = False
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      TxtCode.Text = ""
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      TxtCode.Enabled = False
      BtnProduct.Enabled = False
      If TxtQty.Enabled = True And TxtQty.Visible Then TxtQty.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
   If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
      
   ssql = "Select Productid From SaleOrderbody where Orderid=" & Val(TxtOrderID.Text) & " and Orderdate ='" & DtpOrderDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
   With cn.Execute(ssql)
      If .EOF Then
         Call ActivityLogBin("", eFrmSaleOrderPOS, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      Else
         Call ActivityLogBin("", eFrmSaleOrderPOS, eRemoveRow, TxtOrderID.Text, DtpOrderDate.DateValue, "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
         Call ActivityLogBin(vRandomID, eFrmSaleOrderPOS, eAddTempRecord, TxtOrderID.Text, DtpOrderDate.DateValue, "Pending Remove Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
      End If
   End With
   
   If ObjRegistry.TableVisible = True Then
      If vIsNewRecord = False And UserSecurity.IsAdministrator = False Then
      If cn.Execute("Select isPrinted from SaleOrderHeader where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "' and IsPrinted = 1 ").RecordCount > 0 Then
         MsgBox "Data cannot be deleted Because order has printed", vbInformation, "Alert"
         Exit Sub
      End If
   End If
   End If
   RsBody.Filter = "Code='" & TxtCode.Text & "'"
   If RsBody.RecordCount > 0 Then RsBody.Delete
   ''''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Removed ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
   If ObjRegistry.TableVisible = True And vIsNewRecord = False And TxtCode.Enabled = False And Val(TxtQty.Text) < Grid.Columns("Qty").Value Then
      If cn.Execute("Select isPrinted from SaleOrderHeader where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "' and IsPrinted = 1 ").RecordCount > 0 Then
         MsgBox "Data cannot be changed Because order has printed with higher qty", vbInformation, "Alert"
         Exit Sub
      End If
   End If
   If vAllowNegativeOrder = False Then
      If vIsNewRecord = True Then
         If (Val(vQtyLoose) - Val(TxtQty.Text)) < 0 Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            Grid.Redraw = True
            Call SubClearDetailArea
            If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
            Exit Sub
         End If
      Else
         If (Val(vQtyLoose) - Val(TxtQty.Text) + Val(Grid.Columns("QtyOrigional").Value)) < 0 Then
            MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
            Grid.Redraw = True
            Call SubClearDetailArea
            If TxtCode.Enabled And TxtCode.Visible Then TxtCode.SetFocus
            Exit Sub
         End If
      End If
   End If
   RsBody.Filter = "ProductID = " & Val(TxtPID.Text)
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
'         If Trim(TxtQty.Text) > Val(LblStock.Caption) Then
'            MsgBox "Insufficent Stock.", vbExclamation, "Alert"
'            TxtQty.SetFocus
'            Exit Sub
'         End If
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
                  ssql = "Select Productid From SaleOrderbody where orderid=" & Val(TxtOrderID.Text) & " and Orderdate ='" & DtpOrderDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
                  With cn.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmSaleOrderPOS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmSaleOrderPOS, eEdit, TxtOrderID.Text, DtpOrderDate.DateValue, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
'                  With CN.Execute("select * from registry")
'                     If .RecordCount > 0 Then
'                        If !AllowNegativeOrder = False Then
'                           If vIsNewRecord = True Then
'                              If (Val(vQtyLoose) - Val(TxtQty.Text) - Val(Grid.Columns("Qty").Value)) < 0 Then
'                                 MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                                 Grid.MoveLast
'                                 Grid.Redraw = True
'                                 Exit Sub
'                              End If
'                           Else
'                              If (Val(vQtyLoose) - Val(TxtQty.Text) - Val(Grid.Columns("Qty").Value) + Val(Grid.Columns("QtyOrigional").Value)) < 0 Then
'                                 MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                                 Grid.MoveLast
'                                 Grid.Redraw = True
'                                 Exit Sub
'                              End If
'                           End If
'                        End If
'                     End If
'                     .Close
'                  End With
                  
                  TxtPrice.Text = Round((Val(TxtActualAmount.Text) + Val(Grid.Columns("TotalAmount").Text)) / (Val(TxtQty.Text) + Grid.Columns("Qty").Value), 2)
                  TxtQty.Text = Val(TxtQty.Text) + Grid.Columns("Qty").Value
                  TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text) - Val(Grid.Columns("Qty").Text)
                  vTotDisc = vTotDisc + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
                  vTotalAmount = vTotalAmount + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  
                  'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(TxtDiscVal.Text) - Val(Grid.Columns("DiscVal").Text)
                  TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text) - Val(Grid.Columns("TotalAmount").Text)
                  TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("Qty").Value = Val(TxtQty.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("DiscPC").Value = Val(TxtDiscPC.Text)
                  Grid.Columns("DiscPer").Value = Val(TxtDiscPer.Text)
                  Grid.Columns("DiscVal").Value = Val(TxtDiscVal.Text)
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("Cost").Value = Val(TxtCost.Text)
                  Grid.Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
                  Grid.Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
                  Grid.Columns("EmpComm").Value = IIf(Val(TxtEmpComm.Text) = 0, 0, Val(TxtEmpComm.Text))
                  RsBody!Qty = Val(TxtQty.Text)
                  RsBody!Price = Val(TxtPrice.Text)
                  RsBody!DiscPC = Val(TxtDiscPC.Text)
                  RsBody!DiscPer = Val(TxtDiscPer.Text)
                  RsBody!DiscVal = Val(TxtDiscVal.Text)
                  RsBody!Cost = Val(TxtCost.Text)
                  RsBody!isProduct = Abs(ChkIsProduct.Value)
                  RsBody!Amount = Val(TxtAmount.Text)
                  RsBody!EmpComm = Val(TxtEmpComm.Text)
                  With cn.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmSaleOrderPOS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmSaleOrderPOS, eEdit, TxtOrderID.Text, DtpOrderDate.DateValue, "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmSaleOrderPOS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Pending Update Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
                  
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
   'Grid.Redraw = False
   With Grid
'      With CN.Execute("select * from registry")
'         If .RecordCount > 0 Then
'            If !AllowNegativeOrder = False Then
'               If vIsNewRecord = True Then
'                  If (Val(vQtyLoose) - Val(TxtQty.Text)) < 0 Then
'                     MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                     Grid.Redraw = True
'                     Exit Sub
'                  End If
'               Else
'                  If (Val(vQtyLoose) - Val(TxtQty.Text) + Val(Grid.Columns("QtyOrigional").Value)) < 0 Then
'                     MsgBox "Insufficient Stock for this Product", vbInformation + vbOKOnly, "Error"
'                     Grid.Redraw = True
'                     Exit Sub
'                  End If
'               End If
'            End If
'         End If
'         .Close
'      End With
      If TxtCode.Enabled = True Then
         TxtNetAmount.Caption = Val(TxtNetAmount.Caption) + Val(TxtAmount.Text)
         TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(TxtQty.Text)
         'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(TxtDiscVal.Text)
         vTotDisc = vTotDisc + Val(TxtDiscVal.Text)
         vTotalAmount = vTotalAmount + Val(TxtAmount.Text)
         TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Val(TxtActualAmount.Text)
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmSaleOrderPOS, eAddNewRowByEdit, TxtOrderID.Text, DtpOrderDate.DateValue, "Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQty.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
         Call ActivityLogBin(vRandomID, eFrmSaleOrderPOS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Pending Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtQty.Text) & " Price-" & TxtPrice.Text & " Disc-" & TxtDiscPer.Text & " Amount-" & TxtAmount.Text)
      Else
         ssql = "Select Productid From SaleOrderbody where Orderid=" & Val(TxtOrderID.Text) & " and Orderdate ='" & DtpOrderDate.DateValue & "' and productid = " & Val(Grid.Columns("Code").Text)
         With cn.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmSaleOrderPOS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmSaleOrderPOS, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQty.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            Else
               Call ActivityLogBin("", eFrmSaleOrderPOS, eEdit, TxtOrderID.Text, DtpOrderDate.Date, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("Qty").Text) & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmSaleOrderPOS, eEdit, TxtOrderID.Text, DtpOrderDate.Date, "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtQty.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmSaleOrderPOS, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtOrderID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpOrderDate.Date), "Pending Update Code-" & TxtCode.Text & " Qty-" & Val(TxtQty.Text) & " Price-" & TxtPrice.Text & " Disc-" & Val(TxtDiscPer.Text) & " Amount-" & TxtAmount.Text)
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
      If Trim(TxtCost.Text) <> "" Then
         .Columns("Cost").Value = Val(TxtCost.Text)
      End If
      .Columns("IsProduct").Value = Abs(ChkIsProduct.Value)
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("TotalAmount").Value = Val(TxtActualAmount.Text)
      .Columns("EmpComm").Value = IIf(Val(TxtEmpComm.Text) = 0, 0, Val(TxtEmpComm.Text))
      RsBody!Qty = Val(TxtQty.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!DiscPC = Val(TxtDiscPC.Text)
      RsBody!DiscPer = Val(TxtDiscPer.Text)
      RsBody!DiscVal = Val(TxtDiscVal.Text)
      If Trim(TxtCost.Text) <> "" Then
      RsBody!Cost = Val(TxtCost.Text)
      End If
      If IsNull(RsBody!Cost) Or RsBody!Cost = "" Then RsBody!Cost = 0
      RsBody!isProduct = Abs(ChkIsProduct.Value)
      RsBody!Amount = Val(TxtAmount.Text)
      RsBody!EmpComm = Val(TxtEmpComm.Text)
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
   TxtCode.Enabled = True
   BtnProduct.Enabled = True
   TxtCode.Text = ""
   TxtProductName.Text = ""
   TxtQty.Text = ""
   TxtPrice.Text = ""
   TxtDiscPC.Text = ""
   TxtDiscPer.Text = ""
   TxtDiscVal.Text = ""
   TxtAmount.Text = ""
   TxtCost.Text = ""
   TxtActualAmount.Text = ""
   TxtEmpComm.Text = ""
   ChkIsProduct.Value = 1
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtPID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("code").Text
      TxtProductName.Text = .Columns("ProductName").Text
      TxtQty.Text = .Columns("Qty").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtDiscPC.Text = .Columns("DiscPC").Value
      TxtDiscPer.Text = .Columns("DiscPer").Value
      TxtDiscVal.Text = .Columns("DiscVal").Value
      TxtCost.Text = .Columns("Cost").Value
      TxtEmpComm.Text = .Columns("EmpComm").Value
      TxtAmount.Text = .Columns("Amount").Text
      TxtActualAmount.Text = .Columns("TotalAmount").Text
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
   If Grid.rows = 1 Then Grid.MoveLast
   Exit Sub
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpOrderDate.DateValue + 1 & "',0),0)"
         vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = vQtyLoose & " " & cn.Execute("SELECT dbo.FunGetUnit(" & Val(TxtPID.Text) & ")").Fields(0).Value
         LblStock.Visible = vShowStock
         LblStockCaption.Visible = vShowStock
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetSale()
   On Error GoTo ErrorHandler
   ssql = "select h.*, c.AccountName, OrganizationName, BankMachineName, StoreName, TableName, EmpName, MemberName FROM SaleOrderHeader h left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join BankMachines b on b.BankMachineid = h.BankMachineid left outer join Members m on m.MemberID = h.MemberID inner join stores s on s.storeid = h.storeid left outer join Tables t on t.TableID = h.TableID left outer join Organizations o on o.OrganizationID = h.OrganizationID left outer join Employees e on e.EmpID = h.EmpID where isReplace=0 and h.OrderID=" & Val(TxtOrderID.Text) & " and OrderDate='" & DtpOrderDate.DateValue & "'" & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With cn.Execute(ssql)
      If Not .BOF Then
         DtpDeliveryDate.DateValue = IIf(IsNull(!DeliveryDate), "", !DeliveryDate)
         DTPDeliveryTime.Value = IIf(IsNull(!DeliveryTime), Now, !DeliveryTime)
         If IsNull(!InvType) Or !InvType = "" Then
            CmbType.ListIndex = 0
         Else
            CmbType.Text = !InvType
         End If
         TxtBillID.Text = IIf(IsNull(!BillID), "", !BillID)
         DtpBillDate.DateValue = IIf(IsNull(!BillDate), "01/01/1990", !BillDate)
         DtpPromiseDate.DateValue = !PromiseDate
         TxtStampID.Text = IIf(IsNull(!StampID), "1", !StampID)
         TxtStoreID.Text = !StoreID
         TxtStoreName.Text = !StoreName
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         TxtTableID.Text = IIf(IsNull(!TableId), "", !TableId)
         TxtTableName.Text = IIf(IsNull(!TableName), "", !TableName)
         TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
         TxtEmployeeName.Text = IIf(IsNull(!empname), "", !empname)
         TxtMemberID.Text = IIf(IsNull(!MemberID), "", !MemberID)
         TxtMemberName.Text = IIf(IsNull(!MemberName), "", !MemberName)
         TxtTotalAmount.Caption = !TotalAmount
         TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
         TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
         TxtServiceChargesPer.Text = IIf(IsNull(!ServiceChargesPer), "", !ServiceChargesPer)
         TxtServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtSTaxPer.Text = IIf(IsNull(!STaxPer), "", !STaxPer)
         TxtSTax.Text = IIf(IsNull(!STax), "", !STax)
         TxtManualBillNo.Text = IIf(IsNull(!ManualBillNo), "", !ManualBillNo)
         TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
         TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
         TxtRemarksUrdu.Text = IIf(IsNull(!RemarksUrdu), "", !RemarksUrdu)
         FrmOrderPrint.OptBankCard.Value = !BankCard
         FrmOrderPrint.OptCash.Value = !Cash
         FrmOrderPrint.OptCredit.Value = !Credit
         If FrmOrderPrint.OptBankCard.Value = True Then
            FrmOrderPrint.TxtInvoiceNo.Text = !InvoiceNo
            FrmOrderPrint.TxtCommision.Text = !Commision
            FrmOrderPrint.TxtBankMachineID.Text = !BankMachineID
            FrmOrderPrint.TxtBankMachineName.Text = !BankMachineName
            FrmOrderPrint.TxtCashReceivedCash.Text = ""
            FrmOrderPrint.TxtCustomerID.Text = ""
            FrmOrderPrint.TxtCustomerName.Text = ""
            FrmOrderPrint.TxtCashCustomer.Text = ""
            FrmOrderPrint.TxtBankCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
         End If
         If FrmOrderPrint.OptCash.Value = True Then
            FrmOrderPrint.TxtCommision.Text = ""
            FrmOrderPrint.TxtInvoiceNo.Text = ""
            FrmOrderPrint.TxtBankMachineID.Text = ""
            FrmOrderPrint.TxtBankMachineName.Text = ""
            FrmOrderPrint.TxtCashReceivedCash.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            FrmOrderPrint.TxtCustomerID.Text = ""
            FrmOrderPrint.TxtCustomerName.Text = ""
            FrmOrderPrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
            FrmOrderPrint.TxtBankCustomer.Text = ""
         End If
         If FrmOrderPrint.OptCredit.Value = True Then
            FrmOrderPrint.TxtCommision.Text = ""
            FrmOrderPrint.TxtInvoiceNo.Text = ""
            FrmOrderPrint.TxtBankMachineID.Text = ""
            FrmOrderPrint.TxtBankMachineName.Text = ""
            FrmOrderPrint.TxtCashReceivedCredit.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            FrmOrderPrint.TxtCustomerID.Text = IIf(IsNull(!CustomerID), "", Val(!CustomerID))
            FrmOrderPrint.TxtCustomerName.Text = IIf(IsNull(!AccountName), "", !AccountName)
            FrmOrderPrint.TxtCashCustomer.Text = ""
            FrmOrderPrint.TxtBankCustomer.Text = ""
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
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtBillDisc.Name Then Exit Sub
   TxtBillDiscPer.Text = Round((Val(TxtBillDisc.Text) * 100) / Val(TxtTotalAmount.Caption), 2)
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtBillDiscPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtBillDiscPer.Name Then Exit Sub
   TxtBillDisc.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtBillDiscPer.Text) / 100))
   Call SubCalculateFooter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCode_GotFocus()
   Grid.MoveLast
   Grid.MoveNext
End Sub

Private Sub TxtDiscPC_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDiscPC.Name Then Exit Sub
   If Val(TxtPrice.Text) = 0 Then Exit Sub
   TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / Val(TxtPrice.Text), 2)
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub TxtDiscPC_LostFocus()
'   Select Case Me.ActiveControl.Name
'   Case TxtCode.Name, TxtQty.Name, TxtDiscPC.Name
'      Exit Sub
'   End Select
'   Call GetDataFromTexBoxesToGrid
'End Sub

Private Sub TxtDiscPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDiscPer.Name Then Exit Sub
   TxtDiscPC.Text = Round((Val(TxtPrice.Text) * Val(TxtDiscPer.Text) / 100), 2)
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
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
   TxtDiscPC.Text = Round((Val(TxtPrice.Text) * Val(TxtDiscPer.Text) / 100), 2)
   Call SubCalculateBody
End Sub

Private Sub TxtProductName_Change()
   If ActiveControl.Name <> TxtProductName.Name Then Exit Sub
   Call FindRow
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
   'If ActiveControl.Name <> TxtServiceChargesPer.Name Then Exit Sub
   TxtServiceCharges.Text = SelfRound((Val(TxtTotalAmount.Caption) * Val(TxtServiceChargesPer.Text) / 100))
   Call SubCalculateFooter
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

Private Sub TxtTotalAmount_Change()
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

Private Sub TxtTotalDiscount_Change()
   On Error GoTo ErrorHandler
   If Len(TxtTotalDiscount.Caption) >= 5 Then
      TxtTotalDiscount.FontSize = 36
   Else
      TxtTotalDiscount.FontSize = 48
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTotalQty_Change()
   On Error GoTo ErrorHandler
   If Len(TxtTotalQty.Caption) >= 5 Then
      TxtTotalQty.FontSize = 36
   Else
      TxtTotalQty.FontSize = 48
   End If
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

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpOrderDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = cn.Execute("Select isnull(max(Bin_OrderID),0)+1 from Bin_SaleOrderHeader ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub FindRebate()
   Dim Rebate
   On Error GoTo ErrorHandler
    With cn.Execute("Select * from ProductOffers where Rebate <> 0 and ProductID = " & Val(TxtPID.Text))
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

Private Sub UserActivities()
   If vIsNewRecord = False Then
      With cn.Execute("Select  * from SaleOrderHeader where OrderID =" & TxtOrderID.Text & " And OrderDate = '" & DtpOrderDate.DateValue & "'")
          If Val(TxtEmployeeID.Text) <> IIf(IsNull(!EmpID), 0, !EmpID) Then
              cn.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Updated EmpID-" & !EmpID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
          End If
          If TxtMemberID.Text <> !MemberID Then
              cn.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Updated MemberID-" & !MemberID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
          End If
          If TxtStoreID.Text <> !StoreID Then
              cn.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Updated StoreID-" & !StoredID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
          End If
      End With
      Grid.MoveFirst
      For i = 1 To Grid.rows - 1
         With cn.Execute("Select * from SaleOrderBody Where OrderID = " & TxtOrderID.Text & " and OrderDate ='" & DtpOrderDate.DateValue & "' and Productid = " & Val(Grid.Columns("Productid").Text))
            If .EOF = True Then
               cn.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Inserted New ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & Grid.Columns("Qty").Text & " Price-" & Grid.Columns("Price").Text & " Disc-" & Grid.Columns("DiscPer").Text & " Amount-" & Grid.Columns("Amount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
            Else
               If Grid.Columns("Qty").Text <> !Qty Or Grid.Columns("Price").Text <> !Price Or Grid.Columns("discper").Text <> !DiscPer Then
                  cn.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Updated ProdcutID-" & Grid.Columns("Code").Text & " Qty-" & !Qty & " Price-" & !Price & " Disc-" & !DiscPer & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
               End If
            End If
         End With
      Grid.MoveNext
      Next
   Else
      cn.Execute ("Insert Into UserActivities values ('Sale Order'" & "," & TxtOrderID.Text & ",'" & DtpOrderDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
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

Private Sub BtnOrganization_Click()
   If FunSelectOrganization(ssButton, False) = True Then
'      If TxtCustomerID.Enabled Then TxtCustomerID.SetFocus
   Else
      TxtOrganizationID.SetFocus
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

Private Sub Textbox1_KeyDown(KeyCode As Integer, Shift As Integer)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyDown-Event behaviours for  ''
        '' Enter, Space, Delete & Tab keys to set      ''
        '' Behavior in TxtRemarksUrdu.Text, keys will behave ''
        '' as Normal Text writing behavior.            ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''
     
        'Space Key Behavior
       If KeyCode = 32 Then
        UniCode = &H20
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        KeyCode = 0
        
        'Enter Key Behavior
        ElseIf KeyCode = 13 Then
        UniCode = &HA
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        KeyCode = 0
        
        'Horizontal Tab Behavior
        ElseIf KeyCode = 9 Then
        UniCode = &H9
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        KeyCode = 0
        
        'Delete Key Behavior
        ElseIf KeyCode = 127 Then
        UniCode = &H7F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        KeyCode = 0
        
        End If
        If BtnSave.Enabled = False Then FormStatus = ChangeMode
        
        'This Function Got End There

End Sub

Private Sub Textbox1_KeyPress(KeyAscii As Integer)

        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyPress-Event behaviours for ''
        '' Alfabatic, Numaric & Symbolic keys to write ''
        '' Urdu. I've tried to make it near with Urdu  ''
        '' Phonetic Keyboard Layout.                   ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''
       
'If ModeValue = False Then

        'For Small Letter's Behaviors

        'a Key Behavior
        If KeyAscii = 97 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H627
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'b Key Behavior
        ElseIf KeyAscii = 98 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H628
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'c Key Behavior
        ElseIf KeyAscii = 99 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H686
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'd Key Behavior
        ElseIf KeyAscii = 100 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'e Key Behavior
        ElseIf KeyAscii = 101 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H639
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'f Key Behavior
        ElseIf KeyAscii = 102 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H641
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'g Key Behavior
        ElseIf KeyAscii = 103 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6AF
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'h Key Behavior
        ElseIf KeyAscii = 104 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6BE
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'i Key Behavior
        ElseIf KeyAscii = 105 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6CC
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'j Key Behavior
        ElseIf KeyAscii = 106 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62C
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'k Key Behavior
        ElseIf KeyAscii = 107 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6A9
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'l Key Behavior
        ElseIf KeyAscii = 108 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H644
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'm Key Behavior
        ElseIf KeyAscii = 109 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H645
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'n Key Behavior
        ElseIf KeyAscii = 110 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H646
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'o Key Behavior
        ElseIf KeyAscii = 111 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6C1
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'p Key Behavior
        ElseIf KeyAscii = 112 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H67E
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'q Key Behavior
        ElseIf KeyAscii = 113 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H642
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'r Key Behavior
        ElseIf KeyAscii = 114 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H631
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        's Key Behavior
        ElseIf KeyAscii = 115 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H633
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        't Key Behavior
        ElseIf KeyAscii = 116 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62A
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'u Key Behavior
        ElseIf KeyAscii = 117 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H621
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'v Key Behavior
        ElseIf KeyAscii = 118 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H637
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'w Key Behavior
        ElseIf KeyAscii = 119 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H648
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'x Key Behavior
        ElseIf KeyAscii = 120 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H634
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'y Key Behavior
        ElseIf KeyAscii = 121 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6D2
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'z Key Behavior
        ElseIf KeyAscii = 122 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H632
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        
        ' For Capital Latter's Behaviors
        
        'A Key Behavior
        ElseIf KeyAscii = 65 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H622
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'B Key Behavior
        ElseIf KeyAscii = 66 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &HFBB0
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'C Key Behavior
        ElseIf KeyAscii = 67 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62B
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'D Key Behavior
        ElseIf KeyAscii = 68 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H688
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'E Key Behavior
        ElseIf KeyAscii = 69 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H650
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'F Key Behavior
        ElseIf KeyAscii = 70 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H652
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'G Key Behavior
        ElseIf KeyAscii = 71 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H63A
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'H Key Behavior
        ElseIf KeyAscii = 72 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'I Key Behavior
        ElseIf KeyAscii = 73 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H649
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'J Key Behavior
        ElseIf KeyAscii = 74 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H636
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'K Key Behavior
        ElseIf KeyAscii = 75 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H62E
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'L Key Behavior
        ElseIf KeyAscii = 76 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &HFEFB
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'M Key Behavior
        ElseIf KeyAscii = 77 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H66B
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'N Key Behavior
        ElseIf KeyAscii = 78 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H6BA
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'O Key Behavior
        ElseIf KeyAscii = 79 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H629
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'P Key Behavior
        ElseIf KeyAscii = 80 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H64F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'Q Key Behavior
        ElseIf KeyAscii = 81 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H626
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'R Key Behavior
        ElseIf KeyAscii = 82 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H691
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'S Key Behavior
        ElseIf KeyAscii = 83 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H635
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'T Key Behavior
        ElseIf KeyAscii = 84 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H679
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'U Key Behavior
        ElseIf KeyAscii = 85 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H626
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'V Key Behavior
        ElseIf KeyAscii = 86 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H638
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'W Key Behavior
        ElseIf KeyAscii = 87 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H624
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'Z Key Behavior
        ElseIf KeyAscii = 88 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H698
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'Y Key Behavior
        ElseIf KeyAscii = 89 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &HFBAF
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        'Z Key Behavior
        ElseIf KeyAscii = 90 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H630
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        
        'For Numaric Key's Behaviors
        
        '0 Key Behavior
        ElseIf KeyAscii = 48 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 48
'        UniCode = &H660
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '1 Key Behavior
        ElseIf KeyAscii = 49 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 49
'        UniCode = &H661
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '2 Key Behavior
        ElseIf KeyAscii = 50 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 50
'        UniCode = &H662
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '3 Key Behavior
        ElseIf KeyAscii = 51 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 51
'        UniCode = &H663
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '4 Key Behavior
        ElseIf KeyAscii = 52 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 52
'        UniCode = &H664
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '5 Key Behavior
        ElseIf KeyAscii = 53 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 53
'        UniCode = &H665
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '6 Key Behavior
        ElseIf KeyAscii = 54 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 54
'        UniCode = &H666
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '7 Key Behavior
        ElseIf KeyAscii = 55 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 55
'        UniCode = &H667
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '8 Key Behavior
        ElseIf KeyAscii = 56 Or TxtRemarksUrdu.SelText <> "" Then
        UniCode = 56
'        UniCode = &H668
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '9 Key Behavior
        ElseIf KeyAscii = 57 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = 57
'        UniCode = &H669
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        ' Numaric Keys with 'Shift' Behavior
        
        ') Key Behavior
        ElseIf KeyAscii = 41 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &HFD3F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '! Key Behavior
        ElseIf KeyAscii = 33 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H21
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '@ Key Behavior
        ElseIf KeyAscii = 64 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H40
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '# Key Behavior
        ElseIf KeyAscii = 35 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H23
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '$ Key Behavior
        ElseIf KeyAscii = 36 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H24
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '% Key Behavior
        ElseIf KeyAscii = 37 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H66A
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '^ Key Behavior
        ElseIf KeyAscii = 94 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H5E
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '& Key Behavior
        ElseIf KeyAscii = 38 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H26
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '* Key Behavior
        ElseIf KeyAscii = 42 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H66D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '( Key Behavior
        ElseIf KeyAscii = 40 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &HFD3E
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        
        'For Special Characters
        
        'Symbols
        
        '? Key Behavior
        ElseIf KeyAscii = 63 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H61F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '/ Key Behavior
        ElseIf KeyAscii = 47 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        ', Key Behavior
        ElseIf KeyAscii = 44 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H60C
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '. Key Behavior
        ElseIf KeyAscii = 46 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H640
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '_ Key Behavior
        ElseIf KeyAscii = 95 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H5F
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '- Key Behavior
        ElseIf KeyAscii = 45 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '+ Key Behavior
        ElseIf KeyAscii = 43 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2B
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '= Key Behavior
        ElseIf KeyAscii = 61 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H3D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        ': Key Behavior
        ElseIf KeyAscii = 58 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H3A
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '; Key Behavior
        ElseIf KeyAscii = 59 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H201C
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '< Key Behavior
        ElseIf KeyAscii = 60 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H64E
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '> Key Behavior
        ElseIf KeyAscii = 62 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H650
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '{ Key Behavior
        ElseIf KeyAscii = 123 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2018
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '} Key Behavior
        ElseIf KeyAscii = 125 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2019
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '[ Key Behavior
        ElseIf KeyAscii = 91 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H5B
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '] Key Behavior
        ElseIf KeyAscii = 93 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H5D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '| Key Behavior
        ElseIf KeyAscii = 124 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H7C
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '\ Key Behavior
        ElseIf KeyAscii = 92 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H5C
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '~ Key Behavior
        ElseIf KeyAscii = 126 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H64B
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '` Key Behavior
        ElseIf KeyAscii = 96 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H64D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '" Key Behavior
        ElseIf KeyAscii = 34 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H2190
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        '' Key Behavior
        ElseIf KeyAscii = 39 Or TxtRemarksUrdu.SelText <> "" Then
        TxtRemarksUrdu.SelText = ""
        UniCode = &H201D
        TxtRemarksUrdu.Text = TxtRemarksUrdu.Text + ChrW(UniCode)
        
        End If
        KeyAscii = 0
 '       End If

        'This Function Got End There

End Sub

Private Sub GetSaleInvoice()
   On Error GoTo ErrorHandler
   ssql = "select h.*, c.AccountName, OrganizationName, BankMachineName, StoreName, TableName, EmpName, MemberName FROM SaleHeader h left outer join ChartofAccounts c on h.customerid = c.AccountNo left outer join BankMachines b on b.BankMachineid = h.BankMachineid left outer join Members m on m.MemberID = h.MemberID inner join stores s on s.storeid = h.storeid left outer join Tables t on t.TableID = h.TableID left outer join Organizations o on o.OrganizationID = h.OrganizationID left outer join Employees e on e.EmpID = h.EmpID where isReplace=0 and h.BillID=" & Val(TxtBillID.Text) & " and BillDate='" & DtpBillDate.DateValue & "'"
   With cn.Execute(ssql)
      If Not .BOF Then
         DtpDeliveryDate.DateValue = IIf(IsNull(!DeliveryDate), "", !DeliveryDate)
         DTPDeliveryTime.Value = IIf(IsNull(!DeliveryTime), Now, !DeliveryTime)
         If IsNull(!InvType) Or !InvType = "" Then
            CmbType.ListIndex = 0
         Else
            CmbType.Text = !InvType
         End If
         TxtStampID.Text = IIf(IsNull(!StampID), "1", !StampID)
         TxtStoreID.Text = !StoreID
         TxtStoreName.Text = !StoreName
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         TxtTableID.Text = IIf(IsNull(!TableId), "", !TableId)
         TxtTableName.Text = IIf(IsNull(!TableName), "", !TableName)
         TxtEmployeeID.Text = IIf(IsNull(!EmpID), "", !EmpID)
         TxtEmployeeName.Text = IIf(IsNull(!empname), "", !empname)
         TxtMemberID.Text = IIf(IsNull(!MemberID), "", !MemberID)
         TxtMemberName.Text = IIf(IsNull(!MemberName), "", !MemberName)
         TxtTotalAmount.Caption = !TotalAmount
         TxtBillDiscPer.Text = IIf(IsNull(!BillDiscPer), "", !BillDiscPer)
         TxtBillDisc.Text = IIf(IsNull(!BillDisc), "", !BillDisc)
         TxtServiceChargesPer.Text = IIf(IsNull(!ServiceChargesPer), "", !ServiceChargesPer)
         TxtServiceCharges.Text = IIf(IsNull(!ServiceCharges), "", !ServiceCharges)
         TxtSTaxPer.Text = IIf(IsNull(!STaxPer), "", !STaxPer)
         TxtSTax.Text = IIf(IsNull(!STax), "", !STax)
         TxtManualBillNo.Text = IIf(IsNull(!ManualBillNo), "", !ManualBillNo)
         TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
         TxtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
         TxtRemarksUrdu.Text = IIf(IsNull(!RemarksUrdu), "", !RemarksUrdu)
         FrmOrderPrint.OptBankCard.Value = !BankCard
         FrmOrderPrint.OptCash.Value = !Cash
         FrmOrderPrint.OptCredit.Value = !Credit
         If FrmOrderPrint.OptBankCard.Value = True Then
            FrmOrderPrint.TxtInvoiceNo.Text = !InvoiceNo
            FrmOrderPrint.TxtCommision.Text = !Commision
            FrmOrderPrint.TxtBankMachineID.Text = !BankMachineID
            FrmOrderPrint.TxtBankMachineName.Text = !BankMachineName
            FrmOrderPrint.TxtCashReceivedCash.Text = ""
            FrmOrderPrint.TxtCustomerID.Text = ""
            FrmOrderPrint.TxtCustomerName.Text = ""
            FrmOrderPrint.TxtCashCustomer.Text = ""
            FrmOrderPrint.TxtBankCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
         End If
         If FrmOrderPrint.OptCash.Value = True Then
            FrmOrderPrint.TxtCommision.Text = ""
            FrmOrderPrint.TxtInvoiceNo.Text = ""
            FrmOrderPrint.TxtBankMachineID.Text = ""
            FrmOrderPrint.TxtBankMachineName.Text = ""
            FrmOrderPrint.TxtCashReceivedCash.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            FrmOrderPrint.TxtCustomerID.Text = ""
            FrmOrderPrint.TxtCustomerName.Text = ""
            FrmOrderPrint.TxtCashCustomer.Text = IIf(IsNull(!CustomerName), !AccountName, !CustomerName)
            FrmOrderPrint.TxtBankCustomer.Text = ""
         End If
         If FrmOrderPrint.OptCredit.Value = True Then
            FrmOrderPrint.TxtCommision.Text = ""
            FrmOrderPrint.TxtInvoiceNo.Text = ""
            FrmOrderPrint.TxtBankMachineID.Text = ""
            FrmOrderPrint.TxtBankMachineName.Text = ""
            FrmOrderPrint.TxtCashReceivedCredit.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
            FrmOrderPrint.TxtCustomerID.Text = IIf(IsNull(!CustomerID), "", Val(!CustomerID))
            FrmOrderPrint.TxtCustomerName.Text = IIf(IsNull(!AccountName), "", !AccountName)
            FrmOrderPrint.TxtCashCustomer.Text = ""
            FrmOrderPrint.TxtBankCustomer.Text = ""
         End If
         TxtNetAmount.Caption = !TotalAmount
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

Private Sub PopulateSaleDataToGrid()
On Error GoTo ErrorHandler
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from SaleOrderBody where OrderID = " & Val(TxtOrderID.Text) & " and OrderDate = '" & DtpOrderDate.DateValue & "' and StampID  " & IIf(TxtStampID.Text = "", " is null", "=" & TxtStampID.Text), cn, adOpenDynamic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
      ssql = "select p.ProductName, b.code, b.* from SaleBody b join products p on p.productid = b.productid where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "' order by serialno"
      With cn.Execute(ssql)
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
            Grid.Columns("Code").Text = IIf(IsNull(!Code), "", !Code)
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !Qty
            Grid.Columns("QtyOrigional").Value = !Qty
            Grid.Columns("Price").Value = !Price
            Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
            Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
            Grid.Columns("DiscVal").Value = IIf(IsNull(!DiscVal), "", !DiscVal)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("IsProduct").Value = Abs(!isProduct)
            Grid.Columns("TotalAmount").Value = Val(!Price) * Val(!Qty)
            Grid.Columns("Cost").Value = IIf(IsNull(!Cost), 0, !Cost)
            Grid.Columns("EmpComm").Value = IIf(IsNull(!EmpComm), "", !EmpComm)
            TxtTotalQty.Caption = Val(TxtTotalQty.Caption) + Val(!Qty)
            'TxtTotalDiscount.Caption = Val(TxtTotalDiscount.Caption) + Val(!DiscVal)
            vTotDisc = vTotDisc + Val(!DiscVal)
            vTotalAmount = vTotalAmount + !Amount
            TxtTotalAmount.Caption = Val(TxtTotalAmount.Caption) + Grid.Columns("TotalAmount").Value
            
            RsBody.AddNew
            RsBody!Productid = !Productid
            RsBody!Code = !Code
            RsBody!PackingID = !PackingID
            RsBody!Multiplier = !Multiplier
            RsBody!QtyPack = !QtyPack
            RsBody!Qty = !Qty
            RsBody!Bonus = !Bonus
            RsBody!Price = !Price
            RsBody!RetailPrice = !RetailPrice
            RsBody!IsWSDiscb4ST = !IsWSDiscb4ST
            RsBody!IsWSSaleTax = !IsWSSaleTax
            RsBody!IsRetailSaleTax = !IsRetailSaleTax
            RsBody!TokenVal = !TokenVal
            RsBody!Offer = !Offer
            RsBody!SaleTaxPer = !SaleTaxPer
            RsBody!SaleTaxval = !SaleTaxval
            RsBody!DiscPC = !DiscPC
            RsBody!DiscPer = !DiscPer
            RsBody!DiscVal = !DiscVal
            RsBody!Amount = !Amount
            RsBody!Cost = !Cost
            RsBody!isProduct = !isProduct
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
   
''   RsBodySerial.Filter = 0
''   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
'   RsBodySerial.Open "Select * from SaleBodySerial where BillID=" & Val(TxtBillID.Text) & " and BillDate = '" & DtpBillDate.DateValue & "'", CN, adOpenDynamic, adLockBatchOptimistic
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleOrderHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmSaleOrderPOS) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmSaleOrderPOS & ", " & vUser & "," & TableHeaderFields(eFrmSaleOrderPOS) & " from SaleOrderHeader " & vbCrLf _
             & "Where OrderID = " & TxtOrderID.Text & " and OrderDate = '" & DtpOrderDate.DateValue & "'"
      cn.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.SaleOrderBodyBin (" & TableBodyFields(eFrmSaleOrderPOS) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmSaleOrderPOS) & " from SaleOrderBody " & vbCrLf _
             & "Where OrderID = " & TxtOrderID.Text & " and OrderDate = '" & DtpOrderDate.DateValue & "'"
      cn.Execute vStrSQL
  End If
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
