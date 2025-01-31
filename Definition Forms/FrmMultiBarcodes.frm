VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmMultiBarcodes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmMultiBarcodes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkSaleTax 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Show Sales Tax Value"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2895
      TabIndex        =   45
      Top             =   7290
      Width           =   2265
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmMultiBarcodes.frx":0ECA
      Left            =   2895
      List            =   "FrmMultiBarcodes.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   8190
      Width           =   3276
   End
   Begin VB.CheckBox ChkDiscountedPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discounted Price Include With Barcode"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2895
      TabIndex        =   5
      Top             =   6750
      Width           =   3120
   End
   Begin VB.ComboBox CmbPage 
      Height          =   315
      ItemData        =   "FrmMultiBarcodes.frx":0ECE
      Left            =   2895
      List            =   "FrmMultiBarcodes.frx":0F17
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   7650
      Width           =   3276
   End
   Begin VB.CheckBox ChkPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Price Include With Barcode"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2895
      TabIndex        =   6
      Top             =   7020
      Value           =   1  'Checked
      Width           =   2265
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
      Left            =   12960
      TabIndex        =   23
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
         Height          =   3750
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   24
         Tag             =   "NC"
         Text            =   "FrmMultiBarcodes.frx":1060
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
         TabIndex        =   25
         Top             =   90
         Width           =   135
      End
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      Top             =   2145
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   7
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4155
      Left            =   1515
      TabIndex        =   15
      Top             =   2460
      Width           =   9825
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
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
      stylesets(0).Picture=   "FrmMultiBarcodes.frx":1130
      AllowDelete     =   -1  'True
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedRow"
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
      Columns.Count   =   6
      Columns(0).Width=   2408
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   6826
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3810
      Columns(2).Caption=   "Description"
      Columns(2).Name =   "Description"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1905
      Columns(3).Caption=   "Qty"
      Columns(3).Name =   "Qty2"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1879
      Columns(4).Caption=   "Piece"
      Columns(4).Name =   "Qty"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).NumberFormat=   "########.##"
      Columns(4).FieldLen=   256
      Columns(5).Width=   1879
      Columns(5).Caption=   "GroupID"
      Columns(5).Name =   "GroupID"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   17330
      _ExtentY        =   7329
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   6420
      TabIndex        =   11
      Top             =   9360
      Width           =   1272
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMultiBarcodes.frx":114C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   2520
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2130
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMultiBarcodes.frx":1168
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2880
      TabIndex        =   17
      Top             =   2145
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   9990
      TabIndex        =   3
      Top             =   2145
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   5
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   5100
      TabIndex        =   9
      Top             =   9360
      Width           =   1272
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMultiBarcodes.frx":1184
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStartFrom 
      Height          =   315
      Left            =   1590
      TabIndex        =   4
      Top             =   6960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtTotQty 
      Height          =   315
      Left            =   7050
      TabIndex        =   19
      Top             =   7035
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7740
      TabIndex        =   12
      Top             =   9360
      Width           =   1272
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMultiBarcodes.frx":11A0
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   3780
      TabIndex        =   22
      Top             =   9360
      Width           =   1272
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Preview"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMultiBarcodes.frx":11BC
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtX 
      Height          =   315
      Left            =   10950
      TabIndex        =   28
      Tag             =   "NC"
      Top             =   7238
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      Alignment       =   2
      Appearance      =   0
      MaxLength       =   6
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtY 
      Height          =   315
      Left            =   11685
      TabIndex        =   29
      Tag             =   "NC"
      Top             =   7238
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      Alignment       =   2
      Appearance      =   0
      MaxLength       =   6
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JeweledBut.JeweledButton BtnSingleBarcode 
      Height          =   555
      Left            =   11535
      TabIndex        =   34
      Top             =   3105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   979
      TX              =   "Single Barcode"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmMultiBarcodes.frx":11D8
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpExpiryDate 
      Height          =   312
      Left            =   6384
      TabIndex        =   35
      Top             =   8808
      Width           =   1212
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpMenuFactureDate 
      Height          =   312
      Left            =   4848
      TabIndex        =   37
      Top             =   8808
      Width           =   1212
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
   Begin SITextBox.Txt TxtQty2 
      Height          =   315
      Left            =   8910
      TabIndex        =   2
      Top             =   2145
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   7395
      TabIndex        =   40
      Top             =   1485
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   6750
      TabIndex        =   1
      Top             =   2145
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   556
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   7
   End
   Begin JeweledBut.JeweledButton BtnProductRange 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2295
      TabIndex        =   46
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
      MICON           =   "FrmMultiBarcodes.frx":11F4
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCurrencySymbol 
      Height          =   315
      Left            =   11025
      TabIndex        =   47
      Top             =   7920
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Currency Symbol"
      Height          =   195
      Left            =   11025
      TabIndex        =   48
      Top             =   7695
      Width           =   1185
   End
   Begin VB.Label LblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   6780
      TabIndex        =   44
      Top             =   1935
      Width           =   795
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
      Left            =   4245
      TabIndex        =   43
      Top             =   1305
      Width           =   1005
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
      Left            =   3285
      TabIndex        =   42
      Top             =   1305
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Group ID"
      Height          =   195
      Left            =   7380
      TabIndex        =   41
      Top             =   1275
      Width           =   765
   End
   Begin VB.Label lblQty2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   8910
      TabIndex        =   39
      Top             =   1935
      Width           =   240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacture Date"
      Height          =   195
      Left            =   4845
      TabIndex        =   38
      Top             =   8595
      Width           =   1290
   End
   Begin VB.Label LblExpiryDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date"
      Height          =   192
      Left            =   6384
      TabIndex        =   36
      Top             =   8592
      Width           =   816
   End
   Begin VB.Label Label11 
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
      Left            =   2220
      TabIndex        =   33
      Top             =   8235
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "H Value"
      Height          =   195
      Left            =   10950
      TabIndex        =   32
      Top             =   7028
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "V Value"
      Height          =   195
      Left            =   11685
      TabIndex        =   31
      Top             =   7028
      Width           =   555
   End
   Begin VB.Label LblPrint 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "--- Print Settings ---"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10950
      TabIndex        =   30
      Top             =   6788
      Width           =   1290
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Page Size"
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
      Left            =   1920
      TabIndex        =   27
      Top             =   7710
      Width           =   870
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
      Left            =   11295
      TabIndex        =   26
      Top             =   540
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multiple BarCodes"
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
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   21
      Top             =   270
      Width           =   2565
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity"
      Height          =   195
      Left            =   6990
      TabIndex        =   20
      Top             =   6735
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Start From"
      Height          =   195
      Left            =   1590
      TabIndex        =   18
      Top             =   6735
      Width           =   720
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Piece"
      Height          =   195
      Left            =   10005
      TabIndex        =   16
      Top             =   1935
      Width           =   405
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   2910
      TabIndex        =   14
      Top             =   1935
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      Height          =   195
      Left            =   1500
      TabIndex        =   13
      Top             =   1935
      Width           =   765
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   60
      Width           =   345
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmMultiBarcodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataFile As Integer, Fl As Long, Chunks As Integer
Dim Application1 As New CRAXDRT.Application
Dim Fragment As Integer, Chunk() As Byte, i As Integer, FileName As String
Const ChunkSize As Integer = 16384
Const conChunkSize = 100
Dim strFileNm As String
Dim vCounter As Integer
Dim vIsNewRecord As Boolean
Dim Rs As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim ssql As String
Dim VStrSQL, vUnit As String
Dim vIsNewRow As Boolean
Dim vNoofPages As Integer
Dim vCurrentPage As Integer
Dim vRecNo As Integer
Dim vStartFrom As Integer
Dim vCurrentRecord As Integer
Dim vProductID, vGroupID  As String
Dim vProductQty  As Integer, vQtyLoose As Double
Dim vProductQty2 As Single
Dim vStrBarcode As String
Dim vStrProductQty2 As String
Dim vShowBarcodeDesc As Boolean

Private Sub BtnClear_Click()
   SubClearFields
End Sub

Private Sub SubPrinterSetting()
   On Error GoTo ErrorHandler
   Dim vPrinter() As String
   vPrinter = Split(CmbPrinters.Text, ",")
'   If cn.Execute("Select * From PrinterSetting where size = '" & CmbPage.Text & "'").RecordCount >= 1 Then
'      cn.Execute "UPDATE PrinterSetting set x = " & Val(TxtX.Text) & " , y = " & Val(TxtY.Text) & ", DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "'"
'   Else
'      cn.Execute "INSERT INTO PrinterSetting Values(" & Val(TxtX.Text) & " ," & Val(TxtY.Text) & ",'" & CmbPage.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "')"
'   End If
   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'Multi BarCode' and LocalComputerName = '" & LocalComputerName & "'"
   If cn.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, X, Y, Size, DeviceName, DriverName, Port, IsPreview ) Values ('" & LocalComputerName & "', 'Multi BarCode'," & Val(TxtX.Text) & "," & Val(TxtY.Text) & ",'" & CmbPage.Text & " ','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & 0 & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & CmbPage.Text & "', X = " & Val(TxtX.Text) & ", Y = " & Val(TxtY.Text) & ", DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & 0 & " Where FormType = 'Multi BarCode' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   cn.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   Call SubPrinterSetting
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub SubCalculate()
'   Dim i As Integer
'   With CN.Execute("Select * from Products where Groupid = '045'")
'      For i = 1 To .RecordCount
'         Grid.Columns("ID").Text = !ProductID
'         Grid.Columns("Name").Text = !ProductName
'         Grid.Columns("Qty").Text = "1"
'         Grid.Update
'         Grid.AddNew
'         TxtTotQty.Text = Val(TxtTotQty.Text) + 1
'         .MoveNext
'      Next i
'   End With
'End Sub

Private Sub SubBarCodeGenerate()
   On Error GoTo ErrorHandler
   If Grid.Rows = 1 Then
      MsgBox "No Product ID selected", vbOKOnly, Me.Caption
      TxtProductID.SetFocus
      Exit Sub
   End If
   
   Load FrmBarcodeViewer
   vCounter = IIf(Val(TxtStartFrom.Text) = 0, 0, Val(TxtStartFrom.Text) - 1)
   If CmbPage.ListIndex <= 7 Then
       vNoofPages = IIf((Val(TxtTotQty.Text) + vCounter) Mod CmbPage.ItemData(CmbPage.ListIndex) = 0, (Val(TxtTotQty.Text) + vCounter) \ CmbPage.ItemData(CmbPage.ListIndex), ((Val(TxtTotQty.Text) + vCounter) \ CmbPage.ItemData(CmbPage.ListIndex)) + 1)
    Else
        vNoofPages = 1
    End If
   vCurrentRecord = 0
   cn.Execute ("delete from pic")
   For vCurrentPage = 1 To vCounter
      cn.Execute ("Insert Into Pic Values('" & vProductID & "',null," & vCurrentPage & ",null" & ")")
   Next vCurrentPage
   Grid.MoveFirst
   
   Call SetBarcode
   
   vProductID = Right("00000" + Grid.Columns("ID").Text, 5)
   With cn.Execute("select *, isnull(discpc,0) as Disc from products where productid=" & Val(vProductID))
      FrmBarcodeViewer.TxtBarCode.Text = vStrBarcode
      FrmBarcodeViewer.ParaInProductID = vProductID
      FrmBarcodeViewer.ParaInCompany = ObjRegistry.CompanyShortName
      FrmBarcodeViewer.ParaInProductName = !ProductName & " (" & Val(Grid.Columns("Qty2").Value) & ")"
      FrmBarcodeViewer.ParaInRate = IIf(ChkPrice.Value = 1, !RetailPrice * Val(Val(Grid.Columns("Qty2").Value)), IIf(ChkDiscountedPrice.Value = 1, !RetailPrice - !Disc, 0))
      
   
   End With
      
   FrmBarcodeViewer.cmdEANCreate.Value = True
'   ssql = "Select * from ProductBarcodes where ProductID = '" & vProductID & "' and code = '" & Val(vCode) & "'"
   ssql = "Select * from ProductBarcodes where ProductID = " & Val(vProductID) & " and code = '" & Val((vCode)) & "'"
   If cn.Execute(ssql).RecordCount = 0 Then
       'CN.Execute "Delete From ProductBarcodes where ProductID = '" & vProductID & "' and code like '11%'"
       cn.Execute "INSERT into ProductBarcodes(ProductID,Code,qty) values (" & Val(vProductID) & ",'" & Val((vCode)) & "'," & IIf(Val(Grid.Columns("Qty2").Value) = 0, "Null", Val(Grid.Columns("Qty2").Value)) & ")"
   End If
            
'   FrmBarcodeViewer.Show
   
'   For vCurrentPage = 1 To vNoofPages
      If Rs.State = adStateOpen Then Rs.Close
      Rs.CursorLocation = adUseClient
      Rs.Open "select * from pic", cn, adOpenStatic, adLockOptimistic
      For vCounter = 1 To Val(TxtTotQty.Text)
         If vProductQty = 0 Then
            'CN.Execute "update products set code = '" & Val(FrmBarcodeViewer.txtBarcode.Text) & "' where productid = '" & vProductID & "' and (code is null or code='')"
            Grid.MoveNext
                   
'            vProductID = Grid.Columns("ID").Text
'            vProductQty = Grid.Columns("Qty").Value
'            vProductQty2 = Val(Grid.Columns("Qty2").Value)
'            If InStr(CStr(vProductQty2), ".") <> 0 Then vStrProductQty2 = Split(CStr(vProductQty2), ".")(1)
            Call SetBarcode
            With cn.Execute("select *, isnull(discpc,0) as Disc from products where productid = " & Val(vProductID))
'               FrmBarcodeViewer.TxtBarcode.Text = IIf(vProductQty2 = 0, "0110", "01" & Right("99" + CStr((vStrProductQty2)), 2)) & !GroupID & vProductID
               FrmBarcodeViewer.TxtBarCode.Text = vStrBarcode
               FrmBarcodeViewer.ParaInProductID = vProductID
               FrmBarcodeViewer.ParaInCompany = ObjRegistry.CompanyShortName  'CN.Execute("Select ShortName from company").Fields(0).Value
               FrmBarcodeViewer.ParaInProductName = !ProductName & "(" & Val(Grid.Columns("Qty2").Value) & ")"
               FrmBarcodeViewer.ParaInRate = IIf(ChkPrice.Value = 1, !RetailPrice * Val(Grid.Columns("Qty2").Value), IIf(ChkDiscountedPrice.Value = 1, !RetailPrice - !Disc, 0))
            End With
            FrmBarcodeViewer.cmdEANCreate.Value = True
            ssql = "Select * from ProductBarcodes where ProductID = " & Val(vProductID) & " and code = '" & Val((vCode)) & "'"
            If cn.Execute(ssql).RecordCount = 0 Then
               'CN.Execute "Delete From ProductBarcodes where ProductID = '" & vProductID & "' and code like '11%'"
'               cn.Execute "INSERT into ProductBarcodes(ProductID,Code,qty) values ('" & vProductID & "','" & Val(vCode) & "'," & IIf(Val(Grid.Columns("Qty2").Value) = 0, "Null", Val(Grid.Columns("Qty2").Value)) & ")"
                cn.Execute "INSERT into ProductBarcodes(ProductID,Code,qty) values (" & Val(vProductID) & ",'" & Val((vCode)) & "'," & IIf(Val(Grid.Columns("Qty2").Value) = 0, "Null", Val(Grid.Columns("Qty2").Value)) & ")"
            End If
         End If
         Rs.AddNew
            strFileNm = "c:\" & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(FrmBarcodeViewer.ParaInProductName, "/", "-"), """", "-"), "\", "-"), ".", "-"), "*", "-"), "?", "-"), "&", "-"), ":", "-") & " " & vProductID & ".bmp"
            DataFile = 1
            Open strFileNm For Binary Access Read As DataFile
             Fl = LOF(DataFile)   ' Length of data in file
             If Fl = 0 Then Close DataFile: Exit Sub
             Chunks = Fl \ ChunkSize
             Fragment = Fl Mod ChunkSize
             ReDim Chunk(Fragment)
             Get DataFile, , Chunk()
             Rs(1).AppendChunk Chunk()
             ReDim Chunk(ChunkSize)
             For i = 1 To Chunks
                 Get DataFile, , Chunk()
                 Rs(1).AppendChunk Chunk()
             Next i
             Close DataFile
            vProductQty = vProductQty - 1
            'vCurrentRecord = vCurrentRecord + 1
            'vCounter = vCounter + 1
            Rs(2) = vCounter + vCurrentPage
            Rs(0) = vProductID
'            Rs(3) = Val(vCode)
            Rs(3) = Val(vCode)
            
         Rs.Update
      Next vCounter
      vCounter = 0
      Rs.Close
      Set Rs = Nothing
      'Call BtnPrint_Click
      'CN.Execute ("delete from pic")
'   Next vCurrentPage
   Kill "c:\*.bmp"
   'Call BtnClear_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
       If TxtDescription.Visible = True Then TxtDescription.SetFocus Else TxtQty.SetFocus
   Else
      TxtProductID.SetFocus
   End If
End Sub

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.ParaInWhere = " and isLocked = 0 and isNoCostProduct = 0"
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtProductID.Text = SchProduct.ParaOutID
   End If
   '---------------------------
   VStrSQL = "SELECT ProductName, GroupID from Products where ProductID = " & Val(TxtProductID.Text) & " and isLocked = 0 and isNoCostProduct = 0"
   With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtProductName.Text = !ProductName
         TxtGroupID.Text = !GroupID
         
         If ObjRegistry.ShowSavedStock = True Then
            VStrSQL = "select qtyloose from currentStockStore where Productid = " & Val(TxtProductID.Text)
            With cn.Execute(VStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            VStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & ",NULL,0,0,0,0,0,0,'" & Date + 1 & "',0),0)"
            vQtyLoose = cn.Execute(VStrSQL).Fields(0).Value
         End If
         
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
         
         FunSelectProduct = True
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         TxtProductID.Text = ""
         TxtProductName.Text = ""
         TxtGroupID.Text = ""
         Exit Function
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
      While Not .EOF
         TxtGroupID.Text = !GroupID
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtQty.Text = !QtyLoose
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
'   SchProductRange.Show vbModal, Me
'   If SchProductRange.ParaOutFromID <> "" Then
'   Dim vPID As Long, vCounter As Long
'   vPID = SchProductRange.ParaOutFromID
'   For vCounter = CLng(SchProductRange.ParaOutFromID) To CLng(SchProductRange.ParaOutToID)
'      TxtProductID.Text = vPID
'      FunSelectProduct ssValidate, False
'      TxtQty.Text = SchProductRange.ParaOutQty
'      GetDataFromTexBoxesToGrid
'      vPID = vPID + 1
'      DoEvents
'   Next vCounter
'   End If
   FrmProductRangeGrid.Show vbModal, Me
   RsTemp.Filter = ""
   If RsTemp.RecordCount > 0 Then
      PopulateTempToGrid
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSingleBarcode_Click()
   FrmSingleBarcodes.Show
End Sub

Private Sub ChkDiscountedPrice_Click()
'   If ChkDiscountedPrice.Value = 1 Then ChkPrice.Value = 0
End Sub

Private Sub ChkPrice_Click()
'   If ChkPrice.Value = 1 Then ChkDiscountedPrice.Value = 0
End Sub

Private Sub CmbPage_Click()
   On Error GoTo ErrorHandler
'   With cn.Execute("select * from PrinterSetting where Size = '" & CmbPage.Text & "'")
'     If .RecordCount > 0 Then
'        TxtX.Text = !x
'        TxtY.Text = !Y
'     End If
'     .Close
'   End With
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub JeweledButton1_Click()

End Sub

Private Sub LblPrint_Click()
   TxtX.Enabled = Not TxtX.Enabled
   TxtY.Enabled = Not TxtY.Enabled
   If TxtX.Enabled Then TxtX.SetFocus
   If TxtX.Enabled Then
      LblPrint.ForeColor = vbBlack
   Else
      'TxtFirst.SetFocus
      LblPrint.ForeColor = &H800000
   End If
   If TxtX.Enabled = False Then
      Call SubPrinterSetting
   End If
End Sub

Private Sub TxtProductID_Change()
   If ActiveControl.Name <> TxtProductID.Name Then Exit Sub
   If TxtProductName.Text <> "" Then TxtProductName.Text = ""
End Sub

Private Sub TxtProductID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtProductID_Validate(Cancel As Boolean)
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtProductID.Text) = "" Then Exit Sub
   vTemp = FunSelectProduct(ssValidate, False)
   If vTemp = False Then
      vTemp = FunSelectProduct(ssButton, False)
      Cancel = False
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Multiple Barcodes"
   
   DtpExpiryDate.DateValue = Date
   DtpMenuFactureDate.DateValue = Date
   
   LblStock.Visible = False
   LblStockCaption.Visible = False
   
   CmbPrinters.Clear
   CmbPrinters.AddItem "Default,winspool,LPT1"
   Dim p
   For Each p In Printers
      CmbPrinters.AddItem p.DeviceName & "," & p.DriverName & "," & p.Port
   Next p
   CmbPrinters.ListIndex = 0
'   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(0) & "' where RegistryKey = 'DeviceName'")
'   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(1) & "' where RegistryKey = 'DriverName'")
'   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(2) & "' where RegistryKey = 'Port'")
    
    DtpExpiryDate.DateValue = DateAdd("m", 3, Date)
    
     '''''''''''''''' Form Default Setting  ''''''''''''''''''''''
   ssql = "select * from FormDefaultSetting Where FormType = 'Multi BarCode' and LocalComputerName = '" & LocalComputerName & "'"
   With cn.Execute(ssql)
     If .RecordCount > 0 Then
        If Not IsNull(!DeviceName) Then
            CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
            TxtX.Text = !x
            TxtY.Text = !Y
            CmbPage.Text = !Size
        Else
            CmbPrinters.ListIndex = 0
        End If
     End If
     .Close
   End With
   ''''''''''''''''''''''''''''''''''''''''''''''

'   CmbPage.ListIndex = 5
'   With cn.Execute("select * from PrinterSetting")
'     If .RecordCount > 0 Then
'        TxtX.Text = !x
'        TxtY.Text = !Y
'        CmbPage.Text = !Size
'        If Not IsNull(!DeviceName) Then
'            CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
'        Else
'            CmbPrinters.ListIndex = 0
'        End If
'     End If
'     .Close
'   End With
   TxtX.Enabled = False
   TxtY.Enabled = False
   LblPrint.ForeColor = &H800000
   HelpLocation Me
   
   vShowBarcodeDesc = ObjRegistry.ShowBarcodeDesc
   
   If vShowBarcodeDesc = False Then
      LblDescription.Visible = False
      TxtDescription.Visible = False
      Grid.Columns("Description").Visible = False
      
      lblQty2.Left = lblQty2.Left - TxtDescription.Width
      TxtQty2.Left = TxtQty2.Left - TxtDescription.Width
      
      LblQty.Left = LblQty.Left - TxtDescription.Width
      TxtQty.Left = TxtQty.Left - TxtDescription.Width
      
      Grid.Width = Grid.Width - TxtDescription.Width
      
   End If
   
   If ObjRegistry.ShowBarCodeQty = False Then
    lblQty2.Visible = False
    TxtQty2.Visible = False
    Grid.Columns("Qty2").Visible = False
        
    LblQty.Left = LblQty.Left - TxtQty2.Width
    TxtQty.Left = TxtQty.Left - TxtQty2.Width
    
    Grid.Width = Grid.Width - TxtQty2.Width
   End If
   SubClearFields
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = "Grid" Then
         Grid_DblClick
      Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtProductID.Enabled Then TxtProductID.SetFocus: Call SubClearDetailArea
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtProductID.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtQty.SetFocus
      End Select
   ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtProductID.Name Then
         KeyCode = 0
         TxtStartFrom.SetFocus
   ElseIf Shift = vbCtrlMask Then
      If ActiveControl.Name = Grid.Name Then
         If KeyCode = vbKeyDelete Then
            If Trim(Grid.Columns("ID").Text <> "") Then Call mniRemoveRow_Click
            KeyCode = 0
         Else
            KeyCode = 0: Exit Sub
         End If
      End If
      Select Case KeyCode
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
         Case vbKeyV
            If BtnPreview.Enabled Then BtnPreview_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
       RptReportViewer.Caption = "Multiple Barcode"
       RptReportViewer.Show vbModal
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   If SetReport Then RptReportViewer.Report.PrintOut False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function SetReport() As Boolean
   On Error GoTo ErrorHandler
   SetReport = False
   SubBarCodeGenerate
   Dim vCurrencySymbol As String
   vCurrencySymbol = IIf(TxtCurrencySymbol.Text = "", "Rs.", TxtCurrencySymbol.Text)
   If RsReport.State = adStateOpen Then RsReport.Close
   'TxtActualAmount.Text - (TxtActualAmount.Text * (100 / (100 + Val(TxtSaleTaxPer.Text))))
'   VStrSQL = "select pic.f1, pic.Barcode, p.ProductID, ProductName, ProductName1, SubGroupName, " & IIf(ChkPrice.Value = 0, IIf(ChkDiscountedPrice.Value = 0, "'' as RetailPrice", "'Rs.' + cast(cast(RetailPrice-isnull(discpc,0) as int) as varchar(10)) RetailPrice"), "'Rs.' + cast(cast(RetailPrice as int) as varchar(10)) RetailPrice") & " from pic left outer join Products p on pic.productid = p.ProductID left outer join SubGroups sg on sg.SubGroupID = p.SubGroupID Order by Sr"
'   vStrSQL = "Select pic.f1, pic.Barcode, p.ProductID, Desc1, ProductName + isnull('(' + cast(cast(pb.qty as numeric(5,3)) as varchar(8)) + '}','') as ProductName, " & vbCrLf _
   + " ProductName1, SubGroupName, " & IIf(ChkPrice.Value = 0, IIf(ChkDiscountedPrice.Value = 0, "'' as RetailPrice", "'" & vCurrencySymbol & "' + cast(cast(RetailPrice*isnull(pb.qty,1)-isnull(discpc,0) as int) as varchar(10)) RetailPrice"), "'" & vCurrencySymbol & "' + cast(cast(RetailPrice*isnull(pb.qty,1) as int) as varchar(10)) RetailPrice") & ", DiscPC, '" & vCurrencySymbol & " ' + cast(cast((RetailPrice*isnull(pb.qty,1) - isnull(DiscPC,0)) as int) as varchar(10)) DiscPrice, " & vbCrLf _
   + " b.expirydate, " & IIf(ChkSaleTax.Value = 0, "0", " Round(RetailPrice - RetailPrice * (100 / (100 + SaletaxPer)), 2) ") & " as SaletaxValue  " & vbCrLf _
   + " from pic left outer join Products p on pic.productid = p.ProductID" & vbCrLf _
   + " LEFT OUTER JOIN ProductBarcodes pb on pb.productid = pic.productid and pic.barcode = pb.code " & vbCrLf _
   + " left outer join SubGroups sg on sg.SubGroupID = p.SubGroupID " & vbCrLf _
   + " left outer join (Select  productid, max(expirydate) expirydate from purchasebody group by productid )b on b.productid = p.productid  Order by Sr "
   
   VStrSQL = "Select pic.f1, pic.Barcode, p.ProductID, Desc1, ProductName + isnull('(' + cast(cast(pb.qty as numeric(5,3)) as varchar(8)) + '}','') as ProductName, " & vbCrLf _
   + " ProductName1, SubGroupName, RackName, " & vbCrLf _
   + IIf(ChkPrice.Value = 0, "'' as RetailPrice", "'" & vCurrencySymbol & "' + cast(cast(RetailPrice*isnull(pb.qty,1) as int) as varchar(10)) RetailPrice") & "," & vbCrLf _
   + IIf(ChkDiscountedPrice.Value = 0, "'' as DiscPrice", "'" & vCurrencySymbol & "' + cast(cast(RetailPrice*isnull(pb.qty,1)-isnull(discpc,0) as int) as varchar(10)) DiscPrice") & vbCrLf _
   + " , DiscPC, b.expirydate, " & IIf(ChkSaleTax.Value = 0, "0", " Round(RetailPrice - RetailPrice * (100 / (100 + SaletaxPer)), 2) ") & " as SaletaxValue  " & vbCrLf _
   + " from pic left outer join Products p on pic.productid = p.ProductID" & vbCrLf _
   + " LEFT OUTER JOIN Racks Rk on Rk.RackID  = p.RackID " & vbCrLf _
   + " LEFT OUTER JOIN ProductBarcodes pb on pb.productid = pic.productid and pic.barcode = pb.code " & vbCrLf _
   + " left outer join SubGroups sg on sg.SubGroupID = p.SubGroupID " & vbCrLf _
   + " left outer join (Select  productid, max(expirydate) expirydate from purchasebody group by productid )b on b.productid = p.productid  Order by Sr "
   RsReport.Open VStrSQL, cn, adOpenDynamic, adLockReadOnly
   If CmbPage.ListIndex = 0 Then
'      Set RptReportViewer.Report = New CrpMultiBarA4ShelfTracker
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarA4ShelfTracker.rpt")
   ElseIf CmbPage.ListIndex = 1 Then
'      Set RptReportViewer.Report = New CrpMultiBarCode120
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCode120.rpt")
   ElseIf CmbPage.ListIndex = 2 Then
'      Set RptReportViewer.Report = New CrpMultiBarCode96
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCode96.rpt")
   ElseIf CmbPage.ListIndex = 3 Then
'      Set RptReportViewer.Report = New CrpMultiBarCode84
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCode84.rpt")
   ElseIf CmbPage.ListIndex = 4 Then
'      Set RptReportViewer.Report = New CrpMultiBarCode80
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCode80.rpt")
   ElseIf CmbPage.ListIndex = 5 Then
'      Set RptReportViewer.Report = New CrpMultiBarCode65
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCode65.rpt")
   ElseIf CmbPage.ListIndex = 6 Then
'      Set RptReportViewer.Report = New CrpMultiBarCode50
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCode50.rpt")
   ElseIf CmbPage.ListIndex = 7 Then
'      Set RptReportViewer.Report = New CrpMultiBarCode40
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCode40.rpt")
   ElseIf CmbPage.ListIndex = 8 Then
'      Set RptReportViewer.Report = New CrpMultiBarCode25
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCode25.rpt")
   ElseIf CmbPage.ListIndex = 9 Then
'      Set RptReportViewer.Report = New CrpMultiBarCodeContinues19X28
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCodeContinues19X28.rpt")
   ElseIf CmbPage.ListIndex = 10 Then
'      Set RptReportViewer.Report = New CrpMultiBarCodeContinues19X28
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCodeContinuesExpiry19X28.rpt")
   ElseIf CmbPage.ListIndex = 11 Then
'      Set RptReportViewer.Report = New CrpMultiBarCodeContinues19X30
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCodeContinues19X30.rpt")
   ElseIf CmbPage.ListIndex = 12 Then
'      Set RptReportViewer.Report = New CrpMultiBarCodeContinues19X30
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCodeContinuesExpiry19X30.rpt")
   ElseIf CmbPage.ListIndex = 13 Then
'      Set RptReportViewer.Report = New CrpMultiBarCodeContinues25X32
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCodeContinues25X32.rpt")
   ElseIf CmbPage.ListIndex = 14 Then
'      Set RptReportViewer.Report = New CrpMultiBarCodeContinues28X38
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCodeContinues28X38.rpt")
   ElseIf CmbPage.ListIndex = 15 Then
'      Set RptReportViewer.Report = New CrpMultiBarCodeContinues28X38
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCodeContinuesExpiry28X38.rpt")
   ElseIf CmbPage.ListIndex = 16 Then
'      Set RptReportViewer.Report = New CrpMultiBarCodeContinues25X50
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCodeContinues25X50.rpt")
   ElseIf CmbPage.ListIndex = 17 Then
'      Set RptReportViewer.Report = New CrpMultiBarCodeContinues25X50
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCodeContinues50X100.rpt")
  ElseIf CmbPage.ListIndex = 18 Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCode25Expiry.rpt")
   End If
   
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyShortName 'CN.Execute("Select ShortName from company").Fields(0).Value
   If CmbPage.ListIndex = 10 Or CmbPage.ListIndex = 12 Or CmbPage.ListIndex = 14 Or CmbPage.ListIndex = 15 Then
      RptReportViewer.Report.ParameterFields(2).AddCurrentValue Format(DtpExpiryDate.DateValue, "d/mm/yy")
      RptReportViewer.Report.ParameterFields(3).AddCurrentValue Format(DtpMenuFactureDate.DateValue, "d/mm/yy")
   End If
   Dim vPrinter() As String
   vPrinter = Split(CmbPrinters.Text, ",")
'   RptReportViewer.Report.PrinterName = vPrinter(0)
'   RptReportViewer.Report.DriverName = vPrinter(1)
'   RptReportViewer.Report.PortName = vPrinter(2)
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   
   If CmbPage.ListIndex < 9 Then
      RptReportViewer.Report.PaperSize = crPaperA4
      RptReportViewer.Report.LeftMargin = TxtX.Text
      RptReportViewer.Report.TopMargin = TxtY.Text
   Else
      RptReportViewer.Report.LeftMargin = TxtX.Text
      RptReportViewer.Report.TopMargin = TxtY.Text
   End If
   
   
   'RptReportViewer.Show
   'RptReportViewer.Report.PrintOut False
   'MsgBox "Print has been sent to the printer", vbInformation + vbOKOnly, "Alert"
   SetReport = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Public Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is ComboBox Then
      End If
   Next
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ID").Text = " "
   Grid.Update
   ChkPrice.Value = 1
   BtnProduct.Enabled = True
   TxtProductID.Enabled = True
   If TxtProductID.Visible = True Then TxtProductID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   On Error GoTo ErrorHandler
   If Trim(TxtProductID.Text) = "" Then
      'MsgBox "Enter Group ID.", vbExclamation, "Alert"
      If TxtProductID.Enabled = True Then TxtProductID.SetFocus
      Exit Sub
   End If
   
   If Val(TxtQty.Text) = 0 Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      If TxtQty.Enabled = True Then TxtQty.SetFocus
      Exit Sub
   End If
   
   LblStock.Visible = False
   LblStockCaption.Visible = False
'   Grid.Bookmark = vBm

   
   '-------------------------------------------------------------------
   If Val(Val(TxtQty.Text) * Val(IIf(TxtQty2.Text = "", 1, TxtQty2.Text))) > Val(vQtyLoose) And ObjRegistry.AllowNegativeStockInBarcodes = True Then
      MsgBox "Barcodes are greater than available stock", vbInformation + vbOKOnly, "Error"
      TxtQty.SetFocus
      Exit Sub
   End If
   

   '-------------------------------------------------------------------
   If Trim(Grid.Columns("ID").Text) = "" Then
      TxtTotQty.Text = Val(TxtTotQty.Text) + Val(TxtQty.Text)
   ElseIf Trim(Grid.Columns("ID").Text) = Trim(TxtProductID.Text) Then
      TxtTotQty.Text = Val(TxtTotQty.Text) + Val(TxtQty.Text) - Grid.Columns("Qty").Text
   Else
   
   End If
   If TxtProductID.Enabled = True Then
         Grid.Columns("ID").Text = TxtProductID.Text
   'Else
   '      MsgBox "The record already exist"
   '      SubClearDetailArea
   '      If TxtProductID.Enabled Then TxtProductID.SetFocus
   '      Exit Sub
   End If
   Grid.Redraw = False
   With Grid
      .Columns("Description").Text = TxtDescription.Text
      .Columns("Name").Text = TxtProductName.Text
      .Columns("Qty").Text = TxtQty.Text
      .Columns("Qty2").Text = TxtQty2.Text
      .Columns("GroupID").Text = TxtGroupID.Text
      .MoveLast
      If Trim(.Columns("ID").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("id").Text = " "
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

Private Sub TxtQty_LostFocus()
   If Me.ActiveControl.Name = TxtProductID.Name Then Exit Sub
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub SubClearDetailArea()
   TxtProductID.Enabled = True
   TxtProductID.Text = ""
   TxtProductName.Text = ""
   TxtDescription.Text = ""
   TxtQty.Text = ""
   TxtQty2.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ID").Text
      TxtProductName.Text = .Columns("Name").Text
      TxtDescription.Text = .Columns("Description").Text
      TxtQty.Text = .Columns("Qty").Text
      TxtQty2.Text = .Columns("Qty2").Text
      TxtGroupID.Text = .Columns("GroupID").Text
      
      If ObjRegistry.ShowSavedStock = True Then
            VStrSQL = "select qtyloose from currentStockStore where Productid = " & Val(TxtProductID.Text)
            With cn.Execute(VStrSQL)
               If .RecordCount > 0 Then
                  vQtyLoose = .Fields(0).Value
               Else
                  vQtyLoose = 0
               End If
            End With
         Else
            VStrSQL = "select isnull(dbo.FunStock(" & Val(TxtProductID.Text) & ",NULL,0,0,0,0,0,0,'" & Date + 1 & "',0),0)"
            vQtyLoose = cn.Execute(VStrSQL).Fields(0).Value
         End If
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose(" & Val(TxtProductID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotQty.Text = Val(TxtTotQty.Text) - Grid.Columns("Qty").Value
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
  End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("ID").Text) = "" Then
      TxtProductID.Text = ""
      TxtProductID.Enabled = True
      TxtProductID.SetFocus
      vIsNewRow = True
   Else
      vBm = Grid.Bookmark
      TxtProductID.Enabled = False
      If TxtQty2.Enabled = True And TxtQty2.Visible Then TxtQty2.SetFocus Else TxtQty.SetFocus
      vIsNewRow = False
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(Grid.Columns("ID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ID").Text) = "" Then Exit Sub
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   Grid.SelBookmarks.RemoveAll
   Grid.Refresh
   GetDataBackFromGridToTexBoxes
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SetBarcode()
   vProductQty = Grid.Columns("Qty").Value
   vProductQty2 = Val(Grid.Columns("Qty2").Value)
   vProductID = Right("00000" + Grid.Columns("ID").Text, 5)
   vGroupID = Grid.Columns("GroupID").Text
   
   If vProductQty2 <> 0 Then
      If InStr(CStr(vProductQty2), ".") = 0 Then
         vStrBarcode = "0" & Right("999" + CStr(vProductQty2), 3) & vGroupID & vProductID
      Else
         vStrBarcode = "0" & Right("99" + Split(CStr(vProductQty2), ".")(0), 2) & "0" & Right("999" + Split(CStr(vProductQty2), ".")(1), 3) & vProductID
      End If
   Else
      vStrBarcode = "0110" & vGroupID & vProductID
   End If
End Sub
