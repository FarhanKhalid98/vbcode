VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form DefProducts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15390
   Icon            =   "DefProducts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   770
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1026
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtRackName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   141
      Top             =   8325
      Width           =   2415
   End
   Begin VB.CheckBox ChkSerial 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Serial"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10395
      TabIndex        =   140
      Top             =   9285
      Width           =   1005
   End
   Begin VB.CheckBox Chk3rdScheduleItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "3rd Schedule Item"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13050
      TabIndex        =   137
      Top             =   6705
      Width           =   1620
   End
   Begin VB.CheckBox ChkSearchNotClear 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Search Not Clear"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8640
      TabIndex        =   133
      Tag             =   "NC"
      Top             =   9285
      Width           =   1575
   End
   Begin VB.TextBox TxtDepartmentName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   129
      Top             =   7680
      Width           =   2415
   End
   Begin VB.TextBox TxtDepartmentID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   0
      Top             =   7680
      Width           =   600
   End
   Begin VB.CheckBox ChkDiscB4SaleTax 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4 Sale Tax"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13050
      TabIndex        =   125
      Top             =   6390
      Width           =   1920
   End
   Begin VB.CheckBox ChkDiscB4ExtraScheme 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4 Extra Scheme"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12060
      TabIndex        =   35
      Top             =   5625
      Width           =   2460
   End
   Begin VB.CheckBox ChkDiscB4TradeOffer 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Discount B4 Trade Offer"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12015
      TabIndex        =   33
      Top             =   4860
      Width           =   2460
   End
   Begin VB.TextBox TxtSeasonName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   12360
      TabIndex        =   118
      Top             =   1395
      Width           =   2415
   End
   Begin VB.CheckBox ChkDataNotClear 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Data Not Clear"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7028
      TabIndex        =   113
      Tag             =   "NC"
      Top             =   9285
      Width           =   1455
   End
   Begin VB.CheckBox ChkIsChangedPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Changed Price"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11288
      TabIndex        =   36
      Top             =   8640
      Width           =   1770
   End
   Begin VB.CheckBox ChkDeadProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Dead Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12683
      TabIndex        =   43
      Top             =   8910
      Width           =   1320
   End
   Begin VB.CheckBox ChkRawProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Raw Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11288
      TabIndex        =   42
      Top             =   8910
      Width           =   1365
   End
   Begin VB.CheckBox ChkExpiryDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11288
      TabIndex        =   44
      Top             =   7965
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkWSDiscb4ST 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Use WS Price For Discount b4 ST"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8370
      TabIndex        =   38
      Top             =   8640
      Width           =   2820
   End
   Begin VB.OptionButton OptRPSaleTax 
      BackColor       =   &H00EFC09E&
      Caption         =   "SaleTax"
      Height          =   195
      Left            =   8198
      TabIndex        =   81
      Top             =   6855
      Width           =   1020
   End
   Begin VB.OptionButton OptWSPSaleTax 
      BackColor       =   &H00EFC09E&
      Caption         =   "SaleTax"
      Height          =   195
      Left            =   8198
      TabIndex        =   80
      Top             =   6540
      Width           =   1020
   End
   Begin VB.ComboBox CmbSalePacking 
      Height          =   315
      Left            =   7103
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   5040
      Width           =   1740
   End
   Begin VB.CheckBox ChkNoCostProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "No Cost Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9788
      TabIndex        =   41
      Top             =   8895
      Width           =   1500
   End
   Begin VB.CheckBox ChkClosingProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Closing Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8363
      TabIndex        =   40
      Top             =   8895
      Width           =   1425
   End
   Begin VB.CheckBox ChkLockProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Lock Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7043
      TabIndex        =   39
      Top             =   8895
      Width           =   1320
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
      Height          =   5730
      Left            =   14985
      TabIndex        =   74
      Top             =   3330
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
         Height          =   5340
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   75
         Tag             =   "NC"
         Text            =   "DefProducts.frx":0ECA
         Top             =   330
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
   Begin VB.ComboBox CmbUnits 
      Height          =   315
      Left            =   7103
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5490
      Width           =   1740
   End
   Begin VB.TextBox TxtFilterID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2273
      TabIndex        =   51
      Top             =   2595
      Width           =   2655
   End
   Begin VB.TextBox TxtFilterProductName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2273
      TabIndex        =   50
      Top             =   2190
      Width           =   2655
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   7110
      TabIndex        =   19
      Top             =   6180
      Width           =   1080
      _ExtentX        =   1905
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
   Begin VB.ComboBox CmbPurPacking 
      Height          =   315
      Left            =   7103
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4620
      Width           =   1740
   End
   Begin VB.ComboBox CmbFilterGroup 
      Height          =   315
      Left            =   2273
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Tag             =   "nc"
      Top             =   1800
      Width           =   2670
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4095
      Left            =   1388
      TabIndex        =   52
      Top             =   3015
      Width           =   4035
      ScrollBars      =   2
      _Version        =   196616
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
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
      stylesets(0).Picture=   "DefProducts.frx":1050
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
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "GroupID"
      Columns(0).Name =   "GroupID"
      Columns(0).DataField=   "Column 2"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1667
      Columns(1).Caption=   "Product ID"
      Columns(1).Name =   "ID"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 0"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4392
      Columns(2).Caption=   "Name"
      Columns(2).Name =   "Name"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 1"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   7117
      _ExtentY        =   7223
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
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8318
      TabIndex        =   46
      Top             =   9765
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
      MICON           =   "DefProducts.frx":106C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9638
      TabIndex        =   47
      Top             =   9765
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
      MICON           =   "DefProducts.frx":1088
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid PGrid 
      Height          =   1095
      Left            =   9495
      TabIndex        =   13
      Top             =   4590
      Width           =   2280
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   3
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
      stylesets(0).Picture=   "DefProducts.frx":10A4
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
      Columns.Count   =   3
      Columns(0).Width=   2408
      Columns(0).Caption=   "Packing"
      Columns(0).Name =   "Packing"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   1032
      Columns(1).Caption=   "Mul"
      Columns(1).Name =   "Multiplier"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Packing ID"
      Columns(2).Name =   "PackingID"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   17
      Columns(2).FieldLen=   256
      _ExtentX        =   4022
      _ExtentY        =   1931
      _StockProps     =   79
      BackColor       =   15724527
      Enabled         =   0   'False
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
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7433
      TabIndex        =   57
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2541
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "DefProducts.frx":10C0
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtRetailPrice 
      Height          =   315
      Left            =   7110
      TabIndex        =   22
      Top             =   6810
      Width           =   1080
      _ExtentX        =   1905
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
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   6833
      TabIndex        =   4
      Top             =   2541
      Width           =   600
      _ExtentX        =   1058
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
      IntegralPoint   =   9
   End
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   9518
      TabIndex        =   14
      Top             =   5940
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   7103
      TabIndex        =   9
      Tag             =   "nc"
      Top             =   3795
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtName 
      Height          =   315
      Left            =   8228
      TabIndex        =   10
      Top             =   3780
      Width           =   3585
      _ExtentX        =   6324
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
      IntegralPoint   =   3
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid CGrid 
      Height          =   1095
      Left            =   9525
      TabIndex        =   48
      Top             =   6255
      Width           =   2295
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   3
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
      stylesets(0).Picture=   "DefProducts.frx":10DC
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
      Columns.Count   =   3
      Columns(0).Width=   3466
      Columns(0).Caption=   "Barcode"
      Columns(0).Name =   "Code"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "Packing ID"
      Columns(1).Name =   "PackingID"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   17
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Qty"
      Columns(2).Name =   "Qty"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   4048
      _ExtentY        =   1931
      _StockProps     =   79
      BackColor       =   15724527
      Enabled         =   0   'False
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
   Begin JeweledBut.JeweledButton BtnSubGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7433
      TabIndex        =   61
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2925
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "DefProducts.frx":10F8
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   6833
      TabIndex        =   5
      Top             =   2925
      Width           =   600
      _ExtentX        =   1058
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
      IntegralPoint   =   9
   End
   Begin JeweledBut.JeweledButton BtnCompany 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7433
      TabIndex        =   62
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2148
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "DefProducts.frx":1114
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   6833
      TabIndex        =   3
      Top             =   2148
      Width           =   600
      _ExtentX        =   1058
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
      IntegralPoint   =   9
   End
   Begin SITextBox.Txt TxtPurDisc 
      Height          =   315
      Left            =   7110
      TabIndex        =   23
      Top             =   7170
      Width           =   1080
      _ExtentX        =   1905
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
   Begin JeweledBut.JeweledButton BtnAddCompany 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10208
      TabIndex        =   68
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   2148
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
      MICON           =   "DefProducts.frx":1130
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnAddGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10208
      TabIndex        =   69
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   2541
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
      MICON           =   "DefProducts.frx":114C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnAddSubGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10208
      TabIndex        =   70
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   2925
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
      MICON           =   "DefProducts.frx":1168
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPacking 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8843
      TabIndex        =   71
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   4620
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
      MICON           =   "DefProducts.frx":1184
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnUnit 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8843
      TabIndex        =   72
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   5490
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
      MICON           =   "DefProducts.frx":11A0
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtMaxStockLimit 
      Height          =   315
      Left            =   9758
      TabIndex        =   29
      Top             =   7830
      Width           =   1080
      _ExtentX        =   1905
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
   Begin JeweledBut.JeweledButton BtnSalePacking 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8843
      TabIndex        =   78
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   5040
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
      MICON           =   "DefProducts.frx":11BC
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtWSPrice 
      Height          =   315
      Left            =   7110
      TabIndex        =   21
      Top             =   6495
      Width           =   1080
      _ExtentX        =   1905
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
   Begin SITextBox.Txt TxtSaleTaxPer 
      Height          =   315
      Left            =   13050
      TabIndex        =   37
      Top             =   6030
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Alignment       =   2
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
   Begin SITextBox.Txt TxtTokenVal 
      Height          =   315
      Left            =   7110
      TabIndex        =   27
      Top             =   8475
      Width           =   1080
      _ExtentX        =   1905
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpExpiryDate 
      Height          =   375
      Left            =   11280
      TabIndex        =   45
      Tag             =   "NC"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65543
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin SITextBox.Txt TxtDesc1 
      Height          =   315
      Left            =   7073
      TabIndex        =   12
      Top             =   4185
      Width           =   4740
      _ExtentX        =   8361
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtSaleDiscPer 
      Height          =   315
      Left            =   7110
      TabIndex        =   25
      Top             =   7800
      Width           =   1080
      _ExtentX        =   1905
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
   Begin JeweledBut.JeweledButton BtnBrand 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7433
      TabIndex        =   92
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3330
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "DefProducts.frx":11D8
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBrandID 
      Height          =   315
      Left            =   6833
      TabIndex        =   6
      Top             =   3330
      Width           =   600
      _ExtentX        =   1058
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
      IntegralPoint   =   9
   End
   Begin JeweledBut.JeweledButton BtnAddBrand 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10208
      TabIndex        =   93
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   3330
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
      MICON           =   "DefProducts.frx":11F4
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   7793
      TabIndex        =   95
      Top             =   2541
      Width           =   2415
      _ExtentX        =   4260
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
   End
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   7793
      TabIndex        =   96
      Top             =   2925
      Width           =   2415
      _ExtentX        =   4260
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
   End
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   7793
      TabIndex        =   97
      Top             =   2148
      Width           =   2415
      _ExtentX        =   4260
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
   End
   Begin SITextBox.Txt TxtBrandName 
      Height          =   315
      Left            =   7793
      TabIndex        =   98
      Top             =   3330
      Width           =   2415
      _ExtentX        =   4260
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
   End
   Begin JeweledBut.JeweledButton BtnProductID 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6698
      TabIndex        =   100
      TabStop         =   0   'False
      ToolTipText     =   "Add New"
      Top             =   3795
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
      TX              =   "Miss"
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
      MICON           =   "DefProducts.frx":1210
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   11288
      TabIndex        =   15
      Top             =   5940
      Visible         =   0   'False
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin JeweledBut.JeweledButton BtnMulti 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11828
      TabIndex        =   102
      TabStop         =   0   'False
      ToolTipText     =   "Add New"
      Top             =   3780
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
      TX              =   "Multi"
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
      MICON           =   "DefProducts.frx":122C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7433
      TabIndex        =   103
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1755
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "DefProducts.frx":1248
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   6833
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   1755
      Width           =   600
      _ExtentX        =   1058
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
      IntegralPoint   =   9
   End
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   7793
      TabIndex        =   104
      Tag             =   "NC"
      Top             =   1755
      Width           =   2415
      _ExtentX        =   4260
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
   End
   Begin JeweledBut.JeweledButton BtnAddOrganization 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10208
      TabIndex        =   106
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   1755
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
      MICON           =   "DefProducts.frx":1264
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtMinStockLimit 
      Height          =   315
      Left            =   9758
      TabIndex        =   28
      Top             =   7515
      Width           =   1080
      _ExtentX        =   1905
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
   Begin SITextBox.Txt TxtSaleDisc 
      Height          =   315
      Left            =   7110
      TabIndex        =   24
      Top             =   7485
      Width           =   1080
      _ExtentX        =   1905
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
   Begin SITextBox.Txt TxtEmpComm 
      Height          =   315
      Left            =   7110
      TabIndex        =   26
      Top             =   8115
      Width           =   1080
      _ExtentX        =   1905
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
   Begin SITextBox.Txt TxtServiceCharges 
      Height          =   315
      Left            =   9758
      TabIndex        =   30
      Top             =   8145
      Width           =   1080
      _ExtentX        =   1905
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
   Begin JeweledBut.JeweledButton BtnPub 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7433
      TabIndex        =   109
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1365
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "DefProducts.frx":1280
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPubID 
      Height          =   315
      Left            =   6833
      TabIndex        =   1
      Top             =   1365
      Width           =   600
      _ExtentX        =   1058
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
      IntegralPoint   =   9
   End
   Begin JeweledBut.JeweledButton BtnAddPub 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10208
      TabIndex        =   110
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   1365
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
      MICON           =   "DefProducts.frx":129C
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPubName 
      Height          =   315
      Left            =   7793
      TabIndex        =   111
      Top             =   1365
      Width           =   2415
      _ExtentX        =   4260
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
   End
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   1793
      TabIndex        =   114
      Top             =   9765
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "New"
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
      MICON           =   "DefProducts.frx":12B8
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   3113
      TabIndex        =   115
      Top             =   9765
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Change"
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
      MICON           =   "DefProducts.frx":12D4
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   4433
      TabIndex        =   116
      Top             =   9765
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
      MICON           =   "DefProducts.frx":12F0
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7013
      TabIndex        =   117
      Top             =   9765
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
      MICON           =   "DefProducts.frx":130C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSeason 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   12000
      TabIndex        =   119
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1395
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "DefProducts.frx":1328
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSeasonID 
      Height          =   315
      Left            =   11400
      TabIndex        =   8
      Top             =   1395
      Width           =   600
      _ExtentX        =   1058
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
      IntegralPoint   =   9
   End
   Begin JeweledBut.JeweledButton BtnAddSeason 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   14775
      TabIndex        =   120
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   1395
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
      MICON           =   "DefProducts.frx":1344
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtTradeOffer2 
      Height          =   315
      Left            =   13815
      TabIndex        =   32
      Top             =   4500
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
      Left            =   13050
      TabIndex        =   31
      Top             =   4500
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
      Left            =   13635
      TabIndex        =   34
      Top             =   5220
      Width           =   855
      _ExtentX        =   1508
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
   Begin SITextBox.Txt TxtBottomPrice 
      Height          =   315
      Left            =   12840
      TabIndex        =   126
      Top             =   8160
      Width           =   1095
      _ExtentX        =   1931
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
   Begin JeweledBut.JeweledButton BtnAddDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4815
      TabIndex        =   130
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   7680
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
      MICON           =   "DefProducts.frx":1360
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   131
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7680
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "DefProducts.frx":137C
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtListPrice 
      Height          =   315
      Left            =   8280
      TabIndex        =   20
      Top             =   6180
      Width           =   1080
      _ExtentX        =   1905
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
   Begin SITextBox.Txt TxtPCTCode 
      Height          =   315
      Left            =   13050
      TabIndex        =   138
      Top             =   7200
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Alignment       =   2
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
      Masked          =   1
      IntegralPoint   =   8
   End
   Begin JeweledBut.JeweledButton BtnRack 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   142
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   8325
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "DefProducts.frx":1398
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtRackID 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   8325
      Width           =   600
      _ExtentX        =   1058
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
      IntegralPoint   =   9
   End
   Begin JeweledBut.JeweledButton BtnAddRack 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4815
      TabIndex        =   143
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   8325
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
      MICON           =   "DefProducts.frx":13B4
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblRack 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rack"
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
      Left            =   1440
      TabIndex        =   144
      Top             =   8100
      Width           =   465
   End
   Begin VB.Label LblPCTCode 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PCT Code"
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
      Left            =   13050
      TabIndex        =   139
      Top             =   6975
      Width           =   870
   End
   Begin VB.Label LblListPrice 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Price"
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
      Left            =   8280
      TabIndex        =   136
      Top             =   5880
      Width           =   810
   End
   Begin VB.Label LblGroupName1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name Urdu"
      Height          =   195
      Left            =   10680
      TabIndex        =   135
      Top             =   2280
      Width           =   1290
   End
   Begin MSForms.TextBox TxtGroupName1 
      Height          =   435
      Left            =   10680
      TabIndex        =   134
      ToolTipText     =   "Textbox1"
      Top             =   2520
      Width           =   2985
      VariousPropertyBits=   752896027
      ForeColor       =   0
      BorderStyle     =   1
      Size            =   "5265;767"
      SpecialEffect   =   0
      FontName        =   "@Arial Unicode MS"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label LblDepartment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Left            =   1440
      TabIndex        =   132
      Top             =   7440
      Width           =   990
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bottom Price"
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
      Left            =   12795
      TabIndex        =   128
      Top             =   7965
      Width           =   1110
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date"
      Enabled         =   0   'False
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
      Left            =   11513
      TabIndex        =   127
      Top             =   7965
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label LblExtraSchemePer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Scheme %"
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
      Left            =   12060
      TabIndex        =   124
      Top             =   5310
      Width           =   1380
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
      Left            =   13680
      TabIndex        =   123
      Top             =   4545
      Width           =   120
   End
   Begin VB.Label LblTradeOffer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Offer"
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
      Left            =   12015
      TabIndex        =   122
      Top             =   4590
      Width           =   990
   End
   Begin VB.Label LblSeason 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Season"
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
      Left            =   10665
      TabIndex        =   121
      Top             =   1440
      Width           =   645
   End
   Begin VB.Label LblPublisher 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Writer / Pub"
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
      Left            =   5693
      TabIndex        =   112
      Top             =   1410
      Width           =   1065
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Charges"
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
      Left            =   8288
      TabIndex        =   108
      Top             =   8190
      Width           =   1410
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min Stock Limit"
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
      Left            =   8363
      TabIndex        =   107
      Top             =   7560
      Width           =   1320
   End
   Begin VB.Label LblOrganization 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization"
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
      Left            =   5693
      TabIndex        =   105
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Enabled         =   0   'False
      Height          =   195
      Left            =   11288
      TabIndex        =   101
      Top             =   5715
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label LblName2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name Urdu"
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
      Left            =   10680
      TabIndex        =   99
      Top             =   3120
      Width           =   960
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   435
      Left            =   10665
      TabIndex        =   11
      ToolTipText     =   "Textbox1"
      Top             =   3300
      Width           =   4425
      VariousPropertyBits=   752896027
      ForeColor       =   0
      BorderStyle     =   1
      Size            =   "7805;767"
      SpecialEffect   =   0
      FontName        =   "@Arial Unicode MS"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brands"
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
      Left            =   6173
      TabIndex        =   94
      Top             =   3360
      Width           =   600
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Disc%"
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
      Left            =   6090
      TabIndex        =   91
      Top             =   7875
      Width           =   960
   End
   Begin VB.Label LblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   6023
      TabIndex        =   90
      Top             =   4290
      Width           =   975
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Comm Rs"
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
      Left            =   5835
      TabIndex        =   89
      Top             =   8160
      Width           =   1230
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Price"
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
      Left            =   6240
      TabIndex        =   88
      Top             =   6225
      Width           =   810
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retail Price"
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
      Left            =   6030
      TabIndex        =   87
      Top             =   6855
      Width           =   1020
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Disc/Pc"
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
      Left            =   5910
      TabIndex        =   86
      Top             =   7530
      Width           =   1140
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Disc/Pc"
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
      Left            =   6000
      TabIndex        =   85
      Top             =   7215
      Width           =   1050
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WS Price"
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
      Left            =   6240
      TabIndex        =   84
      Top             =   6540
      Width           =   810
   End
   Begin VB.Label LblSaleTaxPer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Tax%"
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
      Left            =   12090
      TabIndex        =   83
      Top             =   6075
      Width           =   900
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Token Val"
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
      Left            =   6150
      TabIndex        =   82
      Top             =   8520
      Width           =   900
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Packing"
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
      Left            =   5903
      TabIndex        =   79
      Top             =   5070
      Width           =   1140
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max Stock Limit"
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
      Left            =   8363
      TabIndex        =   77
      Top             =   7830
      Width           =   1380
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
      Left            =   14760
      TabIndex        =   73
      Top             =   225
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Units"
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
      Left            =   6593
      TabIndex        =   67
      Top             =   5520
      Width           =   450
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Products"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   66
      Top             =   270
      Width           =   1245
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group"
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
      Left            =   5858
      TabIndex        =   65
      Top             =   2955
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
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
      Left            =   5753
      TabIndex        =   64
      Top             =   3840
      Width           =   930
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
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
      Left            =   5993
      TabIndex        =   63
      Top             =   2175
      Width           =   780
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code :"
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
      Left            =   1658
      TabIndex        =   60
      Top             =   2640
      Width           =   570
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Codes"
      Enabled         =   0   'False
      Height          =   195
      Left            =   9518
      TabIndex        =   59
      Top             =   5715
      Width           =   735
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   14310
      Top             =   -45
      Width           =   345
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
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
      Left            =   6248
      TabIndex        =   58
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Packing"
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
      Left            =   5993
      TabIndex        =   56
      Top             =   4650
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   372.533
      X2              =   372.533
      Y1              =   145
      Y2              =   473
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      Left            =   1613
      TabIndex        =   55
      Top             =   2235
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group :"
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
      Left            =   1583
      TabIndex        =   54
      Top             =   1890
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   7733
      TabIndex        =   53
      Top             =   3855
      Width           =   495
   End
End
Attribute VB_Name = "DefProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.
Dim RsPacking As New ADODB.Recordset
Dim RsProductPacking As New ADODB.Recordset
Dim RsCode As New ADODB.Recordset
Dim Flag As Boolean, isFreeCode As Boolean, ModeValue As Boolean
Dim vPer As Byte, Item As ListItem
Dim vProductID As String
Dim vCompanyID As String
Dim vCounter As Integer
Dim vMaxBinID As Integer
Dim UniCode As Variant
Dim vPurchasePrice As Double
Dim vStrSQL As String

Private Sub BtnAddItemDesc_Click()

End Sub

Private Sub BtnAddOrganization_Click()
On Error GoTo ErrorHandler
'   DefOrganization.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddBrand_Click()
On Error GoTo ErrorHandler
   DefBrands.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddCompany_Click()
On Error GoTo ErrorHandler
   DefCompanies.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddDepartment_Click()
On Error GoTo ErrorHandler
   DefDepartments.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddGroup_Click()
On Error GoTo ErrorHandler
   DefGroups.Show vbModal
   With CN.Execute("Select * FROM Groups")
      CmbFilterGroup.Clear
      CmbFilterGroup.AddItem ""
      Do Until .EOF
         CmbFilterGroup.AddItem !GroupName
         CmbFilterGroup.ItemData(CmbFilterGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
    End With
    Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddPub_Click()
On Error GoTo ErrorHandler
   DefPublishers.Show vbModal
   With CN.Execute("Select * FROM Publishers")
      CmbFilterGroup.Clear
      CmbFilterGroup.AddItem ""
      Do Until .EOF
         CmbFilterGroup.AddItem !PubName
         CmbFilterGroup.ItemData(CmbFilterGroup.NewIndex) = Asc(Left(!PubID, 1)) & Asc(Mid(!PubID, 2, 1)) & Asc(Mid(!PubID, 3, 1))
         .MoveNext
      Loop
    End With
    Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddSubDeapartment_Click()
On Error GoTo ErrorHandler
   DefSubDepartment.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddRack_Click()
On Error GoTo ErrorHandler
   DefRacks.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddSubGroup_Click()
On Error GoTo ErrorHandler
   DefSubGroups.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnBrand_Click()
On Error GoTo ErrorHandler
   If FunSelectBrand(ssButton, False) = True Then
      If TxtSeasonID.Visible Then TxtSeasonID.SetFocus Else TxtID.SetFocus
   Else
      If TxtBrandID.Enabled Then TxtBrandID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnCompany_Click()
On Error GoTo ErrorHandler
   If FunSelectCompany(ssButton, False) = True Then
     If TxtGroupID.Enabled Then TxtGroupID.SetFocus
    Else
     If TxtCompanyID.Enabled Then TxtCompanyID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnGroup_Click()
On Error GoTo ErrorHandler
   If FunSelectGroup(ssButton, False) = True Then
      If TxtSubGroupID.Enabled Then TxtSubGroupID.SetFocus
   Else
      If TxtGroupID.Enabled Then TxtGroupID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnMulti_Click()
On Error GoTo ErrorHandler
   DefMultiProduct.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOrganization_Click()
On Error GoTo ErrorHandler
 If FunSelectOrganization(ssButton, False) = True Then
     If TxtGroupID.Enabled Then TxtGroupID.SetFocus
    Else
     If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPacking_Click()
On Error GoTo ErrorHandler
   DefPackings.Show vbModal
   With CN.Execute("Select * FROM Packings")
      CmbPurPacking.Clear
      CmbPurPacking.AddItem ""
      Do Until .EOF
          CmbPurPacking.AddItem !PackingName
          CmbPurPacking.ItemData(CmbPurPacking.NewIndex) = !PackingID
          .MoveNext
      Loop
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProductID_Click()
On Error GoTo ErrorHandler
   'FunGetMaxID = CN.Execute("Select right('0000' + cast(isnull(max(cast(substring(ProductId,3,10) as smallint)),0) + 1 as varchar),4) from Products").Fields(0) ' Where ProductId like '" & GetGroupID(CmbCompany) & "%'").Fields(0)
   isFreeCode = True
   BtnProductID.Enabled = False
   TxtID.Text = FunGetMaxID
   PopulatePackGrid
   PopulateCodeGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPub_Click()
On Error GoTo ErrorHandler
   If FunSelectPublisher(ssButton, False) = True Then
      If TxtGroupID.Enabled Then TxtGroupID.SetFocus
   Else
      If TxtPubID.Enabled Then TxtPubID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectPublisher(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchPublishers.Show vbModal, Me
        If SchPublishers.ParaOutPubID = "" Then FunSelectPublisher = False: Exit Function
        TxtPubID.Text = SchPublishers.ParaOutPubID
    End If
    '---------------------------
    TxtPubID.Text = Right("000" + CStr(Val(TxtPubID.Text)), 3)
    vStrSQL = " Select * FROM Publishers where PubID='" & TxtPubID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtPubName.Text = !PubName
          'If vIsNewRecord = True Then TxtID.Text = FunGetMaxID
          FunSelectPublisher = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectPublisher = False
          .Close
          TxtPubID.Text = ""
          TxtPubName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnSalePacking_Click()
On Error GoTo ErrorHandler
   DefPackings.Show vbModal
   With CN.Execute("Select * FROM Packings")
      CmbSalePacking.Clear
      CmbSalePacking.AddItem ""
      Do Until .EOF
          CmbSalePacking.AddItem !PackingName
          CmbSalePacking.ItemData(CmbSalePacking.NewIndex) = !PackingID
          .MoveNext
      Loop
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSeason_Click()
   On Error GoTo ErrorHandler
   If FunSelectSeason(ssButton, False) = True Then
     If TxtID.Visible Then TxtID.SetFocus Else TxtSeasonID.SetFocus
   Else
      TxtSeasonID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub ChkSerial_Click()
On Error GoTo ErrorHandler
   If ActiveControl.Name <> ChkSerial.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtRackID_Change()
   On Error GoTo ErrorHandler
   If TxtRackID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtRackID.Name Then Exit Sub
   If TxtRackName.Text <> "" Then TxtRackName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtRackID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtRackID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtRackID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectRack(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectRack(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectRack(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchRack.Show vbModal, Me
        If SchRack.ParaOutRackID = "" Then FunSelectRack = False: Exit Function
        TxtRackID.Text = SchRack.ParaOutRackID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Racks where RackID=" & Val(TxtRackID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtRackName.Text = !RackName
          FunSelectRack = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectRack = False
          .Close
          TxtRackID.Text = ""
          TxtRackName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Sub BtnRack_Click()
   On Error GoTo ErrorHandler
   If FunSelectRack(ssButton, False) = True Then
     If TxtBrandID.Visible Then TxtBrandID.SetFocus Else TxtRackID.SetFocus
   Else
      TxtRackID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSeasonID_Change()
   On Error GoTo ErrorHandler
   If TxtSeasonID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSeasonID.Name Then Exit Sub
   If TxtSeasonName.Text <> "" Then TxtSeasonName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSeasonID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSeasonID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSeasonID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSeason(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSeason(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSeason(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSeasons.Show vbModal, Me
        If SchSeasons.ParaOutSeasonID = "" Then FunSelectSeason = False: Exit Function
        TxtSeasonID.Text = SchSeasons.ParaOutSeasonID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Seasons where SeasonID=" & Val(TxtSeasonID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSeasonName.Text = !SeasonName
          FunSelectSeason = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSeason = False
          .Close
          TxtSeasonID.Text = ""
          TxtSeasonName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnAddSeason_Click()
   On Error GoTo ErrorHandler
   DefSeason.Show vbModal, Me
   If TxtSeasonID.Visible Then TxtSeasonID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub BtnSubGroup_Click()
On Error GoTo ErrorHandler
   If FunSelectSubGroup(ssButton, False) = True Then
      If TxtBrandID.Enabled Then TxtBrandID.SetFocus
   Else
      If TxtSubGroupID.Enabled Then TxtSubGroupID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnUnit_Click()
On Error GoTo ErrorHandler
   DefUnits.Show vbModal
   With CN.Execute("Select * FROM Units")
      CmbUnits.Clear
      CmbUnits.AddItem ""
      Do Until .EOF
         CmbUnits.AddItem !UnitName
         CmbUnits.ItemData(CmbUnits.NewIndex) = !UnitID
         .MoveNext
      Loop
    End With
    Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectBrand(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchBrand.Show vbModal, Me
        If SchBrand.ParaOutBrandID = "" Then FunSelectBrand = False: Exit Function
        TxtBrandID.Text = SchBrand.ParaOutBrandID
    End If
    '---------------------------
    vStrSQL = "Select * FROM Brands where BrandID = " & Val(TxtBrandID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtBrandName.Text = !BrandName
          FunSelectBrand = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectBrand = False
          .Close
          TxtBrandID.Text = ""
          TxtBrandName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchGroup.Show vbModal, Me
        If SchGroup.ParaOutGroupID = "" Then FunSelectGroup = False: Exit Function
        TxtGroupID.Text = SchGroup.ParaOutGroupID
    End If
    '---------------------------
    TxtGroupID.Text = Right("000" + CStr(Val(TxtGroupID.Text)), 3)
    vStrSQL = " Select * FROM Groups where GroupID='" & TxtGroupID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtGroupName.Text = !GroupName
          TxtGroupName1.Text = IIf(IsNull(!GroupName1), "", !GroupName1)
          If ChkSearchNotClear.Value = 0 Then CmbFilterGroup.Text = !GroupName
          'If vIsNewRecord = True Then TxtID.Text = FunGetMaxID
          FunSelectGroup = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectGroup = False
          .Close
          TxtGroupID.Text = ""
          TxtGroupName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   If Err.Number = 383 Then
      CmbFilterGroup.AddItem ""
      With CN.Execute("Select * FROM Groups")
           Do Until .EOF
               CmbFilterGroup.AddItem !GroupName
               CmbFilterGroup.ItemData(CmbFilterGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
               .MoveNext
           Loop
       End With
       Resume
   End If
   Call ShowErrorMessage
End Function

Private Function FunSelectSubGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSubGroup.Show vbModal, Me
        If SchSubGroup.ParaOutSubGroupID = "" Then FunSelectSubGroup = False: Exit Function
        TxtSubGroupID.Text = SchSubGroup.ParaOutSubGroupID
    End If
    '---------------------------
    vStrSQL = " Select * FROM SubGroups where SubGroupID=" & Val(TxtSubGroupID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSubGroupName.Text = !SubGroupName
          FunSelectSubGroup = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSubGroup = False
          .Close
          TxtSubGroupID.Text = ""
          TxtSubGroupName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectCompany(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCompany.Show vbModal, Me
        If SchCompany.ParaOutCompanyID = "" Then FunSelectCompany = False: Exit Function
        TxtCompanyID.Text = SchCompany.ParaOutCompanyID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Companies where CompanyID=" & Val(TxtCompanyID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCompanyName.Text = !CompanyName
          FunSelectCompany = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectCompany = False
          .Close
          TxtCompanyID.Text = ""
          TxtCompanyName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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

Private Sub BtnAddVendor_Click()
On Error GoTo ErrorHandler
   DefVendors.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CGrid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   RsCode.Filter = "Code = '" & CGrid.Columns("Code").Text & "'"
   If RsCode.RecordCount = 1 And CGrid.Columns("Code").Text <> "" Then
      If vIsNewRecord = False Then CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Deleted BarCode-" & CGrid.Columns("Code").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
      RsCode.Delete
   End If
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CGrid_DblClick()
On Error GoTo ErrorHandler
   Call CGrid_LostFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CGrid_GotFocus()
On Error GoTo ErrorHandler
   Flag = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CGrid_LostFocus()
On Error GoTo ErrorHandler
   Flag = False
'   If Trim(Grid.Columns("Code").Text) = "" Then
'      TxtCode.Text = ""
'      TxtCode.Enabled = True
'      TxtCode.SetFocus
'   Else
'      TxtCode.Enabled = False
'      CmbPackName.SetFocus
'   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CGrid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   If Flag Then
      TxtCode.Text = CGrid.Columns("Code").Text
      TxtQty.Text = CGrid.Columns("Qty").Value
      If CGrid.Rows = 1 Then CGrid.MoveLast
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Chk3rdScheduleItem_Click()
On Error GoTo ErrorHandler
   If ActiveControl.Name <> Chk3rdScheduleItem.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkClosingProduct_Click()
On Error GoTo ErrorHandler
   If ActiveControl.Name <> ChkClosingProduct.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkLockProduct_Click()
On Error GoTo ErrorHandler
   If ActiveControl.Name <> ChkLockProduct.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbFilterGroup_GotFocus()
On Error GoTo ErrorHandler
   CmbFilterGroup_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbPurPacking_Click()
On Error GoTo ErrorHandler
   If CmbPurPacking.Visible = False Then Exit Sub
   If Me.ActiveControl.Name <> CmbPurPacking.Name Then Exit Sub
   If ObjRegistry.AllowBothPackingsareSame Then CmbSalePacking.Text = CmbPurPacking.Text
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbSalePacking_Click()
On Error GoTo ErrorHandler
   If CmbSalePacking.Visible = False Then Exit Sub
   If Me.ActiveControl.Name <> CmbSalePacking.Name Then Exit Sub
   If ObjRegistry.AllowBothPackingsareSame Then CmbPurPacking.Text = CmbSalePacking.Text
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbUnits_Click()
On Error GoTo ErrorHandler
   If CmbUnits.Visible = False Then Exit Sub
   If Me.ActiveControl.Name <> CmbUnits.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   On Error GoTo ErrorHandler
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
          Set frmObj = Nothing
      Next
      Set Rs = Nothing
      Set Rs1 = Nothing
      Set RsPacking = Nothing
      Set RsProductPacking = Nothing
      Set DefProducts = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub FraHelp_Click()
On Error GoTo ErrorHandler
   FraHelp.Visible = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
On Error GoTo ErrorHandler
   If Grid.Rows > 0 And BtnOpen.Enabled Then BtnOpen_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbFilterGroup_Click()
   On Error GoTo ErrorHandler
   Set Rs1 = New ADODB.Recordset
   If CmbFilterGroup.ListIndex <= 0 Then
       Rs1.Open "Select * FROM Products Order By ProductName", CN, adOpenStatic, adLockOptimistic
   Else
       Rs1.Open "Select * FROM Products Where GroupID = '" & GetGroupID(CmbFilterGroup) & "' Order By ProductName", CN, adOpenStatic, adLockOptimistic
   End If
   Set Grid.DataSource = Rs1
   Grid.Columns("ID").DataField = "ProductID"
   Grid.Columns("Name").DataField = "ProductName"
   Grid.Columns("GroupID").DataField = "GroupID"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      KeyCode = 0
   ElseIf KeyCode = vbKeyReturn Then
      If ActiveControl.Name = CGrid.Name Then
         CGrid_DblClick
      Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
    
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyU
            ModeValue = False
            KeyCode = 0
         Case vbKeyE
            ModeValue = True
            KeyCode = 0
         Case vbKeyS
             If ActiveControl.Name = PGrid.Name Then UpdatePacking
             If BtnSave.Enabled Then BtnSave_Click
             KeyCode = 0
         Case vbKeyW
             If BtnClear.Enabled Then BtnClear_Click
             KeyCode = 0
         Case vbKeyQ
             If BtnClose.Enabled Then BtnClose_Click
             KeyCode = 0
         Case vbKeyN
             If BtnNew.Enabled Then BtnNew_Click
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
      End Select
   ElseIf KeyCode = vbKeyF12 Then
      PGrid.SetFocus
      KeyCode = 0
   ElseIf KeyCode = vbKeyF6 Then
      TxtName.Text = TxtName.Text & IIf(Len(TxtName.Text) = 0, "", " ") & TxtCompanyName.Text  '+ """"
      keybd_event vbKeyEnd, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF7 Then
      TxtName.Text = TxtName.Text & IIf(Len(TxtName.Text) = 0, "", " ") & TxtGroupName.Text '+ """"
      keybd_event vbKeyEnd, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF8 Then
      TxtName.Text = TxtName.Text & IIf(Len(TxtName.Text) = 0, "", " ") & TxtSubGroupName.Text '+ """"
      keybd_event vbKeyEnd, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF9 Then
      vPer = Val(InputBox("Add Percentage in Purchase Price to Convert Retail Price", "Input", vPer))
      KeyCode = 0
   ElseIf KeyCode = vbKeyDelete Then
      If ActiveControl.Name = CGrid.Name Then
         If CGrid.Columns(0).Text <> "" Then
            TxtCode.Text = ""
            CGrid.SelBookmarks.RemoveAll
            CGrid.SelBookmarks.Add CGrid.Bookmark
            CGrid.DeleteSelected
            CGrid.Refresh
         End If
      End If
    ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtPubID.Name: If FunSelectPublisher(ssFunctionKey, True) = True Then If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus Else If TxtPubID.Enabled Then TxtPubID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, True) = True Then If TxtCompanyID.Enabled Then TxtCompanyID.SetFocus Else If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus
         Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then If TxtGroupID.Enabled Then TxtGroupID.SetFocus Else If TxtCompanyID.Enabled Then TxtCompanyID.SetFocus
         Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then If TxtSubGroupID.Enabled Then TxtSubGroupID.SetFocus Else If TxtGroupID.Enabled Then TxtGroupID.SetFocus
         Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then If TxtBrandID.Enabled Then TxtBrandID.SetFocus Else If TxtSubGroupID.Enabled Then TxtSubGroupID.SetFocus
         Case TxtBrandID.Name: If FunSelectBrand(ssFunctionKey, True) = True Then If TxtID.Enabled Then TxtID.SetFocus Else If TxtBrandID.Enabled Then TxtBrandID.SetFocus
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) = UCase(TxtFilterProductName.Name) Or UCase(Me.ActiveControl.Name) = UCase(TxtFilterID.Name) Then Exit Sub
      If UCase(Me.ActiveControl.Name) Like "TXT*" Or UCase(Me.ActiveControl.Name) Like "TEXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
On Error GoTo ErrorHandler
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Set Rs1 = New ADODB.Recordset
   Rs1.Open "Select * FROM Products Order By ProductName", CN, adOpenStatic, adLockOptimistic
   If ChkSearchNotClear.Value = 0 Then
    Set Grid.DataSource = Rs1
   End If
   FormStatus = SelectionMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
On Error GoTo ErrorHandler
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Unload Me
  Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniProducts", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
  Dim vtbl As String
  Dim vProductID As String
  If Rs1.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
    Dim vid As String
    vid = Rs1!Productid
    vtbl = Common.ChildDataExists("Products", "ProductId='" & vid & "'", "ProductBarCodes,ProductPacking", "ProductID")
    If vtbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
      Exit Sub
    End If
    
'    vMaxBinID = FunGetMaxBinID
'    ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'    CN.Execute ("Insert Into Bin_Products Select " & vMaxBinID & ",'" & Date & "',* from Products Where productID = " & TxtID.Text)
'    CN.Execute ("Insert Into Bin_ProductBarCodes Select " & vMaxBinID & ",'" & Date & "',* from ProductBarCodes Where productID = " & TxtID.Text)
'    CN.Execute ("Insert Into Bin_ProductPacking Select " & vMaxBinID & ",'" & Date & "',* from ProductPacking Where productID = " & TxtID.Text)
'
'    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    '---------------------------------------------
    Call ActivityLogBin("", eFrmProducts, eDelete, TxtID.Text, SystemDate, "Updated Code-" & TxtID.Text & " Pur Price-" & Val(TxtPurPrice.Text) & " WSPrice-" & Val(TxtWSPrice.Text) & " Retail Price-" & Val(TxtRetailPrice.Text) & " Pur Disc-" & Val(TxtPurDisc.Text) & " Sale Disc-" & Val(TxtSaleDisc.Text) & " Emp Comm-" & Val(TxtEmpComm.Text))
'    Call ActivityLog("Products", eDelete, , , vid)
    
    vProductID = TxtID.Text  'TxtPrefix.Text & TxtID.Text
    If CN.Execute("Select * from ProductPacking where ProductID='" & vProductID & "'").RecordCount > 0 Then
       CN.Execute ("Delete from ProductPacking where ProductID='" & vProductID & "'")
    End If
    If CN.Execute("Select * from ProductBarcodes where ProductID='" & vProductID & "'").RecordCount > 0 Then
       CN.Execute ("Delete from ProductBarCodes where ProductID='" & vProductID & "'")
    End If

    Rs1.Delete
    Rs1.ReQuery
    Rs.ReQuery
    PopulateCodeGrid
    PopulatePackGrid
    '---------------------------------------------
    If Rs1.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
    Rs1.MoveNext
    Grid.MoveNext
    If Rs1.EOF Then Rs1.MoveLast
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnNew_Click()
On Error GoTo ErrorHandler
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
  On Error GoTo ErrorHandler
  If Rs1.RecordCount > 0 Then
    If Rs1.BOF = False And Rs1.EOF = False Then
      FormStatus = OpenMode
    End If
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function GetGroupID(cmb As ComboBox) As String
    On Error GoTo ErrorHandler
    If cmb.ListIndex = -1 Then Exit Function
    GetGroupID = Chr(Left(cmb.ItemData(cmb.ListIndex), 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 3, 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 5, 2))
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If FunValidation = False Then Exit Sub
    
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniProducts", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEditDefination = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
''''''''''''' Rizwan
      
      vPurchasePrice = Val(TxtPurPrice.Text)
      If CmbPurPacking.Text <> "" Then
         RsProductPacking.Filter = "ProductID='" & TxtID.Text & "' and PackingID=" & CmbPurPacking.ItemData(CmbPurPacking.ListIndex)
         If RsProductPacking.RecordCount > 0 Then vPurchasePrice = Round(Val(TxtPurPrice.Text) / IIf(RsProductPacking!Multiplier = 0, 1, RsProductPacking!Multiplier), 2)
      End If
      
''''''''''''''''''''

   If Val(TxtRetailPrice.Text) < Val(vPurchasePrice) Then
      If MsgBox("Retail Price ( " & TxtRetailPrice.Text & " ) is Less Than Purchase Price ( " & vPurchasePrice & " ). Do you want to change it?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbYes Then
         TxtRetailPrice.SetFocus
         Exit Sub
      End If
   End If
   'Rs.Filter = ""
   'Rs.ReQuery
   CN.BeginTrans
   
   
'vStrPara = ""

'vStrPara = vStrPara & Val(TxtSID.Text) & "," 'SID
'vStrPara = vStrPara & TxtName.Text & "," 'ProductName
'vStrPara = vStrPara & Val(TxtPurPrice.Text) & "," 'Purchase Price
'vStrPara = vStrPara &  "0," 'is Changed
'vStrPara = vStrPara & TxtName.Text & "," 'ProductName
'vStrPara = vStrPara & IIf(Trim(TextBox1.Text) = "", Null, TextBox1.Text) & "," 'Product Name1 as Urdu Name
'vStrPara = vStrPara & IIf(Trim(TxtCompanyID.Text) = "", Null, TxtCompanyID.Text) & "," 'Company ID
'vStrPara = vStrPara & IIf(Trim(TxtOrganizationID.Text) = "", Null, TxtOrganizationID.Text) & "," 'Organization ID
'vStrPara = vStrPara & TxtGroupID.Text & "," 'Group ID
'vStrPara = vStrPara & IIf(Trim(TxtSubGroupID.Text) = "", Null, TxtSubGroupID.Text) & "," 'Sub Group ID
'vStrPara = vStrPara & IIf(Trim(TxtBrandID.Text) = "", Null, TxtBrandID.Text) & "," 'Brand ID
'vStrPara = vStrPara & IIf(Trim(TxtSeasonID.Text) = "", Null, TxtSeasonID.Text) & "," 'Season ID
'vStrPara = vStrPara & IIf(Trim(TxtDepartmentID.Text) = "", Null, TxtDepartmentID.Text) & "," 'Department ID
'vStrPara = vStrPara & IIf(Trim(TxtPubID.Text) = "", Null, TxtPubID.Text) & "," 'Pub ID
'vStrPara = vStrPara & IIf(CmbPurPacking.Text = "", Null, CmbPurPacking.ItemData(CmbPurPacking.ListIndex)) & "," 'Purchase Packing ID
'vStrPara = vStrPara & IIf(CmbSalePacking.Text = "", Null, CmbSalePacking.ItemData(CmbSalePacking.ListIndex)) & "," 'Sale Packing ID
'vStrPara = vStrPara & IIf(CmbUnits.Text = "", Null, CmbUnits.ItemData(CmbUnits.ListIndex)) & "," 'Unit ID
'vStrPara = vStrPara & Val(TxtWSPrice.Text) & "," 'WS Price
'vStrPara = vStrPara & IIf(TxtListPrice.Text = "", Null, Val(TxtListPrice.Text)) & "," 'List Price
'vStrPara = vStrPara & Val(TxtRetailPrice.Text) & "," 'Retail Price
'vStrPara = vStrPara & Val(TxtSaleDisc.Text) & "," 'Sale Disc
'vStrPara = vStrPara & Val(TxtSaleDiscPer.Text) & "," 'Sale Disc Per
'vStrPara = vStrPara & Val(TxtPurDisc.Text) & "," 'Pur Disc
'vStrPara = vStrPara & IIf(Val(TxtBottomPrice.Text) = 0, Null, TxtBottomPrice.Text) & "," 'Bottom Price
'vStrPara = vStrPara & IIf(Val(TxtMinStockLimit.Text) = 0, Null, TxtMinStockLimit.Text) & "," 'Min Stock Limit
'vStrPara = vStrPara & IIf(Val(TxtMaxStockLimit.Text) = 0, Null, TxtMaxStockLimit.Text) & "," 'Max Stock Limit
'vStrPara = vStrPara & IIf(Val(TxtServiceCharges.Text) = 0, Null, TxtServiceCharges.Text) & "," 'Service Charges
'vStrPara = vStrPara & IIf(Val(TxtEmpComm.Text) = 0, Null, TxtEmpComm.Text) & "," 'Emp Comm
'vStrPara = vStrPara & IIf(Val(TxtTokenVal.Text) = 0, Null, TxtTokenVal.Text) & "," 'Token Val
'vStrPara = vStrPara & IIf(Val(TxtSaleTaxPer.Text) = 0, Null, TxtSaleTaxPer.Text) & "," 'Sale Tax Per
'vStrPara = vStrPara & IIf(Val(TxtPCTCode.Text) = 0, Null, TxtPCTCode.Text) & "," 'PCTCode
'vStrPara = vStrPara & ChkLockProduct.Value & "," 'Lock Product
'vStrPara = vStrPara & ChkDeadProduct.Value & "," 'Dead Product
'vStrPara = vStrPara & ChkNoCostProduct.Value & "," 'No Cost Product
'vStrPara = vStrPara & ChkClosingProduct.Value & "," 'Closing Product
'vStrPara = vStrPara & Chk3rdScheduleItem.Value & "," '3rd Schedule Item
'vStrPara = vStrPara & ChkIsChangedPrice.Value & "," 'Is Changed Price
'vStrPara = vStrPara & ChkRawProduct.Value & "," 'Is Raw Product
'vStrPara = vStrPara & OptWSPSaleTax.Value & "," 'Is WS Sale Tax
'vStrPara = vStrPara & OptRPSaleTax.Value & "," 'R P SaleTax
'vStrPara = vStrPara & ChkWSDiscb4ST.Value & "," 'Is WS Disc b4 ST
'vStrPara = vStrPara & ChkDiscB4TradeOffer.Value & "," 'Is Disc B4 Trade Offer
'vStrPara = vStrPara & ChkDiscB4ExtraScheme.Value & "," 'Is Disc B4 Extra Scheme
'vStrPara = vStrPara & ChkDiscB4SaleTax.Value & "," 'is Disc B4 SaleTax
'vStrPara = vStrPara & ChkSerial.Value & "," 'Is Serial
'vStrPara = vStrPara & IIf(Val(TxtExtraSchemePer.Text) = 0, Null, TxtExtraSchemePer.Text) & "," 'Extra Scheme Per
'vStrPara = vStrPara & IIf(Val(TxtTradeOffer1.Text) = 0, Null, TxtTradeOffer1.Text) & "," 'Trade Offer 1
'vStrPara = vStrPara & IIf(Val(TxtTradeOffer2.Text) = 0, Null, TxtTradeOffer2.Text) & "," 'Trade Offer 2
'vStrPara = vStrPara & IIf(Trim(TxtDesc1.Text) = "", Null, Trim(TxtDesc1.Text)) 'Desc1
'vStrPara = Replace(vStrPara, "''", "Null")
   

'vStrPara = "DECLARE @returnvalue INT EXEC @returnvalue = saleheaderinsert " & vStrPara & " Select @returnvalue"
'   vMasterID = cn.Execute(vStrPara).Fields(0).Value
'   TxtSID.Text = vMasterID
'   MsgBox vMasterID
   
   If vIsNewRecord = True And CN.Execute("select count(*) from products where Productid = '" & TxtID.Text & "'").Fields(0) > 0 Then
'      MsgBox "This ID already exists. A new ID has been generated. Please save again", vbExclamation, "Alert"
      TxtID.Text = FunGetMaxID
'      TxtID.SetFocus
'      Exit Function
   End If

   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open "Select * FROM Products order by ProductName", CN, adOpenDynamic, adLockPessimistic
   Rs.Filter = "ProductID = '" & TxtID.Text & "'"

   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmProducts, eEdit, TxtID.Text, Date, "Effected Code-" & TxtID.Text & " Pur Price-" & Val(Rs!PurPrice) & " WSPrice-" & Val(Rs!WSPrice) & " Retail Price-" & Val(Rs!RetailPrice) & " Pur Disc-" & Val(IIf(IsNull(Rs!PurDiscPC), 0, Rs!PurDiscPC)) & " Sale Disc-" & Val(Rs!DiscPC) & " Emp Comm-" & Val(IIf(IsNull(Rs!EmpComm), 0, Rs!EmpComm)))

   If vIsNewRecord And Rs.RecordCount = 0 Then
      Rs.AddNew
      Rs!Productid = TxtID.Text
      Rs!PurPrice = Val(TxtPurPrice.Text)
      Rs!isChanged = 0
   Else
      Rs!isChanged = 1
      Rs!IsSync = 0
   End If
   If ObjUserSecurity.IsAdministrator = True Then Rs!PurPrice = Val(TxtPurPrice.Text)
   Rs!ProductName = TxtName.Text
   Rs!ProductName1 = IIf(Trim(TextBox1.Text) = "", Null, TextBox1.Text)
   Rs!companyid = IIf(Trim(TxtCompanyID.Text) = "", Null, TxtCompanyID.Text)
   Rs!OrganizationID = IIf(Trim(TxtOrganizationID.Text) = "", Null, TxtOrganizationID.Text)
   Rs!GroupID = TxtGroupID.Text
   Rs!SubGroupID = IIf(Trim(TxtSubGroupID.Text) = "", Null, TxtSubGroupID.Text)
   Rs!BrandID = IIf(Trim(TxtBrandID.Text) = "", Null, TxtBrandID.Text)
   Rs!SeasonID = IIf(Trim(TxtSeasonID.Text) = "", Null, TxtSeasonID.Text)
   Rs!DepartmentID = IIf(Trim(TxtDepartmentID.Text) = "", Null, TxtDepartmentID.Text)
   Rs!PubID = IIf(Trim(TxtPubID.Text) = "", Null, TxtPubID.Text)
   Rs!PurchasePackingID = IIf(CmbPurPacking.Text = "", Null, CmbPurPacking.ItemData(CmbPurPacking.ListIndex))
   Rs!SalePackingID = IIf(CmbSalePacking.Text = "", Null, CmbSalePacking.ItemData(CmbSalePacking.ListIndex))
   Rs!UnitID = IIf(CmbUnits.Text = "", Null, CmbUnits.ItemData(CmbUnits.ListIndex))
   Rs!WSPrice = Val(TxtWSPrice.Text)
   Rs!ListPrice = IIf(TxtListPrice.Text = "", Null, Val(TxtListPrice.Text))
   Rs!RetailPrice = Val(TxtRetailPrice.Text)
   Rs!DiscPC = Val(TxtSaleDisc.Text)
   Rs!DiscPer = Val(TxtSaleDiscPer.Text)
   Rs!PurDiscPC = Val(TxtPurDisc.Text)
   Rs!BottomPrice = IIf(Val(TxtBottomPrice.Text) = 0, Null, TxtBottomPrice.Text)
   Rs!MinStockLimit = IIf(Val(TxtMinStockLimit.Text) = 0, Null, TxtMinStockLimit.Text)
   Rs!MaxStockLimit = IIf(Val(TxtMaxStockLimit.Text) = 0, Null, TxtMaxStockLimit.Text)
   Rs!ServiceCharges = IIf(Val(TxtServiceCharges.Text) = 0, Null, TxtServiceCharges.Text)
   Rs!EmpComm = IIf(Val(TxtEmpComm.Text) = 0, Null, TxtEmpComm.Text)
   Rs!TokenVal = IIf(Val(TxtTokenVal.Text) = 0, Null, TxtTokenVal.Text)
   Rs!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, Null, TxtSaleTaxPer.Text)
   Rs!PCTCode = IIf(Val(TxtPCTCode.Text) = 0, Null, TxtPCTCode.Text)
   Rs!IsLocked = ChkLockProduct.Value
   Rs!isDeadProduct = ChkDeadProduct.Value
   Rs!IsNoCostProduct = ChkNoCostProduct.Value
   Rs!IsClosingProduct = ChkClosingProduct.Value
   Rs!is3rdScheduleItem = Chk3rdScheduleItem.Value
   Rs!isChangedPrice = ChkIsChangedPrice.Value
   Rs!IsRawProduct = ChkRawProduct.Value
   Rs!IsWSSaleTax = OptWSPSaleTax.Value
   Rs!IsRetailSaleTax = OptRPSaleTax.Value
   Rs!IsWSDiscb4ST = ChkWSDiscb4ST.Value
   Rs!IsDiscB4TradeOffer = ChkDiscB4TradeOffer.Value
   Rs!IsDiscB4ExtraScheme = ChkDiscB4ExtraScheme.Value
   Rs!isDiscB4SaleTax = ChkDiscB4SaleTax.Value
   Rs!IsSerial = ChkSerial.Value
   Rs!ExtraSchemePer = IIf(Val(TxtExtraSchemePer.Text) = 0, Null, TxtExtraSchemePer.Text)
   Rs!TradeOffer1 = IIf(Val(TxtTradeOffer1.Text) = 0, Null, TxtTradeOffer1.Text)
   Rs!TradeOffer2 = IIf(Val(TxtTradeOffer2.Text) = 0, Null, TxtTradeOffer2.Text)
   Rs!Desc1 = IIf(Trim(TxtDesc1.Text) = "", Null, Trim(TxtDesc1.Text))
   Rs!RackID = IIf(Trim(TxtRackID.Text) = "", Null, TxtRackID.Text)
   Rs!modified_on = Now
   Rs.Update
   
   With RsCode
      .Filter = 0
      If .RecordCount > 0 Then .MoveFirst
      For vCounter = 1 To .RecordCount
         !Productid = TxtID.Text
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   RsProductPacking.Filter = ""
   RsProductPacking.UpdateBatch
  
  If vIsNewRecord = False Then Call ActivityLogBin("", eFrmProducts, eEdit, TxtID.Text, Date, "Updated Code-" & TxtID.Text & " Pur Price-" & Val(TxtPurPrice.Text) & " WSPrice-" & Val(TxtWSPrice.Text) & " Retail Price-" & Val(TxtRetailPrice.Text) & " Pur Disc-" & Val(TxtPurDisc.Text) & " Sale Disc-" & Val(TxtSaleDisc.Text) & " Emp Comm-" & Val(TxtEmpComm.Text))
  If vIsNewRecord = True Then Call ActivityLogBin("", eFrmProducts, eAdd, TxtID.Text, Date, "Saved New Code -" & TxtID.Text & " Pur Price-" & Val(TxtPurPrice.Text) & " WSPrice-" & Val(TxtWSPrice.Text) & " Retail Price-" & Val(TxtRetailPrice.Text) & " Pur Disc-" & Val(TxtPurDisc.Text) & " Sale Disc-" & Val(TxtSaleDisc.Text) & " Emp Comm-" & Val(TxtEmpComm.Text))
'   If vIsNewRecord = True Then Call ActivityLog("Products", eAdd, , , TxtID.Text)
   CN.CommitTrans
   Rs.ReQuery
    Rs1.ReQuery
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   If TxtGroupID.Text = "" Then
      MsgBox "Please select a Group", vbExclamation, "Alert"
      If TxtGroupID.Enabled And TxtGroupID.Visible Then TxtGroupID.SetFocus
      Exit Function
   End If
   If vIsNewRecord Then
      If Trim(TxtID.Text) = "" Then
         MsgBox "Please specify a Product ID", vbExclamation, "Alert"
         If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
         Exit Function
      End If
      If Not IsNumeric(TxtID.Text) Then
         MsgBox "The Product ID must be numeric", vbExclamation, "Alert"
         If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
         Exit Function
      End If
   End If
   If Trim(TxtName.Text) = "" Then
      MsgBox "Please specify a Product Name", vbExclamation, "Alert"
      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
      Exit Function
   End If
   If Val(TxtRetailPrice.Text) <> 0 Then
      If Val(TxtSaleDiscPer.Text) <> Round((Val(TxtSaleDisc.Text) * 100) / Val(TxtRetailPrice.Text), 2) Then
         MsgBox "Please update the Discount for new Retail Price.", vbExclamation, "Alert"
         If TxtSaleDisc.Enabled And TxtSaleDisc.Visible Then TxtSaleDisc.SetFocus
         Exit Function
      End If
   End If
   If CmbSalePacking.Text <> "" Then
      RsProductPacking.Filter = "ProductID = '" & TxtID.Text & "' and PackingId = " & CmbSalePacking.ItemData(CmbSalePacking.ListIndex)
      If RsProductPacking.RecordCount = 0 Then
         MsgBox "Please select the multiplier of related packing.", vbExclamation, "Alert"
         If PGrid.Enabled And PGrid.Visible Then PGrid.SetFocus
         RsProductPacking.Filter = 0
         Exit Function
      End If
      RsProductPacking.Filter = 0
   End If
   If CmbPurPacking.Text <> "" Then
      RsProductPacking.Filter = "ProductID = '" & TxtID.Text & "' and PackingId = " & CmbPurPacking.ItemData(CmbPurPacking.ListIndex)
      If RsProductPacking.RecordCount = 0 Then
         MsgBox "Please select the multiplier of related packing.", vbExclamation, "Alert"
         If PGrid.Enabled And PGrid.Visible Then PGrid.SetFocus
         RsProductPacking.Filter = 0
         Exit Function
      End If
      RsProductPacking.Filter = 0
   End If

'    If Val(TxtPurchasePrice.Text) = 0 Then
'      MsgBox "The Purchase Price must be Greater than zero", vbExclamation, "Alert"
'      If TxtPurchasePrice.Enabled And TxtPurchasePrice.Visible Then TxtPurchasePrice.SetFocus
'      Exit Function
'    ElseIf Val(TxtSalePrice.Text) = 0 Then
'      MsgBox "The Sale price must be greater than zero", vbExclamation, "Alert"
'      If TxtSalePrice.Enabled And TxtSalePrice.Visible Then TxtSalePrice.SetFocus
'      Exit Function
'    ElseIf Val(TxtPurchaseDiscRatio.Text) > 99.99 Then
'      MsgBox "The Purchase Disc.(%) must be less than 99.99", vbExclamation, "Alert"
'      If TxtPurchaseDiscRatio.Enabled And TxtPurchaseDiscRatio.Visible Then TxtPurchaseDiscRatio.SetFocus
'      Exit Function
'    ElseIf Val(TxtPurchaseDiscRatio.Text) > 0 And Val(TxtPurchaseDiscVal.Text) > 0 Then
'      MsgBox "Only one of the Purchase Disc (%) or Purchase Disc. value must be provided.", vbExclamation, "Alert"
'      If TxtPurchaseDiscRatio.Enabled And TxtPurchaseDiscRatio.Visible Then TxtPurchaseDiscRatio.SetFocus
'      Exit Function
'    ElseIf Val(TxtSaleDiscRatio.Text) > 99.99 Then
'      MsgBox "The Sale Disc.(%) must be less than 99.99", vbExclamation, "Alert"
'      If TxtSaleDiscRatio.Enabled And TxtSaleDiscRatio.Visible Then TxtSaleDiscRatio.SetFocus
'      Exit Function
'    ElseIf Val(TxtSaleDiscRatio.Text) > 0 And Val(TxtSaleDiscVal.Text) > 0 Then
'      MsgBox "Only one of the Sale Disc (%) or Sale Disc. value must be provided.", vbExclamation, "Alert"
'      If TxtSaleDiscRatio.Enabled And TxtSaleDiscRatio.Visible Then TxtSaleDiscRatio.SetFocus
'      Exit Function
'    ElseIf Val(TxtPurchaseSTRatio.Text) > 99.99 Then
'      MsgBox "The Purchase S-Tax (%) must be less than 99.99", vbExclamation, "Alert"
'      If TxtPurchaseSTRatio.Enabled And TxtPurchaseSTRatio.Visible Then TxtPurchaseSTRatio.SetFocus
'      Exit Function
'    ElseIf Val(TxtSaleSTRatio.Text) > 99.99 Then
'      MsgBox "The Sale S-Tax (%) must be less than 99.99", vbExclamation, "Alert"
'      If TxtSaleSTRatio.Enabled And TxtSaleSTRatio.Visible Then TxtSaleSTRatio.SetFocus
'      Exit Function
'    End If
'
  'All Ok, now validation is success
  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Products"
   HelpLocation Me
   
   LblDepartment.Visible = ObjRegistry.isShowDepartment
   TxtDepartmentID.Visible = ObjRegistry.isShowDepartment
   TxtDepartmentName.Visible = ObjRegistry.isShowDepartment
   BtnDepartment.Visible = ObjRegistry.isShowDepartment
   BtnAddDepartment.Visible = ObjRegistry.isShowDepartment
   
   LblOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   BtnAddOrganization.Visible = ObjRegistry.OrganizationVisible
   
   LblTradeOffer.Visible = ObjRegistry.ShowTradeOffer
   TxtTradeOffer1.Visible = ObjRegistry.ShowTradeOffer
   TxtTradeOffer2.Visible = ObjRegistry.ShowTradeOffer
   LblPlusSign.Visible = ObjRegistry.ShowTradeOffer
   ChkDiscB4TradeOffer.Visible = ObjRegistry.ShowTradeOffer
   
   LblExtraSchemePer.Visible = ObjRegistry.ShowTradeOffer
   TxtExtraSchemePer.Visible = ObjRegistry.ShowTradeOffer
   ChkDiscB4ExtraScheme.Visible = ObjRegistry.ShowTradeOffer
   
   LblSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   TxtSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   LblPCTCode.Visible = ObjRegistry.ShowSaleTax
   TxtPCTCode.Visible = ObjRegistry.ShowSaleTax
   Chk3rdScheduleItem.Visible = ObjRegistry.ShowSaleTax
   ChkDiscB4SaleTax.Visible = ObjRegistry.ShowSaleTax

   
   If ObjRegistry.OrganizationVisible = True Then
      TxtOrganizationID.Text = ObjUserSecurity.OrganizationID
      FunSelectOrganization ssValidate, True
   End If
   
   
   CmbFilterGroup.AddItem ""
   If ObjRegistry.HeaderInfoNotClear = True Then
      TxtOrganizationID.Tag = "NC"
      TxtOrganizationName.Tag = "NC"
      TxtDepartmentID.Tag = "NC"
      TxtDepartmentName.Tag = "NC"
      TxtCompanyID.Tag = "NC"
      TxtCompanyName.Tag = "NC"
      TxtGroupID.Tag = "NC"
      TxtGroupName.Tag = "NC"
      TxtSubGroupID.Tag = "NC"
      TxtSubGroupName.Tag = "NC"
      TxtBrandID.Tag = "NC"
      TxtBrandName.Tag = "NC"
      TxtPubID.Tag = "NC"
      TxtPubName.Tag = "NC"
   End If
   With CN.Execute("Select * FROM Groups")
      Do Until .EOF
         CmbFilterGroup.AddItem !GroupName
         CmbFilterGroup.ItemData(CmbFilterGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
   End With
   With CN.Execute("Select * FROM Packings")
      CmbPurPacking.AddItem ""
      CmbSalePacking.AddItem ""
      Do Until .EOF
         CmbPurPacking.AddItem !PackingName
         CmbPurPacking.ItemData(CmbPurPacking.NewIndex) = !PackingID
         CmbSalePacking.AddItem !PackingName
         CmbSalePacking.ItemData(CmbSalePacking.NewIndex) = !PackingID
         .MoveNext
      Loop
   End With
   With CN.Execute("Select * FROM Units")
      CmbUnits.AddItem ""
      Do Until .EOF
         CmbUnits.AddItem !UnitName
         CmbUnits.ItemData(CmbUnits.NewIndex) = !UnitID
         .MoveNext
      Loop
   End With
   isFreeCode = False
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open "Select * FROM Products order by ProductName", CN, adOpenDynamic, adLockOptimistic
   If RsProductPacking.State = adStateOpen Then RsProductPacking.Close
   RsProductPacking.Open "Select * from ProductPacking", CN, adOpenDynamic, adLockBatchOptimistic
   If CmbFilterGroup.ListCount > 0 Then
      CmbFilterGroup.ListIndex = 0
   End If
   vPer = 0
   FormStatus = NewMode
   
   LblName2.Visible = ObjRegistry.AllowUrduProduct
   TextBox1.Visible = ObjRegistry.AllowUrduProduct
   LblGroupName1.Visible = ObjRegistry.AllowUrduProduct
   TxtGroupName1.Visible = ObjRegistry.AllowUrduProduct
'   LblDescription.Visible = Not ObjRegistry.AllowUrduProduct
'   TxtDesc1.Visible = Not ObjRegistry.AllowUrduProduct
   
   CGrid.Columns("Qty").Visible = ObjRegistry.QuantityinBarcodes
   If ObjRegistry.QuantityinBarcodes = True Then
      CGrid.Columns("Qty").Width = 35
      LblQty.Visible = True
      TxtQty.Visible = True
   End If
   
   Dim vWidth As Long, i As Integer
   vWidth = 0
   For i = 0 To CGrid.Cols - 1
      If CGrid.Columns(i).Visible = True Then
         vWidth = vWidth + CGrid.Columns(i).Width
      End If
   Next i
   CGrid.Width = vWidth + 18
   
   LblSeason.Visible = ObjRegistry.isShowSeason
   TxtSeasonID.Visible = ObjRegistry.isShowSeason
   TxtSeasonName.Visible = ObjRegistry.isShowSeason
   BtnSeason.Visible = ObjRegistry.isShowSeason
   BtnAddSeason.Visible = ObjRegistry.isShowSeason
   
   LblPublisher.Visible = ObjRegistry.isShowPublisher
   TxtPubID.Visible = ObjRegistry.isShowPublisher
   TxtPubName.Visible = ObjRegistry.isShowPublisher
   BtnPub.Visible = ObjRegistry.isShowPublisher
   BtnAddPub.Visible = ObjRegistry.isShowPublisher
   
   LblRack.Visible = ObjRegistry.isShowPublisher
   TxtRackID.Visible = ObjRegistry.isShowPublisher
   TxtRackName.Visible = ObjRegistry.isShowPublisher
   BtnRack.Visible = ObjRegistry.isShowPublisher
   BtnAddRack.Visible = ObjRegistry.isShowPublisher
   
   LblListPrice.Visible = ObjRegistry.isShowListPrice
   TxtListPrice.Visible = ObjRegistry.isShowListPrice
   
   ModeValue = False
   
   
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   
   
   Exit Sub
ErrorHandler:
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
      'If Val(TxtID.Text) <> 0 Then
      '   TxtID.Text = Right("00000" + CStr(Val(TxtID.Text) + 1), 5)
      'Else
         TxtID.Text = FunGetMaxID
      'End If
      If ChkDataNotClear.Value = 0 Then Call SubClearFields(True)
      PGrid.Enabled = True
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      If ChkSearchNotClear.Value = 0 Then
        If TxtGroupName.Text <> "" Then CmbFilterGroup.Text = TxtGroupName.Text
      End If
      If Rs.State = adStateOpen Then Rs.Close
      Rs.Open "Select * FROM Products order by ProductName", CN, adOpenDynamic, adLockOptimistic
      
      If ChkDataNotClear.Value = 0 Then
        If RsProductPacking.State = adStateOpen Then RsProductPacking.Close
        RsProductPacking.Open "Select * from ProductPacking", CN, adOpenDynamic, adLockBatchOptimistic
        
        CmbPurPacking.ListIndex = 0
        CmbSalePacking.ListIndex = 0
        CmbUnits.ListIndex = 0
      End If
      CmbFilterGroup.Enabled = False
      TxtFilterProductName.Enabled = False
      TxtFilterID.Enabled = False
      Grid.Enabled = False
      PGrid.Enabled = True
      CGrid.Enabled = True
      If ChkSearchNotClear.Value = 0 Then
        If CmbFilterGroup.ListCount > 0 Then
           CmbFilterGroup.ListIndex = 0
        End If
      End If
      'If CmbFilterGroup.ListCount > 0 Then
      '   TxtGroupID.Text = GetGroupID(CmbFilterGroup)
      '   TxtGroupName.Text = CN.Execute("Select Groupname from Groups where Groupid='" & TxtGroupID.Text & "'").Fields(0).Value
      '   'TxtPrefix.Text = TxtGroupID.Text
      'End If
      PopulatePackGrid
      PopulateCodeGrid
      If TxtCompanyID.Visible And TxtCompanyID.Enabled Then TxtCompanyID.SetFocus
      If ObjRegistry.HeaderInfoNotClear = True And TxtGroupID.Text <> "" Then
         If TxtName.Visible And TxtName.Enabled Then TxtName.SetFocus
      End If
      vIsNewRecord = True
   Case Is = OpenMode
      'If ChkSearchNotClear.Value = 0 Then
        Call SubClearFields(True)
      'Else
      '  Call SubClearFields(False)
      'End If
      Call Grid_RowColChange(0, 0)
      'TxtGroupName.Enabled = False
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnClear.Enabled = True
      Grid.Enabled = False
      PGrid.Enabled = True
      CGrid.Enabled = True
      TxtID.Enabled = False
      CmbFilterGroup.Enabled = True
      TxtCompanyID.SetFocus
      PopulatePackGrid
      PopulateCodeGrid
      vIsNewRecord = False
   Case Is = ChangeMode
      BtnSave.Enabled = True
   Case Is = SelectionMode
      'CmbFilterGroup.ListIndex = 0
      Call SubClearFields(False)
      Grid.Enabled = True
      If ChkSearchNotClear.Value = 0 Then
        Call Grid_RowColChange(0, 0)
      End If
      PGrid.Enabled = False
      CGrid.Enabled = False
      CmbFilterGroup.Enabled = True
      TxtFilterProductName.Enabled = True
      TxtFilterID.Enabled = True
      Grid.SetFocus
      
      PopulatePackGrid
      PopulateCodeGrid
      BtnNew.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnSave.Enabled = False
      BtnClear.Enabled = False
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub PopulatePackGrid()
On Error GoTo ErrorHandler
   Dim vPackSql As String
   Dim vrowcount As Integer
   vPackSql = "Select Packings.PackingID,Packings.PackingName,isnull(ProductPacking.Multiplier,0) as Multiplier from Packings" _
            + " Left outer join (select * from  ProductPacking where Productid = '" & TxtID.Text & "')ProductPacking on Packings.PackingID=ProductPacking.PackingID"
   If RsPacking.State = adStateOpen Then RsPacking.Close
   RsPacking.Open vPackSql, CN, adOpenDynamic, adLockOptimistic
   If RsPacking.RecordCount > 0 Then RsPacking.MoveFirst
   PGrid.RemoveAll
   While Not RsPacking.EOF
      PGrid.AddNew
      PGrid.Columns("Packing").Text = RsPacking!PackingName
      PGrid.Columns("Multiplier").Value = RsPacking!Multiplier
      PGrid.Columns("PackingID").Value = RsPacking!PackingID
      PGrid.Update
      RsPacking.MoveNext
   Wend
   If PGrid.Rows > 0 Then PGrid.FirstRow = 0
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
 End Sub

Private Sub PopulateCodeGrid()
On Error GoTo ErrorHandler
   Dim vSQL As String
   CGrid.Redraw = False
   CGrid.MoveFirst
   CGrid.RemoveAll
   vSQL = "Select * from ProductBarcodes where ProductID='" & TxtID.Text & "'"
   If RsCode.State = adStateOpen Then RsCode.Close
   RsCode.Open vSQL, CN, adOpenDynamic, adLockBatchOptimistic
   CGrid.RemoveAll
   RsCode.Filter = 0
   While Not RsCode.EOF
      CGrid.AddNew
      CGrid.Columns("Code").Text = RsCode!code
      CGrid.Columns("Qty").Value = IIf(IsNull(RsCode!Qty), "", RsCode!Qty)
      CGrid.Update
      RsCode.MoveNext
   Wend
   If CGrid.Rows > 0 Then CGrid.FirstRow = 0
   CGrid.Redraw = False
   With CGrid
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
   End With
   CGrid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
 End Sub
 
Private Sub Grid_Click()
On Error GoTo ErrorHandler
    If Grid.Rows > 0 Then Call Grid_RowColChange(0, 0)
    Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
   Select Case KeyAscii
   Case vbKeyA To vbKeyZ, Asc("a") To Asc("z")
      TxtFilterProductName.Text = TxtFilterProductName.Text & Chr(KeyAscii): TxtFilterProductName.SelStart = Len(TxtFilterProductName.Text): TxtFilterProductName.SetFocus
   Case vbKey0 To vbKey9
      TxtFilterID.Text = TxtFilterID.Text & Chr(KeyAscii): TxtFilterID.SelStart = Len(TxtFilterID.Text): TxtFilterID.SetFocus
   End Select
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ID").Text) <> "" Then
      Rs1.Filter = "ProductID = " & Val(Grid.Columns("ID").Text)
      If Rs1.RecordCount > 0 And Grid.Enabled Then
         TxtOrganizationID.Text = IIf(IsNull(Rs1!OrganizationID), "", Rs1!OrganizationID)
         If Trim(TxtOrganizationID.Text) <> "" Then
            TxtOrganizationName.Text = CN.Execute("Select OrganizationName from Organizations Where Organizationid='" & TxtOrganizationID.Text & "'").Fields(0)
         Else
            TxtOrganizationName.Text = ""
         End If
         TxtCompanyID.Text = IIf(IsNull(Rs1!companyid), "", Rs1!companyid)
         If Trim(TxtCompanyID.Text) <> "" Then
            TxtCompanyName.Text = CN.Execute("Select CompanyName from Companies Where Companyid='" & TxtCompanyID.Text & "'").Fields(0)
         Else
            TxtCompanyName.Text = ""
         End If
         TxtGroupID.Text = Rs1!GroupID
         TxtGroupName.Text = CN.Execute("Select GroupName from Groups Where Groupid='" & TxtGroupID.Text & "'").Fields(0)
         TxtSubGroupID.Text = IIf(IsNull(Rs1!SubGroupID), "", Rs1!SubGroupID)
         If Trim(TxtSubGroupID.Text) <> "" Then
            TxtSubGroupName.Text = CN.Execute("Select SubGroupName from SubGroups Where SubGroupid='" & TxtSubGroupID.Text & "'").Fields(0)
         Else
            TxtSubGroupName.Text = ""
         End If
         
         TxtDepartmentID.Text = IIf(IsNull(Rs1!DepartmentID), "", Rs1!DepartmentID)
         If Trim(TxtDepartmentID.Text) <> "" Then
            TxtDepartmentName.Text = CN.Execute("Select Department from Departments Where Departmentid='" & TxtDepartmentID.Text & "'").Fields(0)
         Else
            TxtDepartmentName.Text = ""
         End If
         
         TxtPubID.Text = IIf(IsNull(Rs1!PubID), "", Rs1!PubID)
         If Trim(TxtPubID.Text) <> "" Then
            TxtPubName.Text = CN.Execute("Select PubName from Publishers Where Pubid='" & TxtPubID.Text & "'").Fields(0)
         Else
            TxtPubName.Text = ""
         End If
         TxtBrandID.Text = IIf(IsNull(Rs1!BrandID), "", Rs1!BrandID)
         If Trim(TxtBrandID.Text) <> "" Then
            TxtBrandName.Text = CN.Execute("Select BrandName from Brands Where BrandID = '" & TxtBrandID.Text & "'").Fields(0)
         Else
            TxtBrandName.Text = ""
         End If
         TxtSeasonID.Text = IIf(IsNull(Rs1!SeasonID), "", Rs1!SeasonID)
         If Trim(TxtSeasonID.Text) <> "" Then
            TxtSeasonName.Text = CN.Execute("Select SeasonName from Seasons Where SeasonID = '" & TxtSeasonID.Text & "'").Fields(0)
         Else
            TxtSeasonName.Text = ""
         End If
         
         TxtRackID.Text = IIf(IsNull(Rs1!RackID), "", Rs1!RackID)
         If Trim(TxtRackID.Text) <> "" Then
            TxtRackName.Text = CN.Execute("Select RackName from Racks Where Rackid='" & TxtRackID.Text & "'").Fields(0)
         Else
            TxtRackName.Text = ""
         End If
         
         TxtID.Text = Grid.Columns("ID").Text
         TxtName.Text = Grid.Columns("Name").Text
         '''''''''''''''''''''''''''''''''''''''''''''
         '===================================================
         With CN.Execute("Select PackingName from Packings Where PackingID = " & IIf(IsNull(Rs1!PurchasePackingID), "0", Rs1!PurchasePackingID))
            If .RecordCount = 0 Then
               CmbPurPacking.ListIndex = 0
            Else
               CmbPurPacking.Text = !PackingName
            End If
         End With
         With CN.Execute("Select PackingName from Packings Where PackingID = " & IIf(IsNull(Rs1!SalePackingID), "0", Rs1!SalePackingID))
            If .RecordCount = 0 Then
               CmbSalePacking.ListIndex = 0
            Else
               CmbSalePacking.Text = !PackingName
            End If
         End With
         '===================================================
         With CN.Execute("Select UnitName from Units Where UnitID = " & IIf(IsNull(Rs1!UnitID), "0", Rs1!UnitID))
            If .RecordCount = 0 Then
               CmbUnits.ListIndex = 0
            Else
               CmbUnits.Text = !UnitName
            End If
         End With
         '====================================================
         vStrSQL = "Select isnull(GroupName1,'') from groups where GroupID = '" & TxtGroupID.Text & "'"
         TxtGroupName1.Text = CN.Execute(vStrSQL).Fields(0).Value
         TextBox1.Text = IIf(IsNull(Rs1!ProductName1), "", Rs1!ProductName1)
         TxtPurPrice.Text = IIf(ObjUserSecurity.IsAdministrator = True, Rs1!PurPrice, 0)
         TxtListPrice.Text = IIf(IsNull(Rs1!ListPrice), "", Rs1!ListPrice)
         TxtRetailPrice.Text = Rs1!RetailPrice
         TxtWSPrice.Text = IIf(IsNull(Rs1!WSPrice), "", Rs1!WSPrice)
         TxtPurDisc.Text = IIf(IsNull(Rs1!PurDiscPC), 0, Rs1!PurDiscPC)
         TxtSaleDisc.Text = IIf(IsNull(Rs1!DiscPC), 0, Rs1!DiscPC)
         TxtSaleDiscPer.Text = IIf(IsNull(Rs1!DiscPer), 0, Rs1!DiscPer)
         TxtMinStockLimit.Text = IIf(IsNull(Rs1!MinStockLimit), 0, Rs1!MinStockLimit)
         TxtMaxStockLimit.Text = IIf(IsNull(Rs1!MaxStockLimit), 0, Rs1!MaxStockLimit)
         TxtTokenVal.Text = IIf(IsNull(Rs1!TokenVal), 0, Rs1!TokenVal)
         TxtSaleTaxPer.Text = IIf(IsNull(Rs1!SaleTaxPer), "", Rs1!SaleTaxPer)
         TxtPCTCode.Text = IIf(IsNull(Rs1!PCTCode), "", Rs1!PCTCode)
         TxtServiceCharges.Text = IIf(IsNull(Rs1!ServiceCharges), 0, Rs1!ServiceCharges)
         TxtEmpComm.Text = IIf(IsNull(Rs1!EmpComm), 0, Rs1!EmpComm)
         TxtDesc1.Text = IIf(IsNull(Rs1!Desc1), "", Rs1!Desc1)
         TxtBottomPrice.Text = IIf(IsNull(Rs1!BottomPrice), 0, Rs1!BottomPrice)
         ChkLockProduct.Value = Abs(Rs1!IsLocked)
         ChkDeadProduct.Value = Abs(IIf(IsNull(Rs1!isDeadProduct), 0, Rs1!isDeadProduct))
         ChkClosingProduct.Value = Abs(Rs1!IsClosingProduct)
         Chk3rdScheduleItem.Value = Abs(IIf(IsNull(Rs1!is3rdScheduleItem), 0, Rs1!is3rdScheduleItem))
         ChkNoCostProduct.Value = Abs(Rs1!IsNoCostProduct)
         ChkIsChangedPrice.Value = Abs(Rs1!isChangedPrice)
         ChkRawProduct.Value = Abs(Rs1!IsRawProduct)
         OptWSPSaleTax.Value = Rs1!IsWSSaleTax
         OptRPSaleTax.Value = Rs1!IsRetailSaleTax
         ChkWSDiscb4ST.Value = Abs(Rs1!IsWSDiscb4ST)
         ChkDiscB4TradeOffer.Value = Abs(Rs1!IsDiscB4TradeOffer)
         ChkDiscB4ExtraScheme.Value = Abs(Rs1!IsDiscB4ExtraScheme)
         ChkDiscB4SaleTax.Value = Abs(Rs1!isDiscB4SaleTax)
         ChkSerial.Value = Abs(Rs1!IsSerial)
         TxtTradeOffer1.Text = IIf(IsNull(Rs1!TradeOffer1), 0, Rs1!TradeOffer1)
         TxtTradeOffer2.Text = IIf(IsNull(Rs1!TradeOffer2), 0, Rs1!TradeOffer2)
         TxtExtraSchemePer.Text = IIf(IsNull(Rs1!ExtraSchemePer), 0, Rs1!ExtraSchemePer)
         PopulatePackGrid
         PopulateCodeGrid
      End If
   End If
   Rs1.Filter = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataBackFromCGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtCode.Text = .Columns("code").Text
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToCGrid()
   If IsNumeric(TxtCode.Text) = True Then
      If (TxtCode.Text) = "" Or Len(TxtCode.Text) < 6 Then
         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   If CGrid.Columns("Code").Text = "" And CN.Execute("Select count(*) from ProductBarcodes where Code = '" & (TxtCode.Text) & "'").Fields(0).Value = 1 Then
      MsgBox "Code Already Exists", vbExclamation, "Alert"
      TxtCode.SetFocus
      Exit Sub
   End If
   If (TxtCode.Text) <> CGrid.Columns("Code").Text Then
      If CGrid.Columns("Code").Text <> "" And CN.Execute("Select count(*) from ProductBarcodes where Code = '" & TxtCode.Text & "'").Fields(0).Value = 1 Then
         MsgBox "Code Already Exists", vbExclamation, "Alert"
         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   If ObjRegistry.IsSingleBarcode = True And CGrid.Row >= 1 Then
    MsgBox "Multiple Barcode can not be eneterd ", vbExclamation, "Alert"
    Exit Sub
   End If
On Error GoTo ErrorHandler
   RsCode.Filter = "Code = '" & IIf(Trim(CGrid.Columns("Code").Text) <> "", CGrid.Columns("Code").Text, (TxtCode.Text)) & "'"
   If RsCode.RecordCount = 0 Then RsCode.AddNew
   CGrid.Columns("Code").Text = (TxtCode.Text)
   CGrid.Columns("Qty").Value = IIf(Val(TxtQty.Text) = 0, "", Val(TxtQty.Text))
   If vIsNewRecord = False Then CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated BarCode-" & CGrid.Columns("Code").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   RsCode!Productid = Val(TxtID.Text)
   RsCode!code = (TxtCode.Text)
   RsCode!Qty = IIf(Val(TxtQty.Text) = 0, Null, Val(TxtQty.Text))
   RsCode.Update
   CGrid.Update
   CGrid.Redraw = False
   CGrid.MoveLast
   With CGrid
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   CGrid.Redraw = True
   TxtCode.Text = ""
   TxtQty.Text = ""
   CGrid.MoveLast
   CGrid.MoveNext
   TxtCode.SetFocus
   Exit Sub
ErrorHandler:
   CGrid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SubClearFields(Enable As Boolean)
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then ctl.Text = ""
         If ctl.Tag = "" Then ctl.Enabled = Enable
      ElseIf TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is ComboBox Then
         If ctl.Tag = "" Then ctl.Enabled = Enable
      ElseIf TypeOf ctl Is JeweledButton Then
         If ctl.Tag <> "" Then ctl.Enabled = Enable
      ElseIf TypeOf ctl Is CheckBox Then
         If ctl.Tag = "" Then ctl.Value = 0
         ctl.Enabled = Enable
      ElseIf TypeOf ctl Is OptionButton Then
         ctl.Enabled = Enable
      End If
      OptWSPSaleTax.Value = True
   Next
   TextBox1.Text = ""
   TxtGroupName1.Text = ""
   TxtGroupName.Enabled = False
   TxtOrganizationName.Enabled = False
   TxtCompanyName.Enabled = False
   TxtSubGroupName.Enabled = False
   TxtBrandName.Enabled = False
   CGrid.CancelUpdate
   CGrid.RemoveAll
   CGrid.AddNew
   CGrid.Columns("Code").Text = " "
   CGrid.Update
   Rs.Filter = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
   On Error GoTo ErrorHandler
   If isFreeCode = True Then
      Dim vSQL As String
      'vSQL = "SELECT min(ProductID)" & vbCrLf _
         + " FROM Products WHERE" & vbCrLf _
         + "     --The next # is missing" & vbCrLf _
         + " ProductID+1 NOT IN (SELECT ProductID FROM Products)" & vbCrLf _
         + " AND" & vbCrLf _
         + "   --We haven't reached the max of values" & vbCrLf _
         + " ProductID+1 < (SELECT Max(ProductID) FROM Products)"
         
       vSQL = "SELECT CASE WHEN MAX(productid) = COUNT(*)" & vbCrLf _
            + " THEN CAST(NULL AS INTEGER)" & vbCrLf _
            + "  -- THEN MAX(column_name) + 1 as other option" & vbCrLf _
            + " WHEN Min(productid) > 1 THEN 1" & vbCrLf _
            + " WHEN MAX(productid) <> COUNT(*) THEN (SELECT MIN(productid)+1" & vbCrLf _
            + " From Products Where (productid + 1)" & vbCrLf _
            + " NOT IN (SELECT productid FROM Products)) ELSE NULL END" & vbCrLf _
            + " FROM Products;"
      FunGetMaxID = CN.Execute(vSQL).Fields(0)
   Else
      vSQL = "Select max(ProductId)+ 1 from Products --Where ProductId like '" & TxtGroupID.Text & "%'"""
      FunGetMaxID = CN.Execute("Select max(ProductId)+1 from Products --Where ProductId like '" & TxtGroupID.Text & "%'").Fields(0)
   End If
   'FunGetMaxID = CN.Execute("Select right('0000' + cast(isnull(max(cast(substring(ProductId,3,10) as smallint)),0) + 1 as varchar),4) from Products").Fields(0) ' Where ProductId like '" & GetGroupID(CmbCompany) & "%'").Fields(0)
   If ObjRegistry.DuplicateCode = True Then
      TxtCode.Text = TxtGroupID.Text & CN.Execute("Select right('0000' + cast(isnull(max(cast(substring(Code,4,10) as int)),0) + 1 as varchar),4) from ProductBarcodes Where len(code)=7 and Code like '" & TxtGroupID.Text & "%'").Fields(0)
   End If
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub ImgExit_Click()
On Error GoTo ErrorHandler
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub LblClose_Click()
On Error GoTo ErrorHandler
   FraHelp.Visible = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo ErrorHandler
   LblHelp.ForeColor = &H800000
   FraHelp.ZOrder 0
   FraHelp.Visible = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo ErrorHandler
   If LblHelp.FontUnderline = True Then Exit Sub
   LblHelp.FontUnderline = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LblHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo ErrorHandler
   LblHelp.ForeColor = vbWhite
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub UpdatePacking()
On Error GoTo ErrorHandler
   If Val(PGrid.Columns("Multiplier").Value) = 0 Then PGrid.Columns("Multiplier").Value = 0
   RsProductPacking.Filter = "ProductID='" & TxtID.Text & "' and PackingID=" & Val(PGrid.Columns("PackingID").Value)
   If RsProductPacking.RecordCount = 0 And Val(PGrid.Columns("Multiplier").Value) > 0 Then
      RsProductPacking.AddNew
      RsProductPacking!PackingID = PGrid.Columns("PackingID").Value
      RsProductPacking!Productid = TxtID.Text
      RsProductPacking!Multiplier = Val(PGrid.Columns("Multiplier").Value)
'      CmbPurPacking.Text = PGrid.Columns("Packing").Text
'      CmbSalePacking.Text = PGrid.Columns("Packing").Text
      If vIsNewRecord = False Then CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Inserted New PackingID-v" & RsProductPacking!PackingID & " Multiplier- " & RsProductPacking!Multiplier & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   ElseIf RsProductPacking.RecordCount = 1 And Val(PGrid.Columns("Multiplier").Value) = 0 Then
      If vIsNewRecord = False Then CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Deleted PackingID-v" & RsProductPacking!PackingID & " Multiplier- " & RsProductPacking!Multiplier & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
      RsProductPacking.Delete
   ElseIf RsProductPacking.RecordCount = 1 Then
      RsProductPacking!Multiplier = Val(PGrid.Columns("Multiplier").Value)
      If vIsNewRecord = False Then CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated PackingID-v" & RsProductPacking!PackingID & " Multiplier- " & RsProductPacking!Multiplier & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
      RsProductPacking.Update
'      CmbPurPacking.Text = PGrid.Columns("Packing").Text
'      CmbSalePacking.Text = PGrid.Columns("Packing").Text
  End If
  Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PGrid_BeforeUpdate(Cancel As Integer)
On Error GoTo ErrorHandler
   If PGrid.Visible = False Then Exit Sub
   If ActiveControl.Name <> PGrid.Name Then Exit Sub
   UpdatePacking
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PGrid_Change()
On Error GoTo ErrorHandler
   If BtnSave.Enabled = False Then BtnSave.Enabled = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PGrid_GotFocus()
On Error GoTo ErrorHandler
   PGrid.Row = 0
   PGrid.Col = 0
'   SendKeys "{Right}"
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PGrid_LostFocus()
On Error GoTo ErrorHandler
   UpdatePacking
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCode_LostFocus()
On Error GoTo ErrorHandler
   If Trim(TxtCode.Text) = "" Then Exit Sub
   GetDataFromTexBoxesToCGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtFilterID_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtFilterID.Name Then Exit Sub
   If Trim(TxtFilterID.Text) = "" Then Grid.MoveFirst: Exit Sub
   If Len(TxtFilterID.Text) > 5 Or Not IsNumeric(TxtFilterID.Text) Then
      With CN.Execute("select * from Productbarcodes where Code = '" & TxtFilterID.Text & "'")
         If .RecordCount > 0 Then
            Rs1.Find "ProductID = " & Val(!Productid), , adSearchForward, 1
         End If
         .Close
      End With
   Else
      Rs1.Find "ProductID = " & Val(TxtFilterID.Text), , adSearchForward, 1
   End If
   If Rs1.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtBrandID_Change()
On Error GoTo ErrorHandler
   If TxtBrandID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtBrandID.Name Then Exit Sub
   If TxtBrandName.Text <> "" Then TxtBrandName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtBrandID_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
If Me.ActiveControl.Name <> TxtBrandID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtBrandName.Text <> "" Then Exit Sub
   If TxtBrandID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBrand(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBrand(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCompanyID_Change()
On Error GoTo ErrorHandler
   If TxtCompanyID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCompanyID.Name Then Exit Sub
   If TxtCompanyName.Text <> "" Then TxtCompanyName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCompanyID_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
If Me.ActiveControl.Name <> TxtCompanyID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCompanyID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCompany(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCompany(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtGroupID_Change()
On Error GoTo ErrorHandler
   If TxtGroupID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   If TxtGroupName.Text <> "" Then TxtGroupName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtGroupID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtGroupName.Text <> "" Then Exit Sub
   If TxtGroupID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectGroup(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectGroup(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub TxtID_LostFocus()
'   On Error GoTo ErrorHandler
'   'If Len(TxtID.Text) = 5 Then Exit Sub
'   'TxtID.Text = Right("00000" + CStr(Val(TxtID.Text)), 5)
'   'Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub

Private Sub TxtName_LostFocus()
   On Error GoTo ErrorHandler
   If ObjRegistry.ProperCase = False Then Exit Sub
   TxtName.Text = StrConv(TxtName.Text, vbProperCase)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtOrganizationID.Text = "" Then Exit Sub
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

Private Sub TxtPubID_Change()
On Error GoTo ErrorHandler
   If TxtPubID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtPubID.Name Then Exit Sub
   If TxtPubName.Text <> "" Then TxtPubName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtPubID_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
If Me.ActiveControl.Name <> TxtPubID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtPubName.Text <> "" Then Exit Sub
   If TxtPubID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectPublisher(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectPublisher(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub TxtPurPrice_Change()
On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtPurPrice.Name Then Exit Sub
   If vPer = 0 Then Exit Sub
   TxtRetailPrice.Text = SelfRound(Val(TxtPurPrice.Text) + (Val(TxtPurPrice.Text) * vPer / 100))
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtQty_LostFocus()
On Error GoTo ErrorHandler
   If Trim(TxtQty.Text) = "" Then Exit Sub
   GetDataFromTexBoxesToCGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSaleDisc_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSaleDisc.Name Then Exit Sub
   If Val(TxtRetailPrice.Text) = 0 Then Exit Sub
   TxtSaleDiscPer.Text = Round((Val(TxtSaleDisc.Text) * 100) / Val(TxtRetailPrice.Text), 2)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSaleDiscPer_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSaleDiscPer.Name Then Exit Sub
   TxtSaleDisc.Text = Round((Val(TxtRetailPrice.Text) * Val(TxtSaleDiscPer.Text) / 100), 2)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSubGroupID_Change()
On Error GoTo ErrorHandler
   If TxtSubGroupID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSubGroupID.Name Then Exit Sub
   If TxtSubGroupName.Text <> "" Then TxtSubGroupName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSubGroupID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtSubGroupID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSubGroupID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSubGroup(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSubGroup(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtFilterProductName_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtFilterProductName.Name Then Exit Sub
   If Trim(TxtFilterProductName.Text) = "" Then Exit Sub
   Set Rs1 = New ADODB.Recordset
   Dim vWords() As String
   Dim vProductName As String
   
   vWords = Split(TxtFilterProductName.Text, " ")
   vProductName = ""
   For i = 0 To UBound(vWords)
       vProductName = vProductName & " and Productname like '%" & Replace(vWords(i), "'", "''") & "%'"
   Next
   Rs1.Open "Select * FROM Products where 1=1 " & vProductName & " Order By ProductName", CN, adOpenDynamic, adLockOptimistic
   Set Grid.DataSource = Rs1
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtName_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtName.Name Then Exit Sub
   If Trim(TxtName.Text) = "" Then Exit Sub
   Set Rs1 = New ADODB.Recordset
   Dim vWords() As String
   Dim vProductName As String
   
   vWords = Split(TxtName.Text, " ")
   vProductName = ""
   For i = 0 To UBound(vWords)
       vProductName = vProductName & " and Productname like '%" & Replace(vWords(i), "'", "''") & "%'"
   Next
   Rs1.Open "Select * FROM Products where 1=1 " & vProductName & " Order By ProductName", CN, adOpenDynamic, adLockOptimistic
   If ChkSearchNotClear.Value = 0 Then
    Set Grid.DataSource = Rs1
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxBinID = CN.Execute("Select isnull(max(BinID),0)+1 from Bin_Products ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
On Error GoTo ErrorHandler
    If vIsNewRecord = False Then
         If TxtOrganizationID.Text <> IIf(IsNull(Rs!OrganizationID), "", Rs!OrganizationID) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated OrganizationID-" & Rs!OrganizationID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtCompanyID.Text <> IIf(IsNull(Rs!companyid), "", Rs!companyid) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated CompanyID-" & Rs!companyid & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtGroupID.Text <> IIf(IsNull(Rs!GroupID), "", Rs!GroupID) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated GroupID-" & Rs!GroupID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtSubGroupID.Text <> IIf(IsNull(Rs!SubGroupID), "", Rs!SubGroupID) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated SubGroupID NO -" & Rs!SubGroupID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtBrandID.Text <> IIf(IsNull(Rs!BrandID), "", Rs!BrandID) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated BrandID NO -" & Rs!BrandID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtName.Text <> Rs!ProductName Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated Product Name-" & Rs!ProductName & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If CmbUnits.ItemData(CmbUnits.ListIndex) <> IIf(IsNull(Rs!UnitID), "", Rs!UnitID) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated UnitID-" & Rs!UnitID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If CmbPurPacking.ItemData(CmbPurPacking.ListIndex) <> IIf(IsNull(Rs!PurchasePackingID), "", Rs!PurchasePackingID) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated PurchasePackingID-" & Rs!PurchasePackingID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If CmbSalePacking.ItemData(CmbSalePacking.ListIndex) <> IIf(IsNull(Rs!SalePackingID), "", Rs!SalePackingID) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated SalePackingID-" & Rs!SalePackingID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPurPrice.Text <> IIf(IsNull(Rs!PurPrice), "", Rs!PurPrice) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated PurPrice NO-" & Rs!PurPrice & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtRetailPrice.Text <> IIf(IsNull(Rs!RetailPrice), "", Rs!RetailPrice) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated RetailPrice-" & Rs!RetailPrice & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtWSPrice.Text <> IIf(IsNull(Rs!WSPrice), "", Rs!WSPrice) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated WSPrice-" & Rs!WSPrice & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPurDisc.Text <> IIf(IsNull(Rs!PurDiscPC), "", Rs!PurDiscPC) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated PurDiscPC-" & Rs!PurDiscPC & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtSaleDisc.Text <> IIf(IsNull(Rs!DiscPC), "", Rs!DiscPC) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated DiscPC-" & Rs!DiscPC & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtMinStockLimit.Text <> IIf(IsNull(Rs!MinStockLimit), "", Rs!MinStockLimit) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated MinStockLimit-" & Rs!MinStockLimit & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtMaxStockLimit.Text <> IIf(IsNull(Rs!MaxStockLimit), "", Rs!MaxStockLimit) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated MaxStockLimit-" & Rs!MaxStockLimit & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If ChkLockProduct.Value <> Abs(Rs!IsLocked) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated IsLocked-" & Rs!IsLocked & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If ChkDeadProduct.Value <> Abs(Rs!isDeadProduct) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated IsDead-" & Rs!isDeadProduct & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If ChkClosingProduct.Value <> Abs(Rs!IsClosingProduct) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated IsClosingProduct-" & Rs!IsClosingProduct & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If Chk3rdScheduleItem.Value <> Abs(Rs!is3rdScheduleItem) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated Is3rdScheduleItem-" & Rs!is3rdScheduleItem & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If ChkNoCostProduct.Value <> Abs(Rs!IsNoCostProduct) Then
            CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated IsNoCostProduct-" & Rs!IsNoCostProduct & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
   Else
        CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TextBox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
On Error GoTo ErrorHandler
        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyDown-Event behaviours for  ''
        '' Enter, Space, Delete & Tab keys to set      ''
        '' Behavior in Textbox1.Text, keys will behave ''
        '' as Normal Text writing behavior.            ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''

   If ModeValue = False Then
      'Space Key Behavior
         If KeyCode = 32 Then
         UniCode = &H20
         TextBox1.Text = TextBox1.Text + ChrW(UniCode)
         KeyCode = 0

        'Enter Key Behavior
'        ElseIf KeyCode = 13 Then
'        UniCode = &HA
'        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
'        KeyCode = 0

        'Horizontal Tab Behavior
'        ElseIf KeyCode = 9 Then
'        UniCode = &H9
'        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
'        KeyCode = 0

         'Delete Key Behavior
         ElseIf KeyCode = 127 Then
         UniCode = &H7F
         TextBox1.Text = TextBox1.Text + ChrW(UniCode)
         KeyCode = 0
         End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
        
        'This Function Got End There
End Sub

Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
On Error GoTo ErrorHandler
        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyPress-Event behaviours for ''
        '' Alfabatic, Numaric & Symbolic keys to write ''
        '' Urdu. I've tried to make it near with Urdu  ''
        '' Phonetic Keyboard Layout.                   ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''

If ModeValue = False Then

        'For Small Letter's Behaviors

        'a Key Behavior
        If KeyAscii = 97 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H627
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'b Key Behavior
        ElseIf KeyAscii = 98 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H628
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'c Key Behavior
        ElseIf KeyAscii = 99 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H686
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'd Key Behavior
        ElseIf KeyAscii = 100 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'e Key Behavior
        ElseIf KeyAscii = 101 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H639
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'f Key Behavior
        ElseIf KeyAscii = 102 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H641
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'g Key Behavior
        ElseIf KeyAscii = 103 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6AF
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'h Key Behavior
        ElseIf KeyAscii = 104 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6BE
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'i Key Behavior
        ElseIf KeyAscii = 105 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6CC
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'j Key Behavior
        ElseIf KeyAscii = 106 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'k Key Behavior
        ElseIf KeyAscii = 107 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6A9
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'l Key Behavior
        ElseIf KeyAscii = 108 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H644
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'm Key Behavior
        ElseIf KeyAscii = 109 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H645
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'n Key Behavior
        ElseIf KeyAscii = 110 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H646
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'o Key Behavior
        ElseIf KeyAscii = 111 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6C1
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'p Key Behavior
        ElseIf KeyAscii = 112 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H67E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'q Key Behavior
        ElseIf KeyAscii = 113 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H642
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'r Key Behavior
        ElseIf KeyAscii = 114 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H631
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        's Key Behavior
        ElseIf KeyAscii = 115 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H633
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        't Key Behavior
        ElseIf KeyAscii = 116 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'u Key Behavior
        ElseIf KeyAscii = 117 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H621
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'v Key Behavior
        ElseIf KeyAscii = 118 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H637
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'w Key Behavior
        ElseIf KeyAscii = 119 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H648
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'x Key Behavior
        ElseIf KeyAscii = 120 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H634
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'y Key Behavior
        ElseIf KeyAscii = 121 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6D2
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'z Key Behavior
        ElseIf KeyAscii = 122 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H632
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)


        ' For Capital Latter's Behaviors

        'A Key Behavior
        ElseIf KeyAscii = 65 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H622
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'B Key Behavior
        ElseIf KeyAscii = 66 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFBB0
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'C Key Behavior
        ElseIf KeyAscii = 67 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'D Key Behavior
        ElseIf KeyAscii = 68 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H688
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'E Key Behavior
        ElseIf KeyAscii = 69 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H650
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'F Key Behavior
        ElseIf KeyAscii = 70 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H652
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'G Key Behavior
        ElseIf KeyAscii = 71 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H63A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'H Key Behavior
        ElseIf KeyAscii = 72 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'I Key Behavior
        ElseIf KeyAscii = 73 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H649
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'J Key Behavior
        ElseIf KeyAscii = 74 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H636
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'K Key Behavior
        ElseIf KeyAscii = 75 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'L Key Behavior
        ElseIf KeyAscii = 76 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFEFB
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'M Key Behavior
        ElseIf KeyAscii = 77 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H66B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'N Key Behavior
        ElseIf KeyAscii = 78 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6BA
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'O Key Behavior
        ElseIf KeyAscii = 79 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H629
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'P Key Behavior
        ElseIf KeyAscii = 80 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'Q Key Behavior
        ElseIf KeyAscii = 81 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H626
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'R Key Behavior
        ElseIf KeyAscii = 82 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H691
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'S Key Behavior
        ElseIf KeyAscii = 83 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H635
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'T Key Behavior
        ElseIf KeyAscii = 84 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H679
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'U Key Behavior
        ElseIf KeyAscii = 85 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H626
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'V Key Behavior
        ElseIf KeyAscii = 86 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H638
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'W Key Behavior
        ElseIf KeyAscii = 87 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H624
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'Z Key Behavior
        ElseIf KeyAscii = 88 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H698
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'Y Key Behavior
        ElseIf KeyAscii = 89 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFBAF
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'Z Key Behavior
        ElseIf KeyAscii = 90 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H630
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)


        'For Numaric Key's Behaviors

        '0 Key Behavior
        ElseIf KeyAscii = 48 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H660
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '1 Key Behavior
        ElseIf KeyAscii = 49 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H661
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '2 Key Behavior
        ElseIf KeyAscii = 50 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H662
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '3 Key Behavior
        ElseIf KeyAscii = 51 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H663
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '4 Key Behavior
        ElseIf KeyAscii = 52 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H664
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '5 Key Behavior
        ElseIf KeyAscii = 53 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H665
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '6 Key Behavior
        ElseIf KeyAscii = 54 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H666
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '7 Key Behavior
        ElseIf KeyAscii = 55 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H667
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '8 Key Behavior
        ElseIf KeyAscii = 56 Or TextBox1.SelText <> "" Then
        UniCode = &H668
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '9 Key Behavior
        ElseIf KeyAscii = 57 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H669
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        ' Numaric Keys with 'Shift' Behavior

        ') Key Behavior
        ElseIf KeyAscii = 41 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFD3F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '! Key Behavior
        ElseIf KeyAscii = 33 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H21
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '@ Key Behavior
        ElseIf KeyAscii = 64 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H40
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '# Key Behavior
        ElseIf KeyAscii = 35 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H23
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '$ Key Behavior
        ElseIf KeyAscii = 36 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H24
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '% Key Behavior
        ElseIf KeyAscii = 37 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H66A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '^ Key Behavior
        ElseIf KeyAscii = 94 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '& Key Behavior
        ElseIf KeyAscii = 38 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H26
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '* Key Behavior
        ElseIf KeyAscii = 42 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H66D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '( Key Behavior
        ElseIf KeyAscii = 40 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFD3E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)


        'For Special Characters

        'Symbols

        '? Key Behavior
        ElseIf KeyAscii = 63 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H61F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '/ Key Behavior
        ElseIf KeyAscii = 47 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        ', Key Behavior
        ElseIf KeyAscii = 44 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H60C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '. Key Behavior
        ElseIf KeyAscii = 46 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H640
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '_ Key Behavior
        ElseIf KeyAscii = 95 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '- Key Behavior
        ElseIf KeyAscii = 45 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '+ Key Behavior
        ElseIf KeyAscii = 43 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '= Key Behavior
        ElseIf KeyAscii = 61 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H3D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        ': Key Behavior
        ElseIf KeyAscii = 58 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H3A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '; Key Behavior
        ElseIf KeyAscii = 59 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H201C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '< Key Behavior
        ElseIf KeyAscii = 60 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '> Key Behavior
        ElseIf KeyAscii = 62 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H650
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '{ Key Behavior
        ElseIf KeyAscii = 123 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2018
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '} Key Behavior
        ElseIf KeyAscii = 125 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2019
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '[ Key Behavior
        ElseIf KeyAscii = 91 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '] Key Behavior
        ElseIf KeyAscii = 93 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '| Key Behavior
        ElseIf KeyAscii = 124 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H7C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '\ Key Behavior
        ElseIf KeyAscii = 92 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '~ Key Behavior
        ElseIf KeyAscii = 126 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '` Key Behavior
        ElseIf KeyAscii = 96 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '" Key Behavior
        ElseIf KeyAscii = 34 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2190
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '' Key Behavior
        ElseIf KeyAscii = 39 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H201D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        End If
        KeyAscii = 0
  End If

        'This Function Got End There
Exit Sub
ErrorHandler:
   Call ShowErrorMessage


End Sub

Private Sub TxtDepartmentID_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtDepartmentID.Name Then Exit Sub
   If TxtDepartmentName.Text <> "" Then TxtDepartmentName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDepartmentID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtDepartmentID.Name Then Exit Sub
   If TxtDepartmentName.Text <> "" Then Exit Sub
   If TxtDepartmentID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectDepartment(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectDepartment(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Function FunSelectDepartment(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchDepartment.Show vbModal, Me
        If SchDepartment.ParaOutDepartmentID = "" Then FunSelectDepartment = False: Exit Function
        TxtDepartmentID.Text = SchDepartment.ParaOutDepartmentID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Departments where DepartmentID=" & Val(TxtDepartmentID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtDepartmentName.Text = !Department
          FunSelectDepartment = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectDepartment = False
          .Close
          TxtDepartmentID.Text = ""
          TxtDepartmentName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnDepartment_Click()
   On Error GoTo ErrorHandler
   If FunSelectDepartment(ssButton, False) = True Then
     If TxtPubID.Visible Then TxtPubID.SetFocus Else TxtGroupID.SetFocus
   Else
      TxtDepartmentID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

