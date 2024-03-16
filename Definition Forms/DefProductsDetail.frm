VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form DefProductsDetail 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15390
   Icon            =   "DefProductsDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   770
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1026
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtVenderName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6945
      TabIndex        =   127
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox TxtDescriptionName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   12255
      TabIndex        =   126
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox TxtItemDescName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   12255
      TabIndex        =   125
      Top             =   1830
      Width           =   2415
   End
   Begin VB.TextBox TxtSeasonName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   12255
      TabIndex        =   124
      Top             =   2610
      Width           =   2415
   End
   Begin VB.TextBox TxtSubDepartmentName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6945
      TabIndex        =   123
      Top             =   2220
      Width           =   2415
   End
   Begin VB.CheckBox ChkDataNotClear 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Data Not Clear"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5895
      TabIndex        =   106
      Tag             =   "NC"
      Top             =   8910
      Width           =   1455
   End
   Begin VB.CheckBox ChkIsChangedPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Changed Price"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10155
      TabIndex        =   35
      Top             =   8265
      Width           =   1770
   End
   Begin VB.TextBox TxtDepartmentID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5985
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1830
      Width           =   600
   End
   Begin VB.TextBox TxtDepartmentName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6945
      TabIndex        =   103
      Top             =   1830
      Width           =   2415
   End
   Begin VB.CheckBox ChkDeadProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Dead Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11550
      TabIndex        =   40
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CheckBox ChkRawProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Raw Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10155
      TabIndex        =   39
      Top             =   8535
      Width           =   1365
   End
   Begin VB.CheckBox ChkExpiryDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10155
      TabIndex        =   41
      Top             =   7590
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkWSDiscb4ST 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Use WS Price For Discount b4 ST"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7230
      TabIndex        =   34
      Top             =   8265
      Width           =   2820
   End
   Begin VB.OptionButton OptRPSaleTax 
      BackColor       =   &H00EFC09E&
      Caption         =   "SaleTax"
      Height          =   195
      Left            =   7065
      TabIndex        =   78
      Top             =   6240
      Width           =   1020
   End
   Begin VB.OptionButton OptWSPSaleTax 
      BackColor       =   &H00EFC09E&
      Caption         =   "SaleTax"
      Height          =   195
      Left            =   7065
      TabIndex        =   77
      Top             =   5925
      Width           =   1020
   End
   Begin VB.ComboBox CmbSalePacking 
      Height          =   315
      Left            =   5970
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4665
      Width           =   1740
   End
   Begin VB.CheckBox ChkNoCostProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "No Cost Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8655
      TabIndex        =   38
      Top             =   8520
      Width           =   1500
   End
   Begin VB.CheckBox ChkClosingProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Closing Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7230
      TabIndex        =   37
      Top             =   8520
      Width           =   1425
   End
   Begin VB.CheckBox ChkLockProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Lock Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5910
      TabIndex        =   36
      Top             =   8520
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
      TabIndex        =   71
      Top             =   3375
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
         TabIndex        =   72
         Tag             =   "NC"
         Text            =   "DefProductsDetail.frx":0ECA
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
         TabIndex        =   73
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.ComboBox CmbUnits 
      Height          =   315
      Left            =   5970
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   5115
      Width           =   1740
   End
   Begin VB.TextBox TxtFilterID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1140
      TabIndex        =   48
      Top             =   2220
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
      Left            =   1140
      TabIndex        =   47
      Top             =   1815
      Width           =   2655
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   5970
      TabIndex        =   22
      Top             =   5565
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
      Left            =   5970
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4245
      Width           =   1740
   End
   Begin VB.ComboBox CmbFilterGroup 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Tag             =   "nc"
      Top             =   1425
      Width           =   2670
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4095
      Left            =   255
      TabIndex        =   49
      Top             =   2640
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
      stylesets(0).Picture=   "DefProductsDetail.frx":1052
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
      Left            =   7185
      TabIndex        =   43
      Top             =   9390
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
      MICON           =   "DefProductsDetail.frx":106E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8505
      TabIndex        =   44
      Top             =   9390
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
      MICON           =   "DefProductsDetail.frx":108A
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid PGrid 
      Height          =   1095
      Left            =   8355
      TabIndex        =   16
      Top             =   4215
      Width           =   2040
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
      stylesets(0).Picture=   "DefProductsDetail.frx":10A6
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
      Columns(1).Width=   714
      Columns(1).Caption=   "Mul"
      Columns(1).Name =   "Multiplier"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Packing ID"
      Columns(2).Name =   "PackingID"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   17
      Columns(2).FieldLen=   256
      _ExtentX        =   3598
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
      Left            =   6615
      TabIndex        =   54
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2610
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
      MICON           =   "DefProductsDetail.frx":10C2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtRetailPrice 
      Height          =   315
      Left            =   5970
      TabIndex        =   24
      Top             =   6195
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
      Left            =   6015
      TabIndex        =   5
      Top             =   2610
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
      Left            =   8385
      TabIndex        =   17
      Top             =   5565
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
      Left            =   5970
      TabIndex        =   12
      Tag             =   "nc"
      Top             =   3420
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
      Left            =   7095
      TabIndex        =   13
      Top             =   3450
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
      Left            =   8385
      TabIndex        =   45
      Top             =   5880
      Width           =   2055
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
      stylesets(0).Picture=   "DefProductsDetail.frx":10DE
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
      Columns(0).Width=   3122
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
      _ExtentX        =   3625
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
      Left            =   6615
      TabIndex        =   58
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3000
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
      MICON           =   "DefProductsDetail.frx":10FA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   6015
      TabIndex        =   6
      Top             =   3000
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
      Left            =   11880
      TabIndex        =   59
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3000
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
      MICON           =   "DefProductsDetail.frx":1116
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   11280
      TabIndex        =   11
      Top             =   3000
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
      Left            =   5970
      TabIndex        =   26
      Top             =   6825
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
      Left            =   14655
      TabIndex        =   65
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   3000
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
      MICON           =   "DefProductsDetail.frx":1132
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnAddGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9390
      TabIndex        =   66
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   2610
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
      MICON           =   "DefProductsDetail.frx":114E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnAddSubGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9390
      TabIndex        =   67
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   3000
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
      MICON           =   "DefProductsDetail.frx":116A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPacking 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7710
      TabIndex        =   68
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   4245
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
      MICON           =   "DefProductsDetail.frx":1186
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnUnit 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7710
      TabIndex        =   69
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   5115
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
      MICON           =   "DefProductsDetail.frx":11A2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtMaxStockLimit 
      Height          =   315
      Left            =   8625
      TabIndex        =   32
      Top             =   7455
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
      Left            =   7710
      TabIndex        =   75
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   4665
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
      MICON           =   "DefProductsDetail.frx":11BE
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtWSPrice 
      Height          =   315
      Left            =   5970
      TabIndex        =   23
      Top             =   5880
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
      Left            =   5970
      TabIndex        =   25
      Top             =   6510
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
   Begin SITextBox.Txt TxtTokenVal 
      Height          =   315
      Left            =   5970
      TabIndex        =   30
      Top             =   8130
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
      Height          =   315
      Left            =   10155
      TabIndex        =   42
      Tag             =   "NC"
      Top             =   7860
      Visible         =   0   'False
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
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
   Begin SITextBox.Txt TxtDesc1 
      Height          =   315
      Left            =   5940
      TabIndex        =   14
      Top             =   3810
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
      Left            =   5970
      TabIndex        =   28
      Top             =   7455
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
      Left            =   11880
      TabIndex        =   90
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2220
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
      MICON           =   "DefProductsDetail.frx":11DA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBrandID 
      Height          =   315
      Left            =   11280
      TabIndex        =   9
      Top             =   2220
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
      Left            =   14655
      TabIndex        =   91
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   2220
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
      MICON           =   "DefProductsDetail.frx":11F6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   6975
      TabIndex        =   93
      Top             =   2610
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
      Left            =   6975
      TabIndex        =   94
      Top             =   3000
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
      Left            =   12240
      TabIndex        =   95
      Top             =   3000
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
      Left            =   12240
      TabIndex        =   96
      Top             =   2220
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
      Left            =   5565
      TabIndex        =   98
      TabStop         =   0   'False
      ToolTipText     =   "Add New"
      Top             =   3420
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
      MICON           =   "DefProductsDetail.frx":1212
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   10155
      TabIndex        =   18
      Top             =   5565
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
      Left            =   10695
      TabIndex        =   100
      TabStop         =   0   'False
      ToolTipText     =   "Add New"
      Top             =   3405
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
      MICON           =   "DefProductsDetail.frx":122E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtMinStockLimit 
      Height          =   315
      Left            =   8625
      TabIndex        =   31
      Top             =   7140
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
      Left            =   5970
      TabIndex        =   27
      Top             =   7140
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
      Left            =   5970
      TabIndex        =   29
      Top             =   7770
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
      Left            =   8625
      TabIndex        =   33
      Top             =   7770
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
   Begin JeweledBut.JeweledButton BtnAddDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9360
      TabIndex        =   105
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   1830
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
      MICON           =   "DefProductsDetail.frx":124A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubDepartmentID 
      Height          =   315
      Left            =   5985
      TabIndex        =   4
      Top             =   2220
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
   Begin JeweledBut.JeweledButton BtnAddSubDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9360
      TabIndex        =   107
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   2220
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
      MICON           =   "DefProductsDetail.frx":1266
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSeason 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11895
      TabIndex        =   109
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2610
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
      MICON           =   "DefProductsDetail.frx":1282
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSeasonID 
      Height          =   315
      Left            =   11295
      TabIndex        =   10
      Top             =   2610
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
      Left            =   14670
      TabIndex        =   110
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   2610
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
      MICON           =   "DefProductsDetail.frx":129E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnItemDesc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11895
      TabIndex        =   112
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1830
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
      MICON           =   "DefProductsDetail.frx":12BA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtItemDescID 
      Height          =   315
      Left            =   11295
      TabIndex        =   8
      Top             =   1830
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
   Begin JeweledBut.JeweledButton BtnAddItemDesc 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   14670
      TabIndex        =   113
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   1830
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
      MICON           =   "DefProductsDetail.frx":12D6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDescription 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11895
      TabIndex        =   115
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1440
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
      MICON           =   "DefProductsDetail.frx":12F2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDescriptionID 
      Height          =   315
      Left            =   11295
      TabIndex        =   7
      Top             =   1440
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
   Begin JeweledBut.JeweledButton BtnAddDescription 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   14670
      TabIndex        =   116
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   1440
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
      MICON           =   "DefProductsDetail.frx":130E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6585
      TabIndex        =   118
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1830
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
      MICON           =   "DefProductsDetail.frx":132A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSubDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6585
      TabIndex        =   119
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2220
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
      MICON           =   "DefProductsDetail.frx":1346
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnVender 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6585
      TabIndex        =   120
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1440
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
      MICON           =   "DefProductsDetail.frx":1362
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtVenderID 
      Height          =   315
      Left            =   5985
      TabIndex        =   2
      Top             =   1440
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
   Begin JeweledBut.JeweledButton BtnAddVendor 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9360
      TabIndex        =   122
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   1440
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
      MICON           =   "DefProductsDetail.frx":137E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   660
      TabIndex        =   128
      Top             =   9390
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
      MICON           =   "DefProductsDetail.frx":139A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   1980
      TabIndex        =   129
      Top             =   9390
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
      MICON           =   "DefProductsDetail.frx":13B6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   3300
      TabIndex        =   130
      Top             =   9390
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
      MICON           =   "DefProductsDetail.frx":13D2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   5880
      TabIndex        =   131
      Top             =   9390
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
      MICON           =   "DefProductsDetail.frx":13EE
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtItemCode 
      Height          =   315
      Left            =   5985
      TabIndex        =   1
      Top             =   1035
      Width           =   1095
      _ExtentX        =   1931
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
      IntegralPoint   =   5
   End
   Begin JeweledBut.JeweledButton BtnItemCode 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7110
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1035
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
      MICON           =   "DefProductsDetail.frx":140A
      BC              =   14737632
      FC              =   0
   End
   Begin MSComctlLib.ListView LvwColour 
      Height          =   2985
      Left            =   10755
      TabIndex        =   133
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   4545
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   5265
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LvwSize 
      Height          =   2985
      Left            =   12960
      TabIndex        =   134
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   4545
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   5265
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin JeweledBut.JeweledButton BtnColour 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   12600
      TabIndex        =   135
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   4545
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
      MICON           =   "DefProductsDetail.frx":1426
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSize 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   14580
      TabIndex        =   136
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   4545
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
      MICON           =   "DefProductsDetail.frx":1442
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      Left            =   5040
      TabIndex        =   132
      Top             =   1080
      Width           =   870
   End
   Begin VB.Label LblVender 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vender"
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
      Left            =   5280
      TabIndex        =   121
      Top             =   1485
      Width           =   615
   End
   Begin VB.Label LblOther 
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
      Left            =   10230
      TabIndex        =   117
      Top             =   1485
      Width           =   975
   End
   Begin VB.Label LblItemDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Description"
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
      Left            =   9810
      TabIndex        =   114
      Top             =   1875
      Width           =   1395
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
      Left            =   10560
      TabIndex        =   111
      Top             =   2655
      Width           =   645
   End
   Begin VB.Label LblSubDepartment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Department"
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
      Left            =   4515
      TabIndex        =   108
      Top             =   2265
      Width           =   1380
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
      Left            =   4905
      TabIndex        =   104
      Top             =   1875
      Width           =   990
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
      Left            =   7155
      TabIndex        =   102
      Top             =   7815
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
      Left            =   7230
      TabIndex        =   101
      Top             =   7185
      Width           =   1320
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Enabled         =   0   'False
      Height          =   195
      Left            =   10155
      TabIndex        =   99
      Top             =   5340
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label LblName2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name 2"
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
      Left            =   5205
      TabIndex        =   97
      Top             =   3810
      Width           =   660
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   435
      Left            =   5925
      TabIndex        =   15
      ToolTipText     =   "Textbox1"
      Top             =   3765
      Width           =   4785
      VariousPropertyBits=   752896027
      ForeColor       =   0
      BorderStyle     =   1
      Size            =   "8440;767"
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
      Left            =   10620
      TabIndex        =   92
      Top             =   2265
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
      Left            =   4950
      TabIndex        =   89
      Top             =   7530
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
      Left            =   4890
      TabIndex        =   88
      Top             =   3915
      Width           =   975
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Comm%"
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
      Left            =   4950
      TabIndex        =   87
      Top             =   7815
      Width           =   1080
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
      Left            =   5100
      TabIndex        =   86
      Top             =   5610
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
      Left            =   4890
      TabIndex        =   85
      Top             =   6240
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
      Left            =   4770
      TabIndex        =   84
      Top             =   7185
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
      Left            =   4860
      TabIndex        =   83
      Top             =   6870
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
      Left            =   5100
      TabIndex        =   82
      Top             =   5925
      Width           =   810
   End
   Begin VB.Label Label20 
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
      Left            =   5010
      TabIndex        =   81
      Top             =   6555
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
      Left            =   5010
      TabIndex        =   80
      Top             =   8175
      Width           =   900
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
      Left            =   10380
      TabIndex        =   79
      Top             =   7590
      Visible         =   0   'False
      Width           =   990
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
      Left            =   4770
      TabIndex        =   76
      Top             =   4695
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
      Left            =   7230
      TabIndex        =   74
      Top             =   7455
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
      TabIndex        =   70
      Top             =   540
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
      Left            =   5460
      TabIndex        =   64
      Top             =   5145
      Width           =   450
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Products Detail"
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
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   63
      Top             =   315
      Width           =   1995
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
      Left            =   5040
      TabIndex        =   62
      Top             =   3030
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
      Left            =   4620
      TabIndex        =   61
      Top             =   3465
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
      Left            =   10440
      TabIndex        =   60
      Top             =   3045
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
      Left            =   525
      TabIndex        =   57
      Top             =   2265
      Width           =   570
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Codes"
      Enabled         =   0   'False
      Height          =   195
      Left            =   8385
      TabIndex        =   56
      Top             =   5340
      Width           =   735
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   14310
      Top             =   270
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
      Left            =   5430
      TabIndex        =   55
      Top             =   2640
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
      Left            =   4860
      TabIndex        =   53
      Top             =   4275
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   297
      X2              =   297
      Y1              =   120
      Y2              =   448
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
      Left            =   480
      TabIndex        =   52
      Top             =   1860
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
      Left            =   450
      TabIndex        =   51
      Top             =   1515
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
      Left            =   6600
      TabIndex        =   50
      Top             =   3480
      Width           =   495
   End
End
Attribute VB_Name = "DefProductsDetail"
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

Private Sub BtnAddBrand_Click()
   On Error GoTo ErrorHandler
   DefBrands.Show vbModal, Me
   If TxtBrandID.Visible Then TxtBrandID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddCompany_Click()
   On Error GoTo ErrorHandler
   DefCompanies.Show vbModal, Me
   If TxtCompanyID.Visible Then TxtCompanyID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddDepartment_Click()
   On Error GoTo ErrorHandler
   DefDepartments.Show vbModal, Me
   If TxtDepartmentID.Visible Then TxtDepartmentID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddGroup_Click()
   On Error GoTo ErrorHandler
   DefGroups.Show vbModal, Me
   With cn.Execute("Select * FROM Groups")
      CmbFilterGroup.Clear
      CmbFilterGroup.AddItem ""
      Do Until .EOF
         CmbFilterGroup.AddItem !GroupName
         CmbFilterGroup.ItemData(CmbFilterGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
    End With
   If TxtGroupID.Visible Then TxtGroupID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAddSubGroup_Click()
   On Error GoTo ErrorHandler
   DefSubGroups.Show vbModal, Me
   If TxtSubGroupID.Visible Then TxtSubGroupID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnBrand_Click()
   On Error GoTo ErrorHandler
   If FunSelectBrand(ssButton, False) = True Then
      If TxtSeasonID.Visible Then TxtSeasonID.SetFocus Else TxtBrandID.SetFocus
   Else
      If TxtBrandID.Enabled Then TxtBrandID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnColour_Click()
   On Error GoTo ErrorHandler
   DefColours.Show vbModal
   PopulateColour
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnCompany_Click()
   On Error GoTo ErrorHandler
   If FunSelectCompany(ssButton, False) = True Then
     If TxtName.Visible Then TxtName.SetFocus Else TxtCompanyID.SetFocus
    Else
     If TxtCompanyID.Enabled Then TxtCompanyID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnDepartment_Click()
   On Error GoTo ErrorHandler
   If FunSelectDepartment(ssButton, False) = True Then
     If TxtSubDepartmentID.Visible Then TxtSubDepartmentID.SetFocus Else TxtDepartmentID.SetFocus
   Else
      TxtDepartmentID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnGroup_Click()
   On Error GoTo ErrorHandler
   If FunSelectGroup(ssButton, False) = True Then
      If TxtSubGroupID.Visible Then TxtSubGroupID.SetFocus Else TxtGroupID.SetFocus
   Else
      If TxtGroupID.Enabled Then TxtGroupID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnItemCode_Click()
   On Error GoTo ErrorHandler
   If FunSelectItemCode(ssButton, False) = True Then
     If TxtVenderID.Visible And TxtVenderID.Enabled Then TxtVenderID.SetFocus Else TxtItemCode.SetFocus
   Else
     If TxtItemCode.Enabled Then TxtItemCode.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnItemDesc_Click()
   On Error GoTo ErrorHandler
   If FunSelectItemDesc(ssButton, False) = True Then
     If TxtBrandID.Visible Then TxtBrandID.SetFocus Else TxtItemDescID.SetFocus
   Else
      TxtItemDescID.SetFocus
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

Private Sub BtnDescription_Click()
   On Error GoTo ErrorHandler
   If FunSelectDescription(ssButton, False) = True Then
     If TxtItemDescID.Visible Then TxtItemDescID.SetFocus Else TxtDescriptionID.SetFocus
   Else
      TxtDescriptionID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPacking_Click()
   On Error GoTo ErrorHandler
   DefPackings.Show vbModal
   With cn.Execute("Select * FROM Packings")
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
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSalePacking_Click()
   On Error GoTo ErrorHandler
   DefPackings.Show vbModal
   With cn.Execute("Select * FROM Packings")
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
     If TxtCompanyID.Visible Then TxtCompanyID.SetFocus Else TxtSeasonID.SetFocus
   Else
      TxtSeasonID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSize_Click()
   On Error GoTo ErrorHandler
   DefSizes.Show vbModal
   PopulateSize
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSubDepartment_Click()
   On Error GoTo ErrorHandler
   If FunSelectSubDepartment(ssButton, False) = True Then
     If TxtGroupID.Visible Then TxtGroupID.SetFocus Else TxtSubDepartmentID.SetFocus
   Else
      TxtSubDepartmentID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSubGroup_Click()
   On Error GoTo ErrorHandler
   If FunSelectSubGroup(ssButton, False) = True Then
      If TxtDescriptionID.Visible Then TxtDescriptionID.SetFocus Else TxtSubGroupID.SetFocus
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
   With cn.Execute("Select * FROM Units")
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
    With cn.Execute(vStrSQL)
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
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtGroupName.Text = !GroupName
          CmbFilterGroup.Text = !GroupName
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
      With cn.Execute("Select * FROM Groups")
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
    With cn.Execute(vStrSQL)
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
    With cn.Execute(vStrSQL)
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

Private Sub BtnAddVendor_Click()
   On Error GoTo ErrorHandler
   DefVendors.Show vbModal, Me
   If TxtVenderID.Visible Then TxtVenderID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CGrid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   RsCode.Filter = "Code='" & CGrid.Columns("Code").Text & "'"
   If RsCode.RecordCount = 1 And CGrid.Columns("Code").Text <> "" Then
      If vIsNewRecord = False Then cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Deleted BarCode-" & CGrid.Columns("Code").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CGrid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   If Flag Then
      TxtCode.Text = CGrid.Columns("Code").Text
      TxtQty.Text = CGrid.Columns("Qty").Text
      If CGrid.Rows = 1 Then CGrid.MoveLast
   End If
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
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbSalePacking_Click()
   On Error GoTo ErrorHandler
   If CmbSalePacking.Visible = False Then Exit Sub
   If Me.ActiveControl.Name <> CmbSalePacking.Name Then Exit Sub
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
       Rs1.Open "Select * FROM Products Order By ProductName", cn, adOpenStatic, adLockOptimistic
   Else
       Rs1.Open "Select * FROM Products Where GroupID = '" & GetGroupID(CmbFilterGroup) & "' Order By ProductName", cn, adOpenStatic, adLockOptimistic
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
         Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then If TxtGroupID.Enabled Then TxtGroupID.SetFocus Else If TxtCompanyID.Enabled Then TxtCompanyID.SetFocus
         Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then If TxtSubGroupID.Enabled Then TxtSubGroupID.SetFocus Else If TxtGroupID.Enabled Then TxtGroupID.SetFocus
         Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then If TxtBrandID.Enabled Then TxtBrandID.SetFocus Else If TxtSubGroupID.Enabled Then TxtSubGroupID.SetFocus
         Case TxtBrandID.Name: If FunSelectBrand(ssFunctionKey, True) = True Then If TxtID.Enabled Then TxtID.SetFocus Else If TxtBrandID.Enabled Then TxtBrandID.SetFocus
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, True) = True Then If TxtDepartmentID.Enabled Then TxtDepartmentID.SetFocus Else If TxtVenderID.Enabled Then TxtVenderID.SetFocus
         Case TxtDepartmentID.Name: If FunSelectDepartment(ssFunctionKey, True) = True Then If TxtSubDepartmentID.Enabled Then TxtSubDepartmentID.SetFocus Else If TxtDepartmentID.Enabled Then TxtDepartmentID.SetFocus
         Case TxtSubDepartmentID.Name: If FunSelectSubDepartment(ssFunctionKey, True) = True Then If TxtSeasonID.Enabled Then TxtSeasonID.SetFocus Else If TxtSubDepartmentID.Enabled Then TxtSubDepartmentID.SetFocus
         Case TxtSeasonID.Name: If FunSelectSeason(ssFunctionKey, True) = True Then If TxtItemDescID.Enabled Then TxtItemDescID.SetFocus Else If TxtSeasonID.Enabled Then TxtSeasonID.SetFocus
         Case TxtItemDescID.Name: If FunSelectItemDesc(ssFunctionKey, True) = True Then If TxtDescriptionID.Enabled Then TxtDescriptionID.SetFocus Else If TxtItemDescID.Enabled Then TxtItemDescID.SetFocus
         Case TxtDescriptionID.Name: If FunSelectDescription(ssFunctionKey, True) = True Then If TxtID.Enabled Then TxtID.SetFocus Else If TxtDescriptionID.Enabled Then TxtDescriptionID.SetFocus
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
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   On Error GoTo ErrorHandler
   Set Rs1 = New ADODB.Recordset
   Rs1.Open "Select * FROM Products Order By ProductName", cn, adOpenStatic, adLockOptimistic
   Set Grid.DataSource = Rs1
   FormStatus = SelectionMode
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  On Error GoTo ErrorHandler
  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniProductsDetail", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
    Call ActivityLog("Products", eDelete, , , vid)
    vProductID = TxtID.Text  'TxtPrefix.Text & TxtID.Text
    If cn.Execute("Select * from ProductPacking where ProductID='" & vProductID & "'").RecordCount > 0 Then
       cn.Execute ("Delete from ProductPacking where ProductID='" & vProductID & "'")
    End If
    If cn.Execute("Select * from ProductBarcodes where ProductID='" & vProductID & "'").RecordCount > 0 Then
       cn.Execute ("Delete from ProductBarCodes where ProductID='" & vProductID & "'")
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
    
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniProductsDetail", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If FunValidation = False Then Exit Sub
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
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

   cn.BeginTrans
   If vIsNewRecord = True And cn.Execute("select count(*) from products where Productid = '" & TxtID.Text & "'").Fields(0) > 0 Then
'      MsgBox "This ID already exists. A new ID has been generated. Please save again", vbExclamation, "Alert"
      TxtID.Text = FunGetMaxID
'      TxtID.SetFocus
'      Exit Function
   End If


   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open "Select * FROM Products order by ProductName", cn, adOpenDynamic, adLockPessimistic
   Rs.Filter = "ProductID = '" & TxtID.Text & "'"
   If vIsNewRecord = False Then Call ActivityLog("Products", eEdit, , , TxtID.Text)
   If vIsNewRecord And Rs.RecordCount = 0 Then
      Rs.AddNew
      Rs!Productid = TxtID.Text
      Rs!PurPrice = Val(TxtPurPrice.Text)
      Rs!isChanged = 0
   Else
      Rs!isChanged = 1
      Rs!IsSync = 0
   End If
   Rs!ItemCode = TxtItemCode.Text
   If ObjUserSecurity.IsAdministrator = True Then Rs!PurPrice = Val(TxtPurPrice.Text)
   Rs!ProductName = TxtName.Text
   Rs!ProductName1 = IIf(Trim(TextBox1.Text) = "", Null, TextBox1.Text)
   Rs!companyid = IIf(Trim(TxtCompanyID.Text) = "", Null, TxtCompanyID.Text)
   Rs!OrganizationID = Null
   Rs!GroupID = TxtGroupID.Text
   Rs!SubGroupID = IIf(Trim(TxtSubGroupID.Text) = "", Null, TxtSubGroupID.Text)
   Rs!BrandID = IIf(Trim(TxtBrandID.Text) = "", Null, TxtBrandID.Text)
   Rs!PubID = Null
   Rs!DepartmentID = IIf(Trim(TxtDepartmentID.Text) = "", Null, TxtDepartmentID.Text)
   Rs!SubDepartmentID = IIf(Trim(TxtSubDepartmentID.Text) = "", Null, TxtSubDepartmentID.Text)
   Rs!SeasonID = IIf(Trim(TxtSeasonID.Text) = "", Null, TxtSeasonID.Text)
   Rs!ItemDescID = IIf(Trim(TxtItemDescID.Text) = "", Null, TxtItemDescID.Text)
   Rs!DescriptionID = IIf(Trim(TxtDescriptionID.Text) = "", Null, TxtDescriptionID.Text)
   Rs!VendorID1 = IIf(Trim(TxtVenderID.Text) = "", Null, TxtVenderID.Text)
   Rs!PurchasePackingID = IIf(CmbPurPacking.Text = "", Null, CmbPurPacking.ItemData(CmbPurPacking.ListIndex))
   Rs!SalePackingID = IIf(CmbSalePacking.Text = "", Null, CmbSalePacking.ItemData(CmbSalePacking.ListIndex))
   Rs!UnitID = IIf(CmbUnits.Text = "", Null, CmbUnits.ItemData(CmbUnits.ListIndex))
   Rs!WSPrice = Val(TxtWSPrice.Text)
   Rs!RetailPrice = Val(TxtRetailPrice.Text)
   Rs!DiscPC = Val(TxtSaleDisc.Text)
   Rs!DiscPer = Val(TxtSaleDiscPer.Text)
   Rs!PurDiscPC = Val(TxtPurDisc.Text)
   Rs!MinStockLimit = IIf(Val(TxtMinStockLimit.Text) = 0, Null, TxtMinStockLimit.Text)
   Rs!MaxStockLimit = IIf(Val(TxtMaxStockLimit.Text) = 0, Null, TxtMaxStockLimit.Text)
   Rs!ServiceCharges = IIf(Val(TxtServiceCharges.Text) = 0, Null, TxtServiceCharges.Text)
   Rs!EmpComm = IIf(Val(TxtEmpComm.Text) = 0, Null, TxtEmpComm.Text)
   Rs!TokenVal = IIf(Val(TxtTokenVal.Text) = 0, Null, TxtTokenVal.Text)
   Rs!SaleTaxPer = IIf(Val(TxtSaleTaxPer.Text) = 0, Null, TxtSaleTaxPer.Text)
   Rs!IsLocked = ChkLockProduct.Value
   Rs!IsDeadProduct = ChkDeadProduct.Value
   Rs!IsNoCostProduct = ChkNoCostProduct.Value
   Rs!IsClosingProduct = ChkClosingProduct.Value
   Rs!isChangedPrice = ChkIsChangedPrice.Value
   Rs!IsRawProduct = ChkRawProduct.Value
   Rs!IsWSSaleTax = OptWSPSaleTax.Value
   Rs!IsRetailSaleTax = OptRPSaleTax.Value
   Rs!IsWSDiscb4ST = ChkWSDiscb4ST.Value
   Rs!Desc1 = IIf(Trim(TxtDesc1.Text) = "", Null, Trim(TxtDesc1.Text))
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
'   RsProductPacking.UpdateBatch
   '' by Farhan
   UpdateColour
   UpdateSize
'   With RsProductPacking
'      .Filter = 0
'      If .RecordCount > 0 Then .MoveFirst
'      For vCounter = 1 To .RecordCount
'         !Productid = TxtID.Text
'         .MoveNext
'      Next vCounter
'      .UpdateBatch
'   End With
   If vIsNewRecord = True Then cn.Execute ("insert into productbarcodes (Productid, Code, IsChanged, Qty) values ('" & TxtID.Text & "','" & TxtItemCode.Text & "',Null,Null)")
   If vIsNewRecord = True Then Call ActivityLog("Products", eAdd, , , TxtID.Text)
   
   cn.CommitTrans
   Rs.ReQuery
   Rs1.ReQuery
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
'   If vIsNewRecord = True And cn.Execute("select count(*) from products where ItemCode = '" & TxtItemCode.Text & "'").Fields(0) > 0 Then
'      MsgBox "Item Code Already Exists.", vbExclamation, "Alert"
'      TxtItemCode.Text = ""
'      TxtItemCode.SetFocus
'      Exit Function
'   End If
'   If TxtItemCode.Text = "" Then
'      MsgBox "Please select a ItemCode", vbExclamation, "Alert"
'      If TxtItemCode.Enabled And TxtItemCode.Visible Then TxtItemCode.SetFocus
'      Exit Function
'   End If
'   If Len(Trim(TxtItemCode.Text)) <> 9 Then
'      MsgBox "Please enter only 9 digit", vbExclamation, "Alert"
'      If TxtItemCode.Enabled And TxtItemCode.Visible Then TxtItemCode.SetFocus
'      Exit Function
'   End If
'   If TxtVenderID.Text = "" Then
'      MsgBox "Please select a Vendor", vbExclamation, "Alert"
'      If TxtVenderID.Enabled And TxtVenderID.Visible Then TxtVenderID.SetFocus
'      Exit Function
'   End If
   
'   If TxtDepartmentID.Text = "" Then
'      MsgBox "Please select a Department", vbExclamation, "Alert"
'      If TxtDepartmentID.Enabled And TxtDepartmentID.Visible Then TxtDepartmentID.SetFocus
'      Exit Function
'   End If
'   If TxtSubDepartmentID.Text = "" Then
'      MsgBox "Please select a Sub Department", vbExclamation, "Alert"
'      If TxtSubDepartmentID.Enabled And TxtSubDepartmentID.Visible Then TxtSubDepartmentID.SetFocus
'      Exit Function
'   End If
'   If TxtSubDepartmentID.Text = "" Then
'      MsgBox "Please select a Sub Department", vbExclamation, "Alert"
'      If TxtSubDepartmentID.Enabled And TxtSubDepartmentID.Visible Then TxtSubDepartmentID.SetFocus
'      Exit Function
'   End If
   If TxtGroupID.Text = "" Then
      MsgBox "Please select a Group", vbExclamation, "Alert"
      If TxtGroupID.Enabled And TxtGroupID.Visible Then TxtGroupID.SetFocus
      Exit Function
   End If
'   If TxtSubGroupID.Text = "" Then
'      MsgBox "Please select a Sub Group", vbExclamation, "Alert"
'      If TxtSubGroupID.Enabled And TxtSubGroupID.Visible Then TxtSubGroupID.SetFocus
'      Exit Function
'   End If
'   If TxtDescriptionID.Text = "" Then
'      MsgBox "Please select a Description", vbExclamation, "Alert"
'      If TxtDescriptionID.Enabled And TxtDescriptionID.Visible Then TxtDescriptionID.SetFocus
'      Exit Function
'   End If
'   If TxtItemDescID.Text = "" Then
'      MsgBox "Please select a Item Description", vbExclamation, "Alert"
'      If TxtItemDescID.Enabled And TxtItemDescID.Visible Then TxtItemDescID.SetFocus
'      Exit Function
'   End If
'   If TxtSeasonID.Text = "" Then
'      MsgBox "Please select a Season", vbExclamation, "Alert"
'      If TxtSeasonID.Enabled And TxtSeasonID.Visible Then TxtSeasonID.SetFocus
'      Exit Function
'   End If
'   If TxtBrandID.Text = "" Then
'      MsgBox "Please select a Brand", vbExclamation, "Alert"
'      If TxtBrandID.Enabled And TxtBrandID.Visible Then TxtBrandID.SetFocus
'      Exit Function
'   End If
'   If TxtCompanyID.Text = "" Then
'      MsgBox "Please select a Company", vbExclamation, "Alert"
'      If TxtCompanyID.Enabled And TxtCompanyID.Visible Then TxtCompanyID.SetFocus
'      Exit Function
'   End If
   i = 0
   Flag = False
   While i <> LvwColour.ListItems.Count
      i = i + 1
      If LvwColour.ListItems(i).Checked = True Then
         i = LvwColour.ListItems.Count
         Flag = True
      End If
   Wend
   If Flag = False Then
      MsgBox "Please select atleast one Colour", vbExclamation, "Alert"
      Exit Function
   End If
   
   i = 0
   Flag = False
   While i <> LvwSize.ListItems.Count
      i = i + 1
      If LvwSize.ListItems(i).Checked = True Then
         i = LvwSize.ListItems.Count
         Flag = True
      End If
   Wend
   If Flag = False Then
      MsgBox "Please select atleast one Size", vbExclamation, "Alert"
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
   
   LvwColour.FullRowSelect = True
   LvwColour.ListItems.Clear
   LvwColour.ColumnHeaders.Add , , "ID", 30, 0
   LvwColour.ColumnHeaders.Add , , "Colour", 70, 0
   LvwColour.View = lvwReport
   
   
   
   LvwSize.FullRowSelect = True
   LvwSize.ListItems.Clear
   LvwSize.ColumnHeaders.Add , , "ID", 30, 0
   LvwSize.ColumnHeaders.Add , , "Size", 70, 0
   LvwSize.View = lvwReport
   
   
   LblDepartment.Visible = ObjRegistry.isShowDepartment
   TxtDepartmentID.Visible = ObjRegistry.isShowDepartment
   TxtDepartmentName.Visible = ObjRegistry.isShowDepartment
   BtnDepartment.Visible = ObjRegistry.isShowDepartment
   BtnAddDepartment.Visible = ObjRegistry.isShowDepartment
   
   LblSubDepartment.Visible = ObjRegistry.isShowSubDepartment
   TxtSubDepartmentID.Visible = ObjRegistry.isShowSubDepartment
   TxtSubDepartmentName.Visible = ObjRegistry.isShowSubDepartment
   BtnSubDepartment.Visible = ObjRegistry.isShowSubDepartment
   BtnAddSubDepartment.Visible = ObjRegistry.isShowSubDepartment
   
   LblSeason.Visible = ObjRegistry.isShowSeason
   TxtSeasonID.Visible = ObjRegistry.isShowSeason
   TxtSeasonName.Visible = ObjRegistry.isShowSeason
   BtnSeason.Visible = ObjRegistry.isShowSeason
   BtnAddSeason.Visible = ObjRegistry.isShowSeason
   
   LblItemDesc.Visible = ObjRegistry.isShowItemDesc
   TxtItemDescID.Visible = ObjRegistry.isShowItemDesc
   TxtItemDescName.Visible = ObjRegistry.isShowItemDesc
   BtnItemDesc.Visible = ObjRegistry.isShowItemDesc
   BtnAddItemDesc.Visible = ObjRegistry.isShowItemDesc
   
   LblDescription.Visible = ObjRegistry.isShowOther
   TxtDescriptionID.Visible = ObjRegistry.isShowOther
   TxtDescriptionName.Visible = ObjRegistry.isShowOther
   BtnDescription.Visible = ObjRegistry.isShowOther
   BtnAddDescription.Visible = ObjRegistry.isShowOther
   
   LblVender.Visible = ObjRegistry.isShowVendor
   TxtVenderID.Visible = ObjRegistry.isShowVendor
   TxtVenderName.Visible = ObjRegistry.isShowVendor
   BtnVender.Visible = ObjRegistry.isShowVendor
   BtnAddVendor.Visible = ObjRegistry.isShowVendor
   
   CmbFilterGroup.AddItem ""
   If ObjRegistry.HeaderInfoNotClear = True Then
      TxtCompanyID.Tag = "NC"
      TxtCompanyName.Tag = "NC"
      TxtGroupID.Tag = "NC"
      TxtGroupName.Tag = "NC"
      TxtSubGroupID.Tag = "NC"
      TxtSubGroupName.Tag = "NC"
      TxtBrandID.Tag = "NC"
      TxtBrandName.Tag = "NC"
      TxtDepartmentID.Tag = "NC"
      TxtDepartmentName.Tag = "NC"
   End If
   With cn.Execute("Select * FROM Groups")
      Do Until .EOF
         CmbFilterGroup.AddItem !GroupName
         CmbFilterGroup.ItemData(CmbFilterGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
   End With
   With cn.Execute("Select * FROM Packings")
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
   With cn.Execute("Select * FROM Units")
      CmbUnits.AddItem ""
      Do Until .EOF
         CmbUnits.AddItem !UnitName
         CmbUnits.ItemData(CmbUnits.NewIndex) = !UnitID
         .MoveNext
      Loop
   End With
   isFreeCode = False
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open "Select * FROM Products order by ProductName", cn, adOpenStatic, adLockOptimistic
   If RsProductPacking.State = adStateOpen Then RsProductPacking.Close
   RsProductPacking.Open "Select * from ProductPacking", cn, adOpenStatic, adLockBatchOptimistic
   If CmbFilterGroup.ListCount > 0 Then
      CmbFilterGroup.ListIndex = 0
   End If
   vPer = 0
   FormStatus = NewMode
   
   LblName2.Visible = ObjRegistry.AllowUrduProduct
   TextBox1.Visible = ObjRegistry.AllowUrduProduct
   LblDescription.Visible = Not ObjRegistry.AllowUrduProduct
   TxtDesc1.Visible = Not ObjRegistry.AllowUrduProduct
   
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

   ModeValue = False
   
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Property Get FormStatus() As FormMode
   On Error GoTo ErrorHandler
   'Nothing
   FormStatus = vMode
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
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
      If TxtGroupName.Text <> "" Then CmbFilterGroup.Text = TxtGroupName.Text
      If Rs.State = adStateOpen Then Rs.Close
      Rs.Open "Select * FROM Products order by ProductName", cn, adOpenStatic, adLockOptimistic
      
      If ChkDataNotClear.Value = 0 Then
      If RsProductPacking.State = adStateOpen Then RsProductPacking.Close
      RsProductPacking.Open "Select * from ProductPacking", cn, adOpenStatic, adLockBatchOptimistic
      
      
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
      If CmbFilterGroup.ListCount > 0 Then
         CmbFilterGroup.ListIndex = 0
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
      Call SubClearFields(True)
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
      TxtItemCode.Enabled = False
      BtnItemCode.Enabled = False
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
      Call Grid_RowColChange(0, 0)
      Grid.Enabled = True
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
   RsPacking.Open vPackSql, cn, adOpenStatic, adLockOptimistic
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
   RsCode.Open vSQL, cn, adOpenStatic, adLockBatchOptimistic
   CGrid.RemoveAll
   RsCode.Filter = 0
   While Not RsCode.EOF
      CGrid.AddNew
      CGrid.Columns("Code").Text = RsCode!code
      CGrid.Columns("Qty").Text = IIf(IsNull(RsCode!Qty), "", RsCode!Qty)
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
 
Private Sub UpdateColour()
   On Error GoTo ErrorHandler
   Dim vSQL As String
   vSQL = "Delete from ProductColours where ProductID = '" & TxtID.Text & "'"
   cn.Execute vSQL
   For i = 1 To LvwColour.ListItems.Count
      If LvwColour.ListItems(i).Checked = True Then
         vSQL = "insert into ProductColours (ProductID, ColourID) Values ('" & TxtID.Text & "','" & LvwColour.ListItems(i).Text & "')"
         cn.Execute vSQL
      End If
   Next i
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateColour()
   On Error GoTo ErrorHandler
   Dim vSQL As String
   Dim vrowcount As Integer
   vSQL = "Select c.ColourID, ColourName, case when pc.ColourID is not null then 1 else 0 end as Checked from Colours c" _
            + " Left outer join (select * from  ProductColours where ProductID = '" & TxtID.Text & "')PC on c.ColourID = pc.ColourID"
            
   LvwColour.ListItems.Clear
   With cn.Execute(vSQL)
      While Not .EOF
         Set Item = LvwColour.ListItems.Add(, , !ColourID)
         Item.SubItems(1) = !ColourName
         Item.Checked = !Checked
         .MoveNext
      Wend
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
 End Sub

Private Sub UpdateSize()
   On Error GoTo ErrorHandler
   Dim vSQL As String
   vSQL = "Delete from ProductSizes where ProductID = '" & TxtID.Text & "'"
   cn.Execute vSQL
   For i = 1 To LvwSize.ListItems.Count
      If LvwSize.ListItems(i).Checked = True Then
         vSQL = "insert into ProductSizes (ProductID, SizeID) Values ('" & TxtID.Text & "','" & LvwSize.ListItems(i).Text & "')"
         cn.Execute vSQL
      End If
   Next i
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
   
Private Sub PopulateSize()
   On Error GoTo ErrorHandler
   Dim vSQL As String
   Dim vrowcount As Integer
   vSQL = "Select s.SizeID, SizeName, case when ps.SizeID is not null then 1 else 0 end as Checked from Sizes s" _
            + " Left outer join (select * from  ProductSizes where ProductID = '" & TxtID.Text & "')PS on s.SizeID = ps.SizeID"
            
   LvwSize.ListItems.Clear
   With cn.Execute(vSQL)
      While Not .EOF
         Set Item = LvwSize.ListItems.Add(, , !SizeID)
         Item.SubItems(1) = !SizeName
         Item.Checked = !Checked
         .MoveNext
      Wend
   End With
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
      If Rs1.RecordCount > 0 And Grid.Enabled Then
         TxtItemCode.Text = IIf(IsNull(Rs1!ItemCode), "", Rs1!ItemCode)
         
         TxtCompanyID.Text = IIf(IsNull(Rs1!companyid), "", Rs1!companyid)
         If Trim(TxtCompanyID.Text) <> "" Then
            TxtCompanyName.Text = cn.Execute("Select CompanyName from Companies Where Companyid='" & TxtCompanyID.Text & "'").Fields(0)
         Else
            TxtCompanyName.Text = ""
         End If
         TxtGroupID.Text = Rs1!GroupID
         TxtGroupName.Text = cn.Execute("Select GroupName from Groups Where Groupid='" & TxtGroupID.Text & "'").Fields(0)
         TxtSubGroupID.Text = IIf(IsNull(Rs1!SubGroupID), "", Rs1!SubGroupID)
         If Trim(TxtSubGroupID.Text) <> "" Then
            TxtSubGroupName.Text = cn.Execute("Select SubGroupName from SubGroups Where SubGroupid='" & TxtSubGroupID.Text & "'").Fields(0)
         Else
            TxtSubGroupName.Text = ""
         End If
         
         TxtDepartmentID.Text = IIf(IsNull(Rs1!DepartmentID), "", Rs1!DepartmentID)
         TxtSubDepartmentID.Text = IIf(IsNull(Rs1!SubDepartmentID), "", Rs1!SubDepartmentID)
         TxtSeasonID.Text = IIf(IsNull(Rs1!SeasonID), "", Rs1!SeasonID)
         TxtItemDescID.Text = IIf(IsNull(Rs1!ItemDescID), "", Rs1!ItemDescID)
         TxtDescriptionID.Text = IIf(IsNull(Rs1!DescriptionID), "", Rs1!DescriptionID)
         TxtVenderID.Text = IIf(IsNull(Rs1!VendorID1), "", Rs1!VendorID1)
         If Trim(TxtDepartmentID.Text) <> "" Then
            TxtDepartmentName.Text = cn.Execute("Select Department from Departments Where Departmentid='" & TxtDepartmentID.Text & "'").Fields(0)
         Else
            TxtDepartmentName.Text = ""
         End If
         If Trim(TxtSubDepartmentID.Text) <> "" Then
            TxtSubDepartmentName.Text = cn.Execute("Select SubDepartmentName from SubDepartments Where SubDepartmentid='" & TxtSubDepartmentID.Text & "'").Fields(0)
         Else
            TxtSubDepartmentName.Text = ""
         End If
         If Trim(TxtSeasonID.Text) <> "" Then
            TxtSeasonName.Text = cn.Execute("Select SeasonName from Seasons Where Seasonid='" & TxtSeasonID.Text & "'").Fields(0)
         Else
            TxtSeasonName.Text = ""
         End If
         If Trim(TxtItemDescID.Text) <> "" Then
            TxtItemDescName.Text = cn.Execute("Select ItemDescName from ItemDescription Where ItemDescid='" & TxtItemDescID.Text & "'").Fields(0)
         Else
            TxtItemDescName.Text = ""
         End If
         If Trim(TxtDescriptionID.Text) <> "" Then
            TxtDescriptionName.Text = cn.Execute("Select DescriptionName from Descriptions Where DescriptionID='" & TxtDescriptionID.Text & "'").Fields(0)
         Else
            TxtDescriptionName.Text = ""
         End If
         If Trim(TxtVenderID.Text) <> "" Then
            TxtVenderName.Text = cn.Execute("Select partyname from Parties Where partyid='" & TxtVenderID.Text & "'").Fields(0)
         Else
            TxtVenderName.Text = ""
         End If
         TxtBrandID.Text = IIf(IsNull(Rs1!BrandID), "", Rs1!BrandID)
         If Trim(TxtBrandID.Text) <> "" Then
            TxtBrandName.Text = cn.Execute("Select BrandName from Brands Where BrandID = '" & TxtBrandID.Text & "'").Fields(0)
         Else
            TxtBrandName.Text = ""
         End If
         TxtID.Text = Grid.Columns("ID").Text
         TxtName.Text = Grid.Columns("Name").Text
         '''''''''''''''''''''''''''''''''''''''''''''
         '===================================================
         With cn.Execute("Select PackingName from Packings Where PackingID = " & IIf(IsNull(Rs1!PurchasePackingID), "0", Rs1!PurchasePackingID))
            If .RecordCount = 0 Then
               CmbPurPacking.ListIndex = 0
            Else
               CmbPurPacking.Text = !PackingName
            End If
         End With
         With cn.Execute("Select PackingName from Packings Where PackingID = " & IIf(IsNull(Rs1!SalePackingID), "0", Rs1!SalePackingID))
            If .RecordCount = 0 Then
               CmbSalePacking.ListIndex = 0
            Else
               CmbSalePacking.Text = !PackingName
            End If
         End With
         '===================================================
         With cn.Execute("Select UnitName from Units Where UnitID = " & IIf(IsNull(Rs1!UnitID), "0", Rs1!UnitID))
            If .RecordCount = 0 Then
               CmbUnits.ListIndex = 0
            Else
               CmbUnits.Text = !UnitName
            End If
         End With
         '====================================================
         TextBox1.Text = IIf(IsNull(Rs1!ProductName1), "", Rs1!ProductName1)
         TxtPurPrice.Text = IIf(ObjUserSecurity.IsAdministrator = True, Rs1!PurPrice, 0)
         TxtRetailPrice.Text = Rs1!RetailPrice
         TxtWSPrice.Text = Rs1!WSPrice
         TxtPurDisc.Text = IIf(IsNull(Rs1!PurDiscPC), 0, Rs1!PurDiscPC)
         TxtSaleDisc.Text = IIf(IsNull(Rs1!DiscPC), 0, Rs1!DiscPC)
         TxtSaleDiscPer.Text = IIf(IsNull(Rs1!DiscPer), 0, Rs1!DiscPer)
         TxtMinStockLimit.Text = IIf(IsNull(Rs1!MinStockLimit), 0, Rs1!MinStockLimit)
         TxtMaxStockLimit.Text = IIf(IsNull(Rs1!MaxStockLimit), 0, Rs1!MaxStockLimit)
         TxtTokenVal.Text = IIf(IsNull(Rs1!TokenVal), 0, Rs1!TokenVal)
         TxtSaleTaxPer.Text = IIf(IsNull(Rs1!SaleTaxPer), 0, Rs1!SaleTaxPer)
         TxtServiceCharges.Text = IIf(IsNull(Rs1!ServiceCharges), 0, Rs1!ServiceCharges)
         TxtEmpComm.Text = IIf(IsNull(Rs1!EmpComm), 0, Rs1!EmpComm)
         TxtDesc1.Text = IIf(IsNull(Rs1!Desc1), "", Rs1!Desc1)
         ChkLockProduct.Value = Abs(Rs1!IsLocked)
         ChkDeadProduct.Value = Abs(Rs1!IsDeadProduct)
         ChkClosingProduct.Value = Abs(Rs1!IsClosingProduct)
         ChkNoCostProduct.Value = Abs(Rs1!IsNoCostProduct)
         ChkIsChangedPrice.Value = Abs(Rs1!isChangedPrice)
         ChkRawProduct.Value = Abs(Rs1!IsRawProduct)
         OptWSPSaleTax.Value = Rs1!IsWSSaleTax
         OptRPSaleTax.Value = Rs1!IsRetailSaleTax
         ChkWSDiscb4ST.Value = Abs(Rs1!IsWSDiscb4ST)
         
         PopulatePackGrid
         PopulateCodeGrid
         PopulateColour
         PopulateSize
      End If
   End If
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
   On Error GoTo ErrorHandler
   If IsNumeric(TxtCode.Text) = True Then
      If (TxtCode.Text) = "" Or Len(TxtCode.Text) < 6 Then
         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   If CGrid.Columns("Code").Text = "" And cn.Execute("Select count(*) from ProductBarcodes where Code='" & (TxtCode.Text) & "'").Fields(0).Value = 1 Then
      MsgBox "Code Already Exists", vbExclamation, "Alert"
      TxtCode.SetFocus
      Exit Sub
   End If
   If (TxtCode.Text) <> CGrid.Columns("Code").Text Then
      If CGrid.Columns("Code").Text <> "" And cn.Execute("Select count(*) from ProductBarcodes where Code='" & (TxtCode.Text) & "'").Fields(0).Value = 1 Then
         MsgBox "Code Already Exists", vbExclamation, "Alert"
         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   RsCode.Filter = "Code='" & IIf(Trim(CGrid.Columns("Code").Text) <> "", CGrid.Columns("Code").Text, (TxtCode.Text)) & "'"
   If RsCode.RecordCount = 0 Then RsCode.AddNew
   CGrid.Columns("Code").Text = (TxtCode.Text)
   CGrid.Columns("Qty").Text = IIf(Val(TxtQty.Text) = 0, "", Val(TxtQty.Text))
   If vIsNewRecord = False Then cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated BarCode-" & CGrid.Columns("Code").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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
   CGrid.CancelUpdate
   CGrid.RemoveAll
   CGrid.AddNew
   CGrid.Columns("Code").Text = " "
   CGrid.Update
   PopulateColour
   PopulateSize
   Rs.Filter = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
   On Error GoTo ErrorHandler
   If isFreeCode = True Then
      Dim vSQL As String
      vSQL = "SELECT right('00000' + cast(isnull(min(cast(ProductID+1 as int)),0) as varchar),5)" & vbCrLf _
         + " FROM Products WHERE" & vbCrLf _
         + "     --The next # is missing" & vbCrLf _
         + " ProductID+1 NOT IN (SELECT ProductID FROM Products)" & vbCrLf _
         + " AND" & vbCrLf _
         + "   --We haven't reached the max of values" & vbCrLf _
         + " ProductID+1 < (SELECT Max(ProductID) FROM Products)"
      FunGetMaxID = cn.Execute(vSQL).Fields(0)
   Else
      FunGetMaxID = cn.Execute("Select right('00000' + cast(isnull(max(cast(ProductId as int)),0) + 1 as varchar),5) from Products --Where ProductId like '" & TxtGroupID.Text & "%'").Fields(0)
   End If
   'FunGetMaxID = CN.Execute("Select right('0000' + cast(isnull(max(cast(substring(ProductId,3,10) as smallint)),0) + 1 as varchar),4) from Products").Fields(0) ' Where ProductId like '" & GetGroupID(CmbCompany) & "%'").Fields(0)
   If ObjRegistry.DuplicateCode = True Then
      TxtCode.Text = TxtGroupID.Text & cn.Execute("Select right('0000' + cast(isnull(max(cast(substring(Code,4,10) as int)),0) + 1 as varchar),4) from ProductBarcodes Where len(code)=7 and Code like '" & TxtGroupID.Text & "%'").Fields(0)
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
      If vIsNewRecord = False Then cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Inserted New PackingID-v" & RsProductPacking!PackingID & " Multiplier- " & RsProductPacking!Multiplier & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   ElseIf RsProductPacking.RecordCount = 1 And Val(PGrid.Columns("Multiplier").Value) = 0 Then
      If vIsNewRecord = False Then cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Deleted PackingID-v" & RsProductPacking!PackingID & " Multiplier- " & RsProductPacking!Multiplier & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
      RsProductPacking.Delete
   ElseIf RsProductPacking.RecordCount = 1 Then
      RsProductPacking!Multiplier = Val(PGrid.Columns("Multiplier").Value)
      If vIsNewRecord = False Then cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated PackingID-v" & RsProductPacking!PackingID & " Multiplier- " & RsProductPacking!Multiplier & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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

Private Sub TxtFilterID_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtFilterID.Name Then Exit Sub
   If Trim(TxtFilterID.Text) = "" Then Grid.MoveFirst: Exit Sub
   If Len(TxtFilterID.Text) > 5 Then
      With cn.Execute("select * from Productbarcodes where Code = '" & TxtFilterID.Text & "'")
         If .RecordCount > 0 Then
            Rs1.Find "ProductID ='" & !Productid & "'", , adSearchForward, 1
         End If
         .Close
      End With
   Else
      Rs1.Find "ProductID ='" & Right("00000" + CStr(Val(TxtFilterID.Text)), 5) & "'", , adSearchForward, 1
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

Private Sub TxtID_LostFocus()
   On Error GoTo ErrorHandler
   If Len(TxtID.Text) = 5 Then Exit Sub
   TxtID.Text = Right("00000" + CStr(Val(TxtID.Text)), 5)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtItemCode_GotFocus()
'   MsgBox "Got"
End Sub

Private Sub TxtItemCode_LostFocus()
'   On Error GoTo ErrorHandler
'   If TxtItemCode.Text = "" Then Exit Sub
'   'If vIsNewRecord = True Then
'   If CN.Execute("Select count(*) from ProductBarcodes where Code='" & (TxtItemCode.Text) & "'").Fields(0).Value = 0 Then
'      TxtCode.Text = TxtItemCode.Text
'      GetDataFromTexBoxesToCGrid
'   End If
'   If TxtVenderID.Visible Then TxtVenderID.SetFocus
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
End Sub

Private Sub TxtName_LostFocus()
   On Error GoTo ErrorHandler
   If ObjRegistry.ProperCase = False Then Exit Sub
   TxtName.Text = StrConv(TxtName.Text, vbProperCase)
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
   Rs1.Open "Select * FROM Products where 1=1 " & vProductName & " Order By ProductName", cn, adOpenStatic, adLockOptimistic
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
   Rs1.Open "Select * FROM Products where 1=1 " & vProductName & " Order By ProductName", cn, adOpenStatic, adLockOptimistic
   Set Grid.DataSource = Rs1
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_Products ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False Then
        If TxtCompanyID.Text <> IIf(IsNull(Rs!companyid), "", Rs!companyid) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated CompanyID-" & Rs!companyid & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtGroupID.Text <> IIf(IsNull(Rs!GroupID), "", Rs!GroupID) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated GroupID-" & Rs!GroupID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtSubGroupID.Text <> IIf(IsNull(Rs!SubGroupID), "", Rs!SubGroupID) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated SubGroupID NO -" & Rs!SubGroupID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtBrandID.Text <> IIf(IsNull(Rs!BrandID), "", Rs!BrandID) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated BrandID NO -" & Rs!BrandID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtName.Text <> Rs!ProductName Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated Product Name-" & Rs!ProductName & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If CmbUnits.ItemData(CmbUnits.ListIndex) <> IIf(IsNull(Rs!UnitID), "", Rs!UnitID) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated UnitID-" & Rs!UnitID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If CmbPurPacking.ItemData(CmbPurPacking.ListIndex) <> IIf(IsNull(Rs!PurchasePackingID), "", Rs!PurchasePackingID) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated PurchasePackingID-" & Rs!PurchasePackingID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If CmbSalePacking.ItemData(CmbSalePacking.ListIndex) <> IIf(IsNull(Rs!SalePackingID), "", Rs!SalePackingID) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated SalePackingID-" & Rs!SalePackingID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPurPrice.Text <> IIf(IsNull(Rs!PurPrice), "", Rs!PurPrice) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated PurPrice NO-" & Rs!PurPrice & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtRetailPrice.Text <> IIf(IsNull(Rs!RetailPrice), "", Rs!RetailPrice) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated RetailPrice-" & Rs!RetailPrice & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtWSPrice.Text <> IIf(IsNull(Rs!WSPrice), "", Rs!WSPrice) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated WSPrice-" & Rs!WSPrice & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPurDisc.Text <> IIf(IsNull(Rs!PurDiscPC), "", Rs!PurDiscPC) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated PurDiscPC-" & Rs!PurDiscPC & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtSaleDisc.Text <> IIf(IsNull(Rs!DiscPC), "", Rs!DiscPC) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated DiscPC-" & Rs!DiscPC & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtMinStockLimit.Text <> IIf(IsNull(Rs!MinStockLimit), "", Rs!MinStockLimit) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated MinStockLimit-" & Rs!MinStockLimit & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtMaxStockLimit.Text <> IIf(IsNull(Rs!MaxStockLimit), "", Rs!MaxStockLimit) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated MaxStockLimit-" & Rs!MaxStockLimit & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If ChkLockProduct.Value <> Abs(Rs!IsLocked) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated IsLocked-" & Rs!IsLocked & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If ChkDeadProduct.Value <> Abs(Rs!IsDeadProduct) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated IsDead-" & Rs!IsDeadProduct & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If ChkClosingProduct.Value <> Abs(Rs!IsClosingProduct) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated IsClosingProduct-" & Rs!IsClosingProduct & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If ChkNoCostProduct.Value <> Abs(Rs!IsNoCostProduct) Then
            cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated IsNoCostProduct-" & Rs!IsNoCostProduct & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
   Else
        cn.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TextBox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyDown-Event behaviours for  ''
        '' Enter, Space, Delete & Tab keys to set      ''
        '' Behavior in Textbox1.Text, keys will behave ''
        '' as Normal Text writing behavior.            ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''

   On Error GoTo ErrorHandler
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

        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyPress-Event behaviours for ''
        '' Alfabatic, Numaric & Symbolic keys to write ''
        '' Urdu. I've tried to make it near with Urdu  ''
        '' Phonetic Keyboard Layout.                   ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''
   On Error GoTo ErrorHandler
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
    With cn.Execute(vStrSQL)
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

Private Sub TxtSubDepartmentID_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtSubDepartmentID.Name Then Exit Sub
   If TxtSubDepartmentName.Text <> "" Then TxtSubDepartmentName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSubDepartmentID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSubDepartmentID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSubDepartmentName.Text <> "" Then Exit Sub
   If TxtSubDepartmentID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSubDepartment(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSubDepartment(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSubDepartment(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSubDepartments.Show vbModal, Me
        If SchSubDepartments.ParaOutSubDepartmentID = "" Then FunSelectSubDepartment = False: Exit Function
        TxtSubDepartmentID.Text = SchSubDepartments.ParaOutSubDepartmentID
    End If
    '---------------------------
    vStrSQL = " Select * FROM SubDepartments where SubDepartmentID=" & Val(TxtSubDepartmentID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSubDepartmentName.Text = !SubDepartmentName
          FunSelectSubDepartment = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSubDepartment = False
          .Close
          TxtSubDepartmentID.Text = ""
          TxtSubDepartmentName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnAddSubDepartment_Click()
   On Error GoTo ErrorHandler
   DefSubDepartment.Show vbModal, Me
   If TxtSubDepartmentID.Visible Then TxtSubDepartmentID.SetFocus
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
    With cn.Execute(vStrSQL)
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

Private Sub TxtDescriptionID_Change()
   On Error GoTo ErrorHandler
   If TxtDescriptionID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtDescriptionID.Name Then Exit Sub
   If TxtDescriptionName.Text <> "" Then TxtDescriptionName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDescriptionID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtDescriptionID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtDescriptionID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectDescription(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectDescription(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectDescription(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchDescription.Show vbModal, Me
        If SchDescription.ParaOutDescriptionID = "" Then FunSelectDescription = False: Exit Function
        TxtDescriptionID.Text = SchDescription.ParaOutDescriptionID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Descriptions where DescriptionID=" & Val(TxtDescriptionID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtDescriptionName.Text = !DescriptionName
          FunSelectDescription = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectDescription = False
          .Close
          TxtDescriptionID.Text = ""
          TxtDescriptionName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnAddDescription_Click()
   On Error GoTo ErrorHandler
   DefDescription.Show vbModal, Me
   If TxtDescriptionID.Visible Then TxtDescriptionID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtItemDescID_Change()
   On Error GoTo ErrorHandler
   If TxtItemDescID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtItemDescID.Name Then Exit Sub
   If TxtItemDescName.Text <> "" Then TxtItemDescName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtItemDescID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtItemDescID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtItemDescID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectItemDesc(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectItemDesc(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectItemDesc(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchItemDesc.Show vbModal, Me
        If SchItemDesc.ParaOutItemDescID = "" Then FunSelectItemDesc = False: Exit Function
        TxtItemDescID.Text = SchItemDesc.ParaOutItemDescID
    End If
    '---------------------------
    vStrSQL = " Select * FROM ItemDescription where ItemDescID=" & Val(TxtItemDescID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtItemDescName.Text = !ItemDescName
          FunSelectItemDesc = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectItemDesc = False
          .Close
          TxtItemDescID.Text = ""
          TxtItemDescName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectItemCode(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchItemCode.Show vbModal, Me
        If SchItemCode.ParaOutItemCode = "" Then FunSelectItemCode = False: Exit Function
        TxtItemCode.Text = SchItemCode.ParaOutItemCode
    End If
    '---------------------------
   FunSelectItemCode = False
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnAddItemDesc_Click()
   On Error GoTo ErrorHandler
   DefItemDescription.Show vbModal, Me
   If TxtItemDescID.Visible Then TxtItemDescID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnVender_Click()
   On Error GoTo ErrorHandler
   If FunSelectVender(ssButton, False) = True Then
      If TxtDepartmentID.Visible Then TxtDepartmentID.SetFocus Else TxtVenderID.SetFocus
   Else
      TxtVenderID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectVender(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInDetail = ""
        SchAccounts.CmbFilter = "Vendors"
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectVender = False: Exit Function
        TxtVenderID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select c.* FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " where BarCode = '" & (TxtVenderID.Text) & "' or (c.AccountNo = '" & (TxtVenderID.Text) & "' and (c.AccountNo like '6%') and c.isDetailed = 1 and c.isLocked = 0)"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtVenderID.Text = !AccountNo
          TxtVenderName.Text = !AccountName
          FunSelectVender = True
          .Close
          Exit Function
      Else
          FunSelectVender = False
          .Close
          TxtVenderID.Text = ""
          TxtVenderName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtVenderID_Change()
   On Error GoTo ErrorHandler
   If TxtVenderID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   If TxtVenderName.Text <> "" Then TxtVenderName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtVenderID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtVenderID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVender(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVender(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
