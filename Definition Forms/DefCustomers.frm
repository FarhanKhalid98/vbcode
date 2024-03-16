VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form DefCustomers 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "DefCustomers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtStoreID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7875
      MaxLength       =   10
      TabIndex        =   83
      Top             =   1665
      Width           =   675
   End
   Begin VB.TextBox TxtTransportName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10710
      MaxLength       =   100
      TabIndex        =   14
      Top             =   7320
      Width           =   3240
   End
   Begin VB.CheckBox ChkHideLockCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Hide Lock Customers"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1620
      TabIndex        =   81
      Tag             =   "NC"
      Top             =   2160
      Width           =   1860
   End
   Begin VB.TextBox TxtRefComm 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10740
      MaxLength       =   5
      TabIndex        =   19
      Top             =   8685
      Width           =   720
   End
   Begin VB.TextBox TxtRemarks 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7875
      MaxLength       =   100
      TabIndex        =   13
      Top             =   7320
      Width           =   2835
   End
   Begin VB.TextBox TxtDescription 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7875
      MaxLength       =   50
      TabIndex        =   4
      Top             =   4800
      Width           =   2595
   End
   Begin VB.OptionButton OptRetail 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Retail"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7920
      TabIndex        =   63
      ToolTipText     =   "Standard users are restricted to perform only those tasks which are explicitly assigned to them by the System administrator."
      Top             =   9600
      Width           =   780
   End
   Begin VB.OptionButton OptWholeSale 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Whole Sale"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8730
      TabIndex        =   62
      ToolTipText     =   $"DefCustomers.frx":0ECA
      Top             =   9600
      Value           =   -1  'True
      Width           =   1125
   End
   Begin VB.CheckBox ChkDateofJoining 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10080
      TabIndex        =   61
      Top             =   9150
      Width           =   195
   End
   Begin VB.TextBox TxtCNIC 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10470
      MaxLength       =   20
      TabIndex        =   9
      Top             =   6045
      Width           =   2595
   End
   Begin VB.TextBox TxtPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   7875
      MaxLength       =   2
      TabIndex        =   55
      Tag             =   "NC"
      Top             =   2355
      Width           =   525
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7875
      MaxLength       =   100
      TabIndex        =   1
      Top             =   2985
      Width           =   5265
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7875
      MaxLength       =   100
      TabIndex        =   2
      Top             =   3525
      Width           =   5265
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7875
      MaxLength       =   30
      TabIndex        =   3
      Top             =   4245
      Width           =   5265
   End
   Begin VB.TextBox TxtPhone1 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7875
      MaxLength       =   30
      TabIndex        =   6
      Top             =   5385
      Width           =   2595
   End
   Begin VB.TextBox TxtMobileNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10470
      MaxLength       =   25
      TabIndex        =   7
      Top             =   5385
      Width           =   2595
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7875
      MaxLength       =   50
      TabIndex        =   10
      Top             =   6705
      Width           =   2385
   End
   Begin VB.TextBox TxtMobileNo2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7875
      MaxLength       =   30
      TabIndex        =   8
      Top             =   6045
      Width           =   2595
   End
   Begin VB.TextBox TxtContactPerson 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10260
      MaxLength       =   30
      TabIndex        =   11
      Top             =   6705
      Width           =   2025
   End
   Begin VB.CheckBox ChkLockCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Lock Customer"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7920
      TabIndex        =   21
      Top             =   9225
      Width           =   1500
   End
   Begin VB.TextBox TxtCreditLimit 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   12285
      MaxLength       =   6
      TabIndex        =   12
      Top             =   6705
      Width           =   855
   End
   Begin VB.TextBox TxtBarCode 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   11475
      MaxLength       =   13
      TabIndex        =   15
      Top             =   8685
      Width           =   2040
   End
   Begin VB.TextBox TxtSectorName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   10515
      TabIndex        =   37
      Tag             =   "NC"
      Top             =   2355
      Width           =   1440
   End
   Begin VB.TextBox TxtSectorID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9450
      MaxLength       =   10
      TabIndex        =   0
      Top             =   2355
      Width           =   705
   End
   Begin VB.TextBox TxtZoneName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   11955
      TabIndex        =   36
      Tag             =   "NC"
      Top             =   2355
      Width           =   1440
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
      Height          =   2850
      Left            =   13440
      TabIndex        =   32
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
         Height          =   2445
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   33
         Tag             =   "NC"
         Text            =   "DefCustomers.frx":0F53
         Top             =   360
         Width           =   3930
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
         TabIndex        =   34
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1620
      MaxLength       =   30
      TabIndex        =   25
      Tag             =   "NC"
      Top             =   2430
      Width           =   5205
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5700
      Left            =   240
      TabIndex        =   26
      Top             =   2775
      Width           =   6915
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
      stylesets(0).Picture=   "DefCustomers.frx":0FDE
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
      Columns(0).Width=   1852
      Columns(0).Caption=   "Customer ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6456
      Columns(1).Caption=   "Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2699
      Columns(2).Caption=   "ContactNo"
      Columns(2).Name =   "ContactNo"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   12197
      _ExtentY        =   10054
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
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   2805
      TabIndex        =   27
      Top             =   9945
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
      MICON           =   "DefCustomers.frx":0FFA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   4125
      TabIndex        =   28
      Top             =   9945
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
      MICON           =   "DefCustomers.frx":1016
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   5445
      TabIndex        =   29
      Top             =   9945
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
      MICON           =   "DefCustomers.frx":1032
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8250
      TabIndex        =   22
      Top             =   9945
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
      MICON           =   "DefCustomers.frx":104E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   9570
      TabIndex        =   23
      Top             =   9945
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
      MICON           =   "DefCustomers.frx":106A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   10890
      TabIndex        =   24
      Top             =   9945
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
      MICON           =   "DefCustomers.frx":1086
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSector 
      Height          =   330
      Left            =   10155
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   2355
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
      MICON           =   "DefCustomers.frx":10A2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSaleAccount 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11130
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   7995
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
      MICON           =   "DefCustomers.frx":10BE
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSaleAccountNo 
      Height          =   315
      Left            =   10365
      TabIndex        =   17
      Top             =   7995
      Width           =   780
      _ExtentX        =   1376
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
   Begin SITextBox.Txt TxtSaleAccountName 
      Height          =   315
      Left            =   11490
      TabIndex        =   43
      Tag             =   "NC"
      Top             =   7995
      Width           =   2010
      _ExtentX        =   3545
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
   Begin JeweledBut.JeweledButton BtnAllocateProductPrice 
      CausesValidation=   0   'False
      Height          =   675
      Left            =   12285
      TabIndex        =   59
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   9060
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1191
      TX              =   "Allocate Product Price"
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
      MICON           =   "DefCustomers.frx":10DA
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpDateofJoining 
      Height          =   315
      Left            =   10080
      TabIndex        =   20
      Tag             =   "NC"
      Top             =   9435
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
   Begin JeweledBut.JeweledButton BtnCustomerType 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11490
      TabIndex        =   65
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   1680
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
      MICON           =   "DefCustomers.frx":10F6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerTypeID 
      Height          =   315
      Left            =   10710
      TabIndex        =   66
      Top             =   1665
      Width           =   780
      _ExtentX        =   1376
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
   Begin SITextBox.Txt TxtCustomerType 
      Height          =   315
      Left            =   11850
      TabIndex        =   67
      Tag             =   "NC"
      Top             =   1680
      Width           =   1560
      _ExtentX        =   2752
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
   Begin SITextBox.Txt TxtEmpID 
      Height          =   315
      Left            =   7890
      TabIndex        =   16
      Top             =   7995
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
   Begin SITextBox.Txt TxtEmpName 
      Height          =   315
      Left            =   8835
      TabIndex        =   71
      Top             =   7995
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
      Left            =   8475
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   7980
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
      MICON           =   "DefCustomers.frx":1112
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtLicenceNO 
      Height          =   315
      Left            =   10470
      TabIndex        =   5
      Top             =   4800
      Width           =   2595
      _ExtentX        =   4577
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
   Begin SITextBox.Txt TxtRefID 
      Height          =   315
      Left            =   7890
      TabIndex        =   18
      Top             =   8685
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
   Begin SITextBox.Txt TxtRefName 
      Height          =   315
      Left            =   8850
      TabIndex        =   76
      Top             =   8685
      Width           =   1890
      _ExtentX        =   3334
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
   Begin JeweledBut.JeweledButton BtnRef 
      Height          =   330
      Left            =   8490
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   8670
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
      MICON           =   "DefCustomers.frx":112E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   8910
      TabIndex        =   84
      Top             =   1665
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Height          =   315
      Left            =   8550
      TabIndex        =   85
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   1665
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "DefCustomers.frx":114A
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   8415
      TabIndex        =   88
      Top             =   2355
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   556
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
      IntegralPoint   =   7
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   8910
      TabIndex        =   87
      Top             =   1470
      Width           =   840
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   7875
      TabIndex        =   86
      Top             =   1470
      Width           =   585
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transport Name"
      Height          =   195
      Left            =   10710
      TabIndex        =   82
      Top             =   7110
      Width           =   1140
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Comm %"
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
      Left            =   10710
      TabIndex        =   80
      Top             =   8460
      Width           =   630
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Reference ID"
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
      Left            =   7890
      TabIndex        =   79
      Top             =   8460
      Width           =   945
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Reference Name"
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
      Left            =   8925
      TabIndex        =   78
      Top             =   8460
      Width           =   1215
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Licence #"
      Height          =   195
      Left            =   10470
      TabIndex        =   75
      Top             =   4575
      Width           =   720
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
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
      Left            =   8910
      TabIndex        =   74
      Top             =   7770
      Width           =   1140
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
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
      Left            =   7875
      TabIndex        =   73
      Top             =   7770
      Width           =   870
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   7875
      TabIndex        =   70
      Top             =   7095
      Width           =   630
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Type ID"
      Height          =   195
      Left            =   10710
      TabIndex        =   69
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Type"
      Height          =   195
      Index           =   3
      Left            =   12150
      TabIndex        =   68
      Top             =   1470
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   2
      Left            =   7875
      TabIndex        =   64
      Top             =   4575
      Width           =   795
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Joining"
      Height          =   195
      Left            =   10380
      TabIndex        =   60
      Top             =   9150
      Width           =   1065
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No."
      Height          =   195
      Left            =   10470
      TabIndex        =   58
      Top             =   5160
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNIC No."
      Height          =   195
      Left            =   10470
      TabIndex        =   57
      Top             =   5835
      Width           =   675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      Height          =   195
      Left            =   7875
      TabIndex        =   56
      Top             =   2145
      Width           =   870
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   7875
      TabIndex        =   54
      Top             =   2775
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   7875
      TabIndex        =   53
      Top             =   3315
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Index           =   0
      Left            =   7875
      TabIndex        =   52
      Top             =   4020
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No.s"
      Height          =   195
      Left            =   7875
      TabIndex        =   51
      Top             =   5175
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      Height          =   195
      Left            =   7875
      TabIndex        =   50
      Top             =   6495
      Width           =   435
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      Height          =   195
      Left            =   10260
      TabIndex        =   49
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No. 2"
      Height          =   195
      Left            =   7875
      TabIndex        =   48
      Top             =   5850
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit"
      Height          =   195
      Left            =   12285
      TabIndex        =   47
      Top             =   6495
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Codes"
      Height          =   195
      Left            =   11475
      TabIndex        =   46
      Top             =   8460
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Name"
      Height          =   195
      Index           =   1
      Left            =   11520
      TabIndex        =   45
      Top             =   7755
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale A/C No"
      Height          =   195
      Left            =   10395
      TabIndex        =   44
      Top             =   7755
      Width           =   900
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector Name"
      Height          =   195
      Left            =   10515
      TabIndex        =   41
      Top             =   2145
      Width           =   930
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector ID"
      Height          =   195
      Left            =   9450
      TabIndex        =   40
      Top             =   2145
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone Name"
      Height          =   195
      Left            =   11955
      TabIndex        =   39
      Top             =   2145
      Width           =   840
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
      Left            =   12825
      TabIndex        =   35
      Top             =   990
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customers"
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
      TabIndex        =   31
      Top             =   270
      Width           =   1500
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   497
      X2              =   496
      Y1              =   162
      Y2              =   568
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   810
      TabIndex        =   30
      Top             =   2460
      Width           =   510
   End
End
Attribute VB_Name = "DefCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vMode As FormMode
Dim vMaxBinID As Integer
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.
Dim vSQL As String

Private Sub BtnAllocateProductPrice_Click()
   FrmAllocateProductPrice.ParaInCustomerID = (TxtPrefix.Text & TxtID.Text)
   FrmAllocateProductPrice.Show
End Sub
Private Sub TxtStoreID_Change()
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If Trim(TxtStoreID.Text) = "" Then Exit Sub
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
Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtCustomerTypeID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
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
    vStrSQL = " Select * FROM Stores where islock = 0 and StoreID=" & Val(TxtStoreID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtStoreName.Text = !StoreName
          FunSelectStore = True
          .Close
          Exit Function
      Else
          FunSelectStore = False
          .Close
          TxtStoreID.Text = ""
          TxtStoreName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub ChkHideLockCustomer_Click()
   If ActiveControl.Name <> ChkHideLockCustomer.Name Then Exit Sub
   Call TxtFilter_Change
End Sub

Private Sub ChkLockCustomer_Click()
   If ActiveControl.Name <> ChkLockCustomer.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
     If ActiveControl.Name = Grid.Name Then Call Grid_DblClick: Exit Sub
     keybd_event 9, 1, 1, 1
     KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, False) = True Then TxtName.SetFocus Else TxtSectorID.SetFocus
         Case TxtCustomerTypeID.Name: If FunSelectCustomerType(ssFunctionKey, False) = True Then If TxtSectorID.Enabled Then TxtSectorID.SetFocus
         Case TxtEmpID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then TxtSaleAccountNo.SetFocus Else TxtEmpID.SetFocus
         Case TxtSaleAccountNo.Name: If FunSelectSaleAccount(ssFunctionKey, True) = True Then ChkLockCustomer.SetFocus
      End Select
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
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
         Case vbKeyN
            If BtnNew.Enabled Then BtnNew_Click
            KeyCode = 0
         Case vbKeyO
            If BtnOpen.Enabled Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   End If
End Sub

Private Sub Grid_DblClick()
   If Grid.Rows > 0 And BtnOpen.Enabled Then BtnOpen_Click
End Sub

Private Sub Grid_GotFocus()
   On Error GoTo ErrorHandler
   Call Grid_RowColChange(0, 0)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKeyA To vbKeyZ, vbKey0 To vbKey9, Asc("a") To Asc("z")
      TxtFilter.Text = Chr(KeyAscii): TxtFilter.SelStart = Len(TxtFilter.Text):  TxtFilter.SetFocus
   End Select
End Sub

Private Sub BtnClear_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FormStatus = SelectionMode
End Sub

Private Sub BtnClose_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & Val(TxtID.Text) & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniCustomers", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
  Dim vtbl As String
  If Rs.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
    Dim vid As String
    vid = Rs!PartyID
    vtbl = Common.ChildDataExists("Parties", "PartyId='" & vid & "'", "") & Common.ChildDataExists("ChartoFAccounts", "AccountNo='" & vid & "'", "Parties")
    If vtbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
      Exit Sub
    End If
    CN.BeginTrans
    
'    vMaxBinID = FunGetMaxBinID
    ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header----------------------------------------------
'    CN.Execute ("Insert Into Bin_Parties Select " & vMaxBinID & ",'" & Date & "',* from Parties Where PartyID = " & TxtPrefix.Text & TxtID.Text)
    
    Call ActivityLog("Customers", eDelete, , , TxtPrefix.Text & TxtID.Text)
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtPrefix.Text & TxtID.Text & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Rs.Delete
    CN.Execute ("Delete From ChartOfAccounts Where AccountNo = '" & vid & "'")
    CN.CommitTrans
    If Rs.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
    Rs.MoveNext
    Grid.MoveNext
    If Rs.EOF Then Rs.MoveLast
  End If
  Exit Sub
ErrorHandler:
  If CN.Errors.Count > 0 Then CN.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Sub BtnNew_Click()
  FormStatus = NewMode
End Sub

Private Sub BtnOpen_Click()
  On Error GoTo ErrorHandler
  If Rs.RecordCount > 0 Then
    If Rs.BOF = False And Rs.EOF = False Then
      FormStatus = OpenMode
    End If
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If FunValidation = False Then Exit Sub
    
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniCustomers", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False Then Call ActivityLog("Customers", eEdit, , , TxtPrefix.Text & TxtID.Text)
   Set Rs = New ADODB.Recordset
   Rs.Open " Select * FROM Parties where PartyID = '" & TxtPrefix.Text & TxtID.Text & "'", CN, adOpenDynamic, adLockOptimistic
   Call UserActivities
   CN.BeginTrans
   If vIsNewRecord Then
   vStrSQL = "Insert into chartofaccounts values ('" & _
      Val(TxtPrefix.Text & TxtID.Text) & "',1,'" & Replace(TxtName.Text, "'", "''") & "','Customers',2,'" & Replace(TxtAddress.Text, "'", "''") & "','62',0,0,1," & ChkLockCustomer.Value & ",1,0,' ',0,0,'" & Date & "',0)"
    CN.Execute vStrSQL
      Rs.AddNew
      Rs!PartyID = TxtPrefix.Text & TxtID.Text
      Rs!isChanged = 0
   Else
      Rs!isChanged = 1
      Rs!IsSync = 0
      Rs!modified_on = Now
      CN.Execute ("Update Chartofaccounts set Accountname = '" & Replace(TxtName.Text, "'", "''") & "',Narration = '" & Replace(TxtAddress.Text, "'", "''") & "', IsSync = 0, isLocked = " & ChkLockCustomer.Value & ", isChanged = 1 Where AccountNo = " & Rs!PartyID)
   End If
   Rs!CustomerTypeID = IIf(Trim(TxtCustomerTypeID.Text) = "", Null, TxtCustomerTypeID.Text)
   Rs!SectorID = IIf(Trim(TxtSectorID.Text) = "", Null, TxtSectorID.Text)
   Rs!StoreID = IIf(Trim(TxtStoreID.Text) = "", Null, TxtStoreID.Text)
   Rs!partyname = TxtName.Text
   Rs!Address = TxtAddress.Text
   Rs!City = TxtCity.Text
   Rs!LicenceNO = IIf(TxtLicenceNO.Text = "", Null, TxtLicenceNO.Text)
   Rs!Phone1 = IIf(TxtPhone1.Text = "", Null, TxtPhone1.Text)
   Rs!Phone2 = Null
   Rs!Mobile = IIf(TxtMobileNo.Text = "", Null, TxtMobileNo.Text)
   Rs!Mobile2 = IIf(TxtMobileNo2.Text = "", Null, TxtMobileNo2.Text)
   Rs!CNIC = IIf(TxtCNIC.Text = "", Null, TxtCNIC.Text)
   Rs!Email = TxtEmail.Text
   Rs!CreditLimit = Val(TxtCreditLimit.Text)
   Rs!ContactPerson = TxtContactPerson.Text
   Rs!SaleAccountNo = IIf(TxtSaleAccountNo.Text = "", Null, TxtSaleAccountNo.Text)
   Rs!EmpID = IIf(Trim(TxtEmpID.Text) = "", Null, TxtEmpID.Text)
   Rs!PartyType = "C"
   Rs!Description = IIf(TxtDescription.Text = "", Null, TxtDescription.Text)
   Rs!Remarks = IIf(TxtRemarks.Text = "", Null, TxtRemarks.Text)
   Rs!TransportName = IIf(TxtTransportName.Text = "", Null, TxtTransportName.Text)
   Rs!isWholeSale = OptWholeSale.Value
   Rs!IsLockParty = ChkLockCustomer.Value
   Rs!BarCode = IIf(Trim(TxtBarCode.Text) = "", Null, Trim(TxtBarCode.Text))
   Rs!DateofJoining = IIf(ChkDateofJoining.Value = 1, IIf(DtpDateofJoining.DateValue <> "", DtpDateofJoining.DateValue, Null), Null)
   Rs!RefID = IIf(TxtRefID.Text = "", Null, TxtRefID.Text)
   Rs!RefComm = IIf(TxtRefComm.Text = "", Null, TxtRefComm.Text)
   Rs.Update
   CN.CommitTrans
   ParaCustID = TxtPrefix.Text & TxtID.Text
   ParaCustName = TxtName.Text
   Set Rs = New ADODB.Recordset
   vSQL = " Select P.*, isnull(Phone1,'') + isnull(Phone2,'') + isnull(mobile,'') + isnull(mobile2,'') ContactNo FROM Parties P inner join ChartOfAccounts C on C.AccountNo = P.PartyID where PartyType = 'C' and ( cast(PartyID as varchar(10)) + PartyName + isnull(Phone1,'') + isnull(Phone2,'') + isnull(mobile,'') + isnull(mobile2,'')) like '%" & Replace(TxtFilter.Text, "'", "''") & "%'" & IIf(ChkHideLockCustomer.Value = 1, " and islocked = 0 and isLockParty = 0 ", "") & " Order by PartyName"
   Rs.Open vSQL, CN, adOpenDynamic, adLockOptimistic
   If vIsNewRecord = True Then Call ActivityLog("Customers", eAdd, , , TxtPrefix.Text & TxtID.Text)
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   If vIsNewRecord Then
      If Trim(TxtID.Text) = "" Then
         MsgBox "Please specify a Customer ID", vbExclamation, "Alert"
         If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
         Exit Function
      End If
      If Not IsNumeric(TxtID.Text) Then
         MsgBox "The Customer ID must be numeric", vbExclamation, "Alert"
         If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
         Exit Function
      End If
'      vSql = "select partyID, PartyName from Parties where (phone1 = '" & TxtPhone1.Text & "' Or mobile = '" & TxtMobileNo.Text & "' or mobile2 = '" & TxtMobileNo2.Text & "') and partyid <> '" & TxtPrefix.Text & TxtID.Text & "'"
'      With cn.Execute(vSql)
'      If .EOF = False Then
'         MsgBox "This Contact No. already exists against of ID " & .Fields("partyID").Value & " Name " & .Fields("PartyName").Value & " Please Enter New Contact No. then Save", vbExclamation, "Alert"
'         Exit Function
'      End If
'      End With
   End If
   If Trim(TxtName.Text) = "" Then
      MsgBox "Please specify a Customer Name", vbExclamation, "Alert"
      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
      Exit Function
   End If
   If ObjRegistry.SectorCompulsory Then
      If Val(TxtSectorID.Text) = 0 Then
         MsgBox "Please specify a Sectore Name", vbExclamation, "Alert"
         If TxtSectorID.Enabled And TxtSectorID.Visible Then TxtSectorID.SetFocus
         Exit Function
      End If
   End If
   
   If TxtID.Enabled = True And CN.Execute("select count(*) from chartofaccounts where accountno = '" & TxtPrefix.Text & TxtID.Text & "'").Fields(0) > 0 Then
      MsgBox "This ID already exists. A new ID has been generated. Please save again", vbExclamation, "Alert"
      TxtID.Text = FunGetMaxID
      TxtID.SetFocus
      Exit Function
   End If
   'All Ok, now validation is success
   FunValidation = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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
   'Me.ControlBox = True
   Me.ScaleMode = 3
   Me.ScaleHeight = 768
   Me.ScaleWidth = 1024
   Me.Height = 11940
   Me.Width = 15450
   
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Customers"
   HelpLocation Me
   Set Rs = New ADODB.Recordset
   vSQL = " Select P.*, isnull(Phone1,'') + isnull(Phone2,'') + isnull(mobile,'') + isnull(mobile2,'') ContactNo FROM Parties P inner join ChartOfAccounts C on C.AccountNo = P.PartyID where PartyType = 'C' and (cast(PartyID as varchar(10)) + PartyName + isnull(Phone1,'') + isnull(Phone2,'') + isnull(mobile,'') + isnull(mobile2,'')) like '%" & Replace(TxtFilter.Text, "'", "''") & "%'" & IIf(ChkHideLockCustomer.Value = 1, " and islocked = 0 and isLockParty = 0 ", "") & " Order by PartyName"
   Rs.Open vSQL, CN, adOpenDynamic, adLockOptimistic
   Grid.Columns("ID").DataField = "PartyId"
   Grid.Columns("Name").DataField = "PartyName"
   Grid.Columns("ContactNo").DataField = "ContactNo"
   FormStatus = NewMode
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
      Call SubClearFields(True)
      BtnAllocateProductPrice.Enabled = False
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtPrefix.Enabled = False
      TxtPrefix.Text = "62"
      TxtID.Text = FunGetMaxID
      TxtFilter.Text = ""
      Grid.Enabled = False
      Set Grid.DataSource = Rs
      TxtFilter.Enabled = False
      ChkHideLockCustomer.Enabled = False
      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      Call SubClearFields(True)
      Call Grid_RowColChange(0, 0)
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnClear.Enabled = True
      BtnAllocateProductPrice.Enabled = True
      Grid.Enabled = False
      TxtID.Enabled = False
      TxtFilter.Enabled = False
      ChkHideLockCustomer.Enabled = False
      TxtName.SetFocus
      TxtFilter.Text = ""
      vIsNewRecord = False
    Case Is = ChangeMode
      BtnSave.Enabled = True
    Case Is = SelectionMode
      Grid.Enabled = True
      Call SubClearFields(False)
      Call Grid_RowColChange(0, 0)
      TxtFilter.Enabled = True
      ChkHideLockCustomer.Enabled = True
      BtnNew.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnSave.Enabled = False
      BtnClear.Enabled = False
      Grid.SetFocus
      'TxtFilter.Text = ""
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   If Rs.RecordCount > 0 And Grid.Enabled Then
'      TxtID.Text = Mid(Grid.Columns("ID").Text, 3)
'      TxtName.Text = Grid.Columns("Name").Text
      TxtID.Text = Mid(CStr(Rs!PartyID), 3)
      TxtName.Text = Rs!partyname
      TxtAddress.Text = IIf(IsNull(Rs!Address), "", Rs!Address)
      TxtCity.Text = IIf(IsNull(Rs!City), "", Rs!City)
      TxtPhone1.Text = IIf(IsNull(Rs!Phone1), "", Rs!Phone1)
      'TxtPhone2.Text = IIf(IsNull(Rs!Phone2), "", Rs!Phone2)
      TxtMobileNo.Text = IIf(IsNull(Rs!Mobile), "", Rs!Mobile)
      TxtMobileNo2.Text = IIf(IsNull(Rs!Mobile2), "", Rs!Mobile2)
      TxtCNIC.Text = IIf(IsNull(Rs!CNIC), "", Rs!CNIC)
      TxtLicenceNO.Text = IIf(IsNull(Rs!LicenceNO), "", Rs!LicenceNO)
      TxtEmail.Text = IIf(IsNull(Rs!Email), "", Rs!Email)
      TxtContactPerson.Text = IIf(IsNull(Rs!ContactPerson), "", Rs!ContactPerson)
      TxtCreditLimit.Text = IIf(IsNull(Rs!CreditLimit), "", Rs!CreditLimit)
      ChkLockCustomer.Value = Abs(Rs!IsLockParty)
      TxtEmpID.Text = IIf(IsNull(Rs!EmpID), "", Rs!EmpID)
      If Trim(TxtEmpID.Text) <> "" Then
         TxtEmpName.Text = CN.Execute("Select EmpName from Employees Where EmpID = '" & TxtEmpID.Text & "'").Fields(0)
      Else
         TxtEmpName.Text = ""
      End If
      TxtSaleAccountNo.Text = IIf(IsNull(Rs!SaleAccountNo), "", Rs!SaleAccountNo)
      If Trim(TxtSaleAccountNo.Text) <> "" Then
         TxtSaleAccountName.Text = CN.Execute("Select AccountName from ChartofAccounts Where AccountNo = '" & TxtSaleAccountNo.Text & "'").Fields(0)
      Else
         TxtSaleAccountName.Text = ""
      End If
      TxtCustomerTypeID.Text = IIf(IsNull(Rs!CustomerTypeID), "", Rs!CustomerTypeID)
      If Trim(TxtCustomerTypeID.Text) <> "" Then
         TxtCustomerType.Text = CN.Execute("Select CustomerType from CustomerTypes Where CustomerTypeID = '" & TxtCustomerTypeID.Text & "'").Fields(0)
      Else
         TxtCustomerType.Text = ""
      End If
      If IsNull(Rs!DateofJoining) Then
         DtpDateofJoining.DateValue = ""
         ChkDateofJoining.Value = 0
      Else
         DtpDateofJoining.DateValue = Rs!DateofJoining
         ChkDateofJoining.Value = 1
      End If
      TxtRefID.Text = IIf(IsNull(Rs!RefID), "", Rs!RefID)
      If Trim(TxtRefID.Text) <> "" Then
         TxtRefName.Text = CN.Execute("Select partyName from Parties Where RefID = '" & TxtRefID.Text & "'").Fields(0)
      Else
         TxtRefName.Text = ""
      End If
      TxtRefComm.Text = IIf(IsNull(Rs!RefComm), "", Rs!RefComm)
      TxtDescription.Text = IIf(IsNull(Rs!Description), "", Rs!Description)
      TxtRemarks.Text = IIf(IsNull(Rs!Remarks), "", Rs!Remarks)
      TxtTransportName.Text = IIf(IsNull(Rs!TransportName), "", Rs!TransportName)
      OptWholeSale.Value = IIf(IsNull(Rs!isWholeSale), 1, Rs!isWholeSale)
      OptRetail.Value = Not OptWholeSale.Value
      TxtBarCode.Text = IIf(IsNull(Rs!BarCode), "", Rs!BarCode)
      TxtStoreID.Text = IIf(IsNull(Rs!StoreID), "", Rs!StoreID)
      If Trim(TxtStoreID.Text) <> "" Then
         TxtStoreName.Text = CN.Execute("Select StoreName from Stores Where StoreID = '" & TxtStoreID.Text & "'").Fields(0)
      Else
         TxtStoreName.Text = ""
      End If
      TxtSectorID.Text = IIf(IsNull(Rs!SectorID), "", Rs!SectorID)
      If TxtSectorID.Text = "" Then
         TxtSectorName.Text = ""
         TxtZoneName.Text = ""
      Else
         With CN.Execute("select * from sectors s inner join Zones t on s.ZoneID = t.ZoneID where sectorid =" & Val(TxtSectorID.Text))
         If .RecordCount > 0 Then
            TxtSectorName.Text = !SectorName
            TxtZoneName.Text = !ZoneName
         End If
         End With
      End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearFields(Enable As Boolean)
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then
            ctl.Text = ""
            ctl.Enabled = Enable
         End If
      ElseIf TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then
            ctl.Text = ""
            ctl.Enabled = Enable
         End If
      ElseIf TypeOf ctl Is JeweledButton Then
         If ctl.Tag = "B" Then
            ctl.Enabled = Enable
         End If
      ElseIf TypeOf ctl Is CheckBox Then
         ctl.Enabled = Enable
      ElseIf TypeOf ctl Is ComboBox Then
         ctl.Enabled = Enable
      End If
   Next
   DtpDateofJoining.Enabled = Enable
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkDateofJoining_Click()
   DtpDateofJoining.Enabled = IIf(ChkDateofJoining.Value = 1, True, False)
End Sub
      
Private Function FunGetMaxID() As String
   FunGetMaxID = CN.Execute("Select isnull(max(cast(substring(cast(accountno as varchar(10)),3,10) as int)),0) + 1 from chartofaccounts Where AccountNo like '62%' and isdetailed=1").Fields(0)
End Function

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtAddress_LostFocus()
   TxtAddress.Text = StrConv(TxtAddress.Text, vbProperCase)
End Sub

Private Sub TxtCity_LostFocus()
   TxtCity.Text = StrConv(TxtCity.Text, vbProperCase)
End Sub

Private Sub TxtContactPerson_LostFocus()
   TxtContactPerson.Text = StrConv(TxtContactPerson.Text, vbProperCase)
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, vbKeyA To vbKeyZ, Asc("@"), Asc("_"), Asc("-"), Asc(" "), vbKeyBack, Asc("a") To Asc("z"), Asc(".")
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtFax_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtFilter_Change()
   On Error GoTo ErrorHandler
   'If Me.ActiveControl.Name <> TxtFilter.Name Then Exit Sub
   'If Trim(TxtFilter.Text) = "" Then Grid.MoveFirst: Exit Sub
   Set Rs = New ADODB.Recordset
'   Rs.Open " Select *, isnull(Phone1,'') + isnull(Phone2,'') + isnull(mobile,'') + isnull(mobile2,'') ContactNo FROM Parties where PartyType = 'C' and (PartyID + PartyName + isnull(Phone1,'') + isnull(Phone2,'') + isnull(mobile,'') + isnull(mobile2,'')) like '%" & Replace(TxtFilter.Text, "'", "''") & "%' Order by PartyName", cn, adOpenDynamic, adLockOptimistic
   vSQL = " Select P.*, isnull(Phone1,'') + isnull(Phone2,'') + isnull(mobile,'') + isnull(mobile2,'') ContactNo FROM Parties P inner join ChartOfAccounts C on C.AccountNo = P.PartyID where PartyType = 'C' and (cast(PartyID as varchar(10))+ PartyName + isnull(Phone1,'') + isnull(Phone2,'') + isnull(mobile,'') + isnull(mobile2,'')) like '%" & Replace(TxtFilter.Text, "'", "''") & "%'" & IIf(ChkHideLockCustomer.Value = 1, " and islocked = 0 and isLockParty = 0 ", "") & " Order by PartyName"
   Rs.Open vSQL, CN, adOpenDynamic, adLockOptimistic
   Set Grid.DataSource = Rs
   'Rs.Find "PartyName like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtMobileNo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtName_Change()
   If Me.ActiveControl.Name <> TxtName.Name Then Exit Sub
   TxtFilter.Text = TxtName.Text
End Sub

Private Sub TxtName_LostFocus()
   TxtName.Text = StrConv(TxtName.Text, vbProperCase)
End Sub

Private Sub TxtPhone1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtPhone2_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxBinID = CN.Execute("Select isnull(max(BinID),0)+1 from Bin_Parties ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
    If vIsNewRecord = False Then
        If TxtName.Text <> IIf(IsNull(Rs!partyname), "", Rs!partyname) Then
            CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Customer-" & Replace(Rs!partyname, "'", "''") & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtAddress.Text <> IIf(IsNull(Rs!Address), "", Rs!Address) Then
            CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Address-" & Rs!Address & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtCity.Text <> IIf(IsNull(Rs!City), "", Rs!City) Then
            CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated City-" & Rs!City & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPhone1.Text <> IIf(IsNull(Rs!Phone1), "", Rs!Phone1) Then
            CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Phone1-" & Rs!Phone1 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
'        If TxtPhone2.Text <> IIf(IsNull(Rs!Phone2), "", Rs!Phone2) Then
'            CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Phone2-" & Rs!Phone2 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
'        End If
        If TxtMobileNo.Text <> IIf(IsNull(Rs!Mobile), "", Rs!Mobile) Then
            CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Mobile-" & Rs!Mobile & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtMobileNo2.Text <> IIf(IsNull(Rs!Mobile), "", Rs!Mobile2) Then
            CN.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated Mobile2-" & Rs!Mobile2 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtEmail.Text <> IIf(IsNull(Rs!Email), "", Rs!Email) Then
            CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Email-" & Rs!Email & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtContactPerson.Text <> IIf(IsNull(Rs!ContactPerson), "", Rs!ContactPerson) Then
            CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated ContactPerson-" & Rs!ContactPerson & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If ChkLockCustomer.Value <> Val(Rs!IsLockParty) Then
            CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated ChkLockCustomer-" & Rs!IsLockParty & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
   Else
        CN.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub BtnSector_Click()
   If FunSelectSector(ssButton, False) = True Then
      If TxtName.Enabled Then TxtName.SetFocus
   Else
      TxtSectorID.SetFocus
   End If
End Sub

Private Sub TxtSectorID_Change()
   If TxtSectorID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   If TxtSectorName.Text <> "" Then TxtSectorName.Text = ""
   If TxtZoneName.Text <> "" Then TxtZoneName.Text = ""
End Sub

Private Sub TxtSectorID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSectorID.Text = "" Then Exit Sub
   If TxtSectorName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSector(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSector(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSector(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSector.Show vbModal, Me
        If SchSector.ParaOutSectorID = "" Then FunSelectSector = False: Exit Function
        TxtSectorID.Text = SchSector.ParaOutSectorID
    End If
    '---------------------------
    vStrSQL = "Select * FROM Sectors s inner join Zones t on t.ZoneID = s.ZoneID where SectorID=" & Val(TxtSectorID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSectorName.Text = !SectorName
          TxtZoneName.Text = !ZoneName
          FunSelectSector = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSector = False
          .Close
          TxtSectorID.Text = ""
          TxtSectorName.Text = ""
          TxtZoneName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectSaleAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectSaleAccount = False: Exit Function
        TxtSaleAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select * FROM ChartOfAccounts where AccountNo='" & TxtSaleAccountNo.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSaleAccountName.Text = !AccountName
          FunSelectSaleAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSaleAccount = False
          .Close
          TxtSaleAccountNo.Text = ""
          TxtSaleAccountName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtSaleAccountNo_Change()
   If TxtSaleAccountNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSaleAccountNo.Name Then Exit Sub
   If TxtSaleAccountName.Text <> "" Then TxtSaleAccountName.Text = ""
End Sub

Private Sub TxtSaleAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSaleAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSaleAccountNo.Text = "" Then Exit Sub
   If TxtSaleAccountName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSaleAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSaleAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSaleAccount_Click()
   If FunSelectSaleAccount(ssButton, False) = True Then
      If ChkLockCustomer.Enabled Then ChkLockCustomer.SetFocus
   Else
      If TxtSaleAccountNo.Enabled Then TxtSaleAccountNo.SetFocus
   End If
End Sub

Private Function FunSelectCustomerType(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCustomerType.Show vbModal, Me
        If SchCustomerType.ParaOutID = "" Then FunSelectCustomerType = False: Exit Function
        TxtCustomerTypeID.Text = SchCustomerType.ParaOutID
    End If
    '---------------------------
    vStrSQL = " Select * FROM CustomerTypes where CustomerTypeID = '" & TxtCustomerTypeID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerType.Text = !CustomerType
          FunSelectCustomerType = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectCustomerType = False
          .Close
          TxtCustomerTypeID.Text = ""
          TxtCustomerType.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtCustomerTypeID_Change()
   If TxtCustomerTypeID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomerTypeID.Name Then Exit Sub
   If TxtCustomerTypeID.Text = "" Then TxtCustomerType.Text = ""
End Sub

Private Sub TxtCustomerTypeID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtCustomerTypeID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerTypeID.Text = "" Then Exit Sub
   If TxtCustomerType.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCustomerType(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCustomerType(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnCustomerType_Click()
   If FunSelectCustomerType(ssButton, False) = True Then
      If TxtSectorID.Enabled Then TxtSectorID.SetFocus
   Else
      If TxtCustomerTypeID.Enabled Then TxtCustomerTypeID.SetFocus
   End If
End Sub

Private Sub TxtEmpID_Change()
   If TxtEmpID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtEmpID.Name Then Exit Sub
   If TxtEmpName.Text <> "" Then TxtEmpName.Text = ""
End Sub

Private Sub TxtEmpID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtEmpID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtEmpName.Text <> "" Then Exit Sub
   If Trim(TxtEmpID.Text) = "" Then Exit Sub
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
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchEmployee.Show vbModal, Me
        If SchEmployee.ParaOutEmployeeID = "" Then FunSelectEmployee = False: Exit Function
        TxtEmpID.Text = SchEmployee.ParaOutEmployeeID
    End If
    '---------------------------
    vStrSQL = " Select EmpName FROM Employees where EmpID = " & Val(TxtEmpID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtEmpName.Text = !empname
          FunSelectEmployee = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectEmployee = False
          .Close
          TxtEmpID.Text = ""
          TxtEmpName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnEmployee_Click()
   If FunSelectEmployee(ssButton, False) = True Then
      TxtSaleAccountNo.SetFocus
   Else
      TxtEmpID.SetFocus
   End If
End Sub

Private Sub TxtRefID_Change()
   If TxtRefID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtRefID.Name Then Exit Sub
   If TxtRefName.Text <> "" Then TxtRefName.Text = ""
End Sub

Private Sub TxtRefID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtRefID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtRefName.Text <> "" Then Exit Sub
   If Trim(TxtRefID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectReference(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectReference(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectReference(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.CmbFilter = "Customers"
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectReference = False: Exit Function
        TxtRefID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = "Select c.AccountNo, c.AccountName as AccountName, Address, City, p.phone1, p.phone2, p.mobile, p.mobile2, p.Description, isnull(p.isWholeSale,1) as isWholeSale, LicenceNO, TransportName, Remarks" & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " where (c.AccountNo = '" & (TxtRefID.Text) & "' or P.Phone1 = '" & (TxtRefID.Text) & "' or P.Phone2 = '" & (TxtRefID.Text) & "' or P.Mobile = '" & (TxtRefID.Text) & "' or P.Mobile2 = '" & (TxtRefID.Text) & "') and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and isDetailed = 1 and isLocked = 0"
    
    vStrSQL = vStrSQL + " union all Select EmpID, EmpName as AccountName, Address, City, '' phone1,'' phone2, '' mobile, '' mobile2, '', 1 as isWholeSale, '' as LicenceNO, '' as TransportName, '' as Remarks" & vbCrLf _
         + " from Employees" & vbCrLf _
         + " where EmpID = '" & (TxtRefID.Text) & "' and isLockEmployee = 0"
    
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtRefID.Text = !AccountNo
          TxtRefName.Text = !AccountName
          FunSelectReference = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectReference = False
          .Close
          TxtRefID.Text = ""
          TxtRefName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnRef_Click()
   If FunSelectReference(ssButton, False) = True Then
      TxtRefID.SetFocus
   Else
      TxtRefComm.SetFocus
   End If
End Sub

