VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form DefVendors 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15450
   Icon            =   "DefVendors.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   770
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNTN 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   12075
      MaxLength       =   100
      TabIndex        =   2
      Top             =   2925
      Width           =   1440
   End
   Begin VB.CheckBox ChkHideLockVendor 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Hide Lock Vendor"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2820
      TabIndex        =   65
      Tag             =   "NC"
      Top             =   2115
      Width           =   1860
   End
   Begin VB.TextBox TxtDescription 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10590
      MaxLength       =   50
      TabIndex        =   5
      Top             =   4545
      Width           =   2595
   End
   Begin VB.TextBox TxtZoneName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   10665
      TabIndex        =   60
      Tag             =   "NC"
      Top             =   1650
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox TxtZoneID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9600
      MaxLength       =   10
      TabIndex        =   59
      Tag             =   "C"
      Top             =   1650
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox TxtPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   7995
      MaxLength       =   2
      TabIndex        =   57
      Tag             =   "NC"
      Top             =   2280
      Width           =   525
   End
   Begin VB.TextBox TxtMobileNo2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7995
      MaxLength       =   30
      TabIndex        =   8
      Top             =   5835
      Width           =   2595
   End
   Begin VB.TextBox TxtCNIC 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10590
      MaxLength       =   20
      TabIndex        =   9
      Top             =   5835
      Width           =   2595
   End
   Begin VB.CheckBox ChkDateofJoining 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11685
      TabIndex        =   14
      Top             =   7605
      Width           =   195
   End
   Begin VB.TextBox TxtMobileNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10590
      MaxLength       =   25
      TabIndex        =   7
      Top             =   5175
      Width           =   2595
   End
   Begin VB.TextBox TxtPhone1 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7995
      MaxLength       =   30
      TabIndex        =   6
      Top             =   5175
      Width           =   2595
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7995
      MaxLength       =   100
      TabIndex        =   1
      Top             =   2925
      Width           =   4095
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   7995
      MaxLength       =   100
      TabIndex        =   3
      Top             =   3555
      Width           =   5265
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7995
      MaxLength       =   30
      TabIndex        =   4
      Top             =   4545
      Width           =   2595
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7995
      MaxLength       =   50
      TabIndex        =   10
      Top             =   6510
      Width           =   5265
   End
   Begin VB.TextBox TxtContactPerson 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7995
      MaxLength       =   30
      TabIndex        =   11
      Top             =   7170
      Width           =   2655
   End
   Begin VB.TextBox TxtBankAccountNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10650
      MaxLength       =   100
      TabIndex        =   12
      Top             =   7170
      Width           =   2595
   End
   Begin VB.CheckBox ChkLockVendor 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Lock Vendor"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7995
      TabIndex        =   16
      Top             =   8325
      Width           =   1320
   End
   Begin VB.TextBox TxtZoneNameSector 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   12075
      TabIndex        =   32
      Tag             =   "NC"
      Top             =   2280
      Width           =   1440
   End
   Begin VB.TextBox TxtSectorID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9570
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "C"
      Top             =   2280
      Width           =   705
   End
   Begin VB.TextBox TxtSectorName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   10635
      TabIndex        =   31
      Tag             =   "NC"
      Top             =   2280
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
      Left            =   13080
      TabIndex        =   27
      Top             =   960
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
         TabIndex        =   28
         Tag             =   "NC"
         Text            =   "DefVendors.frx":0ECA
         Top             =   360
         Width           =   3930
      End
      Begin VB.Label Label15 
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
         TabIndex        =   29
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2820
      MaxLength       =   30
      TabIndex        =   20
      Tag             =   "NC"
      Top             =   2415
      Width           =   4395
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5700
      Left            =   1935
      TabIndex        =   21
      Top             =   2745
      Width           =   5340
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
      stylesets(0).Picture=   "DefVendors.frx":0F55
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
      Columns(0).Width=   1852
      Columns(0).Caption=   "Vendor ID"
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
      TabNavigation   =   1
      _ExtentX        =   9419
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
      Left            =   2880
      TabIndex        =   22
      Top             =   9060
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
      MICON           =   "DefVendors.frx":0F71
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   4200
      TabIndex        =   23
      Top             =   9060
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
      MICON           =   "DefVendors.frx":0F8D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   5520
      TabIndex        =   24
      Top             =   9060
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
      MICON           =   "DefVendors.frx":0FA9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8325
      TabIndex        =   17
      Top             =   9060
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
      MICON           =   "DefVendors.frx":0FC5
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   9645
      TabIndex        =   18
      Top             =   9060
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
      MICON           =   "DefVendors.frx":0FE1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   10965
      TabIndex        =   19
      Top             =   9060
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
      MICON           =   "DefVendors.frx":0FFD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSector 
      Height          =   330
      Left            =   10275
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   2280
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
      MICON           =   "DefVendors.frx":1019
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSaleAccount 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8775
      TabIndex        =   37
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   7815
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
      MICON           =   "DefVendors.frx":1035
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSaleAccountNo 
      Height          =   315
      Left            =   7995
      TabIndex        =   13
      Top             =   7815
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
      Left            =   9135
      TabIndex        =   38
      Tag             =   "NC"
      Top             =   7815
      Width           =   2460
      _ExtentX        =   4339
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpDateofJoining 
      Height          =   315
      Left            =   11685
      TabIndex        =   15
      Tag             =   "NC"
      Top             =   7815
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   9600
      TabIndex        =   53
      Tag             =   "NC"
      Top             =   1020
      Width           =   495
      _ExtentX        =   873
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   10095
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   1020
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
      MICON           =   "DefVendors.frx":1051
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   10455
      TabIndex        =   55
      Tag             =   "NC"
      Top             =   1020
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
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   8520
      TabIndex        =   58
      Top             =   2280
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
   Begin JeweledBut.JeweledButton BtnZone 
      Height          =   330
      Left            =   10305
      TabIndex        =   61
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   1650
      Visible         =   0   'False
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
      MICON           =   "DefVendors.frx":106D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnApplyFixDiscounts 
      CausesValidation=   0   'False
      Height          =   675
      Left            =   11745
      TabIndex        =   66
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   8190
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1191
      TX              =   "Apply Fix Discounts"
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
      MICON           =   "DefVendors.frx":1089
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NTN"
      Height          =   195
      Left            =   12075
      TabIndex        =   67
      Top             =   2715
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   2
      Left            =   10590
      TabIndex        =   64
      Top             =   4320
      Width           =   795
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone Name"
      Height          =   195
      Left            =   10665
      TabIndex        =   63
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone ID"
      Height          =   195
      Left            =   9600
      TabIndex        =   62
      Top             =   1440
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   9600
      TabIndex        =   56
      Top             =   795
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No. 2"
      Height          =   195
      Left            =   7995
      TabIndex        =   52
      Top             =   5640
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNIC No."
      Height          =   195
      Left            =   10590
      TabIndex        =   51
      Top             =   5625
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Joining"
      Height          =   195
      Left            =   11910
      TabIndex        =   50
      Top             =   7605
      Width           =   1065
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
      Height          =   195
      Left            =   10590
      TabIndex        =   49
      Top             =   4950
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Nos"
      Height          =   195
      Left            =   7995
      TabIndex        =   48
      Top             =   4965
      Width           =   795
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   7995
      TabIndex        =   47
      Top             =   2715
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor ID"
      Height          =   195
      Left            =   7995
      TabIndex        =   46
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   7995
      TabIndex        =   45
      Top             =   3345
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Index           =   0
      Left            =   7995
      TabIndex        =   44
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      Height          =   195
      Left            =   7995
      TabIndex        =   43
      Top             =   6300
      Width           =   435
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      Height          =   195
      Left            =   7995
      TabIndex        =   42
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account No"
      Height          =   195
      Left            =   10650
      TabIndex        =   41
      Top             =   6960
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale A/C No"
      Height          =   195
      Left            =   7995
      TabIndex        =   40
      Top             =   7605
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Name"
      Height          =   195
      Index           =   1
      Left            =   9120
      TabIndex        =   39
      Top             =   7605
      Width           =   750
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone Name"
      Height          =   195
      Left            =   12075
      TabIndex        =   36
      Top             =   2070
      Width           =   840
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector ID"
      Height          =   195
      Left            =   9570
      TabIndex        =   35
      Top             =   2070
      Width           =   675
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector Name"
      Height          =   195
      Left            =   10635
      TabIndex        =   34
      Top             =   2070
      Width           =   930
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
      Left            =   11745
      TabIndex        =   30
      Top             =   540
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendors"
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
      Left            =   675
      TabIndex        =   26
      Top             =   315
      Width           =   1200
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   508
      X2              =   507
      Y1              =   160
      Y2              =   566
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   2205
      TabIndex        =   25
      Top             =   2505
      Width           =   510
   End
End
Attribute VB_Name = "DefVendors"
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

Private Sub BtnApplyFixDiscounts_Click()
'   FrmApplyFixDiscounts.ParaInVendorID = (TxtPrefix.Text & TxtID.Text)
'   FrmApplyFixDiscounts.Show
   FrmPurchaseDiscounts.Show
End Sub

Private Sub ChkHideLockVendor_Click()
   If ActiveControl.Name <> ChkHideLockVendor.Name Then Exit Sub
   Call TxtFilter_Change
End Sub

Private Sub ChkLockVendor_Click()
   If ActiveControl.Name <> ChkLockVendor.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkDateofJoining_Click()
   DtpDateofJoining.Enabled = IIf(ChkDateofJoining.Value = 1, True, False)
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
         Case TxtZoneID.Name: If FunSelectZone(ssFunctionKey, False) = True Then TxtSectorID.SetFocus Else TxtZoneID.SetFocus
         Case TxtSaleAccountNo.Name: If FunSelectSaleAccount(ssFunctionKey, True) = True Then ChkLockVendor.SetFocus
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
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
          Set frmObj = Nothing
      Next
      Set Rs = Nothing
      Set DefVendors = Nothing
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
    cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FormStatus = SelectionMode
End Sub

Private Sub BtnClose_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & Val(TxtID.Text) & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniVendors", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
    vtbl = Common.ChildDataExists("Vendors", "PartyId = " & vid, "") & Common.ChildDataExists("ChartoFAccounts", "AccountNo = " & vid, "Parties")
    If vtbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
      Exit Sub
    End If
    cn.BeginTrans
    
'    vMaxBinID = FunGetMaxBinID
    ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header----------------------------------------------
'    CN.Execute ("Insert Into Bin_Parties Select " & vMaxBinID & ",'" & Date & "',* from Parties Where PartyID = " & TxtPrefix.Text & TxtID.Text)

    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtPrefix.Text & TxtID.Text & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Call ActivityLog("Vendors", eDelete, , , TxtPrefix.Text & TxtID.Text)
    Rs.Delete
    cn.Execute ("Delete From ChartOfAccounts Where AccountNo = " & vid)
    cn.CommitTrans
    If Rs.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
    Rs.MoveNext
    Grid.MoveNext
    If Rs.EOF Then Rs.MoveLast
  End If
  Exit Sub
ErrorHandler:
  If cn.Errors.Count > 0 Then cn.RollbackTrans
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
   If FunValidation = False Then Exit Sub
    
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniVendors", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False Then Call ActivityLog("Vendors", eEdit, , , TxtPrefix.Text & TxtID.Text)
   Set Rs = New ADODB.Recordset
   Rs.Open "Select * FROM Parties where PartyID = " & Val(TxtPrefix.Text & TxtID.Text), cn, adOpenDynamic, adLockOptimistic
   cn.BeginTrans
'   Call UserActivities
   If vIsNewRecord Then
      cn.Execute ("Insert into chartofaccounts values (" & _
      TxtPrefix.Text & TxtID.Text & ",1,'" & Replace(TxtName.Text, "'", "''") & "','Vendors',2,'" & Replace(TxtAddress.Text, "'", "''") & "','61',0,0,1," & ChkLockVendor.Value & ",1,0,' ',0,0,'" & Date & "',0)")
      Rs.AddNew
      Rs!PartyID = TxtPrefix.Text & TxtID.Text
      Rs!isChanged = 0
   Else
      Rs!isChanged = 1
      Rs!IsSync = 0
      Rs!modified_on = Now
      cn.Execute ("Update Chartofaccounts set Accountname = '" & Replace(TxtName.Text, "'", "''") & "',Narration = '" & Replace(TxtAddress.Text, "'", "''") & "', IsSync = 0, isLocked = " & ChkLockVendor.Value & ", isChanged = 1 Where AccountNo = " & Rs!PartyID)
   End If
   Rs!SectorID = IIf(Trim(TxtSectorID.Text) = "", Null, TxtSectorID.Text)
   Rs!ZoneID = IIf(Trim(TxtZoneID.Text) = "", Null, TxtZoneID.Text)
   Rs!StoreID = IIf(Trim(TxtStoreID.Text) = "", Null, TxtStoreID.Text)
   Rs!partyname = TxtName.Text
   Rs!Address = IIf(TxtAddress.Text = "", Null, TxtAddress.Text)
   Rs!City = IIf(TxtCity.Text = "", Null, TxtCity.Text)
   Rs!Phone1 = IIf(TxtPhone1.Text = "", Null, TxtPhone1.Text)
   Rs!Phone2 = Null
   Rs!Mobile = IIf(TxtMobileNo.Text = "", Null, TxtMobileNo.Text)
   Rs!Mobile2 = IIf(TxtMobileNo2.Text = "", Null, TxtMobileNo2.Text)
   Rs!CNIC = IIf(TxtCNIC.Text = "", Null, TxtCNIC.Text)
   Rs!NTN = IIf(TxtNTN.Text = "", Null, TxtNTN.Text)
   Rs!Email = IIf(TxtEmail.Text = "", Null, TxtEmail.Text)
   Rs!ContactPerson = IIf(TxtContactPerson.Text = "", Null, TxtContactPerson.Text)
   Rs!AccountNo = IIf(TxtBankAccountNo.Text = "", Null, TxtBankAccountNo.Text)
   Rs!SaleAccountNo = IIf(TxtSaleAccountNo.Text = "", Null, TxtSaleAccountNo.Text)
   Rs!PartyType = "V"
   Rs!Description = IIf(TxtDescription.Text = "", Null, TxtDescription.Text)
   Rs!IsLockParty = ChkLockVendor.Value
   Rs!DateofJoining = IIf(ChkDateofJoining.Value = 1, IIf(DtpDateofJoining.DateValue <> "", DtpDateofJoining.DateValue, Null), Null)
   Rs.Update
   If vIsNewRecord = True Then Call ActivityLog("Vendors", eAdd, , , TxtPrefix.Text & TxtID.Text)
   cn.CommitTrans
   Set Rs = New ADODB.Recordset
   vSQL = " Select P.* FROM Parties P inner join ChartOfAccounts C on C.AccountNo = P.PartyID where PartyType='V' and PartyName like '%" & Replace(TxtFilter.Text, "'", "''") & "%'" & IIf(ChkHideLockVendor.Value = 1, " and islocked = 0 and isLockParty = 0 ", "") & " Order by PartyName"
   Rs.Open vSQL, cn, adOpenDynamic, adLockOptimistic
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
  If vIsNewRecord Then
    If Trim(TxtID.Text) = "" Then
      MsgBox "Please specify a Vendor ID", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
    If Not IsNumeric(TxtID.Text) Or Val(TxtID.Text) < 1 Then
      MsgBox "The Vendor ID must be numeric", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
  End If
  If Trim(TxtName.Text) = "" Then
    MsgBox "Please specify a Vendor Name", vbExclamation, "Alert"
    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
    Exit Function
  End If
  If TxtID.Enabled = True And cn.Execute("select count(*) from chartofaccounts where accountno = '" & TxtPrefix.Text & TxtID.Text & "'").Fields(0) > 0 Then
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
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Vendors"
   HelpLocation Me
   
'   TxtStoreID.Text = ObjRegistry.StoreID
'   FunSelectStore ssValidate, True
'   TxtStoreID.Visible = ObjRegistry.StoreVisible
'   BtnStore.Visible = ObjRegistry.StoreVisible
'   TxtStoreName.Visible = ObjRegistry.StoreVisible
'   LblStoreID.Visible = ObjRegistry.StoreVisible
   
   Set Rs = New ADODB.Recordset
   vSQL = " Select P.* FROM Parties P inner join ChartOfAccounts C on C.AccountNo = P.PartyID where PartyType='V' and PartyName like '%" & Replace(TxtFilter.Text, "'", "''") & "%'" & IIf(ChkHideLockVendor.Value = 1, " and islocked = 0 and isLockParty = 0 ", "") & " Order by PartyName"
   Rs.Open vSQL, cn, adOpenDynamic, adLockOptimistic
   Grid.Columns("ID").DataField = "PartyId"
   Grid.Columns("Name").DataField = "Partyname"
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
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtPrefix.Enabled = False
      TxtPrefix.Text = "61"
      TxtID.Text = FunGetMaxID
      TxtFilter.Text = ""
      Set Grid.DataSource = Rs
      Grid.Enabled = False
      TxtFilter.Enabled = False
      ChkHideLockVendor.Enabled = False
      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
      Set Grid.DataSource = Rs
      vIsNewRecord = True
    Case Is = OpenMode
      Call SubClearFields(True)
      Call Grid_RowColChange(0, 0)
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnClear.Enabled = True
      Grid.Enabled = False
      TxtID.Enabled = False
      TxtFilter.Enabled = False
      ChkHideLockVendor.Enabled = False
      TxtName.SetFocus
      vIsNewRecord = False
    Case Is = ChangeMode
      BtnSave.Enabled = True
    Case Is = SelectionMode
      Grid.Enabled = True
      Call SubClearFields(False)
      Call Grid_RowColChange(0, 0)
      TxtFilter.Enabled = True
      ChkHideLockVendor.Enabled = True
      BtnNew.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnSave.Enabled = False
      BtnClear.Enabled = False
      Grid.SetFocus
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   If Rs.RecordCount > 0 And Grid.Enabled Then
      TxtID.Text = Mid(Grid.Columns("ID").Text, 3)
      TxtName.Text = Grid.Columns("Name").Text
      TxtAddress.Text = IIf(IsNull(Rs!Address), "", Rs!Address)
      TxtCity.Text = IIf(IsNull(Rs!City), "", Rs!City)
      TxtPhone1.Text = IIf(IsNull(Rs!Phone1), "", Rs!Phone1)
'      TxtPhone2.Text = IIf(IsNull(Rs!Phone2), "", Rs!Phone2)
      TxtMobileNo.Text = IIf(IsNull(Rs!Mobile), "", Rs!Mobile)
      TxtMobileNo2.Text = IIf(IsNull(Rs!Mobile2), "", Rs!Mobile2)
      TxtCNIC.Text = IIf(IsNull(Rs!CNIC), "", Rs!CNIC)
      TxtNTN.Text = IIf(IsNull(Rs!NTN), "", Rs!NTN)
      TxtEmail.Text = IIf(IsNull(Rs!Email), "", Rs!Email)
      TxtContactPerson.Text = IIf(IsNull(Rs!ContactPerson), "", Rs!ContactPerson)
      TxtBankAccountNo.Text = IIf(IsNull(Rs!AccountNo), "", Rs!AccountNo)
      TxtSaleAccountNo.Text = IIf(IsNull(Rs!SaleAccountNo), "", Rs!SaleAccountNo)
      TxtDescription.Text = IIf(IsNull(Rs!Description), "", Rs!Description)
      If Trim(TxtSaleAccountNo.Text) <> "" Then
         TxtSaleAccountName.Text = cn.Execute("Select AccountName from ChartofAccounts Where AccountNo = '" & TxtSaleAccountNo.Text & "'").Fields(0)
      Else
         TxtSaleAccountName.Text = ""
      End If
      If IsNull(Rs!DateofJoining) Then
         DtpDateofJoining.DateValue = ""
         ChkDateofJoining.Value = 0
      Else
         DtpDateofJoining.DateValue = Rs!DateofJoining
         ChkDateofJoining.Value = 1
      End If
      ChkLockVendor.Value = Abs(Rs!IsLockParty)
      TxtZoneID.Text = IIf(IsNull(Rs!ZoneID), "", Rs!ZoneID)
      If TxtZoneID.Text <> "" Then
         TxtZoneName.Text = cn.Execute("Select ZoneName from Zones Where ZoneID = '" & TxtZoneID.Text & "'").Fields(0)
      Else
         TxtZoneName.Text = ""
      End If
       TxtStoreID.Text = IIf(IsNull(Rs!StoreID), "", Rs!StoreID)
      If Trim(TxtStoreID.Text) <> "" Then
         TxtStoreName.Text = cn.Execute("Select StoreName from Stores Where StoreID = '" & TxtStoreID.Text & "'").Fields(0)
      Else
         TxtStoreName.Text = ""
      End If
      TxtSectorID.Text = IIf(IsNull(Rs!SectorID), "", Rs!SectorID)
      If TxtSectorID.Text = "" Then
         TxtSectorName.Text = ""
         TxtZoneNameSector.Text = ""
      Else
         With cn.Execute("select * from sectors s inner join Zones t on s.ZoneID = t.ZoneID where sectorid =" & Val(TxtSectorID.Text))
            If .RecordCount > 0 Then
               TxtSectorName.Text = !SectorName
               TxtZoneNameSector.Text = !ZoneName
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

Private Function FunGetMaxID() As String
  FunGetMaxID = cn.Execute("Select isnull(max(cast(substring(cast(accountno as varchar(6)),3,10) as int)),0) + 1 from chartofaccounts Where AccountNo like '61%' and isdetailed=1 ").Fields(0)
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
  'Rs.Find "PartyName like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
   Set Rs = New ADODB.Recordset
'   Rs.Open " Select * FROM Parties where PartyType='V' and PartyName like '%" & Replace(TxtFilter.Text, "'", "''") & "%' Order by PartyName", cn, adOpenStatic, adLockOptimistic
   vSQL = " Select P.* FROM Parties P inner join ChartOfAccounts C on C.AccountNo = P.PartyID where PartyType='V' and PartyName like '%" & Replace(TxtFilter.Text, "'", "''") & "%'" & IIf(ChkHideLockVendor.Value = 1, " and islocked = 0 and isLockParty = 0 ", "") & " Order by PartyName"
   Rs.Open vSQL, cn, adOpenStatic, adLockOptimistic
   Set Grid.DataSource = Rs
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
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_Parties ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
    If vIsNewRecord = False Then
        If TxtName.Text <> IIf(IsNull(Rs!partyname), "", Rs!partyname) Then
            cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated Name-" & Rs!partyname & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtAddress.Text <> IIf(IsNull(Rs!Address), "", Rs!Address) Then
            cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated Address-" & Rs!Address & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtCity.Text <> IIf(IsNull(Rs!City), "", Rs!City) Then
            cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated City-" & Rs!City & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPhone1.Text <> IIf(IsNull(Rs!Phone1), "", Rs!Phone1) Then
            cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated Phone1-" & Rs!Phone1 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
'        If TxtPhone2.Text <> IIf(IsNull(Rs!Phone2), "", Rs!Phone2) Then
'            CN.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated Phone2-" & Rs!Phone2 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
'        End If
        If TxtMobileNo.Text <> IIf(IsNull(Rs!Mobile), "", Rs!Mobile) Then
            cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated Mobile-" & Rs!Mobile & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtMobileNo2.Text <> IIf(IsNull(Rs!Mobile), "", Rs!Mobile2) Then
            cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated Mobile2-" & Rs!Mobile2 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtEmail.Text <> IIf(IsNull(Rs!Email), "", Rs!Email) Then
            cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated Email-" & Rs!Email & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtContactPerson.Text <> IIf(IsNull(Rs!ContactPerson), "", Rs!ContactPerson) Then
            cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated ContactPerson-" & Rs!ContactPerson & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If ChkLockVendor.Value <> Val(Rs!IsLockParty) Then
            cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null , 'Updated ChkLockVendor-" & Rs!IsLockParty & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
   Else
        cn.Execute ("Insert Into UserActivities values ('Vendors'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Function FunSelectZone(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchZone.Show vbModal, Me
        If SchZone.ParaOutZoneID = "" Then FunSelectZone = False: Exit Function
        TxtZoneID.Text = SchZone.ParaOutZoneID
    End If
    '---------------------------
    VStrSQL = "Select * FROM Zones where ZoneID=" & Val(TxtZoneID.Text)
    With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtZoneName.Text = !ZoneName
          FunSelectZone = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectZone = False
          .Close
          TxtZoneID.Text = ""
          TxtZoneName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnZone_Click()
   If FunSelectZone(ssButton, False) = True Then
     If TxtZoneName.Enabled Then TxtZoneName.SetFocus
   Else
      TxtZoneID.SetFocus
   End If
End Sub

Private Sub TxtZoneID_Change()
   If TxtZoneID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtZoneID.Name Then Exit Sub
   If TxtZoneName.Text <> "" Then TxtZoneName.Text = ""
End Sub

Private Sub TxtZoneID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtZoneID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtZoneID.Text = "" Then Exit Sub
   If TxtZoneName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectZone(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectZone(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSector(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSector.Show vbModal, Me
        If SchSector.ParaOutSectorID = "" Then FunSelectSector = False: Exit Function
        TxtSectorID.Text = SchSector.ParaOutSectorID
    End If
    '---------------------------
    VStrSQL = "Select * FROM Sectors s inner join Zones t on t.ZoneID = s.ZoneID where SectorID=" & Val(TxtSectorID.Text)
    With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtSectorName.Text = !SectorName
          TxtZoneNameSector.Text = !ZoneName
          FunSelectSector = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSector = False
          .Close
          TxtSectorID.Text = ""
          TxtSectorName.Text = ""
          TxtZoneNameSector.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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

Private Function FunSelectSaleAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectSaleAccount = False: Exit Function
        TxtSaleAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    VStrSQL = " Select * FROM ChartOfAccounts where AccountNo='" & TxtSaleAccountNo.Text & "'"
    With cn.Execute(VStrSQL)
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
      If ChkLockVendor.Enabled Then ChkLockVendor.SetFocus
   Else
      If TxtSaleAccountNo.Enabled Then TxtSaleAccountNo.SetFocus
   End If
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
   TxtID.Text = FunGetMaxID
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      If TxtSectorID.Enabled Then TxtSectorID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

