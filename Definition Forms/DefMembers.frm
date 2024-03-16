VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form DefMembers 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "DefMembers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkisReSeller 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "ReSeller"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8055
      TabIndex        =   52
      Top             =   8235
      Width           =   1320
   End
   Begin VB.TextBox TxtCreditLimit 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   12465
      MaxLength       =   6
      TabIndex        =   13
      Top             =   7253
      Width           =   855
   End
   Begin VB.TextBox TxtBarCode 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   11115
      MaxLength       =   16
      TabIndex        =   16
      Top             =   7853
      Width           =   2040
   End
   Begin VB.CheckBox ChkPaid 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Paid"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9585
      TabIndex        =   15
      Top             =   7898
      Width           =   1320
   End
   Begin VB.CheckBox ChkDateofBirth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8055
      TabIndex        =   49
      Top             =   6353
      Width           =   195
   End
   Begin VB.CheckBox ChkMembership 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9315
      TabIndex        =   48
      Top             =   6353
      Width           =   195
   End
   Begin VB.CheckBox ChkExpiry 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10890
      TabIndex        =   47
      Top             =   6353
      Width           =   195
   End
   Begin VB.TextBox TxtCNIC 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10665
      MaxLength       =   15
      TabIndex        =   8
      Top             =   5963
      Width           =   2655
   End
   Begin VB.TextBox TxtMemberTypeID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9630
      MaxLength       =   10
      TabIndex        =   1
      Top             =   2303
      Width           =   810
   End
   Begin VB.TextBox TxtPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   7965
      MaxLength       =   2
      TabIndex        =   39
      Tag             =   "NC"
      Top             =   2303
      Width           =   525
   End
   Begin VB.CheckBox ChkLockMember 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Lock Member"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8055
      TabIndex        =   14
      Top             =   7898
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
      Height          =   2850
      Left            =   12960
      TabIndex        =   35
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
         Height          =   2445
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         Tag             =   "NC"
         Text            =   "DefMembers.frx":0ECA
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
         TabIndex        =   37
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtPhone2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10665
      MaxLength       =   30
      TabIndex        =   6
      Top             =   5303
      Width           =   2595
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8010
      MaxLength       =   50
      TabIndex        =   12
      Top             =   7253
      Width           =   4455
   End
   Begin VB.TextBox TxtMobileNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8010
      MaxLength       =   25
      TabIndex        =   7
      Top             =   5963
      Width           =   2655
   End
   Begin VB.TextBox TxtPhone1 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8010
      MaxLength       =   30
      TabIndex        =   5
      Top             =   5303
      Width           =   2655
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8010
      MaxLength       =   30
      TabIndex        =   4
      Top             =   4658
      Width           =   5265
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   8010
      MaxLength       =   100
      TabIndex        =   3
      Top             =   3668
      Width           =   5265
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8490
      MaxLength       =   5
      TabIndex        =   0
      Top             =   2303
      Width           =   825
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8010
      MaxLength       =   30
      TabIndex        =   2
      Top             =   2948
      Width           =   5265
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2925
      MaxLength       =   30
      TabIndex        =   20
      Tag             =   "NC"
      Top             =   1943
      Width           =   4395
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5700
      Left            =   2040
      TabIndex        =   21
      Top             =   2273
      Width           =   5295
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
      stylesets(0).Picture=   "DefMembers.frx":0F55
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
      Columns(0).Caption=   "Member ID"
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
      _ExtentX        =   9340
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
      Left            =   2985
      TabIndex        =   22
      Top             =   8783
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
      MICON           =   "DefMembers.frx":0F71
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   4305
      TabIndex        =   23
      Top             =   8783
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
      MICON           =   "DefMembers.frx":0F8D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   5625
      TabIndex        =   24
      Top             =   8783
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
      MICON           =   "DefMembers.frx":0FA9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8430
      TabIndex        =   17
      Top             =   8783
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
      MICON           =   "DefMembers.frx":0FC5
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   9750
      TabIndex        =   18
      Top             =   8783
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
      MICON           =   "DefMembers.frx":0FE1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   11070
      TabIndex        =   19
      Top             =   8783
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
      MICON           =   "DefMembers.frx":0FFD
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtMemberType 
      Height          =   315
      Left            =   10800
      TabIndex        =   40
      Top             =   2303
      Width           =   1965
      _ExtentX        =   3466
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
   Begin JeweledBut.JeweledButton BtnMemberType 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10440
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   2303
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
      MICON           =   "DefMembers.frx":1019
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpMembershipDate 
      Height          =   315
      Left            =   9480
      TabIndex        =   10
      Tag             =   "NC"
      Top             =   6608
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpExpiryDate 
      Height          =   315
      Left            =   10875
      TabIndex        =   11
      Tag             =   "NC"
      Top             =   6608
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpDateofBirth 
      Height          =   315
      Left            =   8055
      TabIndex        =   9
      Tag             =   "NC"
      Top             =   6608
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
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit"
      Height          =   195
      Left            =   12465
      TabIndex        =   51
      Top             =   7043
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Codes"
      Height          =   195
      Index           =   1
      Left            =   11115
      TabIndex        =   50
      Top             =   7673
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date"
      Height          =   195
      Left            =   11115
      TabIndex        =   46
      Top             =   6353
      Width           =   810
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MemberShip Date"
      Height          =   195
      Index           =   0
      Left            =   9540
      TabIndex        =   45
      Top             =   6353
      Width           =   1275
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNIC No"
      Height          =   195
      Left            =   10665
      TabIndex        =   44
      Top             =   5738
      Width           =   630
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Type ID"
      Height          =   195
      Left            =   9630
      TabIndex        =   43
      Top             =   2093
      Width           =   1185
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Type"
      Height          =   195
      Left            =   10890
      TabIndex        =   42
      Top             =   2093
      Width           =   975
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
      TabIndex        =   38
      Top             =   495
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Members"
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
      TabIndex        =   34
      Top             =   270
      Width           =   1230
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      Height          =   195
      Left            =   8280
      TabIndex        =   33
      Top             =   6353
      Width           =   885
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      Height          =   195
      Left            =   8010
      TabIndex        =   32
      Top             =   7043
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
      Height          =   195
      Left            =   8010
      TabIndex        =   31
      Top             =   5753
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Nos"
      Height          =   195
      Left            =   8010
      TabIndex        =   30
      Top             =   5093
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   8010
      TabIndex        =   29
      Top             =   4433
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   8010
      TabIndex        =   28
      Top             =   3458
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      Height          =   195
      Left            =   7965
      TabIndex        =   27
      Top             =   2093
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   8010
      TabIndex        =   26
      Top             =   2738
      Width           =   420
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   509
      X2              =   508
      Y1              =   128.533
      Y2              =   534.533
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   2310
      TabIndex        =   25
      Top             =   2033
      Width           =   510
   End
End
Attribute VB_Name = "DefMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.

Private Function FunSelectMemberType(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchMemberTypes.Show vbModal, Me
        If SchMemberTypes.ParaOutMemberTypeID = "" Then FunSelectMemberType = False: Exit Function
        TxtMemberTypeID.Text = SchMemberTypes.ParaOutMemberTypeID
    End If
    '---------------------------
    vStrSQL = "Select * FROM MemberTypes where MemberTypeID = " & Val(TxtMemberTypeID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtMemberType.Text = !MemberType
          FunSelectMemberType = True
          .Close
          Exit Function
      Else
          FunSelectMemberType = False
          .Close
          TxtMemberTypeID.Text = ""
          TxtMemberType.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub ChkDateofBirth_Click()
   DtpDateofBirth.Enabled = IIf(ChkDateofBirth.Value = 1, True, False)
End Sub

Private Sub ChkExpiry_Click()
   DtpExpiryDate.Enabled = IIf(ChkExpiry.Value = 1, True, False)
End Sub

Private Sub ChkisReSeller_Click()
   If ActiveControl.Name <> ChkisReSeller.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkMembership_Click()
   DtpMembershipDate.Enabled = IIf(ChkMembership.Value = 1, True, False)
End Sub

Private Sub TxtMemberTypeID_Change()
   If TxtMemberTypeID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtMemberTypeID.Name Then Exit Sub
   If TxtMemberType.Text <> "" Then TxtMemberType.Text = ""
End Sub

Private Sub TxtMemberTypeID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtMemberTypeID.Name Then Exit Sub
   If Trim(TxtMemberTypeID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectMemberType(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectMemberType(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnMemberType_Click()
   If FunSelectMemberType(ssButton, False) = True Then
      TxtName.SetFocus
   Else
      TxtMemberTypeID.SetFocus
   End If
End Sub

Private Sub ChkLockMember_Click()
   If ActiveControl.Name <> ChkLockMember.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkPaid_Click()
   If ActiveControl.Name <> ChkPaid.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
     If ActiveControl.Name = Grid.Name Then Call Grid_DblClick: Exit Sub
     keybd_event 9, 1, 1, 1
     KeyCode = 0
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
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtMemberTypeID.Name: If FunSelectMemberType(ssFunctionKey, True) = True Then TxtMemberTypeID.SetFocus
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" And Me.ActiveControl.Tag = "" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
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
  FormStatus = SelectionMode
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniMembers", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
  Dim vtbl As String
  If Rs.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
    Dim vid As String
    vid = Rs!MemberID
    vtbl = Common.ChildDataExists("Members", "MemberID='" & vid & "'", "") & Common.ChildDataExists("ChartoFAccounts", "AccountNo='" & vid & "'", "Members")
    If vtbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
      Exit Sub
    End If
    cn.BeginTrans
    'Call ActivityLog("Members", eDelete, , , TxtPrefix.Text & TxtID.Text)
    vid = Rs!Prefix & Rs!MemberID
    Rs.Delete
    cn.Execute ("Delete From ChartOfAccounts Where AccountNo = '" & vid & "'")
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
  Dim vStrSQL As String
  If FunValidation = False Then Exit Sub
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniMembers", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
  'If vIsNewRecord = False Then Call ActivityLog("Members", eEdit, , , TxtPrefix.Text & TxtID.Text)
  Set Rs = New ADODB.Recordset
  Rs.CursorLocation = adUseClient
  Rs.Open " Select * FROM Members where MemberID = " & Val(TxtID.Text), cn, adOpenStatic, adLockOptimistic
  cn.BeginTrans
  If vIsNewRecord Then
    vStrSQL = "Insert into chartofaccounts values (" & _
    "'" & TxtPrefix.Text & TxtID.Text & "',1,'" & Replace(TxtName.Text, "'", "''") & "','Members',2,'" & Replace(TxtAddress.Text, "'", "''") & "','64',0,0,1," & ChkLockMember.Value & ",1,0,' ',0,0,'" & Date & "',0)"
    cn.Execute vStrSQL
    Rs.AddNew
    Rs!MemberID = Val(TxtID.Text)
    Rs!Prefix = Val(TxtPrefix.Text)
    'CN.Execute ("Insert into users values (" & CN.Execute("Select isnull(max(userno),0) + 1 from users").Fields(0) & ",'" & TxtName.Text & "','',0,0,0,1,'" & Rs!EmpID & "')")
    Rs!isChanged = 0
  Else
    Rs!isChanged = 1
    Rs!IsSync = 0
    cn.Execute ("Update Chartofaccounts set Accountname='" & Replace(TxtName.Text, "'", "''") & "',Narration = '" & Replace(TxtAddress.Text, "'", "''") & "', isLocked = " & ChkLockMember.Value & ", isChanged = 1 Where AccountNo = '" & Rs!Prefix & Rs!MemberID & "'")
    'CN.Execute ("Update users set username='" & TxtName.Text & "' Where EmpID = '" & Rs!EmpID & "'")
  End If
  If vIsNewRecord Then
  End If
  Rs!MemberName = TxtName.Text
  Rs!MemberTypeID = IIf(Trim(TxtMemberTypeID.Text) = "", Null, Val(TxtMemberTypeID.Text))
  Rs!Address = TxtAddress.Text
  Rs!City = TxtCity.Text
  Rs!Phone1 = TxtPhone1.Text
  Rs!Phone2 = TxtPhone2.Text
  Rs!Mobile = TxtMobileNo.Text
  Rs!DateofBirth = IIf(ChkDateofBirth.Value = 1, IIf(DtpDateofBirth.DateValue <> "", DtpDateofBirth.DateValue, Null), Null)
  Rs!MembershipDate = IIf(ChkMembership.Value = 1, IIf(DtpMembershipDate.DateValue <> "", DtpMembershipDate.DateValue, Null), Null)
  Rs!ExpiryDate = IIf(ChkExpiry.Value = 1, IIf(DtpExpiryDate.DateValue <> "", DtpExpiryDate.DateValue, Null), Null)
  Rs!Email = TxtEmail.Text
  Rs!CreditLimit = Val(TxtCreditLimit.Text)
  Rs!IsLockMember = ChkLockMember.Value
  Rs!isReSeller = ChkisReSeller.Value
  Rs!IsPaid = ChkPaid.Value
  Rs!BarCode = IIf(Trim(TxtBarCode.Text) = "", Null, Trim(TxtBarCode.Text))
  Rs.Update
  cn.CommitTrans
  Set Rs = New ADODB.Recordset
  Rs.CursorLocation = adUseClient
  Rs.Open " Select * FROM Members", cn, adOpenStatic, adLockOptimistic
 'If vIsNewRecord = True Then Call ActivityLog("Members", eAdd, , , TxtPrefix.Text & TxtID.Text)
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
      MsgBox "Please specify a Member ID", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
    If Not IsNumeric(TxtID.Text) Then
      MsgBox "The Member ID must be numeric", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
  End If
  If Trim(TxtName.Text) = "" Then
    MsgBox "Please specify a Member Name", vbExclamation, "Alert"
    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
    Exit Function
  End If
  If TxtID.Enabled = True And cn.Execute("select count(*) from Members where MemberID = " & Val(TxtID.Text)).Fields(0) > 0 Then
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
   SetWindowText Me.hWnd, "Members"
   HelpLocation Me
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open "Select * FROM Members ", cn, adOpenStatic, adLockOptimistic
   Grid.Columns("ID").DataField = "MemberID"
   Grid.Columns("Name").DataField = "MemberName"
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
      TxtID.Enabled = True
      TxtID.Text = FunGetMaxID
      TxtPrefix.Enabled = False
      TxtPrefix.Text = "64"
      ChkMembership.Value = 0
      ChkExpiry.Value = 0
      DtpMembershipDate.DateValue = Date
      DtpExpiryDate.DateValue = DateAdd("y", 1, Date) - 2
      TxtFilter.Text = ""
      Grid.Enabled = False
      TxtFilter.Enabled = False
      Set Grid.DataSource = Rs
      'TxtID.Enabled = False
      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
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
      TxtID.Text = Grid.Columns("ID").Text
      TxtName.Text = Grid.Columns("Name").Text
      TxtMemberTypeID.Text = IIf(IsNull(Rs!MemberTypeID), "", Rs!MemberTypeID)
      If Trim(TxtMemberTypeID.Text) <> "" Then
         TxtMemberType.Text = cn.Execute("select MemberType from MemberTypes where MemberTypeID = " & Val(TxtMemberTypeID.Text)).Fields(0).Value
      Else
         TxtMemberType.Text = ""
      End If
      TxtAddress.Text = IIf(IsNull(Rs!Address), "", Rs!Address)
      TxtCity.Text = IIf(IsNull(Rs!City), "", Rs!City)
      TxtPhone1.Text = IIf(IsNull(Rs!Phone1), "", Rs!Phone1)
      TxtPhone2.Text = IIf(IsNull(Rs!Phone2), "", Rs!Phone2)
      TxtMobileNo.Text = IIf(IsNull(Rs!Mobile), "", Rs!Mobile)
      If IsNull(Rs!DateofBirth) Then
         DtpDateofBirth.DateValue = ""
         ChkDateofBirth.Value = 0
      Else
         DtpDateofBirth.DateValue = Rs!DateofBirth
         ChkDateofBirth.Value = 1
      End If
      TxtCNIC.Text = IIf(IsNull(Rs!CNIC), "", Rs!CNIC)
      If IsNull(Rs!MembershipDate) Then
         DtpMembershipDate.DateValue = ""
         ChkMembership.Value = 0
      Else
         DtpMembershipDate.DateValue = Rs!MembershipDate
         ChkMembership.Value = 1
      End If
      If IsNull(Rs!ExpiryDate) Then
         DtpExpiryDate.DateValue = ""
         ChkExpiry.Value = 0
      Else
         DtpExpiryDate.DateValue = Rs!ExpiryDate
         ChkExpiry.Value = 1
      End If
      TxtEmail.Text = IIf(IsNull(Rs!Email), "", Rs!Email)
      TxtCreditLimit.Text = IIf(IsNull(Rs!CreditLimit), "", Rs!CreditLimit)
      ChkLockMember.Value = Abs(Rs!IsLockMember)
      ChkPaid.Value = IIf(IsNull(Rs!IsPaid), "0", Abs(Rs!IsPaid))
      ChkisReSeller.Value = IIf(IsNull(Rs!isReSeller), "0", Abs(Rs!isReSeller))
      TxtBarCode.Text = IIf(IsNull(Rs!BarCode), "", Rs!BarCode)
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
      ElseIf TypeOf ctl Is ComboBox Then
         ctl.Enabled = Enable
      ElseIf TypeOf ctl Is CheckBox Then
         ctl.Value = 0
         ctl.Enabled = Enable
      ElseIf TypeOf ctl Is SSDateCombo Then
         If ctl.Tag = "" Then
            ctl.Enabled = Enable
         End If
      ElseIf TypeOf ctl Is JeweledButton Then
         If ctl.Tag <> "" Then ctl.Enabled = Enable
      End If
   Next
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
   On Error GoTo ErrorHandler
   If IsNull(ObjRegistry.MemberMin) Or IsNull(ObjRegistry.MemberMax) Or (ObjRegistry.MemberMin = 0 And ObjRegistry.MemberMax = 0) Then
      FunGetMaxID = cn.Execute("Select isnull(max(MemberID),0) + 1 from Members ").Fields(0)
   Else
      FunGetMaxID = cn.Execute("Select isnull(max(MemberID),0) + 1 from Members where MemberID between " & ObjRegistry.MemberMin & " and " & ObjRegistry.MemberMax).Fields(0)
   End If
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
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

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, vbKeyA To vbKeyZ, Asc("@"), Asc("_"), Asc("-"), Asc(" "), vbKeyBack, Asc("a") To Asc("z"), Asc(".")
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtDateofBirth_KeyPress(KeyAscii As Integer)
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
  'Rs.Find "MemberName like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open " Select * FROM Members where MemberName like '%" & Replace(TxtFilter.Text, "'", "''") & "%' Order by MemberName", cn, adOpenStatic, adLockOptimistic
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
