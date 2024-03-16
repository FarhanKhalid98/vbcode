VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmAccountsBalancesDiff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Balances Difference"
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtaccountName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   10515
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1755
      Width           =   3465
   End
   Begin VB.TextBox TxtAccountNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9135
      MaxLength       =   10
      TabIndex        =   19
      Top             =   1740
      Width           =   1020
   End
   Begin VB.CheckBox ChkShowOnlyDifference 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Only Difference"
      Height          =   255
      Left            =   3150
      TabIndex        =   11
      Top             =   1800
      Width           =   2325
   End
   Begin VB.CheckBox ChkOpening 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Include Opening"
      Height          =   255
      Left            =   4890
      TabIndex        =   5
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2325
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7800
      TabIndex        =   1
      Top             =   9945
      Width           =   1275
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
      MICON           =   "FrmAccountsBalancesDiff.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnRestore 
      Height          =   420
      Left            =   6285
      TabIndex        =   0
      Top             =   9945
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Refresh"
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
      MICON           =   "FrmAccountsBalancesDiff.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridOpening 
      Height          =   3435
      Left            =   135
      TabIndex        =   3
      Top             =   2115
      Width           =   8400
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   5
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   12566463
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
      stylesets(0).Picture=   "FrmAccountsBalancesDiff.frx":0038
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   5
      Columns(0).Width=   1561
      Columns(0).Caption=   "Query NO"
      Columns(0).Name =   "QueryNO"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Query Name"
      Columns(1).Name =   "QueryName"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Opening Credit"
      Columns(2).Name =   "OpeningCredit"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Opening Debit"
      Columns(3).Name =   "OpeningDebit"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3096
      Columns(4).Caption=   "Diff"
      Columns(4).Name =   "Diff"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   14817
      _ExtentY        =   6059
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridSubOpening 
      Height          =   3420
      Left            =   8550
      TabIndex        =   4
      Top             =   2115
      Width           =   6675
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   13817275
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
      stylesets(0).Picture=   "FrmAccountsBalancesDiff.frx":0054
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
      BackColorEven   =   15724527
      BackColorOdd    =   16777215
      RowHeight       =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   3
      Columns(0).Width=   5477
      Columns(0).Caption=   "Sub Query Name"
      Columns(0).Name =   "SubQueryName"
      Columns(0).DataField=   "Column 5"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2858
      Columns(1).Caption=   "Opening Credit"
      Columns(1).Name =   "OpeningCredit"
      Columns(1).DataField=   "Column 6"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2831
      Columns(2).Caption=   "Opening Debit"
      Columns(2).Name =   "OpeningDebit"
      Columns(2).DataField=   "Column 7"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   11774
      _ExtentY        =   6032
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   1395
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpTo 
      Height          =   315
      Left            =   3585
      TabIndex        =   7
      Top             =   1380
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
   Begin JeweledBut.JeweledButton CmdPreview 
      Height          =   420
      Left            =   4890
      TabIndex        =   10
      Top             =   1305
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   741
      TX              =   "Execute Accounts Ballance"
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
      MICON           =   "FrmAccountsBalancesDiff.frx":0070
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridDetail 
      Height          =   3435
      Left            =   135
      TabIndex        =   13
      Top             =   6030
      Width           =   8400
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   5
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   12566463
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
      stylesets(0).Picture=   "FrmAccountsBalancesDiff.frx":008C
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   5
      Columns(0).Width=   1561
      Columns(0).Caption=   "Query NO"
      Columns(0).Name =   "QueryNO"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Query Name"
      Columns(1).Name =   "QueryName"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Credit"
      Columns(2).Name =   "Credit"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Debit"
      Columns(3).Name =   "Debit"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3096
      Columns(4).Caption=   "Diff"
      Columns(4).Name =   "Diff"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   14817
      _ExtentY        =   6059
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridSubDetail 
      Height          =   3420
      Left            =   8550
      TabIndex        =   14
      Top             =   6030
      Width           =   6675
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   13817275
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
      stylesets(0).Picture=   "FrmAccountsBalancesDiff.frx":00A8
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
      BackColorEven   =   15724527
      BackColorOdd    =   16777215
      RowHeight       =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   3
      Columns(0).Width=   5477
      Columns(0).Caption=   "Sub Query Name"
      Columns(0).Name =   "SubQueryName"
      Columns(0).DataField=   "Column 5"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2858
      Columns(1).Caption=   "Credit"
      Columns(1).Name =   "Credit"
      Columns(1).DataField=   "Column 6"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2831
      Columns(2).Caption=   "Debit"
      Columns(2).Name =   "Debit"
      Columns(2).DataField=   "Column 7"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   11774
      _ExtentY        =   6032
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
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10155
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1125
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
      MICON           =   "FrmAccountsBalancesDiff.frx":00C4
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   9135
      TabIndex        =   16
      Top             =   1125
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   2
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
   Begin SITextBox.Txt TxtOrganizatonName 
      Height          =   315
      Left            =   10515
      TabIndex        =   18
      Tag             =   "nc"
      Top             =   1125
      Width           =   3465
      _ExtentX        =   6112
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
   Begin JeweledBut.JeweledButton BtnAccount 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   10155
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1740
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmAccountsBalancesDiff.frx":00E0
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
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
      Left            =   9135
      TabIndex        =   25
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
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
      Left            =   10515
      TabIndex        =   24
      Top             =   900
      Width           =   1620
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10530
      TabIndex        =   23
      Top             =   1530
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9135
      TabIndex        =   22
      Top             =   1530
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   135
      TabIndex        =   15
      Top             =   5670
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   135
      TabIndex        =   12
      Top             =   1710
      Width           =   1020
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2280
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3600
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts Balances Difference"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   2
      Top             =   270
      Width           =   4200
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmAccountsBalancesDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset
Dim vStrSQL, vStrPara, vActionNo, vSSID, vRSID As String
Dim vCounter As Integer
Public ParaOutUserNo As String
Dim vOrder As String, vDirection As String, vCol As Byte, vFilter As String, vUser As String, vAction As String

Private Sub LoadGrid()
   LoadGridOpening
   LoadGridDetail
End Sub
Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   Me.ParaOutUserNo = ""
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOrganization_Click()
   If FunSelectOrganizaton(ssButton, False) = True Then
      TxtAccountNo.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnRestore_Click()
On Error GoTo ErrorHandler
      Call LoadGrid
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub ChkShowOnlyDifference_Click()
   LoadGrid
End Sub

Private Sub CmdPreview_Click()
On Error GoTo ErrorHandler

   Me.MousePointer = vbHourglass
   
   CN.Execute "EXECUTE SPAccountsBalancesNew '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & ChkOpening.Value
   'Calculate Average Cost
   CN.Execute "exec SPAverageCost '" & DtpTo.DateValue & "'"
   'Second Insert Closing Stock
   CN.Execute "EXECUTE SPClosingStock '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
   Me.MousePointer = vbDefault
   LoadGrid
   
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    keybd_event 9, 1, 1, 1
    KeyCode = 0
  ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtOrganizationID.Name: If FunSelectOrganizaton(ssFunctionKey, True) = True Then TxtAccountNo.SetFocus
         Case TxtAccountNo.Name: If FunSelectAccount(ssFunctionKey, True) = True Then CmdPreview.SetFocus
      End Select
  End If
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Accounts Balances Difference"
   
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   
   Call LoadGrid
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   On Error GoTo ErrorHandler
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadGridOpening()
   On Error GoTo ErrorHandler
         
   vStrSQL = "Select QueryNO, QueryName, sum(isnull(OpeningCredit,0)) OpeningCredit, Sum(isnull(OpeningDebit,0)) OpeningDebit, sum(isnull(OpeningCredit,0))- Sum(isnull(OpeningDebit,0)) Diff From AccountsBalancesOpening" & vbCrLf _
             & "Where 1=1 " & IIf(Trim(TxtOrganizationID.Text) = "", "", " And OrganizationID = " & TxtOrganizationID.Text) & IIf(TxtAccountNo.Text = "", "", " and AccountNo = " & TxtAccountNo.Text) & vbCrLf _
             & "Group By QueryNO, QueryName" & vbCrLf _
             & IIf(ChkShowOnlyDifference = 1, "having sum(isnull(OpeningCredit,0))- Sum(isnull(OpeningDebit,0)) <> 0", "") & vbCrLf _
             & "Order by QueryNO"
             
   
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   GridOpening.Redraw = False
   GridOpening.MoveFirst
   GridOpening.RemoveAll
   GridOpening.AllowAddNew = True
   While Not Rs.EOF
      GridOpening.Columns("QueryNO").Text = Rs!QueryNO
      GridOpening.Columns("QueryName").Text = Rs!QueryName
      GridOpening.Columns("OpeningCredit").Value = Rs!OpeningCredit
      GridOpening.Columns("OPeningDebit").Value = Rs!OPeningDebit
      GridOpening.Columns("Diff").Value = Rs!Diff
      GridOpening.AddNew
      Rs.MoveNext
   Wend
   GridOpening.MoveFirst
   GridOpening.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub GridOpening_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo ErrorHandler
   Call LoadGridSubOpening
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Call ShowErrorMessage
End Sub



Private Sub LoadGridSubOpening()
On Error GoTo ErrorHandler
   If GridOpening.Rows <= 1 Then Exit Sub
      vStrSQL = "Select QueryNO, QueryName, SubQueryName, sum(isnull(OpeningCredit,0)) OpeningCredit, Sum(isnull(OpeningDebit,0)) OpeningDebit, sum(isnull(OpeningCredit,0))- Sum(isnull(OpeningDebit,0)) Diff From AccountsBalancesOpening" & vbCrLf _
             & "Group By QueryNO, QueryName, SubQueryName" & vbCrLf _
             & "having QueryNO = " & GridOpening.Columns("QueryNO").Text & vbCrLf _
             & "--having sum(OpeningCredit)- Sum(OPeningDebit) <> 0" & vbCrLf _
             & "Order by QueryNO"
   
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set GridSubOpening.DataSource = Rs
   GridSubOpening.Columns("SubQueryName").DataField = "SubQueryName"
   GridSubOpening.Columns("OpeningCredit").DataField = "OpeningCredit"
   GridSubOpening.Columns("OpeningDebit").DataField = "OpeningDebit"
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadGridDetail()
   On Error GoTo ErrorHandler
         
   vStrSQL = "Select QueryNO, QueryName, sum(isnull(Credit,0)) Credit, Sum(isnull(Debit,0)) Debit, sum(isnull(Credit,0))- Sum(isnull(Debit,0)) Diff From AccountsBalancesDetail" & vbCrLf _
             & "Where 1=1 " & IIf(Trim(TxtOrganizationID.Text) = "", "", " And OrganizationID = " & TxtOrganizationID.Text) & IIf(TxtAccountNo.Text = "", "", " and AccountNo = " & TxtAccountNo.Text) & vbCrLf _
             & "Group By QueryNO, QueryName" & vbCrLf _
             & IIf(ChkShowOnlyDifference = 1, "having sum(isnull(Credit,0))- Sum(isnull(Debit,0)) <> 0", "") & vbCrLf _
             & "Order by QueryNO"
             
   
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   GridDetail.Redraw = False
   GridDetail.MoveFirst
   GridDetail.RemoveAll
   GridDetail.AllowAddNew = True
   While Not Rs.EOF
      GridDetail.Columns("QueryNO").Text = Rs!QueryNO
      GridDetail.Columns("QueryName").Text = Rs!QueryName
      GridDetail.Columns("Credit").Value = Rs!Credit
      GridDetail.Columns("Debit").Value = Rs!Debit
      GridDetail.Columns("Diff").Value = Rs!Diff
      GridDetail.AddNew
      Rs.MoveNext
   Wend
   GridDetail.MoveFirst
   GridDetail.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub GridDetail_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo ErrorHandler
   Call LoadGridSubDetail
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Call ShowErrorMessage
End Sub



Private Sub LoadGridSubDetail()
On Error GoTo ErrorHandler
   If GridDetail.Rows <= 1 Then Exit Sub
      vStrSQL = "Select QueryNO, QueryName, SubQueryName, sum(isnull(Credit,0)) Credit, Sum(isnull(Debit,0)) Debit, sum(isnull(Credit,0))- Sum(isnull(Debit,0)) Diff From AccountsBalancesDetail" & vbCrLf _
             & "Group By QueryNO, QueryName, SubQueryName" & vbCrLf _
             & "having QueryNO = " & GridDetail.Columns("QueryNO").Text & vbCrLf _
             & "--having sum(Credit)- Sum(Debit) <> 0" & vbCrLf _
             & "Order by QueryNO"
   
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set GridSubDetail.DataSource = Rs
   GridSubDetail.Columns("SubQueryName").DataField = "SubQueryName"
   GridSubDetail.Columns("Credit").DataField = "Credit"
   GridSubDetail.Columns("Debit").DataField = "Debit"
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub TxtOrganizationID_Change()
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizatonName.Text <> "" Then TxtOrganizatonName.Text = ""
End Sub

Private Sub TxtOrganizationID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtOrganizationID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectOrganizaton(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganizaton(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectOrganizaton(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganizaton = False: Exit Function
       TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    If TxtOrganizationID.Text = "" Then FunSelectOrganizaton = False: Exit Function
    vStrSQL = " Select * FROM Organizations where OrganizationID='" & TxtOrganizationID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizatonName.Text = !OrganizationName
          FunSelectOrganizaton = True
          .Close
          Exit Function
      Else
          FunSelectOrganizaton = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizatonName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInDetail = "False"
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
        TxtAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    If Trim(TxtAccountNo.Text) = "" Then Exit Function
    vStrSQL = " Select AccountNo, AccountName FROM ChartofAccounts where AccountNo= '" & TxtAccountNo.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtaccountName.Text = !AccountName
          FunSelectAccount = True
          Exit Function
      Else
          FunSelectAccount = False
          TxtAccountNo.Text = ""
          TxtaccountName.Text = ""
      End If
      .Close
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtAccountNo_Change()
   If TxtAccountNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   If TxtaccountName.Text <> "" Then TxtaccountName.Text = ""
End Sub

Private Sub TxtAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtAccountNo.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAccount_Click()
   On Error GoTo ErrorHandler
   If FunSelectAccount(ssButton, True) = True Then
      CmdPreview.SetFocus
   Else
      TxtAccountNo.SetFocus
   End If
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

