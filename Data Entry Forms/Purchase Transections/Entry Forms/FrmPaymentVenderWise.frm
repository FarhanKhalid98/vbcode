VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmPaymentVenderWise 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmPaymentVenderWise.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9825
      TabIndex        =   48
      Top             =   8190
      Width           =   1290
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmPaymentVenderWise.frx":0ECA
      Left            =   5010
      List            =   "FrmPaymentVenderWise.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Tag             =   "1"
      Top             =   8085
      Width           =   3585
   End
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   8610
      TabIndex        =   45
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   8100
      Width           =   1170
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
      Left            =   13560
      TabIndex        =   29
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
         Height          =   3750
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   30
         Tag             =   "NC"
         Text            =   "FrmPaymentVenderWise.frx":0ECE
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
         TabIndex        =   31
         Top             =   90
         Width           =   135
      End
   End
   Begin SITextBox.Txt TxtPaymentID 
      Height          =   315
      Left            =   1785
      TabIndex        =   0
      Top             =   2318
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   9
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
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9000
      TabIndex        =   12
      Top             =   8625
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
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
      MICON           =   "FrmPaymentVenderWise.frx":0FC1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7687
      TabIndex        =   8
      Top             =   8618
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
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
      MICON           =   "FrmPaymentVenderWise.frx":0FDD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5077
      TabIndex        =   10
      Top             =   8618
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
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
      MICON           =   "FrmPaymentVenderWise.frx":0FF9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10297
      TabIndex        =   13
      Top             =   8618
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
      MICON           =   "FrmPaymentVenderWise.frx":1015
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6382
      TabIndex        =   9
      Top             =   8618
      Width           =   1275
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
      MICON           =   "FrmPaymentVenderWise.frx":1031
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtVenderID 
      Height          =   315
      Left            =   1725
      TabIndex        =   4
      Top             =   3338
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   16
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
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton BtnVender 
      Height          =   330
      Left            =   2700
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3338
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
      MICON           =   "FrmPaymentVenderWise.frx":104D
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtVenderName 
      Height          =   315
      Left            =   3060
      TabIndex        =   15
      Top             =   3338
      Width           =   2505
      _ExtentX        =   4419
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
      Masked          =   5
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4320
      Left            =   1725
      TabIndex        =   20
      Top             =   3653
      Width           =   11910
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   7
      stylesets.count =   1
      stylesets(0).Name=   "Select"
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
      stylesets(0).Picture=   "FrmPaymentVenderWise.frx":1069
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
      Columns.Count   =   7
      Columns(0).Width=   2355
      Columns(0).Caption=   "Vender ID"
      Columns(0).Name =   "VenderID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4419
      Columns(1).Caption=   "Vender Name"
      Columns(1).Name =   "VenderName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2699
      Columns(2).Caption=   "Previous Payable"
      Columns(2).Name =   "PreviousPayable"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2117
      Columns(3).Caption=   "Amount"
      Columns(3).Name =   "Amount"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   4683
      Columns(4).Caption=   "Description"
      Columns(4).Name =   "Description"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1799
      Columns(5).Caption=   "Discount"
      Columns(5).Name =   "Discount"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2434
      Columns(6).Caption=   "Remaining Amount"
      Columns(6).Name =   "FinalCredit"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   21008
      _ExtentY        =   7620
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpPaymentDate 
      Height          =   315
      Left            =   2745
      TabIndex        =   1
      Top             =   2318
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
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
   Begin SITextBox.Txt TxtPreviousPayable 
      Height          =   315
      Left            =   5565
      TabIndex        =   16
      Top             =   3338
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   1
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   7095
      TabIndex        =   5
      Top             =   3338
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      DecimalPoint    =   2
      IntegralPoint   =   7
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtDiscount 
      Height          =   315
      Left            =   10950
      TabIndex        =   7
      Top             =   3338
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtFinalCredit 
      Height          =   315
      Left            =   11970
      TabIndex        =   17
      Top             =   3338
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      Alignment       =   1
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3772
      TabIndex        =   11
      Top             =   8618
      Width           =   1275
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
      MICON           =   "FrmPaymentVenderWise.frx":1085
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   8295
      TabIndex        =   6
      Top             =   3338
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   100
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
   Begin JeweledBut.JeweledButton BtnEmployee 
      Height          =   330
      Left            =   5115
      TabIndex        =   33
      Top             =   2333
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
      MICON           =   "FrmPaymentVenderWise.frx":10A1
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmployeeName 
      Height          =   315
      Left            =   5475
      TabIndex        =   34
      Top             =   2333
      Width           =   1455
      _ExtentX        =   2566
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
   Begin SITextBox.Txt TxtEmployeeID 
      Height          =   315
      Left            =   4155
      TabIndex        =   2
      Top             =   2333
      Width           =   960
      _ExtentX        =   1693
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
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   7065
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   2333
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
      Left            =   8130
      TabIndex        =   35
      Tag             =   "NC"
      Top             =   2333
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
      Left            =   7770
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2333
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
      MICON           =   "FrmPaymentVenderWise.frx":10BD
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBankID 
      Height          =   315
      Left            =   10050
      TabIndex        =   41
      Top             =   2318
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
   Begin SITextBox.Txt TxtBankName 
      Height          =   315
      Left            =   11115
      TabIndex        =   42
      Top             =   2318
      Width           =   2145
      _ExtentX        =   3784
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
   Begin JeweledBut.JeweledButton BtnBank 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   10755
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2318
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
      MICON           =   "FrmPaymentVenderWise.frx":10D9
      BC              =   12632256
      FC              =   0
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
      Left            =   4245
      TabIndex        =   47
      Top             =   8130
      Width           =   570
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank ID"
      Height          =   195
      Left            =   10050
      TabIndex        =   44
      Top             =   2093
      Width           =   585
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      Height          =   255
      Left            =   4155
      TabIndex        =   40
      Top             =   2138
      Width           =   1215
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      Height          =   375
      Left            =   5460
      TabIndex        =   39
      Top             =   2138
      Width           =   1215
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   8250
      TabIndex        =   38
      Top             =   2093
      Width           =   1350
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   7065
      TabIndex        =   37
      Top             =   2093
      Width           =   1095
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
      Left            =   11250
      TabIndex        =   32
      Top             =   585
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   8280
      TabIndex        =   28
      Top             =   3143
      Width           =   795
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment (Vender Wise)"
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
      Left            =   1920
      TabIndex        =   27
      Top             =   120
      Width           =   4125
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   10950
      TabIndex        =   26
      Top             =   3143
      Width           =   630
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Amount"
      Height          =   195
      Left            =   11970
      TabIndex        =   25
      Top             =   3143
      Width           =   1335
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   7110
      TabIndex        =   24
      Top             =   3143
      Width           =   540
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Payable"
      Height          =   195
      Left            =   5565
      TabIndex        =   23
      Top             =   3143
      Width           =   1230
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender ID"
      Height          =   195
      Left            =   1725
      TabIndex        =   22
      Top             =   3143
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Name"
      Height          =   195
      Left            =   3120
      TabIndex        =   21
      Top             =   3143
      Width           =   975
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date"
      Height          =   195
      Left            =   2760
      TabIndex        =   19
      Top             =   2123
      Width           =   1005
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment ID"
      Height          =   195
      Left            =   1785
      TabIndex        =   18
      Top             =   2123
      Width           =   825
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmPaymentVenderWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter, vGridRows As Integer
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim ssql As String
Dim i As Integer
Dim vStrSQL, vRandomID As String
Dim vMaxBinID As Integer
Dim vNoofPrints As Byte
Dim vPrinter() As String
'----------------------------------

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
      TxtBankID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub CalculateBody()
   TxtFinalCredit.Text = Round(Val(TxtPreviousPayable.Text) - Val(TxtAmount.Text) - Val(TxtDiscount.Text), 2)
End Sub

Private Sub BtnClear_Click()
On Error GoTo ErrorHandler
    '''''''''''''''''' ActivityLogBin For Clear Action
'      Call DeleteTempActivityLogBin(vRandomID)
      vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.Rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("VenderID").Text) <> "" Then
           ssql = "Select PaymentID From PaymentVender where PaymentID=" & Val(TxtPaymentID.Text) & " and VenderID = " & Val(Grid.Columns("VenderID").Text)
            With cn.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmPaymentVender, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtPaymentID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPaymentDate.Date), "Cleared Code-" & Grid.Columns("VenderID").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & " Vendor Name-" & Trim(Grid.Columns("VenderName").Text) & " Discription-" & Trim(Grid.Columns("Description").Text))
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmPaymentVender, eClearSavedRecord, TxtPaymentID.Text, DtpPaymentDate.DateValue, vGridRows & " Payment Invoice/s Cleared")
      Grid.Redraw = True
  ''''''''''''''''''
'    cn.Execute ("Insert Into UserActivities values ('Payment Vendor Wise'" & "," & TxtPaymentID.Text & ",'" & DtpPaymentDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
'''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   cn.Execute ("Insert Into UserActivities values ('Payment Vendor Wise'" & "," & TxtPaymentID.Text & ",'" & DtpPaymentDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Unload Me
End Sub

Private Sub BtnDelete_Click()
On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniPaymentVender", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
'      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   cn.BeginTrans
   Call BinData
   Call ActivityLogBin("", eFrmPaymentVender, eDelete, TxtPaymentID.Text, DtpPaymentDate.DateValue, Grid.Rows - 1 & " Payment Customer/s Deleted ")
'   vMaxBinID = FunGetMaxBinID
   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_PaymentHeader Select " & vMaxBinID & ",'" & Date & "',* from PaymentHeader Where PaymentID = " & TxtPaymentID.Text & " And PaymentDate ='" & DtpPaymentDate.DateValue & "'")
    '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_PaymentVender Select " & vMaxBinID & ",'" & Date & "', * from PaymentVender Where PaymentID = " & TxtPaymentID.Text)
    
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Payment Vendor Wise'" & "," & TxtPaymentID.Text & ",'" & DtpPaymentDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
   Grid.Redraw = False
   Grid.RemoveAll
   cn.Execute "Delete from PaymentVender where PaymentID = " & Val(TxtPaymentID.Text)
   Grid.Redraw = True
   cn.Execute "Delete from PaymentHeader where PaymentID = " & Val(TxtPaymentID.Text)
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchPaymentVender.Show vbModal
   If SchPaymentVender.ParaOutPaymentID <> 0 Then
      TxtPaymentID.Text = SchPaymentVender.ParaOutPaymentID
      cn.Execute ("Insert Into UserActivities values ('Payment Vendor Wise'" & "," & TxtPaymentID.Text & ",'" & DtpPaymentDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetPayment
   End If
End Sub

Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
      TxtAmount.SetFocus
   Else
      TxtVenderID.SetFocus
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
      vStrSQL = " Select H.PaymentID, H.PaymentDate, SM.EmpName, Pty.PartyName, B.VenderID, b.PreviousPayable," & vbCrLf _
      + " Amount, Isnull(B.Discount,0) Discount, B.[Description]" & vbCrLf _
      + " from PaymentHeader H " & vbCrLf _
      + " Inner join PaymentVender B on H.PaymentID = B.PaymentID " & vbCrLf _
      + " Left outer Join Parties pty on pty.partyID = b.VenderID" & vbCrLf _
      + " Left outer Join Employees SM on SM.EmpID = H.EmpID" & vbCrLf _
      + " where h.PaymentID=" & Val(TxtPaymentID.Text)
      
    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
  
If cmbPrintType.Text = "Half Page" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpPaymentVenderHalf1.rpt")
      RptReportViewer.Report.TopMargin = ObjRegistry.Y
      RptReportViewer.Report.LeftMargin = ObjRegistry.x
      RptReportViewer.Report.RightMargin = 225
   ElseIf cmbPrintType.Text = "Thermal" Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpPaymentVenderAurora.rpt")
      RptReportViewer.Report.TopMargin = 0
      RptReportViewer.Report.LeftMargin = 0
      RptReportViewer.Report.RightMargin = 0
   Else
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpPaymentVender.rpt")
   
   End If
    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue CBool(ObjRegistry.PreviousBalanceVisible)
   
   Dim vPrinter() As String
   vPrinter = Split(CmbPrinters.Text, ",")
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
     
   cn.Execute ("Insert Into UserActivities values ('Payment Vendor Wise'" & "," & TxtPaymentID.Text & ",'" & DtpPaymentDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
   
   If ChkIsPreview.Value = 1 Then
      RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniPaymentVender", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
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
'   If Trim(TxtVenderID.Text) = "" Then
'      MsgBox "Enter Employee ID.", vbExclamation, Me.Caption
'      TxtVenderID.SetFocus
'      Exit Sub
'   End If
   If vIsNewRecord Then
      If cn.Execute("Select * from PaymentHeader where PaymentID = " & Val(TxtPaymentID.Text)).RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtPaymentID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
   
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where touserno = " & vUser & " and Entrydate <='" & Date & "'"
    With cn.Execute(vStrSQL)
        If .Fields(0).Value >= DtpPaymentDate.DateValue Then
            MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
    
  RsBody.Filter = 0
  If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one entry for Payment", vbExclamation, "Alert"
      'If TxtPaymentID.Visible And TxtPaymentID.Enabled Then TxtPaymentID.SetFocus
      If DtpPaymentDate.Visible And DtpPaymentDate.Enabled Then DtpPaymentDate.SetFocus
      Exit Sub
  End If
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
  
   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'Payment Vendor' and LocalComputerName = '" & LocalComputerName & "'"
   If cn.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview ) Values ('" & LocalComputerName & "', 'Payment Vendor','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & " Where FormType = 'Payment Vendor' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   cn.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   cn.BeginTrans
   
  Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmPaymentVender, eEdit, TxtPaymentID.Text, DtpPaymentDate.DateValue, "")
   
'   Call UserActivities
   
   ssql = "select * from PaymentHeader where PaymentID =" & Val(TxtPaymentID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, cn, adOpenDynamic, adLockPessimistic
      If .BOF Then
         .AddNew
         !PaymentID = Val(TxtPaymentID.Text)
         !UserNo = vUser
      End If
     !PaymentDate = DtpPaymentDate.DateValue
     !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
     !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
     !BankID = IIf(Trim(TxtBankID.Text) = "", Null, TxtBankID.Text)
'     !UserNo = vUser
     .Update
     .Close
   End With
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !PaymentID = Val(TxtPaymentID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmPaymentVender, eAdd, TxtPaymentID.Text, DtpPaymentDate.DateValue, Grid.Rows - 1 & " New Payment Invoice/s Added ")
   
   cn.CommitTrans
   If MsgBox("Do you want to print this Voucher", vbQuestion + vbYesNo, "Alert") = vbYes Then
      Call BtnPrint_Click
   End If
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
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
      If TxtVenderID.Enabled Then TxtVenderID.SetFocus: Call SubClearDetailArea
   ElseIf Shift = vbCtrlMask Then
      If ActiveControl.Name = Grid.Name Then
         If KeyCode = vbKeyDelete Then
            If Trim(Grid.Columns("VenderID").Text <> "") Then Call mniRemoveRow_Click
            KeyCode = 0
         Else
            KeyCode = 0: Exit Sub
         End If
      End If
      Select Case KeyCode
         Case vbKeyS
            If BtnSave.Enabled Then BtnSave_Click
            KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
               KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
         Case vbKeyO
            If BtnOpen.Enabled Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, True) = True Then If TxtOrganizationID.Enabled And TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus Else TxtBankID.SetFocus Else TxtEmployeeID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtBankID.SetFocus Else TxtOrganizationID.SetFocus
         Case TxtBankID.Name: If FunSelectBank(ssFunctionKey, False) = True Then If TxtVenderID.Enabled And TxtVenderID.Visible Then TxtVenderID.SetFocus Else TxtBankID.SetFocus
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, False) = True Then TxtAmount.SetFocus Else TxtVenderID.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtPaymentID.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtPaymentID.Name Then
         KeyCode = 0
      End If
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
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Payment (Vender Wise)"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   vNoofPrints = IIf(IsNull(ObjRegistry.NoofPrints) Or Val(ObjRegistry.NoofPrints) = 0, 1, ObjRegistry.NoofPrints)
   
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
   ssql = "select * from FormDefaultSetting Where FormType = 'Payment Vendor' and LocalComputerName = '" & LocalComputerName & "'"
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
   
   HelpLocation Me
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
 Unload Me
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select isnull(max(PaymentID),0)+1 from PaymentHeader").Fields(0)
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
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("VenderID").Text = " "
   Grid.Update
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
    Set FrmPaymentVenderWise = Nothing
   End If
   '''''''''''''''''' ActivityLogBin For Close Action
'      Call DeleteTempActivityLogBin(vRandomID)
      If Grid.Rows > 1 And Cancel = 0 Then
         vGridRows = 0
         Grid.Redraw = False
         Grid.MoveFirst
         For vCounter = 2 To Grid.Rows
            vGridRows = vGridRows + 1
            If Trim(Grid.Columns("VenderID").Text) <> "" Then
               ssql = "Select PaymentID From PaymentVender where PaymentID=" & Val(TxtPaymentID.Text) & " and VenderID = " & Val(Grid.Columns("VenderID").Text)
               With cn.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmPaymentVender, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtPaymentID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPaymentDate.Date), "Closed Code-" & Grid.Columns("VenderID").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & " Vendor Name-" & Trim(Grid.Columns("VenderName").Text) & " Discription-" & Trim(Grid.Columns("Description").Text))
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmPaymentVender, eCloseSavedRecord, TxtPaymentID.Text, DtpPaymentDate.DateValue, vGridRows & " Payment Invoice/s Closed")
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
   'TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("FinalCredit").Value
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
   TxtVenderID.Enabled = False
   BtnVender.Enabled = False
   'TxtPaymentID.BackColor = TxtVenderName.BackColor
   'TxtPaymentID.TabStop = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("VenderID").Text) = "" Then
      TxtVenderID.Text = ""
      TxtVenderID.Enabled = True
      BtnVender.Enabled = True
      TxtVenderID.SetFocus
   Else
      TxtVenderID.Enabled = False
      BtnVender.Enabled = False
      TxtAmount.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Val(Grid.Columns("VenderID").Text) = 0 Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Val(Grid.Columns("VenderID").Text) = 0 Then Exit Sub
   ssql = "Select PaymentID From PaymentVensder where PaymentID=" & Val(TxtPaymentID.Text) & " and VenderID = " & Val(Grid.Columns("VenderID").Text)
   With cn.Execute(ssql)
      If .EOF Then
         Call ActivityLogBin("", eFrmPaymentVender, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtPaymentID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPaymentDate.Date), "Removed Code-" & Grid.Columns("VenderID").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & " Vendor Name-" & Trim(Grid.Columns("VenderName").Text) & " Discription-" & Trim(Grid.Columns("Description").Text))
      Else
         Call ActivityLogBin("", eFrmPaymentVender, eRemoveRow, TxtPaymentID.Text, DtpPaymentDate.DateValue, "Removed Code-" & Grid.Columns("VenderID").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & " Vendor Name-" & Trim(Grid.Columns("VenderName").Text) & " Discription-" & Trim(Grid.Columns("Description").Text))
         Call ActivityLogBin(vRandomID, eFrmPaymentVender, eAddTempRecord, TxtPaymentID.Text, DtpPaymentDate.DateValue, "Pending Remove Code-" & Grid.Columns("VenderID").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & " Vendor Name-" & Trim(Grid.Columns("VenderName").Text) & " Discription-" & Trim(Grid.Columns("Description").Text))
      End If
   End With
   RsBody.Filter = "VenderID = " & Val(TxtVenderID.Text)
   If RsBody.RecordCount > 0 Then RsBody.Delete
   cn.Execute ("Insert Into UserActivities values ('Payment Vendor Wise'" & "," & TxtPaymentID.Text & ",'" & DtpPaymentDate.DateValue & "','Removed VendorID-" & Grid.Columns("VenderID").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & " Discount-" & Grid.Columns("Discount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   Grid.Refresh
   RsBody.Filter = 0
   GetDataBackFromGridToTexBoxes
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
      Call SubClearFields
      vRandomID = Rnd() * 11111 & " " & Format(Now, "dd/mm hh:mm:ss")
      Call PopulateDataToGrid
      BtnPrint.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      'TxtPaymentID.Enabled = True
      TxtVenderID.Enabled = True
      BtnVender.Enabled = True
      TxtPaymentID.Text = FunGetMaxID()
      DtpPaymentDate.Enabled = True
      If DtpPaymentDate.Enabled And DtpPaymentDate.Visible Then DtpPaymentDate.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      'TxtPaymentID.Enabled = False
      TxtVenderID.Enabled = True
      BtnVender.Enabled = True
      BtnPrint.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
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

Private Sub GetPayment()
   On Error GoTo ErrorHandler
   
   ssql = " select h.*, Empname, OrganizationName, AccountName as BankName FROM PaymentHeader h inner join PaymentVender b on h.Paymentid = b.Paymentid left outer join Organizations o on o.OrganizationID = h.organizationid left outer join Employees sm on sm.Empid = h.Empid left outer join ChartofAccounts ca on ca.AccountNo = h.BankID where h.PaymentID=" & Val(TxtPaymentID.Text)
   With cn.Execute(ssql)
      If Not .BOF Then
          DtpPaymentDate.DateValue = !PaymentDate
          TxtEmployeeID.Text = IIf(IsNull(!EmpID) = True, "", !EmpID)
          TxtEmployeeName.Text = IIf(IsNull(!EmpName) = True, "", !EmpName)
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID) = True, "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName) = True, "", !OrganizationName)
          TxtBankID.Text = IIf(IsNull(!BankID) = True, "", !BankID)
          TxtBankName.Text = IIf(IsNull(!BankName) = True, "", !BankName)
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
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from PaymentVender where PaymentID = " & Val(TxtPaymentID.Text), cn, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select b.*, PartyName + isnull(' (' + p.Address + ')','')  as VenderName From PaymentVender b inner join parties p on b.Venderid = p.partyid where PaymentID = " & Val(TxtPaymentID.Text)
      With cn.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            Grid.AddNew
            Grid.Columns("VenderID").Text = !VenderID
            Grid.Columns("VenderName").Text = IIf(IsNull(!VenderName), "", !VenderName)
            Grid.Columns("PreviousPayable").Value = !PreviousPayable
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("Description").Text = IIf(IsNull(!Description), "", !Description)
            Grid.Columns("Discount").Value = IIf(IsNull(!Discount), "", !Discount)
            Grid.Columns("FinalCredit").Value = Round(Val(Grid.Columns("PreviousPayable").Value) - Val(Grid.Columns("Amount").Value) - Val(Grid.Columns("Discount").Value), 2)
            'TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(!FinalCredit)
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("VenderID").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtVenderID.Text = .Columns("VenderID").Text
      TxtVenderName.Text = .Columns("VenderName").Text
      TxtPreviousPayable.Text = .Columns("PreviousPayable").Value
      TxtAmount.Text = .Columns("Amount").Value
      TxtDescription.Text = .Columns("Description").Text
      TxtDiscount.Text = .Columns("Discount").Value
      TxtFinalCredit.Text = .Columns("FinalCredit").Value
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtVenderID.Enabled = True
   BtnVender.Enabled = True
   TxtVenderID.Text = ""
   TxtVenderName.Text = ""
   TxtPreviousPayable.Text = ""
   TxtAmount.Text = ""
   TxtDescription.Text = ""
   TxtDiscount.Text = ""
   TxtFinalCredit.Text = ""
End Sub

Private Function FunSelectEmployee(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchEmployee.Show vbModal, Me
        If SchEmployee.ParaOutEmployeeID = "" Then FunSelectEmployee = False: Exit Function
        TxtEmployeeID.Text = SchEmployee.ParaOutEmployeeID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Employees where isLockEmployee = 0 and EmpID = " & Val(TxtEmployeeID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtEmployeeName.Text = !EmpName
          FunSelectEmployee = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectEmployee = False
          .Close
          TxtEmployeeID.Text = ""
          TxtEmployeeName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub BtnEmployee_Click()
   If FunSelectEmployee(ssButton, False) = True Then
      If TxtOrganizationID.Enabled And TxtOrganizationID.Visible Then TxtOrganizationID.SetFocus Else TxtBankID.SetFocus
   Else
      TxtEmployeeID.SetFocus
   End If
End Sub

Private Sub TxtAmount_Change()
   Call CalculateBody
End Sub

Private Sub TxtVenderID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtDiscount_Change()
   Call CalculateBody
End Sub

Private Sub txtDiscount_LostFocus()
Select Case ActiveControl.Name
   Case TxtVenderID.Name, TxtVenderName.Name, TxtVenderID.Name, TxtAmount.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtVenderID_Change()
   If ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   If TxtVenderName.Text <> "" Then
      TxtVenderName.Text = ""
      TxtPreviousPayable.Text = ""
   End If
End Sub

Private Sub TxtVenderID_Validate(Cancel As Boolean)
   If TxtVenderName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtVenderID.Text) = "" Then Exit Sub
   vTemp = FunSelectVender(ssValidate, False)
   If vTemp = False Then
      vTemp = FunSelectVender(ssButton, False)
      Cancel = False
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtEmployeeID_Change()
   If TxtEmployeeID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtEmployeeID.Name Then Exit Sub
   If TxtEmployeeName.Text <> "" Then
      TxtEmployeeName.Text = ""
   End If
End Sub

Private Sub TxtEmployeeID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtEmployeeID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtEmployeeName.Text <> "" Then Exit Sub
   If Trim(TxtEmployeeID.Text) = "" Then Exit Sub
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

Private Sub GetDataFromTexBoxesToGrid()
   Dim vrowcounter As Integer
   If Trim(TxtVenderID.Text) = "" Then
      If TxtVenderID.Enabled = True Then TxtVenderID.SetFocus
      Exit Sub
   End If
'   If Val(TxtAmount.Text) = 0 Then
'      TxtAmount.SetFocus
'      Exit Sub
'   End If
'   If Val(TxtFinalCredit.Text) < 0 Then
'      TxtAmount.SetFocus
'      Exit Sub
'   End If
On Error GoTo ErrorHandler
      RsBody.Filter = "VenderID = " & Val(TxtVenderID.Text)
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("VenderID").Text = TxtVenderID.Text
          RsBody!VenderID = Val(TxtVenderID.Text)
          If vIsNewRecord = False Then Call ActivityLogBin("", eFrmPaymentVender, eAddNewRowByEdit, TxtPaymentID.Text, DtpPaymentDate.DateValue, "Add New Code-" & TxtVenderID.Text & " Amount-" & " " & TxtAmount.Text & " Vendor Name-" & Trim(TxtVenderName.Text) & " Description-" & Trim(TxtDescription.Text))
          Call ActivityLogBin(vRandomID, eFrmPaymentVender, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtPaymentID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPaymentDate.Date), "Pending Add New Code-" & TxtVenderID.Text & " Amount-" & TxtAmount.Text & " Discount-" & TxtDiscount.Text & " Vendor Name-" & Trim(TxtVenderName.Text) & " Description-" & Trim(TxtDescription.Text))
       Else
         ssql = "Select PaymentID From PaymentVender where PaymentID=" & Val(TxtPaymentID.Text) & " and VenderID = " & Val(Grid.Columns("VenderID").Text)
         With cn.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmPaymentVender, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtPaymentID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPaymentDate.Date), "Effected Code-" & Grid.Columns("VenderID").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & " Vendor Name-" & Trim(Grid.Columns("VenderName").Text) & " Discription-" & Trim(Grid.Columns("Description").Text))
               Call ActivityLogBin("", eFrmPaymentVender, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtPaymentID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPaymentDate.Date), "Updated Code-" & TxtVenderID.Text & " Amount-" & TxtAmount.Text & " Discount-" & TxtDiscount.Text & " Vendor Name-" & Trim(TxtVenderName.Text) & " Description-" & Trim(TxtDescription.Text))
            Else
               Call ActivityLogBin("", eFrmPaymentVender, eEdit, TxtPaymentID.Text, DtpPaymentDate.Date, "Effected Code-" & Grid.Columns("VenderID").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & " Vendor Name-" & Trim(Grid.Columns("VenderName").Text) & " Discription-" & Trim(Grid.Columns("Description").Text))
               Call ActivityLogBin("", eFrmPaymentVender, eEdit, TxtPaymentID.Text, DtpPaymentDate.Date, "Updated Code-" & TxtVenderID.Text & " Amount-" & TxtAmount.Text & " Discount-" & TxtDiscount.Text & " Vendor Name-" & Trim(TxtVenderName.Text) & " Description-" & Trim(TxtDescription.Text))
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmPaymentVender, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtPaymentID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpPaymentDate.Date), "Pending Update Code-" & TxtVenderID.Text & " Amount-" & TxtAmount.Text & " Discount-" & TxtDiscount.Text & " Vendor Name-" & Trim(TxtVenderName.Text) & " Description-" & Trim(TxtDescription.Text))
       End If
       With Grid
         .Columns("VenderName").Text = TxtVenderName.Text
         .Columns("PreviousPayable").Value = Val(TxtPreviousPayable.Text)
         .Columns("Amount").Value = Val(TxtAmount.Text)
         .Columns("Description").Text = TxtDescription.Text
         .Columns("Discount").Value = IIf(Val(TxtDiscount.Text) = 0, 0, Val(TxtDiscount.Text))
         .Columns("FinalCredit").Value = Val(TxtFinalCredit.Text)
         
         RsBody!PreviousPayable = IIf(Val(TxtPreviousPayable.Text) = 0, 0, Val(TxtPreviousPayable.Text))
         RsBody!Amount = IIf(Val(TxtAmount.Text) = 0, 0, Val(TxtAmount.Text))
         RsBody!Description = IIf(Trim(TxtDescription.Text) = "", Null, TxtDescription.Text)
         RsBody!Discount = IIf(Val(TxtDiscount.Text) = 0, 0, Val(TxtDiscount.Text))
         RsBody.Update
         .MoveLast
         If Trim(.Columns("VenderID").Text) <> "" Then
            .AllowAddNew = True
            .AddNew
            .Columns("VenderID").Text = " "
            .AllowAddNew = False
         End If
      End With
   Call SubClearDetailArea
   TxtVenderID.SetFocus
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Function FunSelectVender(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchVender.Show vbModal, Me
        If SchVender.ParaOutVenderID = "" Then FunSelectVender = False: Exit Function
        TxtVenderID.Text = SchVender.ParaOutVenderID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Parties where PartyID = " & Val(TxtVenderID.Text) & " AND PartyType <> 'C' and isLockParty = 0"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtVenderName.Text = !PartyName & IIf(IsNull(!Address), "", " (" & !Address & " )")
          TxtPreviousPayable.Text = -1 * cn.Execute("SELECT isnull(dbo.FunCurrentDebit(" & Val(TxtVenderID.Text) & ",'" & DtpPaymentDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0) ").Fields(0).Value
          FunSelectVender = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectVender = False
          .Close
          TxtVenderID.Text = ""
          TxtVenderName.Text = ""
          TxtVenderID.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpPaymentDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_PaymentHeader").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
   If vIsNewRecord = False Then
      With cn.Execute("Select  * from PaymentHeader where PaymentID = " & TxtPaymentID.Text & " And PaymentDate = '" & DtpPaymentDate.DateValue & "'")
        If Val(TxtEmployeeID.Text) <> IIf(IsNull(!EmpID), 0, !EmpID) Then
            cn.Execute ("Insert Into UserActivities values ('Payment Vendor Wise'" & "," & TxtPaymentID.Text & ",'" & DtpPaymentDate.DateValue & "','Updated EmpID-" & !EmpID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
      End With
      Grid.MoveFirst
    
      For i = 1 To Grid.Rows - 1
        With cn.Execute("Select * from PaymentVender Where PaymentID = " & TxtPaymentID.Text & " And VenderID = " & Val(Grid.Columns("VenderID").Text))
             If .EOF = True Then
                cn.Execute ("Insert Into UserActivities values ('Payment Vendor Wise'" & "," & TxtPaymentID.Text & ",'" & DtpPaymentDate.DateValue & "','Inserted New VenderID-" & Grid.Columns("VenderID").Text & " PreviousPayable-" & Grid.Columns("PreviousPayable").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & " Discount-" & Grid.Columns("Discount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("Amount").Text <> !Amount Or Grid.Columns("Discount").Text <> !Discount Then
                   cn.Execute ("Insert Into UserActivities values ('Payment Vendor Wise'" & "," & TxtPaymentID.Text & ",'" & DtpPaymentDate.DateValue & "','Updated VenderID-" & Grid.Columns("VenderID").Text & " PreviousPayable-" & Grid.Columns("PreviousPayable").Text & " Amount-" & Grid.Columns("Amount").Text & " Discount-" & Grid.Columns("Discount").Text & " Discount-" & Grid.Columns("Discount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
         End With
         Grid.MoveNext
      Next
   Else
      cn.Execute ("Insert Into UserActivities values ('Payment Vendor Wise'" & "," & TxtPaymentID.Text & ",'" & DtpPaymentDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Function FunSelectBank(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchAccounts.ParaInDetail = ""
      SchAccounts.ParaInWhereClause = " and c.accountno like '1%'"
      'SchAccounts.cmbfilter.Text = "Banks"
      'SchAccounts.cmbfilter.Enabled = False
      SchAccounts.Show vbModal, Me
      If SchAccounts.ParaOutAccountNo = "" Then FunSelectBank = False: Exit Function
      TxtBankID.Text = SchAccounts.ParaOutAccountNo
   End If
   Dim vStrSQL As String
   vStrSQL = "select * from ChartofAccounts where AccountNo = " & Val(TxtBankID.Text) & " and accountno like '1%'"
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtBankName.Text = !AccountName
         .Close
         FunSelectBank = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      Else
         FunSelectBank = False
         .Close
         TxtBankName.Text = ""
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnBank_Click()
   If FunSelectBank(ssButton, False) = True Then
      If TxtVenderID.Enabled And TxtVenderID.Visible Then TxtVenderID.SetFocus
   Else
      TxtBankID.SetFocus
   End If
End Sub

Private Sub TxtBankID_Change()
   If ActiveControl.Name <> TxtBankID.Name Then Exit Sub
   If TxtBankName.Text <> "" Then TxtBankName.Text = ""
End Sub

Private Sub TxtBankID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtBankID.Name Then Exit Sub
   If TxtBankID.Text = "" Then Exit Sub
   If TxtBankName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBank(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBank(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.PaymentHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmPaymentVender) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmPaymentVender & ", " & vUser & "," & TableHeaderFields(eFrmPaymentVender) & " from PaymentHeader " & vbCrLf _
             & "Where PaymentID = " & TxtPaymentID.Text & " and PaymentDate = '" & DtpPaymentDate.DateValue & "'"
      cn.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.PaymentVenderBin (" & TableBodyFields(eFrmPaymentVender) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmPaymentVender) & " from PaymentVender " & vbCrLf _
             & "Where PaymentID = " & TxtPaymentID.Text
      cn.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


