VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmRecoveryCustomerWise 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmRecoveryCustomerWise.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmRecoveryCustomerWise.frx":0ECA
      Left            =   2655
      List            =   "FrmRecoveryCustomerWise.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Tag             =   "1"
      Top             =   8235
      Width           =   3276
   End
   Begin VB.ComboBox cmbPrintType 
      Height          =   315
      Left            =   5940
      TabIndex        =   59
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   8235
      Width           =   1170
   End
   Begin VB.CheckBox ChkIsPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Preview"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7155
      TabIndex        =   58
      Top             =   8280
      Width           =   1290
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3285
      Width           =   1800
   End
   Begin VB.TextBox TxtCommision 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   1553
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox TxtNetAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      CausesValidation=   0   'False
      Height          =   330
      Left            =   12600
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   8303
      Width           =   1035
   End
   Begin VB.TextBox TxtTotalAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      CausesValidation=   0   'False
      Height          =   330
      Left            =   10200
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   8303
      Width           =   1035
   End
   Begin VB.TextBox TxtTotalDiscount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      CausesValidation=   0   'False
      Height          =   330
      Left            =   11400
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   8303
      Width           =   1050
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
      Left            =   13920
      TabIndex        =   35
      Top             =   6360
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
         TabIndex        =   36
         Tag             =   "NC"
         Text            =   "FrmRecoveryCustomerWise.frx":0ECE
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
         TabIndex        =   37
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnEmployee 
      Height          =   330
      Left            =   5265
      TabIndex        =   33
      Top             =   2378
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
      MICON           =   "FrmRecoveryCustomerWise.frx":0FC1
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtRecoveryID 
      Height          =   315
      Left            =   1785
      TabIndex        =   0
      Top             =   2378
      Width           =   1050
      _ExtentX        =   1852
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
      Left            =   8992
      TabIndex        =   16
      Top             =   9158
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
      MICON           =   "FrmRecoveryCustomerWise.frx":0FDD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7687
      TabIndex        =   12
      Top             =   9158
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
      MICON           =   "FrmRecoveryCustomerWise.frx":0FF9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5077
      TabIndex        =   14
      Top             =   9158
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
      MICON           =   "FrmRecoveryCustomerWise.frx":1015
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10297
      TabIndex        =   17
      Top             =   9158
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
      MICON           =   "FrmRecoveryCustomerWise.frx":1031
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6382
      TabIndex        =   13
      Top             =   9158
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
      MICON           =   "FrmRecoveryCustomerWise.frx":104D
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   315
      Left            =   825
      TabIndex        =   6
      Top             =   3285
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton BtnCustomer 
      Height          =   330
      Left            =   1815
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3285
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
      MICON           =   "FrmRecoveryCustomerWise.frx":1069
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   2175
      TabIndex        =   19
      Top             =   3285
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   5
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4320
      Left            =   -60
      TabIndex        =   24
      Top             =   3600
      Width           =   15315
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   11
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
      stylesets(0).Picture=   "FrmRecoveryCustomerWise.frx":1085
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
      Columns.Count   =   11
      Columns(0).Width=   1561
      Columns(0).Caption=   "Serial"
      Columns(0).Name =   "Serial"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2381
      Columns(1).Caption=   "Customer ID"
      Columns(1).Name =   "CustomerID"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   6694
      Columns(2).Caption=   "Customer Name"
      Columns(2).Name =   "CustomerName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "CompanyName"
      Columns(3).Name =   "CompanyName"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2699
      Columns(4).Caption=   "Previous Receivable"
      Columns(4).Name =   "PreviousReceivable"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1667
      Columns(5).Caption=   "Amount"
      Columns(5).Name =   "Amount"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1588
      Columns(6).Caption=   "ManualNo"
      Columns(6).Name =   "ManualNo"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3440
      Columns(7).Caption=   "Description"
      Columns(7).Name =   "Description"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1455
      Columns(8).Caption=   "Discount"
      Columns(8).Name =   "Discount"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1614
      Columns(9).Caption=   "Remaining Amt"
      Columns(9).Name =   "FinalCredit"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "CompanyID"
      Columns(10).Name=   "CompanyID"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   27014
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpRecoveryDate 
      Height          =   315
      Left            =   2925
      TabIndex        =   1
      Top             =   2378
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
   Begin SITextBox.Txt TxtPreviousReceivable 
      Height          =   315
      Left            =   7800
      TabIndex        =   20
      Top             =   3285
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   9315
      TabIndex        =   8
      Top             =   3285
      Width           =   960
      _ExtentX        =   1693
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
      Left            =   13140
      TabIndex        =   11
      Top             =   3285
      Width           =   825
      _ExtentX        =   1455
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
      Left            =   13980
      TabIndex        =   21
      Top             =   3285
      Width           =   1215
      _ExtentX        =   2143
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
   Begin SITextBox.Txt TxtEmployeeName 
      Height          =   315
      Left            =   5625
      TabIndex        =   5
      Top             =   2378
      Width           =   1680
      _ExtentX        =   2963
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
      Left            =   4305
      TabIndex        =   2
      Top             =   2378
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
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3772
      TabIndex        =   15
      Top             =   9158
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
      MICON           =   "FrmRecoveryCustomerWise.frx":10A1
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDescription 
      Height          =   315
      Left            =   11205
      TabIndex        =   10
      Top             =   3285
      Width           =   1935
      _ExtentX        =   3413
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
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   7350
      TabIndex        =   3
      Tag             =   "NC"
      Top             =   2378
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
      Left            =   8415
      TabIndex        =   46
      Tag             =   "NC"
      Top             =   2378
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
      Left            =   8055
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2378
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
      MICON           =   "FrmRecoveryCustomerWise.frx":10BD
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBankMachineID 
      Height          =   315
      Left            =   10320
      TabIndex        =   4
      Top             =   2378
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
   Begin SITextBox.Txt TxtBankMachineName 
      Height          =   315
      Left            =   11385
      TabIndex        =   51
      Top             =   2378
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
   Begin JeweledBut.JeweledButton BtnBankMachine 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   11025
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2378
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
      MICON           =   "FrmRecoveryCustomerWise.frx":10D9
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtManualNo 
      Height          =   315
      Left            =   10260
      TabIndex        =   9
      Top             =   3285
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      Alignment       =   1
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
      Left            =   1980
      TabIndex        =   61
      Top             =   8235
      Width           =   570
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No."
      Height          =   225
      Left            =   3810
      TabIndex        =   57
      Top             =   8340
      Width           =   1095
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      Height          =   195
      Left            =   6000
      TabIndex        =   56
      Top             =   3090
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual No"
      Height          =   195
      Left            =   10230
      TabIndex        =   55
      Top             =   3090
      Width           =   780
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   9330
      TabIndex        =   54
      Top             =   3090
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Machine ID"
      Height          =   195
      Left            =   10320
      TabIndex        =   53
      Top             =   2138
      Width           =   1245
   End
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   7350
      TabIndex        =   49
      Top             =   2138
      Width           =   1095
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   8535
      TabIndex        =   48
      Top             =   2138
      Width           =   1350
   End
   Begin VB.Label LblWords 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   12390
      TabIndex        =   45
      Top             =   8783
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      Height          =   195
      Left            =   12600
      TabIndex        =   44
      Top             =   8063
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   10200
      TabIndex        =   42
      Top             =   8063
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Discount"
      Height          =   195
      Left            =   11400
      TabIndex        =   41
      Top             =   8063
      Width           =   1035
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
      Left            =   11205
      TabIndex        =   38
      Top             =   630
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   11175
      TabIndex        =   34
      Top             =   3090
      Width           =   795
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      Height          =   375
      Left            =   5610
      TabIndex        =   32
      Top             =   2183
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      Height          =   255
      Left            =   4305
      TabIndex        =   31
      Top             =   2183
      Width           =   1215
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recovery (Customer Wise)"
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
      Left            =   2700
      TabIndex        =   30
      Top             =   270
      Width           =   4680
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   13245
      TabIndex        =   29
      Top             =   3090
      Width           =   630
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remainig Amount"
      Height          =   195
      Left            =   13965
      TabIndex        =   28
      Top             =   3090
      Width           =   1245
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Receivable"
      Height          =   195
      Left            =   7785
      TabIndex        =   27
      Top             =   3090
      Width           =   1470
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerID"
      Height          =   195
      Left            =   870
      TabIndex        =   26
      Top             =   3090
      Width           =   825
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   195
      Left            =   2325
      TabIndex        =   25
      Top             =   3090
      Width           =   1125
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
      Caption         =   "Recovery Date"
      Height          =   195
      Left            =   2940
      TabIndex        =   23
      Top             =   2183
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Recovery ID"
      Height          =   195
      Left            =   1785
      TabIndex        =   22
      Top             =   2183
      Width           =   900
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmRecoveryCustomerWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter, vGridRows As Integer
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim ssql, vRandomID As String
Dim i As Integer
Dim vStrSQL, vManualNo As String
Dim vMaxBinID As Integer
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
      If TxtBankMachineID.Enabled Then TxtBankMachineID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub CalculateBody()
   TxtFinalCredit.Text = Round(Val(TxtPreviousReceivable.Text) - Val(TxtAmount.Text) - Val(TxtDiscount.Text), 2)
End Sub

Private Sub BtnClear_Click()
On Error GoTo ErrorHandler
     '''''''''''''''''' ActivityLogBin For Clear Action
'      Call DeleteTempActivityLogBin(vRandomID)
      vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("CustomerID").Text) <> "" Then
           ssql = "Select RecoveryID From RecoveryCustomer where RecoveryID=" & Val(TxtRecoveryID.Text) & " and CustomerID = " & Val(Grid.Columns("CustomerID").Text)
            With cn.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmRecoveryCustomerWise, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Cleared Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmRecoveryCustomerWise, eClearSavedRecord, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, vGridRows & " Recovery Customer/s Cleared")
      Grid.Redraw = True
  ''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Recovery Customer Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   cn.Execute ("Insert Into UserActivities values ('Recovery Customer Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Unload Me
End Sub

Private Sub BtnDelete_Click()
On Error GoTo ErrorHandler
    ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniRecoveryCustomer", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
   Call ActivityLogBin("", eFrmRecoveryCustomerWise, eDelete, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, Grid.rows - 1 & " Recovery Customer/s Deleted ")
'    vMaxBinID = FunGetMaxBinID
   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_RecoveryHeader Select " & vMaxBinID & ",'" & Date & "',* from RecoveryHeader Where RecoveryID = " & TxtRecoveryID.Text & " And RecoveryDate ='" & DtpRecoveryDate.DateValue & "'")
    '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_RecoveryCustomer Select " & vMaxBinID & ",'" & Date & "', * from RecoveryCustomer Where RecoveryID = " & TxtRecoveryID.Text)

    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cn.Execute ("Insert Into UserActivities values ('Recovery Customer Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Call ActivityLog("Recovery Customer Wise", eDelete, TxtRecoveryID.Text, DtpRecoveryDate.DateValue)
   Grid.Redraw = False
   Grid.RemoveAll
   cn.Execute "Delete from RecoveryCustomer where RecoveryID = " & Val(TxtRecoveryID.Text)
   Grid.Redraw = True
   cn.Execute "Delete from RecoveryHeader where RecoveryID = " & Val(TxtRecoveryID.Text)
   
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchRecoveryCustomer.Show vbModal
   If SchRecoveryCustomer.ParaOutRecoveryID <> 0 Then
      TxtRecoveryID.Text = SchRecoveryCustomer.ParaOutRecoveryID
      cn.Execute ("Insert Into UserActivities values ('Recovery Customer Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
      GetRecovery
   End If
End Sub

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      CmbCompany.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   
      vStrSQL = " Select H.RecoveryID, H.RecoveryDate, SM.EmpName, AccountName as PartyName, B.CustomerID, b.PreviousReceivable," & vbCrLf _
      + " Amount, Isnull(B.Discount,0) Discount, b.Description, UserName" & vbCrLf _
      + " from RecoveryHeader H " & vbCrLf _
      + " Inner join RecoveryCustomer B on H.RecoveryID = B.RecoveryID " & vbCrLf _
      + " Left outer Join ChartofAccounts ca on ca.AccountNo = b.CustomerID" & vbCrLf _
      + " Left outer Join Employees SM on SM.EmpID = H.EmpID" & vbCrLf _
      + " inner Join Users u on U.UserNo = H.UserNo" & vbCrLf _
      + " where h.recoveryID=" & Val(TxtRecoveryID.Text)
      
'      vStrSQL = " Select H.RecoveryID, H.RecoveryDate, isnull(EmpName,'') as EmpName, AccountName as PartyName, B.CustomerID, b.PreviousReceivable," & vbCrLf _
      + " Amount, Isnull(B.Discount,0) Discount, UserName" & vbCrLf _
      + " from RecoveryHeader H " & vbCrLf _
      + " Inner join RecoveryCustomer B on H.RecoveryID = B.RecoveryID " & vbCrLf _
      + " Left outer Join ChartofAccounts pty on pty.AccountNo = b.CustomerID" & vbCrLf _
      + " Left outer Join Employees SM on SM.EmpID = H.EmpID" & vbCrLf _
      + " inner Join Users u on U.UserNo = H.UserNo" & vbCrLf _
      + " where h.recoveryID=" & Val(TxtRecoveryID.Text)
      
   
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
   
   If InStr(1, Printer.DeviceName, "Canon") > 0 Or InStr(1, Printer.DeviceName, "HP") > 0 Then
      Set RptReportViewer.Report = New CrpRecoveryCustomer
   Else
      Set RptReportViewer.Report = New CrpRecoveryCustomerAurora
   End If
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter "abc", "xyz", "ghi"
   
   If ChkIsPreview.Value = 1 Then
      RptReportViewer.Show vbModal, Me
   Else
      RptReportViewer.Report.PrintOut False
   End If
'   CN.Execute ("Insert Into UserActivities values ('Recovery Customer Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Printed','" & Date & "','" & Time & "',5,'Printed'," & vUser & ")")
'   RptReportViewer.Show vbModal, Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniRecoveryCustomer", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
     MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
     Exit Sub
   End If
   If cn.Execute("Select * From AdminClosing where ToUserNo = " & vUser & " and EntryDate = '" & DtpRecoveryDate.DateValue & "'").RecordCount > 0 Then
      MsgBox "You are not authorized to Add Record in Closing Dates.", vbCritical, "Alert"
      Exit Sub
   End If

'  Header Validation
'   If Trim(TxtCustomerID.Text) = "" Then
'      MsgBox "Enter Employee ID.", vbExclamation, Me.Caption
'      TxtCustomerID.SetFocus
'      Exit Sub
'   End If
   If vIsNewRecord Then
      If cn.Execute("Select * from RecoveryHeader where RecoveryID = " & Val(TxtRecoveryID.Text)).RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtRecoveryID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
  RsBody.Filter = 0
  If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one entry for recovery", vbExclamation, "Alert"
      If DtpRecoveryDate.Visible And DtpRecoveryDate.Enabled Then DtpRecoveryDate.SetFocus
      Exit Sub
  End If
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
  
'   If vIsNewRecord = False Then
'      Call ActivityLog("Recovery Customer Wise", eEdit, TxtRecoveryID.Text, DtpRecoveryDate.DateValue)
'   End If

   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   ssql = "select * from FormDefaultSetting Where FormType = 'Recovery Customer Wise' and LocalComputerName = '" & LocalComputerName & "'"
   If cn.Execute(ssql).EOF Then
      ssql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview ) Values ('" & LocalComputerName & "', 'Recovery Customer Wise','" & cmbPrintType.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & ChkIsPreview.Value & ")"
   Else
      ssql = "Update FormDefaultSetting set Size = '" & cmbPrintType.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & ChkIsPreview.Value & " Where FormType = 'Recovery Customer Wise' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   cn.Execute ssql
   ''''''''''''''''''''''''''''''''''''''''''''
   
   cn.BeginTrans
   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmRecoveryCustomerWise, eEdit, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, "Amount: " & Val(TxtNetAmount.Text))
   
'   Call UserActivities
   
   ssql = "select * from RecoveryHeader where RecoveryID =" & Val(TxtRecoveryID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, cn, adOpenDynamic, adLockPessimistic
      If .BOF Then
         .AddNew
         !RecoveryID = Val(TxtRecoveryID.Text)
         !UserNo = vUser
      End If
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !RecoveryDate = DtpRecoveryDate.DateValue
      !EmpID = IIf(Trim(TxtEmployeeID.Text) = "", Null, TxtEmployeeID.Text)
      !BankMachineID = IIf(Trim(TxtBankMachineID.Text) = "", Null, TxtBankMachineID.Text)
      !Commision = IIf(Trim(TxtCommision.Text) = "", Null, Val(TxtCommision.Text))
'      !UserNo = vUser
      .Update
      .Close
   End With
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !RecoveryID = Val(TxtRecoveryID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmRecoveryCustomerWise, eAdd, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, Grid.rows - 1 & " New Recovery Customer/s Added Amount: " & Val(TxtNetAmount.Text))
'   If vIsNewRecord = True Then Call ActivityLog("Recovery Customer Wise", eAdd, TxtRecoveryID.Text, DtpRecoveryDate.DateValue)
   cn.CommitTrans
   
   If MsgBox("Do you want to print this invoice", vbQuestion + vbYesNo, "Alert") = vbYes Then
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
   If KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      Select Case ActiveControl.Name
      Case TxtCustomerID.Name, TxtAmount.Name, TxtDiscount.Name, TxtDescription.Name
         If TxtCustomerID.Enabled Then TxtCustomerID.SetFocus: Call SubClearDetailArea
      End Select
   ElseIf KeyCode = vbKeyReturn Then
      If ActiveControl.Name = "Grid" Then
         Grid_DblClick
      Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf Shift = vbCtrlMask Then
      If ActiveControl.Name = Grid.Name Then
         If KeyCode = vbKeyDelete Then
            If Trim(Grid.Columns("CustomerID").Text <> "") Then Call mniRemoveRow_Click
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
            If BtnDelete.Enabled And BtnDelete.Visible Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtEmployeeID.Name: If FunSelectEmployee(ssFunctionKey, True) = True Then If TxtOrganizationID.Visible And TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus Else If TxtBankMachineID.Enabled Then TxtBankMachineID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then If TxtBankMachineID.Enabled Then TxtBankMachineID.SetFocus Else TxtOrganizationID.SetFocus
         Case TxtBankMachineID.Name: If FunSelectBankMachine(ssFunctionKey, True) = True Then TxtCustomerID.SetFocus Else TxtBankMachineID.SetFocus
         Case TxtCustomerID.Name: If FunSelectCustomer(ssFunctionKey, False) = True Then CmbCompany.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtRecoveryID.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtRecoveryID.Name Then
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
   SetWindowText Me.hWnd, "Recovery (Customer Wise)"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   HelpLocation Me
   
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
   ssql = "select * from FormDefaultSetting Where FormType = 'Purchase Invoice' and LocalComputerName = '" & LocalComputerName & "'"
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
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   With cn.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
         FunSelectOrganization ssValidate, True
      End If
      .Close
   End With
   
   With cn.Execute("Select * from Companies")
      CmbCompany.AddItem ""
      While Not .EOF
         CmbCompany.AddItem !CompanyName
         CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
         .MoveNext
      Wend
      .Close
   End With

   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus

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
   FunGetMaxID = cn.Execute("Select isnull(max(RecoveryID),0)+1 from RecoveryHeader").Fields(0)
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
   TxtTotalAmount.Text = ""
   TxtTotalDiscount.Text = ""
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("CustomerID").Text = " "
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
    Set FrmRecoveryCustomerWise = Nothing
   End If
   '''''''''''''''''' ActivityLogBin For Close Action
'      Call DeleteTempActivityLogBin(vRandomID)
      If Grid.rows > 1 And Cancel = 0 Then
         vGridRows = 0
         Grid.Redraw = False
         Grid.MoveFirst
         For vCounter = 2 To Grid.rows
            vGridRows = vGridRows + 1
            If Trim(Grid.Columns("CustomerID").Text) <> "" Then
               ssql = "Select RecoveryID From RecoveryCustomer where RecoveryID=" & Val(TxtRecoveryID.Text) & " and CustomerID = " & Val(Grid.Columns("CustomerID").Text)
               With cn.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmRecoveryCustomerWise, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Closed Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmRecoveryCustomerWise, eCloseSavedRecord, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, vGridRows & " Recovery Customer/s Closed")
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
   TxtAmount.Text = Val(TxtAmount.Text) - Grid.Columns("Amount").Value
   TxtTotalDiscount.Text = Val(TxtTotalDiscount.Text) - Grid.Columns("Discount").Value
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
   TxtCustomerID.Enabled = False
   BtnCustomer.Enabled = False
   'TxtRecoveryID.BackColor = TxtCustomerName.BackColor
   'TxtRecoveryID.TabStop = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("CustomerID").Text) = "" Then
      TxtCustomerID.Text = ""
      TxtCustomerID.Enabled = True
      BtnCustomer.Enabled = True
      TxtCustomerID.SetFocus
   Else
      TxtCustomerID.Enabled = False
      BtnCustomer.Enabled = False
      TxtAmount.SetFocus
      If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(Grid.Columns("CustomerID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("CustomerID").Text) = "" Then Exit Sub
   ssql = "Select RecoveryID From RecoveryCustomer where RecoveryID=" & Val(TxtRecoveryID.Text) & " and CustomerID = " & Val(Grid.Columns("CustomerID").Text)
   With cn.Execute(ssql)
      If .EOF Then
         Call ActivityLogBin("", eFrmRecoveryCustomerWise, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Removed Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
      Else
         Call ActivityLogBin("", eFrmRecoveryCustomerWise, eRemoveRow, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, "Removed Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
         Call ActivityLogBin(vRandomID, eFrmRecoveryCustomerWise, eAddTempRecord, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, "Pending Remove Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
      End If
   End With
   RsBody.Filter = "CustomerID = " & Val(TxtCustomerID.Text)
   If RsBody.RecordCount > 0 Then RsBody.Delete
   cn.Execute ("Insert Into UserActivities values ('Recovery Customer Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Removed CustomerID-" & Grid.Columns("CustomerID").Text & " PreviousReceivable-" & Grid.Columns("PreviousReceivable").Text & " Disc-" & Grid.Columns("Discount").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
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
      TxtCustomerID.Enabled = True
      BtnCustomer.Enabled = True
      TxtRecoveryID.Text = FunGetMaxID()
      DtpRecoveryDate.Enabled = True
      If DtpRecoveryDate.Enabled And DtpRecoveryDate.Visible Then DtpRecoveryDate.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      TxtCustomerID.Enabled = True
      BtnCustomer.Enabled = True
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

Private Sub GetRecovery()
   On Error GoTo ErrorHandler
   ssql = " select h.*, OrganizationName, EmpName, BankMachineName FROM RecoveryHeader h inner join RecoveryCustomer b on h.recoveryid = b.recoveryid left outer join Organizations o on o.OrganizationID = h.OrganizationID left outer join Employees sm on sm.Empid = h.Empid left outer join BankMachines bm on bm.BankMachineID = h.BankMachineID where h.RecoveryID=" & Val(TxtRecoveryID.Text)
   With cn.Execute(ssql)
      If Not .BOF Then
          DtpRecoveryDate.DateValue = !RecoveryDate
          TxtEmployeeID.Text = IIf(IsNull(!EmpID) = True, "", !EmpID)
          TxtEmployeeName.Text = IIf(IsNull(!empname) = True, "", !empname)
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          TxtBankMachineID.Text = IIf(IsNull(!BankMachineID), "", !BankMachineID)
          TxtBankMachineName.Text = IIf(IsNull(!BankMachineName), "", !BankMachineName)
          TxtCommision.Text = IIf(IsNull(!Commision), "", !Commision)
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
   RsBody.Open "Select * from RecoveryCustomer where RecoveryID =" & Val(TxtRecoveryID.Text), cn, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select b.*, AccountName + isnull(' (' + p.Address + ')','') as CustomerName From RecoveryCustomer b left outer join ChartofAccounts ca on b.customerid = ca.AccountNo Left Outer join Parties p on ca.AccountNo = p.PartyID where RecoveryID =" & Val(TxtRecoveryID.Text)
      With cn.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = 0
         TxtTotalDiscount.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("Serial").Text = Grid.rows
            Grid.Columns("CustomerID").Text = !CustomerID
            Grid.Columns("CustomerName").Text = IIf(IsNull(!CustomerName), "", !CustomerName)
            If !companyid = 0 Or IsNull(!companyid) Then
               Grid.Columns("CompanyName").Text = ""
            Else
               Grid.Columns("CompanyName").Text = cn.Execute("Select CompanyName from Companies where CompanyID=" & !companyid).Fields(0).Value
            End If
            Grid.Columns("PreviousReceivable").Value = !PreviousReceivable
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("ManualNo").Text = IIf(IsNull(!ManualNo), "", !ManualNo)
            Grid.Columns("Description").Text = IIf(IsNull(!Description), "", !Description)
            Grid.Columns("Discount").Value = IIf(IsNull(!Discount), "", !Discount)
            Grid.Columns("FinalCredit").Value = Round(Val(Grid.Columns("PreviousReceivable").Value) - Val(Grid.Columns("Amount").Value) - Val(Grid.Columns("Discount").Value), 2)
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + !Amount
            TxtTotalDiscount.Text = Val(TxtTotalDiscount.Text) + IIf(IsNull(!Discount), 0, !Discount)
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("CustomerID").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      If Trim(.Columns("CompanyName").Text) = "" Then
         CmbCompany.ListIndex = 0
      Else
         CmbCompany.Text = .Columns("CompanyName").Text
      End If
      TxtCustomerID.Text = .Columns("CustomerID").Text
      TxtCustomerName.Text = .Columns("CustomerName").Text
      TxtPreviousReceivable.Text = .Columns("PreviousReceivable").Value
      TxtAmount.Text = .Columns("Amount").Value
      TxtManualNo.Text = .Columns("ManualNo").Text
      TxtDescription.Text = .Columns("Description").Text
      TxtDiscount.Text = .Columns("Discount").Value
      TxtFinalCredit.Text = .Columns("FinalCredit").Value
   End With
   If Grid.rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtCustomerID.Enabled = True
   BtnCustomer.Enabled = True
   TxtCustomerID.Text = ""
   TxtCustomerName.Text = ""
   CmbCompany.ListIndex = 0
   TxtPreviousReceivable.Text = ""
   TxtAmount.Text = ""
   TxtManualNo.Text = ""
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
    vStrSQL = " Select * FROM Employees where EmpID=" & Val(TxtEmployeeID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtEmployeeName.Text = !empname
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
      If TxtOrganizationID.Visible And TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus Else If TxtBankMachineID.Enabled Then TxtBankMachineID.SetFocus
   Else
      TxtEmployeeID.SetFocus
   End If
End Sub

Private Sub TxtAmount_Change()
   Call CalculateBody
End Sub

Private Sub TxtCustomerID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtDiscount_Change()
   Call CalculateBody
End Sub

Private Sub txtDiscount_LostFocus()
Select Case ActiveControl.Name
   Case TxtCustomerID.Name, TxtCustomerName.Name, TxtCustomerID.Name, TxtAmount.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtCustomerID_Change()
   If ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   If TxtCustomerName.Text <> "" Then
      TxtCustomerName.Text = ""
      TxtPreviousReceivable.Text = ""
   End If
End Sub

Private Sub TxtCustomerID_Validate(Cancel As Boolean)
   If TxtCustomerName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCustomerID.Text) = "" Then Exit Sub
   vTemp = FunSelectCustomer(ssValidate, False)
   If vTemp = False Then
      vTemp = FunSelectCustomer(ssButton, False)
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
   If Trim(TxtCustomerID.Text) = "" Then
      If TxtCustomerID.Enabled = True Then TxtCustomerID.SetFocus
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
    If Trim(Grid.Columns("CustomerID").Text) = "" Then
      RsBody.Filter = "CustomerID = " & Val(TxtCustomerID.Text) & " and Description = " & IIf(Trim(TxtDescription.Text) = "", "Null", "'" & TxtDescription.Text & "'") & " and Amount=" & Val(TxtAmount.Text)
   Else
      RsBody.Filter = "CustomerID = " & Val(Grid.Columns("CustomerID").Text) & " and Description = " & IIf(Grid.Columns("Description").Text = "", "Null", "'" & Grid.Columns("Description").Text & "'") & " and Amount=" & Grid.Columns("Amount").Value
   End If
   
   If TxtCustomerID.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("Serial").Text = Grid.rows
         Grid.Columns("CustomerID").Text = TxtCustomerID.Text
          RsBody!CustomerID = TxtCustomerID.Text
          If vIsNewRecord = False Then Call ActivityLogBin("", eFrmRecoveryCustomerWise, eAddNewRowByEdit, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, "Add New Code-" & TxtCustomerID.Text & " Amount-" & TxtAmount.Text & " Disc " & TxtDiscount.Text & " Name- " & Replace(TxtCustomerName.Text, "'", "''") & " Descritption " & TxtDescription.Text)
         Call ActivityLogBin(vRandomID, eFrmRecoveryCustomerWise, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Pending Add New Code-" & TxtCustomerID.Text & " Amount-" & TxtAmount.Text & " Disc " & TxtDiscount.Text & " Name- " & Replace(TxtCustomerName.Text, "'", "''") & " Descritption " & TxtDescription.Text)
      Else
         Grid.Redraw = False
         Grid.MoveFirst
         For vrowcounter = 1 To Grid.rows
            If Grid.Columns("CustomerID").Text = TxtCustomerID.Text And Grid.Columns("Description").Text = Trim(TxtDescription.Text) And Val(Grid.Columns("Amount").Text) = Val(TxtAmount.Text) Then
            ssql = "Select CustomerID From RecoveryCustomer where RecoveryID=" & Val(TxtRecoveryID.Text) & " and CustomerID = " & Val(Grid.Columns("CustomerID").Text)
                  With cn.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmRecoveryCustomerWise, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Effected Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
                     Else
                        Call ActivityLogBin("", eFrmRecoveryCustomerWise, eEdit, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, "Effected Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
                     End If
                  End With
               TxtAmount.Text = Val(TxtAmount.Text) + Val(Grid.Columns("Amount").Text)
               TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
               TxtTotalDiscount.Text = Val(TxtTotalDiscount.Text) + Val(TxtDiscount.Text) - Val(Grid.Columns("Discount").Text)
               Grid.Columns("CustomerName").Text = TxtCustomerName.Text
               Grid.Columns("CompanyName").Text = CmbCompany.Text
               If CmbCompany.Text <> "" Then Grid.Columns("CompanyID").Value = CmbCompany.ItemData(CmbCompany.ListIndex)
               Grid.Columns("PreviousReceivable").Value = Val(TxtPreviousReceivable.Text)
               Grid.Columns("Amount").Value = Val(TxtAmount.Text)
               Grid.Columns("ManualNo").Text = TxtManualNo.Text
               Grid.Columns("Description").Text = TxtDescription.Text
               Grid.Columns("Discount").Value = IIf(Val(TxtDiscount.Text) = 0, 0, Val(TxtDiscount.Text))
               Grid.Columns("FinalCredit").Value = Val(TxtFinalCredit.Text)
               RsBody!PreviousReceivable = IIf(Val(TxtPreviousReceivable.Text) = 0, 0, Val(TxtPreviousReceivable.Text))
               If CmbCompany.Text <> "" Then RsBody!companyid = CmbCompany.ItemData(CmbCompany.ListIndex)
               RsBody!Amount = IIf(Val(TxtAmount.Text) = 0, 0, Val(TxtAmount.Text))
               RsBody!ManualNo = IIf(Trim(TxtManualNo.Text) = "", Null, TxtManualNo.Text)
               RsBody!Description = IIf(Trim(TxtDescription.Text) = "", Null, TxtDescription.Text)
               RsBody!Discount = IIf(Val(TxtDiscount.Text) = 0, 0, Val(TxtDiscount.Text))
                ssql = "Select CustomerID From RecoveryCustomer where RecoveryID=" & Val(TxtRecoveryID.Text) & " and CustomerID = " & Val(Grid.Columns("CustomerID").Text)
                  With cn.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmRecoveryCustomerWise, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Updated Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
                     Else
                        Call ActivityLogBin("", eFrmRecoveryCustomerWise, eEdit, TxtRecoveryID.Text, DtpRecoveryDate.DateValue, "Updated Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("Description").Text & " Descritption " & Grid.Columns("Description").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmRecoveryCustomerWise, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Pending Update Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
               Call SubClearDetailArea
               Grid.MoveLast
               TxtCustomerID.SetFocus
               Grid.Redraw = True
               Exit Sub
            End If
               Grid.MoveNext
            Next vrowcounter
      End If
   Else
      ssql = "Select RecoveryID From RecoveryCustomer where RecoveryID=" & Val(TxtRecoveryID.Text) & " and CustomerID = " & Val(Grid.Columns("CustomerID").Text)
         With cn.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmRecoveryCustomerWise, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Effected Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
               Call ActivityLogBin("", eFrmRecoveryCustomerWise, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Updated Code-" & TxtCustomerID.Text & " Amount-" & TxtAmount.Text & " Disc " & TxtDiscount.Text & " Name- " & TxtCustomerName.Text & " Descritption " & TxtDescription.Text)
            Else
               Call ActivityLogBin("", eFrmRecoveryCustomerWise, eEdit, TxtRecoveryID.Text, DtpRecoveryDate.Date, "Effected Code-" & Grid.Columns("CustomerID").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " Name- " & Grid.Columns("CustomerName").Text & " Descritption " & Grid.Columns("Description").Text)
               Call ActivityLogBin("", eFrmRecoveryCustomerWise, eEdit, TxtRecoveryID.Text, DtpRecoveryDate.Date, "Updated Code-" & TxtCustomerID.Text & " Amount-" & TxtAmount.Text & " Disc " & TxtDiscount.Text & " Name- " & TxtCustomerName.Text & " Descritption " & TxtDescription.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmRecoveryCustomerWise, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtRecoveryID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpRecoveryDate.Date), "Pending Update Code-" & TxtCustomerID.Text & " Amount-" & TxtAmount.Text & " Disc " & TxtDiscount.Text & " Name- " & TxtCustomerName.Text & " Descritption " & TxtDescription.Text)
   End If
      With Grid
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
         TxtTotalDiscount.Text = Val(TxtTotalDiscount.Text) + Val(TxtDiscount.Text) - Val(Grid.Columns("Discount").Text)
         .Columns("CustomerName").Text = TxtCustomerName.Text
         .Columns("CompanyName").Text = CmbCompany.Text
         
         If CmbCompany.Text <> "" Then .Columns("CompanyID").Value = CmbCompany.ItemData(CmbCompany.ListIndex)
         .Columns("PreviousReceivable").Value = Val(TxtPreviousReceivable.Text)
         .Columns("Amount").Value = Val(TxtAmount.Text)
         .Columns("ManualNo").Text = TxtManualNo.Text
         .Columns("Description").Text = TxtDescription.Text
         .Columns("Discount").Value = IIf(Val(TxtDiscount.Text) = 0, 0, Val(TxtDiscount.Text))
         .Columns("FinalCredit").Value = Val(TxtFinalCredit.Text)
         
         RsBody!PreviousReceivable = IIf(Val(TxtPreviousReceivable.Text) = 0, 0, Val(TxtPreviousReceivable.Text))
         If CmbCompany.Text <> "" Then RsBody!companyid = CmbCompany.ItemData(CmbCompany.ListIndex)
         RsBody!Amount = IIf(Val(TxtAmount.Text) = 0, 0, Val(TxtAmount.Text))
         RsBody!ManualNo = IIf(Trim(TxtManualNo.Text) = "", Null, TxtManualNo.Text)
         RsBody!Description = IIf(Trim(TxtDescription.Text) = "", Null, TxtDescription.Text)
         RsBody!Discount = IIf(Val(TxtDiscount.Text) = 0, 0, Val(TxtDiscount.Text))
         RsBody.Update
         .MoveLast
         
         If Trim(.Columns("CustomerID").Text) <> "" Then
            .AllowAddNew = True
            .AddNew
            .Columns("CustomerID").Text = " "
            .AllowAddNew = False
         End If
      End With
   Call SubClearDetailArea
   TxtCustomerID.SetFocus
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.CmbFilter = "Customers"
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%' or c.AccountNo like '5%') and c.isDetailed = 1 and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select c.*, P.Address FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " Left Outer join Members m on c.AccountNo = cast(m.Prefix as varchar(2))  + cast(m.MemberID as varchar(10)) " & vbCrLf & _
              " where p.BarCode = '" & (TxtCustomerID.Text) & "' or m.BarCode = '" & (TxtCustomerID.Text) & "' or (c.AccountNo = " & Val(TxtCustomerID.Text) & " and (c.AccountNo like '6%' or c.AccountNo like '5%') and isDetailed = 1 and isLocked = 0)"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerID.Text = !AccountNo
          TxtCustomerName.Text = !AccountName & IIf(IsNull(!Address), "", " (" & !Address & " )")
          TxtPreviousReceivable.Text = cn.Execute("SELECT isnull(dbo.FunCurrentDebit(" & Val(TxtCustomerID.Text) & ",'" & DtpRecoveryDate.DateValue & "'," & IIf(Val(TxtOrganizationID.Text) = 0, "Null", Val(TxtOrganizationID.Text)) & "),0) ").Fields(0).Value
          TxtManualNo.Text = FunGetMaxManualID
          FunSelectCustomer = True
          .Close
          Exit Function
      Else
          FunSelectCustomer = False
          .Close
          TxtCustomerID.Text = ""
          TxtCustomerName.Text = ""
      End If
    End With
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub CalculateNetAmount()
   TxtNetAmount.Text = Val(TxtTotalAmount.Text) - Val(TxtTotalDiscount.Text)
End Sub

Private Sub TxtNetAmount_Change()
'   Call CalculateNetAmount
   LblWords.Caption = StrConv(Words_Money_Only(Val(TxtNetAmount.Text)), vbProperCase)
End Sub

Private Sub TxtTotalAmount_Change()
   Call CalculateNetAmount
End Sub

Private Sub TxtTotalDiscount_Change()
   Call CalculateNetAmount
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   If DtpRecoveryDate.IsDateValid = False Then Exit Function
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_RecoveryHeader ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
     If vIsNewRecord = False Then
    With cn.Execute("Select  * from RecoveryHeader where RecoveryID =" & TxtRecoveryID.Text & " And RecoveryDate = '" & DtpRecoveryDate.DateValue & "'")
        If Val(TxtEmployeeID.Text) <> IIf(IsNull(!EmpID), 0, !EmpID) Then
            cn.Execute ("Insert Into UserActivities values ('Recovery Customer Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Updated EmpID-" & !EmpID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
    Grid.MoveFirst
    
    For i = 1 To Grid.rows - 1
        With cn.Execute("Select * from RecoveryCustomer Where RecoveryID = " & TxtRecoveryID.Text & " and CustomerID = " & Val(Grid.Columns("CustomerID").Text))
             If .EOF = True Then
                cn.Execute ("Insert Into UserActivities values ('Recovery Customer Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Inserted New CustomerID-" & Grid.Columns("CustomerID").Text & " PreviousReceivable-" & Grid.Columns("PreviousReceivable").Text & " Amount-" & Grid.Columns("Amount").Text & " Disc " & Grid.Columns("Discount").Text & " FinalCredit-" & Grid.Columns("FinalCredit").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
             Else
                If Grid.Columns("PreviousReceivable").Text <> !PreviousReceivable Or Grid.Columns("Amount").Text <> !Amount Or Grid.Columns("Discount").Text <> !Discount Then
                   cn.Execute ("Insert Into UserActivities values ('Recovery Customer Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Updated CustomerID-" & Grid.Columns("CustomerID").Text & " PreviousReceivable-" & !PreviousReceivable & " Disc-" & !Discount & " Amount-" & !Amount & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                End If
            End If
        End With
    Grid.MoveNext
    Next
   Else
    cn.Execute ("Insert Into UserActivities values ('Recovery Customer Wise'" & "," & TxtRecoveryID.Text & ",'" & DtpRecoveryDate.DateValue & "','Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub BtnBankMachine_Click()
   On Error GoTo ErrorHandler
   If FunSelectBankMachine(ssButton, False) = True Then
      If TxtCustomerID.Visible And TxtCustomerID.Enabled Then TxtCustomerID.SetFocus Else TxtBankMachineID.SetFocus
   Else
      TxtBankMachineID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectBankMachine(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchBankMachine.Show vbModal, Me
        If SchBankMachine.ParaOutBankMachineID = "" Then FunSelectBankMachine = False: Exit Function
        TxtBankMachineID.Text = SchBankMachine.ParaOutBankMachineID
    End If
    '---------------------------
    vStrSQL = " Select * FROM BankMachines where BankMachineID=" & Val(TxtBankMachineID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtBankMachineName.Text = !BankMachineName
          TxtCommision.Text = !Commision
          FunSelectBankMachine = True
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          .Close
          Exit Function
      Else
          FunSelectBankMachine = False
          .Close
          TxtBankMachineID.Text = ""
          TxtBankMachineName.Text = ""
          TxtCommision.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtBankMachineID_Change()
   If TxtBankMachineID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtBankMachineID.Name Then Exit Sub
   If TxtBankMachineName.Text <> "" Then
      TxtBankMachineName.Text = ""
      TxtCommision.Text = ""
   End If
End Sub

Private Sub TxtBankMachineID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtBankMachineID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtBankMachineName.Text <> "" Then Exit Sub
   If Trim(TxtBankMachineID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBankMachine(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBankMachine(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxManualID() As String
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("CustomerID").Text) = "" Then
    vManualNo = 0
    Grid.Redraw = False
    Grid.MoveFirst
        For i = 1 To Grid.rows
            If Val(Grid.Columns("ManualNo").Text) > Val(vManualNo) Then
                vManualNo = Val(Grid.Columns("ManualNo").Text)
            End If
            Grid.MoveNext
        Next i
    Grid.Redraw = True
   End If
   If Val(vManualNo) > 0 Then
    FunGetMaxManualID = vManualNo + 1
   Else
    FunGetMaxManualID = ""
   End If
   
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.RecoveryHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmRecoveryCustomerWise) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmRecoveryCustomerWise & ", " & vUser & "," & TableHeaderFields(eFrmRecoveryCustomerWise) & " from RecoveryHeader " & vbCrLf _
             & "Where RecoveryID = " & TxtRecoveryID.Text & " and RecoveryDate = '" & DtpRecoveryDate.DateValue & "'"
      cn.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.RecoveryCustomerBin (" & TableBodyFields(eFrmRecoveryCustomerWise) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmRecoveryCustomerWise) & " from RecoveryCustomer " & vbCrLf _
             & "Where RecoveryID = " & TxtRecoveryID.Text
      cn.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub



