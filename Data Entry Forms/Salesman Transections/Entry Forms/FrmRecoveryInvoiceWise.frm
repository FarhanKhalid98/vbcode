VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmRecoveryInvoiceWise 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmRecoveryInvoiceWise.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnSaleMan 
      Height          =   330
      Left            =   4035
      TabIndex        =   31
      Top             =   1575
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0ECA
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtRecoveryID 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   1575
      Width           =   1050
      _ExtentX        =   1852
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
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7350
      TabIndex        =   16
      Top             =   8115
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6030
      TabIndex        =   12
      Top             =   8115
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3390
      TabIndex        =   14
      Top             =   8115
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8670
      TabIndex        =   17
      Top             =   8115
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4710
      TabIndex        =   13
      Top             =   8115
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSaleID 
      Height          =   315
      Left            =   83
      TabIndex        =   4
      Top             =   2625
      Width           =   1245
      _ExtentX        =   2196
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
   Begin JeweledBut.JeweledButton BtnSale 
      Height          =   330
      Left            =   1328
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2610
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   1688
      TabIndex        =   6
      Top             =   2625
      Width           =   2505
      _ExtentX        =   4419
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
      Height          =   4650
      Left            =   83
      TabIndex        =   20
      Top             =   2940
      Width           =   11835
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "FrmRecoveryInvoiceWise.frx":0F8E
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
      Columns(0).Width=   2831
      Columns(0).Caption=   "Sale ID"
      Columns(0).Name =   "SaleID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4419
      Columns(1).Caption=   "Customer Name"
      Columns(1).Name =   "CustomerName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2619
      Columns(2).Caption=   "Sale Value"
      Columns(2).Name =   "SaleValue"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2672
      Columns(3).Caption=   "Received"
      Columns(3).Name =   "Received"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2619
      Columns(4).Caption=   "Amount"
      Columns(4).Name =   "Amount"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2646
      Columns(5).Caption=   "Discount"
      Columns(5).Name =   "Discount"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2566
      Columns(6).Caption=   "Final Credit"
      Columns(6).Name =   "FinalCredit"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20876
      _ExtentY        =   8202
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpRecoveryDate 
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Top             =   1575
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
   Begin SITextBox.Txt TxtSaleValue 
      Height          =   315
      Left            =   4193
      TabIndex        =   7
      Top             =   2625
      Width           =   1485
      _ExtentX        =   2619
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
      Left            =   7193
      TabIndex        =   9
      Top             =   2625
      Width           =   1485
      _ExtentX        =   2619
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtReceived 
      Height          =   315
      Left            =   5678
      TabIndex        =   8
      Top             =   2625
      Width           =   1515
      _ExtentX        =   2672
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
      DecimalPoint    =   7
      IntegralPoint   =   2
   End
   Begin SITextBox.Txt TxtDiscount 
      Height          =   315
      Left            =   8678
      TabIndex        =   10
      Top             =   2625
      Width           =   1500
      _ExtentX        =   2646
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
      DecimalPoint    =   2
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtFinalCredit 
      Height          =   315
      Left            =   10178
      TabIndex        =   11
      Top             =   2625
      Width           =   1470
      _ExtentX        =   2593
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2055
      TabIndex        =   15
      Top             =   8115
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
      MICON           =   "FrmRecoveryInvoiceWise.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSaleManName 
      Height          =   315
      Left            =   4395
      TabIndex        =   3
      Top             =   1575
      Width           =   3210
      _ExtentX        =   5662
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
   Begin SITextBox.Txt TxtSaleManID 
      Height          =   315
      Left            =   2850
      TabIndex        =   2
      Top             =   1575
      Width           =   1185
      _ExtentX        =   2090
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
      Masked          =   1
      Mandatory       =   1
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Man Name"
      Height          =   375
      Left            =   4380
      TabIndex        =   30
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Man ID"
      Height          =   255
      Left            =   2850
      TabIndex        =   29
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recovery (Invoice Wise)"
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
      Left            =   1920
      TabIndex        =   28
      Top             =   120
      Width           =   4245
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Left            =   8723
      TabIndex        =   27
      Top             =   2430
      Width           =   630
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Credit"
      Height          =   195
      Left            =   10208
      TabIndex        =   26
      Top             =   2430
      Width           =   780
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      Height          =   195
      Left            =   5708
      TabIndex        =   25
      Top             =   2430
      Width           =   690
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   7208
      TabIndex        =   24
      Top             =   2430
      Width           =   540
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Value"
      Height          =   195
      Left            =   4178
      TabIndex        =   23
      Top             =   2430
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale ID"
      Height          =   195
      Left            =   83
      TabIndex        =   22
      Top             =   2430
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   195
      Left            =   1748
      TabIndex        =   21
      Top             =   2430
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
      Left            =   1395
      TabIndex        =   19
      Top             =   1380
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Recovery ID"
      Height          =   195
      Left            =   105
      TabIndex        =   18
      Top             =   1380
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
Attribute VB_Name = "FrmRecoveryInvoiceWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim RsBody As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim sSql As String
Dim vStrSQL As String
'----------------------------------

Private Sub CalculateBody()
   TxtFinalCredit.Text = Val(TxtSaleValue.Text) - Val(TxtReceived.Text) - Val(TxtAmount.Text) - Val(TxtDiscount.Text)
End Sub

Private Sub BtnClear_Click()
On Error GoTo ErrorHandler
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub BtnDelete_Click()
On Error GoTo ErrorHandler
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
'      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   CN.BeginTrans
   Grid.Redraw = False
   Grid.RemoveAll
   CN.Execute "Delete from RecoveryInvoice where RecoveryID = " & Val(TxtRecoveryID.Text)
   Grid.Redraw = True
   CN.Execute "Delete from RecoveryHeader where RecoveryID = " & Val(TxtRecoveryID.Text)
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchRecoveryInvoice.Show vbModal
   If SchRecoveryInvoice.ParaOutRecoveryID <> 0 Then
      TxtRecoveryID.Text = SchRecoveryInvoice.ParaOutRecoveryID
      GetRecovery
   End If
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
      vStrSQL = " Select  H.RecoveryID, H.RecoveryDate, SM.SaleManName, Pty.PartyName, B.SaleID, Sv.SaleValue , SV.Received," & vbCrLf _
      + " Isnull(B.Amount,0) Amount, Isnull(B.Discount,0) Discount" & vbCrLf _
      + " from RecoveryHeader H " & vbCrLf _
      + " Inner join RecoveryInvoice B on H.RecoveryID = B.RecoveryID " & vbCrLf _
      + " Inner join SaleHeader SH on SH.SaleID = H.RecoveryID" & vbCrLf _
      + " Left Join Parties pty on pty.partyID = SH.CustomerID" & vbCrLf _
      + " Left Join SalesMan SM on SM.SaleManID = H.SaleManID" & vbCrLf _
      + " inner Join (select h.saleid, totalamount - isnull(billdisc,0) as SaleValue, (totalamount - isnull(billdisc,0) - isnull(ReceivedAmount,0) - isnull(PreReceived,0) ) as  Received " & vbCrLf _
      + " from saleheader h left outer join " & vbCrLf _
      + " (select b.saleid, sum(b.amount-isnull(b.discount,0)) PreReceived from recoveryInvoice b  inner join RecoveryHeader H on b.RecoveryID = H.RecoveryID where recoveryDate < '" & DtpRecoveryDate.DateValue & "' group by saleid) i on h.saleid = i.saleid ) Sv" & vbCrLf _
      + " on b.SaleID = Sv.SaleID" & vbCrLf _
      + " where h.recoveryID=" & Val(TxtRecoveryID.Text)
      
    If RsReport.State = adStateOpen Then RsReport.Close
    RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
  
    Set RptReportViewer.Report = New CrpRecoveryInvoice
    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String
    vStrComp = "Select CompanyName,Address,City,PhoneNo,email from Company"
    With CN.Execute(vStrComp)
      If .RecordCount > 0 Then
         vCompanyName = !CompanyName
         vAddress = !Address & IIf(IsNull(!City), "", ", " & !City)
         vPhone = IIf(IsNull(!PhoneNo), "", "Phone # " & !PhoneNo)
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue vAddress
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue vPhone
      End If
   End With
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
   'RptReportViewer.Report.PrintOut False
   RptReportViewer.Show vbModal, Me
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSale_Click()
   If FunSelectSale(ssButton, False) = True Then
      TxtAmount.SetFocus
   Else
      TxtSaleID.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
    MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
    Exit Sub
  End If
'  Header Validation
'   If Trim(TxtSaleID.Text) = "" Then
'      MsgBox "Enter SaleMan ID.", vbExclamation, Me.Caption
'      TxtSaleID.SetFocus
'      Exit Sub
'   End If
   If TxtRecoveryID.Enabled Then
      If CN.Execute("Select * from RecoveryHeader where RecoveryID = " & Val(TxtRecoveryID.Text)).RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtRecoveryID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
  RsBody.Filter = 0
  If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one entry for recovery", vbExclamation, "Alert"
      If TxtRecoveryID.Visible And TxtRecoveryID.Enabled Then TxtRecoveryID.SetFocus
      Exit Sub
  End If
  'Body Validation
  ' validation has been performed when a row is added to the grid
  
  'Saving record
   CN.BeginTrans
   sSql = "select * from RecoveryHeader where RecoveryID =" & Val(TxtRecoveryID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, CN, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !RecoveryID = Val(TxtRecoveryID.Text)
      End If
     !RecoveryDate = DtpRecoveryDate.DateValue
     !SalemanID = IIf(Trim(TxtSaleManID.Text) = "", Null, TxtSaleManID.Text)
     !UserNo = vUser
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
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
  ' If CN.Errors.Count > 0 Then CN.RollbackTrans
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
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSave.Enabled And BtnSave.Visible Then BtnSave_Click
            KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
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
         Case TxtSaleManID.Name: If FunSelectSaleman(ssFunctionKey, True) = True Then If TxtSaleID.Enabled Then TxtSaleID.SetFocus
         Case TxtSaleID.Name: If FunSelectSale(ssFunctionKey, False) = True Then TxtAmount.SetFocus
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

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Recovery (Invoice Wise)"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
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
   FunGetMaxID = CN.Execute("Select isnull(max(RecoveryID),0)+1 from RecoveryHeader").Fields(0)
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
   Grid.Columns("SaleID").Text = " "
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
    Set FrmRecoveryInvoiceWise = Nothing
   End If
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
   TxtSaleID.Enabled = False
   BtnSale.Enabled = False
   'TxtRecoveryID.BackColor = TxtCustomerName.BackColor
   'TxtRecoveryID.TabStop = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("SaleID").Text) = "" Then
      TxtSaleID.Text = ""
      TxtSaleID.Enabled = True
      BtnSale.Enabled = True
      TxtSaleID.SetFocus
   Else
      TxtSaleID.Enabled = False
      BtnSale.Enabled = False
      TxtAmount.SetFocus
      If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(Grid.Columns("SaleID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub GetRecovery()
   On Error GoTo ErrorHandler
   sSql = " select h.*, salemanname FROM RecoveryHeader h inner join RecoveryInvoice b on h.recoveryid = b.recoveryid left outer join salesman sm on sm.salemanid = h.salemanid where h.RecoveryID=" & Val(TxtRecoveryID.Text)
   With CN.Execute(sSql)
      If Not .BOF Then
          DtpRecoveryDate.DateValue = !RecoveryDate
          TxtSaleManID.Text = IIf(IsNull(!SalemanID) = True, "", !SalemanID)
          TxtSalemanName.Text = IIf(IsNull(!SalemanName) = True, "", !SalemanName)
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

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("SaleID").Text) = "" Then Exit Sub
   RsBody.Filter = "SaleID='" & TxtSaleID.Text & "'"
   If RsBody.RecordCount > 0 Then RsBody.Delete
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
      Call PopulateDataToGrid
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      BtnSale.Enabled = True
      TxtRecoveryID.Enabled = True
      TxtRecoveryID.Text = FunGetMaxID()
      If DtpRecoveryDate.Enabled And DtpRecoveryDate.Visible Then DtpRecoveryDate.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      TxtRecoveryID.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      BtnSale.Enabled = True
      TxtSaleID.Enabled = True
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

Private Sub PopulateDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from RecoveryInvoice where RecoveryID =" & Val(TxtRecoveryID.Text), CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = " Select  B.SaleID, Pty.PartyName as CustomerName, Sv.SaleValue , SV.Received," & vbCrLf _
      + " Isnull(B.Amount,0) Amount, Isnull(B.Discount,0) Discount" & vbCrLf _
      + " from RecoveryHeader H Inner join RecoveryInvoice B on H.RecoveryID = B.RecoveryID " & vbCrLf _
      + " Inner join SaleHeader SH on SH.SaleID = H.RecoveryID" & vbCrLf _
      + " Left Join Parties pty on pty.partyID = SH.CustomerID" & vbCrLf _
      + " Left Join SalesMan SM on SM.SaleManID = H.SaleManID" & vbCrLf _
      + " inner Join (" & vbCrLf _
      + " select h.saleid, totalamount - isnull(billdisc,0) as SaleValue, (totalamount - isnull(billdisc,0) - isnull(ReceivedAmount,0) - isnull(PreReceived,0) ) as  Received " & vbCrLf _
      + " from saleheader h left outer join " & vbCrLf _
      + " (select b.saleid, sum(b.amount-isnull(b.discount,0)) PreReceived from recoveryInvoice b  inner join RecoveryHeader H on b.RecoveryID = H.RecoveryID where recoveryDate < '4/12/2007' group by saleid) i on h.saleid = i.saleid ) Sv" & vbCrLf _
      + " on b.SaleID = Sv.SaleID" & vbCrLf _
      + " where h.recoveryid = " & Val(TxtRecoveryID.Text)
      
      'sSql = "select b.* from RecoveryInvoice b where RecoveryID =" & Val(TxtRecoveryID.Text)
      With CN.Execute(sSql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            Grid.AddNew
            Grid.Columns("SaleID").Value = !SaleID
            Grid.Columns("CustomerName").Value = IIf(IsNull(!CustomerName), "", !CustomerName)
            Grid.Columns("SaleValue").Value = !SaleValue
            Grid.Columns("Received").Value = IIf(IsNull(!Received), "", !Received)
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("Discount").Value = IIf(IsNull(!Discount), "", !Discount)
            Grid.Columns("FinalCredit").Value = Val(Grid.Columns("SaleValue").Value) - Val(Grid.Columns("Received").Value) - Val(Grid.Columns("Amount").Value) - Val(Grid.Columns("Discount").Value)
            'TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(!FinalCredit)
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("SaleID").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtSaleID.Text = .Columns("SaleID").Text
      TxtCustomerName.Text = .Columns("CustomerName").Text
      TxtSaleValue.Text = .Columns("Salevalue").Value
      TxtReceived.Text = .Columns("Received").Value
      TxtAmount.Text = .Columns("Amount").Value
      TxtDiscount.Text = .Columns("Discount").Value
      TxtFinalCredit.Text = .Columns("FinalCredit").Value
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtSaleID.Enabled = True
   BtnSale.Enabled = True
   TxtSaleID.Text = ""
   TxtCustomerName.Text = ""
   TxtSaleValue.Text = ""
   TxtAmount.Text = ""
   TxtReceived.Text = ""
   TxtDiscount.Text = ""
   TxtFinalCredit.Text = ""
End Sub

Private Function FunSelectSaleman(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSaleman.Show vbModal, Me
        If SchSaleman.ParaOutSalemanID = "" Then FunSelectSaleman = False: Exit Function
        TxtSaleManID.Text = SchSaleman.ParaOutSalemanID
    End If
    '---------------------------
    vStrSQL = " Select * FROM SalesMan where SaleManID=" & Val(TxtSaleManID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSalemanName.Text = !SalemanName
          FunSelectSaleman = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSaleman = False
          .Close
          TxtSaleID.Text = ""
          TxtSalemanName.Text = ""
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

Private Sub BtnSaleman_Click()
   If FunSelectSaleman(ssButton, False) = True Then
      TxtSaleID.SetFocus
   Else
      TxtSaleManID.SetFocus
   End If
End Sub

Private Sub TxtAmount_Change()
   Call CalculateBody
End Sub

Private Sub TxtDiscount_Change()
   Call CalculateBody
End Sub

Private Sub txtDiscount_LostFocus()
Select Case ActiveControl.Name
   Case TxtSaleID.Name, TxtCustomerName.Name, TxtSaleValue.Name, TxtReceived.Name, TxtAmount.Name
      Exit Sub
   End Select
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub TxtSaleID_Change()
   If TxtSaleID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSaleID.Name Then Exit Sub
   If TxtCustomerName.Text <> "" Then
      TxtCustomerName.Text = ""
      TxtSaleValue.Text = ""
      TxtReceived.Text = ""
   End If
End Sub

Private Sub TxtSaleID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtSaleID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSaleID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerName.Text <> "" Then Exit Sub
   If Trim(TxtSaleID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSale(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSale(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSalemanID_Change()
   If TxtSaleManID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSaleManID.Name Then Exit Sub
   If TxtSalemanName.Text <> "" Then
      TxtSalemanName.Text = ""
   End If
End Sub

Private Sub TxtSalemanID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSaleManID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSalemanName.Text <> "" Then Exit Sub
   If Trim(TxtSaleManID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSaleman(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSaleman(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   Dim vrowcounter As Integer
   If Trim(TxtSaleID.Text) = "" Then
      TxtSaleID.SetFocus
      Exit Sub
   End If
   If Val(TxtAmount.Text) = 0 Then
      TxtAmount.SetFocus
      Exit Sub
   End If
   If Val(TxtFinalCredit.Text) < 0 Then
      TxtAmount.SetFocus
      Exit Sub
   End If
On Error GoTo ErrorHandler
      RsBody.Filter = "SaleID ='" & TxtSaleID.Text & "'"
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("SaleID").Text = TxtSaleID.Text
         RsBody!SaleID = TxtSaleID.Text
       End If
       With Grid
         .Columns("CustomerName").Text = TxtCustomerName.Text
         .Columns("SaleValue").Value = Val(TxtSaleValue.Text)
         .Columns("Received").Value = Val(TxtReceived.Text)
         .Columns("Amount").Value = Val(TxtAmount.Text)
         .Columns("Discount").Value = IIf(Val(TxtDiscount.Text) = 0, 0, Val(TxtDiscount.Text))
         .Columns("FinalCredit").Value = Val(TxtFinalCredit.Text)
         RsBody!Amount = Val(TxtAmount.Text)
         RsBody!Discount = IIf(Val(TxtDiscount.Text) = 0, 0, Val(TxtDiscount.Text))
         .MoveLast
         If Trim(.Columns("SaleID").Text) <> "" Then
            .AllowAddNew = True
            .AddNew
            .Columns("SaleID").Text = " "
            .AllowAddNew = False
         End If
      End With
   Call SubClearDetailArea
   TxtSaleID.SetFocus
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Function FunSelectSale(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSale.Show vbModal, Me
        If SchSale.ParaOutSaleID = 0 Then FunSelectSale = False: Exit Function
        TxtSaleID.Text = SchSale.ParaOutSaleID
    End If
    '---------------------------
    vStrSQL = " select h.saleid, PartyName, totalamount - isnull(billdisc,0) as SaleValue, " & vbCrLf _
      + " (totalamount - isnull(billdisc,0) - isnull(ReceivedAmount,0) - isnull(PreReceived,0) ) as  Received " & vbCrLf _
      + " from saleheader h left outer join  " & vbCrLf _
      + " (select saleid, sum(amount-isnull(discount,0)) PreReceived " & vbCrLf _
      + " from recoveryInvoice group by saleid) i on  h.saleid = i.saleid  " & vbCrLf _
      + " inner join parties p on p.partyid = h.CustomerID " & vbCrLf _
      + " where h.saleid = " & Val(TxtSaleID.Text)
      
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerName.Text = !PartyName
          TxtSaleValue.Text = !SaleValue
          TxtReceived.Text = !Received
          FunSelectSale = True
          Call CalculateBody
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSale = False
          .Close
          TxtCustomerName.Text = ""
          TxtSaleValue.Text = ""
          TxtReceived.Text = ""
          Call CalculateBody
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
