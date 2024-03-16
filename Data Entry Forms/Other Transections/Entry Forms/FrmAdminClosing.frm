VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmAdminClosing 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmAdminClosing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtUserFineonShort 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   14280
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   9585
      Width           =   810
   End
   Begin VB.TextBox TxtCompanyFineonShort 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   14280
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   9945
      Width           =   810
   End
   Begin VB.TextBox TxtAdminClssingFinePerOnShort 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   14280
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   9225
      Width           =   810
   End
   Begin VB.TextBox TxtBankPayments 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   8145
      Width           =   1305
   End
   Begin VB.TextBox TxtBankReceived 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4365
      Width           =   1305
   End
   Begin VB.TextBox TxtCreditSaleReturnPaid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   5445
      Width           =   1305
   End
   Begin VB.Frame FrameDetail 
      Caption         =   "Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   5940
      TabIndex        =   60
      Top             =   6165
      Width           =   4290
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridDetail 
         Height          =   2025
         Left            =   120
         TabIndex        =   61
         Top             =   270
         Width           =   4080
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         Col.Count       =   4
         stylesets.count =   2
         stylesets(0).Name=   "SelectedCol"
         stylesets(0).ForeColor=   0
         stylesets(0).BackColor=   12713983
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
         stylesets(0).Picture=   "FrmAdminClosing.frx":0ECA
         stylesets(1).Name=   "SelectedRow"
         stylesets(1).ForeColor=   16777215
         stylesets(1).BackColor=   8388608
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "FrmAdminClosing.frx":0EE6
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
         ActiveCellStyleSet=   "SelectedCol"
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
         SelectTypeRow   =   0
         ForeColorEven   =   0
         BackColorOdd    =   15724527
         RowHeight       =   423
         ExtraHeight     =   106
         Columns.Count   =   4
         Columns(0).Width=   1164
         Columns(0).Caption=   "ID"
         Columns(0).Name =   "ID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Date"
         Columns(1).Name =   "Date"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3254
         Columns(2).Caption=   "Party"
         Columns(2).Name =   "Party"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1640
         Columns(3).Caption=   "Amount"
         Columns(3).Name =   "Amount"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   7197
         _ExtentY        =   3572
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
      Begin JeweledBut.JeweledButton BtnFrameClose 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1530
         TabIndex        =   62
         Top             =   2340
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
         MICON           =   "FrmAdminClosing.frx":0F02
         BC              =   14737632
         FC              =   0
      End
   End
   Begin VB.TextBox TxtCashReceived 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4050
      Width           =   1305
   End
   Begin VB.TextBox TxtPayments 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   7830
      Width           =   1305
   End
   Begin VB.TextBox TxtRecoveryCustomer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3510
      Width           =   1305
   End
   Begin VB.TextBox TxtCashAvailable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   8370
      Width           =   1305
   End
   Begin VB.TextBox TxtSaleReturn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   7290
      Width           =   1305
   End
   Begin VB.TextBox TxtDiscount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5670
      Width           =   1305
   End
   Begin VB.TextBox TxtCreditSale 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5130
      Width           =   1305
   End
   Begin VB.TextBox TxtBankCardSale 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4590
      Width           =   1305
   End
   Begin VB.TextBox TxtPettyCash 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2970
      Width           =   1305
   End
   Begin VB.TextBox TxtTotalSale 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2430
      Width           =   1305
   End
   Begin VB.TextBox TxtServiceCharges 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   6210
      Width           =   1305
   End
   Begin VB.TextBox TxtSTax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4598
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6750
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EFC09E&
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   13500
      TabIndex        =   23
      Top             =   4140
      Width           =   8730
      Begin SITextBox.Txt TxtTotalOpening 
         Height          =   315
         Left            =   4050
         TabIndex        =   24
         Top             =   4230
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Alignment       =   1
         Appearance      =   0
         MaxLength       =   7
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
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid PGrid 
         Height          =   3735
         Left            =   360
         TabIndex        =   25
         Top             =   495
         Width           =   7995
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         Col.Count       =   6
         stylesets.count =   2
         stylesets(0).Name=   "SelectedCol"
         stylesets(0).ForeColor=   0
         stylesets(0).BackColor=   12713983
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
         stylesets(0).Picture=   "FrmAdminClosing.frx":0F1E
         stylesets(1).Name=   "SelectedRow"
         stylesets(1).ForeColor=   16777215
         stylesets(1).BackColor=   8388608
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "FrmAdminClosing.frx":0F3A
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
         ActiveCellStyleSet=   "SelectedCol"
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
         SelectTypeRow   =   0
         ForeColorEven   =   0
         BackColorOdd    =   15724527
         RowHeight       =   423
         ExtraHeight     =   106
         Columns.Count   =   6
         Columns(0).Width=   1852
         Columns(0).Caption=   "Product ID"
         Columns(0).Name =   "ProductID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   5900
         Columns(1).Caption=   "Product Name"
         Columns(1).Name =   "ProductName"
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   1323
         Columns(2).Caption=   "Opening"
         Columns(2).Name =   "Opening"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   2
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1323
         Columns(3).Caption=   "Sale"
         Columns(3).Name =   "Sale"
         Columns(3).Alignment=   1
         Columns(3).CaptionAlignment=   2
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   1323
         Columns(4).Caption=   "Return"
         Columns(4).Name =   "Return"
         Columns(4).Alignment=   1
         Columns(4).CaptionAlignment=   2
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   1323
         Columns(5).Caption=   "Closing"
         Columns(5).Name =   "Closing"
         Columns(5).Alignment=   1
         Columns(5).CaptionAlignment=   2
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   14102
         _ExtentY        =   6588
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
      Begin SITextBox.Txt TxtTotalClosing 
         Height          =   315
         Left            =   7155
         TabIndex        =   27
         Top             =   4230
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Alignment       =   1
         Appearance      =   0
         MaxLength       =   7
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
      End
      Begin JeweledBut.JeweledButton BtnOpeningProduct 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   5985
         TabIndex        =   29
         Top             =   45
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   741
         TX              =   "Go To Opening Product"
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
         MICON           =   "FrmAdminClosing.frx":0F56
         BC              =   14737632
         FC              =   0
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Closing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5670
         TabIndex        =   28
         Top             =   4275
         Width           =   1395
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Opening"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2475
         TabIndex        =   26
         Top             =   4275
         Width           =   1485
      End
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   12188
      MaxLength       =   50
      TabIndex        =   4
      Top             =   8415
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox TxtExcessShort 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11348
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7560
      Width           =   2085
   End
   Begin VB.TextBox TxtAddCollection 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11348
      TabIndex        =   3
      Top             =   6840
      Width           =   2085
   End
   Begin VB.TextBox TxtTotalCash 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11348
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6075
      Width           =   2085
   End
   Begin VB.ComboBox CmbUsers 
      Height          =   315
      ItemData        =   "FrmAdminClosing.frx":0F72
      Left            =   6023
      List            =   "FrmAdminClosing.frx":0F74
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1935
      Width           =   1740
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8963
      TabIndex        =   9
      Top             =   9045
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
      MICON           =   "FrmAdminClosing.frx":0F76
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7643
      TabIndex        =   5
      Top             =   9045
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
      MICON           =   "FrmAdminClosing.frx":0F92
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5003
      TabIndex        =   7
      Top             =   9045
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
      MICON           =   "FrmAdminClosing.frx":0FAE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10283
      TabIndex        =   10
      Top             =   9045
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
      MICON           =   "FrmAdminClosing.frx":0FCA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6323
      TabIndex        =   6
      Top             =   9045
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
      MICON           =   "FrmAdminClosing.frx":0FE6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   3105
      TabIndex        =   11
      Top             =   1950
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpEntryDate 
      Height          =   315
      Left            =   4448
      TabIndex        =   0
      Top             =   1950
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3668
      TabIndex        =   8
      Top             =   9045
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
      MICON           =   "FrmAdminClosing.frx":1002
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3285
      Left            =   6885
      TabIndex        =   21
      Top             =   2835
      Width           =   6555
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   8
      stylesets.count =   2
      stylesets(0).Name=   "SelectedCol"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   12713983
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
      stylesets(0).Picture=   "FrmAdminClosing.frx":101E
      stylesets(1).Name=   "SelectedRow"
      stylesets(1).ForeColor=   16777215
      stylesets(1).BackColor=   8388608
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "FrmAdminClosing.frx":103A
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
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
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   106
      Columns.Count   =   8
      Columns(0).Width=   2117
      Columns(0).Caption=   "Denomination"
      Columns(0).Name =   "Denom"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   714
      Columns(1).Name =   "Mul"
      Columns(1).Alignment=   2
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2778
      Columns(2).Caption=   "Petty Cash Quantity"
      Columns(2).Name =   "PQty"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1296
      Columns(3).Caption=   "Quantity"
      Columns(3).Name =   "Qty"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   794
      Columns(4).Name =   "Equ"
      Columns(4).Alignment=   2
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   2646
      Columns(5).Caption=   "Amount"
      Columns(5).Name =   "Amount"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "PAmount"
      Columns(6).Name =   "PAmount"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "QAmount"
      Columns(7).Name =   "QAmount"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   11562
      _ExtentY        =   5794
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
   Begin JeweledBut.JeweledButton BtnPettyCash 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11648
      TabIndex        =   22
      Top             =   2340
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      TX              =   "Go To Petty Cash"
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
      MICON           =   "FrmAdminClosing.frx":1056
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   7913
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   1935
      Width           =   675
      _ExtentX        =   1191
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
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   8948
      TabIndex        =   30
      Tag             =   "NC"
      Top             =   1935
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8588
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1935
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
      MICON           =   "FrmAdminClosing.frx":1072
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSetSystemDate 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   1883
      TabIndex        =   59
      Top             =   9045
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      TX              =   "Set System Date"
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
      MICON           =   "FrmAdminClosing.frx":108E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnEmail 
      Height          =   420
      Left            =   5940
      TabIndex        =   75
      Top             =   10260
      Visible         =   0   'False
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   741
      TX              =   "E-mail by EASendMail not Work"
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
      MICON           =   "FrmAdminClosing.frx":10AA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnExportPDF 
      Height          =   420
      Left            =   3090
      TabIndex        =   76
      Top             =   9630
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   741
      TX              =   "Export Admin Closing PDF"
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
      MICON           =   "FrmAdminClosing.frx":10C6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnEmailCDO 
      Height          =   420
      Left            =   8393
      TabIndex        =   77
      Top             =   9630
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   741
      TX              =   "E-mail By CDO"
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
      MICON           =   "FrmAdminClosing.frx":10E2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnExportActivityPDF 
      Height          =   420
      Left            =   5610
      TabIndex        =   78
      Top             =   9630
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   741
      TX              =   "Export ActivityLog PDF"
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
      MICON           =   "FrmAdminClosing.frx":10FE
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblUserFineonShort 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "User Fine on Short"
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
      Left            =   12240
      TabIndex        =   73
      Top             =   9630
      Width           =   1605
   End
   Begin VB.Label LblCompanyFineonShort 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Fine on Short"
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
      Left            =   12240
      TabIndex        =   72
      Top             =   9990
      Width           =   1980
   End
   Begin VB.Label LblFinePerOnShort 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Fine Per On Short"
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
      Left            =   12240
      TabIndex        =   70
      Top             =   9315
      Width           =   1530
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   68
      Top             =   7905
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Received"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   66
      Top             =   4140
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Crediit Return Paid (-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   64
      Top             =   5160
      Width           =   2250
   End
   Begin VB.Label LblServiceCh 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Ch. (+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2963
      TabIndex        =   58
      Top             =   6285
      Width           =   1530
   End
   Begin VB.Label LblTotalSale 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sale (+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3068
      TabIndex        =   57
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label LblPattyCash 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Petty Cash (+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3038
      TabIndex        =   56
      Top             =   3060
      Width           =   1455
   End
   Begin VB.Label LblBankCardSale 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Card Sale (-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2573
      TabIndex        =   55
      Top             =   4665
      Width           =   1920
   End
   Begin VB.Label LblCreditSale 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Sale (-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3038
      TabIndex        =   54
      Top             =   5205
      Width           =   1455
   End
   Begin VB.Label LblDiscount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount (-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3293
      TabIndex        =   53
      Top             =   5745
      Width           =   1200
   End
   Begin VB.Label LblSaleReturn 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Return (-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2978
      TabIndex        =   52
      Top             =   7365
      Width           =   1515
   End
   Begin VB.Label LblCashAvailable 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2903
      TabIndex        =   51
      Top             =   8445
      Width           =   1590
   End
   Begin VB.Label LblRecoveryCustomer 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Recovery Customer (+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2108
      TabIndex        =   50
      Top             =   3600
      Width           =   2385
   End
   Begin VB.Label LblPayments 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Payments (-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3173
      TabIndex        =   49
      Top             =   7905
      Width           =   1320
   End
   Begin VB.Label LblCashReceived 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Received  (+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2498
      TabIndex        =   48
      Top             =   4140
      Width           =   1995
   End
   Begin VB.Label LblSalesTax 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax (+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3068
      TabIndex        =   47
      Top             =   6825
      Width           =   1395
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7913
      TabIndex        =   34
      Top             =   1665
      Width           =   855
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8948
      TabIndex        =   33
      Top             =   1665
      Width           =   1245
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
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
      Height          =   225
      Left            =   11648
      TabIndex        =   32
      Top             =   8460
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Excess / Short"
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
      Left            =   9983
      TabIndex        =   20
      Top             =   7695
      Width           =   1275
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Add Collection (+)"
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
      Left            =   9713
      TabIndex        =   18
      Top             =   6975
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cash"
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
      Left            =   10343
      TabIndex        =   17
      Top             =   6255
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6023
      TabIndex        =   15
      Top             =   1665
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3105
      TabIndex        =   14
      Top             =   1665
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4448
      TabIndex        =   13
      Top             =   1665
      Width           =   1095
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Closing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   2700
      TabIndex        =   12
      Top             =   270
      Width           =   2025
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11610
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "FrmAdminClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim Rs As New ADODB.Recordset
Dim RsDetail As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.
Dim vid As String
Dim sSql As String, vStrSQL As String, i As Integer
Const ConnectNormal = 0
Const ConnectSSLAuto = 1
Const ConnectSTARTTLS = 2
Const ConnectDirectSSL = 3
Const ConnectTryTLS = 4

Private Sub SubCalculate2()
   On Error GoTo ErrorHandler
   ' Step 1 - Total Sale
   sSql = " Select isnull(round(Sum((isnull(multiplier,0)* isnull(QtyPack,0) + Qty )*((Price/isnull(multiplier,1))+isnull(sc,0))),0),0) as TotalSale" & vbCrLf _
      + " from SaleHeader h inner join SaleBody b on H.SID = B.SID" & vbCrLf _
      + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
   TxtTotalSale.Text = CN.Execute(sSql).Fields(0).Value

   sSql = " Select isnull(Sum(totalamount),0) as TotalSale" & vbCrLf _
      + " from CustomOrderHeader " & vbCrLf _
      + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and OrderDate = '" & DtpEntryDate.DateValue & "'"
   TxtTotalSale.Text = Val(TxtTotalSale.Text) + CN.Execute(sSql).Fields(0).Value
   
'   sSql = " Select isnull(floor(Sum(Qty*Price)),0) as TotalSale" & vbCrLf _
      + " from ServiceHeader h inner join ServiceBody b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
      + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
'   TxtTotalSale.Text = Val(TxtTotalSale.Text) + CN.Execute(sSql).Fields(0).Value
   
   ' Step 2 - Total PettyCash
   sSql = " Select isnull(sum(Amount),0)amount from PettyCashHeader where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and ToUserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and EntryDate = '" & DtpEntryDate.DateValue & "'"
   TxtPettyCash.Text = CN.Execute(sSql).Fields(0).Value

   ' Step 3 - Total Customer Recovery
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " FROM RecoveryHeader h INNER JOIN RecoveryCustomer b ON h.RecoveryId = B.RecoveryId " & vbCrLf _
      + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.RecoveryDate = '" & DtpEntryDate.DateValue & "'"
   TxtRecoveryCustomer.Text = CN.Execute(sSql).Fields(0).Value

   sSql = " Select isnull(sum(Payment),0) as Amount " & vbCrLf _
      + " FROM CustomOrderDelivery" & vbCrLf _
      + " where Cash=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and DeliveryDate = '" & DtpEntryDate.DateValue & "'"
   TxtRecoveryCustomer.Text = Val(TxtRecoveryCustomer.Text) + CN.Execute(sSql).Fields(0).Value

   ' Step 4 - Bank Card Sale
   sSql = " Select isnull(Sum(Amount + isnull(ServiceCharges,0) + isnull(STax,0) - isnull(CashReceived,0) - isnull(BillDisc,0)),0)  as TotalBankSale" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(Amount) as Amount From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where BankCard = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
   TxtBankCardSale.Text = CN.Execute(sSql).Fields(0).Value

   sSql = " Select isnull(Sum(Amount - isnull(BillDisc,0) + isnull(ServiceCharges,0) + isnull(STax,0) ),0)  as TotalBankSale" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) as Amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where BankCard = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'"
   TxtBankCardSale.Text = Val(TxtBankCardSale.Text) - CN.Execute(sSql).Fields(0).Value

   sSql = " Select isnull(Sum(Amount - isnull(BillDisc,0)),0)  as TotalBankSale" & vbCrLf _
      + " from ServiceHeader h inner join (select BillID, BillDate, sum(Amount) as Amount From ServiceBody Group By BillId, BillDate) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
      + " where BankCard = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
   TxtBankCardSale.Text = Val(TxtBankCardSale.Text) - CN.Execute(sSql).Fields(0).Value

   ' Step 5 - Credit Sale
   sSql = " Select isnull(round(sum(Amount - isnull(CashReceived,0) - isnull(BillDisc,0) + isnull(ServiceCharges,0) + isnull(OtherCharges,0) + isnull(STax,0) ),0),0) as CreditSale" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(Amount) as Amount From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where Credit = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
   TxtCreditSale.Text = CN.Execute(sSql).Fields(0).Value
   
    sSql = " Select isnull(sum(Amount),0) as SaleReturn" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where Credit = 1 and 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'"
   TxtCreditSale.Text = Val(TxtCreditSale.Text) - CN.Execute(sSql).Fields(0).Value
      
   sSql = " Select isnull(sum(TotalAmount-isnull(Advance,0)),0) as CreditSale" & vbCrLf _
      + " from CustomOrderHeader " & vbCrLf _
      + " where Cash = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and OrderDate = '" & DtpEntryDate.DateValue & "'"
   TxtCreditSale.Text = Val(TxtCreditSale.Text) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(TotalAmount),0) as CreditSale" & vbCrLf _
      + " from CustomOrderHeader " & vbCrLf _
      + " where Credit = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and OrderDate = '" & DtpEntryDate.DateValue & "'"
   TxtCreditSale.Text = Val(TxtCreditSale.Text) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(round(sum(Amount - isnull(CashReceived,0) - isnull(BillDisc,0) ),0),0) as CreditSale" & vbCrLf _
      + " from ServiceHeader h inner join (select BillID, BillDate, sum(Amount) as Amount From ServiceBody Group By BillId, BillDate) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
      + " where Credit = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
   TxtCreditSale.Text = Val(TxtCreditSale.Text) + CN.Execute(sSql).Fields(0).Value

   ' Step 6 - Discount
   sSql = " Select isnull(floor(isnull(sum(BillDisc),0) + isnull(sum(discval),0)),0) as Discount" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(discval)discval From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where 1 = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
   TxtDiscount.Text = CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(floor(isnull(sum(BillDisc),0) + isnull(sum(discval),0)),0) as Discount" & vbCrLf _
      + " from ServiceHeader h inner join (select BillId, BillDate, sum(discval)discval From ServiceBody Group By BillId, BillDate) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
      + " where 1 = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
   TxtDiscount.Text = Val(TxtDiscount.Text) + CN.Execute(sSql).Fields(0).Value
   
   ' Step 7 - Service Charges
   sSql = " Select isnull(sum(isnull(ServiceCharges,0)+ isnull(othercharges,0)),0)  as ServiceCharges" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(discval)discval From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where 1 = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
   TxtServiceCharges.Text = CN.Execute(sSql).Fields(0).Value
   
   ' Step 8 - Sales Tax
   sSql = " Select isnull(sum(STax),0) as STax" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(discval)discval From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where 1 = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
   TxtSTax.Text = CN.Execute(sSql).Fields(0).Value
   
   ' Step 9 - Sale Return
   sSql = " Select isnull(round(sum(Amount - isnull(BillDisc,0) + isnull(ServiceCharges,0) + isnull(STax,0) ),0),0) as SaleReturn" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where Cash = 1 and 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'"
   TxtSaleReturn.Text = CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(round(sum(Amount - isnull(BillDisc,0) + isnull(ServiceCharges,0) + isnull(STax,0)),0),0) as SaleReturn" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where BankCard = 1 and 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'"
   TxtSaleReturn.Text = Val(TxtSaleReturn.Text) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(Amount),0) as SaleReturn" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where Credit = 1 and 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'"
   TxtSaleReturn.Text = Val(TxtSaleReturn.Text) + CN.Execute(sSql).Fields(0).Value
   
   
   ' Step 10 - Total Payments
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " FROM DebitVouchers h INNER JOIN DebitVouchersBody b ON h.VoucherNo = B.VoucherNo and h.Storeid = b.Storeid" & vbCrLf _
      + " where 1=1 and BankId is null " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.VoucherDate = '" & DtpEntryDate.DateValue & "'"
   TxtPayments.Text = CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(PaidAmount),0) as Amount " & vbCrLf _
      + " FROM PurchaseHeader " & vbCrLf _
      + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and PurchaseDate = '" & DtpEntryDate.DateValue & "'"
   TxtPayments.Text = Val(TxtPayments.Text) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " from PaymentHeader h inner join PaymentVender v on h.PaymentID = v.PaymentID " & vbCrLf _
      + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and PaymentDate = '" & DtpEntryDate.DateValue & "'"
   TxtPayments.Text = Val(TxtPayments.Text) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " FROM RecoveryHeader h INNER JOIN RecoveryCustomer b ON h.RecoveryId = B.RecoveryId " & vbCrLf _
      + " where 1=1 and BankMachineId is not null " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.RecoveryDate = '" & DtpEntryDate.DateValue & "'"
   TxtPayments.Text = Val(TxtPayments.Text) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " from AdvanceVouchers h inner join AdvanceVouchersBody b on h.VoucherNo = b.VoucherNo" & vbCrLf _
      + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.VoucherDate = '" & DtpEntryDate.DateValue & "'"
   TxtPayments.Text = Val(TxtPayments.Text) + CN.Execute(sSql).Fields(0).Value
   
   ' Step 11 - Total Received Payments
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " FROM CreditVouchers h INNER JOIN CreditVouchersBody b ON h.VoucherNo = B.VoucherNo and h.Storeid = b.Storeid" & vbCrLf _
      + " where 1=1 and BankId is null " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.VoucherDate = '" & DtpEntryDate.DateValue & "'"
   TxtCashReceived.Text = CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(CashReceived),0) as CreditSale" & vbCrLf _
      + " from SaleOrderHeader h " & vbCrLf _
      + " where Credit = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.OrderDate = '" & DtpEntryDate.DateValue & "'"
   TxtCashReceived.Text = Val(TxtCashReceived.Text) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(AdvanceReceived),0) as CreditSale" & vbCrLf _
      + " from BanquetOrder h " & vbCrLf _
      + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BookingDate = '" & DtpEntryDate.DateValue & "'"
   TxtCashReceived.Text = Val(TxtCashReceived.Text) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(Received),0) as CreditSale" & vbCrLf _
      + " from BanquetInvoice h " & vbCrLf _
      + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.InvoiceDate = '" & DtpEntryDate.DateValue & "'"
   TxtCashReceived.Text = Val(TxtCashReceived.Text) + CN.Execute(sSql).Fields(0).Value
   
   
   
   ''''' cash paid on credit Sale Return
   sSql = " Select isnull(sum(cashpaid),0) as SaleReturn" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID)b on H.SID = B.SID" & vbCrLf _
      + " where Credit = 1 and 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'"
   
    TxtCreditSaleReturnPaid.Text = CN.Execute(sSql).Fields(0).Value
    
   ''''' cash paid on Bank Cart Sale Return
   sSql = " Select isnull(sum(cashpaid),0) as SaleReturn" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where BankCard = 1 and 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'"

   TxtCreditSaleReturnPaid.Text = Val(TxtCreditSaleReturnPaid.Text) + CN.Execute(sSql).Fields(0).Value
   
   LoadGrid
   LoadProductGrid
   Call SubFormula2
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubFormula2()
   TxtCashAvailable.Text = (Val(TxtTotalSale.Text) + Val(TxtRecoveryCustomer.Text) + Val(TxtCashReceived.Text) + Val(TxtPettyCash.Text) + Val(TxtServiceCharges.Text) + Val(TxtSTax.Text)) - (Val(TxtBankCardSale.Text) + Val(TxtCreditSale.Text) + Val(TxtDiscount.Text) + Val(TxtSaleReturn.Text) + Val(TxtPayments.Text))
   TxtCashAvailable.Text = Val(TxtCashAvailable.Text) - Val(TxtCreditSaleReturnPaid.Text)
   TxtExcessShort.Text = (Val(TxtTotalCash.Text) + Val(TxtAddCollection.Text)) - Val(TxtCashAvailable.Text)
End Sub

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Me.MousePointer = vbHourglass
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   TxtTotalCash.Text = "0"
   'sSQL = "Select d.Denom, isnull(Qty,0) Qty  FROM (select Denom, Sum(Qty) as Qty from UserClosingHeader where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and EntryDate = '" & DtpEntryDate.DateValue & "' and Group By Denom)h inner join UserClosingBody b on h.ID = b.ID right Outer Join Denominations d on b.Denom = d.Denom Order by d.Denom desc"
   
   sSql = "Select d.Denom, isnull(PQty,0)PQty, isnull(Qty,0) Qty  FROM (select * from UserClosingHeader where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and EntryDate = '" & DtpEntryDate.DateValue & "')h inner join UserClosingBody b on h.ID = b.ID right Outer Join Denominations d on b.Denom = d.Denom Order by d.Denom desc"
   
   
   With CN.Execute(sSql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("Denom").Value = .Fields("Denom").Value
         Grid.Columns("Mul").Text = "X"
         Grid.Columns("Equ").Text = "="
         Grid.Columns("PQty").Value = .Fields("PQty").Value
         Grid.Columns("Qty").Value = .Fields("Qty").Value
         Grid.Columns("PAmount").Value = Val(.Fields("Denom").Value) * (.Fields("PQty").Value)
         Grid.Columns("QAmount").Value = Val(.Fields("Denom").Value) * (.Fields("Qty").Value)
         Grid.Columns("Amount").Value = Val(Grid.Columns("PAmount").Value) + (Grid.Columns("QAmount").Value)
         TxtTotalCash.Text = Val(TxtTotalCash.Text) + Val(Grid.Columns("Amount").Value)
         Grid.Update
'         Grid.AddNew
'         Grid.Columns("Denom").Value = .Fields("Denom").Value
'         Grid.Columns("Mul").Text = "X"
'         Grid.Columns("Equ").Text = "="
'         Grid.Columns("Qty").Value = .Fields("Qty").Value
'         Grid.Columns("Amount").Value = Val(.Fields("Denom").Value) * (.Fields("Qty").Value)
'         TxtTotalCash.Text = Val(TxtTotalCash.Text) + Val(.Fields("Denom").Value) * Val(.Fields("Qty").Value)
'         Grid.Update
         .MoveNext
      Loop
   End With
   Grid.Redraw = True
   'Grid.MoveFirst
   'If Grid.Visible Then Grid.SetFocus
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub LoadProductGrid()
   On Error GoTo ErrorHandler
   Me.MousePointer = vbHourglass
   PGrid.Redraw = False
   PGrid.CancelUpdate
   PGrid.RemoveAll
   TxtTotalClosing.Text = "0"
   TxtTotalOpening.Text = "0"
   
   sSql = " select p.ProductID, ProductName, isnull(sum(opening),0) as Opening, isnull(sum(Sale),0) as Sale, isnull(sum([Return]),0) as [Return], isnull(sum(opening),0)-isnull(sum(Sale),0)+isnull(sum([Return]),0) as Closing" & vbCrLf _
      + " from(" & vbCrLf _
      + " select ProductID, sum(Opening) Opening, 0 as sale, 0 as [Return], 0 as Closing" & vbCrLf _
      + " from OpeningProductHeader h inner join OpeningProductBody b on h.ID = b.ID" & vbCrLf _
      + " where EntryDate = '" & DtpEntryDate.DateValue & "' " & IIf(CmbUsers.ListIndex = 0, "", " and ToUserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & vbCrLf _
      + " Group By ProductID" & vbCrLf _
      + " union all " & vbCrLf _
      + " select p.ProductID, 0 as opening, Sum(Qty) as Sale, 0 as [Return], 0 as closing" & vbCrLf _
      + " from SaleHeader h inner join SaleBody b on h.billid = b.billid and h.billdate = b.billdate " & vbCrLf _
      + " inner join (select ProductID from Products where isclosingProduct = 1)p on p.productid = b.productid" & vbCrLf _
      + " where h.billdate = '" & DtpEntryDate.DateValue & "' " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & vbCrLf _
      + " Group By p.ProductID" & vbCrLf _
      + " union all" & vbCrLf _
      + " select p.ProductID, 0 as Opening, 0 as Sale, Sum(Qty) as [Return], 0 as closing" & vbCrLf _
      + " from SaleReturnHeader h inner join SaleReturnBody b on h.ReturnID = b.ReturnID and h.ReturnDate = b.ReturnDate " & vbCrLf _
      + " inner join (select ProductID from Products where isclosingProduct = 1)p on p.productid = b.productid" & vbCrLf _
      + " where h.ReturnDate = '" & DtpEntryDate.DateValue & "' " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & vbCrLf _
      + " Group By p.ProductID" & vbCrLf _
      + " )d right outer join (select * from Products where isclosingProduct = 1) p on p.ProductID = d.ProductID" & vbCrLf _
      + " group by p.ProductID, ProductName"
     
   With CN.Execute(sSql)
      Do Until .EOF
         PGrid.AddNew
         PGrid.Columns("ProductID").Value = .Fields("ProductID").Value
         PGrid.Columns("ProductName").Value = .Fields("ProductName").Value
         PGrid.Columns("Opening").Value = .Fields("Opening").Value
         PGrid.Columns("Sale").Value = .Fields("Sale").Value
         PGrid.Columns("Return").Value = .Fields("Return").Value
         PGrid.Columns("Closing").Value = .Fields("Closing").Value
         TxtTotalOpening.Text = Val(TxtTotalOpening.Text) + Val(PGrid.Columns("Opening").Value)
         TxtTotalClosing.Text = Val(TxtTotalClosing.Text) + Val(PGrid.Columns("Closing").Value)
         PGrid.Update
         .MoveNext
      Loop
   End With
   PGrid.Redraw = True
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   PGrid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   Call SubClearFields
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
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniAdminClosing", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
   Dim vtbl As String
   
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
   CN.BeginTrans
   
   Call BinData
   Call ActivityLogBin("", eFrmAdminClosing, eDelete, TxtID.Text, DtpEntryDate.DateValue, " Admin Closing Deleted Total Sale: " & Val(TxtTotalSale.Text))
   
   Call SetNonPost(CmbUsers.ItemData(CmbUsers.ListIndex), DtpEntryDate.DateValue)
   CN.Execute "Delete from AdminClosing where ID = " & Val(TxtID.Text)
   
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnEmailCDO_Click()
Dim FileAttach As New Collection
Dim vSubject, vBody As String
vSubject = "Admin Closing-" & CmbUsers.Text & "-" & TxtID.Text & "-" & Format(DtpEntryDate.DateValue, "yyyymmdd")
vBody = "Please Find Attached File"
FileAttach.Add vTmp & "\PDF\AdminClosing-" & CmbUsers.Text & "-" & TxtID.Text & "-" & Format(DtpEntryDate.DateValue, "yyyymmdd") & ".pdf"
FileAttach.Add vTmp & "\PDF\UserActivityLog-" & CmbUsers.Text & "-" & TxtID.Text & "-" & Format(DtpEntryDate.DateValue, "yyyymmdd") & ".pdf"
Call SendEmail(ObjRegistry.FromEmail, ObjRegistry.EmailPwd, ObjRegistry.ToEmail, vSubject, vBody, , , FileAttach)
End Sub

Private Sub BtnExportActivityPDF_Click()

vStrSQL = "Select ALB.*, UserName, ActionName, FormType  " & vbCrLf _
      + " from " & vBinDataBase & ".dbo.ActivityLogBin ALB " & vbCrLf _
      + " Inner join  " & vBinDataBase & ".dbo.FormBin FB on FB.FormNo = ALB.FormNo " & vbCrLf _
      + " Inner join  " & vBinDataBase & ".dbo.ActionBIN AB on AB.ActionNo = ALB.ActionNo " & vbCrLf _
      + " inner join Users U on U.UserNo = ALB.UserNo " & vbCrLf _
      + " where U.UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex) & vbCrLf _
      + " And CONVERT(varchar,ALB.ActivityDate ,101) = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
      + " And AB.ActionNo in (" & ObjRegistry.ActivityActionNo & ")" & vbCrLf _
      + " Order by ActivityID Desc"

   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
      
   Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpUserActivityLog.rpt")
'   Set RptReportViewer.Report = New CrpClosing
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy  'CN.Execute("Select Name from Manufacturer").Fields(0).Value
   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
   
   RptReportViewer.Report.PaperOrientation = crPortrait
   
   ''''''' Export Report as PDF '''''''''''''''''''''
   RptReportViewer.Report.ExportOptions.DiskFileName = vTmp & "\PDF\UserActivityLog-" & CmbUsers.Text & "-" & TxtID.Text & "-" & Format(DtpEntryDate.DateValue, "yyyymmdd") & ".pdf"
   RptReportViewer.Report.ExportOptions.DestinationType = crEDTDiskFile
   RptReportViewer.Report.ExportOptions.FormatType = crEFTPortableDocFormat
   RptReportViewer.Report.Export False
   ''''''''''''''''''''''''''''''''''''''''''''''''''
   
   'RptReportViewer.Report.PrintOut False
'   RptReportViewer.Show
End Sub

Private Sub BtnExportPDF_Click()

vStrSQL = "Select a.EntryDate, TotalSale, PettyCash, RecoveryCustomer, BankCardSale, CreditSale, CashReceived, Discount, ServiceCharges, STax, SaleReturn, a.TotalCash, AddCollection, Payment," & vbCrLf _
      + " Denom, Qty, u.UserName as admin, u1.UserName as ClosingName, StoreName " & vbCrLf _
      + " from AdminClosing a " & vbCrLf _
      + " left outer join UserClosingHeader h on h.UserNo = a.ToUserNo and h.EntryDate = a.EntryDate" & vbCrLf _
      + " left outer join UserClosingBody b on h.ID = b.ID " & vbCrLf _
      + " inner join users u on u.UserNo = a.UserNo" & vbCrLf _
      + " inner join users u1 on u1.userno = a.ToUserNo" & vbCrLf _
      + " inner join Stores st on st.StoreID = a.StoreID" & vbCrLf _
      + " where a.ID = " & Val(TxtID.Text) & " AND a.StoreID = " & Val(TxtStoreID.Text) & vbCrLf _
      + " Order by Denom desc"

   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
      
   Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpClosing.rpt")
'   Set RptReportViewer.Report = New CrpClosing
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy  'CN.Execute("Select Name from Manufacturer").Fields(0).Value
   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
   
   RptReportViewer.Report.PaperOrientation = crPortrait
   
   ''''''' Export Report as PDF '''''''''''''''''''''
   RptReportViewer.Report.ExportOptions.DiskFileName = vTmp & "\PDF\AdminClosing-" & CmbUsers.Text & "-" & TxtID.Text & "-" & Format(DtpEntryDate.DateValue, "yyyymmdd") & ".pdf"
   RptReportViewer.Report.ExportOptions.DestinationType = crEDTDiskFile
   RptReportViewer.Report.ExportOptions.FormatType = crEFTPortableDocFormat
   RptReportViewer.Report.Export False
   ''''''''''''''''''''''''''''''''''''''''''''''''''
   
   'RptReportViewer.Report.PrintOut False
'   RptReportViewer.Show
   

End Sub

Private Sub BtnFrameClose_Click()
   FrameDetail.Visible = False
End Sub

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchAdminClosing.Show vbModal, Me
   If SchAdminClosing.ParaOutID <> 0 Then GetAdminClosing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOpeningProduct_Click()
   On Error GoTo ErrorHandler
   FrmOpeningProducts.Grid.Redraw = False
   If FrmOpeningProducts.RsBody.State = adStateOpen Then
      FrmOpeningProducts.RsBody.CancelBatch
      FrmOpeningProducts.RsBody.Close
   End If
   
   vStrSQL = "Select * From OpeningProductBody where ID = " & Val(FrmOpeningProducts.TxtID.Text)

   FrmOpeningProducts.RsBody.Open vStrSQL, CN, adOpenStatic, adLockBatchOptimistic
   FrmOpeningProducts.Grid.CancelUpdate
   FrmOpeningProducts.Grid.RemoveAll
   'FrmOpeningProducts.vSuppressUpdateEvent = True
   FrmOpeningProducts.TxtTotal.Text = "0"
   PGrid.MoveFirst
   With PGrid
      For i = 1 To PGrid.Rows
         If Val(.Columns("Closing").Value) > 0 Then
            FrmOpeningProducts.RsBody.AddNew
            FrmOpeningProducts.RsBody!ID = Val(FrmOpeningProducts.TxtID.Text)
            FrmOpeningProducts.RsBody!ProductID = .Columns("ProductID").Text
            FrmOpeningProducts.RsBody!Opening = Val(.Columns("Closing").Value)
            FrmOpeningProducts.RsBody.Update
         End If
         FrmOpeningProducts.Grid.AddNew
         FrmOpeningProducts.Grid.Columns("ProductID").Text = .Columns("ProductID").Text
         FrmOpeningProducts.Grid.Columns("ProductName").Text = .Columns("ProductName").Text
         FrmOpeningProducts.Grid.Columns("Opening").Value = .Columns("Closing").Value
         FrmOpeningProducts.TxtTotal.Text = Val(FrmOpeningProducts.TxtTotal.Text) + Val(.Columns("Closing").Value)
         FrmOpeningProducts.Grid.Update
         .MoveNext
      Next i
   End With
   'FrmOpeningProducts.vSuppressUpdateEvent = False
   FrmOpeningProducts.Grid.Redraw = True
   FrmOpeningProducts.Grid.MoveFirst
   FrmOpeningProducts.Grid.FirstRow = 0
   FrmOpeningProducts.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPettyCash_Click()
   On Error GoTo ErrorHandler
   FrmPettyCash.Grid.Redraw = False
   If FrmPettyCash.RsBody.State = adStateOpen Then
      FrmPettyCash.RsBody.CancelBatch
      FrmPettyCash.RsBody.Close
   End If
   
   vStrSQL = "Select * From PettyCashBody where ID = " & Val(FrmPettyCash.TxtID.Text)

   FrmPettyCash.RsBody.Open vStrSQL, CN, adOpenStatic, adLockBatchOptimistic
   FrmPettyCash.Grid.CancelUpdate
   FrmPettyCash.Grid.RemoveAll
   FrmPettyCash.TxtAmount.Text = "0"
   'FrmPettyCash.vSuppressUpdateEvent = True
   Grid.MoveFirst
   With Grid
      For i = 1 To Grid.Rows
         If Val(.Columns("PQty").Value) > 0 Then
            FrmPettyCash.RsBody.AddNew
            FrmPettyCash.RsBody!ID = Val(FrmPettyCash.TxtID.Text)
            FrmPettyCash.RsBody!Denom = Val(.Columns("Denom").Value)
            FrmPettyCash.RsBody!Qty = Val(.Columns("PQty").Value)
            FrmPettyCash.RsBody.Update
         End If
         FrmPettyCash.Grid.AddNew
         FrmPettyCash.Grid.Columns("Denom").Value = .Columns("Denom").Value
         FrmPettyCash.Grid.Columns("Mul").Text = "X"
         FrmPettyCash.Grid.Columns("Equ").Text = "="
         FrmPettyCash.Grid.Columns("Qty").Value = .Columns("PQty").Value
         FrmPettyCash.Grid.Columns("Amount").Value = Val(.Columns("Denom").Value) * (.Columns("PQty").Value)
         FrmPettyCash.TxtAmount.Text = Val(FrmPettyCash.TxtAmount.Text) + Val(FrmPettyCash.Grid.Columns("Amount").Value)
         FrmPettyCash.Grid.Update
         .MoveNext
      Next i
   End With
   'FrmPettyCash.vSuppressUpdateEvent = False
   FrmPettyCash.Grid.Redraw = True
   FrmPettyCash.Grid.MoveFirst
   FrmPettyCash.Grid.FirstRow = 0
   FrmPettyCash.Show
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   If ObjRegistry.UseEmail = True Then BtnEmail_Click
   
      vStrSQL = "Select a.EntryDate, TotalSale, PettyCash, RecoveryCustomer, BankCardSale, CreditSale, CashReceived, Discount, ServiceCharges, STax, SaleReturn, a.TotalCash, AddCollection, Payment," & vbCrLf _
      + " Denom, Qty, u.UserName as admin, u1.UserName as ClosingName, StoreName " & vbCrLf _
      + " from AdminClosing a " & vbCrLf _
      + " left outer join UserClosingHeader h on h.UserNo = a.ToUserNo and h.EntryDate = a.EntryDate" & vbCrLf _
      + " left outer join UserClosingBody b on h.ID = b.ID " & vbCrLf _
      + " inner join users u on u.UserNo = a.UserNo" & vbCrLf _
      + " inner join users u1 on u1.userno = a.ToUserNo" & vbCrLf _
      + " inner join Stores st on st.StoreID = a.StoreID" & vbCrLf _
      + " where a.ID = " & Val(TxtID.Text) & " AND a.StoreID = " & Val(TxtStoreID.Text) & vbCrLf _
      + " Order by Denom desc"

   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
      
   Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpClosing.rpt")
'   Set RptReportViewer.Report = New CrpClosing
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy  'CN.Execute("Select Name from Manufacturer").Fields(0).Value
   RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
   
   RptReportViewer.Report.PaperOrientation = crPortrait
   'RptReportViewer.Report.PrintOut False
   RptReportViewer.Show
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetAdminClosing()
   On Error GoTo ErrorHandler
   'sSQL = "select * from AdminClosing"
   'If Rs.State = adStateOpen Then Rs.Close
   'Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
   sSql = "Select p.*, UserName, StoreName from AdminClosing p inner join users u on p.ToUserNo = u.Userno left outer join Stores s on p.StoreID = s.SToreID where ID = " & SchAdminClosing.ParaOutID & " and p.StoreID = " & TxtStoreID.Text
   With CN.Execute(sSql)
      If Not .BOF Then
          TxtID.Text = !ID
          DtpEntryDate.DateValue = !EntryDate
          TxtStoreID.Text = IIf(IsNull(!StoreID) = False, !StoreID, "")
          TxtStoreName.Text = IIf(IsNull(!StoreName) = False, !StoreName, "")
          TxtTotalSale.Text = !TotalSale
          TxtPettyCash.Text = !PettyCash
          TxtBankCardSale.Text = !BankCardSale
          TxtCreditSale.Text = !CreditSale
          TxtRecoveryCustomer.Text = !RecoveryCustomer
          TxtDiscount.Text = !Discount
          TxtSaleReturn.Text = !SaleReturn
          TxtPayments.Text = !Payment
          TxtCashReceived.Text = IIf(IsNull(!CashReceived), "", !CashReceived)
          TxtServiceCharges.Text = !ServiceCharges
          TxtSTax.Text = !STax
          TxtTotalCash.Text = !TotalCash
          TxtAddCollection.Text = !AddCollection
          TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
          TxtAdminClssingFinePerOnShort.Text = IIf(IsNull(!AdminClssingFinePerOnShort), "", !AdminClssingFinePerOnShort)
          TxtUserFineonShort.Text = IIf(IsNull(!UserFineonShort), "", !UserFineonShort)
          TxtCompanyFineonShort.Text = IIf(IsNull(!CompanyFineonShort), "", !CompanyFineonShort)
          CmbUsers.Text = !UserName
      End If
      .Close
   End With
   FormStatus = OpenMode
   LoadGrid
   LoadProductGrid
   Call SubFormula
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSetSystemDate_Click()
   FrmSystemDate.Show
End Sub

Private Sub CmbUsers_Click()
   If CmbUsers.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbUsers.Name Then Exit Sub
   SubInitialize
End Sub

Private Sub SubInitialize()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False Then Exit Sub
   If CN.Execute("select * from AdminClosing where ToUserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex) & " and EntryDate = '" & DtpEntryDate.DateValue & "' and StoreID = " & TxtStoreID.Text).RecordCount > 0 Then
      MsgBox "This User on that Date have Already Closing. Please specify other.", vbExclamation, "Alert"
      Exit Sub
   End If
   SubCalculate
   If Val(ObjRegistry.AdminClssingFinePerOnShort) > 0 Then
      TxtAdminClssingFinePerOnShort.Text = Val(ObjRegistry.AdminClssingFinePerOnShort)
      TxtUserFineonShort.Text = Round(-1 * TxtExcessShort.Text * TxtAdminClssingFinePerOnShort.Text / 100, 2)
      TxtCompanyFineonShort.Text = -1 * TxtExcessShort.Text - TxtUserFineonShort.Text
   End If
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnEmail_Click()
'Dim oSmtp As New EASendMailObjLib.Mail
'    oSmtp.LicenseCode = "TryIt"
'       ' Set your sender email address
'    oSmtp.FromAddr = ObjRegistry.FromEmail
'
'    ' Add recipient email address
'    oSmtp.AddRecipientEx ObjRegistry.ToEmail, 0
'
'    ' Set email subject
'    oSmtp.Subject = "Admin Closing-" & CmbUsers.Text & "-" & TxtID.Text & "-" & Format(DtpEntryDate.DateValue, "yyyymmdd") & ".pdf"
'
'    ' Set HTML body format
'    oSmtp.BodyFormat = 1
'
'    ' Set HTML body text
'    oSmtp.BodyText = "<font size=5>This is Report of </font> <font color=red><b>Admin Closing</b></font>"
'
'    ' Add attachment from local disk
'    If oSmtp.AddAttachment(vTmp & "\PDF\AdminClosing-" & CmbUsers.Text & "-" & TxtID.Text & "-" & Format(DtpEntryDate.DateValue, "yyyymmdd") & ".pdf") <> 0 Then
'        MsgBox "Failed to add attachment with error:" & oSmtp.GetLastErrDescription()
'    End If
'
'    ' Add attachment from remote website
''    If oSmtp.AddAttachment("http://www.emailarchitect.net/webapp/img/logo.jpg") <> 0 Then
''        MsgBox "Failed to add attachment with error:" & oSmtp.GetLastErrDescription()
''    End If
'
'    ' Your SMTP server address
'    oSmtp.ServerAddr = ObjRegistry.SMTPServerAddress  '"smtp.live.com"
''    oSmtp.ServerAddr = "smtp.live.com"
'
'
'    ' ConnectTryTLS means if server supports SSL/TLS connection, SSL/TLS is used automatically
'    'oSmtp.ConnectType = ConnectTryTLS
'
'    ' If your server uses 587 port
''     oSmtp.ServerPort = 587
'
'    ' User and password for ESMTP authentication, if your server doesn't require
'    ' User authentication, please remove the following codes.
'    oSmtp.UserName = ObjRegistry.FromEmail
'    oSmtp.Password = ObjRegistry.EmailPwd
'
'
'    ' If your server uses 25/587/465 port with SSL/TLS
''     oSmtp.ConnectType = ConnectSSLAuto
'     oSmtp.ServerPort = Val(ObjRegistry.PortNo)  ' 25 or 587 or 465
'
'
'    oSmtp.SSL_starttls = 0
'    oSmtp.SSL_init
'
''    MsgBox "start to send email ..."
'
'    If oSmtp.SendMail() = 0 Then
'        MsgBox "email was sent successfully!"
'    Else
'        MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
'    End If
End Sub

Private Sub DtpEntryDate_Validate(Cancel As Boolean)
   SubInitialize
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Admin Closing"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   Frame1.Top = 133
   Frame1.Left = 109
  
   TxtStoreID.Text = ObjRegistry.StoreID
   With CN.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), ObjRegistry.StoreID, !StoreID)
      End If
      .Close
   End With
   FunSelectStore ssValidate, True
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   BtnSetSystemDate.Visible = ObjRegistry.SystemDate
   
   
   If Val(ObjRegistry.AdminClssingFinePerOnShort) > 0 Then
      LblFinePerOnShort.Visible = True
      TxtAdminClssingFinePerOnShort.Visible = True
      LblUserFineonShort.Visible = True
      TxtUserFineonShort.Visible = True
      LblCompanyFineonShort.Visible = True
      TxtCompanyFineonShort.Visible = True
   Else
      LblFinePerOnShort.Visible = False
      TxtAdminClssingFinePerOnShort.Visible = False
      LblUserFineonShort.Visible = False
      TxtUserFineonShort.Visible = False
      LblCompanyFineonShort.Visible = False
      TxtCompanyFineonShort.Visible = False
   End If
   
   With CN.Execute("Select * FROM Users")
      CmbUsers.Clear
      CmbUsers.AddItem "All Users"
      Do Until .EOF
         CmbUsers.AddItem !UserName
         CmbUsers.ItemData(CmbUsers.NewIndex) = !UserNo
         .MoveNext
      Loop
   End With
   CmbUsers.ListIndex = 0
   FormStatus = NewMode
   
   CmbUsers.Text = CN.Execute("Select username from users where userno = " & vUser).Fields(0)
   Set RsDetail = New ADODB.Recordset
   GridDetail.Columns("ID").DataField = "ID"
   GridDetail.Columns("Date").DataField = "Date"
   GridDetail.Columns("Date").DataField = "Party"
   GridDetail.Columns("Date").DataField = "Amount"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtAddCollection.SetFocus Else TxtStoreID.SetFocus
      End Select
   ElseIf KeyCode = vbKeyEscape Then
      Frame1.Visible = False
   ElseIf KeyCode = vbKeyF12 Then
      Frame1.ZOrder 0
      Frame1.Visible = True
      PGrid.Row = 0
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
         Case vbKeyO
            If BtnOpen.Enabled Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
      End Select
   Else
      If UCase(Me.ActiveControl.Name) Like "TXT*" Or UCase(Me.ActiveControl.Name) Like "DTP*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniAdminClosing", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If FunValidation = False Then Exit Sub
   CN.BeginTrans
   
      
   Set Rs = New ADODB.Recordset
   sSql = "Select * FROM AdminClosing where ID = " & Val(TxtID.Text) & " and StoreID = " & Val(TxtStoreID.Text)
   Rs.Open sSql, CN, adOpenStatic, adLockOptimistic
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmAdminClosing, eEdit, TxtID.Text, DtpEntryDate.DateValue, "Effected User Code-" & Rs!ToUserNo & "  Total Sale-" & Val(Rs!TotalSale))
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmAdminClosing, eEdit, TxtID.Text, DtpEntryDate.DateValue, "Updated User Code-" & CmbUsers.ItemData(CmbUsers.ListIndex) & " Total Sale-" & Val(TxtTotalSale.Text))
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmAdminClosing, eAdd, TxtID.Text, DtpEntryDate.DateValue, "Saved New User Code -" & CmbUsers.ItemData(CmbUsers.ListIndex) & " Total Sale-" & Val(TxtTotalSale.Text))
   If vIsNewRecord Then
      Rs.AddNew
      Rs!ID = TxtID.Text
      Rs!StoreID = TxtStoreID.Text
      Rs!UserNo = vUser
   End If
   Rs!isTransfer = 0
   Rs!EntryDate = DtpEntryDate.DateValue
   Rs!TotalSale = Val(TxtTotalSale.Text)
   Rs!PettyCash = Val(TxtPettyCash.Text)
   Rs!RecoveryCustomer = Val(TxtRecoveryCustomer.Text)
   Rs!BankCardSale = Val(TxtBankCardSale.Text)
   Rs!CreditSale = Val(TxtCreditSale.Text)
   Rs!Discount = Val(TxtDiscount.Text)
   Rs!SaleReturn = Val(TxtSaleReturn.Text)
   Rs!Payment = Val(TxtPayments.Text)
   Rs!CashReceived = Val(TxtCashReceived.Text)
   Rs!ServiceCharges = Val(TxtServiceCharges.Text)
   Rs!STax = Val(TxtSTax.Text)
   Rs!TotalCash = Val(TxtTotalCash.Text)
   Rs!AddCollection = Val(TxtAddCollection.Text)
   Rs!Tag = IIf(Trim(TxtTag.Text) = "", Null, TxtTag.Text)
   Rs!ToUserNo = CmbUsers.ItemData(CmbUsers.ListIndex)
   Rs!Excess = IIf(Trim(TxtExcessShort.Text) > 0, Val(TxtExcessShort.Text), Null)
   Rs!Short = IIf(Trim(TxtExcessShort.Text) < 0, Val(TxtExcessShort.Text), Null)
   Rs!AdminClssingFinePerOnShort = Val(TxtAdminClssingFinePerOnShort.Text)
   Rs!UserFineonShort = Val(TxtUserFineonShort.Text)
   Rs!CompanyFineonShort = Val(TxtCompanyFineonShort.Text)
'   Rs!UserNo = vUser
   Rs.Update
   
   Call SetPost(CmbUsers.ItemData(CmbUsers.ListIndex), DtpEntryDate.DateValue)
   
   CN.CommitTrans
   Me.MousePointer = vbHourglass
   If ObjRegistry.ExportReportASPDF = True Then
      BtnExportPDF_Click
      BtnExportActivityPDF_Click
   End If
   If ObjRegistry.UseEmail = True Then BtnEmailCDO_Click
   FormStatus = NewMode
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
   On Error GoTo ErrorHandler
   FunGetMaxID = CN.Execute("Select isnull(max(ID),0) + 1 from AdminClosing where StoreID = " & Val(TxtStoreID.Text)).Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   If Trim(TxtID.Text) = "" Then
     MsgBox "Please specify ID", vbExclamation, "Alert"
     If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
     Exit Function
   End If
   If Trim(TxtStoreID.Text) = "" Then
     MsgBox "Please specify Store ID", vbExclamation, "Alert"
     If TxtStoreID.Enabled And TxtStoreID.Visible Then TxtStoreID.SetFocus
     Exit Function
   End If
   If vIsNewRecord = True Then
      If CN.Execute("select * from AdminClosing where ToUserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex) & " and EntryDate = '" & DtpEntryDate.DateValue & "' and StoreID = " & TxtStoreID.Text).RecordCount > 0 Then
         MsgBox "This User Has Already Admin Closing on this Date. Please specify the other User or Date.", vbExclamation, "Alert"
         If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      End If
'      If cn.Execute("Select * from AdminClosing where ID = " & Val(TxtID.Text) & " and StoreID = " & Val(TxtStoreID.Text)).RecordCount > 0 Then
'         TxtID.Text = FunGetMaxID
'      End If
   End If

'   If vIsNewRecord = True Then
'      Rs.Filter = " EntryDate = '" & DtpEntryDate.DateValue & "' and ToUserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)
'      If Rs.RecordCount <> 0 Then
'          MsgBox "This User Has ALready Petty Cash. Please specify the other User.", vbExclamation, "Alert"
'          CmbUsers.SetFocus
'          Exit Function
'      End If
'   End If
  'All Ok, now validation is success
   FunValidation = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Property Get FormStatus() As FormMode
   On Error GoTo ErrorHandler
   'Nothing
   FormStatus = vMode
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
   'Based upon the value of vNewValue, we shall decide what controls to enable/disable
   On Error GoTo ErrorHandler
   vMode = vNewValue
   Select Case vNewValue
   Case Is = NewMode
      Call SubClearFields
      'TxtID.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      Frame1.Visible = False
      DtpEntryDate.DateValue = IIf(Format(Now, "hh") > 3, Date, DateAdd("d", -1, Date))
      TxtID.Text = FunGetMaxID
      vIsNewRecord = True
      TxtStoreID.Enabled = True
      BtnStore.Enabled = True
      If DtpEntryDate.Visible And DtpEntryDate.Enabled Then DtpEntryDate.SetFocus
   Case Is = OpenMode
      'TxtID.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      TxtStoreID.Enabled = False
      BtnStore.Enabled = False
      DtpEntryDate.SetFocus
      vIsNewRecord = False
   Case Is = ChangeMode
      BtnOpen.Enabled = False
      BtnPrint.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo ErrorHandler
   Set RptReportViewer.Report = Nothing
   Set FrmAdminClosing = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then
            ctl.Text = ""
         End If
      ElseIf TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then
            ctl.Text = ""
         End If
      End If
   Next
   GridDetail.CancelUpdate
   GridDetail.RemoveAll
   GridDetail.Columns("Party").Visible = True
   GridDetail.Update
   FrameDetail.Visible = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub



Private Sub LblSaleReturn_Click()
GridDetail.Columns("Party").Visible = True
FrameDetail.Caption = LblSaleReturn.Caption
sSql = " Select 'SR' as type, h.ReturnId as ID, h.ReturnDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Round(Amount - isnull(BillDisc, 0) + isnull(ServiceCharges, 0) + isnull(STax, 0), 0), 0) As Amount " & vbCrLf _
 + " from SaleReturnHeader h inner join (select ReturnId, ReturnDate,  amount From SaleReturnBody) b on h.ReturnId = b.ReturnId and h.ReturnDate = b.ReturnDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " Where Cash = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All" & vbCrLf _
 + " Select 'SR' as type, h.ReturnId as ID, h.ReturnDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Round(Amount - isnull(BillDisc, 0) + isnull(ServiceCharges, 0) + isnull(STax, 0), 0), 0) As Amount " & vbCrLf _
 + " from SaleReturnHeader h inner join (select ReturnId, ReturnDate, amount From SaleReturnBody ) b on h.ReturnId = b.ReturnId and h.ReturnDate = b.ReturnDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " Where BankCard = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'SR' as type, h.ReturnId as ID, h.ReturnDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount, 0) As Amount " & vbCrLf _
 + " from SaleReturnHeader h inner join (select ReturnId, ReturnDate, amount From SaleReturnBody) b on h.ReturnId = b.ReturnId and h.ReturnDate = b.ReturnDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " Where Credit = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'"
LoadGridDetail
End Sub

Private Sub LblCashAvailable_Click()
GridDetail.Columns("Party").Visible = True
FrameDetail.Caption = LblCashAvailable.Caption
sSql = " Select 'Cr ' as type, h.voucherNo as ID, h.VoucherDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount, 0) As Amount " & vbCrLf _
 + " FROM CreditVouchers h INNER JOIN CreditVouchersBody b ON h.VoucherNo = B.VoucherNo " & vbCrLf _
 + " inner join parties p on p.partyid = b.accountno " & vbCrLf _
 + " where 1=1 and BankId is null " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.VoucherDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'SO ' as type, h.OrderID as ID, h.OrderDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(CashReceived, 0) As Amount " & vbCrLf _
 + " from SaleOrderHeader h " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where Credit = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.OrderDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'BO ' as type, h.BookingID as ID, h.BookingDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(AdvanceReceived, 0) As Amount " & vbCrLf _
 + " from BanquetOrder h " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BookingDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'BI ' as type, h.BookingID as ID, h.InvoiceDate as Date, '' as party, " & vbCrLf _
 + " isnull(Received, 0) As Amount " & vbCrLf _
 + " from BanquetInvoice h " & vbCrLf _
 + " where 1=1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.InvoiceDate = '" & DtpEntryDate.DateValue & "'"

LoadGridDetail
End Sub

Private Sub LblRecoveryCustomer_Click()
GridDetail.Columns("Party").Visible = True
FrameDetail.Caption = LblRecoveryCustomer.Caption
 sSql = "Select 'RI' as type, h.recoveryId as ID, h.RecoveryDate as Date, ''  as party,  isnull(Amount,0) as Amount " & vbCrLf _
 + " FROM RecoveryHeader h INNER JOIN RecoveryCustomer b ON h.RecoveryId = B.RecoveryId " & vbCrLf _
 + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.RecoveryDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'CD' as type, h.DeliveryID as ID, h.DeliveryDate as Date, Partyname as party, isnull(Payment,0) as Amount " & vbCrLf _
 + " FROM CustomOrderDelivery h " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where Cash=1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and DeliveryDate = '" & DtpEntryDate.DateValue & "'"
 LoadGridDetail
End Sub

Private Sub LblCashReceived_Click()
GridDetail.Columns("Party").Visible = True
FrameDetail.Caption = LblCashReceived.Caption
sSql = " Select 'SI' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount + isnull(ServiceCharges, 0) + isnull(STax, 0) - isnull(CashReceived, 0) - isnull(BillDisc, 0), 0) As TotalBankSale " & vbCrLf _
 + " from SaleHeader h inner join (select BillID, BillDate, Amount From SaleBody ) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where BankCard = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select'SR' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount - isnull(BillDisc, 0) + isnull(ServiceCharges, 0) + isnull(STax, 0), 0) As TotalBankSale " & vbCrLf _
 + " from SaleReturnHeader h inner join (select ReturnID, ReturnDate, Amount From SaleReturnBody ) b on h.ReturnID = b.ReturnID and h.ReturnDate = b.ReturnDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where BankCard = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'SV' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount - isnull(BillDisc, 0), 0) As TotalBankSale " & vbCrLf _
 + " from ServiceHeader h inner join (select BillID, BillDate, Amount as Amount From ServiceBody ) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where BankCard = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
 LoadGridDetail
End Sub

Private Sub LblServiceCh_Click()
GridDetail.Columns("Party").Visible = True
FrameDetail.Caption = LblServiceCh.Caption
sSql = "Select 'SI Charges' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(isnull(ServiceCharges, 0) + isnull(othercharges, 0), 0) As Amount " & vbCrLf _
 + " from SaleHeader h inner join (select SId, BillDate, Sum(discval) discval From SaleBody Group by SId, BillDate ) b on h.SID = b.SID and h.BillDate = b.BillDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " Where 1 = 1 and isnull(isnull(ServiceCharges, 0) + isnull(othercharges, 0), 0) <> 0 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
LoadGridDetail
End Sub

Private Sub LblSalesTax_Click()
GridDetail.Columns("Party").Visible = True
FrameDetail.Caption = LblSalesTax.Caption
sSql = "Select 'SI Tax' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(STax, 0) As Amount " & vbCrLf _
 + " from SaleHeader h inner join (select SId, BillDate, Sum(discval) discval From SaleBody Group by SId, BillDate ) b on h.SID = b.SID and h.BillDate = b.BillDate   " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where 1 = 1 and isnull(STax, 0) <> 0 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
LoadGridDetail
End Sub

Private Sub LblTotalSale_Click()
   GridDetail.Columns("Party").Visible = True
   FrameDetail.Caption = LblTotalSale.Caption
   sSql = " Select 'SI' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
   + " isnull(round((isnull(multiplier,0)* isnull(QtyPack,0) + Qty )*((Price/isnull(multiplier,1))+isnull(sc,0)),0),0) as Amount " & vbCrLf _
   + " from SaleHeader h inner join SaleBody b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
   + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
   + " where 1=1  " & IIf(CmbUsers.ListIndex = 0, "", " and ToUserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.billDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
   + " Union All " & vbCrLf _
   + " Select 'CO' as type, h.orderid as ID, h.orderdate as Date, Partyname as party, isnull(totalamount,0) as Amount " & vbCrLf _
   + " from CustomOrderHeader h " & vbCrLf _
   + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
   + " where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and OrderDate = '" & DtpEntryDate.DateValue & "'"
   
   LoadGridDetail
End Sub

Private Sub LblPattyCash_Click()
   GridDetail.Columns("Party").Visible = False
   FrameDetail.Caption = LblPattyCash.Caption
   sSql = " Select ID, EntryDate as Date, '' party, Amount from PettyCashHeader where 1=1 " & IIf(CmbUsers.ListIndex = 0, "", " and ToUserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and EntryDate = '" & DtpEntryDate.DateValue & "'"
   LoadGridDetail
End Sub

Private Sub LblPayments_Click()
GridDetail.Columns("Party").Visible = True
FrameDetail.Caption = LblPayments.Caption
sSql = "Select 'Dr ' as type, h.voucherNo as ID, h.VoucherDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount, 0) As Amount " & vbCrLf _
 + " FROM DebitVouchers h INNER JOIN DebitVouchersBody b ON h.VoucherNo = B.VoucherNo " & vbCrLf _
 + " inner join parties p on p.partyid = b.accountno " & vbCrLf _
 + " where 1=1 and BankId is null " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.VoucherDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'PI ' as type, h.PurID as ID, h.purchaseDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(PaidAmount, 0) As Amount " & vbCrLf _
 + " FROM PurchaseHeader h " & vbCrLf _
 + " inner join parties p on p.partyid = h.vendorid " & vbCrLf _
 + " where 1=1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and PurchaseDate = '" & DtpEntryDate.DateValue & "'"
 sSql = sSql + "" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'Payment' as type, h.Paymentid as ID, h.PaymentDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount, 0) As Amount " & vbCrLf _
 + " from PaymentHeader h inner join PaymentVender v on h.PaymentID = v.PaymentID " & vbCrLf _
 + " inner join parties p on p.partyid = v.venderid " & vbCrLf _
 + " where 1=1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and PaymentDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'Recovery' as type, h.RecoveryID as ID, h.RecoveryDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount, 0) As Amount " & vbCrLf _
 + " FROM RecoveryHeader h INNER JOIN RecoveryCustomer b ON h.RecoveryId = B.RecoveryId " & vbCrLf _
 + " inner join parties p on p.partyid = b.customerid " & vbCrLf _
 + " where 1=1 and BankMachineId is not null " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.RecoveryDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'SI Tax' as type, h.VoucherNo as ID, h.VoucherDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount, 0) As Amount " & vbCrLf _
 + " from AdvanceVouchers h inner join AdvanceVouchersBody b on h.VoucherNo = b.VoucherNo " & vbCrLf _
 + " inner join parties p on p.partyid = b.accountno " & vbCrLf _
 + " where 1=1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.VoucherDate = '" & DtpEntryDate.DateValue & "'"
LoadGridDetail
End Sub

Private Sub LblBankCardSale_Click()
   GridDetail.Columns("Party").Visible = True
   FrameDetail.Caption = LblBankCardSale.Caption
   sSql = "Select 'SI Bank' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount + isnull(ServiceCharges, 0) + isnull(STax, 0) - isnull(CashReceived, 0) - isnull(BillDisc, 0), 0) As TotalBankSale " & vbCrLf _
 + " from SaleHeader h inner join (select BillID, BillDate, Amount From SaleBody ) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where BankCard = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'SR Bank' as type, h.ReturnID as ID, h.ReturnDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount - isnull(BillDisc, 0) + isnull(ServiceCharges, 0) + isnull(STax, 0), 0) As TotalBankSale " & vbCrLf _
 + " from SaleReturnHeader h inner join (select ReturnID, ReturnDate, Amount From SaleReturnBody ) b on h.ReturnID = b.ReturnID and h.ReturnDate = b.ReturnDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where BankCard = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.ReturnDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'SV Bank' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Amount - isnull(BillDisc, 0), 0) As TotalBankSale " & vbCrLf _
 + " from ServiceHeader h inner join (select BillID, BillDate, Amount From ServiceBody ) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where BankCard = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
 LoadGridDetail
   
End Sub

Private Sub LblCreditSale_Click()
GridDetail.Columns("Party").Visible = True
FrameDetail.Caption = LblCreditSale.Caption
 sSql = "Select 'SI Cr' as type, h.billID as ID, h.BillDate as Date, Partyname as party," & vbCrLf _
 + " isnull(Round(Amount - isnull(CashReceived, 0) - isnull(BillDisc, 0) + isnull(ServiceCharges, 0) + isnull(OtherCharges, 0) + isnull(STax, 0), 0), 0) As Amount " & vbCrLf _
 + " from SaleHeader h inner join (select BillID, BillDate, Amount as Amount From SaleBody ) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where Credit = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'CO Cash' as type, h.OrderID as ID, h.OrderDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(TotalAmount - isnull(Advance, 0), 0) As Amount " & vbCrLf _
 + " from CustomOrderHeader h " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where Cash = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.OrderDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'CO Cr ' as type, h.OrderID as ID, h.OrderDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(TotalAmount, 0) As Amount " & vbCrLf _
 + " from CustomOrderHeader h " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where Credit = 1 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.OrderDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'SV Cr' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(Round(Amount - isnull(CashReceived, 0) - isnull(BillDisc, 0), 0), 0) As Amount " & vbCrLf _
 + " from ServiceHeader h inner join (select BillID, BillDate, Amount From ServiceBody ) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " where Credit = 1  " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
 LoadGridDetail
End Sub

Private Sub LblDiscount_Click()
GridDetail.Columns("Party").Visible = True
FrameDetail.Caption = LblDiscount.Caption
sSql = " Select 'SI Disc' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(floor(isnull(BillDisc, 0) + isnull(discval, 0)), 0) As Amount " & vbCrLf _
 + " from SaleHeader h inner join (select SId, BillDate, Sum(discval) discval From SaleBody Group by SId, BillDate ) b on h.SID = b.SID and h.BillDate = b.BillDate  " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " Where 1 = 1 and isnull(floor(isnull(BillDisc, 0) + isnull(discval, 0)), 0) <> 0 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'" & vbCrLf _
 + " Union All " & vbCrLf _
 + " Select 'SV Disc' as type, h.billID as ID, h.BillDate as Date, Partyname as party, " & vbCrLf _
 + " isnull(floor(isnull(BillDisc, 0) + isnull(discval, 0)), 0) As Amount " & vbCrLf _
 + " from ServiceHeader h inner join (select BillId, BillDate, discval From ServiceBody ) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
 + " inner join parties p on p.partyid = h.customerid " & vbCrLf _
 + " Where 1 = 1 and isnull(floor(isnull(BillDisc, 0) + isnull(discval, 0)), 0) <> 0 " & IIf(CmbUsers.ListIndex = 0, "", " and UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)) & " and h.BillDate = '" & DtpEntryDate.DateValue & "'"
LoadGridDetail
End Sub

Private Sub TxtAddCollection_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtAddCollection.Name Then Exit Sub
   Call SubFormula
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
    vStrSQL = " Select * FROM Stores where StoreID=" & Val(TxtStoreID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtStoreName.Text = !StoreName
          FunSelectStore = True
          .Close
          If BtnSave.Enabled = False And BtnSave.Visible = True Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectStore = False
          .Close
          TxtStoreID.Text = ""
          TxtStoreName.Text = ""
          If BtnSave.Enabled = False And BtnSave.Visible = True Then FormStatus = ChangeMode
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
   TxtID.Text = FunGetMaxID()
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtAddCollection.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub LoadGridDetail()
On Error GoTo ErrorHandler
   If FrameDetail.Visible = False Then FrameDetail.Visible = True
   FrameDetail.Caption = FrameDetail.Caption & " Detail"
   GridDetail.Redraw = False
   GridDetail.CancelUpdate
   GridDetail.RemoveAll
   GridDetail.Update
   GridDetail.AllowAddNew = True
   With CN.Execute(sSql)
      While Not .EOF
      GridDetail.AddNew
      GridDetail.Columns("ID").Text = !ID & " - " & !Type
      GridDetail.Columns("Date").Text = !Date
      GridDetail.Columns("Party").Text = !Party
      GridDetail.Columns("Amount").Value = !Amount
      .MoveNext
      Wend
      
   End With
   GridDetail.Redraw = True
'   If RsDetail.State = adStateOpen Then RsDetail.Close
'   RsDetail.Open sSql, CN, adOpenDynamic, adLockReadOnly
'   Set GridDetail.DataSource = RsDetail
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculate()
   On Error GoTo ErrorHandler
   Call CalculateAmount(CmbUsers.ItemData(CmbUsers.ListIndex), DtpEntryDate.DateValue)
   ' Step 1 - Total Sale
   
   TxtTotalSale.Text = vTotalSale
   
   ' Step 2 - Total PettyCash
   
   TxtPettyCash.Text = vPettyCash

   ' Step 3 - Total Customer Recovery
  
   TxtRecoveryCustomer.Text = vRecoveryCustomer

   ' Step 4 - Bank Card Sale
   
   TxtBankCardSale.Text = vBankCardSale

   ' Step 5 - Credit Sale

   TxtCreditSale.Text = vCreditSale
   
   ' Step 6 - Discount
   
   TxtDiscount.Text = vDiscount
    
   
   ' Step 7 - Service Charges
   
   TxtServiceCharges.Text = vServiceCharges
   
   ' Step 8 - Sales Tax
   
   TxtSTax.Text = vSTax
   
   ' Step 9 - Sale Return
   
   TxtSaleReturn.Text = vSaleReturn
      
   
   ' Step 10 - Total Payments
   
   TxtPayments.Text = vPayments
    
    
   ' Step 11 - Total Received Payments
   
   TxtCashReceived.Text = vCashReceived
   
   ' Step 12 - Bank Payments
   
   TxtBankPayments.Text = vBankPayments
   
    ' Step 13 - Bank Received Payments
   
   TxtBankReceived.Text = vBankReceived
   
   ''''' cash paid on credit Sale Return
   
    TxtCreditSaleReturnPaid.Text = vCreditSaleReturnPaid
   
   
   LoadGrid
   LoadProductGrid
   Call SubFormula
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub SubFormula()
'   TxtCashAvailable.Text = vCashAvailable
   TxtCashAvailable.Text = (Val(TxtTotalSale.Text) + Val(TxtRecoveryCustomer.Text) + Val(TxtCashReceived.Text) + Val(TxtPettyCash.Text) + Val(TxtServiceCharges.Text) + Val(TxtSTax.Text)) - (Val(TxtBankCardSale.Text) + Val(TxtCreditSale.Text) + Val(TxtDiscount.Text) + Val(TxtSaleReturn.Text) + Val(TxtPayments.Text))
   TxtCashAvailable.Text = Val(TxtCashAvailable.Text) - Val(TxtCreditSaleReturnPaid.Text)
   TxtExcessShort.Text = (Val(TxtTotalCash.Text) + Val(TxtAddCollection.Text)) - Val(TxtCashAvailable.Text)
End Sub

 
'****************************************************************
'*  Purpose :   To Send eMail
'*
'*  Inputs  :   strRecipient(String)    Recipient comma seperated
'*              strSubject(String)      Subject
'*              strBody                  Body
'*              colAttachments          Collection of attachments
'*                                      file paths.
'*
'*  Returns :   Boolean about the sent status
'****************************************************************
Public Function SendEmail(ByVal strSender As String, _
                        ByVal strEmailPwd As String, _
                        ByVal strRecipient As String, _
                        ByVal strSubject As String, _
                        ByVal strBody As String, _
                        Optional ByVal strCc As String, _
                        Optional ByVal strBcc As String, _
                        Optional ByVal colAttachments As Collection _
                         ) As Boolean
    Dim cdoMsg As New CDO.Message
    Dim cdoConf As New CDO.Configuration
    Dim schema As String
    Dim Flds
    Dim attachment
    Dim strHTML
    
    On Error GoTo ErrTrap
    Const cdoSendUsingPort = 2
    
    'Set cdoMsg =  CreateObject("CDO.Message")
    'Set cdoConf = CreateObject("CDO.Configuration")
    
    Set Flds = cdoConf.Fields
        
    schema = "http://schemas.microsoft.com/cdo/configuration/"

    With Flds
        .Item(schema & "sendusing") = 2
        .Item(schema & "smtpserver") = ObjRegistry.SMTPServerAddress '"smtp.gmail.com"
        .Item(schema & "smtpserverport") = Val(ObjRegistry.PortNo)  ' 25 or 587 or 465
        .Item(schema & "smtpauthenticate") = 1
        .Item(schema & "sendusername") = strSender
        .Item(schema & "sendpassword") = strEmailPwd
        .Item(schema & "smtpusessl") = 1
        .Update
    End With
    
    ' Apply the settings to the message.
    With cdoMsg
        Set .Configuration = cdoConf
        .To = strRecipient
        .From = strSender
        .Subject = strSubject
        .TextBody = strBody
        If Not colAttachments Is Nothing Then
            For Each attachment In colAttachments
                .AddAttachment attachment
            Next
        End If
        If strCc <> "" Then .CC = strCc
        If strBcc <> "" Then .BCC = strBcc
        .Send
    End With
    
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set Flds = Nothing
        
    SendEmail = True
    Exit Function
ErrTrap:
Err.Raise Err.Number, "", "Error from Functions.SendEmail" & Err.Description
    SendEmail = False
End Function

Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.AdminClosingBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmAdminClosing) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmAdminClosing & ", " & vUser & "," & TableHeaderFields(eFrmAdminClosing) & " from AdminClosing " & vbCrLf _
             & "Where ID = " & TxtID.Text & " and EntryDate = '" & DtpEntryDate.DateValue & "'"
      CN.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

