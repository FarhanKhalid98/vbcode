VERSION 5.00
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmConsumptionProducts 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmConsumptionProducts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3840
      Left            =   11385
      TabIndex        =   31
      Top             =   810
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
         Height          =   3435
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   32
         Tag             =   "NC"
         Text            =   "FrmConsumptionProducts.frx":0ECA
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
         TabIndex        =   33
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6690
      TabIndex        =   16
      Top             =   7845
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
      MICON           =   "FrmConsumptionProducts.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   5363
      TabIndex        =   13
      Top             =   7845
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
      MICON           =   "FrmConsumptionProducts.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   2723
      TabIndex        =   15
      Top             =   7845
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
      MICON           =   "FrmConsumptionProducts.frx":0FE2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8010
      TabIndex        =   17
      Top             =   7845
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
      MICON           =   "FrmConsumptionProducts.frx":0FFE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4043
      TabIndex        =   14
      Top             =   7845
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
      MICON           =   "FrmConsumptionProducts.frx":101A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnManufacturedProduct 
      Height          =   330
      Left            =   3795
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2505
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
      MICON           =   "FrmConsumptionProducts.frx":1036
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   4155
      TabIndex        =   3
      Top             =   2505
      Width           =   2520
      _ExtentX        =   4445
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   6435
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   2895
      TabIndex        =   1
      Top             =   2505
      Width           =   900
      _ExtentX        =   1588
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   2160
      Left            =   2895
      TabIndex        =   24
      Top             =   2820
      Width           =   6225
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
      stylesets(0).Picture=   "FrmConsumptionProducts.frx":1052
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
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ProductID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2249
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4419
      Columns(2).Caption=   "Product Name"
      Columns(2).Name =   "ProductName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1270
      Columns(3).Caption=   "Rate"
      Columns(3).Name =   "Rate"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1058
      Columns(4).Caption=   "Qty"
      Columns(4).Name =   "Qty"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "PackingID"
      Columns(5).Name =   "PackingID"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1482
      Columns(6).Caption=   "Amount"
      Columns(6).Name =   "Amount"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   10980
      _ExtentY        =   3810
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   6758
      TabIndex        =   0
      Tag             =   "NC"
      Top             =   1545
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   7793
      TabIndex        =   25
      Tag             =   "NC"
      Top             =   1545
      Width           =   2340
      _ExtentX        =   4128
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
      Left            =   7433
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1545
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
      MICON           =   "FrmConsumptionProducts.frx":106E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotalQty 
      Height          =   315
      Left            =   10140
      TabIndex        =   29
      Top             =   6780
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
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
   End
   Begin SITextBox.Txt TxtRate 
      Height          =   315
      Left            =   6675
      TabIndex        =   4
      Top             =   2505
      Width           =   720
      _ExtentX        =   1270
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
      DecimalPoint    =   3
      IntegralPoint   =   5
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   7395
      TabIndex        =   5
      Top             =   2505
      Width           =   600
      _ExtentX        =   1058
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
      DecimalPoint    =   3
      IntegralPoint   =   5
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   7995
      TabIndex        =   6
      Top             =   2505
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   3
      IntegralPoint   =   5
      Mandatory       =   1
   End
   Begin JeweledBut.JeweledButton BtnUsedProduct 
      Height          =   330
      Left            =   3780
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5475
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
      MICON           =   "FrmConsumptionProducts.frx":108A
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtUsedProductName 
      Height          =   315
      Left            =   4140
      TabIndex        =   9
      Top             =   5475
      Width           =   2520
      _ExtentX        =   4445
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
   Begin SITextBox.Txt TxtUsedProductID 
      Height          =   315
      Left            =   2880
      TabIndex        =   7
      Top             =   5475
      Width           =   900
      _ExtentX        =   1588
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridUsed 
      Height          =   1905
      Left            =   2880
      TabIndex        =   37
      Top             =   5790
      Width           =   6225
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   6
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
      stylesets(0).Picture=   "FrmConsumptionProducts.frx":10A6
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
      Columns.Count   =   6
      Columns(0).Width=   2249
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ProductID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4419
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "ProductName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1270
      Columns(2).Caption=   "Rate"
      Columns(2).Name =   "Rate"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1058
      Columns(3).Caption=   "Qty"
      Columns(3).Name =   "Qty"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "PackingID"
      Columns(4).Name =   "PackingID"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1482
      Columns(5).Caption=   "Amount"
      Columns(5).Name =   "Amount"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   10980
      _ExtentY        =   3360
      _StockProps     =   79
      Caption         =   "Used Products"
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
   Begin SITextBox.Txt TxtUsedRate 
      Height          =   315
      Left            =   6660
      TabIndex        =   10
      Top             =   5475
      Width           =   720
      _ExtentX        =   1270
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
      DecimalPoint    =   3
      IntegralPoint   =   5
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtUsedQty 
      Height          =   315
      Left            =   7380
      TabIndex        =   11
      Top             =   5475
      Width           =   600
      _ExtentX        =   1058
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
      DecimalPoint    =   3
      IntegralPoint   =   5
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtUsedAmount 
      Height          =   315
      Left            =   7980
      TabIndex        =   12
      Top             =   5475
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Alignment       =   1
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
      Masked          =   2
      DecimalPoint    =   3
      IntegralPoint   =   5
      Mandatory       =   1
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpManufacturedDate 
      Height          =   315
      Left            =   4485
      TabIndex        =   43
      Top             =   1545
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
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
   Begin SITextBox.Txt TxtManufacturedID 
      Height          =   315
      Left            =   2895
      TabIndex        =   44
      Top             =   1545
      Width           =   1440
      _ExtentX        =   2540
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured Date"
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
      Left            =   4500
      TabIndex        =   46
      Top             =   1350
      Width           =   1650
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured ID"
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
      Left            =   2895
      TabIndex        =   45
      Top             =   1350
      Width           =   1440
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Used Product Name"
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
      Left            =   4140
      TabIndex        =   42
      Top             =   5280
      Width           =   1710
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
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
      Left            =   2880
      TabIndex        =   41
      Top             =   5280
      Width           =   930
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      Left            =   7380
      TabIndex        =   40
      Top             =   5280
      Width           =   300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Left            =   6660
      TabIndex        =   39
      Top             =   5280
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   7980
      TabIndex        =   38
      Top             =   5280
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   7995
      TabIndex        =   36
      Top             =   2310
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Left            =   6675
      TabIndex        =   35
      Top             =   2310
      Width           =   420
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
      Left            =   11385
      TabIndex        =   34
      Top             =   540
      Width           =   435
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantites"
      Height          =   195
      Left            =   10050
      TabIndex        =   30
      Top             =   6510
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
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
      Left            =   6758
      TabIndex        =   28
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
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
      Left            =   7793
      TabIndex        =   27
      Top             =   1350
      Width           =   1005
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consumption Products"
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
      TabIndex        =   23
      Top             =   180
      Width           =   3960
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      Left            =   7395
      TabIndex        =   22
      Top             =   2310
      Width           =   300
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   6435
      TabIndex        =   21
      Top             =   405
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   2895
      TabIndex        =   19
      Top             =   2310
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured Product Name"
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
      Left            =   4155
      TabIndex        =   18
      Top             =   2310
      Width           =   2445
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmConsumptionProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim vCounter1 As Integer
Dim RsBody As New ADODB.Recordset
Dim RsUsedBody As New ADODB.Recordset
'Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim sSql As String
Dim vStrSQL As String
'----------------------------------

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

Private Function FunSelectFinishedProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchFinishedProduct.Show vbModal, Me
      If SchFinishedProduct.ParaOutID = "" Then FunSelectFinishedProduct = False: Exit Function
      TxtCode.Text = SchFinishedProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCode.Text) = "" Then Exit Function
    If Len(TxtCode.Text) <= 5 Then
      TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
    End If
    If TxtCode.Text = "" Then FunSelectFinishedProduct = False: Exit Function
    vStrSQL = " SELECT p.productid, Code, ProductName " & vbCrLf _
           + " from ProductProcessInfoHeader f inner join Products p on f.finishedproductid = p.productid" & vbCrLf _
           + " left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
  
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !ProductID
         TxtProductName.Text = !ProductName
         TxtRate.Text = CN.Execute("Select Cost from currentStock where productID = '" & TxtProductID.Text & "'").Fields(0)
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         FunSelectFinishedProduct = True
         .Close
         Exit Function
      Else
         TxtProductID.Text = ""
         TxtCode.Text = ""
         TxtProductName.Text = ""
         TxtQty.Text = ""
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         FunSelectFinishedProduct = False
         .Close
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   CN.BeginTrans
   'Products Body
   Grid.Redraw = False
   Grid.MoveFirst
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("ProductID").Text) <> "" Then
         CN.Execute "Delete from ConsumptionProductsBody where ManufacturedID=" & Val(TxtManufacturedID.Text) & " and ProductID='" & Grid.Columns("Productid").Text & "'"
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
'   'Products Used
   GridUsed.Redraw = False
   GridUsed.MoveFirst
   For vCounter = 1 To GridUsed.Rows
      If Trim(GridUsed.Columns("ProductID").Text) <> "" Then
         CN.Execute "Delete from ConsumptionProductsUsed where ManufacturedID=" & Val(TxtManufacturedID.Text) & " and ProductID='" & GridUsed.Columns("Productid").Text & "'"
      End If
      GridUsed.MoveNext
   Next vCounter
   GridUsed.RemoveAll
   GridUsed.Redraw = True
   'Header
   Call ActivityLog("Manufactured Products", eDelete, TxtManufacturedID.Text)
   CN.Execute "Delete from ConsumptionProductsHeader where ManufacturedID= " & (TxtManufacturedID.Text)
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchConsumptionProducts.Show vbModal
   If SchConsumptionProducts.ParaOutManufacturedID = 0 Then Exit Sub
   TxtManufacturedID.Text = SchConsumptionProducts.ParaOutManufacturedID
   GetManufacturedProduct
End Sub

'Private Sub BtnPrint_Click()
'On Error GoTo ErrorHandler
'   vStrSql = "select u.username, h.TransferID, h.TransferDate, h.TotalAmount as tbill, isnull(h.discount,0) as discount, isnull(h.cashReceived,0) as cashReceived, p.productname, b.qty, b.price-b.discountvalue as price, b.amount" _
'            + " from StockTransferHeader h inner join StockTransferBody b on h.TransferID = b.TransferID and h.TransferDate = b.TransferDate" _
'            + " inner join products p on p.productid = b.productid" _
'            + " inner join users u on u.UserNo = h.UserNo" _
'            + " where h.TransferID= " & Val(TxtTransferID.Text) & " and h.TransferDate='" & DtpTransferDate.DateValue & "' order by SerialNo"
'
'    If RsReport.State = adStateOpen Then RsReport.Close
'    RsReport.Open vStrSql, CN, adOpenStatic, adLockReadOnly
'
'    Set RptReportViewer.Report = New CrpPurchaseInvoice
'    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
'    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
'    'RptReportViewer.Report.PaperSize = crPaperA4
'    'RptReportViewer.Report.PaperSize = crPaperUser
'    'RptReportViewer.Report.SetUserPaperSize 1400, 1200
'    'RptReportViewer.Report.PaperOrientation = crPortrait
'    'RptReportViewer.Show
'    RptReportViewer.Report.PrintOut False
'Exit Sub
'ErrorHandler:
'    Call ShowErrorMessage
'End Sub

Private Sub BtnManufacturedProduct_Click()
   If FunSelectFinishedProduct(ssButton, True) = True Then
      TxtQty.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
    MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
    Exit Sub
  End If
'  '/******** dummy retriction ***********/
'   If vIsNewRecord = False Then
'      MsgBox "You are not authorized to modify a record", vbCritical, "Error"
'   End If
'   '/************************************/
'  Header Validation
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      TxtStoreID.SetFocus
      Exit Sub
   End If
   If vIsNewRecord = True Then
      If CN.Execute("select * from ConsumptionProductsHeader where ManufacturedId=" & TxtManufacturedID.Text).RecordCount > 0 Then
         MsgBox "Manufactured ID Already Exist.", vbExclamation, Me.Caption
         TxtManufacturedID.SetFocus
         Exit Sub
      End If
   End If
   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one Manufactured Product", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
   End If
   CN.BeginTrans
'   ' delete from manufacture product used
'   With CN.Execute("Select * from ConsumptionProductsUsed where ManufacturedID=" & Val(TxtManufacturedID.Text))
'      While Not .EOF
'         CN.Execute "Delete from ConsumptionProductsUsed where ManufacturedID=" & !ManufacturedID & " and ProductID='" & !ProductID & "'"
'         .MoveNext
'      Wend
'      .Close
'   End With
   If vIsNewRecord = False Then Call ActivityLog("Manufactured Products", eEdit, TxtManufacturedID.Text)
   sSql = "select * from ConsumptionProductsHeader where ManufacturedID=" & Val(TxtManufacturedID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, CN, adOpenStatic, adLockOptimistic
      If .BOF Then
         .AddNew
         !ManufacturedID = Val(TxtManufacturedID.Text)
      End If
      !ManufacturedDate = DtpManufacturedDate.DateValue
      !StoreID = Val(TxtStoreID.Text)
      !UserNo = ObjUserSecurity.UserNo
      .Update
      .Close
   End With
   
  'Body Validation
  ' validation has been performed when a row is added to the grid
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ManufacturedID = TxtManufacturedID.Text
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   
   'Body Validation
  ' validation has been performed when a row is added to the gridused
   With RsUsedBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ManufacturedID = TxtManufacturedID.Text
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   
   If vIsNewRecord = True Then Call ActivityLog("Manufactured Products", eAdd, TxtManufacturedID.Text)
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from ConsumptionProductsBody where ManufacturedID=" & Val(TxtManufacturedID.Text), CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      sSql = "select p.productname, code, b.* from ConsumptionProductsBody b join products p on p.productid = b.productid where ManufacturedID=" & Val(TxtManufacturedID.Text)
      With CN.Execute(sSql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !ProductID
            Grid.Columns("Code").Text = !Code
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !QtyLoose
            Grid.Columns("Rate").Value = !Rate
            Grid.Columns("Amount").Value = !Amount
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
End Sub

Private Sub PopulateDataToUsed()
   RsUsedBody.Filter = 0
   If RsUsedBody.State = adStateOpen Then RsUsedBody.Close
   RsUsedBody.Open "Select * from ConsumptionProductsUsed where ManufacturedID=" & Val(TxtManufacturedID.Text), CN, adOpenStatic, adLockBatchOptimistic
   If RsUsedBody.RecordCount > 0 Then
      sSql = "select p.productname,  u.* from ConsumptionProductsUsed u join products p on p.productid = u.productid where ManufacturedID=" & Val(TxtManufacturedID.Text)
      With CN.Execute(sSql)
         GridUsed.Redraw = False
         GridUsed.MoveFirst
         GridUsed.RemoveAll
         GridUsed.AllowAddNew = True
         While Not .EOF
            GridUsed.AddNew
            GridUsed.Columns("ProductID").Text = !ProductID
            GridUsed.Columns("ProductName").Text = !ProductName
            GridUsed.Columns("Qty").Value = !QtyLoose
            GridUsed.Columns("Rate").Value = !Rate
            GridUsed.Columns("Amount").Value = !Amount
            .MoveNext
         Wend
         .Close
      End With
      GridUsed.AddNew
      GridUsed.Columns("ProductID").Text = " "
      GridUsed.AllowAddNew = False
      GridUsed.Redraw = True
   End If
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
      Call PopulateDataToUsed
      TxtManufacturedID.Text = FunGetMaxID
      TxtCode.Enabled = True
      BtnManufacturedProduct.Enabled = True
      BtnStore.Enabled = True
      TxtStoreID.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtManufacturedID.Enabled = True
      If TxtManufacturedID.Enabled And TxtManufacturedID.Visible Then TxtManufacturedID.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnStore.Enabled = False
      TxtStoreID.Enabled = False
      TxtCode.Enabled = True
      BtnManufacturedProduct.Enabled = True
      TxtManufacturedID.Enabled = False
      vIsNewRecord = False
   Case Is = ChangeMode
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtCode.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub GridUsed_DblClick()
   Call GridUsed_LostFocus
End Sub

Private Sub GridUsed_GotFocus()
   TxtUsedProductID.Enabled = False
   BtnUsedProduct.Enabled = False
End Sub

Private Sub GridUsed_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub GridUsed_LostFocus()
   If Trim(GridUsed.Columns("ProductID").Text) = "" Then
      TxtUsedProductID.Text = ""
      TxtUsedProductID.Enabled = True
      BtnUsedProduct.Enabled = True
      TxtUsedProductID.SetFocus
   Else
      TxtUsedProductID.Enabled = False
      BtnUsedProduct.Enabled = False
      TxtUsedQty.SetFocus
   End If
End Sub

Private Sub GridUsed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(GridUsed.Columns("Productid").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub GridUsed_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    GetDataBackFromGridUsedToTexBoxes
End Sub

Private Sub TxtUsedProductID_Change()
If TxtUsedProductID.Enabled = False Then Exit Sub
If TxtUsedAmount.Visible = False Then Exit Sub
 If Me.ActiveControl.Name <> TxtUsedProductID.Name Then Exit Sub
   If TxtUsedProductName.Text <> "" Then
      TxtUsedProductID.Text = ""
      TxtUsedProductName.Text = ""
      TxtUsedQty.Text = ""
   End If
If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub TxtUsedProductID_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtUsedProductID.Text) = "" Then Exit Sub
   vTemp = Not FunSelectUsedProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectUsedProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnUsedProduct_Click()
   If FunSelectUsedProduct(ssButton, True) = True Then
      TxtUsedQty.SetFocus
   Else
      TxtUsedProductID.SetFocus
   End If
End Sub

Private Function FunSelectUsedProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectUsedProduct = False: Exit Function
      TxtUsedProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtUsedProductID.Text) = "" Then Exit Function
    If Len(TxtUsedProductID.Text) <= 5 Then
      TxtUsedProductID.Text = Right("00000" + CStr(Val(TxtUsedProductID.Text)), 5)
    End If
    If TxtUsedProductID.Text = "" Then FunSelectUsedProduct = False: Exit Function
    vStrSQL = " SELECT p.Productid, ProductName" & vbCrLf _
           + " from Products p" & vbCrLf _
           + " where p.productid = '" & TxtUsedProductID.Text & "'"
  
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtUsedProductID.Text = !ProductID
         TxtUsedProductName.Text = !ProductName
         TxtUsedRate.Text = CN.Execute("Select Cost from CurrentStock where ProductID = '" & TxtUsedProductID.Text & "'").Fields(0)
         FunSelectUsedProduct = True
         .Close
         Exit Function
      Else
         FunSelectUsedProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtUsedProductID.Text = ""
         TxtUsedProductName.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub GetDataFromTexBoxesToGridused()
On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   If Trim(TxtUsedProductID.Text) = "" Then
      TxtUsedProductID.SetFocus
      Exit Sub
   End If
   If Val(TxtUsedQty.Text) = 0 Then
      TxtUsedQty.SetFocus
      Exit Sub
   End If
   RsUsedBody.Filter = "ProductID='" & TxtUsedProductID.Text & "'"
   If TxtUsedProductID.Enabled Then
      If RsUsedBody.RecordCount = 0 Then
         RsUsedBody.AddNew
         GridUsed.Columns("ProductID").Text = TxtUsedProductID.Text
         RsUsedBody!ProductID = TxtProductID.Text
         RsUsedBody!ProductID = TxtUsedProductID.Text
      Else
         GridUsed.Redraw = False
         GridUsed.MoveFirst
            For vrowcounter = 1 To GridUsed.Rows
               If GridUsed.Columns("ProductID").Text = TxtUsedProductID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearUsedDetailArea
'                  Call AddDataToGridUsed
                  TxtUsedQty.Text = Val(TxtUsedQty.Text) + GridUsed.Columns("Qty").Value
                  TxtUsedAmount.Text = Val(TxtUsedAmount.Text) + GridUsed.Columns("Amount").Value
                  GridUsed.Columns("ProductName").Text = TxtUsedProductName.Text
                  GridUsed.Columns("Rate").Text = TxtUsedRate.Text
                  GridUsed.Columns("Qty").Value = Val(TxtUsedQty.Text)
                  GridUsed.Columns("Amount").Value = Val(TxtUsedAmount.Text)
                  RsUsedBody!QtyLoose = Val(TxtUsedQty.Text)
                  RsUsedBody!Rate = Val(TxtUsedRate.Text)
                  RsUsedBody!Amount = Val(TxtUsedAmount.Text)
                  GridUsed.MoveLast
                  Call SubClearUsedDetailArea
                  TxtUsedProductID.SetFocus
                  GridUsed.Redraw = True
                  Exit Sub
               End If
               GridUsed.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearUsedDetailArea
         GridUsed.MoveLast
         TxtUsedProductID.SetFocus
         Exit Sub
      End If
   End If
   GridUsed.Redraw = False
   With GridUsed
'      If TxtUsedProductID.Enabled = False Then Call AddDataToGridUsed
      .Columns("ProductName").Text = TxtUsedProductName.Text
      .Columns("Qty").Text = TxtUsedQty.Text
      .Columns("Rate").Text = TxtUsedRate.Text
      .Columns("Amount").Value = Val(TxtUsedAmount.Text)
'      If TxtUsedProductID.Enabled = True Then Call AddDataToGridUsed
      RsUsedBody!QtyLoose = Val(TxtUsedQty.Text)
      RsUsedBody!Rate = Val(TxtUsedRate.Text)
      RsUsedBody!Amount = Val(TxtUsedAmount.Text)
      .MoveLast
      If Trim(.Columns("ProductID").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("ProductID").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearUsedDetailArea
   TxtUsedProductID.SetFocus
'   Call CalculateTotal
   GridUsed.Redraw = True
   Exit Sub
ErrorHandler:
   GridUsed.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub GetDataBackFromGridUsedToTexBoxes()
   On Error GoTo ErrorHandler
   With GridUsed
      TxtUsedProductID.Text = .Columns("ProductID").Text
      TxtUsedProductName.Text = .Columns("ProductName").Text
      TxtUsedQty.Text = .Columns("Qty").Text
      TxtUsedRate.Text = .Columns("Rate").Text
      TxtUsedAmount.Text = .Columns("Amount").Text
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
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
      If TxtCode.Enabled Then TxtCode.SetFocus: Call SubClearDetailArea
   ElseIf Shift = vbCtrlMask Then
      If ActiveControl.Name = Grid.Name Then
         If KeyCode = vbKeyDelete Then
            If Trim(Grid.Columns("ProductID").Text <> "") Then Call mniRemoveRow_Click
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
'         Case vbKeyP
'            If BtnPrint.Enabled Then BtnPrint_Click
'            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtStoreID.SetFocus
         Case TxtCode.Name: If FunSelectFinishedProduct(ssFunctionKey, False) = True Then TxtQty.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
         KeyCode = 0
         If BtnSave.Enabled Then BtnSave.SetFocus
      End If
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub LblClose_Click()
   FraHelp.Visible = False
End Sub

Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblHelp.ForeColor = &H800000
   FraHelp.ZOrder 0
   FraHelp.Visible = True
End Sub

Private Sub LblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If LblHelp.FontUnderline = True Then Exit Sub
   LblHelp.FontUnderline = True
End Sub

Private Sub LblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblHelp.ForeColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Manufactured Products"
   HelpLocation Me
   With CN.Execute("select * from registry")
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
         FunSelectStore ssValidate, True
         TxtStoreID.Visible = !StoreVisible
         BtnStore.Visible = !StoreVisible
         TxtStoreName.Visible = !StoreVisible
         LblStoreID.Visible = !StoreVisible
         LblStoreName.Visible = !StoreVisible
      End If
      .Close
   End With
   FormStatus = NewMode
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
      End If
   Next
   
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("Code").Text = " "
   Grid.Update
   
   GridUsed.CancelUpdate
   GridUsed.RemoveAll
   GridUsed.AddNew
   GridUsed.Columns("ProductID").Text = " "
   GridUsed.Update
   
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
    Set RsBody = Nothing
    Set FrmConsumptionProducts = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
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
   TxtCode.Enabled = False
   BtnManufacturedProduct.Enabled = False
   'TxtCode.BackColor = TxtProductName.BackColor
   'TxtCode.TabStop = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("Code").Text) = "" Then
      TxtCode.Text = ""
      TxtCode.Enabled = True
      BtnManufacturedProduct.Enabled = True
      TxtCode.SetFocus
   Else
      TxtCode.Enabled = False
      BtnManufacturedProduct.Enabled = False
      TxtQty.SetFocus
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(Grid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub GridUsed_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
'   TxtTotalQty.Text = Val(TxtTotalQty.Text) - Grid.Columns("Qty").Value
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub CalculateTotal()
On Error GoTo ErrorHandler
   Dim bm As Variant
   Dim i As Integer
   GridUsed.Redraw = False
   GridUsed.MoveFirst
   TxtTotalQty.Text = ""
   For i = 0 To GridUsed.Rows - 1
      bm = GridUsed.GetBookmark(i)
      TxtTotalQty.Text = Val(TxtTotalQty.Text) + GridUsed.Columns(2).CellValue(bm)
   Next i
   GridUsed.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name = "Grid" Then
    If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
    RsBody.Filter = "Code='" & TxtCode.Text & "'"
    If RsBody.RecordCount > 0 Then RsBody.Delete
    Grid.SelBookmarks.RemoveAll
    Grid.SelBookmarks.Add Grid.Bookmark
'    RemoveDataToGridUsed
    Grid.DeleteSelected
    Grid.Refresh
    RsBody.Filter = 0
    Grid.MoveLast
'    Call CalculateTotal
    GetDataBackFromGridToTexBoxes
   ElseIf Me.ActiveControl.Name = "GridUsed" Then
    If Trim(GridUsed.Columns("ProductID").Text) = "" Then Exit Sub
    RsUsedBody.Filter = "ProductID='" & TxtUsedProductID.Text & "'"
    If RsUsedBody.RecordCount > 0 Then RsUsedBody.Delete
    GridUsed.SelBookmarks.RemoveAll
    GridUsed.SelBookmarks.Add GridUsed.Bookmark
'    RemoveDataToGridUsed
    GridUsed.DeleteSelected
    GridUsed.Refresh
    RsUsedBody.Filter = 0
    GridUsed.MoveLast
'    Call CalculateTotal
    GetDataBackFromGridUsedToTexBoxes
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   If Trim(TxtCode.Text) = "" Then
      TxtCode.SetFocus
      Exit Sub
   End If
   If Val(TxtQty.Text) = 0 Then
      TxtQty.SetFocus
      Exit Sub
   End If
   RsBody.Filter = "Code='" & TxtCode.Text & "'"
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ProductID").Text = TxtProductID.Text
         Grid.Columns("Code").Text = TxtCode.Text
         RsBody!ProductID = TxtProductID.Text
         RsBody!Code = TxtCode.Text
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("Code").Text = TxtProductID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
'                  Call AddDataToGridUsed
                  TxtQty.Text = Val(TxtQty.Text) + Grid.Columns("Qty").Value
                  TxtAmount.Text = Val(TxtAmount.Text) + Grid.Columns("Amount").Value
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("Rate").Text = TxtRate.Text
                  Grid.Columns("Qty").Value = Val(TxtQty.Text)
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  RsBody!QtyLoose = Val(TxtQty.Text)
                  RsBody!Rate = Val(TxtRate.Text)
                  RsBody!Amount = Val(TxtAmount.Text)
                  Grid.MoveLast
                  Call SubClearDetailArea
                  TxtCode.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearDetailArea
         Grid.MoveLast
         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   Grid.Redraw = False
   With Grid
'      If TxtCode.Enabled = False Then Call AddDataToGridUsed
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("Qty").Text = TxtQty.Text
      .Columns("Rate").Text = TxtRate.Text
      .Columns("Amount").Value = Val(TxtAmount.Text)
'      If TxtCode.Enabled = True Then Call AddDataToGridUsed
      RsBody!QtyLoose = Val(TxtQty.Text)
      RsBody!Rate = Val(TxtRate.Text)
      RsBody!Amount = Val(TxtAmount.Text)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearDetailArea
   TxtCode.SetFocus
'   Call CalculateTotal
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub AddDataToGridUsed()
   If Trim(Grid.Columns("ProductID").Text) = "" Then Exit Sub
   'GridUsed.Redraw = False
   Dim Flag1 As Boolean
   With CN.Execute("exec spFinishedProducts '" & Grid.Columns("ProductID").Text & "'")
      For vCounter1 = 1 To .RecordCount
         Flag1 = True
         GridUsed.MoveFirst
         For vCounter = 1 To GridUsed.Rows
            If GridUsed.Columns("ProductID").Text = !ProductID Then
               GridUsed.Columns("Rate").Value = !Rate
               If TxtCode.Enabled = True Then
                  GridUsed.Columns("Qty").Value = GridUsed.Columns("Qty").Value + (!QtyLoose * Val(TxtQty.Text))
               Else
                  GridUsed.Columns("Qty").Value = GridUsed.Columns("Qty").Value - (!QtyLoose * Grid.Columns("Qty").Value) + (!QtyLoose * Val(TxtQty.Text))
               End If
               GridUsed.Columns("Amount").Value = GridUsed.Columns("Rate").Value * GridUsed.Columns("Qty").Value
               Flag1 = False
            End If
            GridUsed.MoveNext
         Next vCounter
         GridUsed.MoveLast
         If Flag1 = True Then
            GridUsed.AllowAddNew = True
            GridUsed.AddNew
            GridUsed.Columns("ProductID").Text = !ProductID
            GridUsed.Columns("ProductName").Text = !ProductName
            GridUsed.Columns("Rate").Value = !Rate
            GridUsed.Columns("Qty").Value = (!QtyLoose * Grid.Columns("Qty").Value)
            GridUsed.Columns("Amount").Value = GridUsed.Columns("Rate").Value * GridUsed.Columns("Qty").Value
            GridUsed.Update
            GridUsed.AllowAddNew = False
         End If
         .MoveNext
      Next vCounter1
      .Close
      GridUsed.Redraw = True
   End With
   If BtnSave.Enabled Then FormStatus = SelectionMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub RemoveDataToGridUsed()
   If Trim(Grid.Columns("ProductID").Text) = "" Then Exit Sub
   GridUsed.Redraw = False
   GridUsed.MoveFirst
   With CN.Execute("exec spFinishedProducts '" & Grid.Columns("ProductID").Text & "'")
      For vCounter1 = 1 To .RecordCount
         GridUsed.MoveFirst
         For vCounter = 1 To GridUsed.Rows
            If GridUsed.Columns("ProductID").Text = !ProductID Then
               GridUsed.Columns("Qty").Value = GridUsed.Columns("Qty").Value - (!QtyLoose * Grid.Columns("Qty").Value)
               GridUsed.Columns("Amount").Value = GridUsed.Columns("Rate").Value * GridUsed.Columns("Qty").Value
            End If
            If GridUsed.Columns("Qty").Value = 0 Then
               GridUsed.SelBookmarks.RemoveAll
               GridUsed.SelBookmarks.Add GridUsed.Bookmark
               GridUsed.DeleteSelected
               Grid.Refresh
            Else
               GridUsed.MoveNext
            End If
         Next vCounter
         .MoveNext
      Next vCounter1
   End With
   GridUsed.Redraw = True
   If BtnSave.Enabled Then FormStatus = SelectionMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtCode.Enabled = True
   BtnManufacturedProduct.Enabled = True
   TxtCode.Text = ""
   TxtProductName.Text = ""
   TxtQty.Text = ""
   TxtRate.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub SubClearUsedDetailArea()
   TxtUsedProductID.Enabled = True
   BtnUsedProduct.Enabled = True
   TxtUsedProductID.Text = ""
   TxtUsedProductName.Text = ""
   TxtUsedQty.Text = ""
   TxtUsedRate.Text = ""
   TxtUsedAmount.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("Code").Text
      TxtProductName.Text = .Columns("ProductName").Text
      TxtQty.Text = .Columns("Qty").Text
      TxtRate.Text = .Columns("Rate").Text
      TxtAmount.Text = .Columns("Amount").Text
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetManufacturedProduct()
   On Error GoTo ErrorHandler
   sSql = "select ManufacturedID, ManufacturedDate, h.StoreID, StoreName FROM ConsumptionProductsHeader h join Stores s on h.storeid = s.storeid where ManufacturedID=" & Val(TxtManufacturedID.Text)
   With CN.Execute(sSql)
      If Not .BOF Then
          TxtManufacturedID.Text = !ManufacturedID
          DtpManufacturedDate.DateValue = !ManufacturedDate
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
      End If
      .Close
   End With
   Call PopulateDataToGrid
   Call PopulateDataToUsed
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
   End If
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtCode_Validate(Cancel As Boolean)
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = FunSelectFinishedProduct(ssValidate, False)
   If vTemp = False Then   '   vTemp = FunSelectFinishedProduct(ssButton, False)
      Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtQty_Change()
    If ActiveControl.Name <> TxtQty.Name Then Exit Sub
    TxtAmount.Text = Val(TxtRate.Text) * Val(TxtQty.Text)
End Sub

Private Sub TxtQty_LostFocus()
   Select Case ActiveControl.Name
   Case TxtCode.Name, TxtRate.Name
      Exit Sub
   End Select
   GetDataFromTexBoxesToGrid
End Sub


Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = CN.Execute("Select isnull(max(ManufacturedID),0)+1 from ConsumptionProductsHeader").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtRate_Change()
    If ActiveControl.Name <> TxtRate.Name Then Exit Sub
    TxtAmount.Text = Val(TxtRate.Text) * Val(TxtQty.Text)
End Sub

Private Sub TxtStoreID_Change()
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then
      TxtStoreName.Text = ""
   End If
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
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtUsedQty_Change()
    If ActiveControl.Name <> TxtUsedQty.Name Then Exit Sub
    TxtUsedAmount.Text = Val(TxtUsedRate.Text) * Val(TxtUsedQty.Text)
End Sub

Private Sub TxtUsedQty_LostFocus()
  Select Case ActiveControl.Name
   Case TxtUsedProductID.Name, TxtUsedRate.Name
      Exit Sub
   End Select
   GetDataFromTexBoxesToGridused
End Sub

Private Sub TxtUsedRate_Change()
    If ActiveControl.Name <> TxtUsedRate.Name Then Exit Sub
    TxtUsedAmount.Text = Val(TxtUsedRate.Text) * Val(TxtUsedQty.Text)
End Sub
